from __future__ import annotations

import json
import logging
import re
import shutil
import sys
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Dict, Hashable, Iterable, Iterator, List, TypeVar
from zoneinfo import ZoneInfo

from dotenv import load_dotenv
from flask import (
    Flask,
    Response,
    abort,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)

from pdf_merge import merge_pdfs
from word_to_pdf_converter import convert_word_to_pdf
from email_service import EmailConfig, send_email_with_attachment


PERSON_NAME_SEPARATOR_PATTERN = re.compile(r"[・･.,，、．｡\s]+")
PERSON_NORMALIZATION_PATTERN = re.compile(r"[\s・･.,，、．｡]+")


JST = ZoneInfo("Asia/Tokyo")


def _resolve_base_dir() -> Path:
    """Return the directory containing the running application."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _resolve_base_dir()


def _load_environment_variables() -> None:
    """Load environment variables from .env files in common locations."""

    candidates = []

    # .env distributed alongside the executable/script.
    candidates.append(BASE_DIR / ".env")

    # .env bundled inside PyInstaller one-file archives lives under _MEIPASS.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(meipass) / ".env")

    # Support running from an arbitrary working directory during development.
    cwd = Path.cwd()
    if cwd != BASE_DIR:
        candidates.append(cwd / ".env")

    loaded_any = False
    seen: set[Path] = set()
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        if not candidate.exists():
            continue
        load_dotenv(candidate)
        loaded_any = True

    if not loaded_any:
        logging.getLogger(__name__).debug(
            "No .env file loaded. Checked locations: %s",
            ", ".join(str(path) for path in seen) or "(none)",
        )


_load_environment_variables()


DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
WORK_DIR = DATA_DIR / "work"
ORDER_FILE = BASE_DIR / "order.json"

for directory in (UPLOAD_DIR, WORK_DIR):
    directory.mkdir(parents=True, exist_ok=True)


T = TypeVar("T", bound=Hashable)


def _iter_unique(items: Iterable[T]) -> Iterator[T]:
    """Yield items while preserving the first occurrence order."""

    seen: set[T] = set()
    for item in items:
        if item in seen:
            continue
        seen.add(item)
        yield item


def _deduplicate_list(items: Iterable[T]) -> List[T]:
    """Return a list containing the first occurrence of every item in ``items``."""

    return list(_iter_unique(items))


def _cleanup_data_directories() -> None:
    """Remove generated files from the upload and work directories."""

    for directory in (UPLOAD_DIR, WORK_DIR):
        if not directory.exists():
            continue
        for path in directory.iterdir():
            if path.is_dir():
                shutil.rmtree(path, ignore_errors=True)
            else:
                try:
                    path.unlink()
                except FileNotFoundError:
                    continue
        directory.mkdir(parents=True, exist_ok=True)


def _schedule_delayed_cleanup(delay_seconds: int = 600) -> None:
    """Schedule a cleanup of generated files after the specified delay."""

    def _cleanup_task() -> None:
        with jobs_lock:
            active_jobs = any(
                job.status not in {"completed", "failed"} for job in jobs.values()
            )
        if active_jobs:
            _schedule_delayed_cleanup(delay_seconds)
            return
        _cleanup_data_directories()

    timer = threading.Timer(delay_seconds, _cleanup_task)
    timer.daemon = True
    timer.start()


_cleanup_data_directories()


@dataclass
class ZipEntry:
    identifier: str
    display_name: str
    archive_name: str
    team_name: str | None = None
    persons: List[str] | None = None
    sanitized_name: str | None = None


UNGROUPED_TEAM_KEY = "__ungrouped__"


def _normalize_team_key(value: str | None) -> str:
    """Return a consistent key for grouping teams."""

    if not isinstance(value, str):
        return UNGROUPED_TEAM_KEY
    candidate = value.strip()
    if not candidate:
        return UNGROUPED_TEAM_KEY
    if candidate in {UNGROUPED_TEAM_KEY, "班なし"}:
        return UNGROUPED_TEAM_KEY
    return candidate


@dataclass
class OrderPreferences:
    """Serializable storage for preferred team/member ordering."""

    team_sequence: List[str]
    member_sequences: Dict[str, List[str]]

    @classmethod
    def empty(cls) -> "OrderPreferences":
        return cls(team_sequence=[], member_sequences={})

    def to_dict(self) -> Dict[str, List[str]]:
        return {
            "team_sequence": self.team_sequence,
            "member_sequences": self.member_sequences,
        }

    @classmethod
    def from_dict(cls, payload: Dict[str, object]) -> "OrderPreferences":
        raw_team_sequence = payload.get("team_sequence", [])
        team_sequence: List[str] = []
        if isinstance(raw_team_sequence, list):
            team_sequence = _deduplicate_list(
                _normalize_team_key(value)
                for value in raw_team_sequence
                if isinstance(value, str)
            )

        raw_members = payload.get("member_sequences", {})
        member_sequences: Dict[str, List[str]] = {}
        if isinstance(raw_members, dict):
            for raw_key, raw_value in raw_members.items():
                if not isinstance(raw_key, str) or not isinstance(raw_value, list):
                    continue
                team_key = _normalize_team_key(raw_key)
                members = _normalize_member_names(raw_value)
                if not members:
                    continue
                existing = member_sequences.setdefault(team_key, [])
                for member in members:
                    if member not in existing:
                        existing.append(member)

        if UNGROUPED_TEAM_KEY in member_sequences and not member_sequences[
            UNGROUPED_TEAM_KEY
        ]:
            member_sequences.pop(UNGROUPED_TEAM_KEY, None)

        if UNGROUPED_TEAM_KEY not in member_sequences:
            team_sequence = [
                team for team in team_sequence if team != UNGROUPED_TEAM_KEY
            ]

        return cls(team_sequence=team_sequence, member_sequences=member_sequences)

    @classmethod
    def from_legacy_list(cls, items: List[str]) -> "OrderPreferences":
        member_sequences: Dict[str, List[str]] = {}
        team_sequence: List[str] = []
        for name in items:
            if not isinstance(name, str):
                continue
            team_key = UNGROUPED_TEAM_KEY
            stripped = name
            if name.startswith("[") and "]" in name:
                prefix, _, remainder = name.partition("]")
                team_name = prefix[1:]
                team_key = team_name or UNGROUPED_TEAM_KEY
                stripped = remainder.strip()
            team_key = _normalize_team_key(team_key)
            if team_key not in team_sequence:
                team_sequence.append(team_key)
            members = member_sequences.setdefault(team_key, [])
            if stripped and stripped not in members:
                members.append(stripped)
        return cls(team_sequence=team_sequence, member_sequences=member_sequences)


def _normalize_member_names(members: Iterable[str]) -> List[str]:
    cleaned: List[str] = []
    for member in members:
        if not isinstance(member, str):
            continue
        stripped = member.strip()
        if stripped and stripped not in cleaned:
            cleaned.append(stripped)
    return cleaned

class OrderManager:
    def __init__(self, storage_file: Path) -> None:
        self.storage_file = storage_file
        self._lock = threading.Lock()

    def _write_preferences(self, preferences: OrderPreferences) -> None:
        self.storage_file.parent.mkdir(parents=True, exist_ok=True)
        with self.storage_file.open("w", encoding="utf-8") as fh:
            json.dump(preferences.to_dict(), fh, ensure_ascii=False, indent=2)

    def load_preferences(self) -> OrderPreferences:
        if not self.storage_file.exists():
            return OrderPreferences.empty()

        try:
            with self.storage_file.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except json.JSONDecodeError:
            return OrderPreferences.empty()

        if isinstance(data, dict):
            return OrderPreferences.from_dict(data)

        if isinstance(data, list) and all(isinstance(item, str) for item in data):
            return OrderPreferences.from_legacy_list(data)

        return OrderPreferences.empty()

    def save_member_sequence(self, team_key: str, members: List[str]) -> None:
        normalized_team = _normalize_team_key(team_key)
        cleaned_members = _normalize_member_names(members)

        with self._lock:
            preferences = self.load_preferences()
            if cleaned_members:
                if normalized_team not in preferences.team_sequence:
                    preferences.team_sequence.append(normalized_team)
                preferences.member_sequences[normalized_team] = cleaned_members
            else:
                preferences.member_sequences.pop(normalized_team, None)
                preferences.team_sequence = [
                    key for key in preferences.team_sequence if key != normalized_team
                ]
            self._write_preferences(preferences)

    def delete_member_sequence(self, team_key: str) -> None:
        normalized_team = _normalize_team_key(team_key)
        with self._lock:
            preferences = self.load_preferences()
            preferences.member_sequences.pop(normalized_team, None)
            preferences.team_sequence = [
                key for key in preferences.team_sequence if key != normalized_team
            ]
            self._write_preferences(preferences)

    def initial_layout(self, entries: List[ZipEntry]) -> tuple[List[str], Dict[str, List[ZipEntry]]]:
        preferences = self.load_preferences()
        team_map: Dict[str, List[ZipEntry]] = {}
        team_appearance: List[str] = []

        for entry in entries:
            if entry.team_name:
                team_key = _normalize_team_key(entry.team_name)
            else:
                team_key = UNGROUPED_TEAM_KEY
            team_map.setdefault(team_key, []).append(entry)
            if team_key not in team_appearance:
                team_appearance.append(team_key)

        team_sequence: List[str] = []
        for team_key in preferences.team_sequence:
            if team_key in team_map and team_key not in team_sequence:
                team_sequence.append(team_key)
        for team_key in team_appearance:
            if team_key not in team_sequence:
                team_sequence.append(team_key)

        ordered_entries: Dict[str, List[ZipEntry]] = {}
        for team_key in team_sequence:
            members = preferences.member_sequences.get(team_key, [])
            items = team_map.get(team_key, [])
            ordered_entries[team_key] = self._sort_team_entries(items, members)

        return team_sequence, ordered_entries

    def _sort_team_entries(self, items: List[ZipEntry], member_order: List[str]) -> List[ZipEntry]:
        if not items:
            return []

        fallback_positions = {
            entry.identifier: index for index, entry in enumerate(sorted(items, key=lambda e: e.display_name.lower()))
        }

        def sort_key(entry: ZipEntry) -> tuple[int, int]:
            persons = entry.persons or []
            matched_index = _find_member_order_index(member_order, persons)
            if matched_index is not None:
                return (0, matched_index)
            return (1, fallback_positions[entry.identifier])

        return sorted(items, key=sort_key)


order_manager = OrderManager(ORDER_FILE)


@dataclass
class JobState:
    id: str
    email: str
    status: str
    message: str
    created_at: datetime
    updated_at: datetime
    order: List[str]
    zip_path: Path
    zip_original_name: str
    entries: Dict[str, ZipEntry]
    team_options: List[str]
    merged_pdf: Path | None = None
    processing_started_at: datetime | None = None
    processing_completed_at: datetime | None = None
    report_number: str | None = None
    progress_current: int = 0
    progress_total: int = 0
    email_delivery_status: str = "pending"
    conversion_order: List[str] = field(default_factory=list)
    conversion_statuses: Dict[str, str] = field(default_factory=dict)
    conversion_threads: Dict[str, int] = field(default_factory=dict)
    show_conversion_progress: bool = False

    def to_dict(self) -> Dict[str, object]:
        elapsed_seconds: float | None = None
        elapsed_display: str | None = None
        if self.processing_started_at and self.processing_completed_at:
            elapsed_seconds = (
                self.processing_completed_at - self.processing_started_at
            ).total_seconds()
            elapsed_display = _format_elapsed(elapsed_seconds)

        progress_percent: int | None = None
        if self.progress_total > 0:
            ratio = self.progress_current / self.progress_total
            progress_percent = int(round(min(max(ratio, 0.0), 1.0) * 100))

        return {
            "id": self.id,
            "email": self.email,
            "status": self.status,
            "message": self.message,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
            "team_options": self.team_options,
            "report_number": self.report_number,
            "final_pdf_name": self.merged_pdf.name if self.merged_pdf else None,
            "elapsed_seconds": elapsed_seconds,
            "elapsed_display": elapsed_display,
            "progress_current": self.progress_current,
            "progress_total": self.progress_total,
            "progress_percent": progress_percent,
            "email_delivery_status": self.email_delivery_status,
            "conversion_progress": [
                {
                    "display_name": name,
                    "status": self.conversion_statuses.get(name, "queued"),
                    "thread": self.conversion_threads.get(name),
                }
                for name in self.conversion_order
            ],
            "show_conversion_progress": self.show_conversion_progress,
        }


app = Flask(__name__)
app.secret_key = "pdf-report-converter"

executor = ThreadPoolExecutor(max_workers=2)
jobs: Dict[str, JobState] = {}
jobs_lock = threading.Lock()
upload_sessions: Dict[str, Dict[str, object]] = {}


EMAIL_CONFIG = EmailConfig.from_env()

DEFAULT_RECIPIENT_EMAIL = "roboken.report.tool@gmail.com"


def _infer_team_directory_level(paths: List[PurePosixPath]) -> int | None:
    directories = [path.parts[:-1] for path in paths]
    max_depth = max((len(parts) for parts in directories), default=0)
    for level in range(max_depth):
        names = {parts[level] for parts in directories if len(parts) > level}
        if len(names) > 1:
            return level
    return None


def _append_duplicate_suffix(base_name: str, counter: int) -> str:
    path = Path(base_name)
    return f"{path.stem} ({counter}){path.suffix}"


def _format_elapsed(total_seconds: float) -> str:
    seconds = int(round(total_seconds))
    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    parts: List[str] = []
    if hours:
        parts.append(f"{hours}時間")
    if minutes or hours:
        parts.append(f"{minutes}分")
    parts.append(f"{seconds}秒")
    return "".join(parts)


def _sanitize_report_filename(original_name: str) -> str:
    path = Path(original_name)
    suffix = path.suffix
    stem = path.stem

    stem = re.sub(r"[･.,，、．｡]+", "・", stem)
    stem = re.sub(r"[₋_＿\s]+", " ", stem)
    stem = re.sub(r"報告会", "報告書", stem)
    stem = re.sub(r"(報告書)(?![\s・,，、])", r"\1 ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()

    if "回" in stem and "報告書" not in stem:
        pos = stem.find("回")
        if pos != -1:
            insert_pos = pos + 1
            stem = stem[:insert_pos] + "報告書" + stem[insert_pos:]

    if "報告書" not in stem:
        stem = stem.replace("回 ", "回報告書 ")
        stem = re.sub(r"(報告書)(?![\s・,，、])", r"\1 ", stem)

    return f"{stem}{suffix}"


def _extract_person_names(sanitized_name: str) -> List[str]:
    stem = Path(sanitized_name).stem
    if "報告書" in stem:
        _, _, remainder = stem.partition("報告書")
    else:
        remainder = stem
    remainder = remainder.strip()
    if not remainder:
        return []
    tokens = PERSON_NAME_SEPARATOR_PATTERN.split(remainder)
    persons = [token.strip() for token in tokens if token.strip()]
    return persons


def _normalize_person_token(value: str) -> str:
    normalized = PERSON_NORMALIZATION_PATTERN.sub("", value)
    return normalized.lower()


def _find_member_order_index(member_order: List[str], persons: List[str]) -> int | None:
    if not member_order or not persons:
        return None

    normalized_order: List[tuple[str, int]] = []
    for index, name in enumerate(member_order):
        if not isinstance(name, str):
            continue
        normalized = _normalize_person_token(name)
        if normalized:
            normalized_order.append((normalized, index))

    if not normalized_order:
        return None

    best_index: int | None = None
    for person in persons:
        if not isinstance(person, str):
            continue
        normalized_person = _normalize_person_token(person)
        if not normalized_person:
            continue
        for normalized_name, order_index in normalized_order:
            if not normalized_name:
                continue
            if (
                normalized_person == normalized_name
                or normalized_person in normalized_name
                or normalized_name in normalized_person
            ):
                if best_index is None or order_index < best_index:
                    best_index = order_index
                if best_index == 0:
                    return best_index
                break

    return best_index


def _extract_report_number_from_name(name: str | None) -> str | None:
    if not name:
        return None
    stem = Path(name).stem
    match = re.search(r"第\s*(\d{1,3})\s*回", stem)
    if match:
        return match.group(1)
    digits = re.findall(r"\d+", stem)
    if digits:
        return digits[0]
    return None


def _determine_report_number(zip_original_name: str, entries: List[ZipEntry]) -> str:
    number = _extract_report_number_from_name(zip_original_name)
    if number:
        return number

    counts: Dict[str, int] = {}
    for entry in entries:
        candidate = _extract_report_number_from_name(entry.sanitized_name)
        if not candidate:
            candidate = _extract_report_number_from_name(entry.display_name)
        if not candidate:
            continue
        counts[candidate] = counts.get(candidate, 0) + 1

    if counts:
        def sort_key(item: tuple[str, int]) -> tuple[int, int]:
            value, count = item
            try:
                numeric_value = int(value)
            except ValueError:
                numeric_value = -1
            return (-count, -numeric_value)

        return sorted(counts.items(), key=sort_key)[0][0]

    return "1"


def _team_display_label(team_key: str) -> str:
    return "班なし" if team_key == UNGROUPED_TEAM_KEY else team_key


def _team_labels_from_preferences(preferences: OrderPreferences) -> List[str]:
    labels: List[str] = []
    seen: set[str] = set()

    for team_key in preferences.team_sequence:
        label = _team_display_label(team_key)
        if label and label not in seen:
            labels.append(label)
            seen.add(label)

    for team_key in preferences.member_sequences:
        label = _team_display_label(team_key)
        if label and label not in seen:
            labels.append(label)
            seen.add(label)

    return labels


def _build_display_name(
    sanitized_name: str,
    team_name: str | None,
    duplicate_counter: Dict[str, int],
) -> str:
    base_name = Path(sanitized_name).name
    if team_name:
        prefix = f"[{team_name}] "
        if base_name.startswith(prefix):
            prefixed_name = base_name
        else:
            prefixed_name = f"{prefix}{base_name}"
    else:
        prefixed_name = base_name

    key = prefixed_name.lower()
    duplicate_counter[key] = duplicate_counter.get(key, 0) + 1
    occurrence = duplicate_counter[key]
    if occurrence == 1:
        return prefixed_name
    return _append_duplicate_suffix(prefixed_name, occurrence)


def _extract_entries(zip_path: Path, *, original_name: str | None = None) -> List[ZipEntry]:
    entries: List[ZipEntry] = []
    word_infos: List[tuple[zipfile.ZipInfo, PurePosixPath]] = []

    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            suffix = Path(info.filename).suffix.lower()
            if suffix not in {".doc", ".docx"}:
                continue
            word_infos.append((info, PurePosixPath(info.filename)))

    if not word_infos:
        return entries

    team_level = _infer_team_directory_level([path for _, path in word_infos])
    default_team_name: str | None = None
    if team_level is None:
        if original_name:
            default_team_name = Path(original_name).stem
        else:
            default_team_name = zip_path.stem
    duplicate_counter: Dict[str, int] = {}

    for info, path in word_infos:
        directories = path.parts[:-1]
        team_name: str | None = None
        if team_level is not None and len(directories) > team_level:
            team_name = directories[team_level]
        elif default_team_name:
            team_name = default_team_name

        sanitized_name = _sanitize_report_filename(path.name)
        display_name = _build_display_name(sanitized_name, team_name, duplicate_counter)
        persons = _extract_person_names(sanitized_name)

        entries.append(
            ZipEntry(
                identifier=str(uuid.uuid4()),
                display_name=display_name,
                archive_name=str(path),
                team_name=team_name,
                persons=persons,
                sanitized_name=sanitized_name,
            )
        )

    return entries


def _apply_team_prefixes(extract_dir: Path, entries: Dict[str, ZipEntry]) -> None:
    for entry in entries.values():
        if not entry.team_name:
            continue

        source_path = extract_dir / Path(entry.archive_name)
        if not source_path.exists():
            continue

        target_path = source_path.with_name(entry.display_name)
        if target_path != source_path:
            target_path.parent.mkdir(parents=True, exist_ok=True)
            source_path.rename(target_path)

        entry.archive_name = str(target_path.relative_to(extract_dir).as_posix())

def _create_job_state(
    job_id: str,
    email: str,
    order: List[str],
    zip_path: Path,
    entry_map: Dict[str, ZipEntry],
    zip_original_name: str,
    team_options: List[str],
) -> JobState:
    now = datetime.now(JST)
    return JobState(
        id=job_id,
        email=email,
        status="queued",
        message="処理を待機しています。",
        created_at=now,
        updated_at=now,
        order=order,
        zip_path=zip_path,
        zip_original_name=zip_original_name,
        entries=entry_map,
        team_options=team_options,
        progress_total=len(order) if order else 0,
    )


def _initialize_conversion_progress(job_id: str, entries: List[ZipEntry]) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return
        now = datetime.now(JST)
        job.conversion_order = [entry.display_name for entry in entries]
        job.conversion_statuses = {
            entry.display_name: "queued" for entry in entries
        }
        job.conversion_threads = {
            entry.display_name: index + 1 for index, entry in enumerate(entries)
        }
        job.show_conversion_progress = True
        job.updated_at = now


def _update_conversion_status(job_id: str, display_name: str, status: str) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return
        now = datetime.now(JST)
        if display_name not in job.conversion_order:
            job.conversion_order.append(display_name)
        job.conversion_statuses[display_name] = status
        if display_name not in job.conversion_threads:
            job.conversion_threads[display_name] = len(job.conversion_threads) + 1
        job.updated_at = now


def _update_job(
    job_id: str,
    *,
    status: str | None = None,
    message: str | None = None,
    merged_pdf: Path | None = None,
    progress_increment: int | None = None,
    progress_total: int | None = None,
    email_delivery_status: str | None = None,
    show_conversion_progress: bool | None = None,
) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return
        now = datetime.now(JST)
        if status:
            job.status = status
            if status == "running" and job.processing_started_at is None:
                job.processing_started_at = now
            if status in {"completed", "failed"}:
                job.processing_completed_at = now
        if message:
            job.message = message
        if merged_pdf:
            job.merged_pdf = merged_pdf
        if progress_total is not None:
            job.progress_total = max(progress_total, 0)
            if job.progress_total == 0:
                job.progress_current = 0
            elif job.progress_current > job.progress_total:
                job.progress_current = job.progress_total
        if email_delivery_status is not None:
            job.email_delivery_status = email_delivery_status
        if show_conversion_progress is not None:
            job.show_conversion_progress = show_conversion_progress
        if progress_increment:
            job.progress_current += progress_increment
            if job.progress_total > 0:
                job.progress_current = min(job.progress_current, job.progress_total)
        if status == "completed" and job.progress_total > 0:
            job.progress_current = job.progress_total
            job.show_conversion_progress = False
        if status == "failed":
            job.show_conversion_progress = False
        job.updated_at = now


def _process_job(job_id: str) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return

    try:
        _update_job(job_id, status="running", message="ZIPファイルを展開しています…")
        work_root = WORK_DIR / job_id
        extract_dir = work_root / "extracted"
        pdf_dir = work_root / "pdf"
        extract_dir.mkdir(parents=True, exist_ok=True)
        pdf_dir.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(job.zip_path) as archive:
            archive.extractall(path=extract_dir)

        _apply_team_prefixes(extract_dir, job.entries)

        ordered_entries: List[ZipEntry] = []
        for display_name in job.order:
            entry = job.entries.get(display_name)
            if entry:
                ordered_entries.append(entry)

        if not ordered_entries:
            raise RuntimeError("処理対象のドキュメントが見つかりませんでした。")

        total_steps = len(ordered_entries)
        _initialize_conversion_progress(job_id, ordered_entries)
        _update_job(
            job_id,
            progress_total=total_steps,
            message="PDF変換の準備をしています…",
        )

        report_number = _determine_report_number(job.zip_original_name, ordered_entries)
        job.report_number = report_number

        _update_job(job_id, message="WordファイルをPDFに変換中・・・")

        pdf_paths: List[Path | None] = [None] * len(ordered_entries)

        def _convert_entry(index: int, entry: ZipEntry) -> tuple[int, Path]:
            _update_job(job_id, message="WordファイルをPDFに変換中・・・")
            _update_conversion_status(job_id, entry.display_name, "running")
            source_path = extract_dir / entry.archive_name
            if not source_path.exists():
                _update_conversion_status(job_id, entry.display_name, "failed")
                raise FileNotFoundError(
                    f"展開後のファイルが見つかりません: {entry.archive_name}"
                )
            try:
                pdf_path = convert_word_to_pdf(source_path, pdf_dir)
            except Exception:
                _update_conversion_status(job_id, entry.display_name, "failed")
                raise
            _update_conversion_status(job_id, entry.display_name, "completed")
            _update_job(job_id, progress_increment=1)
            return index, pdf_path

        max_workers = max(len(ordered_entries), 1)
        with ThreadPoolExecutor(max_workers=max_workers) as conversion_pool:
            futures = {
                conversion_pool.submit(_convert_entry, index, entry): index
                for index, entry in enumerate(ordered_entries)
            }
            try:
                for future in as_completed(futures):
                    index, pdf_path = future.result()
                    pdf_paths[index] = pdf_path
            except Exception:
                for future in futures:
                    future.cancel()
                raise

        pdf_paths_resolved = [path for path in pdf_paths if path is not None]
        if len(pdf_paths_resolved) != len(ordered_entries):
            raise RuntimeError("PDF変換結果の整合性に問題があります。")

        merged_path = work_root / f"第{report_number}回報告書.pdf"
        _update_job(job_id, message="PDFファイルを結合しています…")
        merge_pdfs(pdf_paths_resolved, merged_path)

        recipient_email = (job.email or "").strip()
        should_send_email = EMAIL_CONFIG.is_configured and bool(recipient_email)

        _update_job(
            job_id,
            status="completed",
            message="PDFの結合が完了しました。",
            merged_pdf=merged_path,
            email_delivery_status="sending" if should_send_email else "",
            show_conversion_progress=False,
        )

        if should_send_email:
            def _background_email_sender() -> None:
                try:
                    send_email_with_attachment(
                        config=EMAIL_CONFIG,
                        recipient=recipient_email,
                        subject=f"第{report_number}回報告書",
                        body="",
                        attachment_path=merged_path,
                    )
                except Exception:  # noqa: BLE001
                    app.logger.exception("Failed to send email for job %s", job_id)
                    _update_job(job_id, email_delivery_status="")
                else:
                    _update_job(job_id, email_delivery_status="sent")
                finally:
                    _schedule_delayed_cleanup()

            email_thread = threading.Thread(target=_background_email_sender, daemon=True)
            email_thread.start()
        else:
            _schedule_delayed_cleanup()
    except Exception as exc:  # noqa: BLE001
        _update_job(
            job_id,
            status="failed",
            message=f"エラーが発生しました: {exc}",
            show_conversion_progress=False,
        )
    finally:
        job.zip_path.unlink(missing_ok=True)


@app.route("/", methods=["GET"])
def index() -> str:
    return render_template("index.html")


@app.route("/prepare", methods=["POST"])
def prepare_upload() -> Response | str:
    zip_file = request.files.get("zip_file")

    if not zip_file or not zip_file.filename:
        flash("アップロードするZIPファイルを選択してください。")
        return redirect(url_for("index"))

    job_id = uuid.uuid4().hex
    zip_path = UPLOAD_DIR / f"{job_id}.zip"
    zip_file.save(zip_path)

    entries = _extract_entries(zip_path, original_name=zip_file.filename)
    if not entries:
        zip_path.unlink(missing_ok=True)
        flash("アップロードされたZIPにWordファイルが見つかりませんでした。")
        return redirect(url_for("index"))

    preferences = order_manager.load_preferences()
    preference_team_options = _team_labels_from_preferences(preferences)

    entry_map: Dict[str, ZipEntry] = {entry.display_name: entry for entry in entries}
    team_sequence, team_entries = order_manager.initial_layout(list(entry_map.values()))
    ordered_display_names: List[str] = []
    team_blocks: List[Dict[str, object]] = []

    for team_key in team_sequence:
        items = team_entries.get(team_key, [])
        if not items:
            continue
        block_entries = []
        for item in items:
            ordered_display_names.append(item.display_name)
            block_entries.append(
                {
                    "display_name": item.display_name,
                    "team": item.team_name or "",
                    "persons": item.persons or [],
                }
            )
        team_blocks.append(
            {
                "key": team_key,
                "label": _team_display_label(team_key),
                "count": len(block_entries),
                "entries": block_entries,
            }
            )

    team_options = [block["label"] for block in team_blocks]
    session_team_options = preference_team_options or team_options

    default_member_sequences = {
        key: list(value)
        for key, value in preferences.member_sequences.items()
    }

    upload_sessions[job_id] = {
        "entries": entry_map,
        "team_options": session_team_options,
        "zip_filename": zip_file.filename or "",
    }

    initial_state = [
        {
            "key": block["key"],
            "label": block["label"],
            "entries": [
                {
                    "display_name": entry["display_name"],
                    "team": entry["team"],
                    "persons": entry.get("persons", []),
                }
                for entry in block["entries"]
            ],
        }
        for block in team_blocks
    ]

    return render_template(
        "order.html",
        job_id=job_id,
        ordered_display_names=ordered_display_names,
        team_blocks=team_blocks,
        initial_state=initial_state,
        default_member_sequences=default_member_sequences,
    )


def _collect_preference_teams(preferences: OrderPreferences) -> List[Dict[str, object]]:
    teams: List[Dict[str, object]] = []
    seen: set[str] = set()

    def append_team(team_key: str) -> None:
        if team_key in seen:
            return
        teams.append(
            {
                "key": team_key,
                "label": _team_display_label(team_key),
                "members": list(preferences.member_sequences.get(team_key, [])),
            }
        )
        seen.add(team_key)

    for team_key in preferences.team_sequence:
        append_team(team_key)

    for team_key in preferences.member_sequences.keys():
        append_team(team_key)

    return teams


@app.route("/default-order-editor", methods=["GET"])
def default_order_editor() -> str:
    team_param = request.args.get("team", "").strip()
    preferences = order_manager.load_preferences()
    teams = _collect_preference_teams(preferences)

    available_keys = {team["key"] for team in teams}
    initial_team_key = team_param if team_param in available_keys else None
    if not initial_team_key:
        if teams:
            initial_team_key = teams[0]["key"]
        else:
            initial_team_key = UNGROUPED_TEAM_KEY

    initial_data = {
        "teams": teams,
        "team_order": [team["key"] for team in teams],
        "initial_team_key": initial_team_key,
        "ungrouped_key": UNGROUPED_TEAM_KEY,
        "ungrouped_label": _team_display_label(UNGROUPED_TEAM_KEY),
    }

    return render_template(
        "default_order_editor.html",
        initial_data=initial_data,
    )


@app.route("/default-order-editor/save", methods=["POST"])
def save_default_member_order() -> Response:
    payload = request.get_json(silent=True) or {}
    team_key_raw = payload.get("team_key", "")
    members_raw = payload.get("members", [])

    if isinstance(team_key_raw, str):
        team_key = team_key_raw.strip()
    else:
        team_key = str(team_key_raw or "").strip()

    if not isinstance(members_raw, list):
        return jsonify({"error": "メンバー情報の形式が正しくありません。"}), 400

    normalized_members: List[str] = []
    for item in members_raw:
        if not isinstance(item, str):
            continue
        stripped = item.strip()
        if stripped and stripped not in normalized_members:
            normalized_members.append(stripped)

    if not team_key:
        return jsonify({"error": "班名を入力してください。"}), 400

    normalized_team = _normalize_team_key(team_key)

    order_manager.save_member_sequence(normalized_team, normalized_members)

    return jsonify(
        {
            "status": "ok",
            "team_key": normalized_team,
            "label": _team_display_label(normalized_team),
            "members": normalized_members,
        }
    )


@app.route("/default-order-editor/delete", methods=["POST"])
def delete_default_member_order() -> Response:
    payload = request.get_json(silent=True) or {}
    team_key_raw = payload.get("team_key", "")

    if isinstance(team_key_raw, str):
        team_key = team_key_raw.strip()
    else:
        team_key = str(team_key_raw or "").strip()

    if not team_key:
        return jsonify({"error": "班を指定してください。"}), 400

    normalized_team = _normalize_team_key(team_key)

    order_manager.delete_member_sequence(normalized_team)

    return jsonify({"status": "ok", "team_key": normalized_team})


@app.route("/api/default-order", methods=["GET"])
def api_default_order() -> Response:
    preferences = order_manager.load_preferences()
    teams = _collect_preference_teams(preferences)
    return jsonify(
        {
            "team_sequence": preferences.team_sequence,
            "member_sequences": preferences.member_sequences,
            "teams": teams,
        }
    )


@app.route("/start", methods=["POST"])
def start_processing() -> str:
    job_id = request.form.get("job_id", "")
    order_data = request.form.get("order", "")

    if not job_id or job_id not in upload_sessions:
        flash("アップロード情報を確認できませんでした。もう一度お試しください。")
        return redirect(url_for("index"))
    if not order_data:
        flash("処理を開始する前に順番を確定してください。")
        return redirect(url_for("index"))

    order = [name for name in order_data.split("|") if name]
    if not order:
        flash("ドキュメントの並び順が空です。もう一度お試しください。")
        return redirect(url_for("index"))

    session_data = upload_sessions.pop(job_id)
    entry_map = session_data.get("entries", {}) if isinstance(session_data, dict) else {}
    team_options_raw = session_data.get("team_options", []) if isinstance(session_data, dict) else []
    zip_original_name = session_data.get("zip_filename", "") if isinstance(session_data, dict) else ""

    if not isinstance(entry_map, dict) or not entry_map:
        flash("処理対象のデータを取得できませんでした。もう一度お試しください。")
        return redirect(url_for("index"))

    team_options = [str(option) for option in team_options_raw if isinstance(option, str)]
    if not team_options:
        preference_team_options = _team_labels_from_preferences(
            order_manager.load_preferences()
        )
        if preference_team_options:
            team_options = preference_team_options

    if not team_options:
        labels = {
            _team_display_label(entry.team_name or UNGROUPED_TEAM_KEY)
            for entry in entry_map.values()
        }
        team_options = sorted(labels)

    zip_original_name = str(zip_original_name) if zip_original_name else ""

    zip_path = UPLOAD_DIR / f"{job_id}.zip"
    job_state = _create_job_state(
        job_id,
        DEFAULT_RECIPIENT_EMAIL,
        order,
        zip_path,
        entry_map,
        zip_original_name or zip_path.name,
        team_options,
    )

    with jobs_lock:
        jobs[job_id] = job_state

    executor.submit(_process_job, job_id)

    return render_template("status.html", job_id=job_id)


@app.route("/download/<job_id>", methods=["GET"])
def download_merged_pdf(job_id: str) -> Response:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job or job.status != "completed" or not job.merged_pdf:
            abort(404)
        file_path = job.merged_pdf

    if not file_path.exists():
        abort(404)

    return send_file(
        file_path,
        as_attachment=True,
        download_name=file_path.name,
    )


@app.route("/status/<job_id>", methods=["GET"])
def job_status(job_id: str) -> Response:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "ジョブが見つかりません。"}), 404
        return jsonify(job.to_dict())

def get_app():
    """Trayランチャー等からインポートするための取り出し口"""
    return app

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=8000)