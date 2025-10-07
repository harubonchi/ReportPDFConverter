from __future__ import annotations

import json
import re
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PurePosixPath
from typing import Dict, List

from dotenv import load_dotenv
from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)

from pdf_merge import merge_pdfs
from word_to_pdf_converter import convert_word_to_pdf
from email_service import EmailConfig, send_email_with_attachment


load_dotenv()


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
WORK_DIR = DATA_DIR / "work"
ORDER_FILE = DATA_DIR / "order.json"

for directory in (UPLOAD_DIR, WORK_DIR):
    directory.mkdir(parents=True, exist_ok=True)


@dataclass
class ZipEntry:
    identifier: str
    display_name: str
    archive_name: str
    team_name: str | None = None
    persons: List[str] | None = None


UNGROUPED_TEAM_KEY = "__ungrouped__"


@dataclass
class OrderPreferences:
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

class OrderManager:
    def __init__(self, storage_file: Path) -> None:
        self.storage_file = storage_file
        self._lock = threading.Lock()

    def load_preferences(self) -> OrderPreferences:
        if not self.storage_file.exists():
            return OrderPreferences.empty()
        try:
            with self.storage_file.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
        except json.JSONDecodeError:
            return OrderPreferences.empty()

        if isinstance(data, dict):
            team_sequence = [
                team
                for team in data.get("team_sequence", [])
                if isinstance(team, str)
            ]
            member_sequences: Dict[str, List[str]] = {}
            raw_members = data.get("member_sequences", {})
            if isinstance(raw_members, dict):
                for key, value in raw_members.items():
                    if isinstance(key, str) and isinstance(value, list):
                        member_sequences[key] = [
                            name for name in value if isinstance(name, str)
                        ]
            return OrderPreferences(team_sequence=team_sequence, member_sequences=member_sequences)

        if isinstance(data, list) and all(isinstance(item, str) for item in data):
            return self._from_legacy_list(data)

        return OrderPreferences.empty()

    def _from_legacy_list(self, items: List[str]) -> OrderPreferences:
        member_sequences: Dict[str, List[str]] = {}
        team_sequence: List[str] = []
        for name in items:
            team_key = UNGROUPED_TEAM_KEY
            if name.startswith("[") and "]" in name:
                prefix, _, remainder = name.partition("]")
                team_name = prefix[1:]
                team_key = team_name or UNGROUPED_TEAM_KEY
                stripped = remainder.strip()
            else:
                stripped = name
            if team_key not in team_sequence:
                team_sequence.append(team_key)
            member_sequences.setdefault(team_key, [])
            if stripped and stripped not in member_sequences[team_key]:
                member_sequences[team_key].append(stripped)
        return OrderPreferences(team_sequence=team_sequence, member_sequences=member_sequences)

    def save(self, order: List[str], entries: Dict[str, ZipEntry]) -> None:
        preferences = self._build_preferences(order, entries)
        with self._lock:
            self.storage_file.parent.mkdir(parents=True, exist_ok=True)
            with self.storage_file.open("w", encoding="utf-8") as fh:
                json.dump(preferences.to_dict(), fh, ensure_ascii=False, indent=2)

    def _build_preferences(self, order: List[str], entries: Dict[str, ZipEntry]) -> OrderPreferences:
        team_sequence: List[str] = []
        member_sequences: Dict[str, List[str]] = {}

        for display_name in order:
            entry = entries.get(display_name)
            if not entry:
                continue
            team_key = entry.team_name or UNGROUPED_TEAM_KEY
            if team_key not in team_sequence:
                team_sequence.append(team_key)
            member_list = member_sequences.setdefault(team_key, [])
            persons = entry.persons or []
            if not persons:
                continue
            for person in persons:
                if person not in member_list:
                    member_list.append(person)

        return OrderPreferences(team_sequence=team_sequence, member_sequences=member_sequences)

    def initial_layout(self, entries: List[ZipEntry]) -> tuple[List[str], Dict[str, List[ZipEntry]]]:
        preferences = self.load_preferences()
        team_map: Dict[str, List[ZipEntry]] = {}
        team_appearance: List[str] = []

        for entry in entries:
            team_key = entry.team_name or UNGROUPED_TEAM_KEY
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
            indices = [member_order.index(person) for person in persons if person in member_order]
            if indices:
                return (0, min(indices))
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
    entries: Dict[str, ZipEntry]
    merged_pdf: Path | None = None

    def to_dict(self) -> Dict[str, str]:
        return {
            "id": self.id,
            "email": self.email,
            "status": self.status,
            "message": self.message,
            "created_at": self.created_at.isoformat(),
            "updated_at": self.updated_at.isoformat(),
        }


app = Flask(__name__)
app.secret_key = "gain-report-emailer"

executor = ThreadPoolExecutor(max_workers=2)
jobs: Dict[str, JobState] = {}
jobs_lock = threading.Lock()
upload_sessions: Dict[str, Dict[str, ZipEntry]] = {}


EMAIL_CONFIG = EmailConfig.from_env()


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


def _sanitize_report_filename(original_name: str) -> str:
    path = Path(original_name)
    name = path.name

    name = re.sub(r"[,，、]", "・", name)
    name = re.sub(r"[₋_＿\s]+", " ", name)
    name = re.sub(r"報告会", "報告書", name)
    name = re.sub(r"\s+", " ", name).strip()

    if "回" in name and "報告書" not in name:
        pos = name.find("回")
        if pos != -1:
            insert_pos = pos + 1
            name = name[:insert_pos] + "報告書" + name[insert_pos:]

    if "報告書" not in name:
        name = name.replace("回 ", "回報告書 ")

    return name


def _extract_person_names(sanitized_name: str) -> List[str]:
    stem = Path(sanitized_name).stem
    if "報告書" in stem:
        _, _, remainder = stem.partition("報告書")
    else:
        remainder = stem
    remainder = remainder.strip()
    if not remainder:
        return []
    tokens = re.split(r"[・,，、\s]+", remainder)
    persons = [token.strip() for token in tokens if token.strip()]
    return persons


def _team_display_label(team_key: str) -> str:
    return "班なし" if team_key == UNGROUPED_TEAM_KEY else team_key


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


def _extract_entries(zip_path: Path) -> List[ZipEntry]:
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
    duplicate_counter: Dict[str, int] = {}

    for info, path in word_infos:
        directories = path.parts[:-1]
        team_name: str | None = None
        if team_level is not None and len(directories) > team_level:
            team_name = directories[team_level]

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

def _create_job_state(job_id: str, email: str, order: List[str], zip_path: Path, entry_map: Dict[str, ZipEntry]) -> JobState:
    now = datetime.utcnow()
    return JobState(
        id=job_id,
        email=email,
        status="queued",
        message="Job queued.",
        created_at=now,
        updated_at=now,
        order=order,
        zip_path=zip_path,
        entries=entry_map,
    )


def _update_job(job_id: str, *, status: str | None = None, message: str | None = None, merged_pdf: Path | None = None) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return
        if status:
            job.status = status
        if message:
            job.message = message
        if merged_pdf:
            job.merged_pdf = merged_pdf
        job.updated_at = datetime.utcnow()


def _process_job(job_id: str) -> None:
    with jobs_lock:
        job = jobs.get(job_id)
    if not job:
        return

    try:
        _update_job(job_id, status="running", message="Extracting ZIP archive…")
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
            raise RuntimeError("No documents found to process.")

        pdf_paths: List[Path] = []
        for entry in ordered_entries:
            _update_job(job_id, message=f"Converting {entry.display_name} to PDF…")
            source_path = extract_dir / entry.archive_name
            if not source_path.exists():
                raise FileNotFoundError(f"Extracted file missing: {entry.archive_name}")
            pdf_path = convert_word_to_pdf(source_path, pdf_dir)
            pdf_paths.append(pdf_path)

        merged_path = work_root / "merged.pdf"
        _update_job(job_id, message="Merging PDF files…")
        merge_pdfs(pdf_paths, merged_path)

        if EMAIL_CONFIG.is_configured:
            _update_job(job_id, message="Sending email with merged PDF…")
            send_email_with_attachment(
                config=EMAIL_CONFIG,
                recipient=job.email,
                subject="Merged report",
                body=(
                    "The merged PDF report you requested is attached."
                ),
                attachment_path=merged_path,
            )
        else:
            raise RuntimeError("Email configuration is incomplete. Please update environment variables.")

        order_manager.save(job.order, job.entries)
        _update_job(job_id, status="completed", message="Report emailed successfully.", merged_pdf=merged_path)
    except Exception as exc:  # noqa: BLE001
        _update_job(job_id, status="failed", message=str(exc))
    finally:
        job.zip_path.unlink(missing_ok=True)


@app.route("/", methods=["GET"])
def index() -> str:
    preferences = order_manager.load_preferences()
    saved_order: List[str] = []
    for team_key in preferences.team_sequence:
        members = preferences.member_sequences.get(team_key, [])
        if not members:
            continue
        label = _team_display_label(team_key)
        for person in members:
            if team_key == UNGROUPED_TEAM_KEY:
                saved_order.append(person)
            else:
                saved_order.append(f"[{label}] {person}")
    return render_template("index.html", saved_order=saved_order)


@app.route("/prepare", methods=["POST"])
def prepare_upload() -> Response | str:
    email = request.form.get("email", "").strip()
    zip_file = request.files.get("zip_file")

    if not email:
        flash("Email address is required.")
        return redirect(url_for("index"))
    if not zip_file or not zip_file.filename:
        flash("Please select a ZIP file to upload.")
        return redirect(url_for("index"))

    job_id = uuid.uuid4().hex
    zip_path = UPLOAD_DIR / f"{job_id}.zip"
    zip_file.save(zip_path)

    entries = _extract_entries(zip_path)
    if not entries:
        zip_path.unlink(missing_ok=True)
        flash("The uploaded ZIP does not contain any Word documents.")
        return redirect(url_for("index"))

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

    upload_sessions[job_id] = entry_map

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
        email=email,
        ordered_display_names=ordered_display_names,
        team_blocks=team_blocks,
        initial_state=initial_state,
    )


@app.route("/start", methods=["POST"])
def start_processing() -> str:
    job_id = request.form.get("job_id", "")
    email = request.form.get("email", "").strip()
    order_data = request.form.get("order", "")

    if not job_id or job_id not in upload_sessions:
        flash("Unable to locate the upload session. Please try again.")
        return redirect(url_for("index"))
    if not email:
        flash("Email address is required.")
        return redirect(url_for("index"))
    if not order_data:
        flash("Please confirm the document order before starting.")
        return redirect(url_for("index"))

    order = [name for name in order_data.split("|") if name]
    if not order:
        flash("Document order is empty. Please try again.")
        return redirect(url_for("index"))

    entry_map = upload_sessions.pop(job_id)
    zip_path = UPLOAD_DIR / f"{job_id}.zip"
    job_state = _create_job_state(job_id, email, order, zip_path, entry_map)

    with jobs_lock:
        jobs[job_id] = job_state

    executor.submit(_process_job, job_id)

    return render_template("status.html", job_id=job_id)


@app.route("/status/<job_id>", methods=["GET"])
def job_status(job_id: str) -> Response:
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job not found."}), 404
        return jsonify(job.to_dict())


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)