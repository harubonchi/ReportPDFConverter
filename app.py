from __future__ import annotations

import json
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
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


class OrderManager:
    def __init__(self, storage_file: Path) -> None:
        self.storage_file = storage_file
        self._lock = threading.Lock()

    def load(self) -> List[str]:
        if not self.storage_file.exists():
            return []
        try:
            with self.storage_file.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
            if isinstance(data, list) and all(isinstance(item, str) for item in data):
                return data
        except json.JSONDecodeError:
            pass
        return []

    def save(self, order: List[str]) -> None:
        with self._lock:
            self.storage_file.parent.mkdir(parents=True, exist_ok=True)
            with self.storage_file.open("w", encoding="utf-8") as fh:
                json.dump(order, fh, ensure_ascii=False, indent=2)

    def initial_order(self, file_names: List[str]) -> List[str]:
        stored = self.load()
        ordered: List[str] = [name for name in stored if name in file_names]
        for name in sorted(file_names):
            if name not in ordered:
                ordered.append(name)
        return ordered


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


def _extract_entries(zip_path: Path) -> List[ZipEntry]:
    entries: List[ZipEntry] = []
    duplicate_counter: Dict[str, int] = {}
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            suffix = Path(info.filename).suffix.lower()
            if suffix not in {".doc", ".docx"}:
                continue
            base_name = Path(info.filename).name
            duplicate_counter[base_name] = duplicate_counter.get(base_name, 0) + 1
            display_name = base_name
            if duplicate_counter[base_name] > 1:
                display_name = f"{base_name} ({duplicate_counter[base_name]})"
            entries.append(
                ZipEntry(
                    identifier=str(uuid.uuid4()),
                    display_name=display_name,
                    archive_name=info.filename,
                )
            )
    return entries


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

        order_manager.save(job.order)
        _update_job(job_id, status="completed", message="Report emailed successfully.", merged_pdf=merged_path)
    except Exception as exc:  # noqa: BLE001
        _update_job(job_id, status="failed", message=str(exc))
    finally:
        job.zip_path.unlink(missing_ok=True)


@app.route("/", methods=["GET"])
def index() -> str:
    saved_order = order_manager.load()
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
    ordered_display_names = order_manager.initial_order([entry.display_name for entry in entries])
    upload_sessions[job_id] = entry_map

    return render_template(
        "order.html",
        job_id=job_id,
        email=email,
        ordered_display_names=ordered_display_names,
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