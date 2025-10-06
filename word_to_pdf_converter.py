from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from docx2pdf import convert as docx2pdf_convert


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _convert_with_docx2pdf(source: Path, destination: Path) -> None:
    docx2pdf_convert(str(source), str(destination))


def _convert_with_libreoffice(source: Path, output_dir: Path) -> Path:
    command = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(source),
    ]
    completed = subprocess.run(command, check=False, capture_output=True)
    if completed.returncode != 0:
        raise ConversionError(
            "LibreOffice conversion failed: " + completed.stderr.decode("utf-8", errors="ignore")
        )
    output_path = output_dir / (source.stem + ".pdf")
    if not output_path.exists():
        raise ConversionError("LibreOffice conversion did not produce an output file.")
    return output_path


def convert_word_to_pdf(source: Path, output_dir: Path) -> Path:
    """Convert a Word document to PDF and return the resulting path."""

    if source.suffix.lower() not in {".doc", ".docx"}:
        raise ConversionError(f"Unsupported file type: {source.suffix}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source.stem}.pdf"

    try:
        if source.suffix.lower() == ".docx":
            _convert_with_docx2pdf(source, output_path)
        else:
            # For legacy .doc files attempt to use LibreOffice if available.
            if shutil.which("soffice") is None:
                raise ConversionError("LibreOffice (soffice) is required to convert .doc files.")
            return _convert_with_libreoffice(source, output_dir)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Failed to convert {source.name} to PDF: {exc}") from exc

    if not output_path.exists():
        raise ConversionError(f"Conversion completed but {output_path} was not created.")

    return output_path