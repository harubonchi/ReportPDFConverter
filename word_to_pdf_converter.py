from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

from docx2pdf import convert as docx2pdf_convert


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _convert_with_docx2pdf(source: Path, destination: Path) -> None:
    docx2pdf_convert(str(source), str(destination))


LIBREOFFICE_PROFILE_DIR = Path.home() / ".config/libreoffice/4/user"


def _libreoffice_python_candidates() -> list[Path]:
    candidates: list[Path] = []
    env_value = os.environ.get("LIBREOFFICE_PYTHON")
    if env_value:
        candidates.append(Path(env_value))
    candidates.extend(
        [
            Path("/usr/lib/libreoffice/program/python"),
            Path("/usr/bin/libreoffice-python"),
            Path("/Applications/LibreOffice.app/Contents/MacOS/python"),
            Path("/Applications/LibreOffice.app/Contents/Resources/python"),
            Path("C:/Program Files/LibreOffice/program/python.exe"),
            Path("C:/Program Files (x86)/LibreOffice/program/python.exe"),
        ]
    )
    return candidates


def _find_libreoffice_python() -> Path:
    for candidate in _libreoffice_python_candidates():
        if candidate.exists() and os.access(candidate, os.X_OK):
            return candidate
    raise ConversionError(
        "LibreOffice Python executable not found. Set LIBREOFFICE_PYTHON environment variable."
    )


def _convert_with_libreoffice(source: Path, output_dir: Path) -> Path:
    script_path = Path(__file__).resolve().parent / "libreoffice_uno_converter.py"
    if not script_path.exists():
        raise ConversionError("UNO conversion script not found.")

    output_path = output_dir / (source.stem + ".pdf")
    output_path.parent.mkdir(parents=True, exist_ok=True)

    libreoffice_python = _find_libreoffice_python()

    command = [
        str(libreoffice_python),
        str(script_path),
        str(source.resolve()),
        str(output_path.resolve()),
        "--line-spacing",
        "1.15",
    ]

    LIBREOFFICE_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    command.extend(["--user-profile", str(LIBREOFFICE_PROFILE_DIR.resolve())])

    completed = subprocess.run(command, check=False, capture_output=True, text=True)
    if completed.returncode != 0:
        stderr = completed.stderr.strip()
        raise ConversionError(
            "LibreOffice conversion failed: " + (stderr or "unknown error from UNO script")
        )

    if not output_path.exists():
        raise ConversionError("LibreOffice conversion did not produce an output file.")

    return output_path


def _ensure_libreoffice_available() -> None:
    if shutil.which("soffice") is None:
        raise ConversionError("LibreOffice (soffice) is required to convert Word documents.")


def _docx2pdf_supported() -> bool:
    return sys.platform in {"win32", "cygwin", "darwin"}


def convert_word_to_pdf(source: Path, output_dir: Path) -> Path:
    """Convert a Word document to PDF and return the resulting path."""

    if source.suffix.lower() not in {".doc", ".docx"}:
        raise ConversionError(f"Unsupported file type: {source.suffix}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source.stem}.pdf"

    suffix = source.suffix.lower()

    if suffix == ".docx":
        if not _docx2pdf_supported():
            _ensure_libreoffice_available()
            return _convert_with_libreoffice(source, output_dir)

        try:
            _convert_with_docx2pdf(source, output_path)
        except NotImplementedError:
            _ensure_libreoffice_available()
            return _convert_with_libreoffice(source, output_dir)
        except Exception as exc:  # noqa: BLE001
            raise ConversionError(f"Failed to convert {source.name} to PDF: {exc}") from exc

        if not output_path.exists():
            raise ConversionError(f"Conversion completed but {output_path} was not created.")

        return output_path

    _ensure_libreoffice_available()

    try:
        return _convert_with_libreoffice(source, output_dir)
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Failed to convert {source.name} to PDF: {exc}") from exc
