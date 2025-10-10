"""Conversion utilities turning Word documents into PDF files on Windows."""

from __future__ import annotations

import platform
from pathlib import Path

SUPPORTED_EXTENSIONS = {".doc", ".docx"}
_IS_WINDOWS = platform.system().lower() == "windows"


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _convert_with_win32com(source: Path, destination: Path) -> None:  # pragma: no cover - requires Windows
    """Convert ``source`` to ``destination`` using Microsoft Word automation."""

    try:
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on environment
        raise ConversionError("pywin32 is required for Word to PDF conversion on Windows.") from exc

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = None
    try:
        doc = word.Documents.Open(str(source), ReadOnly=True)
        doc.ExportAsFixedFormat(
            OutputFileName=str(destination),
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:  # pragma: no cover - depends on Word COM behaviour
                pass
        word.Quit()


def convert_word_to_pdf(source: Path, output_dir: Path) -> Path:
    """Convert a Word document to PDF and return the output path."""
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ConversionError(f"Unsupported file type: {source.suffix}")
    if not _IS_WINDOWS:
        raise ConversionError(
            "Word to PDF conversion is supported only on Windows with Microsoft Word installed."
        )
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source.stem}.pdf"
    try:
        _convert_with_win32com(source, output_path)
    except ConversionError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Failed to convert {source.name} to PDF: {exc}") from exc
    if not output_path.exists():
        raise ConversionError(f"Conversion completed but {output_path} was not created.")
    return output_path


if __name__ == "__main__":
    raise SystemExit(
        "This module is intended to be imported and used within the application. "
        "Run the Flask app instead of executing this file directly."
    )
