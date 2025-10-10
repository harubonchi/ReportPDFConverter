"""Utilities for merging multiple PDF files into a single document."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PyPDF2 import PdfMerger


def merge_pdfs(pdf_paths: Iterable[Path], output_path: Path) -> Path:
    """Merge ``pdf_paths`` in order and write the result to ``output_path``."""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    merger = PdfMerger()
    try:
        for pdf in pdf_paths:
            merger.append(str(Path(pdf)))
        with output_path.open("wb") as fh:
            merger.write(fh)
    finally:
        merger.close()
    return output_path
