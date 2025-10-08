"""Utilities for preprocessing DOCX files before PDF conversion.

This module focuses on reducing the footprint of embedded images so that the
subsequent Word -> PDF conversion is less likely to run out of memory.  Large
images are resized and recompressed before the temporary DOCX is handed to the
converter.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
import math
import zipfile

from PIL import Image


_MEDIA_PREFIX = "word/media/"


@dataclass
class _OptimizationSettings:
    """Configuration for image optimisation."""

    max_pixels: int = 5_000_000
    max_bytes: int = 6 * 1024 * 1024
    jpeg_quality: int = 85


def _should_optimize(data: bytes, settings: _OptimizationSettings) -> bool:
    if len(data) > settings.max_bytes:
        return True
    try:
        with Image.open(BytesIO(data)) as image:
            image.load()
            return image.width * image.height > settings.max_pixels
    except Exception:  # noqa: BLE001 - image decoding errors fall back
        return False


def _resize_dimensions(width: int, height: int, max_pixels: int) -> tuple[int, int]:
    total_pixels = width * height
    if total_pixels <= max_pixels:
        return width, height
    scale = math.sqrt(max_pixels / total_pixels)
    resized_width = max(1, int(width * scale))
    resized_height = max(1, int(height * scale))
    return resized_width, resized_height


def _optimise_image_bytes(data: bytes, settings: _OptimizationSettings) -> bytes:
    try:
        with Image.open(BytesIO(data)) as image:
            image.load()
            target_size = _resize_dimensions(image.width, image.height, settings.max_pixels)
            if target_size != (image.width, image.height):
                image = image.resize(target_size, Image.LANCZOS)

            buffer = BytesIO()
            format_hint = (image.format or "PNG").upper()
            save_kwargs: dict[str, object]
            if format_hint in {"JPEG", "JPG"}:
                save_kwargs = {"format": "JPEG", "quality": settings.jpeg_quality, "optimize": True}
                if image.mode not in {"L", "RGB"}:
                    image = image.convert("RGB")
            elif format_hint == "PNG":
                save_kwargs = {"format": "PNG", "optimize": True, "compress_level": 9}
            else:
                # Unsupported or uncommon formats are converted to PNG for better compatibility.
                if image.mode not in {"RGB", "RGBA", "L"}:
                    image = image.convert("RGBA" if "A" in image.getbands() else "RGB")
                save_kwargs = {"format": "PNG", "optimize": True, "compress_level": 9}

            image.save(buffer, **save_kwargs)
            new_data = buffer.getvalue()
    except Exception:  # noqa: BLE001 - in case Pillow cannot process the image
        return data

    # Only use the optimised payload if it is actually smaller.
    if len(new_data) < len(data):
        return new_data
    return data


def preprocess_docx_images(source: Path, destination: Path) -> Path:
    """Create a copy of ``source`` with oversized images downscaled.

    Parameters
    ----------
    source:
        Path to the original DOCX file.
    destination:
        Path where the preprocessed DOCX should be written.  The parent
        directory must exist.

    Returns
    -------
    Path
        Path to the preprocessed DOCX file.  ``destination`` is always
        returned, even when no images were touched.
    """

    if source.suffix.lower() != ".docx":
        raise ValueError("preprocess_docx_images only supports .docx files")

    settings = _OptimizationSettings()

    with zipfile.ZipFile(source, "r") as src_zip, zipfile.ZipFile(destination, "w") as dst_zip:
        for item in src_zip.infolist():
            data = src_zip.read(item.filename)
            if item.filename.startswith(_MEDIA_PREFIX) and _should_optimize(data, settings):
                data = _optimise_image_bytes(data, settings)
            dst_zip.writestr(item, data)

    return destination