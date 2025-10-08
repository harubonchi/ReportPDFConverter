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
    target_max_bytes: int = 100 * 1024
    resize_step: float = 0.85
    min_dimension: int = 1

    def jpeg_quality_sequence(self) -> tuple[int, ...]:
        """Return descending JPEG quality levels used during optimisation."""

        qualities = [self.jpeg_quality]
        # Gradually reduce the quality, ensuring it eventually bottoms out.
        for candidate in (75, 65, 55, 45, 35, 30, 25, 20, 15, 10):
            if candidate < qualities[-1]:
                qualities.append(candidate)
        return tuple(dict.fromkeys(qualities))


def _should_optimize(data: bytes, settings: _OptimizationSettings) -> bool:
    if len(data) > settings.max_bytes or len(data) > settings.target_max_bytes:
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


def _save_as_jpeg(image: Image.Image, quality: int) -> bytes:
    buffer = BytesIO()
    image.save(buffer, format="JPEG", quality=quality, optimize=True)
    return buffer.getvalue()


def _save_as_png(image: Image.Image) -> bytes:
    buffer = BytesIO()
    image.save(buffer, format="PNG", optimize=True, compress_level=9)
    return buffer.getvalue()


def _ensure_minimum_resize(previous_size: tuple[int, int], settings: _OptimizationSettings) -> tuple[int, int]:
    width, height = previous_size
    new_width = max(settings.min_dimension, int(width * settings.resize_step))
    new_height = max(settings.min_dimension, int(height * settings.resize_step))
    if (new_width, new_height) == previous_size and (width > settings.min_dimension or height > settings.min_dimension):
        if width > settings.min_dimension:
            new_width = max(settings.min_dimension, width - 1)
        if height > settings.min_dimension:
            new_height = max(settings.min_dimension, height - 1)
    return new_width, new_height


def _compress_jpeg_to_target(image: Image.Image, settings: _OptimizationSettings) -> bytes:
    if image.mode not in {"L", "RGB"}:
        image = image.convert("RGB")
    target = settings.target_max_bytes
    best_data: bytes | None = None
    candidate = image

    while True:
        for quality in settings.jpeg_quality_sequence():
            data = _save_as_jpeg(candidate, quality)
            if len(data) <= target:
                return data
            if best_data is None or len(data) < len(best_data):
                best_data = data

        if candidate.width <= settings.min_dimension and candidate.height <= settings.min_dimension:
            return best_data if best_data is not None else _save_as_jpeg(candidate, settings.jpeg_quality)

        new_size = _ensure_minimum_resize(candidate.size, settings)
        if new_size == candidate.size:
            return best_data if best_data is not None else _save_as_jpeg(candidate, settings.jpeg_quality)
        candidate = candidate.resize(new_size, Image.LANCZOS)


def _compress_png_to_target(image: Image.Image, settings: _OptimizationSettings, *, preserve_alpha: bool) -> bytes:
    if preserve_alpha:
        desired_mode = "RGBA" if "A" in image.getbands() else "RGB"
    else:
        desired_mode = "RGB"
    if image.mode != desired_mode:
        image = image.convert(desired_mode)

    target = settings.target_max_bytes
    best_data: bytes | None = None
    candidate = image

    while True:
        data = _save_as_png(candidate)
        if len(data) <= target:
            return data
        if best_data is None or len(data) < len(best_data):
            best_data = data

        if candidate.width <= settings.min_dimension and candidate.height <= settings.min_dimension:
            return best_data if best_data is not None else data

        new_size = _ensure_minimum_resize(candidate.size, settings)
        if new_size == candidate.size:
            return best_data if best_data is not None else data
        candidate = candidate.resize(new_size, Image.LANCZOS)


def _optimise_image_bytes(data: bytes, settings: _OptimizationSettings) -> bytes:
    try:
        with Image.open(BytesIO(data)) as image:
            image.load()
            target_size = _resize_dimensions(image.width, image.height, settings.max_pixels)
            if target_size != (image.width, image.height):
                image = image.resize(target_size, Image.LANCZOS)

            format_hint = (image.format or "PNG").upper()
            has_alpha = "A" in image.getbands()

            if format_hint in {"JPEG", "JPG"} and not has_alpha:
                new_data = _compress_jpeg_to_target(image, settings)
            elif format_hint == "PNG" and has_alpha:
                new_data = _compress_png_to_target(image, settings, preserve_alpha=True)
            else:
                if not has_alpha:
                    jpeg_data = _compress_jpeg_to_target(image, settings)
                else:
                    jpeg_data = None

                png_data = _compress_png_to_target(image, settings, preserve_alpha=has_alpha)

                if jpeg_data is None:
                    new_data = png_data
                else:
                    new_data = jpeg_data if len(jpeg_data) <= len(png_data) else png_data
    except Exception:  # noqa: BLE001 - in case Pillow cannot process the image
        return data

    if len(data) > settings.target_max_bytes:
        return new_data
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