"""Utilities for LibreOffice UNO based conversion with fixed line spacing."""
from __future__ import annotations

import argparse
import contextlib
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Iterable

import uno
from com.sun.star.beans import PropertyValue  # type: ignore[attr-defined]
from com.sun.star.connection import NoConnectException  # type: ignore[attr-defined]
from com.sun.star.lang import DisposedException  # type: ignore[attr-defined]
from com.sun.star.style.LineSpacingMode import PROP as LINE_SPACING_PROP  # type: ignore[attr-defined]


def _create_property(name: str, value: object) -> PropertyValue:
    prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    prop.Name = name
    prop.Value = value
    return prop


def _create_line_spacing(height: int) -> object:
    spacing = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    spacing.Mode = LINE_SPACING_PROP
    spacing.Height = height
    return spacing


class LibreOfficeSession:
    def __init__(self, process: subprocess.Popen[bytes]) -> None:
        self.process = process

    def terminate(self) -> None:
        if self.process.poll() is None:
            self.process.terminate()
            with contextlib.suppress(subprocess.TimeoutExpired):
                self.process.wait(timeout=10)


def _start_office(pipe_name: str, user_profile: Path | None) -> subprocess.Popen[bytes]:
    command: list[str] = [
        "soffice",
        "--headless",
        "--nologo",
        "--nodefault",
        "--nofirststartwizard",
        f"--accept=pipe,name={pipe_name};urp;",
    ]
    if user_profile is not None:
        command.insert(1, f"-env:UserInstallation={user_profile.resolve().as_uri()}")
    return subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def _connect_to_office(pipe_name: str, timeout: float = 30.0):  # type: ignore[no-untyped-def]
    start = time.time()
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    while time.time() - start < timeout:
        try:
            return resolver.resolve(
                f"uno:pipe,name={pipe_name};urp;StarOffice.ComponentContext"
            )
        except (NoConnectException, DisposedException):
            time.sleep(0.5)
    raise RuntimeError("Failed to connect to LibreOffice UNO pipe within timeout.")


def _iterate_paragraphs_from_text(text) -> Iterable[object]:  # type: ignore[no-untyped-def]
    enum = text.createEnumeration()
    while enum.hasMoreElements():
        element = enum.nextElement()
        if element.supportsService("com.sun.star.text.Paragraph"):
            yield element
        elif element.supportsService("com.sun.star.text.TextTable"):
            for cell_name in element.getCellNames():
                cell = element.getCellByName(cell_name)
                yield from _iterate_paragraphs_from_text(cell.Text)


def _apply_line_spacing(document, spacing_value: int) -> None:  # type: ignore[no-untyped-def]
    spacing_struct = _create_line_spacing(spacing_value)

    style_families = document.getStyleFamilies()
    if style_families.hasByName("ParagraphStyles"):
        paragraph_styles = style_families.getByName("ParagraphStyles")
        for style_name in paragraph_styles.getElementNames():
            style = paragraph_styles.getByName(style_name)
            if hasattr(style, "ParaLineSpacing"):
                style.ParaLineSpacing = _create_line_spacing(spacing_value)

    for paragraph in _iterate_paragraphs_from_text(document.Text):
        paragraph.ParaLineSpacing = spacing_struct

    if hasattr(document, "getTextFrames"):
        text_frames = document.getTextFrames()
        for frame_name in text_frames.getElementNames():
            frame = text_frames.getByName(frame_name)
            if hasattr(frame, "Text"):
                for paragraph in _iterate_paragraphs_from_text(frame.Text):
                    paragraph.ParaLineSpacing = _create_line_spacing(spacing_value)

    if hasattr(document, "getFootnotes"):
        footnotes = document.getFootnotes()
        for footnote in footnotes:
            text = getattr(footnote, "Text", None)
            if text is not None:
                for paragraph in _iterate_paragraphs_from_text(text):
                    paragraph.ParaLineSpacing = _create_line_spacing(spacing_value)

    if hasattr(document, "getEndnotes"):
        endnotes = document.getEndnotes()
        for endnote in endnotes:
            text = getattr(endnote, "Text", None)
            if text is not None:
                for paragraph in _iterate_paragraphs_from_text(text):
                    paragraph.ParaLineSpacing = _create_line_spacing(spacing_value)


def _convert_document(ctx, source: Path, destination: Path, spacing_value: int) -> None:  # type: ignore[no-untyped-def]
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    load_props = (_create_property("Hidden", True),)

    source_url = uno.systemPathToFileUrl(str(source.resolve()))
    document = desktop.loadComponentFromURL(source_url, "_blank", 0, load_props)
    try:
        _apply_line_spacing(document, spacing_value)
        output_url = uno.systemPathToFileUrl(str(destination.resolve()))
        pdf_props = (_create_property("FilterName", "writer_pdf_Export"),)
        document.storeToURL(output_url, pdf_props)
    finally:
        with contextlib.suppress(Exception):  # noqa: BLE001
            document.close(True)
        with contextlib.suppress(Exception):  # noqa: BLE001
            document.dispose()


def convert_with_fixed_line_spacing(
    source: Path, destination: Path, user_profile: Path | None, spacing_ratio: float
) -> None:
    spacing_value = int(round(spacing_ratio * 100))
    pipe_name = f"uno_pipe_{os.getpid()}"
    process = _start_office(pipe_name, user_profile)
    session = LibreOfficeSession(process)
    try:
        ctx = _connect_to_office(pipe_name)
        _convert_document(ctx, source, destination, spacing_value)
    finally:
        session.terminate()


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convert a document to PDF using LibreOffice UNO with fixed line spacing."
    )
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--user-profile", type=Path, default=None)
    parser.add_argument("--line-spacing", type=float, default=1.15)
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    try:
        convert_with_fixed_line_spacing(
            args.source,
            args.destination,
            args.user_profile,
            args.line_spacing,
        )
    except Exception as exc:  # noqa: BLE001
        print(f"Failed to convert document: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())