"""Conversion utilities turning Word documents into PDF files on Windows.

This module is hardened for concurrent usage:
- Serializes access to Microsoft Word across threads/processes to avoid
  COM re-entrancy failures (e.g. ExportAsFixedFormat errors).
- Initializes COM per-call, and always tears it down, to prevent orphaned
  WINWORD.EXE instances and cross-thread COM issues.
"""

from __future__ import annotations

import platform
import os
import time
import tempfile
from pathlib import Path

SUPPORTED_EXTENSIONS = {".doc", ".docx"}
_IS_WINDOWS = platform.system().lower() == "windows"

_LOCKFILE_PATH = Path(tempfile.gettempdir()) / "ReportPDFConverter.word.lock"


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


# Note: 以前は Word の COM 競合回避のためにクロスプロセス/スレッドのロックを行っていましたが、
# 並列変換の要件に合わせ、ロックは撤廃しました。各スレッドが個別に Word インスタンスを
# 起動し、COM の初期化/終了を確実に行うことで並列動作に対応します。


def _convert_with_win32com(source: Path, destination: Path) -> None:  # pragma: no cover - requires Windows
    """Convert ``source`` to ``destination`` using Microsoft Word automation.

    この関数はスレッド毎に独立した Word インスタンスを生成し、COM をスレッド内で
    初期化/終了します。これにより複数スレッドからの並列変換に対応します。
    """

    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover - depends on environment
        raise ConversionError("pywin32 is required for Word to PDF conversion on Windows.") from exc

    # Retry transient COM errors when Word is busy.
    RPC_E_CALL_REJECTED = -2147418111
    RPC_S_SERVER_UNAVAILABLE = -2147023174

    def _attempt_once() -> None:
        pythoncom.CoInitialize()
        try:
            word = win32com.client.gencache.EnsureDispatch("Word.Application")
            try:
                word.Visible = False
                # 変換の阻害となるダイアログを抑止
                try:
                    word.DisplayAlerts = 0  # wdAlertsNone
                except Exception:
                    pass

                # 出力先の準備（既存ファイルは削除）
                destination.parent.mkdir(parents=True, exist_ok=True)
                try:
                    if destination.exists():
                        destination.unlink(missing_ok=True)  # type: ignore[arg-type]
                except Exception:
                    pass

                doc = None
                try:
                    doc = word.Documents.Open(
                        str(source),
                        ReadOnly=True,
                        ConfirmConversions=False,
                        AddToRecentFiles=False,
                    )
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(destination),
                        ExportFormat=17,          # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,            # wdExportOptimizeForPrint
                        Range=0,                  # wdExportAllDocument
                        Item=0,                   # wdExportDocumentContent
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1,        # wdExportCreateNoBookmarks
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                    )
                finally:
                    if doc is not None:
                        try:
                            doc.Close(False)
                        except Exception:
                            pass
            finally:
                try:
                    word.Quit()
                except Exception:
                    pass
        finally:
            pythoncom.CoUninitialize()

    # Perform attempts with exponential backoff for transient COM errors.
    max_attempts = 3
    attempt = 0
    while True:
        attempt += 1
        try:
            _attempt_once()
            break
        except Exception as exc:  # noqa: BLE001
            # Inspect HRESULT if available (pywintypes.com_error or pythoncom.com_error)
            hresult = None
            if hasattr(exc, "hresult"):
                hresult = getattr(exc, "hresult")
            elif getattr(exc, "args", None):
                try:
                    hresult = int(exc.args[0])
                except Exception:
                    hresult = None

            if (
                attempt < max_attempts
                and hresult in {RPC_E_CALL_REJECTED, RPC_S_SERVER_UNAVAILABLE}
            ):
                time.sleep(1.0 * (2 ** (attempt - 1)))
                continue
            raise


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
