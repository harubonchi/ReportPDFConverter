"""Conversion utilities turning Word documents into PDF files."""

from __future__ import annotations

import os
import platform
from pathlib import Path

import jpype
from jpype import JClass

# JVM/クラスパスのデフォルト値は環境変数で上書きできるようにし、
# コンテナや Linux 環境でもそのまま動作するようにする。
ASPOSE_JAR_NAME = os.environ.get("ASPOSE_JAR_NAME", "aspose-words-22.12-jdk17-unlocked.jar")
# 任意で JVM の明示パスを指定したい場合は JVM_PATH を利用する。
_ENV_JVM_PATH = os.environ.get("JVM_PATH")
SUPPORTED_EXTENSIONS = {".doc", ".docx"}
_IS_WINDOWS = platform.system().lower() == "windows"


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _get_aspose_jar() -> Path:
    """Return the path to the Aspose Words JAR, validating its existence."""

    jar_path = Path(__file__).with_name(ASPOSE_JAR_NAME).resolve()
    if not jar_path.exists():
        raise ConversionError(f"Aspose Words JAR not found at {jar_path}.")
    return jar_path


def _jvm_path() -> str:
    """Resolve the JVM shared library path in a cross-platform manner."""
    if _ENV_JVM_PATH:
        jvm = Path(_ENV_JVM_PATH).expanduser().resolve()
        if not jvm.exists():
            raise ConversionError(f"Specified JVM_PATH does not exist: {jvm}")
        return str(jvm)

    try:
        return jpype.getDefaultJVMPath()
    except (jpype.JVMNotFoundException, OSError) as exc:  # pragma: no cover - depends on runtime
        raise ConversionError(
            "Unable to locate the JVM shared library. Set JAVA_HOME or JVM_PATH environment variables."
        ) from exc


def _start_jvm() -> None:
    """Initialise the JVM once per process."""

    if jpype.isJVMStarted():
        return
    jar_path = _get_aspose_jar()
    jvm_path = _jvm_path()
    # 明示JVM + クラスパス。jpype.startJVM は同一プロセス内で一度だけ呼び出す。
    jpype.startJVM(jvm_path, convertStrings=False, classpath=[str(jar_path)])


def _java_diagnostics() -> None:
    """JVM/Javaの基本情報を表示（自己診断用）。"""
    System = JClass("java.lang.System")
    props = {
        "java.version": System.getProperty("java.version"),
        "java.vendor": System.getProperty("java.vendor"),
        "java.vm.name": System.getProperty("java.vm.name"),
        "java.vm.version": System.getProperty("java.vm.version"),
        "os.name": System.getProperty("os.name"),
        "os.arch": System.getProperty("os.arch"),
    }
    print("=== Java Diagnostics ===")
    for k, v in props.items():
        print(f"{k}: {v}")
    print("========================")


def _convert_with_aspose(source: Path, destination: Path) -> None:
    """Convert ``source`` to ``destination`` using the Aspose Java API."""

    _start_jvm()
    Document = JClass("com.aspose.words.Document")
    document = Document(str(source))
    document.save(str(destination))


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
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source.stem}.pdf"
    try:
        converter = _convert_with_win32com if _IS_WINDOWS else _convert_with_aspose
        converter(source, output_path)
    except ConversionError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise ConversionError(f"Failed to convert {source.name} to PDF: {exc}") from exc
    if not output_path.exists():
        raise ConversionError(f"Conversion completed but {output_path} was not created.")
    return output_path


# -----------------------------
# 自己診断スモークテスト + input.docx 実変換
# -----------------------------
def _self_test() -> None:
    """JVM起動→Java環境表示→空PDF生成で検証."""
    _start_jvm()
    _java_diagnostics()

    # Aspose で空ドキュメント→PDF保存（最小I/O）
    Document = JClass("com.aspose.words.Document")
    DocumentBuilder = JClass("com.aspose.words.DocumentBuilder")
    tmp_pdf = Path(__file__).with_name("_selftest.pdf").resolve()

    doc = Document()
    builder = DocumentBuilder(doc)
    builder.writeln("Aspose.Words self-test OK.")
    doc.save(str(tmp_pdf))

    size = tmp_pdf.stat().st_size if tmp_pdf.exists() else 0
    print(f"Self-test PDF written: {tmp_pdf} ({size} bytes)")
    if size <= 0:
        raise ConversionError("Self-test failed: PDF was not created or is empty.")


def _test_convert_input_docx() -> None:
    """プロジェクト直下の input.docx を output.pdf に変換して検証."""
    source = Path(__file__).with_name("input.docx").resolve()
    if not source.exists():
        raise ConversionError(f"Test file not found: {source}")

    out_dir = source.parent  # 同ディレクトリに出力
    pdf_path = convert_word_to_pdf(source, out_dir)
    size = pdf_path.stat().st_size if pdf_path.exists() else 0
    print(f"Converted: {pdf_path} ({size} bytes)")
    if size <= 0:
        raise ConversionError("Conversion test failed: output.pdf is empty.")


if __name__ == "__main__":
    # 1) JVM/Aspose 自己診断
    _self_test()
    # 2) input.docx → output.pdf 実変換テスト
    _test_convert_input_docx()