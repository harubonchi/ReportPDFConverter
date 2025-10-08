from __future__ import annotations

from pathlib import Path
import os
import shlex
import tempfile

import jpype
from jpype import JClass, JException

# JVM/クラスパスのデフォルト値は環境変数で上書きできるようにし、
# コンテナや Linux 環境でもそのまま動作するようにする。
ASPOSE_JAR_NAME = os.environ.get("ASPOSE_JAR_NAME", "aspose-words-20.12-jdk17.jar")
# 任意で JVM の明示パスを指定したい場合は JVM_PATH を利用する。
_ENV_JVM_PATH = os.environ.get("JVM_PATH")
_ENV_JVM_MIN_HEAP = os.environ.get("JVM_MIN_HEAP") or os.environ.get("JVM_MIN_HEAP_MB")
_ENV_JVM_MAX_HEAP = os.environ.get("JVM_MAX_HEAP") or os.environ.get("JVM_MAX_HEAP_MB")
_ENV_JVM_OPTIONS = os.environ.get("JVM_OPTIONS")
_DEFAULT_MIN_HEAP = "256m"
_DEFAULT_MAX_HEAP = "2048m"
_DEFAULT_TEMP_DIR = Path(tempfile.gettempdir()) / "aspose-words-cache"
ASPOSE_TEMP_DIR = Path(os.environ.get("ASPOSE_TEMP_DIR", str(_DEFAULT_TEMP_DIR)))
SUPPORTED_EXTENSIONS = {".doc", ".docx"}


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _get_aspose_jar() -> Path:
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
    if jpype.isJVMStarted():
        return
    jar_path = _get_aspose_jar()
    jvm_path = _jvm_path()
    # 明示JVM + クラスパス。jpype.startJVM は同一プロセス内で一度だけ呼び出す。
    heap_args = _resolve_heap_arguments()
    extra_args = shlex.split(_ENV_JVM_OPTIONS) if _ENV_JVM_OPTIONS else []
    jpype.startJVM(
        jvm_path,
        *heap_args,
        *extra_args,
        convertStrings=False,
        classpath=[str(jar_path)],
    )


def _normalize_heap_size(value: str) -> str:
    candidate = value.strip()
    if not candidate:
        raise ValueError("Empty heap size")
    lowered = candidate.lower()
    if lowered.endswith(("kb", "mb", "gb")):
        return candidate
    if lowered[-1] in {"k", "m", "g"}:
        return candidate
    if lowered[-1].isdigit():
        return f"{candidate}m"
    return candidate


def _resolve_heap_arguments() -> list[str]:
    args: list[str] = []

    def _heap_arg(prefix: str, value: str | None, default: str | None) -> None:
        candidate = value if value is not None else default
        if not candidate:
            return
        try:
            normalized = _normalize_heap_size(candidate)
        except ValueError:
            return
        args.append(f"{prefix}{normalized}")

    _heap_arg("-Xms", _ENV_JVM_MIN_HEAP, _DEFAULT_MIN_HEAP)
    _heap_arg("-Xmx", _ENV_JVM_MAX_HEAP, _DEFAULT_MAX_HEAP)
    return args


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
    _start_jvm()
    Document = JClass("com.aspose.words.Document")
    LoadOptions = JClass("com.aspose.words.LoadOptions")
    FileCorruptedException = JClass("com.aspose.words.FileCorruptedException")
    try:
        MemorySettings = JClass("com.aspose.words.MemorySettings")
        MemorySetting = JClass("com.aspose.words.MemorySetting")
    except (TypeError, RuntimeError):  # pragma: no cover - depends on Aspose version
        MemorySettings = None
        MemorySetting = None

    try:
        PdfSaveOptions = JClass("com.aspose.words.PdfSaveOptions")
    except (TypeError, RuntimeError):  # pragma: no cover - depends on Aspose version
        PdfSaveOptions = None

    temp_dir = ASPOSE_TEMP_DIR.expanduser()
    temp_dir.mkdir(parents=True, exist_ok=True)
    temp_dir = temp_dir.resolve()

    if MemorySettings:
        try:
            MemorySettings.setTempFolder(str(temp_dir))
        except (AttributeError, TypeError):  # pragma: no cover - defensive
            pass

    load_options = LoadOptions()
    load_options.setTempFolder(str(temp_dir))
    try:
        load_options.setUpdateDirtyFields(False)
    except AttributeError:  # pragma: no cover - depends on Aspose version
        pass

    if MemorySettings and MemorySetting:
        try:
            MemorySettings.setMemorySetting(MemorySetting.MEMORY_PREFERENCE)
        except (AttributeError, TypeError):  # pragma: no cover - defensive
            pass

    pdf_save_options = None
    if PdfSaveOptions:
        try:
            pdf_save_options = PdfSaveOptions()
            pdf_save_options.setTempFolder(str(temp_dir))
        except (AttributeError, TypeError):  # pragma: no cover - defensive
            pdf_save_options = None

    try:
        document = Document(str(source), load_options)
        if pdf_save_options is not None:
            document.save(str(destination), pdf_save_options)
        else:
            document.save(str(destination))
    except JException as exc:
        if isinstance(exc, FileCorruptedException):
            raise ConversionError(
                "Aspose.Words reported that the document is corrupted. "
                "Large or complex files may require additional JVM heap space."
            ) from exc
        raise


def convert_word_to_pdf(source: Path, output_dir: Path) -> Path:
    """Convert a Word document to PDF using Aspose.Words."""
    if source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ConversionError(f"Unsupported file type: {source.suffix}")
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{source.stem}.pdf"
    try:
        _convert_with_aspose(source, output_path)
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
