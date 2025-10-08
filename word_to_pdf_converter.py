from __future__ import annotations

from pathlib import Path
import os
import jpype
from jpype import JClass

# === あなたのホストJDKを固定指定（例） ===
JAVA_HOME = r"C:\Program Files\Java\jdk-18.0.2"

# 正規の Aspose.Words for Java の JAR を同ディレクトリに置く
ASPOSE_JAR_NAME = "aspose-words-20.12-jdk17.jar"
SUPPORTED_EXTENSIONS = {".doc", ".docx"}


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _get_aspose_jar() -> Path:
    jar_path = Path(__file__).with_name(ASPOSE_JAR_NAME).resolve()
    if not jar_path.exists():
        raise ConversionError(f"Aspose Words JAR not found at {jar_path}.")
    return jar_path


def _jvm_path() -> str:
    jvm = Path(JAVA_HOME, "bin", "server", "jvm.dll")
    if not jvm.exists():
        raise ConversionError(f"jvm.dll not found at {jvm}. Check JAVA_HOME.")
    # プロセス内だけ JAVA_HOME/PATH を整える（念のため）
    os.environ["JAVA_HOME"] = JAVA_HOME
    os.environ["PATH"] = str(Path(JAVA_HOME, "bin")) + os.pathsep + os.environ.get("PATH", "")
    return str(jvm)


def _start_jvm() -> None:
    if jpype.isJVMStarted():
        return
    jar_path = _get_aspose_jar()
    jvm_path = _jvm_path()
    # 明示JVM + クラスパス
    jpype.startJVM(jvm_path, f"-Djava.class.path={str(jar_path)}")

def ensure_jvm_started() -> None:
    """Ensure that the JVM is started.

    Flaskアプリから利用するとき、バックグラウンドスレッド内でJVMを起動すると
    Windows環境で失敗する場合があるため、メインスレッドから明示的に起動できる
    ように公開する。
    """

    _start_jvm()


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
    document = Document(str(source))
    document.save(str(destination))


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
    _test_convert_input_docx()
