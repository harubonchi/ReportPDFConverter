from __future__ import annotations

from pathlib import Path

import jpype
from jpype import JClass, JVMNotFoundException, getDefaultJVMPath


ASPOSE_JAR_NAME = "aspose-words-20.12-jdk17.jar"
SUPPORTED_EXTENSIONS = {".doc", ".docx"}


class ConversionError(RuntimeError):
    """Raised when a Word document cannot be converted to PDF."""


def _get_aspose_jar() -> Path:
    jar_path = Path(__file__).with_name(ASPOSE_JAR_NAME)
    if not jar_path.exists():
        raise ConversionError(
            f"Aspose Words JAR not found at {jar_path}. Ensure it is available in the project directory."
        )
    return jar_path


def _start_jvm() -> None:
    if jpype.isJVMStarted():
        return

    jar_path = _get_aspose_jar()

    try:
        jvm_path = getDefaultJVMPath()
    except (JVMNotFoundException, FileNotFoundError) as exc:  # pragma: no cover - depends on env
        raise ConversionError(
            "Java runtime not found. Ensure JAVA_HOME is configured inside the container."
        ) from exc

    jpype.startJVM(jvm_path, f"-Djava.class.path={jar_path}")


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
