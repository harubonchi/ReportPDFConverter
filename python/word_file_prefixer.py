# coding: UTF-8
"""Utility script for prefixing Word files with their group names.

This module is responsible for traversing an extracted zip directory that
contains sub-directories per group (e.g. "R班", "N班").  Every Word file in a
group directory will receive a prefix based on the group name such as
"【R班】".  The script can also reorder the group directories by adding
zero-padded numeric prefixes so that packaging the directory again preserves
the desired order.

Usage example
-------------

    python word_file_prefixer.py /path/to/extracted \
        --group-order R班 N班 S班

The command above prefixes all Word files and renames the group directories to
"01_R班", "02_N班", "03_S班" in that order.
"""

from __future__ import annotations

import argparse
import dataclasses
import re
from pathlib import Path
from typing import Iterator, List, Sequence, Set


WORD_EXTENSIONS: Set[str] = {".doc", ".docx"}
ORDER_PREFIX_PATTERN = re.compile(r"^(?P<index>\d+)[_-](?P<name>.+)$")


@dataclasses.dataclass(slots=True)
class GroupDirectory:
    """Represents a group directory and its resolved group name."""

    original_path: Path
    temp_path: Path | None = None
    group_name: str = dataclasses.field(init=False)

    def __post_init__(self) -> None:
        self.group_name = extract_group_name(self.original_path.name)

    def ensure_temp_name(self) -> None:
        """Rename the directory to a temporary unique name.

        This avoids collisions while we compute the final ordering.  The
        temporary name always starts with ``__tmp__`` so that it can be easily
        recognised.
        """

        if self.temp_path is not None:
            return

        parent = self.original_path.parent
        index = 0
        while True:
            candidate = parent / f"__tmp__{index:03d}__{self.original_path.name}"
            if not candidate.exists():
                self.original_path.rename(candidate)
                self.temp_path = candidate
                return
            index += 1

    def finalise_name(self, new_name: str) -> None:
        """Rename the temporary directory to ``new_name``."""

        if self.temp_path is None:
            raise RuntimeError("Temporary path not set before finalising name")

        final_path = self.temp_path.parent / new_name
        self.temp_path.rename(final_path)
        self.original_path = final_path
        self.temp_path = None


def extract_group_name(directory_name: str) -> str:
    """Return the group name stripped from any numeric ordering prefix.

    >>> extract_group_name("01_R班")
    'R班'
    >>> extract_group_name("R班")
    'R班'
    """

    match = ORDER_PREFIX_PATTERN.match(directory_name)
    if match:
        return match.group("name")
    return directory_name


def iter_group_directories(root_path: Path) -> Iterator[GroupDirectory]:
    """Yield :class:`GroupDirectory` instances for each direct sub-directory."""

    for child in sorted(root_path.iterdir()):
        if child.is_dir():
            yield GroupDirectory(child)


def _sorted_groups(
    groups: Sequence[GroupDirectory], desired_order: Sequence[str]
) -> List[GroupDirectory]:
    normalised_order = {name: index for index, name in enumerate(desired_order)}
    return sorted(
        groups,
        key=lambda group: (
            normalised_order.get(group.group_name, len(normalised_order)),
            group.group_name,
        ),
    )


def reorder_group_directories(
    groups: Sequence[GroupDirectory],
    desired_order: Sequence[str],
) -> None:
    """Rename directories to include zero padded ordering prefixes.

    The ``desired_order`` sequence defines the priority.  Any groups not
    explicitly listed are appended afterwards in alphabetical order.  Prefixes
    are added in the form ``01_`` so that the physical directory order reflects
    the requested arrangement (useful when compressing back to a zip file).
    """

    if not groups:
        return

    seen_names: Set[str] = set()
    for group in groups:
        if group.group_name in seen_names:
            raise ValueError(
                f"Duplicate group directory detected for '{group.group_name}'",
            )
        seen_names.add(group.group_name)

    ordered_groups = _sorted_groups(groups, desired_order)

    padding = max(2, len(str(len(ordered_groups))))

    # Temporarily rename directories to avoid name collisions.
    for group in groups:
        group.ensure_temp_name()

    # Finalise the order with numeric prefixes.
    for index, group in enumerate(ordered_groups, start=1):
        prefix = f"{index:0{padding}d}_"
        group.finalise_name(f"{prefix}{group.group_name}")


def add_prefix_to_word_files(group: GroupDirectory) -> List[Path]:
    """Prefix Word files within ``group`` with the group name.

    Returns a list of files that were renamed.
    """

    renamed_files: List[Path] = []
    prefix = f"【{group.group_name}】"

    for file_path in iter_word_files(group.original_path):
        if file_path.name.startswith(prefix):
            continue

        new_name = generate_unique_name(file_path, prefix)
        target = file_path.with_name(new_name)
        file_path.rename(target)
        renamed_files.append(target)

    return renamed_files


def iter_word_files(directory: Path) -> Iterator[Path]:
    """Yield all Word files within ``directory`` recursively."""

    for path in directory.rglob("*"):
        if path.is_file() and path.suffix.lower() in WORD_EXTENSIONS:
            yield path


def generate_unique_name(file_path: Path, prefix: str) -> str:
    """Generate a unique new name for ``file_path`` using ``prefix``."""

    stem = file_path.stem
    suffix = file_path.suffix
    base_name = f"{prefix}{stem}"
    candidate = f"{base_name}{suffix}"

    counter = 1
    while (file_path.parent / candidate).exists():
        candidate = f"{base_name}({counter}){suffix}"
        counter += 1

    return candidate


def parse_arguments(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """Parse command-line arguments for the prefixer utility."""

    parser = argparse.ArgumentParser(
        description="Prefix Word files in group folders with their group names",
    )
    parser.add_argument(
        "root",
        type=Path,
        help="Path to the extracted zip folder containing group directories",
    )
    parser.add_argument(
        "--group-order",
        nargs="*",
        metavar="GROUP",
        help="Optional list describing the desired order of groups",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Display the operations without performing any file changes",
    )

    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_arguments(argv)
    root_path: Path = args.root

    if not root_path.exists() or not root_path.is_dir():
        raise FileNotFoundError(f"{root_path} is not a valid directory")

    group_directories = list(iter_group_directories(root_path))

    if args.group_order:
        if args.dry_run:
            preview_order(group_directories, args.group_order)
        else:
            reorder_group_directories(group_directories, args.group_order)

    if args.dry_run:
        preview_prefix_changes(group_directories)
        return

    for group in group_directories:
        renamed = add_prefix_to_word_files(group)
        if renamed:
            print(f"Prefixed {len(renamed)} file(s) in {group.original_path.name}")


def preview_order(groups: Sequence[GroupDirectory], desired_order: Sequence[str]) -> None:
    """Print the new order without applying changes."""

    ordered_groups = _sorted_groups(groups, desired_order)
    print("Planned group order:")
    for index, group in enumerate(ordered_groups, start=1):
        print(f"  {index:02d}. {group.group_name}")


def preview_prefix_changes(groups: Sequence[GroupDirectory]) -> None:
    """Preview the Word files that would be renamed."""

    for group in groups:
        prefix = f"【{group.group_name}】"
        targets = list(iter_word_files(group.original_path))
        print(f"Group: {group.group_name} ({group.original_path})")
        for file_path in targets:
            if file_path.name.startswith(prefix):
                print(f"  [skip] {file_path.name}")
            else:
                new_name = generate_unique_name(file_path, prefix)
                print(f"  [rename] {file_path.name} -> {new_name}")


if __name__ == "__main__":
    main()