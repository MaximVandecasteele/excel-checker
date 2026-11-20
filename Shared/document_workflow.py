from __future__ import annotations

import re
from pathlib import Path
from typing import Callable, TypeVar


Result = TypeVar("Result")


def _safe_filename(name: str) -> str:
    return re.sub(r"[^\w.-]+", "_", name, flags=re.UNICODE).strip("._") or "leerling"


def review_folder(
    reference_path: Path,
    input_folder: Path,
    output_folder: Path,
    *,
    file_extension: str,
    compare: Callable[[Path, Path], Result],
    write_report: Callable[[Result, Path], None],
    filename_suffix: str = "",
    filename_match: str = "suffix",
    report_suffix: str = "_rapport.html",
) -> list[Path]:
    """Run a document-specific comparison and report writer over a folder tree."""
    normalized_extension = file_extension.casefold()
    if not normalized_extension.startswith(".") or normalized_extension == ".":
        raise ValueError("file_extension must be a file extension starting with a dot.")
    if filename_match not in {"suffix", "contains"}:
        raise ValueError("filename_match must be 'suffix' or 'contains'.")

    output_folder.mkdir(parents=True, exist_ok=True)
    filename_filter = re.sub(r"\s+", "", filename_suffix).casefold()
    reference_resolved = reference_path.resolve()
    student_files = sorted(
        path for path in input_folder.rglob("*")
        if path.is_file()
        and path.suffix.casefold() == normalized_extension
        and (
            filename_filter in re.sub(r"\s+", "", path.stem).casefold()
            if filename_match == "contains"
            else re.sub(r"\s+", "", path.stem).casefold().endswith(filename_filter)
        )
        and not path.name.startswith("~$")
        and path.resolve() != reference_resolved
    )

    report_paths: list[Path] = []
    used_names: dict[str, int] = {}
    for student_file in student_files:
        relative_stem = student_file.relative_to(input_folder).with_suffix("")
        base = _safe_filename(str(relative_stem))
        used_names[base] = used_names.get(base, 0) + 1
        duplicate_suffix = f"_{used_names[base]}" if used_names[base] > 1 else ""
        report_path = output_folder / f"{base}{duplicate_suffix}{report_suffix}"
        write_report(compare(reference_path, student_file), report_path)
        report_paths.append(report_path)
    return report_paths
