"""Plain data containers and exceptions shared by the engine and the GUI."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


class ReportError(Exception):
    """A problem with a user-friendly message that the GUI can show directly."""


class ValidationFailed(ReportError):
    """One or more input files failed validation."""

    def __init__(self, results: list["FileValidationResult"]):
        self.results = results
        lines = ["Validation failed:"]
        for r in results:
            if not r.ok:
                lines.append(f"{r.path.name}:")
                lines += [f"  - {e}" for e in r.errors]
        super().__init__("\n".join(lines))


@dataclass
class FileValidationResult:
    path: Path
    rows: int = 0
    columns: int = 0
    headers: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.errors


@dataclass
class ProcessingResult:
    master_path: Path
    created_new: bool = False
    files_processed: int = 0          # files added in this run
    rows_processed: int = 0           # rows added in this run
    columns_validated: int = 0
    total_rows: int = 0               # rows now in the master report
    added_files: list[str] = field(default_factory=list)
    skipped_files: list[str] = field(default_factory=list)
    message: str = ""