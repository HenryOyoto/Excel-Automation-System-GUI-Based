"""Path helpers (dev + PyInstaller) and Excel file auto-detection."""
from __future__ import annotations

import sys
from pathlib import Path

from app.utils.constants import MASTER_FILENAME


def is_frozen() -> bool:
    """True when running from a PyInstaller-built executable."""
    return getattr(sys, "frozen", False)


def resource_path(relative: str) -> Path:
    """Locate a bundled read-only resource (icons etc.) in dev or in the EXE."""
    if is_frozen():
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parents[2]
    return base / relative


def project_root() -> Path:
    """Writable base folder: project root in dev, EXE folder when frozen."""
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def default_input_dir() -> Path:
    return project_root() / "data" / "input"


def default_output_dir() -> Path:
    return project_root() / "data" / "output"


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def find_excel_files(folder: Path) -> list[Path]:
    """Auto-detect input workbooks in a folder.

    Skips Excel lock/temp files (~$...) and the master report itself, so the
    app never tries to consume its own output.
    """
    if not folder.is_dir():
        return []
    return sorted(
        p for p in folder.glob("*.xlsx")
        if not p.name.startswith("~$") and p.name != MASTER_FILENAME
    )