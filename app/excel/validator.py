"""File and data validation for input workbooks."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from app.models.data_models import FileValidationResult
from app.utils import constants as C

MAX_EXAMPLES = 5


def _blank(series: pd.Series) -> pd.Series:
    """True where a cell is missing or only whitespace."""
    return series.isna() | (series.astype(str).str.strip() == "")


def _rows(index) -> str:
    """Format DataFrame index labels as Excel row numbers (row 1 = header)."""
    labels = list(index)
    shown = ", ".join(str(i + 2) for i in labels[:MAX_EXAMPLES])
    extra = len(labels) - MAX_EXAMPLES
    return shown + (f" (+{extra} more)" if extra > 0 else "")


def validate_file(path: Path) -> tuple[FileValidationResult, pd.DataFrame | None]:
    """Validate one workbook. Returns (result, cleaned DataFrame or None)."""
    path = Path(path)
    res = FileValidationResult(path=path)

    if not path.exists():
        res.errors.append("File does not exist.")
        return res, None
    if path.suffix.lower() != ".xlsx":
        res.errors.append(f"Not an .xlsx workbook (found '{path.suffix or 'no extension'}').")
        return res, None

    try:
        df = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    except PermissionError:
        res.errors.append("Cannot read the file (permission denied, or it is open in Excel). "
                          "Close it and try again.")
        return res, None
    except Exception as exc:  # BadZipFile, KeyError, InvalidFileException, ...
        res.errors.append(f"Workbook is corrupt or unreadable: {exc}")
        return res, None

    df = df.dropna(how="all")  # ignore fully blank rows; index keeps original positions
    if len(df.columns) == 0:
        res.errors.append("Workbook is empty (no data found).")
        return res, None

    df.columns = [str(c).strip() for c in df.columns]
    res.headers = list(df.columns)
    res.columns = len(df.columns)
    res.rows = len(df)

    if res.columns < C.MIN_COLUMNS:
        res.errors.append(f"Only {res.columns} columns found (minimum {C.MIN_COLUMNS}).")
    missing = [c for c in C.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        res.errors.append("Missing required column(s): " + ", ".join(missing))
        return res, None  # nothing else can be checked reliably
    if res.rows < C.MIN_ROWS:
        res.errors.append(f"Only {res.rows} data rows found (minimum {C.MIN_ROWS}).")

    # Required (typed-in) fields must not be empty.
    for col in C.BASE_COLUMNS:
        bad = df.index[_blank(df[col])]
        if len(bad):
            res.errors.append(f"'{col}' is empty in row(s): {_rows(bad)}")

    # Dates must be real dates.
    dates = pd.to_datetime(df[C.COL_DATE], errors="coerce", format="mixed")
    bad = df.index[dates.isna() & ~_blank(df[C.COL_DATE])]
    if len(bad):
        res.errors.append(f"Invalid date(s) in row(s): {_rows(bad)}")

    # Numeric fields must be numbers; ranges must make sense.
    for col in C.NUMERIC_COLUMNS:
        nums = pd.to_numeric(df[col], errors="coerce")
        bad = df.index[nums.isna() & ~_blank(df[col])]
        if len(bad):
            res.errors.append(f"'{col}' has non-numeric values in row(s): {_rows(bad)}")
        neg = df.index[nums < 0]
        if len(neg):
            res.errors.append(f"'{col}' has negative values in row(s): {_rows(neg)}")
        if col in (C.COL_DISCOUNT, C.COL_TAX_RATE):
            over = df.index[nums > 1]
            if len(over):
                res.errors.append(f"'{col}' must be a fraction such as 0.10 for 10% "
                                  f"(row(s): {_rows(over)})")

    return res, (df if res.ok else None)


def validate_files(paths: list[Path]) -> tuple[list[FileValidationResult], dict[Path, pd.DataFrame]]:
    """Validate several files, then check that their columns are consistent."""
    results: list[FileValidationResult] = []
    frames: dict[Path, pd.DataFrame] = {}
    for p in paths:
        res, df = validate_file(p)
        results.append(res)
        if df is not None:
            frames[Path(p)] = df

    # Cross-file consistency: every file must have the same set of columns.
    reference = next((r for r in results if r.headers and r.ok), None)
    if reference:
        for r in results:
            if r is reference or not r.headers:
                continue
            missing = set(reference.headers) - set(r.headers)
            extra = set(r.headers) - set(reference.headers)
            if missing or extra:
                r.errors.append(
                    f"Columns differ from {reference.path.name} "
                    f"(missing: {sorted(missing) or 'none'}, extra: {sorted(extra) or 'none'})."
                )
                frames.pop(Path(r.path), None)
    return results, frames