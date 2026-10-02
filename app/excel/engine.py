"""Orchestrates: detect -> validate -> consolidate -> analyse -> write master.

CLI (auto-detects every .xlsx in the input folder):
    python -m app.excel.engine
    python -m app.excel.engine --input data\\input --output data\\output [--rebuild]
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Callable

import pandas as pd

from app.excel import processor as P
from app.excel.report_generator import write_master
from app.excel.validator import validate_files
from app.models.data_models import (FileValidationResult, ProcessingResult,
                                    ReportError, ValidationFailed)
from app.utils import constants as C
from app.utils.file_utils import (default_input_dir, default_output_dir,
                                  ensure_dir, find_excel_files)

ProgressFn = Callable[[float, str], None]
LogFn = Callable[[str], None]


def validate_only(paths: list[Path], log: LogFn | None = None) -> list[FileValidationResult]:
    """Validate files and log a line per file (used by the GUI's Validate button)."""
    log = log or (lambda m: None)
    results, _ = validate_files([Path(p) for p in paths])
    for r in results:
        if r.ok:
            log(f"[OK]   {r.path.name}: {r.rows} rows x {r.columns} columns")
        else:
            log(f"[FAIL] {r.path.name}")
            for e in r.errors:
                log(f"         - {e}")
    return results


def load_existing_master(path: Path) -> pd.DataFrame:
    """Read the Consolidated Data sheet of an existing master report."""
    try:
        df = pd.read_excel(path, sheet_name=C.SHEET_CONSOLIDATED, engine="openpyxl")
    except PermissionError:
        raise ReportError(f"Cannot read the existing '{path.name}'. "
                          "If it is open in Excel, close it and try again.")
    except Exception as exc:
        raise ReportError(f"The existing '{path.name}' is corrupt or not a valid master report "
                          f"({exc}). Move or delete it, or run a rebuild.")
    missing = [c for c in C.REQUIRED_COLUMNS + [C.COL_SOURCE] if c not in df.columns]
    if missing:
        raise ReportError(f"The existing '{path.name}' is missing column(s): {', '.join(missing)}. "
                          "Move or delete it, or run a rebuild.")
    return df.dropna(how="all")


def build_master(paths: list[Path], output_dir: Path, rebuild: bool = False,
                 progress: ProgressFn | None = None, log: LogFn | None = None) -> ProcessingResult:
    """Create or update Master_Sales_Report.xlsx from the given input files."""
    progress = progress or (lambda f, m: None)
    log = log or (lambda m: None)

    paths = [Path(p) for p in paths]
    if not paths:
        raise ReportError("No Excel files selected. Add files or place them in the input folder.")
    if not output_dir:
        raise ReportError("No output folder selected.")
    names = [p.name for p in paths]
    dupes = sorted({n for n in names if names.count(n) > 1})
    if dupes:
        raise ReportError(f"Two selected files share the same name: {', '.join(dupes)}.")

    output_dir = Path(output_dir)
    try:
        ensure_dir(output_dir)
    except PermissionError:
        raise ReportError(f"No permission to use the output folder: {output_dir}")
    except OSError as exc:
        raise ReportError(f"Cannot create the output folder '{output_dir}': {exc}")
    master_path = output_dir / C.MASTER_FILENAME

    # 1. Validate
    progress(0.05, "Validating files...")
    log(f"Validating {len(paths)} file(s)...")
    results, frames = validate_files(paths)
    for r in results:
        log(f"  [{'OK' if r.ok else 'FAIL'}] {r.path.name}: {r.rows} rows x {r.columns} columns")
    if any(not r.ok for r in results):
        raise ValidationFailed(results)

    # 2. Decide what is new
    existing, known = None, set()
    if master_path.exists() and not rebuild:
        progress(0.30, "Reading existing master report...")
        existing = load_existing_master(master_path)
        known = set(existing[C.COL_SOURCE].astype(str))
        log(f"Existing master found ({len(existing)} rows from {len(known)} file(s)).")
    elif master_path.exists():
        log("Rebuild requested: existing master will be replaced.")

    new_paths = [p for p in paths if p.name not in known]
    skipped = [p.name for p in paths if p.name in known]
    for n in skipped:
        log(f"  Skipping {n}: already in the master report.")

    if existing is not None and not new_paths:
        progress(1.0, "Nothing to add.")
        return ProcessingResult(
            master_path=master_path, skipped_files=skipped, total_rows=len(existing),
            message=f"{C.MASTER_FILENAME} already contains all {len(skipped)} file(s). Nothing to add.")

    # 3. Consolidate old + new
    progress(0.45, "Consolidating data...")
    parts = []
    if existing is not None:
        old = P.to_typed(existing)
        old[C.COL_SOURCE] = existing[C.COL_SOURCE].astype(str).values
        parts.append(old)
    for p in new_paths:
        typed = P.to_typed(frames[p])
        typed[C.COL_SOURCE] = p.name
        parts.append(typed)
        log(f"  Adding {p.name} ({len(typed)} rows)")
    combined = P.consolidate(parts)

    # 4. Write (temp file first, so a failure never destroys an existing master)
    progress(0.70, "Generating master report...")
    tmp = master_path.with_name(master_path.name + ".tmp")
    try:
        write_master(combined, tmp)
        os.replace(tmp, master_path)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise ReportError(f"Cannot write '{master_path.name}'. If it is open in Excel, close it "
                          "and try again, or check the folder permissions.")
    except Exception as exc:
        tmp.unlink(missing_ok=True)
        raise ReportError(f"Report generation failed: {exc}") from exc

    added_rows = sum(r.rows for r in results if r.path in new_paths)
    created = existing is None
    verb = "Created" if created else "Updated"
    message = (f"{verb} {C.MASTER_FILENAME}: added {len(new_paths)} file(s) "
               f"({added_rows} rows); {len(combined)} rows in total.")
    progress(1.0, "Done.")
    log(message)
    return ProcessingResult(
        master_path=master_path, created_new=created, files_processed=len(new_paths),
        rows_processed=added_rows,
        columns_validated=sum(r.columns for r in results if r.path in new_paths),
        total_rows=len(combined), added_files=[p.name for p in new_paths],
        skipped_files=skipped, message=message)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Build Master_Sales_Report.xlsx from input workbooks.")
    ap.add_argument("--input", type=Path, default=default_input_dir(), help="folder to scan")
    ap.add_argument("--output", type=Path, default=default_output_dir(), help="report folder")
    ap.add_argument("--rebuild", action="store_true", help="ignore any existing master")
    args = ap.parse_args(argv)

    files = find_excel_files(args.input)
    if not files:
        print(f"No .xlsx files found in {args.input}")
        return 1
    print(f"Detected {len(files)} Excel file(s) in {args.input}")
    try:
        result = build_master(files, args.output, rebuild=args.rebuild, log=print)
    except ReportError as exc:
        print(f"\nERROR: {exc}")
        return 1
    print(f"\n{result.message}\nOutput: {result.master_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())