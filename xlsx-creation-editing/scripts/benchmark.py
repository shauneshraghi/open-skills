"""Benchmark xlsx-creation-editing operations against the Apache POI corpus.

Design:
- Uses POI_PATH env var (default: poi/test-data/spreadsheet relative to repo root)
  per the cross-platform convention; all paths via pathlib.Path.
- tempfile.mkdtemp() for all output to avoid hardcoded paths.
- Measures four wall-clock operations per file: open, iterate, write_100, validate.
- Prints a fixed-width table; accepts --quick for first 3 fixtures only.
"""

from __future__ import annotations

import argparse
import os
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
import sys
sys.path.insert(0, str(ROOT))

import openpyxl

from scripts.create_xlsx import create_workbook, save, write_cell
from scripts.validate_xlsx import validate_xlsx

POI_PATH = Path(
    os.environ.get(
        "POI_PATH",
        str(Path(__file__).resolve().parents[3] / "poi" / "test-data" / "spreadsheet"),
    )
)

BENCH_FIXTURES = [
    "sample.xlsx",
    "simple-monthly-budget.xlsx",
    "ExcelTables.xlsx",
    "Formatting.xlsx",
    "ConditionalFormattingSamples.xlsx",
    "123233_charts.xlsx",
    "styles.xlsx",
    "DateFormatTests.xlsx",
    "AverageTaxRates.xlsx",
    "simple-table-named-range.xlsx",
]


def _time_open(path: Path) -> float:
    start = time.perf_counter()
    openpyxl.load_workbook(str(path), data_only=True)
    return (time.perf_counter() - start) * 1000


def _time_iterate(path: Path) -> float:
    start = time.perf_counter()
    wb = openpyxl.load_workbook(str(path), data_only=True)
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for _ in row:
                pass
    return (time.perf_counter() - start) * 1000


def _time_write_100() -> float:
    tmpdir = Path(tempfile.mkdtemp())
    start = time.perf_counter()
    wb = create_workbook()
    ws = wb.active
    for r in range(1, 11):
        for c in range(1, 11):
            write_cell(ws, r, c, r * c)
    save(wb, tmpdir / "bench_write.xlsx")
    return (time.perf_counter() - start) * 1000


def _time_validate(path: Path) -> float:
    start = time.perf_counter()
    validate_xlsx(path)
    return (time.perf_counter() - start) * 1000


def main() -> None:
    parser = argparse.ArgumentParser(description="Benchmark xlsx-creation-editing")
    parser.add_argument("--quick", action="store_true", help="First 3 fixtures only")
    args = parser.parse_args()

    if not POI_PATH.exists():
        print(f"ERROR: POI_PATH not found: {POI_PATH}")
        sys.exit(1)

    fixtures = BENCH_FIXTURES[:3] if args.quick else BENCH_FIXTURES
    col_w = max(len(f) for f in fixtures) + 2

    header = f"{'Fixture':<{col_w}}  {'open':>8}  {'iterate':>8}  {'write_100':>9}  {'validate':>9}"
    print(header)
    print("-" * len(header))

    for fname in fixtures:
        path = POI_PATH / fname
        if not path.exists():
            print(f"{fname:<{col_w}}  (not found)")
            continue
        try:
            t_open = _time_open(path)
            t_iter = _time_iterate(path)
            t_write = _time_write_100()
            t_val = _time_validate(path)
            print(
                f"{fname:<{col_w}}  {t_open:7.1f}ms  {t_iter:7.1f}ms"
                f"  {t_write:8.1f}ms  {t_val:8.1f}ms"
            )
        except Exception as exc:
            print(f"{fname:<{col_w}}  ERROR: {exc}")


if __name__ == "__main__":
    main()
