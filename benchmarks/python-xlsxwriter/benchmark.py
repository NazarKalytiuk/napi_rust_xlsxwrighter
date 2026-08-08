#!/usr/bin/env python3

import argparse
import json
import os
from pathlib import Path
import platform
import sys
import time

import xlsxwriter

DEFAULT_ROWS = 200_000
DEFAULT_COLUMNS = 20
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384


def positive_integer(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate the shared XLSX benchmark workload with XlsxWriter."
    )
    parser.add_argument("--rows", type=positive_integer, default=DEFAULT_ROWS)
    parser.add_argument("--columns", type=positive_integer, default=DEFAULT_COLUMNS)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("python-xlsxwriter-generation.xlsx"),
    )
    args = parser.parse_args()

    if args.rows > MAX_EXCEL_ROWS:
        parser.error(f"rows must not exceed Excel's limit of {MAX_EXCEL_ROWS}")
    if args.columns > MAX_EXCEL_COLUMNS:
        parser.error(f"columns must not exceed Excel's limit of {MAX_EXCEL_COLUMNS}")

    return args


def max_rss_bytes() -> int | None:
    try:
        import resource
    except ImportError:
        return None

    maximum = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    if sys.platform == "darwin":
        return int(maximum)
    return int(maximum) * 1_024


def make_row(row: int, columns: int) -> list[str | float]:
    return [
        f"Row {row} Col {column}"
        if column % 2 == 0
        else row * column * 0.123
        for column in range(columns)
    ]


def run(args: argparse.Namespace) -> None:
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    total_start = time.perf_counter_ns()
    workbook = xlsxwriter.Workbook(
        output,
        {"constant_memory": True, "tmpdir": os.fspath(output.parent)},
    )
    worksheet = workbook.add_worksheet("Data")

    write_start = time.perf_counter_ns()
    for row in range(args.rows):
        worksheet.write_row(row, 0, make_row(row, args.columns))
    write_ms = (time.perf_counter_ns() - write_start) / 1_000_000

    save_start = time.perf_counter_ns()
    workbook.close()
    save_ms = (time.perf_counter_ns() - save_start) / 1_000_000
    total_ms = (time.perf_counter_ns() - total_start) / 1_000_000

    result = {
        "implementation": "python/xlsxwriter",
        "libraryVersion": xlsxwriter.__version__,
        "runtimeVersion": platform.python_version(),
        "mode": "constant-memory",
        "writeMode": "row",
        "rows": args.rows,
        "columns": args.columns,
        "cells": args.rows * args.columns,
        "writeMs": write_ms,
        "saveMs": save_ms,
        "totalMs": total_ms,
        "maxRssBytes": max_rss_bytes(),
        "outputBytes": output.stat().st_size,
    }
    print(json.dumps(result, separators=(",", ":")))


if __name__ == "__main__":
    try:
        run(parse_args())
    except Exception as error:
        print(f"Python XlsxWriter benchmark failed: {error}", file=sys.stderr)
        raise SystemExit(1) from error
