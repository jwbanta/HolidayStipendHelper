"""Inspect workbook and CSV structure without changing the source files."""

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook


@dataclass
class TableSummary:
    name: str
    visibility: str | None
    row_count: int
    column_count: int
    samples: list[tuple[Any, ...]]


def summarize_rows(
    name: str, visibility: str | None, rows: Iterable[Iterable[Any]], sample_count: int
) -> TableSummary:
    row_count = 0
    column_count = 0
    samples = []
    for row in rows:
        values = tuple(row)
        row_count += 1
        column_count = max(column_count, len(values))
        if len(samples) < sample_count:
            samples.append(values)
    return TableSummary(name, visibility, row_count, column_count, samples)


def inspect_file(path: Path, sample_count: int = 0) -> list[TableSummary]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        with path.open(encoding="utf-8-sig", newline="") as file:
            return [summarize_rows(path.name, None, csv.reader(file, strict=True), sample_count)]
    if suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, read_only=True, data_only=False)
        try:
            return [
                summarize_rows(
                    sheet.title,
                    sheet.sheet_state,
                    sheet.iter_rows(values_only=True),
                    sample_count,
                )
                for sheet in workbook.worksheets
            ]
        finally:
            workbook.close()
    raise ValueError(f"unsupported file type {suffix or '(none)'}; use .csv, .xlsx, or .xlsm")
