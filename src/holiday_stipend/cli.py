"""Command-line entry point."""

import argparse
import csv
from pathlib import Path
import sys
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .discovery import inspect_file


def nonnegative_integer(value: str) -> int:
    try:
        number = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("must be a non-negative integer") from exc
    if number < 0:
        raise argparse.ArgumentTypeError("must be a non-negative integer")
    return number


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Inspect employee list files.")
    commands = parser.add_subparsers(dest="command", required=True)
    discover = commands.add_parser("discover", help="inventory Excel and CSV files")
    discover.add_argument("files", nargs="+", type=Path, help="files to inspect")
    discover.add_argument(
        "--sample-rows",
        type=nonnegative_integer,
        default=0,
        metavar="N",
        help="show the first N rows of each table (may contain personal information)",
    )
    args = parser.parse_args(argv)

    failed = False
    for path in args.files:
        try:
            tables = inspect_file(path, args.sample_rows)
        except (OSError, UnicodeError, csv.Error, BadZipFile, InvalidFileException, ParseError, EOFError, ValueError) as exc:
            print(f"Error inspecting {str(path)!r}: {exc}", file=sys.stderr)
            failed = True
            continue
        print(f"File: {str(path)!r}")
        for table in tables:
            visibility = f" ({table.visibility})" if table.visibility is not None else ""
            print(f"  Table: {table.name!r}{visibility}")
            print(f"    Rows: {table.row_count}; columns: {table.column_count}")
            for index, row in enumerate(table.samples, start=1):
                print(f"    Row {index}: {row!r}")
    return int(failed)


if __name__ == "__main__":
    sys.exit(main())
