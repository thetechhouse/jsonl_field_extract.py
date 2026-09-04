#!/usr/bin/env python3
"""Group an existing JSONL file by email and export it to Excel.

This is a standalone post-processing utility. It does not search or read the
original source dataset. Every input record is retained, including duplicates.
Rows are grouped by a case-insensitive email address and rows containing an
unseparated UK mobile number are placed first within each email group.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sqlite3
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, BinaryIO, Iterator, Sequence

try:
    import orjson  # type: ignore
except ImportError:
    orjson = None

try:
    import xlsxwriter  # type: ignore
except ImportError:
    xlsxwriter = None


EMAIL_ADDRESS_RE = re.compile(
    r"\b[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9-]+(?:\.[A-Z0-9-]+)+\b",
    re.IGNORECASE,
)
UK_MOBILE_RE = re.compile(r"(?<!\d)(?:\+447\d{9}|07\d{9})(?!\d)")
ILLEGAL_XML_RE = re.compile(
    "[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]"
)
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384
EXCEL_MAX_CELL_CHARS = 32_767
METADATA_COLUMNS = ("_group_email", "_contains_uk_mobile", "_source_line")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Group an existing JSONL result by email, prioritise records with "
            "UK mobile numbers, and export every record to an Excel workbook."
        )
    )
    parser.add_argument("input", nargs="?", help="Existing JSONL result file")
    parser.add_argument("output", nargs="?", help="Destination .xlsx file")
    parser.add_argument(
        "--prefer-domain",
        help=(
            "When a record has several email addresses, prefer one containing "
            "this text, e.g. @abc.com"
        ),
    )
    parser.add_argument(
        "--temp-dir",
        help=(
            "Directory for the disk-backed grouping database. Defaults to the "
            "Excel output directory."
        ),
    )
    parser.add_argument(
        "--rows-per-sheet",
        type=int,
        default=EXCEL_MAX_ROWS - 1,
        help=(
            "Maximum data rows per worksheet; the header uses one additional "
            f"row (default: {EXCEL_MAX_ROWS - 1})"
        ),
    )
    parser.add_argument(
        "--overwrite", action="store_true", help="Replace an existing output workbook"
    )
    return parser.parse_args()


def prompt_nonempty(prompt: str) -> str:
    while True:
        value = input(prompt).strip()
        if value:
            return value


def json_loads(raw: bytes) -> Any:
    return orjson.loads(raw) if orjson is not None else json.loads(raw)


def compact_json(value: Any) -> str:
    if orjson is not None:
        return orjson.dumps(value).decode("utf-8", errors="replace")
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def human_size(size: int) -> str:
    units = ("B", "KiB", "MiB", "GiB", "TiB")
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} TiB"


def iter_email_named_values(value: Any, inside_email_field: bool = False) -> Iterator[Any]:
    """Yield scalar values below keys whose names contain 'email'."""
    if isinstance(value, dict):
        for key, child in value.items():
            child_is_email = inside_email_field or "email" in str(key).casefold()
            yield from iter_email_named_values(child, child_is_email)
    elif isinstance(value, list):
        for child in value:
            yield from iter_email_named_values(child, inside_email_field)
    elif inside_email_field:
        yield value


def scalar_text(value: Any) -> str:
    if isinstance(value, str):
        return value
    if value is None:
        return "null"
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return compact_json(value)


def email_group(record: Any, preferred_domain: str | None) -> str | None:
    addresses: list[str] = []
    for value in iter_email_named_values(record):
        addresses.extend(
            match.group(0).casefold()
            for match in EMAIL_ADDRESS_RE.finditer(scalar_text(value))
        )
    if preferred_domain:
        needle = preferred_domain.casefold()
        for address in addresses:
            if needle in address:
                return address
    return addresses[0] if addresses else None


def contains_uk_mobile(value: Any) -> bool:
    if isinstance(value, dict):
        return any(contains_uk_mobile(child) for child in value.values())
    if isinstance(value, list):
        return any(contains_uk_mobile(child) for child in value)
    return isinstance(value, str) and UK_MOBILE_RE.search(value) is not None


def flatten_record(value: Any, prefix: str = "", result: dict[str, Any] | None = None) -> dict[str, Any]:
    """Flatten nested objects; retain arrays as compact JSON in one cell."""
    if result is None:
        result = {}
    if isinstance(value, dict):
        if not value and prefix:
            result[prefix] = "{}"
        for key, child in value.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            flatten_record(child, child_prefix, result)
    elif isinstance(value, list):
        result[prefix or "value"] = compact_json(value)
    else:
        result[prefix or "value"] = value
    return result


def safe_excel_string(value: str) -> str:
    value = ILLEGAL_XML_RE.sub("", value)
    if len(value) <= EXCEL_MAX_CELL_CHARS:
        return value
    marker = "… [truncated to Excel's 32,767-character cell limit]"
    return value[: EXCEL_MAX_CELL_CHARS - len(marker)] + marker


class GroupDatabase:
    def __init__(self, path: Path, preferred_domain: str | None) -> None:
        self.path = path
        self.preferred_domain = preferred_domain
        self.connection = sqlite3.connect(str(path))
        self.connection.execute("PRAGMA journal_mode=OFF")
        self.connection.execute("PRAGMA synchronous=OFF")
        self.connection.execute("PRAGMA temp_store=FILE")
        self.connection.execute("PRAGMA cache_size=-131072")
        self.connection.execute("PRAGMA locking_mode=EXCLUSIVE")
        self.connection.execute(
            "CREATE TABLE records ("
            "sequence INTEGER PRIMARY KEY, "
            "source_line INTEGER NOT NULL, "
            "sort_key TEXT NOT NULL, "
            "email_address TEXT, "
            "phone_priority INTEGER NOT NULL, "
            "record BLOB NOT NULL)"
        )
        self.total_records = 0
        self.phone_records = 0
        self.no_email_records = 0

    def add(self, record: Any, raw: bytes, source_line: int) -> None:
        address = email_group(record, self.preferred_domain)
        has_phone = contains_uk_mobile(record)
        sequence = self.total_records
        if address is None:
            sort_key = f"1:{sequence:020d}"
            self.no_email_records += 1
        else:
            sort_key = f"0:{address}"
        if has_phone:
            self.phone_records += 1
        self.connection.execute(
            "INSERT INTO records(sequence, source_line, sort_key, email_address, "
            "phone_priority, record) VALUES (?, ?, ?, ?, ?, ?)",
            (
                sequence,
                source_line,
                sort_key,
                address,
                0 if has_phone else 1,
                sqlite3.Binary(raw),
            ),
        )
        self.total_records += 1

    def finish_index(self) -> None:
        self.connection.commit()
        self.connection.execute(
            "CREATE INDEX records_group_order "
            "ON records(sort_key, phone_priority, sequence)"
        )
        self.connection.commit()

    def ordered_records(self) -> Iterator[tuple[int, str | None, bool, bytes]]:
        cursor = self.connection.execute(
            "SELECT source_line, email_address, phone_priority, record "
            "FROM records ORDER BY sort_key, phone_priority, sequence"
        )
        for source_line, address, phone_priority, raw in cursor:
            yield source_line, address, phone_priority == 0, raw

    def distinct_email_groups(self) -> int:
        row = self.connection.execute(
            "SELECT COUNT(DISTINCT email_address) FROM records "
            "WHERE email_address IS NOT NULL"
        ).fetchone()
        return int(row[0]) if row else 0

    def close(self) -> None:
        self.connection.close()


def ingest_jsonl(input_path: Path, database: GroupDatabase) -> tuple[list[str], dict[str, int]]:
    fields: set[str] = set()
    widths: dict[str, int] = {}
    max_data_columns = EXCEL_MAX_COLUMNS - len(METADATA_COLUMNS)
    with input_path.open("rb") as handle:
        for line_number, raw in enumerate(handle, 1):
            if not raw.strip():
                continue
            try:
                record = json_loads(raw)
            except (ValueError, json.JSONDecodeError) as exc:
                raise RuntimeError(
                    f"Invalid JSON on line {line_number:,} of {input_path}: {exc}"
                ) from exc
            flattened = flatten_record(record)
            fields.update(flattened)
            if len(fields) > max_data_columns:
                raise RuntimeError(
                    f"The data contains more than Excel's supported {max_data_columns:,} "
                    "source columns after flattening."
                )
            for key, value in flattened.items():
                display_length = len(scalar_text(value))
                widths[key] = min(60, max(widths.get(key, len(key)), display_length))
            database.add(record, raw, line_number)
            if database.total_records % 1_000_000 == 0:
                print(f"Loaded {database.total_records:,} records…")
    return sorted(fields, key=str.casefold), widths


def unique_display_columns(source_fields: Sequence[str]) -> list[tuple[str, str]]:
    reserved = set(METADATA_COLUMNS)
    columns: list[tuple[str, str]] = []
    used = set(reserved)
    for source_key in source_fields:
        display_key = source_key
        if display_key in used:
            display_key = f"json.{display_key}"
            suffix = 2
            while display_key in used:
                display_key = f"json{suffix}.{source_key}"
                suffix += 1
        used.add(display_key)
        columns.append((display_key, source_key))
    return columns


def write_cell(worksheet: Any, row: int, column: int, value: Any, cell_format: Any = None) -> None:
    if value is None:
        worksheet.write_blank(row, column, None, cell_format)
    elif isinstance(value, bool):
        worksheet.write_boolean(row, column, value, cell_format)
    elif isinstance(value, int) and not isinstance(value, bool):
        if abs(value) > 999_999_999_999_999:
            worksheet.write_string(row, column, str(value), cell_format)
        else:
            worksheet.write_number(row, column, value, cell_format)
    elif isinstance(value, float):
        if math.isfinite(value):
            worksheet.write_number(row, column, value, cell_format)
        else:
            worksheet.write_string(row, column, str(value), cell_format)
    else:
        # write_string deliberately prevents formula and hyperlink interpretation.
        worksheet.write_string(row, column, safe_excel_string(scalar_text(value)), cell_format)


def configure_data_sheet(
    workbook: Any,
    sheet_number: int,
    headers: Sequence[str],
    source_columns: Sequence[tuple[str, str]],
    widths: dict[str, int],
    header_format: Any,
) -> Any:
    worksheet = workbook.add_worksheet(f"Records_{sheet_number:03d}")
    worksheet.hide_gridlines(2)
    worksheet.freeze_panes(1, len(METADATA_COLUMNS))
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.repeat_rows(0)
    worksheet.set_margins(0.25, 0.25, 0.5, 0.5)
    worksheet.set_row(0, 24)
    for column, header in enumerate(headers):
        worksheet.write_string(0, column, header, header_format)
    worksheet.set_column(0, 0, 34)
    worksheet.set_column(1, 1, 20)
    worksheet.set_column(2, 2, 14)
    for offset, (_, source_key) in enumerate(source_columns, len(METADATA_COLUMNS)):
        worksheet.set_column(offset, offset, max(10, min(60, widths.get(source_key, 12) + 2)))
    return worksheet


def export_workbook(
    database: GroupDatabase,
    output_path: Path,
    temporary_xlsx: Path,
    source_fields: Sequence[str],
    widths: dict[str, int],
    input_path: Path,
    rows_per_sheet: int,
) -> int:
    assert xlsxwriter is not None
    workbook = xlsxwriter.Workbook(
        str(temporary_xlsx),
        {
            "constant_memory": True,
            "strings_to_formulas": False,
            "strings_to_urls": False,
            "strings_to_numbers": False,
        },
    )
    workbook.use_zip64()
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "align": "center",
            "valign": "vcenter",
            "bottom": 1,
            "bottom_color": "#163A5C",
        }
    )
    phone_format = workbook.add_format(
        {"bold": True, "font_color": "#7F6000", "bg_color": "#FFF2CC", "align": "center"}
    )
    summary_title = workbook.add_format(
        {"bold": True, "font_size": 16, "font_color": "#FFFFFF", "bg_color": "#1F4E78"}
    )
    summary_label = workbook.add_format({"bold": True, "font_color": "#1F4E78"})
    summary_number = workbook.add_format({"num_format": "#,##0"})

    summary = workbook.add_worksheet("Summary")
    summary.hide_gridlines(2)
    summary.set_portrait()
    summary.fit_to_pages(1, 1)
    summary.set_margins(0.4, 0.4, 0.5, 0.5)
    summary.merge_range("A1:B1", "JSONL Email Grouping Summary", summary_title)
    summary.set_row(0, 28)
    summary.set_column("A:A", 27)
    summary.set_column("B:B", 64)

    source_columns = unique_display_columns(source_fields)
    headers = [*METADATA_COLUMNS, *(display for display, _ in source_columns)]
    sheet_number = 1
    worksheet = configure_data_sheet(
        workbook, sheet_number, headers, source_columns, widths, header_format
    )
    data_row = 1
    sheets: list[tuple[Any, int]] = []

    for source_line, address, has_phone, raw in database.ordered_records():
        if data_row > rows_per_sheet:
            sheets.append((worksheet, data_row - 1))
            sheet_number += 1
            worksheet = configure_data_sheet(
                workbook, sheet_number, headers, source_columns, widths, header_format
            )
            data_row = 1
        record = json_loads(raw)
        flattened = flatten_record(record)
        write_cell(worksheet, data_row, 0, address or "")
        if has_phone:
            write_cell(worksheet, data_row, 1, "YES", phone_format)
        else:
            write_cell(worksheet, data_row, 1, "")
        write_cell(worksheet, data_row, 2, source_line)
        for column, (_, source_key) in enumerate(source_columns, len(METADATA_COLUMNS)):
            write_cell(worksheet, data_row, column, flattened.get(source_key))
        data_row += 1
        if (data_row - 1) % 1_000_000 == 0:
            print(f"Written {data_row - 1:,} rows to Excel…")

    sheets.append((worksheet, data_row - 1))
    for data_sheet, final_data_row in sheets:
        if final_data_row >= 1:
            data_sheet.autofilter(0, 0, final_data_row, len(headers) - 1)

    summary_values = [
        ("Input JSONL", str(input_path.resolve())),
        ("Output workbook", str(output_path.resolve())),
        ("Records retained", database.total_records),
        ("Distinct email groups", database.distinct_email_groups()),
        ("Records containing UK mobile", database.phone_records),
        ("Records without extracted email", database.no_email_records),
        ("Data worksheets", sheet_number),
        ("Deduplication", "None — every input record is retained"),
        ("Phone formats", "+447xxxxxxxxx and 07xxxxxxxxx"),
    ]
    for row, (label, value) in enumerate(summary_values, 2):
        summary.write_string(row, 0, label, summary_label)
        if isinstance(value, int):
            summary.write_number(row, 1, value, summary_number)
        else:
            summary.write_string(row, 1, safe_excel_string(str(value)))
    summary.freeze_panes(2, 0)
    workbook.close()
    return sheet_number


def main() -> int:
    args = parse_args()
    if xlsxwriter is None:
        print(
            "Error: XlsxWriter is required. Install it with:\n"
            "  python3 -m pip install XlsxWriter",
            file=sys.stderr,
        )
        return 2
    if not 1 <= args.rows_per_sheet <= EXCEL_MAX_ROWS - 1:
        print(
            f"Error: --rows-per-sheet must be between 1 and {EXCEL_MAX_ROWS - 1:,}.",
            file=sys.stderr,
        )
        return 2

    input_path = Path(args.input or prompt_nonempty("Existing JSONL file: ")).expanduser()
    output_path = Path(args.output or prompt_nonempty("Excel output file: ")).expanduser()
    if not input_path.is_file():
        print(f"Error: input file does not exist: {input_path}", file=sys.stderr)
        return 2
    if output_path.suffix.casefold() != ".xlsx":
        output_path = output_path.with_suffix(".xlsx")
    if output_path.exists() and not args.overwrite:
        print(f"Error: output exists; use --overwrite to replace it: {output_path}", file=sys.stderr)
        return 2

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_dir = Path(args.temp_dir).expanduser() if args.temp_dir else output_path.parent
    temp_dir.mkdir(parents=True, exist_ok=True)
    database_path: Path | None = None
    temporary_xlsx: Path | None = None
    database: GroupDatabase | None = None
    started = time.monotonic()

    try:
        with tempfile.NamedTemporaryFile(
            prefix=".jsonl-excel-groups-", suffix=".sqlite3", dir=temp_dir, delete=False
        ) as database_file:
            database_path = Path(database_file.name)
        database = GroupDatabase(database_path, args.prefer_domain)
        print(f"Reading existing JSONL only: {input_path}")
        fields, widths = ingest_jsonl(input_path, database)
        print(
            f"Loaded {database.total_records:,} record(s); building the email grouping index…"
        )
        database.finish_index()

        with tempfile.NamedTemporaryFile(
            prefix=".jsonl-email-groups-",
            suffix=".xlsx",
            dir=output_path.parent,
            delete=False,
        ) as output_file:
            temporary_xlsx = Path(output_file.name)
        sheet_count = export_workbook(
            database,
            output_path,
            temporary_xlsx,
            fields,
            widths,
            input_path,
            args.rows_per_sheet,
        )
        database.close()
        database = None
        os.replace(temporary_xlsx, output_path)
        temporary_xlsx = None
    except KeyboardInterrupt:
        print("\nCancelled; incomplete workbook was not retained.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"\nError: {exc}", file=sys.stderr)
        return 1
    finally:
        if database is not None:
            try:
                database.close()
            except sqlite3.Error:
                pass
        if database_path is not None:
            try:
                database_path.unlink()
            except OSError:
                pass
        if temporary_xlsx is not None:
            try:
                temporary_xlsx.unlink()
            except OSError:
                pass

    elapsed = time.monotonic() - started
    print(
        f"\nComplete: created {output_path} "
        f"({human_size(output_path.stat().st_size)}), "
        f"{sheet_count:,} data worksheet(s), elapsed {elapsed:,.1f}s."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
