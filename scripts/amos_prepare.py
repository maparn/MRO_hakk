#!/usr/bin/env python3
"""Magnetic MRO: prepare AMOS reference exports for subsequent STAT calculations.

Reads two ORIGINAL workbooks in read-only mode:
    AMOS_uus_süsteem_info.xlsx   (partno, description, measure_unit,
                                 avg_price, sales, purch_price)
    AMOS_UOM.xlsx                (partno, measure_unit_ext, measure_unit,
                                 rate, exchange_base)

Creates only excel_files/output/amos_prepared.xlsx with sheets:
    Prepared  - one row per source AMOS record; no auto-selected price
    UOM_Map   - full measure-unit mapping, including ambiguous rules
    Issues    - rows needing attention before downstream calculations
    Summary   - counts and rules/limitations

NEVER edits source workbooks or the teammate's STAT workbook.
No unit conversions, price selection, currency conversions or price formulas.

Run from the MRO_hakk folder:
    python scripts/amos_prepare.py
Optional:
    python scripts/amos_prepare.py --amos "path/to/new.xlsx" --uom "path/to/UOM.xlsx"
"""

from __future__ import annotations

import argparse
import math
from collections import Counter, defaultdict
from contextlib import contextmanager
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


if __package__:
    from .report_io import literal_cell, save_report, validate_output
    from .uom_rules import annotate_rules, normalize_key as norm
else:
    from report_io import literal_cell, save_report, validate_output
    from uom_rules import annotate_rules, normalize_key as norm


ROOT = Path(__file__).resolve().parent.parent
AMOS_NAME = "AMOS_uus_süsteem_info.xlsx"
UOM_NAME = "AMOS_UOM.xlsx"
DEFAULT_OUTPUT = ROOT / "excel_files" / "output" / "amos_prepared.xlsx"

AMOS_FIELDS = ("partno", "description", "measure_unit", "avg_price", "sales", "purch_price")
UOM_FIELDS = ("partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base")
UOM_CONTEXT = ("description", "mutation", "mutator", "status", "mutation_time",
               "created_by", "created_date")


def visible(value: object) -> str:
    return "" if value is None else str(value).strip()


def locate(filename: str) -> Path:
    paths = (ROOT / "excel_files" / "input" / filename, ROOT / filename)
    for path in paths:
        if path.is_file():
            return path
    raise FileNotFoundError(f"Missing {filename}. Put it in MRO_hakk/ or excel_files/input/.")


@contextmanager
def open_records(path: Path, required: tuple[str, ...], optional: tuple[str, ...] = ()):
    """Provide repeatable row iteration on the same open workbook snapshot."""
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Sheet1" not in book.sheetnames:
            raise ValueError(f"Expected Sheet1 in {path.name}. Found {book.sheetnames}")
        sheet = book["Sheet1"]
        sheet.reset_dimensions()  # Some AMOS exports declare an incorrect used range.
        header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header is None:
            raise ValueError(f"Empty file: {path.name}")
        names = [norm(h).lower() for h in header]
        missing = [name for name in required if name not in names]
        if missing:
            raise ValueError(f"{path.name}: missing columns: {', '.join(missing)}")
        if len(set(names)) != len(names):
            raise ValueError(f"{path.name}: duplicate column headers")
        indexes = {field: names.index(field) for field in (*required, *optional) if field in names}

        def rows():
            have_data = False
            for excel_row, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
                if not any(v is not None and visible(v) for v in values):
                    continue
                have_data = True
                record = {field: values[i] if i < len(values) else None
                          for field, i in indexes.items()}
                record["_excel_row"] = excel_row
                yield record
            if not have_data:
                raise ValueError(f"{path.name}: Sheet1 has column headers but no data rows")

        yield rows
    finally:
        book.close()


def read_records(path: Path, required: tuple[str, ...],
                 optional: tuple[str, ...] = ()) -> list[dict]:
    with open_records(path, required, optional) as rows:
        return list(rows())


def parse_price(value: object) -> tuple[float | None, str]:
    """Validate numbers. None != zero; do not invent a replacement value."""
    if value is None or visible(value) == "":
        return None, "MISSING"
    try:
        # Decimal(str(...)) avoids binary float comparisons against zero.
        amount = Decimal(visible(value).replace(",", "."))
    except InvalidOperation:
        return None, "INVALID"
    if not amount.is_finite():
        return None, "INVALID"
    if amount < 0:
        return None, "NEGATIVE"
    number = float(amount)
    if not math.isfinite(number) or (amount != 0 and number == 0):
        return None, "INVALID"
    return number, "POSITIVE" if amount > 0 else "ZERO"


def assess_units(part: str, records: list[dict], source_unit: str):
    """Summarize rule candidates; their direction and commercial use require confirmation."""
    if not source_unit:
        return "SOURCE_UNIT_MISSING", 0, 0, "MISSING_SOURCE_UNIT"
    if not records:
        reason = "NO_UOM_RULE" if part else "MISSING_PARTNO"
        return "NO_UOM_RULE", 0, 0, reason

    internal = [r for r in records if norm(r["measure_unit"]) == source_unit]
    external = [r for r in records if norm(r["measure_unit_ext"]) == source_unit]
    if not internal and not external:
        return "UNIT_NOT_LISTED_IN_UOM", 0, 0, "UNIT_NOT_LISTED_IN_UOM"

    matched_rows = {r["_excel_row"]: r for r in (*internal, *external)}
    valid = sum(not r["rule_issues"] for r in matched_rows.values())
    if valid < len(matched_rows):
        decision = "REVIEW_INVALID_RULE"
    elif valid > 1:
        decision = "REVIEW_MULTIPLE_RULES"
    else:
        decision = "REVIEW_DIRECTION_AND_POLICY"
    if internal and external:
        status = "MATCHES_BOTH_SIDES"
    else:
        status = "MATCHES_INTERNAL_UNIT" if internal else "MATCHES_EXTERNAL_UNIT"
    return status, len(matched_rows), valid, decision


def make_sheet(book: Workbook, name: str, headers: list[str], widths: list[int]):
    """Streaming writer: handles 66k+ records without storing an Excel cell grid."""
    sheet = book.create_sheet(name)
    sheet.freeze_panes = "A2"
    for i, width in enumerate(widths, 1):
        sheet.column_dimensions[get_column_letter(i)].width = width
    styled = []
    for text in headers:
        cell = WriteOnlyCell(sheet, value=text)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        styled.append(cell)
    sheet.append(styled)
    return sheet


def append(sheet, values) -> None:
    sheet.append([literal_cell(sheet, v) for v in values])


def generate(amos_path: Path, uom_path: Path, output: Path) -> None:
    inputs = (amos_path, uom_path)
    validate_output(output, ROOT, inputs, owner=DEFAULT_OUTPUT.name)

    print(f"Loading UOM:  {uom_path.name}")
    uom = read_records(uom_path, UOM_FIELDS, UOM_CONTEXT)
    annotate_rules(uom)

    print(f"Loading AMOS: {amos_path.name}")
    with open_records(amos_path, AMOS_FIELDS) as rows:
        # Count keys first, then stream all source records into the write-only report.
        part_counts = Counter(filter(None, (norm(r["partno"]) for r in rows())))
        write_prepared_report(rows(), part_counts, uom, amos_path, uom_path, output)


def write_prepared_report(amos, part_counts, uom, amos_path, uom_path, output):
    uom_by_part = defaultdict(list)
    for r in uom:
        part = norm(r["partno"])
        if part:
            uom_by_part[part].append(r)

    book = Workbook(write_only=True)
    prepared_headers = [
        "source_row", "partno", "description", "measure_unit", "avg_price",
        "purch_price", "sales", "avg_price_state", "purch_price_state",
        "price_status", "uom_rule_count", "uom_status", "issues",
        "avg_price_raw", "purch_price_raw", "partno_raw", "measure_unit_raw",
        "uom_candidate_count", "uom_valid_candidate_count", "uom_conversion_status",
    ]
    prepared = make_sheet(book, "Prepared", prepared_headers,
                          [12, 23, 45, 16, 17, 17, 14, 20, 20, 21, 18, 31, 48, 18, 18, 23, 18, 23, 27, 36])
    uom_sheet = make_sheet(book, "UOM_Map", [
        "source_row", "partno", "measure_unit_ext", "measure_unit", "rate",
        "exchange_base", "exact_rule_repetitions", "uom_status",
        "rate_raw", "exchange_base_raw", "partno_raw", "measure_unit_ext_raw",
        "measure_unit_raw", *UOM_CONTEXT,
    ], [12, 23, 23, 20, 15, 18, 25, 33, 15, 20, 23, 24, 20, 45, 20, 20, 14, 20, 20, 20])
    issue_sheet = make_sheet(book, "Issues", [
        "source", "source_row", "partno", "measure_unit", "issues",
    ], [14, 14, 23, 20, 65])
    summary_sheet = make_sheet(book, "Summary", ["Metric", "Value"], [48, 75])

    metrics = Counter()
    for r in amos:
        metrics["AMOS_DATA_ROWS"] += 1
        part = norm(r["partno"])
        unit = norm(r["measure_unit"])
        avg, avg_state = parse_price(r["avg_price"])
        purch, purch_state = parse_price(r["purch_price"])
        avg_ok = avg_state == "POSITIVE"
        purch_ok = purch_state == "POSITIVE"
        if avg_ok and purch_ok:
            price_status = "BOTH_PRESENT"
        elif avg_ok:
            price_status = "AVG_ONLY"
        elif purch_ok:
            price_status = "PURCH_ONLY"
        else:
            price_status = "NO_USABLE_PRICE"

        uom_entries = uom_by_part.get(part, [])
        unit_status, candidates, valid_candidates, conversion = assess_units(part, uom_entries, unit)
        errors = []
        if not part:
            errors.append("MISSING_PARTNO")
        elif part == "-":
            errors.append("CHECK_PLACEHOLDER_PARTNO")
        if part and part_counts[part] > 1:
            errors.append("REPEATED_AMOS_PARTNO")
        if not unit:
            errors.append("MISSING_UNIT")
        if avg_state in ("INVALID", "NEGATIVE"):
            errors.append("BAD_AVG_PRICE")
        if purch_state in ("INVALID", "NEGATIVE"):
            errors.append("BAD_PURCH_PRICE")
        if price_status == "NO_USABLE_PRICE":
            errors.append("NO_USABLE_PRICE")
        if unit_status == "UNIT_NOT_LISTED_IN_UOM":
            errors.append("CHECK_UNIT_AGAINST_UOM")
        if conversion in ("REVIEW_INVALID_RULE", "REVIEW_MULTIPLE_RULES"):
            errors.append(conversion)

        metrics[f"PRICE_{price_status}"] += 1
        metrics[f"UOM_{unit_status}"] += 1
        if errors:
            metrics["ROWS_WITH_ISSUES"] += 1
            append(issue_sheet, ["AMOS", r["_excel_row"], part, unit, "; ".join(errors)])

        append(prepared, [
            r["_excel_row"], part, r["description"], unit, avg, purch,
            r["sales"], avg_state, purch_state, price_status,
            len(uom_entries), unit_status, "; ".join(errors),
            r["avg_price"], r["purch_price"], r["partno"], r["measure_unit"],
            candidates, valid_candidates, conversion,
        ])

    for r in uom:
        metrics["UOM_DATA_ROWS"] += 1
        part = norm(r["partno"])
        ext = norm(r["measure_unit_ext"])
        internal = norm(r["measure_unit"])
        rate, _ = parse_price(r["rate"])
        base, _ = parse_price(r["exchange_base"])
        repeat = r["exact_rule_repetitions"]
        problems = r["rule_issues"]
        if problems:
            metrics["UOM_ROWS_WITH_ISSUES"] += 1
            append(issue_sheet, ["UOM", r["_excel_row"], part, ext,
                                 "; ".join(problems)])
        append(uom_sheet, [r["_excel_row"], part, ext, internal,
                           rate, base, repeat, "; ".join(problems) or "OK",
                           r["rate"], r["exchange_base"], r["partno"],
                           r["measure_unit_ext"], r["measure_unit"],
                           *(r.get(field) for field in UOM_CONTEXT)])

    for label, value in [
        ("Source AMOS file", amos_path.name),
        ("Source UOM file", uom_path.name),
        ("AMOS rows", metrics["AMOS_DATA_ROWS"]),
        ("Unique AMOS part numbers", len(part_counts)),
        ("UOM rows", metrics["UOM_DATA_ROWS"]),
        ("Unique UOM part numbers", len(uom_by_part)),
        ("AMOS rows with issues", metrics["ROWS_WITH_ISSUES"]),
        ("UOM rows with issues", metrics["UOM_ROWS_WITH_ISSUES"]),
        ("BOTH_PRESENT", metrics["PRICE_BOTH_PRESENT"]),
        ("AVG_ONLY", metrics["PRICE_AVG_ONLY"]),
        ("PURCH_ONLY", metrics["PRICE_PURCH_ONLY"]),
        ("NO_USABLE_PRICE", metrics["PRICE_NO_USABLE_PRICE"]),
        ("No UOM rule (not automatically an error)", metrics["UOM_NO_UOM_RULE"]),
        ("Unit not listed in existing UOM rules", metrics["UOM_UNIT_NOT_LISTED_IN_UOM"]),
        ("Input workbooks modified", "NO"),
        ("STAT workbook modified", "NO"),
        ("Unit conversion applied", "NO - factors require company confirmation"),
        ("Price chosen for quotation", "NO - avg_price vs purch_price not confirmed"),
        ("Currency conversion applied", "NO"),
        ("Missing price fallback +20% applied", "NO - to be handled by pricing module"),
    ]:
        append(summary_sheet, [label, value])

    save_report(book, output, ROOT, (amos_path, uom_path), owner=DEFAULT_OUTPUT.name)
    print(f"DONE: {output}")
    print(f"AMOS rows: {metrics['AMOS_DATA_ROWS']:,}; UOM rules: {len(uom):,}")
    print(f"Rows missing both usable prices: {metrics['PRICE_NO_USABLE_PRICE']:,}")
    print("Original Excel/STAT and colleague's formulas were NOT changed.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Prepare Magnetic MRO AMOS data without changing STAT")
    parser.add_argument("--amos", type=Path, help="Path to AMOS_uus_süsteem_info.xlsx")
    parser.add_argument("--uom", type=Path, help="Path to AMOS_UOM.xlsx")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    amos_path = args.amos if args.amos else locate(AMOS_NAME)
    uom_path = args.uom if args.uom else locate(UOM_NAME)
    if not amos_path.is_file() or not uom_path.is_file():
        parser.error("AMOS or UOM source file does not exist")
    generate(amos_path, uom_path, args.output)


if __name__ == "__main__":
    main()
