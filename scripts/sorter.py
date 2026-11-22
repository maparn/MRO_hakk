#!/usr/bin/env python3
"""Magnetic MRO: auditable STAT material pricing report.

Run from the project directory: python scripts/sorter.py
Input files are located in excel_files/input or the project root.
The report calculates each material row, then sums by Reference and averages those
Reference totals by Task Card. Technical sheets retain the matching evidence.
"""

from __future__ import annotations

import argparse
import json
import warnings
from collections import Counter, defaultdict
from decimal import Decimal
from itertools import chain, islice
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

warnings.filterwarnings(
    "ignore",
    message="Workbook contains no default style, apply openpyxl's default",
    category=UserWarning,
    module="openpyxl.styles.stylesheet",
)

if __package__:
    from .amos_prepare import AMOS_FIELDS, open_records, parse_price
    from .report_io import save_report, validate_output
    from .uom_rules import annotate_rules, normalize_key as key, numeric_factor, rule_signature
else:
    from amos_prepare import AMOS_FIELDS, open_records, parse_price
    from report_io import save_report, validate_output
    from uom_rules import annotate_rules, normalize_key as key, numeric_factor, rule_signature

PROJECT = Path(__file__).resolve().parents[1]
STAT_FILENAME = "STAT_INITIAL - Copy.xlsx"
UOM_FILENAME = "AMOS_UOM.xlsx"
OUTPUT_DEFAULT = PROJECT / "excel_files" / "output" / "STAT_UOM_vordlusraport.xlsx"
STAT_SHEET = "INITIAL"
UOM_SHEET = "Sheet1"
STAT_PARTNO_COL = 6
STAT_UNIT_COL = 9
UOM_FIELDS = ("partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base")
STAT_PART_HEADERS = {"PN", "PARTNO", "PART NO", "PART NUMBER"}
STAT_UNIT_HEADERS = {"UNIT", "MEASURE_UNIT", "MEASURE UNIT", "UOM"}

STATUS_ET = {
    "CALCULATED_HISTORICAL_MAX": "Valitud kõrgem ajalooline ühikuhind",
    "CURRENT_AMOS_PRICE": "Arvutatud AMOS keskmise hinnaga",
    "FALLBACK_APPLIED": "Arvutatud ajaloolise ühikuhinnaga +20%",
    "MANUAL_REVIEW_MISSING_PART_NUMBER": "Puuduv tootekood, vajab käsitsi kontrolli",
    "MANUAL_REVIEW_AMOS_NOT_REQUESTED": "AMOS fail puudub, vajab käsitsi kontrolli",
    "MANUAL_REVIEW_AMOS_RECORD_MISSING": "Tootekood puudub AMOS failist, vajab käsitsi kontrolli",
    "MANUAL_REVIEW_MULTIPLE_AMOS_RECORDS": "AMOS failis on mitu vastet, vajab käsitsi kontrolli",
    "MANUAL_REVIEW_MISSING_HISTORICAL_PRICE": "ANDMED PUUDUVAD — ajalooline ühikuhind puudub",
    "MANUAL_REVIEW_INVALID_QUANTITY": "Kogus puudub või ei ole korrektne",
    "MANUAL_REVIEW_MISSING_UNIT": "Mõõtühik puudub, vajab käsitsi kontrolli",
    "MANUAL_REVIEW_AMBIGUOUS_UOM_RULE": "Mõõtühiku reegel on vastuoluline",
    "MANUAL_REVIEW_INVALID_UOM_FACTOR": "Mõõtühiku teisendustegur ei ole korrektne",
}


def locate_input(filename: str) -> Path:
    for path in (PROJECT / "excel_files" / "input" / filename, PROJECT / filename):
        if path.is_file():
            return path
    raise FileNotFoundError(f"Ei leia faili '{filename}'. Kontrolli excel_files/input kausta.")


def read_stat(path: Path, part_col: int, unit_col: int, *, quantity_col: int | None = None,
              task_card_col: int | None = None, reference_col: int | None = None,
              description_col: int | None = None, nrc_col: int | None = None,
              factor_col: int | None = None, historical_price_col: int | None = None):
    context_cols = [col for col in (
        quantity_col, task_card_col, reference_col, description_col, nrc_col,
        factor_col, historical_price_col,
    ) if col is not None]
    if any(not 1 <= col <= 16384 for col in context_cols):
        raise ValueError("Invalid STAT context column number.")
    if not all(1 <= col <= 16384 for col in (part_col, unit_col)) or part_col == unit_col:
        raise ValueError("STAT partno ja ühiku jaoks on vaja kahte erinevat kehtivat veergu.")
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if STAT_SHEET not in wb.sheetnames:
            raise ValueError(f"STAT leht '{STAT_SHEET}' puudub; olemas: {wb.sheetnames}")
        sheet = wb[STAT_SHEET]
        sheet.reset_dimensions()
        rows = []
        width = 0
        current_task_card = None
        current_reference = None
        for row_no, cells in enumerate(sheet.iter_rows(values_only=True), 1):
            width = max(width, len(cells))
            if not any(c is not None and str(c).strip() for c in cells):
                continue
            def cell(col):
                return cells[col - 1] if col is not None and col <= len(cells) else None

            partno, unit = cell(part_col), cell(unit_col)
            if key(partno) in STAT_PART_HEADERS and key(unit) in STAT_UNIT_HEADERS:
                continue
            # Ignore report footers such as GRAND TOTAL / SUM AVERAGE. Keep a
            # genuinely incomplete material row when it still has material fields.
            if not any(key(value) for value in (
                partno, unit, cell(quantity_col or 8), cell(description_col or 7),
            )):
                continue
            task_card = cell(task_card_col or 3)
            reference = cell(reference_col or 1)
            if key(task_card):
                current_task_card = task_card
            if key(reference):
                current_reference = reference
            rows.append({"row": row_no, "partno": partno, "part_key": key(partno),
                         "unit": unit, "unit_key": key(unit), "stat_quantity": cell(quantity_col or 8),
                         "task_card": current_task_card, "reference": current_reference,
                         "description": cell(description_col or 7), "nrc": cell(nrc_col or 4),
                         "factor": cell(factor_col or 5),
                         "historical_unit_price_eur": cell(historical_price_col or 11)})
        if not rows:
            raise ValueError(f"STAT lehel '{STAT_SHEET}' puuduvad andmeread.")
        if max(part_col, unit_col, *context_cols) > width:
            raise ValueError(f"STAT: valitud veerg on väljaspool lehe laiust ({width}).")
        return rows
    finally:
        wb.close()


def annotate_uom_rows(rows):
    for row in rows:
        row["measure_unit_ext"] = row["external"]
        row["measure_unit"] = row["internal"]
    annotate_rules(rows)


def read_uom(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if UOM_SHEET not in wb.sheetnames:
            raise ValueError(f"UOM leht '{UOM_SHEET}' puudub; olemas: {wb.sheetnames}")
        sheet = wb[UOM_SHEET]
        sheet.reset_dimensions()  # AMOS exports can declare A1:B4 regardless of real size.
        iterator = sheet.iter_rows(values_only=True)
        header = next(iterator, ())
        names = [key(value).lower() for value in header]
        counts = Counter(names)
        for field in UOM_FIELDS:
            if counts[field] != 1:
                raise ValueError(f"UOM: veerg '{field}' peab esinema täpselt üks kord (leitud {counts[field]}).")
        positions = {field: names.index(field) for field in UOM_FIELDS}
        rows = []
        for row_no, cells in enumerate(iterator, 2):
            if not any(c is not None and str(c).strip() for c in cells):
                continue
            row = {name: cells[i] if i < len(cells) else None for name, i in positions.items()}
            row.update(row=row_no, part_key=key(row["partno"]), external=row["measure_unit_ext"],
                       internal=row["measure_unit"])
            rows.append(row)
        if not rows:
            raise ValueError("UOM lehel puuduvad andmeread.")
        annotate_rules(rows)
        return rows
    finally:
        wb.close()


def append_sheet(wb, name, header, rows, number_formats=None):
    """Stream a styled worksheet so large Work Packages fit web memory limits."""
    ws = wb.create_sheet(name)
    ws.freeze_panes = "A2"
    number_formats = number_formats or {}
    iterator = iter(rows)
    sample = list(islice(iterator, 100))
    widths = [len(str(label)) for label in header]
    for row in sample:
        for i, value in enumerate(row):
            widths[i] = max(widths[i], len(str(value)) if value is not None else 0)
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = min(52, max(15, width + 2))
    styled_header = []
    for label in header:
        cell = WriteOnlyCell(ws, value=label)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        styled_header.append(cell)
    ws.append(styled_header)
    row_count = 0
    for row in chain(sample, iterator):
        cells = []
        for column, value in enumerate(row, 1):
            cell = WriteOnlyCell(ws, value=value)
            if isinstance(value, str):
                cell.data_type = "s"
            if column in number_formats and value is not None:
                cell.number_format = number_formats[column]
            cells.append(cell)
        ws.append(cells)
        row_count += 1
    ws.auto_filter.ref = f"A1:{get_column_letter(len(header))}{row_count + 1}"
    return ws


def match_side(stat_unit, uom):
    unit = key(stat_unit)
    if not unit:
        return "NONE"
    external = unit == key(uom["external"])
    internal = unit == key(uom["internal"])
    if external and internal:
        return "BOTH"
    if external:
        return "EXTERNAL"
    return "INTERNAL" if internal else "NONE"


def decide(stat, candidates, matched):
    """Classify observable evidence only. The target price unit is not supplied by STAT."""
    if not stat["part_key"]:
        return "MISSING_PARTNO"
    if not stat["unit_key"]:
        return "MISSING_STAT_UNIT"
    if not candidates:
        return "NO_UOM_RULE"
    if not matched:
        return "UNIT_NOT_LISTED"
    issues = {issue for row, side in matched for issue in row["rule_issues"]}
    if "CONFLICTING_UOM_RULES" in issues:
        return "CONFLICTING_UOM_RULES"
    if "INVALID_UOM_FACTOR" in issues or "MISSING_UOM_KEY" in issues:
        return "INVALID_MATCHING_RULE"
    if "REPEATED_UOM_RULE" in issues:
        return "REPEATED_UOM_RULE"
    if len(matched) > 1:
        return "AMBIGUOUS_UNIT_MATCH"
    return "SINGLE_" + matched[0][1] + "_MATCH"


def read_amos_for_stat(path, stat_rows):
    needed = {r["part_key"] for r in stat_rows if r["part_key"]}
    index = defaultdict(list)
    with open_records(path, AMOS_FIELDS) as records:
        for row in records():
            part = key(row["partno"])
            if part in needed:
                index[part].append(row)
    return index


def amos_evidence(index, part):
    if index is None:
        status = "NOT_REQUESTED"
    else:
        rows = index.get(part, [])
        if len(rows) == 1:
            row = rows[0]
            avg, avg_state = parse_price(row["avg_price"])
            purch, purch_state = parse_price(row["purch_price"])
            return ["UNIQUE", row["_excel_row"], row["measure_unit"], avg, avg_state, purch, purch_state]
        status = "MULTIPLE" if rows else "NO_MATCH"
    return [status, None, None, None, None, None, None]


def parse_unit_aliases(items):
    aliases = {}
    for item in items:
        if "=" not in item:
            raise ValueError(f"Ühiku alias peab olema kujul LT=L: {item}")
        source, target = (key(value) for value in item.split("=", 1))
        if not source or not target:
            raise ValueError(f"Ühiku alias peab sisaldama kahte ühikut: {item}")
        aliases[source] = target
    return aliases


def parse_physical_price_conversions(records):
    """Validate explicitly approved AMOS-price-unit to STAT-price-unit mappings."""
    conversions = {}
    for record in records or ():
        if not record.get("enabled", False):
            continue
        amos_unit = key(record.get("amos_unit"))
        stat_unit = key(record.get("stat_unit"))
        multiplier = numeric_factor(record.get("price_multiplier"))
        approval = str(record.get("approval_reference") or "").strip()
        if not amos_unit or not stat_unit or multiplier is None or multiplier <= 0:
            raise ValueError("Approved physical conversion requires units and a positive price_multiplier")
        if not approval or approval.upper().startswith("PENDING"):
            raise ValueError("Enabled physical conversion requires approval_reference")
        conversion_key = (amos_unit, stat_unit)
        if conversion_key in conversions:
            raise ValueError(f"Duplicate physical conversion: {amos_unit}>{stat_unit}")
        conversions[conversion_key] = {
            "price_multiplier": multiplier, "approval_reference": approval,
        }
    return conversions


def read_historical_avg(path: Path, sheet_name: str):
    """Read the historical service average keyed by Task Card."""
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return {}
        sheet = wb[sheet_name]
        sheet.reset_dimensions()
        iterator = sheet.iter_rows(values_only=True)
        header = next(iterator, ())
        names = [key(value).replace(" ", "").lower() for value in header]
        if "taskcard" not in names or "avg" not in names:
            raise ValueError(f"{sheet_name}: oodatud veerud TaskCard ja avg puuduvad")
        task_index, avg_index = names.index("taskcard"), names.index("avg")
        result = {}
        for row_no, cells in enumerate(iterator, 2):
            task_card = key(cells[task_index] if task_index < len(cells) else None)
            value = cells[avg_index] if avg_index < len(cells) else None
            amount = numeric_factor(value)
            if not task_card or amount is None or amount < 0:
                continue
            if task_card in result and result[task_card] != amount:
                raise ValueError(f"{sheet_name}: Task Card {task_card} ajalooline AVG kordub erineva väärtusega")
            result[task_card] = amount
        return result
    finally:
        wb.close()


def calculate_line(stat, uoms, amos_index, aliases, fallback_markup=Decimal("20"),
                   physical_conversions=None):
    """Calculate one material row in EUR using current price or row-level legacy fallback."""
    historical_price, historical_state = parse_price(stat.get("historical_unit_price_eur"))
    result = {
        "effective_stat_unit": aliases.get(stat["unit_key"], stat["unit_key"]),
        "converted_quantity": None,
        "conversion_rate": None,
        "conversion_base": None,
        "conversion_status": "NOT_CALCULATED",
        "price_conversion_factor": None,
        "conversion_source": None,
        "unit_price_eur": None,
        "price_source": None,
        "line_total_eur": None,
        "calculation_status": "NOT_CALCULATED",
        "avg_price_state": None,
        "current_price_status": "MANUAL_REVIEW_REQUIRED",
        "fallback_status": "NOT_APPLIED",
        "historical_unit_price_eur": historical_price,
        "historical_price_state": historical_state,
        "selected_unit_price_eur": None,
        "applied_markup_percent": None,
        "fallback_reason": None,
        "review_required": True,
    }

    quantity = numeric_factor(stat.get("stat_quantity"))

    def unresolved(status, reason=None):
        result.update(calculation_status=status, fallback_reason=reason,
                      price_source="MANUAL_REVIEW_REQUIRED",
                      fallback_status=("MISSING_HISTORICAL_PRICE"
                                       if status == "MANUAL_REVIEW_MISSING_HISTORICAL_PRICE"
                                       else "MANUAL_REVIEW_REQUIRED"),
                      review_required=True)
        return result

    def historical_fallback(reason):
        if historical_state != "POSITIVE":
            return unresolved("MANUAL_REVIEW_MISSING_HISTORICAL_PRICE", reason)
        if quantity is None or quantity < 0:
            return unresolved("MANUAL_REVIEW_INVALID_QUANTITY", reason)
        selected = Decimal(str(historical_price)) * (
            Decimal("1") + fallback_markup / Decimal("100"))
        total = quantity * selected
        result.update(
            converted_quantity=float(quantity), conversion_status="NOT_APPLICABLE_FALLBACK",
            price_conversion_factor=1.0, conversion_source="HISTORICAL_FALLBACK",
            unit_price_eur=float(selected), selected_unit_price_eur=float(selected),
            price_source="HISTORICAL_PRICE_PLUS_20_PERCENT",
            applied_markup_percent=float(fallback_markup), fallback_reason=reason,
            line_total_eur=float(total), calculation_status="FALLBACK_APPLIED",
            fallback_status="HISTORICAL_FALLBACK_APPLIED", review_required=False,
        )
        return result

    def select_max_current_price(current_stat_unit_price, converted_quantity):
        # Compare historical and AMOS prices in the same STAT unit.
        historical = (
            Decimal(str(historical_price))
            if historical_state == "POSITIVE"
            else None
        )

        use_old = (
            historical is not None
            and historical > current_stat_unit_price
        )

        chosen = historical if use_old else current_stat_unit_price

        result.update(
            converted_quantity=float(converted_quantity),
            unit_price_eur=float(chosen),
            selected_unit_price_eur=float(chosen),
            price_source=(
                "HISTORICAL_UNIT_PRICE"
                if use_old else "CURRENT_AMOS_PRICE"
            ),
            applied_markup_percent=0.0,
            line_total_eur=float(quantity * chosen),
            calculation_status=(
                "CALCULATED_HISTORICAL_MAX"
                if use_old else "CURRENT_AMOS_PRICE"
            ),
            review_required=False,
        )
        return result

    if not stat.get("part_key"):
        return unresolved("MANUAL_REVIEW_MISSING_PART_NUMBER")
    if amos_index is None:
        return unresolved("MANUAL_REVIEW_AMOS_NOT_REQUESTED")
    amos_rows = amos_index.get(stat["part_key"], [])
    if not amos_rows:
        result["current_price_status"] = "MATERIAL_NOT_FOUND"
        return unresolved("MANUAL_REVIEW_AMOS_RECORD_MISSING", "AMOS_RECORD_MISSING")
    if len(amos_rows) != 1:
        return unresolved("MANUAL_REVIEW_MULTIPLE_AMOS_RECORDS", "MULTIPLE_AMOS_RECORDS")
    amos_row = amos_rows[0]
    price, price_state = parse_price(amos_row["avg_price"])
    result["avg_price_state"] = price_state
    result["current_price_status"] = {
        "POSITIVE": "CURRENT_PRICE_VALID", "ZERO": "CURRENT_PRICE_ZERO",
        "MISSING": "CURRENT_PRICE_MISSING", "INVALID": "CURRENT_PRICE_INVALID",
        "NEGATIVE": "CURRENT_PRICE_INVALID",
    }[price_state]
    if price_state != "POSITIVE":
        reason = {"ZERO": "CURRENT_PRICE_ZERO", "MISSING": "MISSING_CURRENT_PRICE",
                  "INVALID": "INVALID_PRICE", "NEGATIVE": "INVALID_PRICE"}[price_state]
        return historical_fallback(reason)

    if quantity is None or quantity < 0:
        return unresolved("MANUAL_REVIEW_INVALID_QUANTITY")
    source_unit = result["effective_stat_unit"]
    target_unit = key(amos_row["measure_unit"])
    if not source_unit or not target_unit:
        return unresolved("MANUAL_REVIEW_MISSING_UNIT")
    if source_unit == target_unit:
        converted = quantity
        price_factor = Decimal("1")
        result.update(conversion_rate=1.0, conversion_base=1.0,
                      conversion_source=("UNIT_ALIAS" if source_unit != stat["unit_key"]
                                         else "IDENTICAL_UNIT"),
                      conversion_status=("UNIT_ALIAS_APPLIED"
                                         if source_unit != stat["unit_key"] else "SAME_UNIT"))
    else:
        candidates = [
            row for row in uoms
            if key(row["external"]) == source_unit
            and key(row["internal"]) == target_unit
        ]
        matching = [row for row in candidates if not row["rule_issues"]]
        if len(matching) != 1:
            if not candidates:
                physical = (physical_conversions or {}).get((target_unit, source_unit))
                if physical is None:
                    return historical_fallback("MISSING_UOM_RULE")
                price_factor = physical["price_multiplier"]
                converted = quantity
                result.update(
                    conversion_status="APPROVED_PHYSICAL_PRICE_CONVERSION",
                    price_conversion_factor=float(price_factor),
                    conversion_source="APPROVED_PHYSICAL_CONFIG",
                    fallback_reason=None,
                )
                selected = Decimal(str(price)) * price_factor
                return select_max_current_price(selected, quantity)
            if len(candidates) > 1 or any("CONFLICTING_UOM_RULES" in row["rule_issues"]
                                        for row in candidates):
                return unresolved("MANUAL_REVIEW_AMBIGUOUS_UOM_RULE", "AMBIGUOUS_UOM_RULE")
            return unresolved("MANUAL_REVIEW_INVALID_UOM_FACTOR", "INVALID_UOM_FACTOR")
        rule = matching[0]
        rate = numeric_factor(rule["rate"])
        base = numeric_factor(rule["exchange_base"])
        price_factor = rate / base
        converted = quantity * price_factor
        result.update(conversion_rate=float(rate), conversion_base=float(base),
                      price_conversion_factor=float(price_factor),
                      conversion_source="AMOS_UOM_RULE", conversion_status="CONVERSION_APPLIED")

    selected = Decimal(str(price)) * price_factor
    return select_max_current_price(selected, converted)


def build_aggregations(details, historical_avg, exchange_rate=None):
    """Aggregate rows by Reference, average by Task Card, then select max(old, new)."""
    grouped = defaultdict(list)
    for stat, calculation in details:
        grouped[(key(stat.get("reference")), key(stat.get("task_card")))].append((stat, calculation))

    reference_rows = []
    reference_records = []
    for (reference_key, task_key), entries in grouped.items():
        reference = next((s.get("reference") for s, _ in entries if key(s.get("reference"))), None)
        task_card = next((s.get("task_card") for s, _ in entries if key(s.get("task_card"))), None)
        totals = [Decimal(str(c["line_total_eur"])) for _, c in entries
                  if c["line_total_eur"] is not None]
        fallback_rows = sum(c["price_source"] == "HISTORICAL_PRICE_PLUS_20_PERCENT"
                            for _, c in entries)
        unresolved_rows = sum(c["line_total_eur"] is None for _, c in entries)
        subtotal = sum(totals, Decimal("0"))
        complete_total = subtotal if unresolved_rows == 0 and reference_key else None
        status = "COMPLETE" if complete_total is not None else "INCOMPLETE"
        record = {
            "reference": reference, "task_card": task_card, "material_rows": len(entries),
            "calculated_rows": len(totals), "fallback_rows": fallback_rows,
            "unresolved_rows": unresolved_rows, "calculated_subtotal_eur": subtotal,
            "reference_total_eur": complete_total, "status": status,
        }
        reference_records.append(record)
        reference_rows.append([
            reference, task_card, len(entries), len(totals), fallback_rows, unresolved_rows,
            float(subtotal), float(complete_total) if complete_total is not None else None, status,
        ])
    reference_rows.sort(key=lambda row: (str(row[0] or ""), str(row[1] or "")))

    by_task = defaultdict(list)
    for record in reference_records:
        by_task[key(record["task_card"])].append(record)
    task_rows = []
    new_sum_avg_eur = Decimal("0")
    final_sum_eur = Decimal("0")
    incomplete_tasks = 0
    for task_key, records in by_task.items():
        task_card = next((r["task_card"] for r in records if key(r["task_card"])), None)
        complete = [r["reference_total_eur"] for r in records if r["reference_total_eur"] is not None]
        incomplete = len(records) - len(complete)
        total = sum(complete, Decimal("0"))
        average = total / len(complete) if complete else None
        if average is not None:
            new_sum_avg_eur += average
        if incomplete:
            incomplete_tasks += 1
        historical = historical_avg.get(task_key)
        status = "COMPLETE" if incomplete == 0 and task_key else "INCOMPLETE"
        difference = average - historical if average is not None and historical is not None else None
        final_price = None
        selected_source = None
        if average is not None and historical is not None:
            if average >= historical:
                final_price, selected_source = average, "NEW"
            else:
                final_price, selected_source = historical, "OLD"
        elif average is not None:
            final_price, selected_source = average, "NEW"
        elif historical is not None:
            final_price, selected_source = historical, "OLD"
        if final_price is not None:
            final_sum_eur += final_price
        usd = final_price * exchange_rate if final_price is not None and exchange_rate is not None else None
        task_rows.append([
            task_card, len(records), len(complete), incomplete, float(total),
            float(average) if average is not None else None,
            float(historical) if historical is not None else None,
            float(final_price) if final_price is not None else None, selected_source,
            float(difference) if difference is not None else None,
            float(usd) if usd is not None else None, status,
        ])
    task_rows.sort(key=lambda row: str(row[0] or ""))
    final_sum_usd = final_sum_eur * exchange_rate if exchange_rate is not None else None
    overall_status = "COMPLETE" if incomplete_tasks == 0 else "INCOMPLETE"
    task_rows.append([
        "SUM_AVG", None, None, incomplete_tasks, None, float(new_sum_avg_eur), None,
        float(final_sum_eur), None, None,
        float(final_sum_usd) if final_sum_usd is not None else None, overall_status,
    ])
    return reference_rows, task_rows, float(final_sum_eur), overall_status


def build_report(stat_rows, uom_rows, output: Path, stat_col: int, unit_col: int, inputs=(),
                 *, amos_index=None, historical_avg=None, fallback_markup=Decimal("20"), aliases=None,
                 approval_limit=Decimal("500"), exchange_rate=None, exchange_date=None,
                 exchange_status="NOT_CONFIGURED", physical_conversions=None,
                 project_root=PROJECT):
    historical_avg = historical_avg or {}
    aliases = aliases or {}
    if not all("rule_issues" in row and "measure_unit_ext" in row for row in uom_rows):
        annotate_uom_rows(uom_rows)
    stat_parts = {r["part_key"] for r in stat_rows if r["part_key"]}
    by_part = defaultdict(list)
    by_external = defaultdict(list)
    by_signature = defaultdict(list)
    uom_missing = []
    for row in uom_rows:
        if not row["part_key"]:
            uom_missing.append([row["row"], row["external"], row["internal"]])
            continue
        by_part[row["part_key"]].append(row)
        sig = rule_signature(row)
        if all(sig[:2]):
            by_external[sig[:2]].append(row)
        if all(sig[:3]):
            by_signature[sig].append(row)

    candidate_count = sum(len(by_part.get(r["part_key"], ())) for r in stat_rows)
    if max(candidate_count, len(stat_rows), len(uom_rows)) > 1_048_575:
        raise ValueError("STAT, UOM or candidate rows exceed the Excel worksheet limit with a header.")

    decisions, matches, review, only_stat, stat_missing, calculation_details = [], [], [], [], [], []
    material_rows, exception_rows = [], []
    for stat in stat_rows:
        part, source_row, unit = stat["part_key"], stat["row"], stat["unit"]
        uoms = by_part.get(part, [])
        candidates = [(r, match_side(unit, r)) for r in uoms]
        matched = [(r, side) for r, side in candidates if side != "NONE"]
        status = decide(stat, uoms, matched)
        calculation = calculate_line(stat, uoms, amos_index, aliases, fallback_markup,
                                     physical_conversions)
        if calculation["conversion_status"] in {"SAME_UNIT", "UNIT_ALIAS_APPLIED"}:
            status = calculation["conversion_status"]
        calculation_details.append((stat, calculation))
        evidence = amos_evidence(amos_index, part)
        decisions.append([source_row, stat["partno"], unit, len(uoms), len(matched),
                          sum(not r["rule_issues"] for r, side in matched), status,
                          calculation["conversion_status"] in {"SAME_UNIT", "UNIT_ALIAS_APPLIED", "CONVERSION_APPLIED"},
                          stat.get("stat_quantity"), stat.get("task_card"), stat.get("reference"),
                          *evidence,
                          calculation["effective_stat_unit"], calculation["converted_quantity"],
                          calculation["conversion_rate"], calculation["conversion_base"],
                          calculation["conversion_status"], calculation["unit_price_eur"],
                          calculation["price_source"], calculation["line_total_eur"],
                          calculation["calculation_status"]])
        selected_price = calculation["selected_unit_price_eur"]
        approval_status = (
            "REVIEW_OVER_THRESHOLD"
            if selected_price is not None and approval_limit is not None
            and Decimal(str(selected_price)) > approval_limit
            else "WITHIN_THRESHOLD" if selected_price is not None else "NOT_CALCULATED"
        )
        material_row = [
            source_row, stat.get("reference"), stat.get("task_card"), stat.get("nrc"),
            stat.get("factor"), stat.get("partno"), stat.get("description"),
            stat.get("stat_quantity"), stat.get("unit"),
            calculation["historical_unit_price_eur"], calculation["historical_price_state"],
            calculation["current_price_status"], calculation["fallback_status"],
            evidence[0], evidence[1], evidence[2], evidence[3], evidence[4],
            calculation["effective_stat_unit"], calculation["conversion_rate"],
            calculation["conversion_base"], calculation["converted_quantity"],
            calculation["price_conversion_factor"], calculation["conversion_source"],
            calculation["conversion_status"],
            (calculation["selected_unit_price_eur"]
             if calculation["selected_unit_price_eur"] is not None else "ANDMED PUUDUVAD"),
            calculation["price_source"],
            calculation["applied_markup_percent"], calculation["fallback_reason"],
            calculation["line_total_eur"], "EUR", calculation["calculation_status"],
            STATUS_ET.get(calculation["calculation_status"], calculation["calculation_status"]),
            float(approval_limit) if approval_limit is not None else None, approval_status,
            calculation["review_required"],
        ]
        material_rows.append(material_row)
        if (calculation["review_required"]
                or calculation["price_source"] == "HISTORICAL_PRICE_PLUS_20_PERCENT"
                or approval_status == "REVIEW_OVER_THRESHOLD"):
            exception_rows.append(material_row)
        # Every row here is a possible lookup alternative linked to one STAT row.
        for r, side in candidates:
            matches.append([part, source_row, unit, r["row"], r["external"], r["internal"],
                            r["rate"], r["exchange_base"], status, side, "; ".join(r["rule_issues"]) or "OK"])
        if not status.startswith("SINGLE_") and status not in {"SAME_UNIT", "UNIT_ALIAS_APPLIED"}:
            review.append([part, source_row, unit,
                           "; ".join(f"{r['external']}→{r['internal']}" for r in uoms), status])
        if not part:
            stat_missing.append([source_row, unit])
        elif not uoms:
            only_stat.append([part, source_row, unit])

    only_uom, variants = [], []
    for part, rules in sorted(by_part.items()):
        if part not in stat_parts:
            only_uom.extend([part, r["row"], r["external"], r["internal"]] for r in rules)
        distinct = {rule_signature(r) for r in rules}
        if len(distinct) > 1:
            variants.append([part, len(distinct), ", ".join(
                f"{r['external']}→{r['internal']} ({r['row']})" for r in rules)])
    duplicates = []
    for sig, rules in by_signature.items():
        if len(rules) > 1:
            r = rules[0]
            duplicates.append([*sig[:3], len(rules), ", ".join(str(x["row"]) for x in rules),
                               r["rate"], r["exchange_base"]])
    conflicts = []
    for (part, external), rules in sorted(by_external.items()):
        if "CONFLICTING_UOM_RULES" in rules[0]["rule_issues"]:
            conflicts.extend([part, external, r["internal"], r["rate"], r["exchange_base"], r["row"]]
                             for r in rules)
    quality = ([r["row"], r["partno"], r["external"], r["internal"], r["rate"], r["exchange_base"],
                "; ".join(r["rule_issues"])] for r in uom_rows if r["rule_issues"])

    calculated_count = sum(c["line_total_eur"] is not None for _, c in calculation_details)
    fallback_count = sum(c["price_source"] == "HISTORICAL_PRICE_PLUS_20_PERCENT"
                         for _, c in calculation_details)
    unresolved_count = len(stat_rows) - calculated_count
    missing_uom_count = sum(c["fallback_reason"] == "MISSING_UOM_RULE"
                            for _, c in calculation_details)
    reference_rows, task_rows, sum_avg, overall_status = build_aggregations(
        calculation_details, historical_avg, exchange_rate)
    sum_avg_usd = Decimal(str(sum_avg)) * exchange_rate if exchange_rate is not None else None
    stats = {"stat_rows": len(stat_rows), "uom_rows": len(uom_rows),
             "matched_partnos": len(stat_parts.intersection(by_part)),
             "candidate_pairs": len(matches), "review_rows": len(review), "uom_variants": len(variants),
             "calculated_rows": calculated_count, "fallback_rows": fallback_count,
             "unresolved_rows": unresolved_count, "missing_uom_rows": missing_uom_count,
             "final_total_eur": sum_avg, "sum_avg_eur": sum_avg,
             "overall_status": overall_status}
    wb = Workbook(write_only=True)
    summary = [
        ["Töödeldud materjaliread", len(stat_rows)], ["Arvutatud read", calculated_count],
        ["Ajaloolise hinna fallback read", fallback_count], ["Puuduva UOM reegli read", missing_uom_count],
        ["Lahendamata read", unresolved_count], ["Reference grupid", len(reference_rows)],
        ["Task Card grupid", max(0, len(task_rows) - 1)], ["Lõplik summa EUR", sum_avg],
        ["Lõplik summa USD", float(sum_avg_usd) if sum_avg_usd is not None else None],
        ["Koondstaatuse olek", overall_status],
        ["EUR/USD kurss", float(exchange_rate) if exchange_rate is not None else None],
        ["Kursi kuupäev", exchange_date], ["Kursi olek", exchange_status],
        ["Ajaloolise hinna juurdehindlus %", float(fallback_markup)],
        ["Ühikuhinna kontrolllävi EUR", float(approval_limit) if approval_limit is not None else None],
        ["Kontrolllävi ei ole hinnalagi", "JAH"], ["Algfaile muudeti", "EI"],
    ]
    append_sheet(wb, "Executive Summary", ["Näitaja", "Tulemus"], summary)
    material_headers = [
        "source_row", "reference", "task_card", "nrc", "factor", "partno", "description",
        "original_quantity", "original_unit", "historical_unit_price_eur", "historical_price_state",
        "current_price_status", "fallback_status",
        "amos_lookup_status", "amos_source_row", "amos_unit", "amos_avg_price_eur", "amos_price_state",
        "effective_source_unit", "conversion_rate", "conversion_base", "converted_quantity",
        "price_conversion_factor", "conversion_source", "conversion_status",
        "selected_unit_price_eur", "price_source", "applied_markup_percent", "fallback_reason",
        "final_line_eur", "currency", "calculation_status", "status_et",
        "approval_threshold_eur", "approval_status", "review_required",
    ]
    eur_format = '#,##0.00 "EUR"'
    append_sheet(wb, "Material Calculations", material_headers, material_rows,
                 {10: eur_format, 17: eur_format, 26: eur_format,
                  30: eur_format, 34: eur_format})
    append_sheet(wb, "Reference Summary", [
        "reference", "task_card", "material_rows", "calculated_rows", "fallback_rows",
        "unresolved_rows", "calculated_subtotal_eur", "reference_total_eur", "status",
    ], reference_rows)
    append_sheet(wb, "Task Card Summary", [
        "task_card", "reference_count", "complete_references", "incomplete_references",
        "reference_totals_sum_eur", "new_task_card_avg_eur", "old_task_card_avg_eur",
        "final_task_card_price_eur", "selected_price_source",
        "difference_new_minus_old_eur", "final_task_card_price_usd", "status",
    ], task_rows, {5: eur_format, 6: eur_format, 7: eur_format, 8: eur_format,
                   10: eur_format, 11: '#,##0.00 "USD"'})
    append_sheet(wb, "Exceptions", material_headers, exception_rows)
    append_sheet(wb, "STAT decisions", ["stat_row", "partno", "stat_unit", "part_rule_rows",
        "unit_match_rows", "clean_unit_match_rows", "unit_decision", "conversion_approved",
        "stat_quantity", "task_card", "reference", "amos_lookup_status", "amos_source_row", "amos_unit",
        "avg_price", "avg_price_state", "purch_price", "purch_price_state",
        "effective_stat_unit", "converted_quantity", "conversion_rate", "conversion_base",
        "conversion_status", "unit_price_eur", "price_source", "line_total_eur",
        "calculation_status"], decisions)
    append_sheet(wb, "Vastavused", ["partno", "STAT rida", "STAT ühik", "UOM rida", "UOM ext ühik",
        "UOM sisemine ühik", "rate", "exchange_base", "Olek", "match_side", "rule_issues"], matches)
    append_sheet(wb, "Unit ulevaatus", ["partno", "STAT rida", "STAT ühik", "UOM variandid", "Märkus"], review)
    append_sheet(wb, "Ainult STAT", ["partno", "STAT rida", "STAT ühik"], only_stat)
    append_sheet(wb, "Ainult UOM", ["partno", "UOM rida", "UOM ext", "UOM sisemine"], only_uom)
    append_sheet(wb, "UOM variandid", ["partno", "Erinevaid reegleid", "Ühikud ja read"], variants)
    append_sheet(wb, "Duplikaadid", ["partno", "Väline ühik", "Sisemine ühik", "Kordusi", "Read",
        "rate", "exchange_base"], duplicates)
    append_sheet(wb, "UOM conflicts", ["partno", "measure_unit_ext", "measure_unit", "rate",
        "exchange_base", "source_row"], conflicts)
    append_sheet(wb, "UOM issues", ["source_row", "partno", "measure_unit_ext", "measure_unit",
        "rate", "exchange_base", "rule_issues"], quality)
    append_sheet(wb, "STAT puuduv partno", ["STAT rida", "STAT ühik"], stat_missing)
    append_sheet(wb, "UOM puuduv partno", ["UOM rida", "UOM ext", "UOM sisemine"], uom_missing)
    try:
        save_report(wb, output, project_root, inputs, owner=OUTPUT_DEFAULT.name)
    finally:
        wb.close()
    return stats


def main():
    parser = argparse.ArgumentParser(description="Kontrolli STAT ja UOM mõõtühikute vastavusi")
    parser.add_argument("--stat", type=Path, help="STAT Excel faili tee")
    parser.add_argument("--uom", type=Path, help="AMOS UOM Excel faili tee")
    parser.add_argument("--amos", type=Path, help="Optional original AMOS export for price candidates")
    parser.add_argument("--quantity-col", type=int, help="STAT quantity column (H=8)")
    parser.add_argument("--task-card-col", type=int, help="STAT Task Card column (C=3)")
    parser.add_argument("--reference-col", type=int, default=1, help="STAT Reference column (A=1)")
    parser.add_argument("--description-col", type=int, default=7, help="STAT description column (G=7)")
    parser.add_argument("--nrc-col", type=int, default=4, help="STAT NRC column (D=4)")
    parser.add_argument("--factor-col", type=int, default=5, help="STAT factor column (E=5)")
    parser.add_argument("--historical-price-col", type=int, default=11,
                        help="STAT historical EUR unit price column (K=11)")
    parser.add_argument("--historical-sheet", default="AVG", help="Historical service average sheet")
    parser.add_argument("--fallback-markup", type=Decimal, default=Decimal("20"),
                        help="Markup percent when AMOS avg_price is zero")
    parser.add_argument("--approval-limit", type=Decimal, default=Decimal("500"),
                        help="Review threshold in EUR; it never caps the price")
    parser.add_argument("--eur-usd-rate", type=Decimal, help="Official EUR to USD rate")
    parser.add_argument("--exchange-date", help="Exchange-rate date")
    parser.add_argument("--exchange-status", default="NOT_CONFIGURED")
    parser.add_argument("--unit-alias", action="append", default=[], help="Unit alias, for example LT=L")
    parser.add_argument("--physical-price-conversion-json", action="append", default=[],
                        help="Enabled approved physical price conversion as JSON")
    parser.add_argument("--output", type=Path, default=OUTPUT_DEFAULT, help="Raporti Excel faili tee")
    parser.add_argument("--part-col", type=int, default=STAT_PARTNO_COL, help="STAT partno veerg (F=6)")
    parser.add_argument("--unit-col", type=int, default=STAT_UNIT_COL, help="STAT ühiku veerg (I=9)")
    args = parser.parse_args()
    stat_path = args.stat or locate_input(STAT_FILENAME)
    uom_path = args.uom or locate_input(UOM_FILENAME)
    if not stat_path.is_file() or not uom_path.is_file():
        parser.error("Sisendfaili ei leitud. Kontrolli --stat ja --uom teid.")
    if args.amos is not None and not args.amos.is_file():
        parser.error("AMOS source file not found")
    inputs = [stat_path, uom_path]
    if args.amos is not None:
        inputs.append(args.amos)
    validate_output(args.output, PROJECT, inputs, owner=OUTPUT_DEFAULT.name)
    stat_rows = read_stat(stat_path, args.part_col, args.unit_col, quantity_col=args.quantity_col,
                          task_card_col=args.task_card_col, reference_col=args.reference_col,
                          description_col=args.description_col, nrc_col=args.nrc_col,
                          factor_col=args.factor_col, historical_price_col=args.historical_price_col)
    amos_index = read_amos_for_stat(args.amos, stat_rows) if args.amos is not None else None
    uom_rows = read_uom(uom_path)
    historical_avg = read_historical_avg(stat_path, args.historical_sheet)
    aliases = parse_unit_aliases(args.unit_alias)
    physical_conversions = parse_physical_price_conversions(
        [json.loads(item) for item in args.physical_price_conversion_json])
    stats = build_report(stat_rows, uom_rows, args.output, args.part_col, args.unit_col, inputs,
                         amos_index=amos_index, historical_avg=historical_avg,
                         fallback_markup=args.fallback_markup, aliases=aliases,
                         approval_limit=args.approval_limit, exchange_rate=args.eur_usd_rate,
                         exchange_date=args.exchange_date, exchange_status=args.exchange_status,
                         physical_conversions=physical_conversions)
    print(f"Raport salvestatud: {args.output}")
    for label, value in stats.items():
        print(f"{label}: {value}")
    if len(stat_rows) < 2:
        print("HOIATUS: STAT failis on väga vähe andmeridu; testi täisandmestikuga.")


if __name__ == "__main__":
    main()
