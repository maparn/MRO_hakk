#!/usr/bin/env python3
"""Magnetic MRO: STAT vs AMOS UOM diagnostic report.

Run from the project directory: python scripts/sorter.py
Input files are located in excel_files/input or the project root.
STAT decisions contain one result per source row. Vastavused shows candidate rules.
No rule in this report authorizes a unit conversion or price calculation.
"""

from __future__ import annotations

import argparse
import warnings
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook
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
    from .report_io import literal_cell, save_report, validate_output
    from .uom_rules import annotate_rules, normalize_key as key, numeric_factor, rule_signature
else:
    from amos_prepare import AMOS_FIELDS, open_records, parse_price
    from report_io import literal_cell, save_report, validate_output
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


def locate_input(filename: str) -> Path:
    for path in (PROJECT / "excel_files" / "input" / filename, PROJECT / filename):
        if path.is_file():
            return path
    raise FileNotFoundError(f"Ei leia faili '{filename}'. Kontrolli excel_files/input kausta.")


def read_stat(path: Path, part_col: int, unit_col: int, *, quantity_col: int | None = None,
              task_card_col: int | None = None, reference_col: int | None = None):
    context_cols = [col for col in (quantity_col, task_card_col, reference_col) if col is not None]
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
            task_card = cell(task_card_col or 3)
            reference = cell(reference_col)
            if key(task_card):
                current_task_card = task_card
            if key(reference):
                current_reference = reference
            rows.append({"row": row_no, "partno": partno, "part_key": key(partno),
                         "unit": unit, "unit_key": key(unit), "stat_quantity": cell(quantity_col or 8),
                         "task_card": current_task_card, "reference": current_reference})
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


def append_sheet(wb, name, header, rows):
    ws = wb.create_sheet(name)
    ws.append(header)
    widths = [len(str(label)) for label in header]
    for index, row in enumerate(rows):
        ws.append([literal_cell(ws, value) for value in row])
        if index < 100:
            for i, value in enumerate(row):
                widths[i] = max(widths[i], len(str(value)) if value is not None else 0)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, cell in enumerate(ws[1]):
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i + 1)].width = min(52, max(15, widths[i] + 2))
    ws.row_dimensions[1].height = 32
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


def calculate_line(stat, uoms, amos_index, aliases):
    """Calculate one material line in EUR without guessing missing business inputs."""
    result = {
        "effective_stat_unit": aliases.get(stat["unit_key"], stat["unit_key"]),
        "converted_quantity": None,
        "conversion_rate": None,
        "conversion_base": None,
        "conversion_status": "NOT_CALCULATED",
        "unit_price_eur": None,
        "price_source": "AMOS_AVG_PRICE",
        "line_total_eur": None,
        "calculation_status": "NOT_CALCULATED",
        "avg_price_state": None,
    }
    if amos_index is None:
        result["calculation_status"] = "AMOS_NOT_REQUESTED"
        return result
    amos_rows = amos_index.get(stat["part_key"], [])
    if len(amos_rows) != 1:
        result["calculation_status"] = "AMOS_MULTIPLE" if amos_rows else "AMOS_NOT_FOUND"
        return result
    amos_row = amos_rows[0]
    price, price_state = parse_price(amos_row["avg_price"])
    result["avg_price_state"] = price_state
    if price_state == "ZERO":
        result["calculation_status"] = "AVG_PRICE_ZERO_SERVICE_FALLBACK"
        return result
    if price_state != "POSITIVE":
        result["calculation_status"] = f"AVG_PRICE_{price_state}"
        return result

    quantity = numeric_factor(stat.get("stat_quantity"))
    if quantity is None or quantity < 0:
        result["calculation_status"] = "MISSING_OR_INVALID_QUANTITY"
        return result
    source_unit = result["effective_stat_unit"]
    target_unit = key(amos_row["measure_unit"])
    if not source_unit or not target_unit:
        result["calculation_status"] = "MISSING_UNIT"
        return result
    if source_unit == target_unit:
        converted = quantity
        result.update(conversion_rate=1.0, conversion_base=1.0,
                      conversion_status="SAME_UNIT_OR_ALIAS")
    else:
        matching = [
            row for row in uoms
            if not row["rule_issues"]
            and key(row["external"]) == source_unit
            and key(row["internal"]) == target_unit
        ]
        if len(matching) != 1:
            result["calculation_status"] = "AMBIGUOUS_CONVERSION" if matching else "NO_CONVERSION_RULE"
            return result
        rule = matching[0]
        rate = numeric_factor(rule["rate"])
        base = numeric_factor(rule["exchange_base"])
        converted = quantity * rate / base
        result.update(conversion_rate=float(rate), conversion_base=float(base),
                      conversion_status="UOM_RULE_APPLIED")

    total = converted * Decimal(str(price))
    result.update(converted_quantity=float(converted), unit_price_eur=price,
                  line_total_eur=float(total), calculation_status="OK")
    return result


def build_service_summary(details, historical_avg, fallback_markup, approval_limit,
                          exchange_rate=None, exchange_date=None, exchange_status="NOT_CONFIGURED"):
    services = defaultdict(list)
    for stat, calculation in details:
        services[key(stat.get("task_card"))].append((stat, calculation))
    rows = []
    resolved_total = Decimal("0")
    for task_card, entries in services.items():
        reference = next((stat.get("reference") for stat, _ in entries if key(stat.get("reference"))), None)
        valid_totals = [Decimal(str(calc["line_total_eur"])) for _, calc in entries
                        if calc["line_total_eur"] is not None]
        calculated_sum = sum(valid_totals, Decimal("0"))
        all_ok = all(calc["calculation_status"] == "OK" for _, calc in entries)
        historical = historical_avg.get(task_card)
        issues = sorted({calc["calculation_status"] for _, calc in entries
                         if calc["calculation_status"] != "OK"})
        if not task_card:
            final, source = None, "MISSING_TASK_CARD"
        elif all_ok and historical is not None:
            if calculated_sum >= historical:
                final, source = calculated_sum, "NEW_PRICE_HIGHER_OR_EQUAL"
            else:
                final, source = historical, "OLD_PRICE_HIGHER"
        elif all_ok:
            final, source = calculated_sum, "NEW_PRICE_ONLY"
        elif historical is not None:
            final = historical * (Decimal("1") + fallback_markup / Decimal("100"))
            source = "OLD_PRICE_PLUS_MARKUP_FALLBACK"
        else:
            final, source = None, "REVIEW_REQUIRED"
        final_usd = final * exchange_rate if final is not None and exchange_rate is not None else None
        if final is None:
            approval_status = "PRICE_REVIEW_REQUIRED"
        elif approval_limit is not None and final > approval_limit:
            approval_status = "REVIEW_OVER_LIMIT"
        else:
            approval_status = "WITHIN_LIMIT"
        if final is not None:
            resolved_total += final
        rows.append([
            task_card or None, reference, len(entries), len(valid_totals), float(calculated_sum),
            float(historical) if historical is not None else None, float(fallback_markup),
            float(final) if final is not None else None, source,
            float(approval_limit) if approval_limit is not None else None, approval_status,
            float(exchange_rate) if exchange_rate is not None else None, exchange_date,
            exchange_status, float(final_usd) if final_usd is not None else None,
            "; ".join(issues),
        ])
    rows.sort(key=lambda row: str(row[0] or ""))
    resolved_usd = resolved_total * exchange_rate if exchange_rate is not None else None
    rows.append(["SUM_AVG", None, None, None, None, None, None, float(resolved_total),
                 "FINAL_TOTAL_EUR", float(approval_limit) if approval_limit is not None else None,
                 "TOTAL", float(exchange_rate) if exchange_rate is not None else None,
                 exchange_date, exchange_status,
                 float(resolved_usd) if resolved_usd is not None else None, None])
    return rows, float(resolved_total)


def build_report(stat_rows, uom_rows, output: Path, stat_col: int, unit_col: int, inputs=(),
                 *, amos_index=None, historical_avg=None, fallback_markup=Decimal("20"), aliases=None,
                 approval_limit=Decimal("500"), exchange_rate=None, exchange_date=None,
                 exchange_status="NOT_CONFIGURED"):
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
    for stat in stat_rows:
        part, source_row, unit = stat["part_key"], stat["row"], stat["unit"]
        uoms = by_part.get(part, [])
        candidates = [(r, match_side(unit, r)) for r in uoms]
        matched = [(r, side) for r, side in candidates if side != "NONE"]
        status = decide(stat, uoms, matched)
        calculation = calculate_line(stat, uoms, amos_index, aliases)
        if calculation["conversion_status"] == "SAME_UNIT_OR_ALIAS":
            status = "UNIT_ALIAS_APPLIED" if calculation["effective_stat_unit"] != stat["unit_key"] else "SAME_AMOS_UNIT"
        calculation_details.append((stat, calculation))
        decisions.append([source_row, stat["partno"], unit, len(uoms), len(matched),
                          sum(not r["rule_issues"] for r, side in matched), status,
                          calculation["conversion_status"] in {"SAME_UNIT_OR_ALIAS", "UOM_RULE_APPLIED"},
                          stat.get("stat_quantity"), stat.get("task_card"), stat.get("reference"),
                          *amos_evidence(amos_index, part),
                          calculation["effective_stat_unit"], calculation["converted_quantity"],
                          calculation["conversion_rate"], calculation["conversion_base"],
                          calculation["conversion_status"], calculation["unit_price_eur"],
                          calculation["price_source"], calculation["line_total_eur"],
                          calculation["calculation_status"]])
        # Every row here is a possible lookup alternative linked to one STAT row.
        for r, side in candidates:
            matches.append([part, source_row, unit, r["row"], r["external"], r["internal"],
                            r["rate"], r["exchange_base"], status, side, "; ".join(r["rule_issues"]) or "OK"])
        if not status.startswith("SINGLE_") and status not in {"SAME_AMOS_UNIT", "UNIT_ALIAS_APPLIED"}:
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

    stats = {"stat_rows": len(stat_rows), "uom_rows": len(uom_rows),
             "matched_partnos": len(stat_parts.intersection(by_part)),
             "candidate_pairs": len(matches), "review_rows": len(review), "uom_variants": len(variants)}
    service_rows, sum_avg = build_service_summary(
        calculation_details, historical_avg, fallback_markup, approval_limit,
        exchange_rate, exchange_date, exchange_status)
    stats["sum_avg_eur"] = sum_avg
    wb = Workbook()
    wb.remove(wb.active)
    service_sheet = append_sheet(wb, "Service Summary", ["task_card", "reference", "stat_rows",
        "calculated_rows", "calculated_line_sum_eur", "historical_avg_eur",
        "fallback_markup_percent", "final_service_eur", "final_price_source",
        "approval_limit_eur", "approval_status", "eur_usd_rate", "exchange_rate_date",
        "exchange_rate_status", "final_service_usd", "issues"], service_rows)
    for column in (5, 6, 8, 10):
        for cells in service_sheet.iter_cols(min_col=column, max_col=column, min_row=2):
            for cell in cells:
                cell.number_format = '#,##0.00 "EUR"'
    for cells in service_sheet.iter_cols(min_col=15, max_col=15, min_row=2):
        for cell in cells:
            cell.number_format = '#,##0.00 "USD"'
    summary = [
        ["STAT andmeread", len(decisions)], ["UOM andmeread", len(uom_rows)],
        ["STAT unikaalsed partno-d", len(stat_parts)], ["UOM unikaalsed partno-d", len(by_part)],
        ["Sobivad partno-d", stats["matched_partnos"]], ["STAT × UOM kandidaadipaarid", len(matches)],
        ["Ühikuid vajab kontrolli (STAT read)", len(review)], ["Ainult STAT (read)", len(only_stat)],
        ["Ainult UOM (read)", len(only_uom)], ["Mitu erinevat UOM reeglit (kogu UOM, partno-d)", len(variants)],
        ["Täpselt korduvad UOM reeglid (grupid)", len(duplicates)],
        ["Vastuolulised UOM reeglid (read)", len(conflicts)],
        ["STAT partno puudub", len(stat_missing)], ["UOM partno puudub", len(uom_missing)],
        ["STAT partno veerg", get_column_letter(stat_col)], ["STAT mõõtühiku veerg", get_column_letter(unit_col)],
        ["SUM_AVG EUR", sum_avg], ["Kinnitamise limiit EUR", float(approval_limit)],
        ["EUR/USD kurss", float(exchange_rate) if exchange_rate is not None else "PUUDUB"],
        ["Kursi kuupäev", exchange_date or "PUUDUB"], ["Kursi olek", exchange_status],
        ["Algfaile muudeti", "EI"],
        ["Ühikute teisendusi arvutati", "JAH, ainult kinnitatud reeglitega"],
        ["Hinnaarvutusi tehti", "JAH, AMOS avg_price on EUR"],
    ]
    append_sheet(wb, "Kokkuvote", ["Näitaja", "Tulemus"], summary)
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
        save_report(wb, output, PROJECT, inputs, owner=OUTPUT_DEFAULT.name)
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
    parser.add_argument("--reference-col", type=int, help="STAT Reference column (unset by default)")
    parser.add_argument("--historical-sheet", default="AVG", help="Historical service average sheet")
    parser.add_argument("--fallback-markup", type=Decimal, default=Decimal("20"),
                        help="Markup percent when AMOS avg_price is zero")
    parser.add_argument("--approval-limit", type=Decimal, default=Decimal("500"),
                        help="Review threshold in EUR; it never caps the price")
    parser.add_argument("--eur-usd-rate", type=Decimal, help="Official EUR to USD rate")
    parser.add_argument("--exchange-date", help="Exchange-rate date")
    parser.add_argument("--exchange-status", default="NOT_CONFIGURED")
    parser.add_argument("--unit-alias", action="append", default=[], help="Unit alias, for example LT=L")
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
                          task_card_col=args.task_card_col, reference_col=args.reference_col)
    amos_index = read_amos_for_stat(args.amos, stat_rows) if args.amos is not None else None
    uom_rows = read_uom(uom_path)
    historical_avg = read_historical_avg(stat_path, args.historical_sheet)
    aliases = parse_unit_aliases(args.unit_alias)
    stats = build_report(stat_rows, uom_rows, args.output, args.part_col, args.unit_col, inputs,
                         amos_index=amos_index, historical_avg=historical_avg,
                         fallback_markup=args.fallback_markup, aliases=aliases,
                         approval_limit=args.approval_limit, exchange_rate=args.eur_usd_rate,
                         exchange_date=args.exchange_date, exchange_status=args.exchange_status)
    print(f"Raport salvestatud: {args.output}")
    for label, value in stats.items():
        print(f"{label}: {value}")
    if len(stat_rows) < 2:
        print("HOIATUS: STAT failis on väga vähe andmeridu; testi täisandmestikuga.")


if __name__ == "__main__":
    main()
