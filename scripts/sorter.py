#!/usr/bin/env python3
"""Magnetic MRO: STAT vs AMOS UOM diagnostic report.

Run from the project directory: python scripts/sorter.py
Input files are located in excel_files/input or the project root.
STAT decisions contain one result per source row. Vastavused shows candidate rules.
No rule in this report authorizes a unit conversion or price calculation.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

if __package__:
    from .amos_prepare import AMOS_FIELDS, open_records, parse_price
    from .report_io import literal_cell, save_report, validate_output
    from .uom_rules import annotate_rules, normalize_key as key, rule_signature
else:
    from amos_prepare import AMOS_FIELDS, open_records, parse_price
    from report_io import literal_cell, save_report, validate_output
    from uom_rules import annotate_rules, normalize_key as key, rule_signature

PROJECT = Path(__file__).resolve().parents[1]
STAT_FILENAME = "STAT_INITIAL - Copy.xlsx"
UOM_FILENAME = "AMOS_UOM.xlsx"
OUTPUT_DEFAULT = PROJECT / "excel_files" / "output" / "STAT_UOM_vordlusraport.xlsx"
STAT_SHEET = "INITIAL"
UOM_SHEET = "Sheet1"
STAT_PARTNO_COL = 6
STAT_UNIT_COL = 9
UOM_FIELDS = ("partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base")


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
        for row_no, cells in enumerate(sheet.iter_rows(values_only=True), 1):
            width = max(width, len(cells))
            if not any(c is not None and str(c).strip() for c in cells):
                continue
            def cell(col):
                return cells[col - 1] if col is not None and col <= len(cells) else None

            partno, unit = cell(part_col), cell(unit_col)
            rows.append({"row": row_no, "partno": partno, "part_key": key(partno),
                         "unit": unit, "unit_key": key(unit), "stat_quantity": cell(quantity_col or 8),
                         "task_card": cell(task_card_col or 3), "reference": cell(reference_col)})
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


def build_report(stat_rows, uom_rows, output: Path, stat_col: int, unit_col: int, inputs=(),
                 *, amos_index=None):
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

    decisions, matches, review, only_stat, stat_missing = [], [], [], [], []
    for stat in stat_rows:
        part, source_row, unit = stat["part_key"], stat["row"], stat["unit"]
        uoms = by_part.get(part, [])
        candidates = [(r, match_side(unit, r)) for r in uoms]
        matched = [(r, side) for r, side in candidates if side != "NONE"]
        status = decide(stat, uoms, matched)
        decisions.append([source_row, stat["partno"], unit, len(uoms), len(matched),
                          sum(not r["rule_issues"] for r, side in matched), status, False,
                          stat.get("stat_quantity"), stat.get("task_card"), stat.get("reference"),
                          *amos_evidence(amos_index, part)])
        # Every row here is a possible lookup alternative linked to one STAT row.
        for r, side in candidates:
            matches.append([part, source_row, unit, r["row"], r["external"], r["internal"],
                            r["rate"], r["exchange_base"], status, side, "; ".join(r["rule_issues"]) or "OK"])
        if not status.startswith("SINGLE_"):
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
    wb = Workbook()
    wb.remove(wb.active)
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
        ["Algfaile muudeti", "EI"], ["Ühikute teisendusi arvutati", "EI"], ["Hinnaarvutusi tehti", "EI"],
    ]
    append_sheet(wb, "Kokkuvote", ["Näitaja", "Tulemus"], summary)
    append_sheet(wb, "STAT decisions", ["stat_row", "partno", "stat_unit", "part_rule_rows",
        "unit_match_rows", "clean_unit_match_rows", "unit_decision", "conversion_approved",
        "stat_quantity", "task_card", "reference", "amos_lookup_status", "amos_source_row", "amos_unit",
        "avg_price", "avg_price_state", "purch_price", "purch_price_state"], decisions)
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
    stats = build_report(stat_rows, uom_rows, args.output, args.part_col, args.unit_col, inputs, amos_index=amos_index)
    print(f"Raport salvestatud: {args.output}")
    for label, value in stats.items():
        print(f"{label}: {value}")
    if len(stat_rows) < 2:
        print("HOIATUS: STAT failis on väga vähe andmeridu; testi täisandmestikuga.")


if __name__ == "__main__":
    main()
