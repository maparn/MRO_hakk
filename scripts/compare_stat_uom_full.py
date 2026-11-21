"""Full STAT × UOM candidate-pair report for manual inspection.

Run from the repository root: python scripts/compare_stat_uom_full.py
Vastavused contains every same-part pair. The unit comparison on this
legacy sheet is against the internal UOM only. This report applies no
conversion and does not provide rows for summing quantities or prices.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path
from sys import stderr
from zipfile import BadZipFile

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

if __package__:
    from . import report_io, sorter
else:
    import report_io
    import sorter


KAUST = Path(__file__).resolve().parents[1]
RAPORT_FAIL = KAUST / "excel_files" / "output" / "STAT_UOM_vordlusraport_friend.xlsx"
MAX_DATA_ROWS = 1_048_575  # One Excel row is reserved for each sheet header.

PAIR_HEADERS = (
    "partno võti", "STAT rea nr", "STAT partno", "STAT measure_unit",
    "UOM rea nr", "UOM partno", "UOM measure_unit", "UOM measure_unit_ext",
    "rate", "exchange_base", "STAT vs sisemine ühik", "UOM reegli probleemid",
)
UOM_HEADERS = (
    "partno võti", "UOM rea nr", "UOM partno", "UOM unit", "UOM measure_unit_ext",
    "rate", "exchange_base", "UOM reegli probleemid",
)
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")


def make_sheet(book, name, headers, widths=()):
    sheet = book.create_sheet(name)
    sheet.freeze_panes = "A2"
    cells = []
    for index, header in enumerate(headers, 1):
        width = widths[index - 1] if widths else max(15, len(header) + 3)
        sheet.column_dimensions[get_column_letter(index)].width = width
        cell = WriteOnlyCell(sheet, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cells.append(cell)
    sheet.append(cells)
    return sheet


def index_by_part(records):
    grouped = defaultdict(list)
    missing = []
    for row in records:
        if row["part_key"]:
            grouped[row["part_key"]].append(row)
        else:
            missing.append(row)
    return grouped, missing


def internal_comparison(stat_unit_key, internal):
    uom_key = sorter.key(internal)
    if not stat_unit_key and not uom_key:
        return "MISSING_BOTH_UNITS"
    if not stat_unit_key:
        return "MISSING_STAT_UNIT"
    if not uom_key:
        return "MISSING_INTERNAL_UNIT"
    return "SAME_INTERNAL_UNIT" if stat_unit_key == uom_key else "DIFFERENT_INTERNAL_UNIT"


def uom_values(rule):
    return [
        rule["row"], rule["partno"], rule["internal"], rule["external"],
        rule["rate"], rule["exchange_base"], "; ".join(rule["rule_issues"]),
    ]


def generate(stat_path, uom_path, output, part_col, unit_col):
    inputs = (stat_path, uom_path)
    report_io.validate_output(output, KAUST, inputs, owner=RAPORT_FAIL.name)
    stat_rows = sorter.read_stat(stat_path, part_col, unit_col)
    uom_rows = sorter.read_uom(uom_path)
    stats, missing_stats = index_by_part(stat_rows)
    rules, missing_rules = index_by_part(uom_rows)
    common = stats.keys() & rules.keys()
    pair_count = sum(len(stats[key]) * len(rules[key]) for key in common)
    if max(len(stat_rows), len(uom_rows), pair_count) > MAX_DATA_ROWS:
        raise ValueError("Paaride aruanne ületaks Exceli töölehe reapiiri.")

    book = Workbook(write_only=True)
    summary = make_sheet(book, "Kokkuvote", ["Näitaja", "Tulemus"], widths=(48, 76))
    pairs = make_sheet(book, "Vastavused", PAIR_HEADERS)
    review = make_sheet(book, "Unit erinevused", PAIR_HEADERS)
    only_stat = make_sheet(book, "Ainult STAT", ["partno võti", "STAT rea nr", "STAT partno", "STAT unit"])
    only_uom = make_sheet(book, "Ainult UOM", UOM_HEADERS)
    repeated = make_sheet(book, "Duplikaadid", [*UOM_HEADERS, "Identse reegli kordusi"])
    conflicts = make_sheet(book, "UOM konfliktid", UOM_HEADERS)
    variants = make_sheet(book, "UOM variandid", UOM_HEADERS)
    missing_stat = make_sheet(book, "STAT puuduv partno", ["STAT rea nr", "STAT measure_unit"])
    missing_uom = make_sheet(book, "UOM puuduv partno", ["UOM rea nr", "UOM measure_unit", "UOM measure_unit_ext"])
    counts = Counter()

    def append(sheet, values):
        sheet.append([report_io.literal_cell(sheet, value) for value in values])
        counts[sheet.title] += 1

    for row in missing_stats:
        append(missing_stat, [row["row"], row["unit"]])
    for rule in missing_rules:
        append(missing_uom, [rule["row"], rule["internal"], rule["external"]])

    alternative_parts = 0
    for key in sorted(stats.keys() | rules.keys()):
        stat_group = stats.get(key, ())
        rule_group = rules.get(key, ())
        external_units = {sorter.key(rule["external"]) for rule in rule_group}
        has_alternatives = len(external_units - {""}) > 1
        alternative_parts += has_alternatives

        for rule in rule_group:
            values = [key, *uom_values(rule)]
            if not stat_group:
                append(only_uom, values)
            if has_alternatives:
                append(variants, values)
            if rule["exact_rule_repetitions"] > 1:
                append(repeated, [*values, rule["exact_rule_repetitions"]])
            if "CONFLICTING_UOM_RULES" in rule["rule_issues"]:
                append(conflicts, values)

        for row in stat_group:
            stat_values = [key, row["row"], row["partno"], row["unit"]]
            if not rule_group:
                append(only_stat, stat_values)
            for rule in rule_group:
                comparison = internal_comparison(row["unit_key"], rule["internal"])
                number, partno, internal, external, rate, base, issues = uom_values(rule)
                values = [*stat_values, number, partno, internal, external, rate, base, comparison, issues]
                append(pairs, values)
                if comparison != "SAME_INTERNAL_UNIT":
                    append(review, values)

    summary_rows = [
        ("STAT-faili sisuga read", len(stat_rows)),
        ("UOM-faili sisuga read", len(uom_rows)),
        ("Ühised unikaalsed partno-d", len(common)),
        ("STAT read, mille partno leidub UOM-is", sum(len(stats[key]) for key in common)),
        ("Kandidaatpaarid partno alusel", counts[pairs.title]),
        ("Paarid: sisemine ühik erineb või puudub", counts[review.title]),
        ("Ainult STAT-failis (read)", counts[only_stat.title]),
        ("Ainult UOM-failis (read)", counts[only_uom.title]),
        ("Mitme välise ühikuga partno-d kogu UOM-is", alternative_parts),
        ("Korduva identse UOM-reegli read", counts[repeated.title]),
        ("Konfliktse UOM-reegli read", counts[conflicts.title]),
        ("STAT read puuduva partno väärtusega", counts[missing_stat.title]),
        ("UOM read puuduva partno väärtusega", counts[missing_uom.title]),
        ("Võrdlusvõti", f"STAT veerg {get_column_letter(part_col)} ↔ UOM partno"),
        ("Mõõtühiku võrdlus", f"STAT veerg {get_column_letter(unit_col)} ↔ UOM measure_unit (sisemine)"),
        ("Vastavused", "Kõik sama partno-ga STAT × UOM paarid. Ridu ei summeerita."),
        ("Ühikute teisendus", "Ei rakendata. Sama sisemine ühik ei kinnita teisendusreeglit."),
    ]
    for metric in summary_rows:
        append(summary, metric)
    for sheet in book:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(sheet.column_dimensions))}{counts[sheet.title] + 1}"
    report_io.save_report(book, output, KAUST, inputs, owner=RAPORT_FAIL.name)
    return dict(summary_rows)


def main():
    parser = argparse.ArgumentParser(description="STAT × UOM kõigi kandidaatpaaride diagnostika")
    parser.add_argument("--stat", type=Path)
    parser.add_argument("--uom", type=Path)
    parser.add_argument("--output", type=Path, default=RAPORT_FAIL)
    parser.add_argument("--part-col", type=int, default=sorter.STAT_PARTNO_COL)
    parser.add_argument("--unit-col", type=int, default=sorter.STAT_UNIT_COL)
    args = parser.parse_args()
    try:
        stat = args.stat or sorter.locate_input(sorter.STAT_FILENAME)
        uom = args.uom or sorter.locate_input(sorter.UOM_FILENAME)
        metrics = generate(stat, uom, args.output, args.part_col, args.unit_col)
    except (OSError, ValueError, BadZipFile) as error:
        print(f"Raporti loomine ebaõnnestus: {error}", file=stderr)
        return 1
    print(f"Raport salvestatud: {args.output}")
    for name, value in metrics.items():
        if isinstance(value, int):
            print(f"{name}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
