#!/usr/bin/env python3
"""Magnetic MRO: STAT vs AMOS UOM võrdlusraport (ainult lugemine).

Kasutamine Ubuntu/WSL-is, MRO_hakk kaustast:
    python scripts/sorter.py
    python scripts/sorter.py --stat "excel_files/input/STAT_INITIAL - Copy.xlsx"

Failid võivad paikneda projekti juurkaustas või excel_files/input kaustas.
Väljund: excel_files/output/STAT_UOM_vordlusraport.xlsx

NB! See skript on KONTROLL, mitte ühikute teisendaja ega hinnakalkulaator.
Sama partno mitu UOM-rida võib olla normaalne (eri välised mõõtühikud).
Ära käsitle ühikute tekstilisi erinevusi ilma ettevõtte reegliteta veana.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PROJECT = Path(__file__).resolve().parents[1]
STAT_FILENAME = "STAT_INITIAL - Copy.xlsx"
UOM_FILENAME = "AMOS_UOM.xlsx"
OUTPUT_DEFAULT = PROJECT / "excel_files" / "output" / "STAT_UOM_vordlusraport.xlsx"
STAT_SHEET = "INITIAL"
UOM_SHEET = "Sheet1"

# STAT_INITIAL näite kontrollitud paigutus: F = partno, I = mõõtühik.
STAT_PARTNO_COL = 6
STAT_UNIT_COL = 9


def key(value):
    return "" if value is None else str(value).strip().upper()


def locate_input(filename: str) -> Path:
    """Otsi faili kasutaja projektist; ära muuda algandmeid."""
    choices = [PROJECT / "excel_files" / "input" / filename, PROJECT / filename]
    for path in choices:
        if path.is_file():
            return path
    raise FileNotFoundError(
        f"Ei leia faili '{filename}'. Aseta see kausta excel_files/input või projekti juurkausta."
    )


def read_stat(path: Path, part_col: int, unit_col: int):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if STAT_SHEET not in wb.sheetnames:
            raise ValueError(f"STAT leht '{STAT_SHEET}' puudub; olemas: {wb.sheetnames}")
        rows = []
        for row_no, cells in enumerate(wb[STAT_SHEET].iter_rows(values_only=True), start=1):
            if not any(c is not None and str(c).strip() for c in cells):
                continue
            partno = cells[part_col - 1] if len(cells) >= part_col else None
            unit = cells[unit_col - 1] if len(cells) >= unit_col else None
            rows.append({"row": row_no, "partno": partno, "part_key": key(partno),
                         "unit": unit, "unit_key": key(unit)})
        return rows
    finally:
        wb.close()


def read_uom(path: Path):
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        if UOM_SHEET not in wb.sheetnames:
            raise ValueError(f"UOM leht '{UOM_SHEET}' puudub; olemas: {wb.sheetnames}")
        sheet = wb[UOM_SHEET]
        # AMOS ekspordi mõõtmed võivad olla valed (nt A1:B4, kuigi 12 veergu ja 1418 rida).
        # reset_dimensions() sunnib openpyxl-i tegelikud read lõpuni läbi lugema.
        sheet.reset_dimensions()
        header = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if not header:
            raise ValueError("UOM fail on tühi")
        positions = {key(value).lower(): i for i, value in enumerate(header)}
        required = ("partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base")
        missing = [h for h in required if h not in positions]
        if missing:
            raise ValueError(f"UOM failis puuduvad veerud: {', '.join(missing)}")
        out = []
        for row_no, cells in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(c is not None and str(c).strip() for c in cells):
                continue
            def cell(name):
                i = positions[name]
                return cells[i] if i < len(cells) else None
            partno = cell("partno")
            out.append({"row": row_no, "partno": partno, "part_key": key(partno),
                        "external": cell("measure_unit_ext"),
                        "internal": cell("measure_unit"),
                        "rate": cell("rate"), "exchange_base": cell("exchange_base")})
        return out
    finally:
        wb.close()


def append_sheet(wb, name, header, rows):
    ws = wb.create_sheet(name)
    ws.append(header)
    for row in rows:
        # Kaitse Exceli valemina tõlgendatava kasutajateksti eest raportis.
        clean = []
        for value in row:
            if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
                value = "'" + value
            clean.append(value)
        ws.append(clean)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32
    # Mõistlikud veerulaiused (ära skaneeri kõiki kümneid tuhandeid ridu).
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        width = max(len(str(c.value or "")) for c in col[:100]) if col else 10
        ws.column_dimensions[letter].width = min(48, max(15, width + 2))
    return ws


def build_report(stat_rows, uom_rows, output: Path, stat_col: int, unit_col: int):
    stat_by_part, uom_by_part = defaultdict(list), defaultdict(list)
    stat_missing, uom_missing = [], []
    for row in stat_rows:
        if row["part_key"]:
            stat_by_part[row["part_key"]].append(row)
        else:
            stat_missing.append([row["row"], row["unit"]])
    for row in uom_rows:
        if row["part_key"]:
            uom_by_part[row["part_key"]].append(row)
        else:
            uom_missing.append([row["row"], row["external"], row["internal"]])

    matches, unit_review, only_stat, only_uom, variants, duplicates = [], [], [], [], [], []
    for partno in sorted(set(stat_by_part) | set(uom_by_part)):
        stats, uoms = stat_by_part.get(partno, []), uom_by_part.get(partno, [])
        if not stats:
            for uom in uoms:
                only_uom.append([partno, uom["row"], uom["external"], uom["internal"]])
            continue
        if not uoms:
            for stat in stats:
                only_stat.append([partno, stat["row"], stat["unit"]])
            continue

        if len(uoms) > 1:
            variants.append([partno, len(uoms), ", ".join(
                f"{u['external']}→{u['internal']} (rida {u['row']})" for u in uoms)])

        # Kui mitu teisendusvarianti eksisteerib, on need alternatiivid,
        # mitte üks ühine teisendustegur ega automaatselt dubleeritud read.
        for stat in stats:
            exact = [u for u in uoms if stat["unit_key"] and stat["unit_key"] in
                     (key(u["external"]), key(u["internal"]))]
            status = "Ühik vastab vähemalt ühele UOM variandile" if exact else "Ühik vajab käsitsi kontrolli"
            choices = exact if exact else uoms
            for u in choices:
                matches.append([partno, stat["row"], stat["unit"], u["row"],
                                u["external"], u["internal"], u["rate"],
                                u["exchange_base"], status])
            if not exact:
                unit_review.append([partno, stat["row"], stat["unit"],
                                    "; ".join(f"{u['external']}→{u['internal']}" for u in uoms),
                                    "Ei ole kinnitatud teisendusreegel; kontrolli käsitsi"])

    # Tõeline kordus: sama partno + sama välise ja sisemise ühiku kombinatsioon.
    groups = defaultdict(list)
    for uom in uom_rows:
        if uom["part_key"]:
            groups[(uom["part_key"], key(uom["external"]), key(uom["internal"]))].append(uom)
    for (partno, external, internal), group in sorted(groups.items()):
        if len(group) > 1:
            duplicates.append([partno, external, internal, len(group),
                               ", ".join(str(r["row"]) for r in group)])

    wb = Workbook()
    wb.remove(wb.active)
    summary = [
        ["Näitaja", "Tulemus"],
        ["STAT andmeread", len(stat_rows)],
        ["UOM andmeread", len(uom_rows)],
        ["STAT unikaalsed partno-d", len(stat_by_part)],
        ["UOM unikaalsed partno-d", len(uom_by_part)],
        ["Sobivad partno-d", len(set(stat_by_part) & set(uom_by_part))],
        ["Ühikuid vajab kontrolli (STAT read)", len(unit_review)],
        ["Ainult STAT (read)", len(only_stat)],
        ["Ainult UOM (read)", len(only_uom)],
        ["Mitu UOM varianti (partno-d)", len(variants)],
        ["Korduvad UOM ühikupaarid (võimalikud vead)", len(duplicates)],
        ["STAT partno puudub", len(stat_missing)],
        ["UOM partno puudub", len(uom_missing)],
        ["STAT partno veerg", get_column_letter(stat_col)],
        ["STAT mõõtühiku veerg", get_column_letter(unit_col)],
        ["Algfaile muudeti", "EI"],
        ["Ühikute teisendusi arvutati", "EI"],
        ["Hinnaarvutusi tehti", "EI"],
    ]
    append_sheet(wb, "Kokkuvote", summary[0], summary[1:])
    append_sheet(wb, "Vastavused", ["partno", "STAT rida", "STAT ühik", "UOM rida", "UOM ext ühik",
                                   "UOM sisemine ühik", "rate", "exchange_base", "Olek"], matches)
    append_sheet(wb, "Unit ulevaatus", ["partno", "STAT rida", "STAT ühik",
                                      "UOM variandid", "Märkus"], unit_review)
    append_sheet(wb, "Ainult STAT", ["partno", "STAT rida", "STAT ühik"], only_stat)
    append_sheet(wb, "Ainult UOM", ["partno", "UOM rida", "UOM ext", "UOM sisemine"], only_uom)
    append_sheet(wb, "UOM variandid", ["partno", "Variantide arv", "Ühikud ja read"], variants)
    append_sheet(wb, "Duplikaadid", ["partno", "Väline ühik", "Sisemine ühik", "Kordusi", "Read"], duplicates)
    append_sheet(wb, "STAT puuduv partno", ["STAT rida", "STAT ühik"], stat_missing)
    append_sheet(wb, "UOM puuduv partno", ["UOM rida", "UOM ext", "UOM sisemine"], uom_missing)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    wb.close()
    return {"stat_rows": len(stat_rows), "uom_rows": len(uom_rows),
            "matched_partnos": len(set(stat_by_part) & set(uom_by_part)),
            "review_rows": len(unit_review), "uom_variants": len(variants)}


def main():
    parser = argparse.ArgumentParser(description="Kontrolli STAT ja UOM mõõtühikute vastavusi")
    parser.add_argument("--stat", type=Path, help="STAT Excel faili tee")
    parser.add_argument("--uom", type=Path, help="AMOS UOM Excel faili tee")
    parser.add_argument("--output", type=Path, default=OUTPUT_DEFAULT, help="Raporti Excel faili tee")
    parser.add_argument("--part-col", type=int, default=STAT_PARTNO_COL,
                        help="STAT partno veeru number (vaikimisi F=6)")
    parser.add_argument("--unit-col", type=int, default=STAT_UNIT_COL,
                        help="STAT ühiku veeru number (vaikimisi I=9)")
    args = parser.parse_args()
    if args.part_col < 1 or args.unit_col < 1:
        parser.error("Veeru number peab olema positiivne")

    stat_path = args.stat or locate_input(STAT_FILENAME)
    uom_path = args.uom or locate_input(UOM_FILENAME)
    if not stat_path.is_file() or not uom_path.is_file():
        parser.error("Sisendfaili ei leitud. Kontrolli --stat ja --uom teid.")
    if args.output.resolve() in {stat_path.resolve(), uom_path.resolve()}:
        parser.error("Väljund ei tohi olla sama fail mis sisend!")

    print(f"STAT: {stat_path}")
    print(f"UOM:  {uom_path}")
    print("Loen lähtefailid...")
    stat_rows = read_stat(stat_path, args.part_col, args.unit_col)
    uom_rows = read_uom(uom_path)
    stats = build_report(stat_rows, uom_rows, args.output, args.part_col, args.unit_col)
    print(f"Raport salvestatud: {args.output}")
    for k, v in stats.items():
        print(f"{k}: {v}")
    if len(stat_rows) < 2:
        print("HOIATUS: STAT failis on väga vähe andmeridu; testi täisandmestikuga.")
    print("Algfaile ei muudetud. Ühikute teisendusi ja hinnaarvutusi EI tehtud.")


if __name__ == "__main__":
    main()
