"""Integration regressions. All generated workbooks stay in a temporary project."""

import hashlib
import shutil
import subprocess
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from scripts.report_io import literal_cell, save_report, validate_output
from scripts.exchange_rates import parse_ecb_rate
from scripts.sorter import build_service_summary
from scripts.uom_rules import annotate_rules


PROJECT = Path(__file__).resolve().parents[1]
PREP = "amos_prepared.xlsx"
SORT = "STAT_UOM_vordlusraport.xlsx"
FULL = "STAT_UOM_vordlusraport_friend.xlsx"


def workbook(path, title, rows):
    book = Workbook()
    sheet = book.active
    sheet.title = title
    for row in rows:
        sheet.append(row)
    book.save(path)
    book.close()


def records(path, title):
    book = load_workbook(path, read_only=True)
    try:
        values = book[title].values
        headers = next(values)
        return [dict(zip(headers, row)) for row in values]
    finally:
        book.close()


class ReportIntegration(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        shutil.copytree(PROJECT / "scripts", self.root / "scripts",
                        ignore=shutil.ignore_patterns("__pycache__"))
        shutil.copy(PROJECT / "amos_prepare.py", self.root)
        inputs = self.root / "excel_files/input"
        inputs.mkdir(parents=True)
        self.output = self.root / "excel_files/output"
        self.amos = inputs / "AMOS_uus_süsteem_info.xlsx"
        self.uom = inputs / "AMOS_UOM.xlsx"
        self.stat = inputs / "STAT_INITIAL - Copy.xlsx"
        workbook(self.amos, "Sheet1", [
            ("partno", "description", "measure_unit", "avg_price", "sales", "purch_price"),
            ("-999-22G1", " cable ", "ME", "1,25", 0, 0),
            ("PAINT", "Paint", "L", -1, 0, None),
        ])
        workbook(self.uom, "Sheet1", [
            ("partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base", "description"),
            ("-999-22G1", "CM", "ME", 0, 1, "Unusable rule"),
            ("PAINT", "EA", "L", 5, 1, "Container"),
            ("PAINT", "ML", "L", 1, 1000, "Millilitres"),
        ])
        workbook(self.stat, "INITIAL", [
            ("Reference", None, "TaskCard", "NRCNum", "Factor", "PN", "Description",
             "Qty", "Unit"),
            (None, None, "TASK-A", None, None, "-999-22G1", None, 1, "ME", "=1+2"),
            (None, None, "TASK-B", None, None, " paint ", None, 2, "L"),
            (None, None, "TASK-C", None, None, "PAINT", None, 3, "LT"),
        ])

    def cli(self, *args, success=True):
        run = subprocess.run([sys.executable, "-B", *map(str, args)], cwd=self.root,
                             capture_output=True, text=True)
        if success:
            self.assertEqual(run.returncode, 0, run.stderr + run.stdout)
        else:
            self.assertNotEqual(run.returncode, 0)

    def test_reports_keep_source_keys_and_stat_granularity(self):
        before = {p: hashlib.sha256(p.read_bytes()).digest() for p in (self.amos, self.uom, self.stat)}
        self.cli("amos_prepare.py")
        self.cli("scripts/sorter.py", "--amos", self.amos)
        self.cli("scripts/compare_stat_uom_full.py")
        prepared = records(self.output / PREP, "Prepared")
        self.assertEqual([r["partno"] for r in prepared], ["-999-22G1", "PAINT"])
        self.assertEqual(prepared[0]["avg_price"], 1.25)
        self.assertEqual(prepared[0]["uom_conversion_status"], "REVIEW_INVALID_RULE")
        self.assertEqual(prepared[1]["avg_price_state"], "NEGATIVE")
        self.assertEqual(prepared[1]["purch_price_state"], "MISSING")
        self.assertEqual(prepared[1]["uom_conversion_status"], "REVIEW_MULTIPLE_RULES")
        self.assertEqual(records(self.output / PREP, "UOM_Map")[0]["description"], "Unusable rule")

        decisions = records(self.output / SORT, "STAT decisions")
        self.assertEqual([r["stat_row"] for r in decisions], [2, 3, 4])
        self.assertEqual([r["unit_decision"] for r in decisions],
                         ["SAME_AMOS_UNIT", "AMBIGUOUS_UNIT_MATCH", "UNIT_NOT_LISTED"])
        self.assertEqual([r["conversion_approved"] for r in decisions], [True, False, False])
        self.assertEqual([r["stat_quantity"] for r in decisions], [1, 2, 3])
        self.assertEqual([r["task_card"] for r in decisions], ["TASK-A", "TASK-B", "TASK-C"])
        self.assertTrue(all(r["amos_lookup_status"] == "UNIQUE" for r in decisions))
        self.assertEqual(decisions[0]["avg_price"], 1.25)
        self.assertTrue(all(r["reference"] is None for r in decisions))
        for filename in (SORT, FULL):
            self.assertEqual(len(records(self.output / filename, "Vastavused")), 5)
        for path, digest in before.items():
            self.assertEqual(hashlib.sha256(path.read_bytes()).digest(), digest)

    def test_bad_schema_signals_failure_and_preserves_previous_report(self):
        self.cli("scripts/compare_stat_uom_full.py")
        old = (self.output / FULL).read_bytes()
        workbook(self.uom, "Sheet1", [
            ("partno", "partno", "measure_unit_ext", "measure_unit", "rate", "exchange_base"),
            ("A", "B", "EA", "EA", 1, 1),
        ])
        self.cli("scripts/compare_stat_uom_full.py", success=False)
        self.assertEqual((self.output / FULL).read_bytes(), old)

    def test_output_ownership_and_formula_protection(self):
        self.cli("amos_prepare.py", "--output", self.stat, success=False)
        self.cli("amos_prepare.py", "--output", self.output / SORT, success=False)
        self.output.mkdir(exist_ok=True)
        path = self.output / "manual.xlsx"
        workbook(path, "Manual", [(1, 2)])
        before = path.read_bytes()
        self.cli("amos_prepare.py", "--output", path, success=False)
        self.assertEqual(path.read_bytes(), before)
        self.cli("amos_prepare.py")
        path = self.output / PREP
        book = load_workbook(path)
        book["Summary"]["D5"] = "=1+1"
        book.save(path)
        book.close()
        before = path.read_bytes()
        self.cli("amos_prepare.py", success=False)
        self.assertEqual(path.read_bytes(), before)

    def test_failed_save_cleans_only_its_own_temporary_file(self):
        self.output.mkdir()
        existing = self.output / PREP
        sentinel = self.output / "amos_prepared.incomplete.xlsx"
        sentinel.write_bytes(b"pre-existing file")
        book = Workbook()
        save_report(book, existing, self.root, (), owner=PREP)
        before = existing.read_bytes()
        with patch.object(Workbook, "save", side_effect=OSError("disk full")):
            with self.assertRaises(OSError):
                save_report(Workbook(), existing, self.root, (), owner=PREP)
        self.assertEqual(existing.read_bytes(), before)
        self.assertEqual(sentinel.read_bytes(), b"pre-existing file")
        self.assertFalse(list(self.output.glob(".mro-*.tmp")))

    def test_output_symlink_to_input_is_refused(self):
        self.output.mkdir()
        link = self.output / "alias.xlsx"
        link.symlink_to(self.stat)
        with self.assertRaises(ValueError):
            validate_output(link, self.root, (), owner=PREP)

    def test_exact_repeat_conflict_and_legitimate_alternative_are_distinct(self):
        def rule(ext, rate):
            return dict(partno="A", measure_unit_ext=ext, measure_unit="EA",
                        rate=rate, exchange_base=1)
        rows = [rule("PACK", 1), rule("pack", "1,00"), rule("PACK", 10),
                rule("BAG", 5), rule("?", "NaN")]
        annotate_rules(rows)
        self.assertEqual(rows[0]["exact_rule_repetitions"], 2)
        self.assertEqual(set(rows[1]["rule_issues"]), {"REPEATED_UOM_RULE", "CONFLICTING_UOM_RULES"})
        self.assertEqual(rows[2]["rule_issues"], ("CONFLICTING_UOM_RULES",))
        self.assertEqual(rows[3]["rule_issues"], ())
        self.assertEqual(rows[4]["rule_issues"], ("INVALID_UOM_FACTOR",))

    def test_excel_literals_round_trip_without_prefixes_or_formulas(self):
        path = self.root / "strings.xlsx"
        book = Workbook(write_only=True)
        sheet = book.create_sheet("Data")
        texts = ["-999-22G1", "-", "=1+2", "+SKU", "@code", "#N/A", "000023"]
        sheet.append([literal_cell(sheet, s) for s in texts])
        book.save(path)
        book.close()
        book = load_workbook(path, read_only=True)
        try:
            row = next(book["Data"].rows)
            self.assertEqual([c.value for c in row], texts)
            self.assertTrue(all(c.data_type == "s" for c in row))
        finally:
            book.close()

    def test_zero_avg_price_uses_historical_service_average_plus_twenty_percent(self):
        book = load_workbook(self.amos)
        book["Sheet1"]["D3"] = 0
        book.save(self.amos)
        book.close()

        book = load_workbook(self.stat)
        avg = book.create_sheet("AVG")
        avg.append(["TaskCard", "avg"])
        avg.append(["TASK-B", 100])
        avg.append(["TASK-C", 200])
        book.save(self.stat)
        book.close()

        self.cli("scripts/sorter.py", "--amos", self.amos, "--unit-alias", "LT=L",
                 "--fallback-markup", "20")
        services = records(self.output / SORT, "Service Summary")
        by_task = {row["task_card"]: row for row in services}
        self.assertEqual(by_task["TASK-B"]["final_service_eur"], 120)
        self.assertEqual(by_task["TASK-C"]["final_service_eur"], 240)
        self.assertEqual(by_task["SUM_AVG"]["final_service_eur"], 361.25)
        self.assertEqual(by_task["TASK-B"]["final_price_source"], "OLD_PRICE_PLUS_MARKUP_FALLBACK")

    def test_service_uses_higher_of_complete_new_and_old_price_without_markup(self):
        details = [
            ({"task_card": "NEW-HIGH", "reference": "A"},
             {"line_total_eur": 120, "calculation_status": "OK", "avg_price_state": "POSITIVE"}),
            ({"task_card": "OLD-HIGH", "reference": "B"},
             {"line_total_eur": 80, "calculation_status": "OK", "avg_price_state": "POSITIVE"}),
        ]
        rows, total = build_service_summary(
            details, {"NEW-HIGH": Decimal("100"), "OLD-HIGH": Decimal("100")},
            Decimal("20"), Decimal("500"), exchange_rate=Decimal("1.25"),
            exchange_date="2026-11-20", exchange_status="LIVE")
        by_task = {row[0]: row for row in rows}
        self.assertEqual(by_task["NEW-HIGH"][7], 120)
        self.assertEqual(by_task["NEW-HIGH"][8], "NEW_PRICE_HIGHER_OR_EQUAL")
        self.assertEqual(by_task["OLD-HIGH"][7], 100)
        self.assertEqual(by_task["OLD-HIGH"][8], "OLD_PRICE_HIGHER")
        self.assertEqual(by_task["NEW-HIGH"][14], 150)
        self.assertEqual(total, 220)

    def test_approval_limit_flags_but_does_not_cap_price(self):
        details = [({"task_card": "TASK", "reference": None},
                    {"line_total_eur": 1000, "calculation_status": "OK",
                     "avg_price_state": "POSITIVE"})]
        rows, total = build_service_summary(details, {}, Decimal("20"), Decimal("500"))
        self.assertEqual(rows[0][7], 1000)
        self.assertEqual(rows[0][10], "REVIEW_OVER_LIMIT")
        self.assertEqual(total, 1000)

    def test_ecb_xml_parser_reads_eur_usd_rate_and_date(self):
        payload = b'''<gesmes:Envelope xmlns:gesmes="x" xmlns="y"><Cube><Cube time="2026-11-20">
            <Cube currency="USD" rate="1.2345"/></Cube></Cube></gesmes:Envelope>'''
        self.assertEqual(parse_ecb_rate(payload), (Decimal("1.2345"), "2026-11-20"))
