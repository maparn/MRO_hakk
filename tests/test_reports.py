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
from scripts.sorter import build_aggregations, calculate_line, parse_physical_price_conversions
from scripts.web_processing import load_converter_html, process_workbooks
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
             "Qty", "Unit", "USD", "EUR"),
            ("REF-A", None, "TASK-A", None, None, "-999-22G1", None, 1, "ME", 0, 10),
            ("REF-B", None, "TASK-B", None, None, " paint ", None, 2, "L", 0, 5),
            ("REF-C", None, "TASK-C", None, None, "PAINT", None, 3, "LT", 0, 6),
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
                         ["SAME_UNIT", "AMBIGUOUS_UNIT_MATCH", "UNIT_NOT_LISTED"])
        self.assertEqual([r["conversion_approved"] for r in decisions], [True, False, False])
        self.assertEqual([r["stat_quantity"] for r in decisions], [1, 2, 3])
        self.assertEqual([r["task_card"] for r in decisions], ["TASK-A", "TASK-B", "TASK-C"])
        self.assertTrue(all(r["amos_lookup_status"] == "UNIQUE" for r in decisions))
        self.assertEqual(decisions[0]["avg_price"], 1.25)
        self.assertEqual([r["reference"] for r in decisions], ["REF-A", "REF-B", "REF-C"])
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

    @staticmethod
    def stat_line(part="PART", unit="EA", quantity=2, historical=10,
                  reference="REF-1", task_card="TASK-1"):
        return {
            "row": 2, "partno": part, "part_key": part.upper(), "unit": unit,
            "unit_key": unit.upper(), "stat_quantity": quantity,
            "historical_unit_price_eur": historical, "reference": reference,
            "task_card": task_card,
        }

    def test_identical_units_use_current_amos_price(self):
        stat = self.stat_line(quantity=3)
        result = calculate_line(stat, [], {"PART": [{"measure_unit": "EA", "avg_price": 7}]}, {})
        self.assertEqual(result["line_total_eur"], 21)
        self.assertEqual(result["price_source"], "CURRENT_AMOS_PRICE")
        self.assertEqual(result["conversion_status"], "SAME_UNIT")

    def test_zero_current_price_uses_row_historical_price_plus_markup(self):
        stat = self.stat_line(quantity=3, historical=10)
        result = calculate_line(stat, [], {"PART": [{"measure_unit": "EA", "avg_price": 0}]}, {},
                                Decimal("20"))
        self.assertEqual(result["selected_unit_price_eur"], 12)
        self.assertEqual(result["line_total_eur"], 36)
        self.assertEqual(result["price_source"], "HISTORICAL_PRICE_PLUS_20_PERCENT")
        self.assertEqual(result["fallback_reason"], "CURRENT_PRICE_ZERO")
        self.assertEqual(result["calculation_status"], "FALLBACK_APPLIED")
        self.assertEqual(result["current_price_status"], "CURRENT_PRICE_ZERO")

    def test_historical_markup_is_applied_exactly_once(self):
        result = calculate_line(
            self.stat_line(quantity=3, historical=50), [],
            {"PART": [{"measure_unit": "EA", "avg_price": 0}]}, {}, Decimal("20"),
        )
        self.assertEqual(result["selected_unit_price_eur"], 60)
        self.assertEqual(result["line_total_eur"], 180)

    def test_blank_current_price_uses_row_historical_price_plus_markup(self):
        stat = self.stat_line(quantity=2, historical=15)
        result = calculate_line(stat, [], {"PART": [{"measure_unit": "EA", "avg_price": None}]}, {},
                                Decimal("20"))
        self.assertEqual(result["line_total_eur"], 36)
        self.assertEqual(result["fallback_reason"], "MISSING_CURRENT_PRICE")

    def test_missing_historical_price_does_not_block_valid_current_price(self):
        result = calculate_line(
            self.stat_line(quantity=3, historical=None), [],
            {"PART": [{"measure_unit": "EA", "avg_price": 7.5}]}, {})
        self.assertEqual(result["selected_unit_price_eur"], 7.5)
        self.assertEqual(result["line_total_eur"], 22.5)
        self.assertEqual(result["price_source"], "CURRENT_AMOS_PRICE")

    def test_both_current_and_historical_prices_missing_require_review(self):
        result = calculate_line(
            self.stat_line(quantity=3, historical=None), [],
            {"PART": [{"measure_unit": "EA", "avg_price": None}]}, {})
        self.assertIsNone(result["selected_unit_price_eur"])
        self.assertEqual(result["current_price_status"], "CURRENT_PRICE_MISSING")
        self.assertEqual(result["fallback_status"], "MISSING_HISTORICAL_PRICE")
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_MISSING_HISTORICAL_PRICE")

    def test_missing_uom_rule_uses_real_world_row_level_fallback(self):
        stat = self.stat_line(part="17-4PH/CRES/AMS5643BAR/0.5", unit="MM",
                              quantity=70, historical=52.86)
        amos = {stat["part_key"]: [{"measure_unit": "ME", "avg_price": 254.44}]}
        result = calculate_line(stat, [], amos, {}, Decimal("20"))
        self.assertAlmostEqual(result["selected_unit_price_eur"], 63.432)
        self.assertAlmostEqual(result["line_total_eur"], 4440.24)
        self.assertEqual(result["fallback_reason"], "MISSING_UOM_RULE")
        self.assertEqual(result["conversion_status"], "NOT_APPLICABLE_FALLBACK")
        self.assertEqual(result["fallback_status"], "HISTORICAL_FALLBACK_APPLIED")

    def test_missing_uom_and_historical_price_requires_manual_review(self):
        stat = self.stat_line(unit="MM", historical=None)
        result = calculate_line(stat, [], {"PART": [{"measure_unit": "ME", "avg_price": 10}]}, {})
        self.assertIsNone(result["line_total_eur"])
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_MISSING_HISTORICAL_PRICE")
        self.assertEqual(result["fallback_status"], "MISSING_HISTORICAL_PRICE")

    def test_absent_amos_record_is_not_silently_approved_from_history(self):
        result = calculate_line(self.stat_line(historical=25), [], {}, {})
        self.assertIsNone(result["line_total_eur"])
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_AMOS_RECORD_MISSING")
        self.assertTrue(result["review_required"])
        self.assertEqual(result["current_price_status"], "MATERIAL_NOT_FOUND")

    def test_stat_footer_is_not_counted_as_a_material(self):
        workbook(self.stat, "INITIAL", [
            ("Reference", None, "TaskCard", None, None, "PN", "Description", "Qty", "Unit"),
            ("REF", None, "TASK", None, None, "PART", "Material", 2, "EA"),
            (None, None, None, None, None, None, None, None, None,
             None, None, None, None, None, None, None, "GRAND TOTAL:", "=SUM(R2:R2)"),
        ])
        from scripts.sorter import read_stat
        rows = read_stat(self.stat, 6, 9, quantity_col=8, task_card_col=3, reference_col=1)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["part_key"], "PART")

    def test_similar_part_numbers_never_match(self):
        stat = self.stat_line(part="9MK619321-1")
        amos = {"9MK619321-10": [{"measure_unit": "EA", "avg_price": 100}]}
        result = calculate_line(stat, [], amos, {})
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_AMOS_RECORD_MISSING")

    def test_multiple_exact_amos_rows_require_manual_review(self):
        amos = {"PART": [
            {"measure_unit": "EA", "avg_price": 10},
            {"measure_unit": "EA", "avg_price": 11},
        ]}
        result = calculate_line(self.stat_line(), [], amos, {})
        self.assertIsNone(result["line_total_eur"])
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_MULTIPLE_AMOS_RECORDS")

    def test_conflicting_uom_rules_require_manual_review_without_duplication(self):
        stat = self.stat_line(unit="PACK", quantity=2, historical=10)
        rules = [
            {"external": "PACK", "internal": "EA", "rate": 5,
             "exchange_base": 1, "rule_issues": ("CONFLICTING_UOM_RULES",)},
            {"external": "PACK", "internal": "EA", "rate": 10,
             "exchange_base": 1, "rule_issues": ("CONFLICTING_UOM_RULES",)},
        ]
        result = calculate_line(stat, rules, {"PART": [{"measure_unit": "EA", "avg_price": 3}]}, {})
        self.assertIsNone(result["line_total_eur"])
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_AMBIGUOUS_UOM_RULE")

    def test_invalid_uom_factor_requires_manual_review(self):
        stat = self.stat_line(unit="PACK", quantity=2, historical=10)
        rules = [{"external": "PACK", "internal": "EA", "rate": 0,
                  "exchange_base": 1, "rule_issues": ("INVALID_UOM_FACTOR",)}]
        result = calculate_line(stat, rules, {"PART": [{"measure_unit": "EA", "avg_price": 3}]}, {})
        self.assertIsNone(result["line_total_eur"])
        self.assertEqual(result["calculation_status"], "MANUAL_REVIEW_INVALID_UOM_FACTOR")

    def test_repeated_part_rows_keep_their_own_historical_prices(self):
        amos = {"PART": [{"measure_unit": "EA", "avg_price": 0}]}
        first = calculate_line(self.stat_line(quantity=2, historical=10), [], amos, {})
        second = calculate_line(self.stat_line(quantity=2, historical=30), [], amos, {})
        self.assertEqual(first["line_total_eur"], 24)
        self.assertEqual(second["line_total_eur"], 72)

    def test_reference_sum_then_task_card_average(self):
        details = [
            (self.stat_line(reference="REF-A", task_card="TASK", quantity=1),
             {"line_total_eur": 100, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-A", task_card="TASK", quantity=1),
             {"line_total_eur": 50, "price_source": "HISTORICAL_PRICE_PLUS_20_PERCENT"}),
            (self.stat_line(reference="REF-B", task_card="TASK", quantity=1),
             {"line_total_eur": 250, "price_source": "CURRENT_AMOS_PRICE"}),
        ]
        refs, tasks, sum_avg, status = build_aggregations(details, {}, Decimal("1.2"))
        self.assertEqual([row[7] for row in refs], [150, 250])
        self.assertEqual(tasks[0][5], 200)
        self.assertEqual(tasks[0][7], 200)
        self.assertEqual(tasks[0][8], "NEW")
        self.assertEqual(tasks[0][10], 240)
        self.assertEqual(sum_avg, 200)
        self.assertEqual(status, "COMPLETE")

    def test_higher_old_or_new_task_card_avg_is_selected(self):
        details = [
            (self.stat_line(reference="REF-A", task_card="TASK-NEW"),
             {"line_total_eur": 650, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-B", task_card="TASK-OLD"),
             {"line_total_eur": 600, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-C", task_card="TASK-EQUAL"),
             {"line_total_eur": 300, "price_source": "CURRENT_AMOS_PRICE"}),
        ]
        old = {"TASK-NEW": Decimal("500"), "TASK-OLD": Decimal("800"),
               "TASK-EQUAL": Decimal("300")}
        _, tasks, final_total, status = build_aggregations(details, old, Decimal("1.1"))
        by_task = {row[0]: row for row in tasks[:-1]}
        self.assertEqual((by_task["TASK-NEW"][7], by_task["TASK-NEW"][8]), (650, "NEW"))
        self.assertEqual((by_task["TASK-OLD"][7], by_task["TASK-OLD"][8]), (800, "OLD"))
        self.assertEqual((by_task["TASK-EQUAL"][7], by_task["TASK-EQUAL"][8]), (300, "NEW"))
        self.assertEqual(by_task["TASK-OLD"][10], 880)
        self.assertEqual(final_total, 1750)
        self.assertEqual(status, "COMPLETE")

    def test_distinct_task_cards_never_mix_reference_totals(self):
        details = [
            (self.stat_line(reference="REF-A", task_card="TASK-1"),
             {"line_total_eur": 100, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-B", task_card="TASK-1"),
             {"line_total_eur": 300, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-A", task_card="TASK-2"),
             {"line_total_eur": 900, "price_source": "CURRENT_AMOS_PRICE"}),
        ]
        _, tasks, _, _ = build_aggregations(details, {}, None)
        by_task = {row[0]: row for row in tasks[:-1]}
        self.assertEqual(by_task["TASK-1"][5], 200)
        self.assertEqual(by_task["TASK-2"][5], 900)

    def test_unresolved_reference_is_not_presented_as_valid_zero(self):
        details = [
            (self.stat_line(reference="REF-A", task_card="TASK"),
             {"line_total_eur": None, "price_source": "MANUAL_REVIEW_REQUIRED"}),
        ]
        refs, tasks, sum_avg, status = build_aggregations(details, {}, Decimal("1.2"))
        self.assertIsNone(refs[0][7])
        self.assertIsNone(tasks[0][5])
        self.assertEqual(sum_avg, 0)
        self.assertEqual(status, "INCOMPLETE")

    def test_incomplete_new_avg_still_selects_max_but_remains_flagged(self):
        details = [
            (self.stat_line(reference="REF-A", task_card="TASK"),
             {"line_total_eur": 900, "price_source": "CURRENT_AMOS_PRICE"}),
            (self.stat_line(reference="REF-B", task_card="TASK"),
             {"line_total_eur": None, "price_source": "MANUAL_REVIEW_REQUIRED"}),
        ]
        _, tasks, final_total, status = build_aggregations(
            details, {"TASK": Decimal("500")}, Decimal("1.2"))
        self.assertEqual(tasks[0][5], 900)
        self.assertEqual((tasks[0][7], tasks[0][8]), (900, "NEW"))
        self.assertEqual(tasks[0][11], "INCOMPLETE")
        self.assertEqual(final_total, 900)
        self.assertEqual(status, "INCOMPLETE")

    def test_missing_exchange_rate_keeps_eur_and_leaves_usd_blank(self):
        details = [(self.stat_line(), {"line_total_eur": 100, "price_source": "CURRENT_AMOS_PRICE"})]
        _, tasks, sum_avg, status = build_aggregations(details, {}, None)
        self.assertEqual(sum_avg, 100)
        self.assertIsNone(tasks[0][10])
        self.assertEqual(status, "COMPLETE")

    def test_missing_old_avg_uses_new_and_missing_new_uses_old(self):
        calculated = [(self.stat_line(task_card="NEW"),
                       {"line_total_eur": 125, "price_source": "CURRENT_AMOS_PRICE"})]
        _, tasks, total, _ = build_aggregations(calculated, {}, Decimal("1.2"))
        self.assertEqual((tasks[0][7], tasks[0][8], tasks[0][10]), (125, "NEW", 150))
        unresolved = [(self.stat_line(task_card="OLD"),
                       {"line_total_eur": None, "price_source": "MANUAL_REVIEW_REQUIRED"})]
        _, tasks, total, status = build_aggregations(
            unresolved, {"OLD": Decimal("400")}, Decimal("1.2"))
        self.assertEqual((tasks[0][7], tasks[0][8], tasks[0][10]), (400, "OLD", 480))
        self.assertEqual(total, 400)
        self.assertEqual(status, "INCOMPLETE")

    def test_uom_rate_and_exchange_base_are_applied_in_the_documented_direction(self):
        stat = {
            "part_key": "SEAL", "unit_key": "FT", "stat_quantity": 10,
        }
        rules = [{
            "external": "FT", "internal": "CM", "rate": 305,
            "exchange_base": 10, "rule_issues": (),
        }]
        amos = {"SEAL": [{"measure_unit": "CM", "avg_price": 2}]}
        result = calculate_line(stat, rules, amos, {})
        self.assertEqual(result["converted_quantity"], 305)
        self.assertEqual(result["conversion_rate"], 305)
        self.assertEqual(result["conversion_base"], 10)
        self.assertEqual(result["line_total_eur"], 610)
        self.assertEqual(result["selected_unit_price_eur"], 61)
        self.assertEqual(result["price_conversion_factor"], 30.5)
        self.assertEqual(result["conversion_source"], "AMOS_UOM_RULE")
        self.assertEqual(result["calculation_status"], "CURRENT_AMOS_PRICE")

    def test_approved_physical_me_price_is_converted_to_mm_without_changing_quantity(self):
        stat = self.stat_line(part="BAR", unit="MM", quantity=70, historical=52.86)
        conversions = parse_physical_price_conversions([{
            "amos_unit": "ME", "stat_unit": "MM", "price_multiplier": "0.001",
            "enabled": True, "approval_reference": "TEST_APPROVAL",
        }])
        result = calculate_line(
            stat, [], {"BAR": [{"measure_unit": "ME", "avg_price": 254.44}]}, {},
            Decimal("20"), conversions)
        self.assertEqual(result["converted_quantity"], 70)
        self.assertAlmostEqual(result["selected_unit_price_eur"], 0.25444)
        self.assertAlmostEqual(result["line_total_eur"], 17.8108)
        self.assertEqual(result["conversion_source"], "APPROVED_PHYSICAL_CONFIG")
        self.assertEqual(result["conversion_status"], "APPROVED_PHYSICAL_PRICE_CONVERSION")

    def test_approved_physical_mm_price_is_converted_to_me(self):
        stat = self.stat_line(part="BAR", unit="ME", quantity=2, historical=50)
        conversions = parse_physical_price_conversions([{
            "amos_unit": "MM", "stat_unit": "ME", "price_multiplier": "1000",
            "enabled": True, "approval_reference": "TEST_APPROVAL",
        }])
        result = calculate_line(
            stat, [], {"BAR": [{"measure_unit": "MM", "avg_price": 0.25}]}, {},
            Decimal("20"), conversions)
        self.assertEqual(result["converted_quantity"], 2)
        self.assertEqual(result["selected_unit_price_eur"], 250)
        self.assertEqual(result["line_total_eur"], 500)

    def test_disabled_physical_conversion_preserves_historical_fallback(self):
        records = [{"amos_unit": "ME", "stat_unit": "MM", "price_multiplier": "0.001",
                    "enabled": False, "approval_reference": "PENDING"}]
        conversions = parse_physical_price_conversions(records)
        result = calculate_line(
            self.stat_line(part="BAR", unit="MM", quantity=70, historical=52.86), [],
            {"BAR": [{"measure_unit": "ME", "avg_price": 254.44}]}, {}, Decimal("20"),
            conversions)
        self.assertAlmostEqual(result["line_total_eur"], 4440.24)
        self.assertEqual(result["fallback_reason"], "MISSING_UOM_RULE")

    def test_pending_physical_conversion_cannot_be_enabled_accidentally(self):
        with self.assertRaises(ValueError):
            parse_physical_price_conversions([{
                "amos_unit": "ME", "stat_unit": "MM", "price_multiplier": "0.001",
                "enabled": True, "approval_reference": "PENDING_CONFIRMATION",
            }])

    def test_physical_conversion_is_not_applied_when_official_rule_exists(self):
        stat = self.stat_line(part="BAR", unit="MM", quantity=70, historical=52.86)
        official = [{"external": "MM", "internal": "ME", "rate": 1,
                     "exchange_base": 1000, "rule_issues": ()}]
        conversions = {("ME", "MM"): {"price_multiplier": Decimal("0.002"),
                                         "approval_reference": "TEST"}}
        result = calculate_line(
            stat, official, {"BAR": [{"measure_unit": "ME", "avg_price": 254.44}]}, {},
            Decimal("20"), conversions)
        self.assertAlmostEqual(result["selected_unit_price_eur"], 0.25444)
        self.assertEqual(result["conversion_source"], "AMOS_UOM_RULE")

    def test_unit_alias_preserves_quantity_before_price_multiplication(self):
        stat = {
            "part_key": "PAINT", "unit_key": "LT", "stat_quantity": 0.1,
        }
        amos = {"PAINT": [{"measure_unit": "L", "avg_price": 106.09}]}
        result = calculate_line(stat, [], amos, {"LT": "L"})
        self.assertEqual(result["effective_stat_unit"], "L")
        self.assertEqual(result["converted_quantity"], 0.1)
        self.assertAlmostEqual(result["line_total_eur"], 10.609)
        self.assertEqual(result["conversion_status"], "UNIT_ALIAS_APPLIED")

    def test_ecb_xml_parser_reads_eur_usd_rate_and_date(self):
        payload = b'''<gesmes:Envelope xmlns:gesmes="x" xmlns="y"><Cube><Cube time="2026-11-20">
            <Cube currency="USD" rate="1.2345"/></Cube></Cube></gesmes:Envelope>'''
        self.assertEqual(parse_ecb_rate(payload), (Decimal("1.2345"), "2026-11-20"))

    def test_web_workflow_returns_downloadable_report_without_persisting_uploads(self):
        config = {
            "input_files": {
                "stat": "excel_files/input/STAT_INITIAL - Copy.xlsx",
                "amos": "excel_files/input/AMOS_uus_süsteem_info.xlsx",
                "uom": "excel_files/input/AMOS_UOM.xlsx",
            },
            "output_file": "excel_files/output/STAT_UOM_vordlusraport.xlsx",
            "stat_columns": {
                "part_number": 6, "unit": 9, "quantity": 8,
                "task_card": 3, "reference": 1,
            },
            "business_rules": {
                "historical_sheet": "AVG", "fallback_markup_percent": 20,
                "approval_limit_eur": 1000, "unit_aliases": {"LT": "L"},
            },
        }
        result = process_workbooks(
            {"stat": self.stat.read_bytes(), "amos": self.amos.read_bytes(),
             "uom": self.uom.read_bytes()},
            config,
            {"rate": "1.25", "date": "2026-11-20", "status": "TEST"},
        )
        self.assertEqual(result["filename"], SORT)
        self.assertTrue(result["report"].startswith(b"PK"))
        self.assertEqual(result["summary"][-1]["task_card"], "SUM_AVG")
        self.assertEqual(result["summary"][-1]["final_task_card_price_eur"], 34.85)
        self.assertEqual(result["exchange"]["status"], "TEST")
        self.assertEqual(len(result["material_calculations"]), 3)

    def test_zero_price_fallback_and_missing_data_are_transferred_to_web_report(self):
        workbook(self.amos, "Sheet1", [
            ("partno", "description", "measure_unit", "avg_price", "sales", "purch_price"),
            ("PART-ZERO", "Zero", "EA", 0, 0, 0),
            ("PART-MISSING", "Missing", "EA", None, 0, None),
        ])
        workbook(self.stat, "INITIAL", [
            ("Reference", None, "TaskCard", None, None, "PN", "Description", "Qty", "Unit", None, "EUR"),
            ("REF-1", None, "TASK-1", None, None, "PART-ZERO", "Zero", 3, "EA", None, 50),
            ("REF-2", None, "TASK-2", None, None, "PART-MISSING", "Missing", 2, "EA", None, None),
        ])
        config = {
            "input_files": {"stat": "excel_files/input/STAT_INITIAL - Copy.xlsx",
                            "amos": "excel_files/input/AMOS_uus_süsteem_info.xlsx",
                            "uom": "excel_files/input/AMOS_UOM.xlsx"},
            "output_file": "excel_files/output/STAT_UOM_vordlusraport.xlsx",
            "stat_columns": {"part_number": 6, "unit": 9, "quantity": 8,
                             "task_card": 3, "reference": 1},
            "business_rules": {"historical_sheet": "AVG", "fallback_markup_percent": 20,
                               "approval_limit_eur": 500, "unit_aliases": {}},
        }
        result = process_workbooks(
            {"stat": self.stat.read_bytes(), "amos": self.amos.read_bytes(),
             "uom": self.uom.read_bytes()}, config,
            {"rate": "1.2", "date": "2026-11-20", "status": "TEST"})
        rows = {row["partno"]: row for row in result["material_calculations"]}
        zero = rows["PART-ZERO"]
        self.assertEqual(zero["amos_avg_price_eur"], 0)
        self.assertEqual(zero["selected_unit_price_eur"], 60)
        self.assertEqual(zero["final_line_eur"], 180)
        self.assertEqual(zero["price_source"], "HISTORICAL_PRICE_PLUS_20_PERCENT")
        self.assertEqual(zero["calculation_status"], "FALLBACK_APPLIED")
        missing = rows["PART-MISSING"]
        self.assertIsNone(missing["amos_avg_price_eur"])
        self.assertEqual(missing["selected_unit_price_eur"], "ANDMED PUUDUVAD")
        self.assertTrue(missing["review_required"])

    def test_repeated_web_processing_does_not_reuse_previous_work_package(self):
        config = {
            "input_files": {"stat": "excel_files/input/STAT_INITIAL - Copy.xlsx",
                            "amos": "excel_files/input/AMOS_uus_süsteem_info.xlsx",
                            "uom": "excel_files/input/AMOS_UOM.xlsx"},
            "output_file": "excel_files/output/STAT_UOM_vordlusraport.xlsx",
            "stat_columns": {"part_number": 6, "unit": 9, "quantity": 8,
                             "task_card": 3, "reference": 1},
            "business_rules": {"historical_sheet": "AVG", "fallback_markup_percent": 20,
                               "approval_limit_eur": 500, "unit_aliases": {"LT": "L"}},
        }
        exchange = {"rate": "1.2", "date": "2026-11-20", "status": "TEST"}
        first = process_workbooks(
            {"stat": self.stat.read_bytes(), "amos": self.amos.read_bytes(),
             "uom": self.uom.read_bytes()}, config, exchange)
        workbook(self.stat, "INITIAL", [
            ("Reference", None, "TaskCard", None, None, "PN", None, "Qty", "Unit", None, "EUR"),
            ("REF-Z", None, "TASK-Z", None, None, "-999-22G1", None, 10, "ME", None, 1),
        ])
        second = process_workbooks(
            {"stat": self.stat.read_bytes(), "amos": self.amos.read_bytes(),
             "uom": self.uom.read_bytes()}, config, exchange)
        self.assertEqual(len(first["material_calculations"]), 3)
        self.assertEqual(len(second["material_calculations"]), 1)
        self.assertEqual(second["material_calculations"][0]["task_card"], "TASK-Z")

    def test_converter_assets_are_inlined_and_missing_assets_are_nonfatal(self):
        assets = self.root / "assets"
        assets.mkdir()
        self.assertIsNone(load_converter_html(assets))
        (assets / "unit_converter.html").write_text(
            '<html><script src="unit_converter.js"></script></html>', encoding="utf-8")
        (assets / "unit_converter.js").write_text("window.converterLoaded = true;", encoding="utf-8")
        rendered = load_converter_html(assets)
        self.assertIn("window.converterLoaded = true;", rendered)
        self.assertNotIn('src="unit_converter.js"', rendered)
