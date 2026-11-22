"""Isolated in-memory-facing workflow used by the Streamlit web application."""

from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import load_workbook

from .sorter import (
    OUTPUT_DEFAULT,
    build_report,
    parse_unit_aliases,
    read_amos_for_stat,
    read_historical_avg,
    read_stat,
    read_uom,
)


def _temporary_path(root: Path, configured_path: str) -> Path:
    relative = Path(configured_path)
    if relative.is_absolute() or ".." in relative.parts:
        raise ValueError("Invalid configured file path")
    path = (root / relative).resolve()
    if root.resolve() not in path.parents:
        raise ValueError("Configured file path escapes the temporary workspace")
    return path


def _read_sheet(report: Path, sheet_name: str) -> list[dict]:
    book = load_workbook(report, read_only=True, data_only=True)
    try:
        values = book[sheet_name].iter_rows(values_only=True)
        headers = next(values)
        return [dict(zip(headers, row)) for row in values]
    finally:
        book.close()


def process_workbooks(payloads: dict[str, bytes], config: dict, exchange: dict) -> dict:
    """Process uploaded workbooks without persisting user data after the call."""
    expected = {"stat", "amos", "uom"}
    if set(payloads) != expected or any(not payloads[name] for name in expected):
        raise ValueError("Three non-empty Excel files are required")

    with TemporaryDirectory(prefix="mro-web-") as directory:
        root = Path(directory)
        inputs = {
            name: _temporary_path(root, config["input_files"][name])
            for name in expected
        }
        for name, path in inputs.items():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(payloads[name])

        output = _temporary_path(root, config["output_file"])
        output.parent.mkdir(parents=True, exist_ok=True)
        columns = config["stat_columns"]
        rules = config.get("business_rules", {})
        stat_rows = read_stat(
            inputs["stat"], columns["part_number"], columns["unit"],
            quantity_col=columns["quantity"], task_card_col=columns["task_card"],
            reference_col=columns["reference"], description_col=columns.get("description", 7),
            nrc_col=columns.get("nrc", 4), factor_col=columns.get("factor", 5),
            historical_price_col=columns.get("historical_unit_price_eur", 11),
        )
        uom_rows = read_uom(inputs["uom"])
        amos_index = read_amos_for_stat(inputs["amos"], stat_rows)
        historical_avg = read_historical_avg(
            inputs["stat"], rules.get("historical_sheet", "AVG"))
        aliases = parse_unit_aliases([
            f"{source}={target}"
            for source, target in rules.get("unit_aliases", {}).items()
        ])
        stats = build_report(
            stat_rows, uom_rows, output, columns["part_number"], columns["unit"],
            list(inputs.values()), amos_index=amos_index, historical_avg=historical_avg,
            fallback_markup=Decimal(str(rules.get("fallback_markup_percent", 20))),
            aliases=aliases,
            approval_limit=Decimal(str(rules.get("approval_limit_eur", 500))),
            exchange_rate=Decimal(str(exchange["rate"])),
            exchange_date=exchange["date"], exchange_status=exchange["status"],
            project_root=root,
        )
        task_summary = _read_sheet(output, "Task Card Summary")
        reference_summary = _read_sheet(output, "Reference Summary")
        material_calculations = _read_sheet(output, "Material Calculations")
        return {
            "report": output.read_bytes(),
            "filename": OUTPUT_DEFAULT.name,
            "summary": task_summary,
            "task_summary": task_summary,
            "reference_summary": reference_summary,
            "material_calculations": material_calculations,
            "stats": stats,
            "exchange": exchange,
        }
