#!/usr/bin/env python3
"""One-command runner for the Magnetic MRO test workflow."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
INPUT_DIR = (ROOT / "excel_files" / "input").resolve()
OUTPUT_DIR = (ROOT / "excel_files" / "output").resolve()
GENERATED_OUTPUT_NAMES = {
    "amos_prepared.xlsx",
    "STAT_UOM_vordlusraport.xlsx",
    "STAT_UOM_vordlusraport_friend.xlsx",
}


def load_config(path: Path) -> dict:
    with path.open(encoding="utf-8") as source:
        config = json.load(source)
    if set(config.get("input_files", {})) != {"stat", "amos", "uom"}:
        raise ValueError("config.json peab sisaldama stat, amos ja uom sisendfaile")
    if "output_file" not in config:
        raise ValueError("config.json peab sisaldama output_file väärtust")
    return config


def checked_path(relative: str, directory: Path) -> Path:
    path = (ROOT / relative).resolve()
    if path.parent != directory or path.suffix.lower() != ".xlsx":
        raise ValueError(f"Fail peab asuma otse kaustas {directory}: {relative}")
    return path


def configured_paths(config: dict) -> tuple[dict[str, Path], Path]:
    inputs = {
        name: checked_path(relative, INPUT_DIR)
        for name, relative in config["input_files"].items()
    }
    output = checked_path(config["output_file"], OUTPUT_DIR)
    if output.name != "STAT_UOM_vordlusraport.xlsx":
        raise ValueError("Põhiraporti nimi peab olema STAT_UOM_vordlusraport.xlsx")
    return inputs, output


def remove_files(paths: list[Path]) -> None:
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except PermissionError as error:
            raise PermissionError(f"Sulge fail Excelis ja proovi uuesti: {path}") from error


def generated_outputs() -> list[Path]:
    return [OUTPUT_DIR / name for name in sorted(GENERATED_OUTPUT_NAMES)]


def run(command: list[str]) -> None:
    print("\n>", " ".join(command))
    subprocess.run(command, cwd=ROOT, check=True)


def main() -> int:
    parser = argparse.ArgumentParser(description="Magnetic MRO testitöövoog")
    parser.add_argument("--config", type=Path, default=ROOT / "config.json")
    parser.add_argument("--fresh", action="store_true", help="eemalda vanad väljundid enne käivitamist")
    parser.add_argument("--clean", action="store_true", help="eemalda ainult genereeritud väljundid")
    parser.add_argument("--clean-all", action="store_true", help="eemalda seadistatud sisendid ja väljundid")
    parser.add_argument("--run-tests", action="store_true", help="käivita automaattestid enne töötlemist")
    args = parser.parse_args()

    config = load_config(args.config.resolve())
    inputs, output = configured_paths(config)

    if args.clean_all:
        remove_files(generated_outputs() + list(inputs.values()))
        print("Input ja output on puhastatud. Laadi uued sisendfailid input-kausta.")
        return 0
    if args.clean:
        remove_files(generated_outputs())
        print("Genereeritud väljundfailid on eemaldatud.")
        return 0

    missing = [path for path in inputs.values() if not path.is_file()]
    if missing:
        print("Puuduvad sisendfailid:", file=sys.stderr)
        for path in missing:
            print(f"- {path}", file=sys.stderr)
        return 2

    if args.fresh:
        remove_files(generated_outputs())
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    python = sys.executable
    columns = config["stat_columns"]
    rules = config.get("business_rules", {})
    aliases = rules.get("unit_aliases", {})
    alias_arguments = [item for pair in aliases.items() for item in ("--unit-alias", f"{pair[0]}={pair[1]}")]
    if args.run_tests:
        run([python, "-B", "-m", "unittest", "discover", "-s", "tests", "-q"])
    run([python, "-B", "scripts/sorter.py", "--stat", str(inputs["stat"]),
         "--uom", str(inputs["uom"]), "--amos", str(inputs["amos"]),
         "--part-col", str(columns["part_number"]), "--unit-col", str(columns["unit"]),
         "--quantity-col", str(columns["quantity"]), "--task-card-col", str(columns["task_card"]),
         "--reference-col", str(columns["reference"]),
         "--historical-sheet", str(rules.get("historical_sheet", "AVG")),
         "--fallback-markup", str(rules.get("fallback_markup_percent", 20)),
         *alias_arguments, "--output", str(output)])

    print(f"\nVALMIS. Uuendatud raport: {output}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, subprocess.CalledProcessError) as error:
        print(f"\nVIGA: {error}", file=sys.stderr)
        raise SystemExit(1)
