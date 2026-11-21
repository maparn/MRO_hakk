"""Ownership, literal cell values and atomic replacement of generated reports."""

import os
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.parsers import expat
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.cell import WriteOnlyCell


# Reports created before ownership metadata was introduced.
LEGACY_SHEETS = {
    "amos_prepared.xlsx": ("Prepared", "UOM_Map", "Issues", "Summary"),
    "STAT_UOM_vordlusraport.xlsx": (
        "Kokkuvote", "Vastavused", "Unit ulevaatus", "Ainult STAT", "Ainult UOM",
        "UOM variandid", "Duplikaadid", "STAT puuduv partno", "UOM puuduv partno",
    ),
    "STAT_UOM_vordlusraport_friend.xlsx": (
        "Kokkuvote", "Vastavused", "Unit erinevused", "Ainult STAT", "Ainult UOM",
        "Duplikaadid", "STAT puuduv partno", "UOM puuduv partno",
    ),
}


def literal_cell(sheet, value):
    """Preserve imported strings exactly, including formula-like text and identifiers."""
    if not isinstance(value, str):
        return value
    cell = WriteOnlyCell(sheet, value=value)
    cell.data_type = "s"
    return cell


def validate_output(output, root, inputs, *, owner):
    """Only this project's dedicated, direct output directory may receive reports."""
    root = Path(root).resolve()
    directory = root / "excel_files" / "output"
    output = Path(output).absolute()
    if directory.resolve() != directory or output.parent.resolve() != directory:
        raise ValueError(f"Report must be directly inside {directory}")
    if output.is_symlink() or output.suffix.lower() != ".xlsx":
        raise ValueError(f"Expected a regular .xlsx report path: {output}")
    if owner not in LEGACY_SHEETS:
        raise ValueError(f"Unknown report type: {owner}")
    if any(output.name.casefold() == name.casefold() for name in LEGACY_SHEETS if name != owner):
        raise ValueError(f"Path is reserved for another report: {output}")
    if output.exists():
        if not output.is_file():
            raise ValueError(f"Expected a file: {output}")
        if any(Path(source).exists() and output.samefile(source) for source in inputs):
            raise ValueError(f"Output overlaps an input workbook: {output}")
    return output


def _check_existing(output, owner):
    with output.open("rb") as source:
        book = load_workbook(source, read_only=True)
        try:
            legacy = (book.properties.identifier is None and output.name == owner
                      and tuple(book.sheetnames) == LEGACY_SHEETS[owner])
            if not legacy and book.properties.identifier != f"MRO_hakk/{owner}":
                raise ValueError(f"Existing workbook is not owned by this report: {output}")
        finally:
            book.close()

        # Even a generated workbook may have been edited by the Excel developer.
        def reject_formula(name, _attributes):
            if name.rsplit(" ", 1)[-1] == "f":
                raise ValueError(f"Existing report contains Excel formulas: {output}")

        with ZipFile(source) as archive:
            for name in archive.namelist():
                if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                    parser = expat.ParserCreate(namespace_separator=" ")
                    parser.StartElementHandler = reject_formula
                    with archive.open(name) as sheet:
                        parser.ParseFile(sheet)


def save_report(book, output, root, inputs, *, owner):
    output = validate_output(output, root, inputs, owner=owner)
    output.parent.mkdir(parents=True, exist_ok=True)
    book.properties.identifier = f"MRO_hakk/{owner}"
    with NamedTemporaryFile(dir=output.parent, prefix=".mro-", suffix=".tmp", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        book.save(temporary)
        if output.exists():
            _check_existing(output, owner)
        os.replace(temporary, output)
    finally:
        temporary.unlink(missing_ok=True)
        book.close()
