from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from pnfl_pdbtoexcel.config import set_config_path, set_play_path, set_team
from pnfl_pdbtoexcel.pdb_to_excel import PdbWorkbookCreator


TESTS_DIR = Path(__file__).resolve().parents[1]
PDB_PATH = TESTS_DIR / "data" / "2045-2047.pdb"
GAMEPLAN_DATA_DIR = Path(__file__).resolve().parents[3] / "fbpro98-gameplan" / "tests" / "data"
OFFENSE_PLN = GAMEPLAN_DATA_DIR / "offense.pln"
DEFENSE_PLN = GAMEPLAN_DATA_DIR / "defense.pln"
PLAYPOOL_DIR = Path(__file__).resolve().parents[3] / "pnfl-playpool" / "tests" / "data" / "plays"

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EXPECTED_SHEETS_BASE = [
    "Options",
    "Run Plays",
    "Pass Plays",
    "Def Plays",
    "Tendencies",
]


def _require_fixtures() -> None:
    if not PDB_PATH.is_file():
        pytest.skip(f"Missing PDB fixture: {PDB_PATH}")
    if not PLAYPOOL_DIR.is_dir():
        pytest.skip(f"Missing playpool fixture dir: {PLAYPOOL_DIR}")


def _setup_config(tmp_path: Path) -> None:
    config_path = tmp_path / "pdb_to_excel.ini"
    config_path.write_text(
        f"[Settings]\nPlayPath={PLAYPOOL_DIR}\n",
        encoding="utf-8",
    )
    set_config_path(config_path)
    set_team(None)
    set_play_path(None)


def _read_sheet_names(workbook_path: Path) -> list[str]:
    with ZipFile(workbook_path) as archive:
        tree = ET.fromstring(archive.read("xl/workbook.xml"))
    return [
        sheet.attrib["name"]
        for sheet in tree.findall(f".//{{{MAIN_NS}}}sheet")
    ]


def _read_sheet_row_count(workbook_path: Path, sheet_name: str) -> int:
    with ZipFile(workbook_path) as archive:
        tree = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_tree = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_tree.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        for sheet in tree.findall(f".//{{{MAIN_NS}}}sheet"):
            if sheet.attrib["name"] != sheet_name:
                continue
            rid = sheet.attrib[f"{{{REL_NS}}}id"]
            sheet_xml = archive.read(f"xl/{rel_targets[rid]}")
            sheet_tree = ET.fromstring(sheet_xml)
            return len(sheet_tree.findall(f".//{{{MAIN_NS}}}row"))
    raise AssertionError(f"Worksheet '{sheet_name}' not found")


def test_workbook_without_gameplans(tmp_path: Path) -> None:
    _require_fixtures()
    _setup_config(tmp_path)
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_files(str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), True, True, False)

    assert workbook_path.is_file()
    sheets = _read_sheet_names(workbook_path)
    for expected in EXPECTED_SHEETS_BASE:
        assert expected in sheets

    assert _read_sheet_row_count(workbook_path, "Run Plays") == 3703
    assert _read_sheet_row_count(workbook_path, "Pass Plays") == 7889
    assert _read_sheet_row_count(workbook_path, "Def Plays") == 9843
    assert _read_sheet_row_count(workbook_path, "Tendencies") == 369


def test_workbook_with_gameplans(tmp_path: Path) -> None:
    _require_fixtures()
    if not OFFENSE_PLN.is_file() or not DEFENSE_PLN.is_file():
        pytest.skip("Missing gameplan fixtures")
    _setup_config(tmp_path)
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_files(str(PDB_PATH), str(DEFENSE_PLN), str(OFFENSE_PLN))
    creator.create_workbook(str(workbook_path), True, True, False)

    assert workbook_path.is_file()
    sheets = _read_sheet_names(workbook_path)
    for expected in EXPECTED_SHEETS_BASE:
        assert expected in sheets

    assert _read_sheet_row_count(workbook_path, "Run Plays") == 3703
    assert _read_sheet_row_count(workbook_path, "Pass Plays") == 7889
    assert _read_sheet_row_count(workbook_path, "Def Plays") == 9843
    assert _read_sheet_row_count(workbook_path, "Tendencies") == 369


def test_workbook_skip_calcs_and_totals(tmp_path: Path) -> None:
    _require_fixtures()
    _setup_config(tmp_path)
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_files(str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), False, False, False)

    assert workbook_path.is_file()
    sheets = _read_sheet_names(workbook_path)
    for expected in EXPECTED_SHEETS_BASE:
        assert expected in sheets

    # Without totals, fewer rows (no `Total Stats team entries)
    run_rows = _read_sheet_row_count(workbook_path, "Run Plays")
    pass_rows = _read_sheet_row_count(workbook_path, "Pass Plays")
    def_rows = _read_sheet_row_count(workbook_path, "Def Plays")
    assert run_rows < 3703
    assert pass_rows < 7891
    assert def_rows < 9843
