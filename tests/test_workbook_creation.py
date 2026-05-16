from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest

from pnfl_pdbtoexcel.config import CategoryOrder, Config
from pnfl_pdbtoexcel.pdb import PLAY_DATA
from pnfl_pdbtoexcel.workbook_creator import PdbWorkbookCreator

TESTS_DIR = Path(__file__).resolve().parent
WORKSPACE_ROOT = Path(__file__).resolve().parents[2]
PDB_PATH = TESTS_DIR / "data" / "2045-2047.pdb"
GAMEPLAN_DATA_DIR = WORKSPACE_ROOT / "fbpro98-gameplan" / "tests" / "data"
OFFENSE_PLN = GAMEPLAN_DATA_DIR / "offense.pln"
DEFENSE_PLN = GAMEPLAN_DATA_DIR / "defense.pln"
PLAYPOOL_DIR = WORKSPACE_ROOT / "pnfl-playpool" / "tests" / "data" / "plays"

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

EXPECTED_SHEETS_WITH_CATEGORIES = EXPECTED_SHEETS_BASE + [
    "Run Categories",
    "Pass Categories",
    "Def Categories",
]


def _require_fixtures() -> None:
    if not PDB_PATH.is_file():
        pytest.skip(f"Missing PDB fixture: {PDB_PATH}")
    if not PLAYPOOL_DIR.is_dir():
        pytest.skip(f"Missing playpool fixture dir: {PLAYPOOL_DIR}")


def _make_config() -> tuple[Config, CategoryOrder]:
    config = Config(play_path=str(PLAYPOOL_DIR))
    category_order: CategoryOrder = {
        PLAY_DATA.PLAY_TYPE.RUN: ["RL", "RM", "RR", "GLR"],
        PLAY_DATA.PLAY_TYPE.PASS: ["PSL", "PSM", "PSR", "PML", "PMM", "PMR", "PLR", "PRD", "GLP"],
        PLAY_DATA.PLAY_TYPE.DEFENSE: [
            "RunLeft",
            "RunMiddle",
            "RunRight",
            "RunDazzle",
            "PassShort",
            "PassMedium",
            "PassLong",
            "PassDazzle",
            "GLrun",
            "GLpass",
        ],
    }
    return config, category_order


def test_validate_category_order_rejects_missing_run_categories() -> None:
    _, complete = _make_config()
    incomplete: CategoryOrder = {
        **complete,
        PLAY_DATA.PLAY_TYPE.RUN: ["RL", "RM"],  # missing GLR, RR
    }
    with pytest.raises(ValueError, match=r"RUN: missing \['GLR', 'RR'\]"):
        PdbWorkbookCreator._validate_category_order(incomplete)


def test_validate_category_order_rejects_missing_play_type() -> None:
    _, complete = _make_config()
    missing_defense = {k: v for k, v in complete.items() if k != PLAY_DATA.PLAY_TYPE.DEFENSE}
    with pytest.raises(ValueError, match="DEFENSE: missing from CategoryOrder"):
        PdbWorkbookCreator._validate_category_order(missing_defense)


def test_validate_category_order_accepts_complete() -> None:
    _, complete = _make_config()
    PdbWorkbookCreator._validate_category_order(complete)


def _read_sheet_names(workbook_path: Path) -> list[str]:
    with ZipFile(workbook_path) as archive:
        tree = ET.fromstring(archive.read("xl/workbook.xml"))
    return [sheet.attrib["name"] for sheet in tree.findall(f".//{{{MAIN_NS}}}sheet")]


def _read_sheet_row_count(workbook_path: Path, sheet_name: str) -> int:
    with ZipFile(workbook_path) as archive:
        tree = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_tree = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_tree.findall(f"{{{PKG_REL_NS}}}Relationship")
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
    config, category_order = _make_config()
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_config(config, category_order, str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), True, True, False)

    assert workbook_path.is_file()
    # Exact match: also verifies the category worksheets are NOT present.
    assert _read_sheet_names(workbook_path) == EXPECTED_SHEETS_BASE

    assert _read_sheet_row_count(workbook_path, "Run Plays") == 3703
    assert _read_sheet_row_count(workbook_path, "Pass Plays") == 7889
    assert _read_sheet_row_count(workbook_path, "Def Plays") == 9843
    assert _read_sheet_row_count(workbook_path, "Tendencies") == 369


def test_workbook_with_gameplans(tmp_path: Path) -> None:
    _require_fixtures()
    if not OFFENSE_PLN.is_file() or not DEFENSE_PLN.is_file():
        pytest.skip("Missing gameplan fixtures")
    config, category_order = _make_config()
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_config(config, category_order, str(PDB_PATH), str(DEFENSE_PLN), str(OFFENSE_PLN))
    creator.create_workbook(str(workbook_path), True, True, False)

    assert workbook_path.is_file()
    # Exact match: also verifies the category worksheets are NOT present.
    assert _read_sheet_names(workbook_path) == EXPECTED_SHEETS_BASE

    assert _read_sheet_row_count(workbook_path, "Run Plays") == 3703
    assert _read_sheet_row_count(workbook_path, "Pass Plays") == 7889
    assert _read_sheet_row_count(workbook_path, "Def Plays") == 9843
    assert _read_sheet_row_count(workbook_path, "Tendencies") == 369


def test_workbook_skip_calcs_and_totals(tmp_path: Path) -> None:
    _require_fixtures()
    config, category_order = _make_config()
    workbook_path = tmp_path / "output.xlsx"

    creator = PdbWorkbookCreator.from_config(config, category_order, str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), False, False, False)

    assert workbook_path.is_file()
    # Exact match: also verifies the category worksheets are NOT present.
    assert _read_sheet_names(workbook_path) == EXPECTED_SHEETS_BASE

    # Without totals, fewer rows (no Total Stats team entries)
    run_rows = _read_sheet_row_count(workbook_path, "Run Plays")
    pass_rows = _read_sheet_row_count(workbook_path, "Pass Plays")
    def_rows = _read_sheet_row_count(workbook_path, "Def Plays")
    assert run_rows < 3703
    assert pass_rows < 7891
    assert def_rows < 9843


# ---------------------------------------------------------------------------
# Cell-value tests
#
# These read actual cell values out of the generated workbook and assert them
# for the header plus a few representative rows, covering every column of each
# play worksheet. Raw columns (Att, Yards, Fumbles, ...) trace to PLAY_DATA,
# which tests/test_pdb_parsing.py snapshot-verifies; the derived columns (Avg,
# percentages, Y/Att, ...) are each arithmetically consistent with the raw
# columns in the same row. Category worksheets are intentionally not covered.
# ---------------------------------------------------------------------------


def _col_to_index(cell_ref: str) -> int:
    """Map an A1-style cell reference to a zero-based column index ('C2' -> 2)."""
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for ch in letters:
        index = index * 26 + (ord(ch.upper()) - ord("A") + 1)
    return index - 1


def _read_sheet_cells(workbook_path: Path, sheet_name: str) -> list[list]:
    """Return every worksheet row as a position-correct list of typed cells.

    String cells stay str; numeric cells become int (when integral) or float
    rounded to 3 places. xlsxwriter omits empty-string cells, so gaps are
    filled with "" and every row is padded to the header width.
    """
    with ZipFile(workbook_path) as archive:
        tree = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_tree = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_tree.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        try:
            sst = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared = ["".join(t.text or "" for t in si.iter(f"{{{MAIN_NS}}}t")) for si in sst]
        except KeyError:
            shared = []
        rid = None
        for sheet in tree.findall(f".//{{{MAIN_NS}}}sheet"):
            if sheet.attrib["name"] == sheet_name:
                rid = sheet.attrib[f"{{{REL_NS}}}id"]
        if rid is None:
            raise AssertionError(f"Worksheet '{sheet_name}' not found")
        sheet_tree = ET.fromstring(archive.read(f"xl/{rel_targets[rid]}"))

    rows: list[list] = []
    for row in sheet_tree.findall(f".//{{{MAIN_NS}}}row"):
        cells: dict[int, object] = {}
        for c in row.findall(f"{{{MAIN_NS}}}c"):
            v = c.find(f"{{{MAIN_NS}}}v")
            if v is None or v.text is None:
                continue
            if c.attrib.get("t") == "s":
                cells[_col_to_index(c.attrib["r"])] = shared[int(v.text)]
            else:
                number = float(v.text)
                cells[_col_to_index(c.attrib["r"])] = int(number) if number.is_integer() else round(number, 3)
        width = max(cells) + 1 if cells else 0
        rows.append([cells.get(i, "") for i in range(width)])

    header_width = len(rows[0]) if rows else 0
    for row_cells in rows:
        if len(row_cells) < header_width:
            row_cells.extend([""] * (header_width - len(row_cells)))
    return rows


@pytest.fixture(scope="module")
def workbook_no_gameplans(tmp_path_factory) -> Path:
    """Build the workbook once (no gameplans, calculations + totals on)."""
    _require_fixtures()
    config, category_order = _make_config()
    workbook_path = tmp_path_factory.mktemp("workbook") / "output.xlsx"
    creator = PdbWorkbookCreator.from_config(config, category_order, str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), True, True, False)
    return workbook_path


def test_run_plays_cell_values(workbook_no_gameplans: Path) -> None:
    rows = _read_sheet_cells(workbook_no_gameplans, "Run Plays")
    assert rows[0] == [
        "Team", "Category", "Slot 1", "Slot 2", "Play", "Type",
        "Rushes", "Yards", "Avg", "Fumbles", "Fumble %", "TD", "TD %",
    ]  # fmt: skip
    assert rows[1] == ["Atlanta", "RL", "", "", "AF22rl12", "", 3, 9, 3, 0, 0, 0, 0]
    assert rows[3] == ["Atlanta", "RL", "", "", "AZ26RL62", "", 3, 10, 3.3, 0, 0, 1, 0.333]
    # Total Stats row, with the Type column populated ("QB draw").
    assert rows[-1] == ["Total Stats", "GLR", "", "", "WR10GR01", "QB draw", 15, 43, 2.9, 1, 0.067, 4, 0.267]


def test_pass_plays_cell_values(workbook_no_gameplans: Path) -> None:
    rows = _read_sheet_cells(workbook_no_gameplans, "Pass Plays")
    assert rows[0] == [
        "Team", "Category", "Slot 1", "Slot 2", "Play", "Type",
        "Comp", "Att", "Comp %", "Yards", "Y/Comp", "Y/Att",
        "Fumbles", "Int", "Int %", "Sacks", "Sack %", "TD", "TD %",
    ]  # fmt: skip
    assert rows[1] == ["Atlanta", "PSL", "", "", "AF1AwagR", "", 3, 5, 0.6, 25, 8.3, 5, 0, 0, 0, 0, 0, 1, 0.2]
    assert rows[3] == ["Atlanta", "PSL", "", "", "AT1AWhXd", "", 3, 4, 0.75, 24, 8, 6, 0, 0, 0, 0, 0, 1, 0.25]
    assert rows[-1] == ["Total Stats", "GLP", "", "", "TBGXqouT", "", 1, 4, 0.25, 26, 26, 6.5, 0, 0, 0, 0, 0, 0, 0]


def test_def_plays_cell_values(workbook_no_gameplans: Path) -> None:
    rows = _read_sheet_cells(workbook_no_gameplans, "Def Plays")
    assert rows[0] == [
        "Team", "Category", "Slot 1", "Slot 2", "Play", "Type",
        "Calls", "Yards", "Avg", "vs Run", "Yards", "Avg", "vs Pass", "Yards", "Avg",
        "Fumbles", "Int", "TO %", "Sacks", "Sack %", "TD/Def", "TD/Off", "TD/Off %",
    ]  # fmt: skip
    # First row, with Type ("3-4") populated. Calls/Yards/Avg = run + pass combined.
    assert rows[1] == [
        "Atlanta", "RunLeft", "", "", "AF31rl3H", "3-4",
        5, 8, 1.6, 4, 8, 2, 1, 0, 0, 1, 0, 0.2, 1, 0.2, 0, 1, 0.2,
    ]  # fmt: skip
    # Negative yardage row.
    assert rows[3] == [
        "Atlanta", "RunLeft", "", "", "AF32rlZD", "3-4",
        3, -1, -0.3, 1, -1, -1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    ]  # fmt: skip
    assert rows[-1] == [
        "Total Stats", "GLpass", "", "", "SF61GP1R", "",
        1, 2, 2, 0, 0, 0, 1, 2, 2, 0, 0, 0, 0, 0, 0, 0, 0,
    ]  # fmt: skip


def test_tendencies_cell_values(workbook_no_gameplans: Path) -> None:
    rows = _read_sheet_cells(workbook_no_gameplans, "Tendencies")
    assert rows[0] == ["Team", "Situation", "Runs", "Passes"]
    assert rows[1] == ["Atlanta", "First and 0-1", 9, 1]
    assert rows[3] == ["Atlanta", "First and 6-10", 515, 752]
    assert rows[-1] == ["Washington", "Fourth and 10+", 0, 3]


# ---------------------------------------------------------------------------
# Category worksheets
#
# These exist only when Config.include_category_worksheets is True. The same
# trust basis as the play-worksheet tests applies: raw columns aggregate
# PLAY_DATA, derived columns are arithmetically consistent within each row.
# ---------------------------------------------------------------------------


@pytest.fixture(scope="module")
def workbook_with_categories(tmp_path_factory) -> Path:
    """Build the workbook once with category worksheets enabled."""
    _require_fixtures()
    _, category_order = _make_config()
    config = Config(play_path=str(PLAYPOOL_DIR), include_category_worksheets=True)
    workbook_path = tmp_path_factory.mktemp("workbook_categories") / "output.xlsx"
    creator = PdbWorkbookCreator.from_config(config, category_order, str(PDB_PATH), None, None)
    creator.create_workbook(str(workbook_path), True, True, False)
    return workbook_path


def test_category_worksheets_added(workbook_with_categories: Path) -> None:
    assert _read_sheet_names(workbook_with_categories) == EXPECTED_SHEETS_WITH_CATEGORIES


def test_run_categories_cell_values(workbook_with_categories: Path) -> None:
    rows = _read_sheet_cells(workbook_with_categories, "Run Categories")
    assert rows[0] == ["Team", "Category", "Rushes", "Yards", "Avg", "Fumbles", "Fumble %", "TD", "TD %"]
    assert rows[1] == ["Atlanta", "RL", 97, 407, 4.2, 2, 0.021, 8, 0.082]
    assert rows[3] == ["Atlanta", "RR", 175, 750, 4.3, 4, 0.023, 7, 0.04]
    assert rows[-1] == ["Total Stats", "GLR", 798, 2378, 3, 14, 0.018, 114, 0.143]


def test_pass_categories_cell_values(workbook_with_categories: Path) -> None:
    rows = _read_sheet_cells(workbook_with_categories, "Pass Categories")
    assert rows[0] == [
        "Team", "Category", "Comp", "Att", "Comp %", "Yards", "Y/Comp", "Y/Att",
        "Fumbles", "Int", "Int %", "Sacks", "Sack %", "TD", "TD %",
    ]  # fmt: skip
    assert rows[1] == ["Atlanta", "PSL", 157, 254, 0.62, 1263, 8, 5, 1, 4, 0.016, 2, 0.008, 4, 0.016]
    assert rows[3] == ["Atlanta", "PSR", 70, 115, 0.61, 520, 7.4, 4.5, 1, 2, 0.017, 5, 0.043, 4, 0.035]
    assert rows[-1] == ["Total Stats", "GLP", 486, 655, 0.74, 1706, 3.5, 2.6, 5, 9, 0.014, 20, 0.031, 137, 0.209]


def test_def_categories_cell_values(workbook_with_categories: Path) -> None:
    rows = _read_sheet_cells(workbook_with_categories, "Def Categories")
    assert rows[0] == [
        "Team", "Category", "vs Run", "Yards", "Avg", "vs Pass", "Yards", "Avg",
        "Fumbles", "Int", "TO %", "Sacks", "Sack %", "TD/Def", "TD/Off", "TD/Off %",
    ]  # fmt: skip
    assert rows[1] == ["Atlanta", "RunLeft", 459, 2206, 4.8, 553, 3938, 7.1, 11, 10, 0.021, 27, 0.027, 0, 37, 0.037]
    assert rows[3] == ["Atlanta", "RunRight", 28, 52, 1.9, 20, 98, 4.9, 1, 0, 0.019, 2, 0.038, 0, 4, 0.075]
    assert rows[-1] == ["Total Stats", "GLpass", 87, 84, 1, 133, 839, 6.3, 2, 5, 0.032, 5, 0.023, 0, 68, 0.308]
