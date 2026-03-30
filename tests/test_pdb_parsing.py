import ctypes
import json
from pathlib import Path

import pytest

from pnfl_pdbtoexcel.PdbToExcel import InvalidPDBError, PDB, PLAY_DATA


TEST_DATA_DIR = Path(__file__).resolve().parent / "data"
REAL_PDB_PATH = TEST_DATA_DIR / "2045-2047.pdb"
PLAYS_FIXTURE_PATH = TEST_DATA_DIR / "2045-2047.plays.json"

PLAY_TYPE_NAMES = {
    PLAY_DATA.PLAY_TYPE.RUN: "RUN",
    PLAY_DATA.PLAY_TYPE.PASS: "PASS",
    PLAY_DATA.PLAY_TYPE.SPECIAL: "SPECIAL",
    PLAY_DATA.PLAY_TYPE.DEFENSE: "DEFENSE",
    PLAY_DATA.PLAY_TYPE.ONSIDE: "ONSIDE",
}
PLAY_DATA_FIELD_NAMES = [field_definition[0] for field_definition in PLAY_DATA._fields_]


def _normalize_play_data(play_data):
    return {
        field_name: int(getattr(play_data, field_name))
        for field_name in PLAY_DATA_FIELD_NAMES
        if field_name not in ("team_name", "play_name")
    }


def _normalize_plays(pdb):
    normalized = {}
    for play_type, plays_for_type in pdb.plays.items():
        type_name = PLAY_TYPE_NAMES[play_type]
        normalized[type_name] = {}
        for (team_name, play_name), play_data in sorted(plays_for_type.items()):
            normalized[type_name][f"{team_name}|{play_name}"] = _normalize_play_data(play_data)
    return normalized


def test_real_pdb_matches_self_plays_snapshot_fixture():
    pdb = PDB(REAL_PDB_PATH)
    expected_plays = json.loads(PLAYS_FIXTURE_PATH.read_text(encoding="utf-8"))

    assert _normalize_plays(pdb) == expected_plays


def test_real_pdb_contains_expected_tendency_count_and_sample_records():
    pdb = PDB(REAL_PDB_PATH)

    run_key = ("Pittsburgh", "SF62wish")
    pass_key = ("Chicago", "KC2AFLYZ")
    defense_key = ("Minnesota", "LV32rmQ2")

    assert len(pdb.tendencies) == 23
    assert run_key in pdb.plays[PLAY_DATA.PLAY_TYPE.RUN]
    assert pass_key in pdb.plays[PLAY_DATA.PLAY_TYPE.PASS]
    assert defense_key in pdb.plays[PLAY_DATA.PLAY_TYPE.DEFENSE]


def test_pdb_raises_for_invalid_data_type(tmp_path):
    pdb_path = tmp_path / "invalid-type.pdb"
    with open(pdb_path, "wb") as pdb_file:
        pdb_file.write(bytes([9]))
        pdb_file.write(b"\x00" * ctypes.sizeof(PLAY_DATA))

    with pytest.raises(InvalidPDBError):
        PDB(pdb_path)
