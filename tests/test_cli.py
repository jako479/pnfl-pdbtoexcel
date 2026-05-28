from __future__ import annotations

from pathlib import Path

import pytest

from pnfl_pdbtoexcel.cli import parse_args
from pnfl_pdbtoexcel.config import load_category_order, load_config
from pnfl_pdbtoexcel.pdb import PLAY_DATA


def test_parse_args_requires_pdbfile_and_outputfile() -> None:
    with pytest.raises(SystemExit):
        parse_args([])
    with pytest.raises(SystemExit):
        parse_args(["test.pdb"])


def test_parse_args_accepts_positional_args(tmp_path: Path) -> None:
    pdb = tmp_path / "test.pdb"
    pdb.touch()
    args = parse_args([str(pdb), "output.xlsx"])
    assert args.pdbfile == str(pdb)
    assert args.outputfile == "output.xlsx"
    assert args.pln_def is None
    assert args.pln_off is None
    assert args.pln_def_2 is None
    assert args.pln_off_2 is None
    assert args.config is None
    assert args.skip_calcs is False
    assert args.skip_totals is False
    assert args.play_path is None


def test_parse_args_accepts_all_options(tmp_path: Path) -> None:
    pdb = tmp_path / "test.pdb"
    pdb.touch()
    defense = tmp_path / "def.pln"
    defense.touch()
    defense_2 = tmp_path / "def2.pln"
    defense_2.touch()
    offense = tmp_path / "off.pln"
    offense.touch()
    offense_2 = tmp_path / "off2.pln"
    offense_2.touch()
    config = tmp_path / "config.ini"
    config.touch()

    args = parse_args(
        [
            str(pdb),
            "output.xlsm",
            "-d",
            str(defense),
            "-d2",
            str(defense_2),
            "-o",
            str(offense),
            "-o2",
            str(offense_2),
            "--config",
            str(config),
            "--skip-calcs",
            "--skip-totals",
            "--play-path",
            r"E:\PNFL",
        ]
    )
    assert args.pln_def == str(defense)
    assert args.pln_def_2 == str(defense_2)
    assert args.pln_off == str(offense)
    assert args.pln_off_2 == str(offense_2)
    assert args.config == Path(str(config))
    assert args.skip_calcs is True
    assert args.skip_totals is True
    assert args.play_path == r"E:\PNFL"


def test_parse_args_rejects_non_pdb_extension() -> None:
    with pytest.raises(SystemExit):
        parse_args(["test.txt", "output.xlsx"])


def test_parse_args_rejects_non_excel_output() -> None:
    with pytest.raises(SystemExit):
        parse_args(["test.pdb", "output.csv"])


def test_config_play_path_override(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"
    config_path.write_text("[Settings]\nPlayPath=C:\\from-config\n", encoding="utf-8")
    assert load_config(path=config_path).play_path == "C:\\from-config"
    assert load_config(path=config_path, play_path=r"D:\from-cli").play_path == r"D:\from-cli"


def test_config_falls_back_to_defaults(tmp_path: Path) -> None:
    nonexistent = tmp_path / "nonexistent.ini"
    c = load_config(path=nonexistent)
    assert c.play_path == r"C:\SIERRA\FbPro98\PNFL"
    assert c.exclude_sacks_from_pass_attempts is True


def test_config_exclude_sacks_from_pass_attempts(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"

    config_path.write_text("[Settings]\nExcludeSacksFromPassAttempts=Yes\n", encoding="utf-8")
    assert load_config(path=config_path).exclude_sacks_from_pass_attempts is True

    config_path.write_text("[Settings]\nExcludeSacksFromPassAttempts=No\n", encoding="utf-8")
    assert load_config(path=config_path).exclude_sacks_from_pass_attempts is False


def test_loads_category_order(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"
    config_path.write_text(
        "[Settings]\n\n"
        "[CategoryOrder]\n"
        "RunCategories =\n    RL\n    RM\n    RR\n"
        "PassCategories =\n    PSL\n    PSM\n"
        "DefenseCategories =\n    RunLeft\n    PassShort\n",
        encoding="utf-8",
    )
    co = load_category_order(path=config_path)
    assert co[PLAY_DATA.PLAY_TYPE.RUN] == ["RL", "RM", "RR"]
    assert co[PLAY_DATA.PLAY_TYPE.PASS] == ["PSL", "PSM"]
    assert co[PLAY_DATA.PLAY_TYPE.DEFENSE] == ["RunLeft", "PassShort"]


def test_empty_category_order(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"
    config_path.write_text("[Settings]\n", encoding="utf-8")
    co = load_category_order(path=config_path)
    assert co[PLAY_DATA.PLAY_TYPE.RUN] == []
    assert co[PLAY_DATA.PLAY_TYPE.PASS] == []
    assert co[PLAY_DATA.PLAY_TYPE.DEFENSE] == []
