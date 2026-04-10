from __future__ import annotations

import pytest

from pnfl_pdbtoexcel.cli import parse_args
from pnfl_pdbtoexcel.config import load_config

from pathlib import Path


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
    assert args.plnfile_defense is None
    assert args.plnfile_offense is None
    assert args.plnfile_defense_2 is None
    assert args.plnfile_offense_2 is None
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

    args = parse_args([
        str(pdb),
        "output.xlsm",
        "-d", str(defense),
        "-d2", str(defense_2),
        "-o", str(offense),
        "-o2", str(offense_2),
        "--config", str(config),
        "-c",
        "-t",
        "--play-path", r"E:\PNFL",
    ])
    assert args.plnfile_defense == str(defense)
    assert args.plnfile_defense_2 == str(defense_2)
    assert args.plnfile_offense == str(offense)
    assert args.plnfile_offense_2 == str(offense_2)
    assert args.config == str(config)
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
    assert load_config(config_path=config_path).Settings.PlayPath == "C:\\from-config"
    assert load_config(config_path=config_path, play_path=r"D:\from-cli").Settings.PlayPath == r"D:\from-cli"


def test_config_falls_back_to_defaults(tmp_path: Path) -> None:
    nonexistent = tmp_path / "nonexistent.ini"
    c = load_config(config_path=nonexistent)
    assert c.Settings.PlayPath == r"C:\SIERRA\FbPro98\PNFL"


def test_config_loads_category_order(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"
    config_path.write_text(
        "[Settings]\n\n"
        "[CategoryOrder]\n"
        "RunCategories =\n    RL\n    RM\n    RR\n"
        "PassCategories =\n    PSL\n    PSM\n"
        "DefenseCategories =\n    RunLeft\n    PassShort\n",
        encoding="utf-8",
    )
    c = load_config(config_path=config_path)
    assert c.CategoryOrder.RunCategories == ["RL", "RM", "RR"]
    assert c.CategoryOrder.PassCategories == ["PSL", "PSM"]
    assert c.CategoryOrder.DefenseCategories == ["RunLeft", "PassShort"]


def test_config_empty_category_order(tmp_path: Path) -> None:
    config_path = tmp_path / "convert-pdb.ini"
    config_path.write_text("[Settings]\n", encoding="utf-8")
    c = load_config(config_path=config_path)
    assert c.CategoryOrder.RunCategories == []
    assert c.CategoryOrder.PassCategories == []
    assert c.CategoryOrder.DefenseCategories == []
