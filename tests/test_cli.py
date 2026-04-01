from __future__ import annotations

import pytest

from pnfl_pdbtoexcel.cli import parse_args
from pnfl_pdbtoexcel.config import get_config, set_config_path, set_pnfl_path, set_team

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
    assert args.config is None
    assert args.skip_calcs is False
    assert args.skip_totals is False
    assert args.team is None
    assert args.pnfl_path is None


def test_parse_args_accepts_all_options(tmp_path: Path) -> None:
    pdb = tmp_path / "test.pdb"
    pdb.touch()
    defense = tmp_path / "def.pln"
    defense.touch()
    offense = tmp_path / "off.pln"
    offense.touch()
    config = tmp_path / "config.ini"
    config.touch()

    args = parse_args([
        str(pdb),
        "output.xlsm",
        "-d", str(defense),
        "-o", str(offense),
        "--config", str(config),
        "-c",
        "-t",
        "--team", "Denver",
        "--pnfl-path", r"E:\PNFL",
    ])
    assert args.plnfile_defense == str(defense)
    assert args.plnfile_offense == str(offense)
    assert args.config == str(config)
    assert args.skip_calcs is True
    assert args.skip_totals is True
    assert args.team == "Denver"
    assert args.pnfl_path == r"E:\PNFL"


def test_parse_args_rejects_non_pdb_extension() -> None:
    with pytest.raises(SystemExit):
        parse_args(["test.txt", "output.xlsx"])


def test_parse_args_rejects_non_excel_output() -> None:
    with pytest.raises(SystemExit):
        parse_args(["test.pdb", "output.csv"])


def test_config_team_override(tmp_path: Path) -> None:
    config_path = tmp_path / "pdb_to_excel.ini"
    config_path.write_text("[Settings]\nTeam=Vikings\n", encoding="utf-8")
    set_config_path(config_path)
    set_team(None)
    set_pnfl_path(None)
    assert get_config()["Settings"]["Team"] == "Vikings"

    set_config_path(config_path)
    set_team("Denver")
    assert get_config()["Settings"]["Team"] == "Denver"
    set_team(None)


def test_config_pnfl_path_override(tmp_path: Path) -> None:
    config_path = tmp_path / "pdb_to_excel.ini"
    config_path.write_text("[Settings]\nPnflPath=C:\\from-config\n", encoding="utf-8")
    set_config_path(config_path)
    set_team(None)
    set_pnfl_path(None)
    assert get_config()["Settings"]["PnflPath"] == "C:\\from-config"

    set_config_path(config_path)
    set_pnfl_path(r"D:\from-cli")
    assert get_config()["Settings"]["PnflPath"] == r"D:\from-cli"
    set_pnfl_path(None)


def test_config_falls_back_to_defaults(tmp_path: Path) -> None:
    from pnfl_pdbtoexcel import config as config_module
    original_path = config_module._config_path
    original_candidates = config_module.CONFIG_CANDIDATES
    config_module._config_path = None
    config_module._config = None
    config_module.CONFIG_CANDIDATES = [tmp_path / "nonexistent.ini"]
    try:
        c = get_config()
        assert c["Settings"]["PnflPath"] == r"C:\SIERRA\FbPro98\PNFL"
        assert c["Settings"]["Team"] == ""
    finally:
        config_module._config_path = original_path
        config_module._config = None
        config_module.CONFIG_CANDIDATES = original_candidates
