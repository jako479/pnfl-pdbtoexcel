from __future__ import annotations

import configparser
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

from pnfl_pdbtoexcel.pdb import PLAY_DATA

PACKAGE_DIR = Path(__file__).resolve().parent

DEFAULT_PLAY_PATH = r"C:\SIERRA\FbPro98\PNFL"

CONFIG_CANDIDATES = [
    Path.cwd() / "convert-pdb.dev.ini",
    Path.cwd() / "convert-pdb.ini",
    Path.cwd() / "config" / "convert-pdb.dev.ini",
    Path.cwd() / "config" / "convert-pdb.ini",
]

type CategoryOrder = Mapping[PLAY_DATA.PLAY_TYPE, list[str]]


@dataclass(frozen=True)
class Config:
    play_path: str = DEFAULT_PLAY_PATH
    calculate_total_stats: bool = True
    calculate_percentages: bool = True
    include_category_worksheets: bool = False


def get_runtime_path(filename: str) -> Path:
    return PACKAGE_DIR / "resources" / filename


def find_config_path() -> Path:
    return next(
        (c for c in CONFIG_CANDIDATES if c.is_file()),
        CONFIG_CANDIDATES[0],
    )


def load_config(
    path: Path | None = None,
    *,
    play_path: str | None = None,
) -> Config:
    cp = _read_config(path or find_config_path())
    return Config(
        play_path=play_path or cp.get("Settings", "PlayPath", fallback=DEFAULT_PLAY_PATH),
        calculate_total_stats=cp.getboolean("Settings", "CalculateTotalStats", fallback=True),
        calculate_percentages=cp.getboolean("Settings", "CalculatePercentages", fallback=True),
        # Hidden feature: undocumented in the released config and CLI help.
        include_category_worksheets=cp.getboolean("Settings", "IncludeCategoryWorksheets", fallback=False),
    )


def load_category_order(path: Path | None = None) -> CategoryOrder:
    cp = _read_config(path or find_config_path())
    return {
        PLAY_DATA.PLAY_TYPE.RUN: _parse_category_list(cp, "RunCategories"),
        PLAY_DATA.PLAY_TYPE.PASS: _parse_category_list(cp, "PassCategories"),
        PLAY_DATA.PLAY_TYPE.DEFENSE: _parse_category_list(cp, "DefenseCategories"),
    }


def _read_config(path: Path) -> configparser.ConfigParser:
    cp = configparser.ConfigParser()
    cp.read(path, encoding="utf-8")
    return cp


def _parse_category_list(cp: configparser.ConfigParser, key: str) -> list[str]:
    raw = cp.get("CategoryOrder", key, fallback="")
    return [line.strip() for line in raw.splitlines() if line.strip()]
