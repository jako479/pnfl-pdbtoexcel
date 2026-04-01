from __future__ import annotations

import configparser
import hashlib
import socket
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent.parent

DEFAULT_PNFL_PATH = r"C:\SIERRA\FbPro98\PNFL"

CONFIG_CANDIDATES = [
    SCRIPT_DIR / "pdb_to_excel.ini",
    PROJECT_DIR / "config" / "pdb_to_excel.ini",
    SCRIPT_DIR / "PdbToExcel.ini",
    PROJECT_DIR / "config" / "PdbToExcel.ini",
]

_config_path: Path | None = None
_config: dict[str, dict[str, object]] | None = None
_team_override: str | None = None
_pnfl_path_override: str | None = None


def get_runtime_path(filename: str) -> Path:
    return SCRIPT_DIR / filename


def set_config_path(config_path: str | Path) -> None:
    global _config_path, _config
    _config_path = Path(config_path).expanduser().resolve()
    _config = None


def set_team(team: str | None) -> None:
    global _team_override, _config
    _team_override = team
    _config = None


def set_pnfl_path(pnfl_path: str | None) -> None:
    global _pnfl_path_override, _config
    _pnfl_path_override = pnfl_path
    _config = None


def get_config_path() -> Path:
    global _config_path
    if _config_path is None:
        _config_path = next(
            (c for c in CONFIG_CANDIDATES if c.is_file()),
            CONFIG_CANDIDATES[0],
        )
    return _config_path


def get_config() -> dict[str, dict[str, object]]:
    global _config
    if _config is not None:
        return _config

    md5 = hashlib.md5(socket.gethostname().encode())
    cp = configparser.ConfigParser()
    cp.read(get_config_path(), encoding="utf-8")

    _config = {
        "Settings": {
            "Team": cp.get("Settings", "Team", fallback=""),
            "PnflPath": cp.get("Settings", "PnflPath", fallback=DEFAULT_PNFL_PATH),
            "CalculateTotalStats": cp.getboolean("Settings", "CalculateTotalStats", fallback=True),
            "CalculateCategoryStats": md5.hexdigest() == "5c4b925bf527c4f8581815a35a10d658",
            "CalculateGroupedCategoryStats": (
                md5.hexdigest() == "5c4b925bf527c4f8581815a35a10d658" and 1
            ),
        },
        "AdditionalColumns": {
            "RunFumblePercentage": cp.getboolean("AdditionalColumns", "RunFumblePercentage", fallback=True),
            "RunTouchdownPercentage": cp.getboolean("AdditionalColumns", "RunTouchdownPercentage", fallback=True),
            "PassInterceptionPercentage": cp.getboolean("AdditionalColumns", "PassInterceptionPercentage", fallback=True),
            "PassSackPercentage": cp.getboolean("AdditionalColumns", "PassSackPercentage", fallback=True),
            "PassTouchdownPercentage": cp.getboolean("AdditionalColumns", "PassTouchdownPercentage", fallback=True),
            "DefenseTurnoverPercentage": cp.getboolean("AdditionalColumns", "DefenseTurnoverPercentage", fallback=True),
            "DefenseSackPercentage": cp.getboolean("AdditionalColumns", "DefenseSackPercentage", fallback=True),
            "DefenseOffTdPercentage": cp.getboolean("AdditionalColumns", "DefenseOffTdPercentage", fallback=True),
        },
    }

    if _team_override is not None:
        _config["Settings"]["Team"] = _team_override
    if _pnfl_path_override is not None:
        _config["Settings"]["PnflPath"] = _pnfl_path_override

    return _config
