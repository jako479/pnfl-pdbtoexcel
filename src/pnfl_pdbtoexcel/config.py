from __future__ import annotations

import configparser
import hashlib
import socket
from dataclasses import dataclass
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent.parent

DEFAULT_PLAY_PATH = r"C:\SIERRA\FbPro98\PNFL"

CONFIG_CANDIDATES = [
    Path.cwd() / "pdb_to_excel.ini",
    SCRIPT_DIR / "pdb_to_excel.ini",
    PROJECT_DIR / "config" / "pdb_to_excel.ini",
]

_config_path: Path | None = None
_config: AppConfig | None = None
_team_override: str | None = None
_play_path_override: str | None = None


@dataclass
class Settings:
    Team: str = ""
    PlayPath: str = DEFAULT_PLAY_PATH
    CalculateTotalStats: bool = True
    CalculateCategoryStats: bool = False
    CalculateGroupedCategoryStats: bool = False


@dataclass
class AdditionalColumns:
    RunFumblePercentage: bool = True
    RunTouchdownPercentage: bool = True
    PassInterceptionPercentage: bool = True
    PassSackPercentage: bool = True
    PassTouchdownPercentage: bool = True
    DefenseTurnoverPercentage: bool = True
    DefenseSackPercentage: bool = True
    DefenseOffTdPercentage: bool = True


@dataclass
class AppConfig:
    Settings: Settings
    AdditionalColumns: AdditionalColumns


def get_runtime_path(filename: str) -> Path:
    return SCRIPT_DIR / filename


def get_config_path() -> Path:
    global _config_path
    if _config_path is None:
        _config_path = next(
            (c for c in CONFIG_CANDIDATES if c.is_file()),
            CONFIG_CANDIDATES[0],
        )
    return _config_path


def set_config_path(config_path: str | Path) -> None:
    global _config_path, _config
    _config_path = Path(config_path).expanduser().resolve()
    _config = None


def set_team(team: str | None) -> None:
    global _team_override, _config
    _team_override = team
    _config = None


def set_play_path(play_path: str | None) -> None:
    global _play_path_override, _config
    _play_path_override = play_path
    _config = None


def get_config() -> AppConfig:
    global _config
    if _config is not None:
        return _config

    md5 = hashlib.md5(socket.gethostname().encode())
    cp = configparser.ConfigParser()
    cp.read(get_config_path(), encoding="utf-8")

    is_dev_machine = md5.hexdigest() == "5c4b925bf527c4f8581815a35a10d658"

    settings = Settings(
        Team=cp.get("Settings", "Team", fallback=""),
        PlayPath=cp.get("Settings", "PlayPath", fallback=DEFAULT_PLAY_PATH),
        CalculateTotalStats=cp.getboolean("Settings", "CalculateTotalStats", fallback=True),
        CalculateCategoryStats=is_dev_machine,
        CalculateGroupedCategoryStats=is_dev_machine,
    )

    additional_columns = AdditionalColumns(
        RunFumblePercentage=cp.getboolean(
            "AdditionalColumns", "RunFumblePercentage", fallback=True
        ),
        RunTouchdownPercentage=cp.getboolean(
            "AdditionalColumns", "RunTouchdownPercentage", fallback=True
        ),
        PassInterceptionPercentage=cp.getboolean(
            "AdditionalColumns", "PassInterceptionPercentage", fallback=True
        ),
        PassSackPercentage=cp.getboolean("AdditionalColumns", "PassSackPercentage", fallback=True),
        PassTouchdownPercentage=cp.getboolean(
            "AdditionalColumns", "PassTouchdownPercentage", fallback=True
        ),
        DefenseTurnoverPercentage=cp.getboolean(
            "AdditionalColumns", "DefenseTurnoverPercentage", fallback=True
        ),
        DefenseSackPercentage=cp.getboolean(
            "AdditionalColumns", "DefenseSackPercentage", fallback=True
        ),
        DefenseOffTdPercentage=cp.getboolean(
            "AdditionalColumns", "DefenseOffTdPercentage", fallback=True
        ),
    )

    if _team_override is not None:
        settings.Team = _team_override
    if _play_path_override is not None:
        settings.PlayPath = _play_path_override

    _config = AppConfig(Settings=settings, AdditionalColumns=additional_columns)
    return _config
