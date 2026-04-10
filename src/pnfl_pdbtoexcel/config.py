from __future__ import annotations

import configparser
import hashlib
import socket
from dataclasses import dataclass
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PACKAGE_DIR.parent.parent

DEFAULT_PLAY_PATH = r"C:\SIERRA\FbPro98\PNFL"

CONFIG_CANDIDATES = [
    Path.cwd() / "convert-pdb.dev.ini",
    Path.cwd() / "convert-pdb.ini",
    PROJECT_DIR / "config" / "convert-pdb.dev.ini",
    PROJECT_DIR / "config" / "convert-pdb.ini",
    PACKAGE_DIR / "convert-pdb.dev.ini",
    PACKAGE_DIR / "convert-pdb.ini",
]


@dataclass(frozen=True)
class Settings:
    PlayPath: str = DEFAULT_PLAY_PATH
    CalculateTotalStats: bool = True
    CalculatePercentages: bool = True
    CalculateCategoryStats: bool = False


@dataclass(frozen=True)
class CategoryOrder:
    RunCategories: list[str]
    PassCategories: list[str]
    DefenseCategories: list[str]


@dataclass(frozen=True)
class AppConfig:
    Settings: Settings
    CategoryOrder: CategoryOrder


def get_runtime_path(filename: str) -> Path:
    return PACKAGE_DIR / "resources" / filename


def find_config_path() -> Path:
    return next(
        (c for c in CONFIG_CANDIDATES if c.is_file()),
        CONFIG_CANDIDATES[0],
    )


def _parse_category_list(cp: configparser.ConfigParser, key: str) -> list[str]:
    raw = cp.get("CategoryOrder", key, fallback="")
    return [line.strip() for line in raw.splitlines() if line.strip()]


def load_config(
    config_path: Path | None = None,
    play_path: str | None = None,
) -> AppConfig:
    path = config_path or find_config_path()
    cp = configparser.ConfigParser()
    cp.read(path, encoding="utf-8")

    md5 = hashlib.md5(socket.gethostname().encode())
    is_dev_machine = md5.hexdigest() == "5c4b925bf527c4f8581815a35a10d658"

    return AppConfig(
        Settings=Settings(
            PlayPath=play_path or cp.get("Settings", "PlayPath", fallback=DEFAULT_PLAY_PATH),
            CalculateTotalStats=cp.getboolean("Settings", "CalculateTotalStats", fallback=True),
            CalculatePercentages=cp.getboolean("Settings", "CalculatePercentages", fallback=True),
            CalculateCategoryStats=is_dev_machine,
        ),
        CategoryOrder=CategoryOrder(
            RunCategories=_parse_category_list(cp, "RunCategories"),
            PassCategories=_parse_category_list(cp, "PassCategories"),
            DefenseCategories=_parse_category_list(cp, "DefenseCategories"),
        ),
    )
