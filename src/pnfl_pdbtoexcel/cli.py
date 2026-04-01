from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence
from pathlib import Path

from .config import set_config_path, set_pnfl_path, set_team, get_config
from .pdb_to_excel import PdbWorkbookCreator


def _valid_existing_file(param: str, expected_extensions: tuple[str, ...]) -> str:
    filepath = Path(param).expanduser()
    if filepath.suffix.lower() not in expected_extensions:
        extensions = ", ".join(expected_extensions)
        raise argparse.ArgumentTypeError(
            f"File must have one of these extensions: {extensions}"
        )
    if not filepath.is_file():
        raise argparse.ArgumentTypeError(f"File not found: {filepath}")
    return str(filepath)


def _valid_output_file(param: str) -> str:
    filepath = Path(param).expanduser()
    if filepath.suffix.lower() not in (".xlsm", ".xlsx"):
        raise argparse.ArgumentTypeError("File must have a xlsm or xlsx extension")
    return str(filepath)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pnfl-pdbtoexcel",
        description="Create an Excel workbook from a WinLogStats PDB and optional FBPro 98 game plans.",
    )
    parser.add_argument(
        "pdbfile",
        type=lambda v: _valid_existing_file(v, (".pdb",)),
        help="WinLogStats database file (.PDB)",
    )
    parser.add_argument(
        "outputfile",
        type=_valid_output_file,
        help="save to this XLSX/XLSM file",
    )
    parser.add_argument(
        "-d", "--plnfile-defense",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="defensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-o", "--plnfile-offense",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="offensive game plan file (.PLN)",
    )
    parser.add_argument(
        "--config",
        type=lambda v: _valid_existing_file(v, (".ini",)),
        help="use this INI file instead of the default config lookup",
    )
    parser.add_argument(
        "-c", "--skip-calcs",
        action="store_true",
        help="prevents the extra calculation columns (overrides config settings)",
    )
    parser.add_argument(
        "-t", "--skip-totals",
        action="store_true",
        help="prevents totalling stats (overrides config settings)",
    )
    parser.add_argument("--team", help="team name (overrides config Settings.Team)")
    parser.add_argument(
        "--pnfl-path",
        help="path to PNFL play tree (overrides config Settings.PnflPath)",
    )
    return parser


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = build_parser()
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.INFO,
        format="%(levelname)s: %(message)s",
    )

    if args.config:
        set_config_path(args.config)
    set_team(args.team)
    set_pnfl_path(args.pnfl_path)

    config = get_config()
    calculate_totals = config["Settings"]["CalculateTotalStats"] and not args.skip_totals

    creator = PdbWorkbookCreator(args.pdbfile, args.plnfile_defense, args.plnfile_offense)
    creator.create_workbook(args.outputfile, not args.skip_calcs, calculate_totals, False)
    return 0
