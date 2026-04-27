from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence
from pathlib import Path

from pnfl_pdbtoexcel.config import load_config
from pnfl_pdbtoexcel.pdb_to_excel import PdbWorkbookCreator


def _valid_existing_file(param: str, expected_extensions: tuple[str, ...]) -> str:
    filepath = Path(param).expanduser()
    if filepath.suffix.lower() not in expected_extensions:
        extensions = ", ".join(expected_extensions)
        raise argparse.ArgumentTypeError(f"File must have one of these extensions: {extensions}")
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
        prog="pnfl convert-pdb",
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
        "-d",
        "--plnfile-defense",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="defensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-d2",
        "--plnfile-defense-2",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="second defensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-o",
        "--plnfile-offense",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="offensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-o2",
        "--plnfile-offense-2",
        type=lambda v: _valid_existing_file(v, (".pln",)),
        help="second offensive game plan file (.PLN)",
    )
    parser.add_argument(
        "--config",
        type=lambda v: _valid_existing_file(v, (".ini",)),
        help="use this INI file instead of the default config lookup",
    )
    parser.add_argument(
        "-c",
        "--skip-calcs",
        action="store_true",
        help="prevents the extra calculation columns (overrides config settings)",
    )
    parser.add_argument(
        "-t",
        "--skip-totals",
        action="store_true",
        help="prevents totalling stats (overrides config settings)",
    )
    parser.add_argument(
        "--play-path",
        help="path to PNFL play files directory (overrides config Settings.PlayPath)",
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

    config = load_config(
        config_path=args.config,
        play_path=args.play_path,
    )
    calculate_totals = config.Settings.CalculateTotalStats and not args.skip_totals

    creator = PdbWorkbookCreator.from_files(
        config,
        args.pdbfile,
        args.plnfile_defense,
        args.plnfile_offense,
        args.plnfile_defense_2,
        args.plnfile_offense_2,
    )
    creator.create_workbook(args.outputfile, not args.skip_calcs, calculate_totals, False)
    return 0
