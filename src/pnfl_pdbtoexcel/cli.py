from __future__ import annotations

import argparse
import logging
from collections.abc import Sequence
from pathlib import Path

from xlsxwriter.exceptions import XlsxWriterException

from pnfl_pdbtoexcel.main import convert_pdb

PROG = "pnfl convert-pdb"
logger = logging.getLogger(__name__)


def _valid_file_extension(param: str, expected_extensions: tuple[str, ...]) -> str:
    filepath = Path(param).expanduser()
    if filepath.suffix.lower() not in expected_extensions:
        extensions = ", ".join(expected_extensions)
        raise argparse.ArgumentTypeError(f"File must have one of these extensions: {extensions}")
    return str(filepath)


def _valid_output_file(param: str) -> str:
    filepath = Path(param).expanduser()
    if filepath.suffix.lower() not in (".xlsm", ".xlsx"):
        raise argparse.ArgumentTypeError("File must have a xlsm or xlsx extension")
    return str(filepath)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog=PROG,
        description="Create an Excel workbook from a WinLogStats PDB and optional FBPro 98 game plans.",
    )
    parser.add_argument(
        "pdbfile",
        type=lambda v: _valid_file_extension(v, (".pdb",)),
        help="WinLogStats database file (.PDB)",
    )
    parser.add_argument(
        "outputfile",
        type=_valid_output_file,
        help="save to this XLSX/XLSM file",
    )
    parser.add_argument(
        "-d",
        "--pln-def",
        type=lambda v: _valid_file_extension(v, (".pln",)),
        help="defensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-d2",
        "--pln-def-2",
        type=lambda v: _valid_file_extension(v, (".pln",)),
        help="second defensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-o",
        "--pln-off",
        type=lambda v: _valid_file_extension(v, (".pln",)),
        help="offensive game plan file (.PLN)",
    )
    parser.add_argument(
        "-o2",
        "--pln-off-2",
        type=lambda v: _valid_file_extension(v, (".pln",)),
        help="second offensive game plan file (.PLN)",
    )
    parser.add_argument(
        "--config",
        type=lambda v: Path(_valid_file_extension(v, (".ini",))),
        help="use this INI file instead of the default config lookup",
    )
    parser.add_argument(
        "--skip-calcs",
        action="store_true",
        help="prevents the extra calculation columns (overrides config settings)",
    )
    parser.add_argument(
        "--skip-totals",
        action="store_true",
        help="prevents totalling stats (overrides config settings)",
    )
    parser.add_argument(
        "--play-path",
        dest="play_path",
        help="path to PNFL play files directory (overrides config [Settings] PlayPath)",
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
    for path in (
        args.pdbfile,
        args.pln_def,
        args.pln_def_2,
        args.pln_off,
        args.pln_off_2,
        args.config,
    ):
        if path is not None and not Path(path).is_file():
            logger.error("%s: %s: file not found", PROG, path)
            return 1
    try:
        convert_pdb(
            pdb_path=args.pdbfile,
            output_path=args.outputfile,
            pln_defense=args.pln_def,
            pln_offense=args.pln_off,
            pln_defense_2=args.pln_def_2,
            pln_offense_2=args.pln_off_2,
            config_path=args.config,
            play_path_override=args.play_path,
            skip_calcs=args.skip_calcs,
            skip_totals=args.skip_totals,
        )
    except (OSError, XlsxWriterException) as error:
        logger.error("%s: %s", PROG, error)
        return 1
    return 0
