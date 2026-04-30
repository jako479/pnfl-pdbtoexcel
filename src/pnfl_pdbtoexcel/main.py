from __future__ import annotations

from pathlib import Path

from pnfl_pdbtoexcel.config import load_category_order, load_config
from pnfl_pdbtoexcel.workbook_creator import PdbWorkbookCreator


def convert_pdb(
    *,
    pdb_path: str,
    output_path: str,
    pln_defense: str | None = None,
    pln_offense: str | None = None,
    pln_defense_2: str | None = None,
    pln_offense_2: str | None = None,
    config_path: Path | None = None,
    play_path_override: str | None = None,
    skip_calcs: bool = False,
    skip_totals: bool = False,
) -> None:
    """Build an Excel workbook from a PDB and optional gameplan files."""
    config = load_config(path=config_path, play_path=play_path_override)
    category_order = load_category_order(path=config_path)
    calculate_totals = config.calculate_total_stats and not skip_totals

    creator = PdbWorkbookCreator.from_config(
        config,
        category_order,
        pdb_path,
        pln_defense,
        pln_offense,
        pln_defense_2,
        pln_offense_2,
    )
    creator.create_workbook(output_path, not skip_calcs, calculate_totals, False)
