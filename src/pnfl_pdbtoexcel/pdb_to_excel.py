from __future__ import annotations

import logging
from os import PathLike

from fbpro98_gameplan import Gameplan, read_gameplan
from pnfl_playpool import PlayPool, PlayRecord

from .config import get_config
from .pdb import PDB, PLAY_DATA
from .workbook import ExcelPdbWorkbook


logger = logging.getLogger(__name__)

StrPath = str | PathLike[str]

DELETED_PLAYS = [
    "ATF0ELOB",
]

TOTAL_STATS_FILTER = {
    "GLR": 2.7,
    "RL": 3.6,
    "RM": 4.6,
    "RR": 4.6,
    "GLP": (0.7, None),
    "PSL": (0.60, 6.0),
    "PSM": (0.60, 5.5),
    "PSR": (0.60, 6.0),
    "PML": (0.50, 6.8),
    "PMM": (0.50, 6.8),
    "PMR": (0.50, 6.0),
    "PLR": (0.40, 6.0),
    "PRD": (0.29, None),
}

DEFENSE_OUTPUT_CATEGORIES = {
    "RunRight",
    "RunMiddle",
    "RunLeft",
    "PassShort",
    "PassMedium",
    "PassLong",
    "RunDazzle",
    "PassDazzle",
    "GLrun",
    "GLpass",
}


class PdbWorkbookCreator:
    def __init__(
        self,
        pdb_filename: StrPath,
        pln_def_filename: StrPath | None,
        pln_off_filename: StrPath | None,
    ) -> None:
        config = get_config()
        self.play_pool = PlayPool.from_directory(config["Settings"]["PnflPath"])

        self.pln_offense: Gameplan | None = None
        self.pln_defense: Gameplan | None = None

        self.pdb = PDB(pdb_filename)
        self.pdb.convert_invalid_play_data(self.play_pool)

        if pln_def_filename:
            self.pln_defense = read_gameplan(pln_def_filename)
        if pln_off_filename:
            self.pln_offense = read_gameplan(pln_off_filename)

    def create_workbook(
        self,
        filename: StrPath,
        perform_calculations: bool,
        calculate_totals: bool,
        filter_total_stats: bool,
    ) -> None:
        logger.info("Attempting to create '%s'", filename)
        if not perform_calculations:
            logger.info("Skipping extra calculations")

        with ExcelPdbWorkbook(filename, perform_calculations) as workbook:
            config = get_config()
            combined_plays: dict[bytes, PLAY_DATA] | None = {} if calculate_totals else None
            missing_play_files_logged: set[str] = set()

            for play_in_pdb, play_name, _, play_record in self._iter_category_source_plays(missing_play_files_logged):
                play_slot = self._get_play_slot(play_in_pdb, play_name)
                workbook.add_play(play_in_pdb, play_slot, play_record)
                if combined_plays is not None:
                    self._add_play_stats_to_total_play(combined_plays, play_in_pdb)

            if combined_plays is not None:
                self._add_total_plays_to_workbook(workbook, combined_plays, filter_total_stats)

            for tendency_data in self.pdb.tendencies:
                workbook.add_tendency(tendency_data)

            if config["Settings"]["CalculateCategoryStats"]:
                team_categories_data, categories_data = self._collect_category_stats(
                    calculate_totals=calculate_totals,
                    group_categories=bool(config["Settings"]["CalculateGroupedCategoryStats"]),
                )

                for team_category, category_data in team_categories_data.items():
                    workbook.add_category(team_category, category_data)

                if calculate_totals:
                    for category_name, category_data in categories_data.items():
                        workbook.add_category(
                            ("`Total Stats", category_name), category_data,
                        )

        logger.info("Conversion complete")

    def _iter_tracked_plays(self):
        play_types = (
            PLAY_DATA.PLAY_TYPE.RUN,
            PLAY_DATA.PLAY_TYPE.PASS,
            PLAY_DATA.PLAY_TYPE.DEFENSE,
        )
        for play_type in play_types:
            plays_list = sorted(
                self.pdb.plays[play_type].values(), key=lambda x: x.team_name,
            )
            yield from plays_list

    def _iter_category_source_plays(self, missing_play_files_logged: set[str] | None = None):
        for play_in_pdb in self._iter_tracked_plays():
            play_name = play_in_pdb.play_name.decode("ASCII")
            team_name = play_in_pdb.team_name.decode("ASCII")
            play_record = self.play_pool.find_by_name(play_name)
            if play_record is None:
                self._log_unknown_play(play_name, missing_play_files_logged)
                continue
            if not self._should_export_play(play_in_pdb, play_name, play_record):
                continue
            yield play_in_pdb, play_name, team_name, play_record

    def _log_unknown_play(
        self, play_name: str, missing_play_files_logged: set[str] | None = None,
    ) -> None:
        if missing_play_files_logged is not None and play_name not in missing_play_files_logged:
            if play_name in DELETED_PLAYS:
                logger.info("Skipping deleted play '%s'", play_name)
            else:
                logger.warning("Play file not found for play '%s'", play_name)
            missing_play_files_logged.add(play_name)

    def _should_export_play(
        self, play_in_pdb: PLAY_DATA, play_name: str, play_record: PlayRecord,
    ) -> bool:
        if (
            play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE
            and play_record.pool_category not in DEFENSE_OUTPUT_CATEGORIES
        ):
            return False
        return play_name not in DELETED_PLAYS

    def _get_play_slot(self, play_in_pdb: PLAY_DATA, play_name: str) -> str:
        play_in_plan = None
        if self.pln_offense and play_in_pdb.play_type in (
            PLAY_DATA.PLAY_TYPE.RUN, PLAY_DATA.PLAY_TYPE.PASS,
        ):
            play_in_plan = self.pln_offense.normal_plays.get(play_name)
        elif self.pln_defense and play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            play_in_plan = self.pln_defense.normal_plays.get(play_name)
        if play_in_plan:
            row = play_in_plan.slot // 4 + 1
            col = play_in_plan.slot % 4 + 1
            return f"{row}-{col}"
        return ""

    def _collect_category_stats(self, calculate_totals: bool, group_categories: bool):
        categories_data: dict[str, PLAY_DATA] = {}
        team_categories_data: dict[tuple, PLAY_DATA] = {}

        for play_in_pdb, _, team_name, play_record in self._iter_category_source_plays():
            self._add_play_stats_to_team_categories(
                team_categories_data=team_categories_data,
                play_in_pdb=play_in_pdb,
                team_name=team_name,
                category=play_record.pool_category,
                group_categories=group_categories,
            )
            if calculate_totals:
                self._add_play_stats_to_category(
                    categories_data, play_in_pdb, play_record.pool_category,
                )

        return team_categories_data, categories_data

    @staticmethod
    def _add_play_stats_to_total_play(
        combined_plays: dict[bytes, PLAY_DATA], play_in_pdb: PLAY_DATA,
    ) -> None:
        if play_in_pdb.play_name in combined_plays:
            combined_play_data = combined_plays[play_in_pdb.play_name]
        else:
            combined_play_data = PLAY_DATA()
            combined_play_data.play_type = play_in_pdb.play_type
            combined_play_data.team_name = b"`Total Stats"
            combined_play_data.play_name = play_in_pdb.play_name
        combined_play_data += play_in_pdb
        combined_plays[play_in_pdb.play_name] = combined_play_data

    def _add_total_plays_to_workbook(
        self,
        workbook: ExcelPdbWorkbook,
        combined_plays: dict[bytes, PLAY_DATA],
        filter_total_stats: bool,
    ) -> None:
        for play_in_pdb in combined_plays.values():
            play_name = play_in_pdb.play_name.decode("ASCII")
            play_record = self.play_pool.find_by_name(play_name)
            if play_record is None:
                continue
            if not self._should_export_play(play_in_pdb, play_name, play_record):
                continue
            if filter_total_stats and not self._play_meets_criteria(play_in_pdb, play_record):
                continue
            play_slot = self._get_play_slot(play_in_pdb, play_name)
            workbook.add_play(play_in_pdb, play_slot, play_record)

    @staticmethod
    def _add_play_stats_to_team_categories(
        team_categories_data: dict[tuple, PLAY_DATA],
        play_in_pdb: PLAY_DATA,
        team_name: str,
        category: str,
        group_categories: bool = False,
    ) -> None:
        team_category = (team_name, category)
        PdbWorkbookCreator._add_play_stats_to_team_category(
            team_categories_data, play_in_pdb, team_category,
        )

        if group_categories:
            if play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.RUN:
                PdbWorkbookCreator._add_play_stats_to_team_category(
                    team_categories_data, play_in_pdb, (team_name, "TOTAL RUNS"),
                )
            elif play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.PASS:
                if category in ("PSL", "PSM", "PSR"):
                    PdbWorkbookCreator._add_play_stats_to_team_category(
                        team_categories_data, play_in_pdb, (team_name, "TOTAL PS"),
                    )
                if category in ("PML", "PMM", "PMR"):
                    PdbWorkbookCreator._add_play_stats_to_team_category(
                        team_categories_data, play_in_pdb, (team_name, "TOTAL PM"),
                    )
                PdbWorkbookCreator._add_play_stats_to_team_category(
                    team_categories_data, play_in_pdb, (team_name, "TOTAL PASSES"),
                )

    @staticmethod
    def _add_play_stats_to_team_category(
        team_categories_data: dict[tuple, PLAY_DATA],
        play_in_pdb: PLAY_DATA,
        team_category: tuple,
    ) -> None:
        if team_category in team_categories_data:
            team_category_data = team_categories_data[team_category]
        else:
            team_category_data = PLAY_DATA()
            team_category_data.play_type = play_in_pdb.play_type
        team_category_data += play_in_pdb
        team_categories_data[team_category] = team_category_data

    @staticmethod
    def _add_play_stats_to_category(
        categories_data: dict[str, PLAY_DATA],
        play_in_pdb: PLAY_DATA,
        category: str,
    ) -> None:
        if category in categories_data:
            category_data = categories_data[category]
        else:
            category_data = PLAY_DATA()
            category_data.play_type = play_in_pdb.play_type
        category_data += play_in_pdb
        categories_data[category] = category_data

    @staticmethod
    def _play_meets_criteria(play_in_pdb: PLAY_DATA, play_record: PlayRecord) -> bool:
        if play_record.pool_category not in TOTAL_STATS_FILTER:
            return True

        if play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            if play_in_pdb.play_count < 7:
                return False
            avg = int(play_in_pdb.total_yards) / int(play_in_pdb.play_count) if play_in_pdb.play_count > 0 else 0.0
            return avg >= TOTAL_STATS_FILTER[play_record.pool_category]
        elif play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            if play_in_pdb.play_count < 7:
                return False
            comp_pct = int(play_in_pdb.completions) / int(play_in_pdb.play_count)
            ypa = int(play_in_pdb.total_yards) / int(play_in_pdb.play_count)
            min_comp_pct = TOTAL_STATS_FILTER[play_record.pool_category][0]
            min_ypa = TOTAL_STATS_FILTER[play_record.pool_category][1]
            return comp_pct >= min_comp_pct or (min_ypa is not None and ypa >= min_ypa)
        return True
