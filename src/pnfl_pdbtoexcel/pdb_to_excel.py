from __future__ import annotations

import logging
from collections.abc import Iterator
from os import PathLike

from fbpro98_gameplan import GamePlan, read_gameplan
from pnfl_playpool import DefensivePlayRecord, OffensivePlayRecord, PlayPool, PlayRecord, SpecialTeamsPlayRecord

from pnfl_pdbtoexcel.config import AppConfig
from pnfl_pdbtoexcel.pdb import PDB, PLAY_DATA
from pnfl_pdbtoexcel.workbook import ExcelPdbWorkbook

logger = logging.getLogger(__name__)

StrPath = str | PathLike[str]
NormalPlayRecord = OffensivePlayRecord | DefensivePlayRecord

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
        config: AppConfig,
        play_pool: PlayPool,
        pdb: PDB,
        pln_defense: GamePlan | None = None,
        pln_offense: GamePlan | None = None,
        pln_defense_2: GamePlan | None = None,
        pln_offense_2: GamePlan | None = None,
    ) -> None:
        self.config = config
        self.play_pool = play_pool
        self.pdb = pdb
        self.pln_defense = pln_defense
        self.pln_offense = pln_offense
        self.pln_defense_2 = pln_defense_2
        self.pln_offense_2 = pln_offense_2

    @classmethod
    def from_files(
        cls,
        config: AppConfig,
        pdb_filename: StrPath,
        pln_def_filename: StrPath | None = None,
        pln_off_filename: StrPath | None = None,
        pln_def_filename_2: StrPath | None = None,
        pln_off_filename_2: StrPath | None = None,
    ) -> PdbWorkbookCreator:
        """Convenience factory — builds all dependencies from file paths.

        This is not dependency injection itself; it's the factory that does
        the building so the constructor doesn't have to. Production code
        calls this; test code calls __init__ directly with fakes.
        """
        play_pool = PlayPool.from_directory(config.Settings.PlayPath)
        pdb = PDB(pdb_filename)
        pdb.convert_invalid_play_data(play_pool)
        pln_defense = read_gameplan(pln_def_filename) if pln_def_filename else None
        pln_offense = read_gameplan(pln_off_filename) if pln_off_filename else None
        pln_defense_2 = read_gameplan(pln_def_filename_2) if pln_def_filename_2 else None
        pln_offense_2 = read_gameplan(pln_off_filename_2) if pln_off_filename_2 else None
        return cls(config, play_pool, pdb, pln_defense, pln_offense, pln_defense_2, pln_offense_2)

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

        offense_slot_count = (1 if self.pln_offense else 0) + (1 if self.pln_offense_2 else 0)
        defense_slot_count = (1 if self.pln_defense else 0) + (1 if self.pln_defense_2 else 0)

        with ExcelPdbWorkbook(
            self.config,
            filename,
            perform_calculations,
            offense_slot_count,
            defense_slot_count,
        ) as workbook:
            combined_plays: dict[bytes, PLAY_DATA] | None = {} if calculate_totals else None
            missing_play_files_logged: set[str] = set()

            resolved_plays: list[tuple[PLAY_DATA, str, str, NormalPlayRecord]] = []
            for play_in_pdb, play_name, team_name, play_record in self._iter_category_source_plays(
                missing_play_files_logged
            ):
                resolved_plays.append((play_in_pdb, play_name, team_name, play_record))
                play_slots = self._get_play_slots(play_in_pdb, play_name)
                workbook.add_play(play_in_pdb, play_slots, play_record)
                if combined_plays is not None:
                    self._add_play_stats_to_total_play(combined_plays, play_in_pdb)

            if combined_plays is not None:
                self._add_total_plays_to_workbook(workbook, combined_plays, filter_total_stats)

            for tendency_data in self.pdb.tendencies:
                workbook.add_tendency(tendency_data)

            if self.config.Settings.CalculateCategoryStats:
                team_categories_data, categories_data = self._collect_category_stats(
                    resolved_plays,
                    calculate_totals=calculate_totals,
                )

                for team_category, category_data in team_categories_data.items():
                    workbook.add_category(team_category, category_data)

                if calculate_totals:
                    for category_name, category_data in categories_data.items():
                        workbook.add_category(
                            ("Total Stats", category_name),
                            category_data,
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
                self.pdb.plays[play_type].values(),
                key=lambda x: x.team_name,
            )
            yield from plays_list

    def _iter_category_source_plays(
        self,
        missing_play_files_logged: set[str] | None = None,
    ) -> Iterator[tuple[PLAY_DATA, str, str, NormalPlayRecord]]:
        for play_in_pdb in self._iter_tracked_plays():
            play_name = play_in_pdb.play_name.decode("ASCII")
            team_name = play_in_pdb.team_name.decode("ASCII")
            play_record = self.play_pool.find_by_name(play_name)
            if play_record is None:
                self._log_unknown_play(play_name, missing_play_files_logged)
                continue
            if not self._should_export_play(play_in_pdb, play_name, play_record):
                continue
            # Special teams filtered by _should_export_play; only normal plays remain.
            assert isinstance(play_record, (OffensivePlayRecord, DefensivePlayRecord))
            yield play_in_pdb, play_name, team_name, play_record

    def _log_unknown_play(
        self,
        play_name: str,
        missing_play_files_logged: set[str] | None = None,
    ) -> None:
        if missing_play_files_logged is not None and play_name not in missing_play_files_logged:
            if play_name in DELETED_PLAYS:
                logger.info("Skipping deleted play '%s'", play_name)
            else:
                logger.warning("Play file not found for play '%s'", play_name)
            missing_play_files_logged.add(play_name)

    def _should_export_play(
        self,
        play_in_pdb: PLAY_DATA,
        play_name: str,
        play_record: PlayRecord,
    ) -> bool:
        if play_name in DELETED_PLAYS:
            return False
        # The PDB sometimes classifies special teams plays (e.g. SFFGPass) as
        # RUN or PASS type. Skip them — they don't belong in the play worksheets.
        if isinstance(play_record, SpecialTeamsPlayRecord):
            return False
        if (
            play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE
            and isinstance(play_record, DefensivePlayRecord)
            and play_record.pool_category not in DEFENSE_OUTPUT_CATEGORIES
        ):
            return False
        return True

    def _get_play_slots(self, play_in_pdb: PLAY_DATA, play_name: str) -> tuple[str, str]:
        if play_in_pdb.play_type in (PLAY_DATA.PLAY_TYPE.RUN, PLAY_DATA.PLAY_TYPE.PASS):
            plan_1 = self.pln_offense
            plan_2 = self.pln_offense_2
        elif play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            plan_1 = self.pln_defense
            plan_2 = self.pln_defense_2
        else:
            return ("", "")
        return (
            self._format_slot(plan_1.normal_plays.get(play_name)) if plan_1 else "",
            self._format_slot(plan_2.normal_plays.get(play_name)) if plan_2 else "",
        )

    @staticmethod
    def _format_slot(play_in_plan) -> str:
        if not play_in_plan:
            return ""
        row = play_in_plan.slot // 4 + 1
        col = play_in_plan.slot % 4 + 1
        return f"{row}-{col}"

    def _collect_category_stats(
        self,
        resolved_plays: list[tuple[PLAY_DATA, str, str, NormalPlayRecord]],
        calculate_totals: bool,
    ):
        categories_data: dict[str, PLAY_DATA] = {}
        team_categories_data: dict[tuple[str, str], PLAY_DATA] = {}

        for play_in_pdb, _, team_name, play_record in resolved_plays:
            self._add_play_stats_to_team_category(
                team_categories_data,
                play_in_pdb,
                (team_name, play_record.pool_category),
            )
            if calculate_totals:
                self._add_play_stats_to_category(
                    categories_data,
                    play_in_pdb,
                    play_record.pool_category,
                )

        return team_categories_data, categories_data

    @staticmethod
    def _add_play_stats_to_total_play(
        combined_plays: dict[bytes, PLAY_DATA],
        play_in_pdb: PLAY_DATA,
    ) -> None:
        if play_in_pdb.play_name in combined_plays:
            combined_play_data = combined_plays[play_in_pdb.play_name]
        else:
            combined_play_data = PLAY_DATA()
            combined_play_data.play_type = play_in_pdb.play_type
            combined_play_data.team_name = b"Total Stats"
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
            assert isinstance(play_record, (OffensivePlayRecord, DefensivePlayRecord))
            if filter_total_stats and not self._play_meets_criteria(play_in_pdb, play_record):
                continue
            play_slots = self._get_play_slots(play_in_pdb, play_name)
            workbook.add_play(play_in_pdb, play_slots, play_record)

    @staticmethod
    def _add_play_stats_to_team_category(
        team_categories_data: dict[tuple[str, str], PLAY_DATA],
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
    def _play_meets_criteria(play_in_pdb: PLAY_DATA, play_record: NormalPlayRecord) -> bool:
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
