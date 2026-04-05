from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import xlsxwriter
from xlsxwriter.worksheet import Worksheet

from pnfl_playpool import DefensivePlayRecord, OffensivePlayRecord, PlayRecord

from .config import AppConfig, get_config, get_runtime_path
from .pdb import PLAY_DATA


@dataclass
class _WorksheetState:
    worksheet: Worksheet
    columns: int
    rows: int = 1


@dataclass
class _Formats:
    text: xlsxwriter.format.Format
    avg: xlsxwriter.format.Format
    percent_0: xlsxwriter.format.Format
    percent_1: xlsxwriter.format.Format
    center: xlsxwriter.format.Format
    options_header: xlsxwriter.format.Format
    option_label: xlsxwriter.format.Format
    option: xlsxwriter.format.Format
    option_note: xlsxwriter.format.Format


class ExcelPdbWorkbook:

    def __init__(self, config: AppConfig, filename, perform_calculations):
        self.config = config
        self.filename = filename
        self.perform_calculations = perform_calculations

    @classmethod
    def from_config(cls, filename, perform_calculations) -> ExcelPdbWorkbook:
        """Convenience factory — reads config so the caller doesn't have to."""
        return cls(get_config(), filename, perform_calculations)

    def __enter__(self):
        filepath = Path(self.filename)
        filepath.parent.mkdir(exist_ok=True)

        self.workbook = xlsxwriter.Workbook(filepath)
        if filepath.suffix == ".xlsm":
            if self.config.Settings.CalculateCategoryStats:
                self.workbook.add_vba_project(str(get_runtime_path("vbaProject_categories.bin")))
            else:
                self.workbook.add_vba_project(str(get_runtime_path("vbaProject.bin")))

        self.fmt = self._create_formats()
        self._create_options_worksheet()
        self.run = self._create_run_worksheet()
        self.pass_ = self._create_pass_worksheet()
        self.def_ = self._create_def_worksheet()
        self.tendencies = self._create_tendencies_worksheet()

        self.run_categories: _WorksheetState | None = None
        self.pass_categories: _WorksheetState | None = None
        self.def_categories: _WorksheetState | None = None
        if self.config.Settings.CalculateCategoryStats:
            self.run_categories = self._create_run_categories_worksheet()
            self.pass_categories = self._create_pass_categories_worksheet()
            self.def_categories = self._create_def_categories_worksheet()

        return self

    def __exit__(self, *args):
        if self.workbook:
            self._add_conditional_format()
            self.workbook.close()

    # -- Public API -----------------------------------------------------------

    def add_play(self, play_data, play_slot, play_record):
        if play_data.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            self._add_run_play(play_data, play_slot, play_record)
        elif play_data.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            self._add_pass_play(play_data, play_slot, play_record)
        elif play_data.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            self._add_defense_play(play_data, play_slot, play_record)

    def add_category(self, team_category, category_data):
        if category_data.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            self._add_run_category(team_category, category_data)
        elif category_data.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            self._add_pass_category(team_category, category_data)
        elif category_data.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            self._add_defense_category(team_category, category_data)

    def add_tendency(self, t):
        team = t.team_name.decode("ASCII")
        downs = [
            ("First", "first_down"),
            ("Second", "second_down"),
            ("Third", "third_down"),
            ("Fourth", "fourth_down"),
        ]
        buckets = [
            ("0-1", t.run_zero_to_one, t.pass_zero_to_one),
            ("2-5", t.run_two_to_five, t.pass_two_to_five),
            ("6-10", t.run_six_to_ten, t.pass_six_to_ten),
            ("10+", t.run_ten_plus, t.pass_ten_plus),
        ]
        for down_label, down_attr in downs:
            for bucket_label, run_data, pass_data in buckets:
                row_data = [
                    team,
                    f"{down_label} and {bucket_label}",
                    getattr(run_data, down_attr),
                    getattr(pass_data, down_attr),
                ]
                self.tendencies.worksheet.write_row(self.tendencies.rows, 0, row_data)
                self.tendencies.rows += 1

    # -- Worksheet creation ---------------------------------------------------

    def _create_formats(self) -> _Formats:
        return _Formats(
            text=self.workbook.add_format({"num_format": "@"}),
            avg=self.workbook.add_format({"num_format": "0.0"}),
            percent_0=self.workbook.add_format({"num_format": "0%"}),
            percent_1=self.workbook.add_format({"num_format": "0.0%"}),
            center=self.workbook.add_format({"align": "center"}),
            options_header=self.workbook.add_format(
                {"bg_color": "#FABF8F", "bold": True, "font_size": 10}
            ),
            option_label=self.workbook.add_format(
                {"bg_color": "#FDE9D9", "bold": True, "font_size": 10}
            ),
            option=self.workbook.add_format(
                {"bg_color": "#FFFF00", "bold": True, "font_size": 10, "border": 2, "border_color": "black"}
            ),
            option_note=self.workbook.add_format(
                {"bg_color": "#FDE9D9", "bold": True, "font_size": 8}
            ),
        )

    def _create_options_worksheet(self) -> None:
        ws = self.workbook.add_worksheet("Options")
        ws.set_vba_name("wsOptions")
        ws.set_column_pixels(0, 0, 6)
        ws.write_row(0, 1, ["OPTIONS - MISCELLANEOUS", "", "", "", ""], self.fmt.options_header)
        ws.write_row(1, 1, ["Highlight Selected Row", "", "", "", ""], self.fmt.option_label)
        ws.write(2, 1, "Yes", self.fmt.option)
        ws.write_row(2, 2, ["      (when enabled, hold 'Ctrl' to paste data)", "", "", ""], self.fmt.option_note)

    def _create_run_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Run Plays")
        ws.set_vba_name("wsRunPlays")
        header = ["Team", "Category", "Slot", "Play", "Type", "Attempts", "Yards", "Avg", "Fumbles", "TD", "Pts", "Custom1", "Custom2"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 80)
        ws.set_column_pixels(2, 2, 47, self.fmt.text)
        ws.set_column_pixels(3, 3, 100)
        ws.set_column_pixels(4, 4, 59)
        ws.set_column_pixels(7, 7, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.RunFumblePercentage:
                idx = 9
                header.insert(idx, "Fumble %")
                ws.set_column_pixels(idx, idx, 70, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.RunTouchdownPercentage:
                idx = 10 + extra
                header.insert(idx, "TD %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:C1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_pass_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Pass Plays")
        ws.set_vba_name("wsPassPlays")
        header = ["Team", "Category", "Slot", "Play", "Type", "Comp", "Att", "Comp %", "Yards", "Y/Comp", "Y/Att", "Fumbles", "Int", "Sacks", "TD", "Pts", "Custom1", "Custom2"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 80)
        ws.set_column_pixels(2, 2, 47, self.fmt.text)
        ws.set_column_pixels(3, 3, 100)
        ws.set_column_pixels(4, 4, 49)
        ws.set_column_pixels(5, 5, 49)
        ws.set_column_pixels(6, 6, 49)
        ws.set_column_pixels(7, 7, None, self.fmt.percent_0)
        ws.set_column_pixels(9, 9, None, self.fmt.avg)
        ws.set_column_pixels(10, 10, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.PassInterceptionPercentage:
                idx = 13
                header.insert(idx, "Int %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.PassSackPercentage:
                idx = 14 + extra
                header.insert(idx, "Sack %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.PassTouchdownPercentage:
                idx = 15 + extra
                header.insert(idx, "TD %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:C1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_def_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Def Plays")
        ws.set_vba_name("wsDefPlays")
        header = ["Team", "Category", "Slot", "Play", "Type", "vs Run", "Avg", "vs Pass", "Avg", "Fumbles", "Int", "Sacks", "TD/Def", "TD/Off", "Custom1", "Custom2"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 126)
        ws.set_column_pixels(2, 2, 47, self.fmt.text)
        ws.set_column_pixels(3, 3, 100)
        ws.set_column_pixels(4, 4, 37)
        ws.set_column_pixels(5, 5, None, self.fmt.text)
        ws.set_column_pixels(6, 6, None, self.fmt.avg)
        ws.set_column_pixels(7, 7, None, self.fmt.text)
        ws.set_column_pixels(8, 8, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.DefenseTurnoverPercentage:
                header.insert(11, "TO %")
                ws.set_column_pixels(11, 11, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.DefenseSackPercentage:
                idx = 12 + extra
                header.insert(idx, "Sack %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.DefenseOffTdPercentage:
                idx = 14 + extra
                header.insert(idx, "TD/Off %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:C1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_tendencies_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Tendencies")
        ws.set_vba_name("wsTendencies")
        header = ["Team", "Situation", "Runs", "Passes"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 140)
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:A1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_run_categories_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Run Categories")
        ws.set_vba_name("wsRunCategories")
        header = ["Team", "Category", "Attempts", "Yards", "Avg", "Fumbles", "TD", "Pts"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 80)
        ws.set_column_pixels(4, 4, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.RunFumblePercentage:
                idx = 6
                header.insert(idx, "Fumble %")
                ws.set_column_pixels(idx, idx, 70, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.RunTouchdownPercentage:
                idx = 7 + extra
                header.insert(idx, "TD %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:B1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_pass_categories_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Pass Categories")
        ws.set_vba_name("wsPassCategories")
        header = ["Team", "Category", "Comp", "Att", "Comp %", "Yards", "Y/Comp", "Y/Att", "Fumbles", "Int", "Sacks", "TD", "Pts"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 80)
        ws.set_column_pixels(2, 2, 49)
        ws.set_column_pixels(3, 3, 49)
        ws.set_column_pixels(4, 4, None, self.fmt.percent_0)
        ws.set_column_pixels(6, 6, None, self.fmt.avg)
        ws.set_column_pixels(7, 7, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.PassInterceptionPercentage:
                idx = 10
                header.insert(idx, "Int %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.PassSackPercentage:
                idx = 11 + extra
                header.insert(idx, "Sack %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.PassTouchdownPercentage:
                idx = 12 + extra
                header.insert(idx, "TD %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:B1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    def _create_def_categories_worksheet(self) -> _WorksheetState:
        ws = self.workbook.add_worksheet("Def Categories")
        ws.set_vba_name("wsDefCategories")
        header = ["Team", "Category", "vs Run", "Avg", "vs Pass", "Avg", "Fumbles", "Int", "Sacks", "TD/Def", "TD/Off"]
        ws.set_column_pixels(0, 0, 120)
        ws.set_column_pixels(1, 1, 126)
        ws.set_column_pixels(2, 2, None, self.fmt.text)
        ws.set_column_pixels(3, 3, None, self.fmt.avg)
        ws.set_column_pixels(4, 4, None, self.fmt.text)
        ws.set_column_pixels(5, 5, None, self.fmt.avg)
        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.DefenseTurnoverPercentage:
                idx = 8
                header.insert(idx, "TO %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.DefenseSackPercentage:
                idx = 9 + extra
                header.insert(idx, "Sack %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
            if self.config.AdditionalColumns.DefenseOffTdPercentage:
                idx = 11 + extra
                header.insert(idx, "TD/Off %")
                ws.set_column_pixels(idx, idx, None, self.fmt.percent_1)
                extra += 1
        ws.write_row(0, 0, header)
        ws.freeze_panes(1, 0)
        ws.autofilter("A1:B1")  # pyright: ignore[reportCallIssue]
        return _WorksheetState(ws, len(header))

    # -- Row writers ----------------------------------------------------------

    def _add_run_play(self, play_data, play_slot, play_record):
        play_type = ""
        if isinstance(play_record, OffensivePlayRecord) and play_record.qb_draw:
            play_type = "QB draw"

        avg = int(play_data.total_yards) / int(play_data.play_count) if play_data.play_count > 0 else 0.0

        row_data = [
            play_data.team_name.decode("ASCII"),
            play_record.pool_category,
            play_slot,
            play_data.play_name.decode("ASCII"),
            play_type,
            play_data.play_count,
            play_data.total_yards,
            round(avg, 1),
            play_data.fumbles,
            play_data.touchdowns_offense,
            play_data.touchdowns_offense * 6,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.RunFumblePercentage:
                row_data.insert(9, round(int(play_data.fumbles) / int(play_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.RunTouchdownPercentage:
                row_data.insert(10 + extra, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra += 1

        self.run.worksheet.write_row(self.run.rows, 0, row_data)
        self.run.rows += 1

    def _add_pass_play(self, play_data, play_slot, play_record):
        play_type = ""
        if isinstance(play_record, OffensivePlayRecord) and play_record.screen:
            play_type = "Screen"

        avg_per_completion = int(play_data.total_yards) / int(play_data.completions) if play_data.completions > 0 else 0.0
        avg_per_attempt = int(play_data.total_yards) / int(play_data.play_count) if play_data.play_count > 0 else 0.0

        row_data = [
            play_data.team_name.decode("ASCII"),
            play_record.pool_category,
            play_slot,
            play_data.play_name.decode("ASCII"),
            play_type,
            play_data.completions,
            play_data.play_count,
            round(int(play_data.completions) / int(play_data.play_count), 2),
            play_data.total_yards,
            round(avg_per_completion, 1),
            round(avg_per_attempt, 1),
            play_data.fumbles,
            play_data.interceptions,
            play_data.sacks,
            play_data.touchdowns_offense,
            play_data.touchdowns_offense * 6,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.PassInterceptionPercentage:
                row_data.insert(13, round(int(play_data.interceptions) / int(play_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.PassSackPercentage:
                row_data.insert(14 + extra, round(int(play_data.sacks) / int(play_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.PassTouchdownPercentage:
                row_data.insert(15 + extra, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra += 1

        self.pass_.worksheet.write_row(self.pass_.rows, 0, row_data)
        self.pass_.rows += 1

    def _add_defense_play(self, play_data, play_slot, play_record):
        play_type = ""
        if isinstance(play_record, DefensivePlayRecord) and play_record.personnel_grouping is not None:
            play_type = play_record.personnel_grouping.value

        rush_avg = int(play_data.rush_yards_allowed) / int(play_data.run_plays_against) if play_data.run_plays_against > 0 else 0.0
        pass_avg = int(play_data.pass_yards_allowed) / int(play_data.pass_plays_against) if play_data.pass_plays_against > 0 else 0.0

        row_data = [
            play_data.team_name.decode("ASCII"),
            play_record.pool_category,
            play_slot,
            play_data.play_name.decode("ASCII"),
            play_type,
            f"{play_data.run_plays_against}/{play_data.rush_yards_allowed}",
            round(rush_avg, 1),
            f"{play_data.pass_plays_against}/{play_data.pass_yards_allowed}",
            round(pass_avg, 1),
            play_data.fumbles,
            play_data.interceptions,
            play_data.sacks,
            play_data.touchdowns_defense,
            play_data.touchdowns_offense,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.DefenseTurnoverPercentage:
                row_data.insert(11, round((int(play_data.fumbles) + int(play_data.interceptions)) / int(play_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.DefenseSackPercentage:
                row_data.insert(12 + extra, round(int(play_data.sacks) / int(play_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.DefenseOffTdPercentage:
                row_data.insert(14 + extra, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra += 1

        self.def_.worksheet.write_row(self.def_.rows, 0, row_data)
        self.def_.rows += 1

    def _add_run_category(self, team_category, category_data):
        avg = int(category_data.total_yards) / int(category_data.play_count) if category_data.play_count > 0 else 0.0

        row_data = [
            team_category[0],
            team_category[1],
            category_data.play_count,
            category_data.total_yards,
            round(avg, 1),
            category_data.fumbles,
            category_data.touchdowns_offense,
            category_data.touchdowns_offense * 6,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.RunFumblePercentage:
                row_data.insert(6, round(int(category_data.fumbles) / int(category_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.RunTouchdownPercentage:
                row_data.insert(7 + extra, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra += 1

        self.run_categories.worksheet.write_row(self.run_categories.rows, 0, row_data)
        self.run_categories.rows += 1

    def _add_pass_category(self, team_category, category_data):
        avg_per_completion = int(category_data.total_yards) / int(category_data.completions) if category_data.completions > 0 else 0.0
        avg_per_attempt = int(category_data.total_yards) / int(category_data.play_count) if category_data.play_count > 0 else 0.0

        row_data = [
            team_category[0],
            team_category[1],
            category_data.completions,
            category_data.play_count,
            round(int(category_data.completions) / int(category_data.play_count), 2),
            category_data.total_yards,
            round(avg_per_completion, 1),
            round(avg_per_attempt, 1),
            category_data.fumbles,
            category_data.interceptions,
            category_data.sacks,
            category_data.touchdowns_offense,
            category_data.touchdowns_offense * 6,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.PassInterceptionPercentage:
                row_data.insert(10, round(int(category_data.interceptions) / int(category_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.PassSackPercentage:
                row_data.insert(11 + extra, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.PassTouchdownPercentage:
                row_data.insert(12 + extra, round(int(category_data.touchdowns_offense) / int(category_data.play_count), 3))
                extra += 1

        self.pass_categories.worksheet.write_row(self.pass_categories.rows, 0, row_data)
        self.pass_categories.rows += 1

    def _add_defense_category(self, team_category, category_data):
        rush_avg = int(category_data.rush_yards_allowed) / int(category_data.run_plays_against) if category_data.run_plays_against > 0 else 0.0
        pass_avg = int(category_data.pass_yards_allowed) / int(category_data.pass_plays_against) if category_data.pass_plays_against > 0 else 0.0

        row_data = [
            team_category[0],
            team_category[1],
            f"{category_data.run_plays_against}/{category_data.rush_yards_allowed}",
            round(rush_avg, 1),
            f"{category_data.pass_plays_against}/{category_data.pass_yards_allowed}",
            round(pass_avg, 1),
            category_data.fumbles,
            category_data.interceptions,
            category_data.sacks,
            category_data.touchdowns_defense,
            category_data.touchdowns_offense,
        ]

        if self.perform_calculations:
            extra = 0
            if self.config.AdditionalColumns.DefenseTurnoverPercentage:
                row_data.insert(8, round((int(category_data.fumbles) + int(category_data.interceptions)) / int(category_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.DefenseSackPercentage:
                row_data.insert(9 + extra, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra += 1
            if self.config.AdditionalColumns.DefenseOffTdPercentage:
                row_data.insert(11 + extra, round(int(category_data.touchdowns_offense) / int(category_data.play_count), 3))
                extra += 1

        self.def_categories.worksheet.write_row(self.def_categories.rows, 0, row_data)
        self.def_categories.rows += 1

    # -- Conditional formatting -----------------------------------------------

    def _add_conditional_format(self):
        border_format = self.workbook.add_format()
        border_format.set_border_color("#0000FF")
        border_format.set_top(1)
        border_format.set_bottom(1)

        formula = {"type": "formula", "criteria": '=CELL("row")=ROW()', "format": border_format}

        sheets = [self.run, self.pass_, self.def_, self.tendencies]
        if self.config.Settings.CalculateCategoryStats:
            sheets.extend([self.run_categories, self.pass_categories, self.def_categories])

        for state in sheets:
            state.worksheet.conditional_format(1, 0, state.rows, state.columns - 1, formula)
