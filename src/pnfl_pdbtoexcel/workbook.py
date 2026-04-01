from pathlib import Path

import xlsxwriter

from pnfl_playpool import DefensivePlayRecord, OffensivePlayRecord, PlayRecord

from .pdb import PLAY_DATA


def _get_pdbtoexcel_module():
    from . import pdb_to_excel as pdbtoexcel_module
    return pdbtoexcel_module


def _get_config():
    return _get_pdbtoexcel_module().get_config()

def _get_runtime_path(filename):
    return _get_pdbtoexcel_module().get_runtime_path(filename)


class ExcelPdbWorkbook:

    def __init__(self, filename, perform_calculations):
        self.filename = filename
        self.perform_calculations = perform_calculations

    def __enter__(self):
        config = _get_config()
        filepath = Path(self.filename)
        filepath.parent.mkdir(exist_ok=True)

        self.workbook = xlsxwriter.Workbook(filepath)
        file_extension = filepath.suffix
        if file_extension == '.xlsm':
            if config['Settings']['CalculateCategoryStats']:
                self.workbook.add_vba_project(str(_get_runtime_path("vbaProject_categories.bin")))
            else:
                self.workbook.add_vba_project(str(_get_runtime_path("vbaProject.bin")))

        text_format = self.workbook.add_format({'num_format': '@'})
        avg_format = self.workbook.add_format({'num_format': '0.0'})
        percent_format_0 = self.workbook.add_format({'num_format': '0%'})
        percent_format_1 = self.workbook.add_format({'num_format': '0.0%'})
        center_format = self.workbook.add_format()
        center_format.set_align('center')
        options_header_format = self.workbook.add_format({'bg_color': '#FABF8F',
                                                          'bold': True,
                                                          'font_size': 10})
        option_label_format = self.workbook.add_format({'bg_color': '#FDE9D9',
                                                        'bold': True,
                                                        'font_size': 10})
        option_format = self.workbook.add_format({'bg_color': '#FFFF00',
                                                  'bold': True,
                                                  'font_size': 10,
                                                  'border': 2,
                                                  'border_color': 'black'})
        option_note_format = self.workbook.add_format({'bg_color': '#FDE9D9',
                                                       'bold': True,
                                                       'font_size': 8})

        self.options_worksheet = self.workbook.add_worksheet('Options')
        self.options_worksheet.set_vba_name('wsOptions')
        self.options_worksheet.set_column_pixels(0, 0, 6)
        row_data = ['OPTIONS - MISCELLANEOUS', '', '', '', '']
        self.options_worksheet.write_row(0, 1, row_data, options_header_format)
        row_data = ['Highlight Selected Row', '', '', '', '']
        self.options_worksheet.write_row(1, 1, row_data, option_label_format)
        self.options_worksheet.write(2, 1, 'Yes', option_format)
        row_data = ['      (when enabled, hold \'Ctrl\' to paste data)', '', '', '']
        self.options_worksheet.write_row(2, 2, row_data, option_note_format)

        self.run_worksheet = self.workbook.add_worksheet('Run Plays')
        self.run_worksheet.set_vba_name('wsRunPlays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'Attempts', 'Yards', 'Avg', 'Fumbles', 'TD', 'Pts', 'Custom1', 'Custom2']
        self.run_worksheet.set_column_pixels(0, 0, 120)
        self.run_worksheet.set_column_pixels(1, 1, 80)
        self.run_worksheet.set_column_pixels(2, 2, 47, text_format)
        self.run_worksheet.set_column_pixels(3, 3, 100)
        self.run_worksheet.set_column_pixels(4, 4, 59)
        self.run_worksheet.set_column_pixels(7, 7, None, avg_format)
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['RunFumblePercentage']:
                column_index = 9
                header.insert(column_index, 'Fumble %')
                self.run_worksheet.set_column_pixels(column_index, column_index, 70, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['RunTouchdownPercentage']:
                column_index = 10 + extra_columns
                header.insert(column_index, 'TD %')
                self.run_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.run_worksheet.write_row(0, 0, header)
        self.run_columns = len(header)
        self.run_rows = 1
        self.run_worksheet.freeze_panes(1, 0)
        self.run_worksheet.autofilter('A1:C1')  # pyright: ignore[reportCallIssue]

        self.pass_worksheet = self.workbook.add_worksheet('Pass Plays')
        self.pass_worksheet.set_vba_name('wsPassPlays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'Comp', 'Att', 'Comp %', 'Yards', 'Y/Comp', 'Y/Att', 'Fumbles', 'Int', 'Sacks', 'TD', 'Pts', 'Custom1', 'Custom2']
        self.pass_worksheet.set_column_pixels(0, 0, 120)
        self.pass_worksheet.set_column_pixels(1, 1, 80)
        self.pass_worksheet.set_column_pixels(2, 2, 47, text_format)
        self.pass_worksheet.set_column_pixels(3, 3, 100)
        self.pass_worksheet.set_column_pixels(4, 4, 49)
        self.pass_worksheet.set_column_pixels(5, 5, 49)
        self.pass_worksheet.set_column_pixels(6, 6, 49)
        self.pass_worksheet.set_column_pixels(7, 7, None, percent_format_0)
        self.pass_worksheet.set_column_pixels(9, 9, None, avg_format)
        self.pass_worksheet.set_column_pixels(10, 10, None, avg_format)
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['PassInterceptionPercentage']:
                column_index = 13
                header.insert(column_index, 'Int %')
                self.pass_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['PassSackPercentage']:
                column_index = 14 + extra_columns
                header.insert(column_index, 'Sack %')
                self.pass_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['PassTouchdownPercentage']:
                column_index = 15 + extra_columns
                header.insert(column_index, 'TD %')
                self.pass_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.pass_worksheet.write_row(0, 0, header)
        self.pass_rows = 1
        self.pass_columns = len(header)
        self.pass_worksheet.freeze_panes(1, 0)
        self.pass_worksheet.autofilter('A1:C1')  # pyright: ignore[reportCallIssue]

        self.def_worksheet = self.workbook.add_worksheet('Def Plays')
        self.def_worksheet.set_vba_name('wsDefPlays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'vs Run', 'Avg', 'vs Pass', 'Avg', 'Fumbles', 'Int', 'Sacks', 'TD/Def', 'TD/Off', 'Custom1', 'Custom2']
        self.def_worksheet.set_column_pixels(0, 0, 120)
        self.def_worksheet.set_column_pixels(1, 1, 126)
        self.def_worksheet.set_column_pixels(2, 2, 47, text_format)
        self.def_worksheet.set_column_pixels(3, 3, 100)
        self.def_worksheet.set_column_pixels(4, 4, 37)
        self.def_worksheet.set_column_pixels(5, 5, None, text_format)
        self.def_worksheet.set_column_pixels(6, 6, None, avg_format)
        self.def_worksheet.set_column_pixels(7, 7, None, text_format)
        self.def_worksheet.set_column_pixels(8, 8, None, avg_format)
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                header.insert(11, 'TO %')
                self.def_worksheet.set_column_pixels(11, 11, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['DefenseSackPercentage']:
                column_index = 12 + extra_columns
                header.insert(column_index, 'Sack %')
                self.def_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['DefenseOffTdPercentage']:
                column_index = 14 + extra_columns
                header.insert(column_index, 'TD/Off %')
                self.def_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.def_worksheet.write_row(0, 0, header)
        self.def_rows = 1
        self.def_columns = len(header)
        self.def_worksheet.freeze_panes(1, 0)
        self.def_worksheet.autofilter('A1:C1')  # pyright: ignore[reportCallIssue]

        self.tendencies_worksheet = self.workbook.add_worksheet('Tendencies')
        self.tendencies_worksheet.set_vba_name('wsTendencies')
        header = ['Team', 'Situation', 'Runs', 'Passes']
        self.tendencies_worksheet.set_column_pixels(0, 0, 120)
        self.tendencies_worksheet.set_column_pixels(1, 1, 140)
        self.tendencies_worksheet.write_row(0, 0, header)
        self.tendencies_rows = 1
        self.tendencies_columns = len(header)
        self.tendencies_worksheet.freeze_panes(1, 0)
        self.tendencies_worksheet.autofilter('A1:A1')  # pyright: ignore[reportCallIssue]

        if config['Settings']['CalculateCategoryStats']:
            self.run_categories_worksheet = self.workbook.add_worksheet('Run Categories')
            self.run_categories_worksheet.set_vba_name('wsRunCategories')
            header = ['Team', 'Category', 'Attempts', 'Yards', 'Avg', 'Fumbles', 'TD', 'Pts']
            self.run_categories_worksheet.set_column_pixels(0, 0, 120)
            self.run_categories_worksheet.set_column_pixels(1, 1, 80)
            self.run_categories_worksheet.set_column_pixels(4, 4, None, avg_format)
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['RunFumblePercentage']:
                    column_index = 6
                    header.insert(column_index, 'Fumble %')
                    self.run_categories_worksheet.set_column_pixels(column_index, column_index, 70, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['RunTouchdownPercentage']:
                    column_index = 7 + extra_columns
                    header.insert(column_index, 'TD %')
                    self.run_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
            self.run_categories_worksheet.write_row(0, 0, header)
            self.run_categories_columns = len(header)
            self.run_categories_rows = 1
            self.run_categories_worksheet.freeze_panes(1, 0)
            self.run_categories_worksheet.autofilter('A1:B1')  # pyright: ignore[reportCallIssue]

            self.pass_categories_worksheet = self.workbook.add_worksheet('Pass Categories')
            self.pass_categories_worksheet.set_vba_name('wsPassCategories')
            header = ['Team', 'Category', 'Comp', 'Att', 'Comp %', 'Yards', 'Y/Comp', 'Y/Att', 'Fumbles', 'Int', 'Sacks', 'TD', 'Pts']
            self.pass_categories_worksheet.set_column_pixels(0, 0, 120)
            self.pass_categories_worksheet.set_column_pixels(1, 1, 80)
            self.pass_categories_worksheet.set_column_pixels(2, 2, 49)
            self.pass_categories_worksheet.set_column_pixels(3, 3, 49)
            self.pass_categories_worksheet.set_column_pixels(4, 4, None, percent_format_0)
            self.pass_categories_worksheet.set_column_pixels(6, 6, None, avg_format)
            self.pass_categories_worksheet.set_column_pixels(7, 7, None, avg_format)
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['PassInterceptionPercentage']:
                    column_index = 10
                    header.insert(column_index, 'Int %')
                    self.pass_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['PassSackPercentage']:
                    column_index = 11 + extra_columns
                    header.insert(column_index, 'Sack %')
                    self.pass_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['PassTouchdownPercentage']:
                    column_index = 12 + extra_columns
                    header.insert(column_index, 'TD %')
                    self.pass_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
            self.pass_categories_worksheet.write_row(0, 0, header)
            self.pass_categories_rows = 1
            self.pass_categories_columns = len(header)
            self.pass_categories_worksheet.freeze_panes(1, 0)
            self.pass_categories_worksheet.autofilter('A1:B1')  # pyright: ignore[reportCallIssue]

            self.def_categories_worksheet = self.workbook.add_worksheet('Def Categories')
            self.def_categories_worksheet.set_vba_name('wsDefCategories')
            header = ['Team', 'Category', 'vs Run', 'Avg', 'vs Pass', 'Avg', 'Fumbles', 'Int', 'Sacks', 'TD/Def', 'TD/Off']
            self.def_categories_worksheet.set_column_pixels(0, 0, 120)
            self.def_categories_worksheet.set_column_pixels(1, 1, 126)
            self.def_categories_worksheet.set_column_pixels(2, 2, None, text_format)
            self.def_categories_worksheet.set_column_pixels(3, 3, None, avg_format)
            self.def_categories_worksheet.set_column_pixels(4, 4, None, text_format)
            self.def_categories_worksheet.set_column_pixels(5, 5, None, avg_format)
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                    column_index = 8
                    header.insert(column_index, 'TO %')
                    self.def_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['DefenseSackPercentage']:
                    column_index = 9 + extra_columns
                    header.insert(column_index, 'Sack %')
                    self.def_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['DefenseOffTdPercentage']:
                    column_index = 11 + extra_columns
                    header.insert(column_index, 'TD/Off %')
                    self.def_categories_worksheet.set_column_pixels(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
            self.def_categories_worksheet.write_row(0, 0, header)
            self.def_categories_rows = 1
            self.def_categories_columns = len(header)
            self.def_categories_worksheet.freeze_panes(1, 0)
            self.def_categories_worksheet.autofilter('A1:B1')  # pyright: ignore[reportCallIssue]
        return self

    def __exit__(self, *args):
        if self.workbook:
            self._add_conditional_format()
            self.workbook.close()

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
        row_data = [t.team_name.decode('ASCII'), 'First and 0-1', t.run_zero_to_one.first_down, t.pass_zero_to_one.first_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'First and 2-5', t.run_two_to_five.first_down, t.pass_two_to_five.first_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'First and 6-10', t.run_six_to_ten.first_down, t.pass_six_to_ten.first_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'First and 10+', t.run_ten_plus.first_down, t.pass_ten_plus.first_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1

        row_data = [t.team_name.decode('ASCII'), 'Second and 0-1', t.run_zero_to_one.second_down, t.pass_zero_to_one.second_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Second and 2-5', t.run_two_to_five.second_down, t.pass_two_to_five.second_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Second and 6-10', t.run_six_to_ten.second_down, t.pass_six_to_ten.second_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Second and 10+', t.run_ten_plus.second_down, t.pass_ten_plus.second_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1

        row_data = [t.team_name.decode('ASCII'), 'Third and 0-1', t.run_zero_to_one.third_down, t.pass_zero_to_one.third_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Third and 2-5', t.run_two_to_five.third_down, t.pass_two_to_five.third_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Third and 6-10', t.run_six_to_ten.third_down, t.pass_six_to_ten.third_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Third and 10+', t.run_ten_plus.third_down, t.pass_ten_plus.third_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1

        row_data = [t.team_name.decode('ASCII'), 'Fourth and 0-1', t.run_zero_to_one.fourth_down, t.pass_zero_to_one.fourth_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Fourth and 2-5', t.run_two_to_five.fourth_down, t.pass_two_to_five.fourth_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Fourth and 6-10', t.run_six_to_ten.fourth_down, t.pass_six_to_ten.fourth_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1
        row_data = [t.team_name.decode('ASCII'), 'Fourth and 10+', t.run_ten_plus.fourth_down, t.pass_ten_plus.fourth_down]
        self.tendencies_worksheet.write_row(self.tendencies_rows, 0, row_data)
        self.tendencies_rows += 1

    def _add_run_play(self, play_data, play_slot, play_record):
        play_type = ''
        if isinstance(play_record, OffensivePlayRecord) and play_record.qb_draw:
            play_type = 'QB draw'

        if play_data.play_count > 0:
            avg = int(play_data.total_yards) / int(play_data.play_count)
        else:
            avg = 0.0

        row_data = [
            play_data.team_name.decode('ASCII'),
            play_record.pool_category,
            play_slot,
            play_name,
            play_type,
            play_data.play_count,
            play_data.total_yards,
            round(avg, 1),
            play_data.fumbles,
            play_data.touchdowns_offense,
            play_data.touchdowns_offense * 6,
        ]

        if self.perform_calculations:
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['RunFumblePercentage']:
                row_data.insert(9, round(int(play_data.fumbles) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['RunTouchdownPercentage']:
                row_data.insert(10 + extra_columns, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra_columns += 1

        self.run_worksheet.write_row(self.run_rows, 0, row_data)
        self.run_rows += 1

    def _add_pass_play(self, play_data, play_slot, play_record):
        play_type = ''
        if isinstance(play_record, OffensivePlayRecord) and play_record.screen:
            play_type = 'Screen'

        if play_data.completions > 0:
            avg_per_completion = int(play_data.total_yards) / int(play_data.completions)
        else:
            avg_per_completion = 0.0

        if play_data.play_count > 0:
            avg_per_attempt = int(play_data.total_yards) / int(play_data.play_count)
        else:
            avg_per_attempt = 0.0

        row_data = [
            play_data.team_name.decode('ASCII'),
            play_record.pool_category,
            play_slot,
            play_data.play_name.decode('ASCII'),
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
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['PassInterceptionPercentage']:
                row_data.insert(13, round(int(play_data.interceptions) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['PassSackPercentage']:
                row_data.insert(14 + extra_columns, round(int(play_data.sacks) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['PassTouchdownPercentage']:
                row_data.insert(15 + extra_columns, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra_columns += 1

        self.pass_worksheet.write_row(self.pass_rows, 0, row_data)
        self.pass_rows += 1

    def _add_defense_play(self, play_data, play_slot, play_record):
        play_type = ''
        if isinstance(play_record, DefensivePlayRecord) and play_record.personnel_grouping is not None:
            play_type = play_record.personnel_grouping.value

        if play_data.run_plays_against > 0:
            rush_avg = int(play_data.rush_yards_allowed) / int(play_data.run_plays_against)
        else:
            rush_avg = 0.0

        if play_data.pass_plays_against > 0:
            pass_avg = int(play_data.pass_yards_allowed) / int(play_data.pass_plays_against)
        else:
            pass_avg = 0.0

        row_data = [
            play_data.team_name.decode('ASCII'),
            play_record.pool_category,
            play_slot,
            play_data.play_name.decode('ASCII'),
            play_type,
            f"{str(play_data.run_plays_against)}/{str(play_data.rush_yards_allowed)}",
            round(rush_avg, 1),
            f"{str(play_data.pass_plays_against)}/{str(play_data.pass_yards_allowed)}",
            round(pass_avg, 1),
            play_data.fumbles,
            play_data.interceptions,
            play_data.sacks,
            play_data.touchdowns_defense,
            play_data.touchdowns_offense,
        ]

        if self.perform_calculations:
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                row_data.insert(11, round((int(play_data.fumbles) + int(play_data.interceptions)) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['DefenseSackPercentage']:
                row_data.insert(12 + extra_columns, round(int(play_data.sacks) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['DefenseOffTdPercentage']:
                row_data.insert(14 + extra_columns, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra_columns += 1

        self.def_worksheet.write_row(self.def_rows, 0, row_data)
        self.def_rows += 1

    def _add_run_category(self, team_category, category_data):
        if category_data.play_count > 0:
            avg = int(category_data.total_yards) / int(category_data.play_count)
        else:
            avg = 0.0

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
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['RunFumblePercentage']:
                row_data.insert(6, round(int(category_data.fumbles) / int(category_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['RunTouchdownPercentage']:
                row_data.insert(7 + extra_columns, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra_columns += 1

        self.run_categories_worksheet.write_row(self.run_categories_rows, 0, row_data)
        self.run_categories_rows += 1

    def _add_pass_category(self, team_category, category_data):
        if category_data.completions > 0:
            avg_per_completion = int(category_data.total_yards) / int(category_data.completions)
        else:
            avg_per_completion = 0.0

        if category_data.play_count > 0:
            avg_per_attempt = int(category_data.total_yards) / int(category_data.play_count)
        else:
            avg_per_attempt = 0.0

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
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['PassInterceptionPercentage']:
                row_data.insert(10, round(int(category_data.interceptions) / int(category_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['PassSackPercentage']:
                row_data.insert(11 + extra_columns, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['PassTouchdownPercentage']:
                row_data.insert(12 + extra_columns, round(int(category_data.touchdowns_offense) / int(category_data.play_count), 3))
                extra_columns += 1

        self.pass_categories_worksheet.write_row(self.pass_categories_rows, 0, row_data)
        self.pass_categories_rows += 1

    def _add_defense_category(self, team_category, category_data):
        if category_data.run_plays_against > 0:
            rush_avg = int(category_data.rush_yards_allowed) / int(category_data.run_plays_against)
        else:
            rush_avg = 0.0

        if category_data.pass_plays_against > 0:
            pass_avg = int(category_data.pass_yards_allowed) / int(category_data.pass_plays_against)
        else:
            pass_avg = 0.0

        row_data = [
            team_category[0],
            team_category[1],
            f"{str(category_data.run_plays_against)}/{str(category_data.rush_yards_allowed)}",
            round(rush_avg, 1),
            f"{str(category_data.pass_plays_against)}/{str(category_data.pass_yards_allowed)}",
            round(pass_avg, 1),
            category_data.fumbles,
            category_data.interceptions,
            category_data.sacks,
            category_data.touchdowns_defense,
            category_data.touchdowns_offense,
        ]

        if self.perform_calculations:
            config = _get_config()
            extra_columns = 0
            if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                row_data.insert(8, round((int(category_data.fumbles) + int(category_data.interceptions)) / int(category_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['DefenseSackPercentage']:
                row_data.insert(9 + extra_columns, round(int(category_data.sacks) / int(category_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['DefenseOffTdPercentage']:
                row_data.insert(11 + extra_columns, round(int(category_data.touchdowns_offense) / int(category_data.play_count), 3))
                extra_columns += 1

        self.def_categories_worksheet.write_row(self.def_categories_rows, 0, row_data)
        self.def_categories_rows += 1

    def _add_conditional_format(self):
        config = _get_config()
        border_format = self.workbook.add_format()
        border_format.set_border_color('#0000FF')
        border_format.set_top(1)
        border_format.set_bottom(1)

        self.run_worksheet.conditional_format(1, 0, self.run_rows, self.run_columns - 1,
                                              {'type': 'formula',
                                               'criteria': '=CELL("row")=ROW()',
                                               'format': border_format})
        self.pass_worksheet.conditional_format(1, 0, self.pass_rows, self.pass_columns - 1,
                                               {'type': 'formula',
                                                'criteria': '=CELL("row")=ROW()',
                                                'format': border_format})
        self.def_worksheet.conditional_format(1, 0, self.def_rows, self.def_columns - 1,
                                              {'type': 'formula',
                                               'criteria': '=CELL("row")=ROW()',
                                               'format': border_format})
        self.tendencies_worksheet.conditional_format(1, 0, self.tendencies_rows, self.tendencies_columns - 1,
                                                     {'type': 'formula',
                                                      'criteria': '=CELL("row")=ROW()',
                                                      'format': border_format})

        if config['Settings']['CalculateCategoryStats']:
            self.run_categories_worksheet.conditional_format(1, 0, self.run_categories_rows, self.run_categories_columns - 1,
                                                             {'type': 'formula',
                                                              'criteria': '=CELL("row")=ROW()',
                                                              'format': border_format})
            self.pass_categories_worksheet.conditional_format(1, 0, self.pass_categories_rows, self.pass_categories_columns - 1,
                                                              {'type': 'formula',
                                                               'criteria': '=CELL("row")=ROW()',
                                                               'format': border_format})
            self.def_categories_worksheet.conditional_format(1, 0, self.def_categories_rows, self.def_categories_columns - 1,
                                                             {'type': 'formula',
                                                              'criteria': '=CELL("row")=ROW()',
                                                              'format': border_format})
