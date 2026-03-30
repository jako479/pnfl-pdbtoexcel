###############################################################################
#
# PdbToExcel - Creates an Excel workbook from a WinLogStats PDB and FBPro 98
# offensive and defensive game plan files (.PLN)
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2024, Brian Jacobs, brian.andrew.jacobs@gmail.com
#

import argparse
import configparser
import hashlib
import math
from pathlib import Path
import socket
import sys
from fbpro98_gameplan import PLN
from .pdb import PDB, PLAY_DATA
from .workbook import ExcelPdbWorkbook


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent.parent


def get_runtime_path(filename):
    return SCRIPT_DIR / filename


def set_config_path(config_path):
    get_config_path.config_path = Path(config_path).expanduser().resolve()
    if hasattr(get_config, "config_dictionary"):
        delattr(get_config, "config_dictionary")


def get_config_path():
    if not hasattr(get_config_path, "config_path"):
        candidates = [
            SCRIPT_DIR / "PdbToExcel.ini",
            PROJECT_DIR / "config" / "PdbToExcel.ini",
        ]
        get_config_path.config_path = next(
            (candidate for candidate in candidates if candidate.is_file()),
            candidates[0],
        )
    return get_config_path.config_path


DELETED_PLAYS = [
    'ATF0ELOB'
]

TOTAL_STATS_FILTER = {
    'GLR': 2.7,
    'RL': 3.6,
    'RM': 4.6,
    'RR': 4.6,
    'GLP': (0.7, None),
    'PSL': (0.60, 6.0),
    'PSM': (0.60, 5.5),
    'PSR': (0.60, 6.0),
    'PML': (0.50, 6.8),
    'PMM': (0.50, 6.8),
    'PMR': (0.50, 6.0),
    'PLR': (0.40, 6.0),
    'PRD': (0.29, None),
}

DEFENSE_OUTPUT_CATEGORIES = {
    'RunRight',
    'RunMiddle',
    'RunLeft',
    'PassShort',
    'PassMedium',
    'PassLong',
    'RunDazzle',
    'PassDazzle',
    'GLrun',
    'GLpass',
}


def get_config():
    if not hasattr(get_config, 'config_dictionary'):
        md5 = hashlib.md5(socket.gethostname().encode())
        cp = configparser.ConfigParser()
        cp.read(get_config_path(), encoding="utf-8")
        config_dict = {}
        config_dict['Settings'] = {}
        config_dict['Settings']['Team'] = cp.get("Settings", "Team", fallback="")
        config_dict['Settings']['PnflPath'] = cp.get("Settings", "PnflPath", fallback="C:\\SIERRA\\FbPro98\\PNFL")
        config_dict['Settings']['CalculateTotalStats'] = cp.getboolean("Settings", "CalculateTotalStats", fallback=True)
        config_dict['Settings']['CalculateCategoryStats'] = (md5.hexdigest() == '5c4b925bf527c4f8581815a35a10d658')
        config_dict['Settings']['CalculateGroupedCategoryStats'] = (md5.hexdigest() == '5c4b925bf527c4f8581815a35a10d658' and 1)
        config_dict['AdditionalColumns'] = {}
        config_dict['AdditionalColumns']['RunFumblePercentage'] = cp.getboolean("AdditionalColumns", "RunFumblePercentage", fallback=True)
        config_dict['AdditionalColumns']['RunTouchdownPercentage'] = cp.getboolean("AdditionalColumns", "RunTouchdownPercentage", fallback=True)
        config_dict['AdditionalColumns']['PassInterceptionPercentage'] = cp.getboolean("AdditionalColumns", "PassInterceptionPercentage", fallback=True)
        config_dict['AdditionalColumns']['PassSackPercentage'] = cp.getboolean("AdditionalColumns", "PassSackPercentage", fallback=True)
        config_dict['AdditionalColumns']['PassTouchdownPercentage'] = cp.getboolean("AdditionalColumns", "PassTouchdownPercentage", fallback=True)
        config_dict['AdditionalColumns']['DefenseTurnoverPercentage'] = cp.getboolean("AdditionalColumns", "DefenseTurnoverPercentage", fallback=True)
        config_dict['AdditionalColumns']['DefenseSackPercentage'] = cp.getboolean("AdditionalColumns", "DefenseSackPercentage", fallback=True)
        config_dict['AdditionalColumns']['DefenseOffTdPercentage'] = cp.getboolean("AdditionalColumns", "DefenseOffTdPercentage", fallback=True)
        get_config.config_dictionary = config_dict
    return get_config.config_dictionary


###############################################################################
#
# PdbWorkbookCreator - A class for creating the Excel workbook from a PDB, and
# offensive and defensive game plans (PLN)
#
###############################################################################

class PdbWorkbookCreator:

    ###########################################################################
    #
    # Initialization/Destruction
    #
    ###########################################################################

    def __init__(self, pdb_filename, pln_def_filename, pln_off_filename):
        # Build the dictionary of play names to play category (e.g. RM, GLP, PSR)
        # The category is determined by the directory in which the play resides.
        config = get_config()
        self.play_pool = PlayPool(config['Settings']['PnflPath'])

        self.pln_offense = None
        self.pln_defense = None

        # Parse the PDB
        self.pdb = PDB(pdb_filename)
        self.pdb.convert_invalid_play_data(self.play_pool)

        # Parse the defensive game plan
        if pln_def_filename:
            self.pln_defense = PLN(pln_def_filename)

        # Parse the offensive game plan
        if pln_off_filename:
            self.pln_offense = PLN(pln_off_filename)

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def create_workbook(self, filename, perform_calculations, calculate_totals, filter_total_stats):
        print(f"Attempting to create \'{filename}\'")
        if not perform_calculations:
            print("Skipping extra calculations")

        with ExcelPdbWorkbook(filename, perform_calculations) as workbook:
            config = get_config()
            combined_plays = {} if calculate_totals else None
            missing_play_files_logged = set()

            for play_in_pdb, play_name, _, play_attributes in self._iter_category_source_plays(missing_play_files_logged):
                play_slot = self._get_play_slot(play_in_pdb, play_name)
                workbook.add_play(play_in_pdb, play_slot, play_attributes)
                if calculate_totals:
                    self._add_play_stats_to_total_play(combined_plays, play_in_pdb)

            if calculate_totals:
                self._add_total_plays_to_workbook(workbook, combined_plays, filter_total_stats)

            # Add all tendency data in the PDB to the workbook
            for row_num, tendency_data in enumerate(self.pdb.tendencies):
                workbook.add_tendency(tendency_data)

            if config['Settings']['CalculateCategoryStats']:
                team_categories_data, categories_data = self._collect_category_stats(
                    calculate_totals=calculate_totals,
                    group_categories=config['Settings']['CalculateGroupedCategoryStats'],
                )

                # Add the team category stats to the workbook
                for team_category, category_data in team_categories_data.items():
                    workbook.add_category(team_category, category_data)

                if calculate_totals:
                    # Add the total category stats to the workbook
                    for category_name, category_data in categories_data.items():
                        workbook.add_category(('`Total Stats', category_name), category_data)

        print("Conversion complete")


    def _iter_tracked_plays(self):
        play_types = (
            PLAY_DATA.PLAY_TYPE.RUN,
            PLAY_DATA.PLAY_TYPE.PASS,
            PLAY_DATA.PLAY_TYPE.DEFENSE,
        )
        for play_type in play_types:
            plays_list = list(self.pdb.plays[play_type].values())
            plays_list.sort(key=lambda x: x.team_name)
            for play_in_pdb in plays_list:
                yield play_in_pdb


    def _iter_category_source_plays(self, missing_play_files_logged=None):
        for play_in_pdb in self._iter_tracked_plays():
            play_name = play_in_pdb.play_name.decode('ASCII')
            team_name = play_in_pdb.team_name.decode('ASCII')
            play_attributes = self.play_pool.get_play_attributes(play_name)
            if self._play_is_unknown_to_play_pool(play_attributes):
                self._log_unknown_play_pool_entry(play_name, missing_play_files_logged)
            if not self._should_export_play(play_in_pdb, play_name, play_attributes):
                continue
            yield play_in_pdb, play_name, team_name, play_attributes


    def _play_is_unknown_to_play_pool(self, play_attributes):
        return play_attributes.category == 'Unknown'


    def _log_unknown_play_pool_entry(self, play_name, missing_play_files_logged=None):
        # This checks the category from PlayPool lookup, not any category stored in the PDB.
        # A play can have valid stats in the database and still come back as 'Unknown' here
        # if there is no matching current play file in the play pool.
        if missing_play_files_logged is not None and play_name not in missing_play_files_logged:
            if play_name in DELETED_PLAYS:
                print(f"Skipping deleted play \'{play_name}\'")
            else:
                print(f"Play file not found for play \'{play_name}\'")
            missing_play_files_logged.add(play_name)


    def _should_export_play(self, play_in_pdb, play_name, play_attributes):
        if (play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE and
                play_attributes.category not in DEFENSE_OUTPUT_CATEGORIES):
            return False
        return play_name not in DELETED_PLAYS


    def _get_play_slot(self, play_in_pdb, play_name):
        play_slot = ''
        play_in_plan = None
        if (self.pln_offense and
                (play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.RUN or
                 play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.PASS)):
            play_in_plan = self.pln_offense.normal_plays.get(play_name)
        elif (self.pln_defense and
                play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE):
            play_in_plan = self.pln_defense.normal_plays.get(play_name)
        if play_in_plan:
            play_slot = self._get_slot_string(play_in_plan.slot)
        return play_slot


    def _collect_category_stats(self, calculate_totals, group_categories):
        categories_data = {}
        team_categories_data = {}

        for play_in_pdb, _, team_name, play_attributes in self._iter_category_source_plays():
            self._add_play_stats_to_team_categories(
                team_categories_data=team_categories_data,
                play_in_pdb=play_in_pdb,
                team_name=team_name,
                category=play_attributes.category,
                group_categories=group_categories,
            )
            if calculate_totals:
                self._add_play_stats_to_category(categories_data, play_in_pdb, play_attributes.category)

        return team_categories_data, categories_data


    def _add_play_stats_to_total_play(self, combined_plays, play_in_pdb):
        if play_in_pdb.play_name in combined_plays:
            combined_play_data = combined_plays[play_in_pdb.play_name]
        else:
            combined_play_data = PLAY_DATA()
            combined_play_data.play_type = play_in_pdb.play_type
            combined_play_data.team_name = bytes('`Total Stats', 'ASCII')
            combined_play_data.play_name = play_in_pdb.play_name
        combined_play_data += play_in_pdb
        combined_plays[play_in_pdb.play_name] = combined_play_data


    def _add_total_plays_to_workbook(self, workbook, combined_plays, filter_total_stats):
        for play_in_pdb in combined_plays.values():
            play_name = play_in_pdb.play_name.decode('ASCII')
            play_attributes = self.play_pool.get_play_attributes(play_name)
            if not self._should_export_play(play_in_pdb, play_name, play_attributes):
                continue
            if filter_total_stats and not self._play_meets_criteria(play_in_pdb, play_attributes):
                continue
            play_slot = self._get_play_slot(play_in_pdb, play_name)
            workbook.add_play(play_in_pdb, play_slot, play_attributes)


    def _add_play_stats_to_team_categories(self, team_categories_data, play_in_pdb, team_name, category, group_categories=False):
        team_category = (team_name, category)
        self._add_play_stats_to_team_category(team_categories_data, play_in_pdb, team_category)

        if group_categories:
            # Also calculate total stats for team by grouped categories
            if play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.RUN:
                team_category = (team_name, "TOTAL RUNS")
                self._add_play_stats_to_team_category(team_categories_data, play_in_pdb, team_category)
            elif play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.PASS:
                if category in ('PSL', 'PSM', 'PSR'):
                    team_category = (team_name, "TOTAL PS")
                    self._add_play_stats_to_team_category(team_categories_data, play_in_pdb, team_category)
                if category in ('PML', 'PMM', 'PMR'):
                    team_category = (team_name, "TOTAL PM")
                    self._add_play_stats_to_team_category(team_categories_data, play_in_pdb, team_category)
                team_category = (team_name, "TOTAL PASSES")
                self._add_play_stats_to_team_category(team_categories_data, play_in_pdb, team_category)


    def _add_play_stats_to_team_category(self, team_categories_data, play_in_pdb, team_category):
        if team_category in team_categories_data:
            team_category_data = team_categories_data[team_category]
        else:
            team_category_data = PLAY_DATA()
            team_category_data.play_type = play_in_pdb.play_type
        team_category_data += play_in_pdb
        team_categories_data[team_category] = team_category_data


    def _add_play_stats_to_category(self, categories_data, play_in_pdb, category):
        # Add play stats to the category data
        if category in categories_data:
            category_data = categories_data[category]
        else:
            category_data = PLAY_DATA()
            category_data.play_type = play_in_pdb.play_type
        category_data += play_in_pdb
        # Add/Replace the category stats
        categories_data[category] = category_data


    def _get_slot_string(self, slot):
        slot_string = ''
        if slot >= 0:
            slot_string = f'{str(math.floor(slot / 4) + 1)}-{(slot % 4) + 1}'
        return slot_string


    def _play_meets_criteria(self, play_in_pdb, play_attributes):
        if play_attributes.category not in TOTAL_STATS_FILTER:
            return True

        if play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            if play_in_pdb.play_count < 7:
                return False
            if play_in_pdb.play_count > 0:
                avg = play_in_pdb.total_yards / play_in_pdb.play_count
            else:
                avg = 0.0
            return (avg >= TOTAL_STATS_FILTER[play_attributes.category])
        elif play_in_pdb.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            if play_in_pdb.play_count < 7:
                return False
            comp_pct = play_in_pdb.completions / play_in_pdb.play_count
            ypa = play_in_pdb.total_yards / play_in_pdb.play_count
            min_comp_pct = TOTAL_STATS_FILTER[play_attributes.category][0]
            min_ypa = TOTAL_STATS_FILTER[play_attributes.category][1]
            return (comp_pct >= min_comp_pct or (min_ypa is not None and ypa >= min_ypa))
        return True


class PlayAttributes:
    """Class for storing attributes about a play as derived from the play file name and path"""
    def __init__(self, category=''):
        self.category = category
        self.run_and_shoot = False
        self.screen = False
        self.three_four = False
        self.four_three = False


###############################################################################
#
# PlayPool - A class for accessing the play files
#
###############################################################################

class PlayPool:

    ###########################################################################
    #
    # Initialization/Destruction
    #
    ###########################################################################

    def __init__(self, dirpath):
        self.play_dict = {}
        for play_file in Path(dirpath).glob('**/*.ply'):
            play_name = play_file.name[:-4].upper()  # excludes the extension
            dir_name1 = play_file.parents[0].name    # parent directory
            dir_name2 = play_file.parents[1].name    # parent of the parent directory
            play_attributes = PlayAttributes()
            if dir_name1.startswith(('34', '43')):
                if dir_name1.startswith('34'):
                    play_attributes.three_four = True
                else:
                    play_attributes.four_three = True
                # combine 34 and 43 defenses into single categories (e.g. RunMiddle)
                play_attributes.category = dir_name1[2:]
            elif dir_name1 == 'Screens':
                play_attributes.screen = True
                play_attributes.category = dir_name2
            elif dir_name2 == 'R&SDefs':
                play_attributes.run_and_shoot = True
                play_attributes.category = dir_name1
            else:
                # All other plays
                play_attributes.category = dir_name1
            self.play_dict[play_name] = play_attributes

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def get_play_attributes(self, play_name):
        return self.play_dict.get(play_name.upper(), PlayAttributes('Unknown'))

    def is_run_play(self, play_name):
        run_categories = ['GLR', 'RL', 'RM', 'RR']
        play_attributes = self.get_play_attributes(play_name)
        return (play_attributes.category in run_categories)

    def is_pass_play(self, play_name):
        pass_categories = ['GLP', 'PLR', 'PML', 'PMM', 'PMR', 'PRD', 'PSL', 'PSM', 'PSR']
        play_attributes = self.get_play_attributes(play_name)
        return (play_attributes.category in pass_categories)


###########################################################################
#
# MAIN ROUTINE
#
###########################################################################

def build_argument_parser():

    def valid_existing_file(param, expected_extensions):
        filepath = Path(param).expanduser()
        if filepath.suffix.lower() not in expected_extensions:
            extensions = ", ".join(expected_extensions)
            raise argparse.ArgumentTypeError(f"File must have one of these extensions: {extensions}")
        if not filepath.is_file():
            raise argparse.ArgumentTypeError(f"File not found: {filepath}")
        return str(filepath)

    def valid_output_file(param):
        filepath = Path(param).expanduser()
        if filepath.suffix.lower() not in ('.xlsm', '.xlsx'):
            raise argparse.ArgumentTypeError('File must have a xlsm or xlsx extension')
        return str(filepath)

    parser = argparse.ArgumentParser(
        description="Create an Excel workbook from a WinLogStats PDB and optional FBPro 98 game plans."
    )
    parser.add_argument("pdbfile", type=lambda value: valid_existing_file(value, ('.pdb',)), help="WinLogStats database file (.PDB)")
    parser.add_argument("-d", "--plnfile-defense", type=lambda value: valid_existing_file(value, ('.pln',)), help="defensive game plan file (.PLN)")
    parser.add_argument("-o", "--plnfile-offense", type=lambda value: valid_existing_file(value, ('.pln',)), help="offensive game plan file (.PLN)")
    parser.add_argument("outputfile", type=valid_output_file, help="save to this XLSX/XLSM file")
    parser.add_argument("--config", type=lambda value: valid_existing_file(value, ('.ini',)), help="use this INI file instead of the default config lookup")
    parser.add_argument("-c", "--skip-calcs", default=False, action='store_true', help="prevents the extra calculation columns (overrides config settings)")
    parser.add_argument("-t", "--skip-totals", default=False, action='store_true', help="prevents totalling stats (overrides config settings)")
    return parser


def run(args=None):

    # Execute as command line script
    if args is None:
        args = sys.argv[1:]
    parser = build_argument_parser()
    args = parser.parse_args(args)

    if args.config:
        set_config_path(args.config)

    config = get_config()

    # calculate totals if enabled in configuration settings and not skipping via command-line
    calculate_totals = (config['Settings']['CalculateTotalStats'] and not args.skip_totals)

    creator = PdbWorkbookCreator(args.pdbfile, args.plnfile_defense, args.plnfile_offense)
    creator.create_workbook(args.outputfile, not args.skip_calcs, calculate_totals, False)

    # Done
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
