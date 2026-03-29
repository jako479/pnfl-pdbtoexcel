###############################################################################
#
# PdbToExcel - Creates an Excel workbook from a WinLogStats PDB and FBPro 98
# offensive and defensive game plan files (.PLN)
#
# SPDX-License-Identifier: BSD-2-Clause
# Copyright 2024, Brian Jacobs, brian.andrew.jacobs@gmail.com
#

import argparse
import ctypes
import configparser
from enum import Enum, IntEnum
import hashlib
import math
import os
from pathlib import Path
import socket
import struct
import sys
from workbook import ExcelPdbWorkbook


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


RENAMED_PLAYS = {
    'WR47PT01': 'WR27PT01',
    'WR48PT01': 'WR28PT01',
}

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


##################################################################
##################################################################
#
# WINLOGSTATS PDB FILE FORMAT
#
##################################################################
#
# 0|PLAY_DATA|0|PLAY_DATA|...|0|PLAY_DATA|1|TENDENCY_DATA|...|1|TENDENCY_DATA|<EOF>
#
# There are 0-N PLAY_DATA sections (one or more for each team in the file)
# There is 0-N TENDENCY_DATA sections (one for each team in the file)
# Occasionally, there is a zero-filled TENDENCY_DATA section
#
##################################################################
##################################################################

###############################################################################
#
# PLAY_DATA - A ctypes struct for storing play data
#
###############################################################################

class PLAY_DATA(ctypes.LittleEndianStructure):

    class PLAY_TYPE(IntEnum):
        RUN = 0
        PASS = 1
        SPECIAL = 2
        DEFENSE = 3
        ONSIDE = 5

    _fields_ = [("play_type",           ctypes.c_uint32),    # 0x00  // 0x00000000 // Run
                                                             #       // 0x00000001 // Pass
                                                             #       // 0x00000002 // Special
                                                             #       // 0x00000003 // Defense
                                                             #       // 0x00000005 // Onside
                ("team_name",           ctypes.c_char*64),   # 0x04
                ("play_name",           ctypes.c_char*128),  # 0x44
                ("total_yards",         ctypes.c_int32),     # 0xC4  // OFFENSE
                ("play_count",          ctypes.c_uint32),    # 0xC8  // OFFENSE/DEFENSE/SPECIAL
                ("completions",         ctypes.c_uint32),    # 0xCC  // OFFENSE
                ("sacks",               ctypes.c_uint32),    # 0xD0	 // OFFENSE/DEFENSE
                ("fumbles",             ctypes.c_uint32),    # 0xD4	 // OFFENSE/DEFENSE
                ("interceptions",       ctypes.c_uint32),    # 0xD8	 // OFFENSE/DEFENSE
                ("touchdowns_offense",  ctypes.c_uint32),    # 0xDC	 // OFFENSE/DEFENSE
                ("touchdowns_defense",  ctypes.c_uint32),    # 0xE0	 // OFFENSE/DEFENSE
                ("unknown1",            ctypes.c_int32),     # 0xE4	 // 0x00000000   // UNKNOWN
                ("unknown2",            ctypes.c_int32),     # 0xE8	 // 0x00000000   // UNKNOWN
                ("points_scored",       ctypes.c_uint32),    # 0xEC	 // OFFENSE (TD points scored by offense or defense)
                                                             #       // DOESN'T COUNT TWO POINT CONVERSIONS
                ("run_plays_against",   ctypes.c_uint32),    # 0xF0	 // DEFENSE
                ("pass_plays_against",  ctypes.c_uint32),    # 0xF4	 // DEFENSE
                ("rush_yards_allowed",  ctypes.c_int32),     # 0xF8  // DEFENSE
                ("pass_yards_allowed",  ctypes.c_int32)]     # 0xFC  // DEFENSE

    def is_valid(self):
        # sanity check
        # check for valid play type, team name, and play name
        if ((self.play_type == self.PLAY_TYPE.RUN or
             self.play_type == self.PLAY_TYPE.PASS or
             self.play_type == self.PLAY_TYPE.SPECIAL or
             self.play_type == self.PLAY_TYPE.DEFENSE or
             self.play_type == self.PLAY_TYPE.ONSIDE) and
                len(self.team_name) > 0 and
                len(self.play_name) > 0):
            return True
        return False

    def __iadd__(self, other):
        self.total_yards += other.total_yards
        self.play_count += other.play_count
        self.completions += other.completions
        self.sacks += other.sacks
        self.fumbles += other.fumbles
        self.interceptions += other.interceptions
        self.touchdowns_offense += other.touchdowns_offense
        self.touchdowns_defense += other.touchdowns_defense
        self.unknown1 += other.unknown1
        self.unknown2 += other.unknown2
        self.points_scored += other.points_scored
        self.run_plays_against += other.run_plays_against
        self.pass_plays_against += other.pass_plays_against
        self.rush_yards_allowed += other.rush_yards_allowed
        self.pass_yards_allowed += other.pass_yards_allowed
        return self


###############################################################################
#
# DOWN_DATA - A ctypes struct for storing situational data by down
#
###############################################################################

class DOWN_DATA(ctypes.LittleEndianStructure):
    _fields_ = [("first_down",  ctypes.c_uint32),           # 0x00
                ("second_down", ctypes.c_uint32),           # 0x04
                ("third_down",  ctypes.c_uint32),           # 0x08
                ("fourth_down", ctypes.c_uint32)]           # 0x0C


###############################################################################
#
# TENDENCY_DATA - A ctypes struct for storing tendency data by situatation
#
###############################################################################

class TENDENCY_DATA(ctypes.LittleEndianStructure):
    _fields_ = [("team_name",           ctypes.c_char*64),  # 0x00
                ("run_zero_to_one",     DOWN_DATA),         # 0x40
                ("pass_zero_to_one",    DOWN_DATA),         # 0x50
                ("run_two_to_five",     DOWN_DATA),         # 0x60
                ("pass_two_to_five",    DOWN_DATA),         # 0x70
                ("run_six_to_ten",      DOWN_DATA),         # 0x80
                ("pass_six_to_ten",     DOWN_DATA),         # 0x90
                ("run_ten_plus",        DOWN_DATA),         # 0xA0
                ("pass_ten_plus",       DOWN_DATA)]         # 0xB0

    def is_valid(self):
        return len(self.team_name) > 0


###############################################################################
#
# InvalidPDBError - A custom exception
#
###############################################################################

class InvalidPDBError(BaseException):
    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


###############################################################################
#
# PDB - A class for opening and parsing a WinLogStats PDB file
#
###############################################################################

class PDB:
    class DATA_TYPE(IntEnum):
        PLAY = 0
        TENDENCY = 1

    ###########################################################################
    #
    # Initialization/Destruction
    #
    ###########################################################################

    def __init__(self, filename):
        self.filename = filename
        self.plays = {PLAY_DATA.PLAY_TYPE.RUN: {},
                      PLAY_DATA.PLAY_TYPE.PASS: {},
                      PLAY_DATA.PLAY_TYPE.SPECIAL: {},
                      PLAY_DATA.PLAY_TYPE.DEFENSE: {},
                      PLAY_DATA.PLAY_TYPE.ONSIDE: {}}
        self.tendencies = []

        file_path = Path(self.filename)

        with open(file_path, 'rb') as pdb:
            # loop to process the entire file
            while True:
                # Read data type identifier of the next data chunk
                data = pdb.read(1)
                if not data:
                    # EOF
                    self.tendencies.sort(key=lambda x: x.team_name)
                    break
                data_type = int.from_bytes(data, byteorder='little')
                if data_type < self.DATA_TYPE.PLAY or data_type > self.DATA_TYPE.TENDENCY:
                    raise InvalidPDBError(f"Invalid data type \'{data}\' at {pdb.tell()-1:#x} in {file_path}")
                if data_type == self.DATA_TYPE.PLAY:
                    # Play data
                    data = pdb.read(ctypes.sizeof(PLAY_DATA))
                    play_in_pdb = PLAY_DATA.from_buffer_copy(data)
                    if play_in_pdb.is_valid():
                        # Store the play, but exclude RUNCLOCK and STOPCLOK plays
                        play_name = play_in_pdb.play_name.decode('ASCII')
                        if play_name != 'RUNCLOCK' and play_name != 'STOPCLOK':
                            team_name = play_in_pdb.team_name.decode('ASCII')
                            if play_name in RENAMED_PLAYS:
                                # Change the name to the new name
                                play_name = RENAMED_PLAYS[play_name]
                                play_in_pdb.play_name = play_name.encode('ASCII')
                            play_key = (team_name, play_name)
                            play_type = PLAY_DATA.PLAY_TYPE(play_in_pdb.play_type)
                            # Possible the team used the old name for some weeks and new name for
                            # other weeks, so must check if the play is already in our data
                            if play_key in self.plays[play_type]:
                                original_data = self.plays[play_type][play_key]
                                # Add the original data to the new data
                                play_in_pdb += original_data
                            self.plays[play_type][play_key] = play_in_pdb
                    else:
                        print(f"Skipping invalid play data at {pdb.tell()-ctypes.sizeof(PLAY_DATA):#x}")
                elif data_type == self.DATA_TYPE.TENDENCY:
                    # Tendency data
                    data = pdb.read(ctypes.sizeof(TENDENCY_DATA))
                    tendency_data = TENDENCY_DATA.from_buffer_copy(data)
                    if tendency_data.is_valid():
                        self.tendencies.append(tendency_data)
                    else:
                        print(f"Skipping invalid tendency data at {pdb.tell()-ctypes.sizeof(TENDENCY_DATA):#x}")

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    # Attempts to fix any data that got categorized incorrectly in WinLogStats, such as
    # any pass play data that got stored as rush data, and vice-versa.
    #
    # For example, in WinLogStats most sacks (but not all) get recorded as rush play
    # data and many rush attempts get recorded as pass completions (presumably because
    # a pitch goes slightly forward.)
    def convert_invalid_play_data(self, play_pool):
        play_type_swap = {
            PLAY_DATA.PLAY_TYPE.RUN: PLAY_DATA.PLAY_TYPE.PASS,
            PLAY_DATA.PLAY_TYPE.PASS: PLAY_DATA.PLAY_TYPE.RUN,
        }
        for original_play_type in (PLAY_DATA.PLAY_TYPE.RUN, PLAY_DATA.PLAY_TYPE.PASS):
            # Loop thru the run play data looking for pass plays, and vice-versa
            for play_key in list(self.plays[original_play_type].keys()):
                play_name = play_key[1]
                if ((original_play_type == PLAY_DATA.PLAY_TYPE.RUN and play_pool.is_pass_play(play_name)) or
                        (original_play_type == PLAY_DATA.PLAY_TYPE.PASS and play_pool.is_run_play(play_name))):
                    new_play_type = play_type_swap[original_play_type]
                    # Convert data from original play type to new play type and add it to
                    # any matching play of the new play type
                    original_play = self.plays[original_play_type][play_key]
                    original_play_count = int(original_play.play_count)
                    if play_key in self.plays[new_play_type]:
                        # Found matching play of the new play type (by team and play name)
                        new_play = self.plays[new_play_type][play_key]
                    else:
                        # No matching play data, so create a new record
                        new_play = PLAY_DATA()
                        new_play.play_type = new_play_type
                        new_play.team_name = original_play.team_name
                        new_play.play_name = original_play.play_name
                    # Combine the run and pass data
                    new_play.total_yards += original_play.total_yards
                    new_play.play_count += original_play.play_count
                    if new_play_type == PLAY_DATA.PLAY_TYPE.PASS and original_play.total_yards < 0:
                        # For pass plays, only count negative yardage as sacks
                        if original_play.total_yards <= -original_play_count:
                            # Yardage makes sense for sacks (assuming sacks should average to
                            # at least 1 yard each)
                            new_play.sacks += original_play_count
                        else:
                            # Yardage doesn't quite make sense for just sacks, so recording 2/3 of
                            # the plays as sacks
                            new_play.sacks += round(original_play_count * (2 / 3))
                    if new_play_type == PLAY_DATA.PLAY_TYPE.RUN:
                        # For run plays, count interceptions as fumbles (forward pitch intercepted?)
                        new_play.fumbles += (original_play.fumbles + original_play.interceptions)
                    else:
                        new_play.fumbles += original_play.fumbles
                        new_play.interceptions += original_play.interceptions
                    new_play.touchdowns_offense += original_play.touchdowns_offense
                    new_play.touchdowns_defense += original_play.touchdowns_defense
                    new_play.unknown1 += original_play.unknown1
                    new_play.unknown2 += original_play.unknown2
                    new_play.points_scored += original_play.points_scored
                    # Add/Replace the play in the new play type data
                    self.plays[new_play_type][play_key] = new_play
                    # Remove the play from the original play type data
                    self.plays[original_play_type].pop(play_key)


###############################################################################
#
# PlayInPlan - A class for storing play data as it exists in a game plan file
#
###############################################################################

class PlayInPlan:
    def __init__(self):
        self.slot = -1
        self.stock_flag = 0
        self.play_category = 0
        self.special_flag = 0
        self.special_category = 0
        self.user_category = 0
        self.filename = ''
        self.play_name = ''
        self.offset = 0
        self.size = 0

    def get_name(self):
        if self.filename:
            return Path(self.filename).stem
        return self.play_name


###############################################################################
#
# PLN - A class for opening and parsing an FBPro98 game plan
#
###############################################################################

class PLN:
    # Constants
    NUMBER_NORMAL_PLAYS = 64
    NUMBER_SPECIAL_PLAYS = 10
    NUMBER_STOCK_SPECIAL_PLAYS = 10
    NUMBER_PLAY_SLOTS = NUMBER_NORMAL_PLAYS + NUMBER_SPECIAL_PLAYS + NUMBER_STOCK_SPECIAL_PLAYS
    G95_HEADER_SIZE = 12
    G95_OFFSETS_TABLE_SIZE = NUMBER_PLAY_SLOTS * 2

    class CHUNK_ID(str, Enum):
        G95 = "G95:"
        J95 = "J95:"
        S98 = "S98:"

    ###########################################################################
    #
    # Initialization/Destruction
    #
    ###########################################################################

    def __init__(self, filename):
        self.filename = filename
        self.normal_plays = {}
        self.special_plays = {}
        self.stock_special_plays = {}
        self.plays_by_slot = {}

        file_path = Path(self.filename)
        with open(file_path, 'rb') as pln:
            buffer = pln.read()
            if len(buffer) < self.G95_HEADER_SIZE + self.G95_OFFSETS_TABLE_SIZE:
                raise InvalidPLNError(f"File too small to contain PLN header and offsets table in {file_path}")

            header = struct.unpack_from('<4sIBBBB', buffer, 0)
            self.id = header[0].decode('ASCII')
            self.g95_size = header[1]
            self.unknown1 = header[2]
            self.unknown2 = header[3]
            self.unknown3 = header[4]
            self.unknown4 = header[5]
            if self.id != self.CHUNK_ID.G95:
                raise InvalidPLNError(f"Invalid header \'{self.id}\' at {0:#x} in {file_path}")
            if (self.unknown1 != 0 or self.unknown2 != 1 or self.unknown3 != 2 or self.unknown4 != 3):
                print(f"Unknown data in header at {4:#X}: {self.unknown1:02x} {self.unknown2:02x} {self.unknown3:02x} {self.unknown4:02x}")

            g95_end = 8 + self.g95_size
            if g95_end > len(buffer):
                raise InvalidPLNError(f"G95 chunk extends past end of file in {file_path}")

            offsets = struct.unpack_from(f'<{self.NUMBER_PLAY_SLOTS}H', buffer, self.G95_HEADER_SIZE)

            for play_index, relative_offset in enumerate(offsets):
                if relative_offset == 0:
                    continue

                record_offset = self.G95_HEADER_SIZE + relative_offset
                if record_offset < self.G95_HEADER_SIZE + self.G95_OFFSETS_TABLE_SIZE or record_offset >= g95_end:
                    raise InvalidPLNError(
                        f"Play offset {relative_offset:#x} for slot {play_index} is out of range in {file_path}"
                    )

                play = self._parse_play(buffer, record_offset, play_index, file_path)
                self.plays_by_slot[play.slot] = play
                self._store_play(play)

    def _parse_play(self, buffer, buffer_offset, play_index, file_path):
        if buffer_offset + 4 > len(buffer):
            raise InvalidPLNError(f"Truncated play header at {buffer_offset:#x} in {file_path}")

        data = struct.unpack_from('<BBBB', buffer, buffer_offset)
        buffer_offset += 4

        play = PlayInPlan()
        play.slot = play_index
        play.stock_flag = data[0]
        play.play_category = data[1]
        play.special_flag = data[2]
        play.special_category = data[2]
        play.user_category = data[3]

        if play.stock_flag == 0:
            play.filename, _ = self._read_c_string(buffer, buffer_offset, file_path)
        elif play.stock_flag == 1:
            if buffer_offset + 16 > len(buffer):
                raise InvalidPLNError(f"Truncated stock play record at {buffer_offset:#x} in {file_path}")
            data = struct.unpack_from('<8sII', buffer, buffer_offset)
            play.play_name = data[0].decode('ASCII', errors='replace').rstrip('\x00 ')
            play.offset = data[1]
            play.size = data[2]
        else:
            raise InvalidPLNError(f"Invalid stock flag {play.stock_flag:#x} at slot {play_index} in {file_path}")

        return play

    def _read_c_string(self, buffer, buffer_offset, file_path):
        string_end = buffer.find(b'\x00', buffer_offset)
        if string_end == -1:
            raise InvalidPLNError(f"Missing null terminator for play record at {buffer_offset:#x} in {file_path}")
        value = buffer[buffer_offset:string_end].decode('ASCII', errors='replace')
        return value, string_end + 1

    def _store_play(self, play):
        if play.slot < self.NUMBER_NORMAL_PLAYS:
            self.normal_plays[play.get_name()] = play
        elif play.slot < self.NUMBER_NORMAL_PLAYS + self.NUMBER_SPECIAL_PLAYS:
            self.special_plays[play.get_name()] = play
        else:
            self.stock_special_plays[play.get_name()] = play


###############################################################################
#
# InvalidPLNError - A custom exception
#
###############################################################################

class InvalidPLNError(BaseException):
    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


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


###############################################################################
#
# ExcelPdbWorkbook - A class for creating the Excel workbook
#
###############################################################################

class _LegacyExcelPdbWorkbook:

    ###########################################################################
    #
    # Initialization/Destruction
    #
    ###########################################################################

    def __init__(self, filename, perform_calculations):
        self.filename = filename
        self.perform_calculations = perform_calculations

    def __enter__(self):
        config = get_config()
        filepath = Path(self.filename)
        filepath.parent.mkdir(exist_ok=True)

        self.workbook = xlsxwriter.Workbook(filepath)
        file_extension = filepath.suffix
        if file_extension == '.xlsm':
            # Add the VBA code to the workbook
            if config['Settings']['CalculateCategoryStats']:
                self.workbook.add_vba_project(str(get_runtime_path("vbaProject_categories.bin")))
            else:
                self.workbook.add_vba_project(str(get_runtime_path("vbaProject.bin")))

        text_format = self.workbook.add_format({'num_format': '@'})
        avg_format = self.workbook.add_format({'num_format': '0.0'})
        percent_format_0 = self.workbook.add_format({'num_format': '0%'})
        percent_format_1 = self.workbook.add_format({'num_format': '0.0%'})
        center_format = self.workbook.add_format()
        center_format.set_align('center')

        # Setup the 'Run Plays' worksheet
        self.run_worksheet = self.workbook.add_worksheet('Run Plays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'Attempts', 'Yards', 'Avg', 'Fumbles', 'TD', 'Pts']
        self.run_worksheet.set_column_pixels(0, 0, 120)         # Team
        self.run_worksheet.set_column_pixels(1, 1, 80)          # Category
        self.run_worksheet.set_column(2, 2, 6.00, text_format)  # Slot (in game plan)
        self.run_worksheet.set_column_pixels(3, 3, 100)         # Play
        self.run_worksheet.set_column_pixels(4, 4, 59)          # Type
        self.run_worksheet.set_column(7, 7, None, avg_format)   # Average
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['RunFumblePercentage']:
                column_index = 9
                header.insert(column_index, 'Fumble %')
                self.run_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['RunTouchdownPercentage']:
                column_index = 10 + extra_columns
                header.insert(column_index, 'TD %')
                self.run_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.run_worksheet.write_row(0, 0, header)
        self.run_columns = len(header)
        self.run_rows = 1
        self.run_worksheet.freeze_panes(1, 0)
        self.run_worksheet.autofilter('A1:C1')  # pyright: ignore[reportCallIssue]

        # Setup the 'Pass Plays' worksheet
        self.pass_worksheet = self.workbook.add_worksheet('Pass Plays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'Cmp', 'Att', 'Comp %', 'Yards', 'Y/Comp', 'Y/Att', 'Fumbles', 'Int', 'Sacks', 'TD', 'Pts']
        self.pass_worksheet.set_column_pixels(0, 0, 120)              # Team
        self.pass_worksheet.set_column_pixels(1, 1, 80)               # Category
        self.pass_worksheet.set_column(2, 2, 6.00, text_format)       # Slot (in game plan)
        self.pass_worksheet.set_column_pixels(3, 3, 100)              # Play
        self.pass_worksheet.set_column_pixels(4, 4, 49)               # Type
        self.pass_worksheet.set_column(5, 5, 40)                      # Completions
        self.pass_worksheet.set_column(6, 6, 40)                      # Attempt
        self.pass_worksheet.set_column(7, 7, None, percent_format_0)  # Completion percentage
        self.pass_worksheet.set_column(9, 9, None, avg_format)        # Yards/Completion
        self.pass_worksheet.set_column(10, 10, None, avg_format)      # Yards/Attempt
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['PassInterceptionPercentage']:
                column_index = 13
                header.insert(column_index, 'Int %')
                self.pass_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['PassSackPercentage']:
                column_index = 14 + extra_columns
                header.insert(column_index, 'Sack %')
                self.pass_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['PassTouchdownPercentage']:
                column_index = 15 + extra_columns
                header.insert(column_index, 'TD %')
                self.pass_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.pass_worksheet.write_row(0, 0, header)
        self.pass_rows = 1
        self.pass_columns = len(header)
        self.pass_worksheet.freeze_panes(1, 0)
        self.pass_worksheet.autofilter('A1:C1')     # pyright: ignore[reportCallIssue]

        # Setup the 'Def Plays' worksheet
        self.def_worksheet = self.workbook.add_worksheet('Def Plays')
        header = ['Team', 'Category', 'Slot', 'Play', 'Type', 'vs Run', 'Avg', 'vs Pass', 'Avg', 'Fumbles', 'Int', 'Sacks', 'TD/Def', 'TD/Off']
        self.def_worksheet.set_column_pixels(0, 0, 120)         # Team
        self.def_worksheet.set_column_pixels(1, 1, 126)         # Category
        self.def_worksheet.set_column(2, 2, 6.00, text_format)  # Slot (in game plan)
        self.def_worksheet.set_column_pixels(3, 3, 100)         # Play
        self.def_worksheet.set_column_pixels(4, 4, 37)          # Type
        self.def_worksheet.set_column(5, 5, None, text_format)  # vs Run
        self.def_worksheet.set_column(6, 6, None, avg_format)   # Avg
        self.def_worksheet.set_column(7, 7, None, text_format)  # vs Pass
        self.def_worksheet.set_column(8, 8, None, avg_format)   # Avg
        if self.perform_calculations:
            extra_columns = 0
            if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                header.insert(11, 'TO %')
                self.def_worksheet.set_column(11, 11, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['DefenseSackPercentage']:
                column_index = 12 + extra_columns
                header.insert(column_index, 'Sack %')
                self.def_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
            if config['AdditionalColumns']['DefenseOffTdPercentage']:
                column_index = 14 + extra_columns
                header.insert(column_index, 'TD/Off %')
                self.def_worksheet.set_column(column_index, column_index, None, percent_format_1)
                extra_columns += 1
        self.def_worksheet.write_row(0, 0, header)
        self.def_rows = 1
        self.def_columns = len(header)
        self.def_worksheet.freeze_panes(1, 0)
        self.def_worksheet.autofilter('A1:C1')      # pyright: ignore[reportCallIssue]

        # Setup the 'Tendencies' worksheet
        self.tendencies_worksheet = self.workbook.add_worksheet('Tendencies')
        header = ['Team', 'Situation', 'Runs', 'Passes']
        self.tendencies_worksheet.set_column_pixels(0, 0, 120)  # Team
        self.tendencies_worksheet.set_column_pixels(1, 1, 140)  # Situation
        self.tendencies_worksheet.write_row(0, 0, header)
        self.tendencies_rows = 1
        self.tendencies_columns = len(header)
        self.tendencies_worksheet.freeze_panes(1, 0)
        self.tendencies_worksheet.autofilter('A1:A1')       # pyright: ignore[reportCallIssue]

        if config['Settings']['CalculateCategoryStats']:
            # Setup the 'Run Categories' worksheet
            self.run_categories_worksheet = self.workbook.add_worksheet('Run Categories')
            header = ['Team', 'Category', 'Attempts', 'Yards', 'Avg', 'Fumbles', 'TD', 'Pts']
            self.run_categories_worksheet.set_column_pixels(0, 0, 120)         # Team
            self.run_categories_worksheet.set_column_pixels(1, 1, 80)          # Category
            self.run_categories_worksheet.set_column(4, 4, None, avg_format)   # Average
            self.run_categories_worksheet.write_row(0, 0, header)
            self.run_categories_columns = len(header)
            self.run_categories_rows = 1
            self.run_categories_worksheet.freeze_panes(1, 0)
            self.run_categories_worksheet.autofilter('A1:B1')       # pyright: ignore[reportCallIssue]
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['RunFumblePercentage']:
                    column_index = 6
                    header.insert(column_index, 'Fumble %')
                    self.run_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['RunTouchdownPercentage']:
                    column_index = 7 + extra_columns
                    header.insert(column_index, 'TD %')
                    self.run_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1

            # Setup the 'Pass Categories' worksheet
            self.pass_categories_worksheet = self.workbook.add_worksheet('Pass Categories')
            header = ['Team', 'Category', 'Cmp', 'Att', 'Comp %', 'Yards', 'Y/Comp', 'Y/Att', 'Fumbles', 'Int', 'Sacks', 'TD', 'Pts']
            self.pass_categories_worksheet.set_column_pixels(0, 0, 120)              # Team
            self.pass_categories_worksheet.set_column_pixels(1, 1, 80)               # Category
            self.pass_categories_worksheet.set_column(2, 2, 40)                      # Completions
            self.pass_categories_worksheet.set_column(3, 3, 40)                      # Attempts
            self.pass_categories_worksheet.set_column(4, 4, None, percent_format_0)  # Completion percentage
            self.pass_categories_worksheet.set_column(6, 6, None, avg_format)        # Yards/Completion
            self.pass_categories_worksheet.set_column(7, 7, None, avg_format)        # Yards/Attempt
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['PassInterceptionPercentage']:
                    column_index = 10
                    header.insert(column_index, 'Int %')
                    self.pass_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['PassSackPercentage']:
                    column_index = 11 + extra_columns
                    header.insert(column_index, 'Sack %')
                    self.pass_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['PassTouchdownPercentage']:
                    column_index = 12 + extra_columns
                    header.insert(column_index, 'TD %')
                    self.pass_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
            self.pass_categories_worksheet.write_row(0, 0, header)
            self.pass_categories_rows = 1
            self.pass_categories_columns = len(header)
            self.pass_categories_worksheet.freeze_panes(1, 0)
            self.pass_categories_worksheet.autofilter('A1:B1')      # pyright: ignore[reportCallIssue]

            # Setup the 'Def Categories' worksheet
            self.def_categories_worksheet = self.workbook.add_worksheet('Def Categories')
            header = ['Team', 'Category', 'vs Run', 'Avg', 'vs Pass', 'Avg', 'Fumbles', 'Int', 'Sacks', 'TD/Def', 'TD/Off']
            self.def_categories_worksheet.set_column_pixels(0, 0, 120)         # Team
            self.def_categories_worksheet.set_column_pixels(1, 1, 126)         # Category
            self.def_categories_worksheet.set_column(2, 2, None, text_format)  # vs Run
            self.def_categories_worksheet.set_column(3, 3, None, avg_format)   # Avg
            self.def_categories_worksheet.set_column(4, 4, None, text_format)  # vs Pass
            self.def_categories_worksheet.set_column(5, 5, None, avg_format)   # Avg
            if self.perform_calculations:
                extra_columns = 0
                if config['AdditionalColumns']['DefenseTurnoverPercentage']:
                    column_index = 8
                    header.insert(column_index, 'TO %')
                    self.def_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['DefenseSackPercentage']:
                    column_index = 9 + extra_columns
                    header.insert(column_index, 'Sack %')
                    self.def_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
                if config['AdditionalColumns']['DefenseOffTdPercentage']:
                    column_index = 11 + extra_columns
                    header.insert(column_index, 'TD/Off %')
                    self.def_categories_worksheet.set_column(column_index, column_index, None, percent_format_1)
                    extra_columns += 1
            self.def_categories_worksheet.write_row(0, 0, header)
            self.def_categories_rows = 1
            self.def_categories_columns = len(header)
            self.def_categories_worksheet.freeze_panes(1, 0)
            self.def_categories_worksheet.autofilter('A1:B1')       # pyright: ignore[reportCallIssue]
        return self

    def __exit__(self, *args):
        if self.workbook:
            self._add_conditional_format()
            self.workbook.close()

    ###########################################################################
    #
    # Public API.
    #
    ###########################################################################

    def add_play(self, play_data, play_slot, play_attributes):
        if play_data.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            self._add_run_play(play_data, play_slot, play_attributes)
        elif play_data.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            self._add_pass_play(play_data, play_slot, play_attributes)
        elif play_data.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            self._add_defense_play(play_data, play_slot, play_attributes)

    def add_category(self, team_category, category_data):
        if category_data.play_type == PLAY_DATA.PLAY_TYPE.RUN:
            self._add_run_category(team_category, category_data)
        elif category_data.play_type == PLAY_DATA.PLAY_TYPE.PASS:
            self._add_pass_category(team_category, category_data)
        elif category_data.play_type == PLAY_DATA.PLAY_TYPE.DEFENSE:
            self._add_defense_category(team_category, category_data)

    def add_tendency(self, t):
        # First down
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

        # Second down
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

        # Third down
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

        # Fourth down
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

    ###########################################################################
    #
    # Private API.
    #
    ###########################################################################

    def _add_run_play(self, play_data: PLAY_DATA, play_slot, play_attributes):
        play_name = play_data.play_name.decode('ASCII')
        type = ''
        if play_name[1] == '1' or play_name[2] == '1':      # Account for 'S10snkG'
            type = 'QB draw'

        if play_data.play_count > 0:
            avg = int(play_data.total_yards) / int(play_data.play_count)
        else:
            avg = 0.0

        row_data = [play_data.team_name.decode('ASCII'),
                    play_attributes.category,
                    play_slot,
                    play_name,
                    type,
                    play_data.play_count,
                    play_data.total_yards,
                    round(avg, 1),
                    play_data.fumbles,
                    play_data.touchdowns_offense,
                    play_data.touchdowns_offense * 6]

        if self.perform_calculations:
            config = get_config()
            extra_columns = 0
            if config['AdditionalColumns']['RunFumblePercentage']:
                row_data.insert(9, round(int(play_data.fumbles) / int(play_data.play_count), 3))
                extra_columns += 1
            if config['AdditionalColumns']['RunTouchdownPercentage']:
                row_data.insert(10 + extra_columns, round(int(play_data.touchdowns_offense) / int(play_data.play_count), 3))
                extra_columns += 1

        self.run_worksheet.write_row(self.run_rows, 0, row_data)
        self.run_rows += 1

    def _add_pass_play(self, play_data, play_slot, play_attributes):
        type = ''
        if play_attributes.screen:
            type = 'Screen'

        if play_data.completions > 0:
            avgPerCompletion = int(play_data.total_yards) / int(play_data.completions)
        else:
            avgPerCompletion = 0.0

        if play_data.play_count > 0:
            avgPerAttempt = int(play_data.total_yards) / int(play_data.play_count)
        else:
            avgPerAttempt = 0.0

        row_data = [play_data.team_name.decode('ASCII'),
                    play_attributes.category,
                    play_slot,
                    play_data.play_name.decode('ASCII'),
                    type,
                    play_data.completions,
                    play_data.play_count,
                    round(int(play_data.completions) / int(play_data.play_count), 2),
                    play_data.total_yards,
                    round(avgPerCompletion, 1),
                    round(avgPerAttempt, 1),
                    play_data.fumbles,
                    play_data.interceptions,
                    play_data.sacks,
                    play_data.touchdowns_offense,
                    play_data.touchdowns_offense * 6]

        if self.perform_calculations:
            config = get_config()
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

    def _add_defense_play(self, play_data, play_slot, play_attributes):
        type = ''
        if play_attributes.three_four:
            type = '3-4'
        elif play_attributes.four_three:
            type = '4-3'
        elif play_attributes.run_and_shoot:
            type = 'R&S'

        if play_data.run_plays_against > 0:
            rush_avg = int(play_data.rush_yards_allowed) / int(play_data.run_plays_against)
        else:
            rush_avg = 0.0

        if play_data.pass_plays_against > 0:
            pass_avg = int(play_data.pass_yards_allowed) / int(play_data.pass_plays_against)
        else:
            pass_avg = 0.0

        row_data = [play_data.team_name.decode('ASCII'),
                    play_attributes.category,
                    play_slot,
                    play_data.play_name.decode('ASCII'),
                    type,
                    f"{str(play_data.run_plays_against)}/{str(play_data.rush_yards_allowed)}",
                    round(rush_avg, 1),
                    f"{str(play_data.pass_plays_against)}/{str(play_data.pass_yards_allowed)}",
                    round(pass_avg, 1),
                    play_data.fumbles,
                    play_data.interceptions,
                    play_data.sacks,
                    play_data.touchdowns_defense,
                    play_data.touchdowns_offense]

        if self.perform_calculations:
            config = get_config()
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

        row_data = [team_category[0],
                    team_category[1],
                    category_data.play_count,
                    category_data.total_yards,
                    round(avg, 1),
                    category_data.fumbles,
                    category_data.touchdowns_offense,
                    category_data.touchdowns_offense * 6]

        if self.perform_calculations:
            config = get_config()
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
            avgPerCompletion = int(category_data.total_yards) / int(category_data.completions)
        else:
            avgPerCompletion = 0.0

        if category_data.play_count > 0:
            avgPerAttempt = int(category_data.total_yards) / int(category_data.play_count)
        else:
            avgPerAttempt = 0.0

        row_data = [team_category[0],
                    team_category[1],
                    category_data.completions,
                    category_data.play_count,
                    round(int(category_data.completions) / int(category_data.play_count), 2),
                    category_data.total_yards,
                    round(avgPerCompletion, 1),
                    round(avgPerAttempt, 1),
                    category_data.fumbles,
                    category_data.interceptions,
                    category_data.sacks,
                    category_data.touchdowns_offense,
                    category_data.touchdowns_offense * 6]

        if self.perform_calculations:
            config = get_config()
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

        row_data = [team_category[0],
                    team_category[1],
                    f"{str(category_data.run_plays_against)}/{str(category_data.rush_yards_allowed)}",
                    round(rush_avg, 1),
                    f"{str(category_data.pass_plays_against)}/{str(category_data.pass_yards_allowed)}",
                    round(pass_avg, 1),
                    category_data.fumbles,
                    category_data.interceptions,
                    category_data.sacks,
                    category_data.touchdowns_defense,
                    category_data.touchdowns_offense]

        if self.perform_calculations:
            config = get_config()
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
        config = get_config()
        # Apply conditional format so user can tell which row is selected when Excel loses focus
        border_format = self.workbook.add_format()
        border_format.set_border_color('#0000FF')
        border_format.set_top(1)
        border_format.set_bottom(1)

        self.run_worksheet.conditional_format(1, 0, self.run_rows, self.run_columns - 1,
                                              {'type':     'formula',
                                               'criteria': '=CELL("row")=ROW()',
                                               'format':   border_format
                                               })
        self.pass_worksheet.conditional_format(1, 0, self.pass_rows, self.pass_columns - 1,
                                               {'type':     'formula',
                                                'criteria': '=CELL("row")=ROW()',
                                                'format':   border_format
                                                })
        self.def_worksheet.conditional_format(1, 0, self.def_rows, self.def_columns - 1,
                                              {'type':     'formula',
                                               'criteria': '=CELL("row")=ROW()',
                                               'format':   border_format
                                               })
        self.tendencies_worksheet.conditional_format(1, 0, self.tendencies_rows, self.tendencies_columns - 1,
                                                     {'type':     'formula',
                                                      'criteria': '=CELL("row")=ROW()',
                                                      'format':   border_format
                                                      })

        if config['Settings']['CalculateCategoryStats']:
            self.run_categories_worksheet.conditional_format(1, 0, self.run_categories_rows, self.run_categories_columns - 1,
                                                             {'type':     'formula',
                                                              'criteria': '=CELL("row")=ROW()',
                                                              'format':   border_format
                                                              })
            self.pass_categories_worksheet.conditional_format(1, 0, self.pass_categories_rows, self.pass_categories_columns - 1,
                                                              {'type':     'formula',
                                                               'criteria': '=CELL("row")=ROW()',
                                                               'format':   border_format
                                                               })
            self.def_categories_worksheet.conditional_format(1, 0, self.def_categories_rows, self.def_categories_columns - 1,
                                                             {'type':     'formula',
                                                              'criteria': '=CELL("row")=ROW()',
                                                              'format':   border_format
                                                              })


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
