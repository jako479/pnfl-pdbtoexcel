from __future__ import annotations

import ctypes
import logging
from enum import IntEnum
from pathlib import Path

logger = logging.getLogger(__name__)


RENAMED_PLAYS = {
    "WR47PT01": "WR27PT01",
    "WR48PT01": "WR28PT01",
}


class PLAY_DATA(ctypes.LittleEndianStructure):
    class PLAY_TYPE(IntEnum):
        RUN = 0
        PASS = 1
        SPECIAL = 2
        DEFENSE = 3
        ONSIDE = 5

    _fields_ = [
        ("play_type", ctypes.c_uint32),
        ("team_name", ctypes.c_char * 64),
        ("play_name", ctypes.c_char * 128),
        ("total_yards", ctypes.c_int32),
        ("play_count", ctypes.c_uint32),
        ("completions", ctypes.c_uint32),
        ("sacks", ctypes.c_uint32),
        ("fumbles", ctypes.c_uint32),
        ("interceptions", ctypes.c_uint32),
        ("touchdowns_offense", ctypes.c_uint32),
        ("touchdowns_defense", ctypes.c_uint32),
        ("unknown1", ctypes.c_int32),
        ("unknown2", ctypes.c_int32),
        ("points_scored", ctypes.c_uint32),
        ("run_plays_against", ctypes.c_uint32),
        ("pass_plays_against", ctypes.c_uint32),
        ("rush_yards_allowed", ctypes.c_int32),
        ("pass_yards_allowed", ctypes.c_int32),
    ]

    def is_valid(self):
        if (
            (
                self.play_type == self.PLAY_TYPE.RUN
                or self.play_type == self.PLAY_TYPE.PASS
                or self.play_type == self.PLAY_TYPE.SPECIAL
                or self.play_type == self.PLAY_TYPE.DEFENSE
                or self.play_type == self.PLAY_TYPE.ONSIDE
            )
            and len(self.team_name) > 0
            and len(self.play_name) > 0
        ):
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


class DOWN_DATA(ctypes.LittleEndianStructure):
    _fields_ = [
        ("first_down", ctypes.c_uint32),
        ("second_down", ctypes.c_uint32),
        ("third_down", ctypes.c_uint32),
        ("fourth_down", ctypes.c_uint32),
    ]


class TENDENCY_DATA(ctypes.LittleEndianStructure):
    _fields_ = [
        ("team_name", ctypes.c_char * 64),
        ("run_zero_to_one", DOWN_DATA),
        ("pass_zero_to_one", DOWN_DATA),
        ("run_two_to_five", DOWN_DATA),
        ("pass_two_to_five", DOWN_DATA),
        ("run_six_to_ten", DOWN_DATA),
        ("pass_six_to_ten", DOWN_DATA),
        ("run_ten_plus", DOWN_DATA),
        ("pass_ten_plus", DOWN_DATA),
    ]

    def is_valid(self):
        return len(self.team_name) > 0


class InvalidPDBError(BaseException):
    def __init__(self, msg):
        self.msg = msg

    def __str__(self):
        return self.msg


class PDB:
    class DATA_TYPE(IntEnum):
        PLAY = 0
        TENDENCY = 1

    def __init__(self, filename):
        self.filename = filename
        self.plays = {
            PLAY_DATA.PLAY_TYPE.RUN: {},
            PLAY_DATA.PLAY_TYPE.PASS: {},
            PLAY_DATA.PLAY_TYPE.SPECIAL: {},
            PLAY_DATA.PLAY_TYPE.DEFENSE: {},
            PLAY_DATA.PLAY_TYPE.ONSIDE: {},
        }
        self.tendencies = []

        file_path = Path(self.filename)

        with open(file_path, "rb") as pdb:
            while True:
                data = pdb.read(1)
                if not data:
                    self.tendencies.sort(key=lambda x: x.team_name)
                    break
                data_type = int.from_bytes(data, byteorder="little")
                if data_type < self.DATA_TYPE.PLAY or data_type > self.DATA_TYPE.TENDENCY:
                    raise InvalidPDBError(
                        f"Invalid data type '{data}' at {pdb.tell()-1:#x} in {file_path}"
                    )
                if data_type == self.DATA_TYPE.PLAY:
                    data = pdb.read(ctypes.sizeof(PLAY_DATA))
                    play_in_pdb = PLAY_DATA.from_buffer_copy(data)
                    if play_in_pdb.is_valid():
                        play_name = play_in_pdb.play_name.decode("ASCII")
                        if play_name != "RUNCLOCK" and play_name != "STOPCLOK":
                            team_name = play_in_pdb.team_name.decode("ASCII")
                            if play_name in RENAMED_PLAYS:
                                play_name = RENAMED_PLAYS[play_name]
                                play_in_pdb.play_name = play_name.encode("ASCII")
                            play_key = (team_name, play_name)
                            play_type = PLAY_DATA.PLAY_TYPE(play_in_pdb.play_type)
                            if play_key in self.plays[play_type]:
                                original_data = self.plays[play_type][play_key]
                                play_in_pdb += original_data
                            self.plays[play_type][play_key] = play_in_pdb
                    else:
                        logger.warning(
                            "Skipping invalid play data at %#x",
                            pdb.tell() - ctypes.sizeof(PLAY_DATA),
                        )
                elif data_type == self.DATA_TYPE.TENDENCY:
                    data = pdb.read(ctypes.sizeof(TENDENCY_DATA))
                    tendency_data = TENDENCY_DATA.from_buffer_copy(data)
                    if tendency_data.is_valid():
                        self.tendencies.append(tendency_data)
                    else:
                        logger.warning(
                            "Skipping invalid tendency data at %#x",
                            pdb.tell() - ctypes.sizeof(TENDENCY_DATA),
                        )

    def convert_invalid_play_data(self, play_pool):
        from pnfl_playpool import OffensivePlayRecord

        play_type_swap = {
            PLAY_DATA.PLAY_TYPE.RUN: PLAY_DATA.PLAY_TYPE.PASS,
            PLAY_DATA.PLAY_TYPE.PASS: PLAY_DATA.PLAY_TYPE.RUN,
        }
        for original_play_type in (PLAY_DATA.PLAY_TYPE.RUN, PLAY_DATA.PLAY_TYPE.PASS):
            for play_key in list(self.plays[original_play_type].keys()):
                play_name = play_key[1]
                record = play_pool.find_by_name(play_name)
                if not isinstance(record, OffensivePlayRecord):
                    continue
                if (
                    (original_play_type == PLAY_DATA.PLAY_TYPE.RUN and record.is_pass)
                    or (original_play_type == PLAY_DATA.PLAY_TYPE.PASS and record.is_run)
                ):
                    new_play_type = play_type_swap[original_play_type]
                    original_play = self.plays[original_play_type][play_key]
                    original_play_count = int(original_play.play_count)
                    if play_key in self.plays[new_play_type]:
                        new_play = self.plays[new_play_type][play_key]
                    else:
                        new_play = PLAY_DATA()
                        new_play.play_type = new_play_type
                        new_play.team_name = original_play.team_name
                        new_play.play_name = original_play.play_name
                    new_play.total_yards += original_play.total_yards
                    new_play.play_count += original_play.play_count
                    if new_play_type == PLAY_DATA.PLAY_TYPE.PASS and original_play.total_yards < 0:
                        if original_play.total_yards <= -original_play_count:
                            new_play.sacks += original_play_count
                        else:
                            new_play.sacks += round(original_play_count * (2 / 3))
                    if new_play_type == PLAY_DATA.PLAY_TYPE.RUN:
                        new_play.fumbles += original_play.fumbles + original_play.interceptions
                    else:
                        new_play.fumbles += original_play.fumbles
                        new_play.interceptions += original_play.interceptions
                    new_play.touchdowns_offense += original_play.touchdowns_offense
                    new_play.touchdowns_defense += original_play.touchdowns_defense
                    new_play.unknown1 += original_play.unknown1
                    new_play.unknown2 += original_play.unknown2
                    new_play.points_scored += original_play.points_scored
                    self.plays[new_play_type][play_key] = new_play
                    self.plays[original_play_type].pop(play_key)


__all__ = [
    "DOWN_DATA",
    "InvalidPDBError",
    "PDB",
    "PLAY_DATA",
    "RENAMED_PLAYS",
    "TENDENCY_DATA",
]
