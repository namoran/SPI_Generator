# Copyright (c) 2010-2024 openpyxl

"""
Represent a cell's coordinate in any collection
"""

from openpyxl.utils.cell import get_column_letter, coordinate_to_tuple

class Coordinate:

    __slots__ = ("row", "column")


    def __init__(self, row, column):
        self.row = row
        self.column = column


    def __str__(self):
        return f"{get_column_letter(self.column)}{self.row}"


    @property
    def column_letter(self):
        return get_column_letter(self.column)


    @classmethod
    def from_string(cls, coord):
        row, column = coordinate_to_tuple(coord)
        return cls(row, column)


    def __eq__(self, other):
        return self.row == other.row and self.column == other.column


    def __ne__(self, other):
        return not self == other
