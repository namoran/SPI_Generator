# Copyright (c) 2010-2024 openpyxl

import pytest

@pytest.fixture
def Coordinate():

    from ..coordinate import Coordinate
    return Coordinate


class TestCoordinate:


    def test_ctor(self, Coordinate):
        coord = Coordinate(1, 3)
        assert coord.row == 1 and coord.column == 3


    def test_str(self, Coordinate):
        coord = Coordinate(9, 8)
        assert str(coord) == "H9"


    def test_column_letter(self, Coordinate):
        coord = Coordinate(3, 4)
        assert coord.column_letter == "D"


    def test_from_string(self, Coordinate):
        coord = Coordinate.from_string("B6")
        assert coord.row == 6 and coord.column == 2


    def test_eq(self, Coordinate):
        c1 = Coordinate(1, 6)
        c2 = Coordinate(1, 6)
        assert c1 == c2


    def test_ne(self, Coordinate):
        c1 = Coordinate(1, 1)
        c2 = Coordinate(1, 6)
        assert c1 != c2
