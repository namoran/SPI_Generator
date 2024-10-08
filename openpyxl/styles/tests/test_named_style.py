# Copyright (c) 2010-2024 openpyxl

import pytest

from array import array

from ..fonts import Font
from ..borders import Border
from ..fills import PatternFill
from ..alignment import Alignment
from ..protection import Protection
from ..cell_style import CellStyle, StyleArray

from openpyxl import Workbook

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def NamedStyle():
    from ..named_styles import NamedStyle
    return NamedStyle


class TestNamedStyle:

    def test_ctor(self, NamedStyle):
        style = NamedStyle()

        assert style.font == Font()
        assert style.border == Border()
        assert style.fill == PatternFill()
        assert style.protection == Protection()
        assert style.alignment == Alignment()
        assert style.number_format == "General"
        assert style._wb is None


    def test_dict(self, NamedStyle):
        style = NamedStyle()
        assert dict(style) == {'name':'Normal', 'hidden':'0', }


    def test_bind(self, NamedStyle):
        style = NamedStyle()

        wb = Workbook()
        style.bind(wb)

        assert style._wb is wb


    def test_as_tuple(self, NamedStyle):
        style = NamedStyle()
        assert style.as_tuple() == array('i', (0, 0, 0, 0, 0, 0, 0, 0, 0))


    def test_as_xf(self, NamedStyle):
        style = NamedStyle()
        style.alignment = Alignment(horizontal="left")

        xf = style.as_xf()
        assert xf == CellStyle(numFmtId=0, fontId=0, fillId=0, borderId=0,
                              applyNumberFormat=None,
                              applyFont=None,
                              applyFill=None,
                              applyBorder=None,
                              applyAlignment=True,
                              applyProtection=None,
                              alignment=Alignment(horizontal="left"),
                              protection=None,
                              )


    def test_as_name(self, NamedStyle, _NamedCellStyle):
        style = NamedStyle()

        name = style.as_name()
        assert name == _NamedCellStyle(name='Normal', xfId=0, hidden=False)


    @pytest.mark.parametrize("attr, key, collection, expected",
                             [
                                 ('font', 'fontId', '_fonts', 0),
                                 ('fill', 'fillId', '_fills', 0),
                                 ('border', 'borderId', '_borders', 0),
                                 ('alignment', 'alignmentId', '_alignments', 0),
                                 ('protection', 'protectionId', '_protections', 0),
                                 ('number_format', 'numFmtId', '_number_formats', 164),
                             ]
                             )
    def test_recalculate(self, NamedStyle, attr, key, collection, expected):
        style = NamedStyle()
        wb = Workbook()
        wb._number_formats.append("###")
        style.bind(wb)
        style._style = StyleArray([1, 1, 1, 1, 1, 1, 1, 1, 1])

        obj = getattr(wb, collection)[0]
        setattr(style, attr, obj)
        assert getattr(style._style, key) == expected


    def test_no_mutable_defaults(self, NamedStyle):
        ns1 = NamedStyle()
        ns2 = NamedStyle()
        for attr in ("font", "fill", "border", "alignment", "protection"):
            assert getattr(ns1, attr) is not getattr(ns2, attr)


@pytest.fixture
def _NamedCellStyle():
    from ..named_styles import _NamedCellStyle
    return _NamedCellStyle


class TestNamedCellStyle:

    def test_ctor(self, _NamedCellStyle):
        named_style = _NamedCellStyle(xfId=0, name="Normal", builtinId=0)
        xml = tostring(named_style.to_tree())
        expected = """
        <cellStyle name="Normal" xfId="0" builtinId="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, _NamedCellStyle):
        src = """
        <cellStyle name="Followed Hyperlink" xfId="10" builtinId="9" hidden="1"/>
        """
        node = fromstring(src)
        named_style = _NamedCellStyle.from_tree(node)
        assert named_style == _NamedCellStyle(
            name="Followed Hyperlink",
            xfId=10,
            builtinId=9,
            hidden=True
        )


@pytest.fixture
def _NamedCellStyleList():
    from ..named_styles import _NamedCellStyleList
    return _NamedCellStyleList


class TestNamedCellStyleList:

    def test_ctor(self, _NamedCellStyleList):
        styles = _NamedCellStyleList()
        xml = tostring(styles.to_tree())
        expected = """
        <cellStyles count ="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, _NamedCellStyleList):
        src = """
        <cellStyles />
        """
        node = fromstring(src)
        styles = _NamedCellStyleList.from_tree(node)
        assert styles == _NamedCellStyleList()


    def test_duplicate_names(self, _NamedCellStyleList):
        src = """
        <cellStyles count="11">
          <cellStyle name="Followed Hyperlink" xfId="2" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="4" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="6" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="8" builtinId="9" hidden="1"/>
          <cellStyle name="Followed Hyperlink" xfId="10" builtinId="9" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="1" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="3" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="5" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="7" builtinId="8" hidden="1"/>
          <cellStyle name="Hyperlink" xfId="9" builtinId="8" hidden="1"/>
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
        </cellStyles>
        """
        node = fromstring(src)
        styles = _NamedCellStyleList.from_tree(node)
        cleaned = styles.remove_duplicates()

        assert [s.name for s in cleaned] == ['Normal', 'Hyperlink', 'Followed Hyperlink']


    def test_duplicate_ids(self, _NamedCellStyleList):
        src = """
        <cellStyles count="18">
          <cellStyle name="Column0Style" xfId="1" />
          <cellStyle name="Column10Style" xfId="1"/>
          <cellStyle name="Column11Style" xfId="1"/>
          <cellStyle name="Column12Style" xfId="4"/>
          <cellStyle name="Column13Style" xfId="4"/>
          <cellStyle name="Column1Style" xfId="1"/>
          <cellStyle name="Column2Style" xfId="3"/>
          <cellStyle name="Column3Style" xfId="4"/>
          <cellStyle name="Column4Style" xfId="1"/>
          <cellStyle name="Column5Style" xfId="1"/>
          <cellStyle name="Column6Style" xfId="1"/>
          <cellStyle name="Column7Style" xfId="1"/>
          <cellStyle name="Column8Style" xfId="1"/>
          <cellStyle name="Column9Style" xfId="1"/>
          <cellStyle name="Heading" xfId="2"/>
          <cellStyle name="Hyperlink 2" xfId="6"/>
          <cellStyle name="Normal" xfId="0" builtinId="0"/>
          <cellStyle name="Normal 2" xfId="5"/>
        </cellStyles>
        """
        node = fromstring(src)
        styles = _NamedCellStyleList.from_tree(node)
        cleaned = styles.remove_duplicates()

        assert [s.name for s in cleaned] == ["Normal", "Column0Style", "Heading",
                                             "Column2Style", "Column12Style", "Normal 2",
                                             "Hyperlink 2"
                                             ]


@pytest.fixture
def NamedStyleList():
    from ..named_styles import NamedStyleList
    return NamedStyleList


class TestNamedStyleList:

    def test_append_valid(self, NamedStyle, NamedStyleList):
        styles = NamedStyleList()
        style = NamedStyle(name="special")
        styles.append(style)
        assert style in styles


    def test_append_invalid(self, NamedStyleList):
        styles = NamedStyleList()
        with pytest.raises(TypeError):
            styles.append(1)


    def test_duplicate(self, NamedStyleList, NamedStyle):
        styles = NamedStyleList()
        style = NamedStyle(name="special")
        styles.append(style)
        with pytest.raises(ValueError):
            styles.append(style)


    def test_names(self, NamedStyleList, NamedStyle):
        styles = NamedStyleList()
        style = NamedStyle(name="special")
        styles.append(style)
        assert styles.names == ['special']


    def test_idx(self, NamedStyleList, NamedStyle):
        styles = NamedStyleList()
        style = NamedStyle(name="special")
        styles.append(style)
        assert styles[0] == style


    def test_key(self, NamedStyleList, NamedStyle):
        styles = NamedStyleList()
        style = NamedStyle(name="special")
        styles.append(style)
        assert styles['special'] == style


    def test_key_error(self, NamedStyleList):
        styles = NamedStyleList()
        with pytest.raises(KeyError):
            styles['special']
