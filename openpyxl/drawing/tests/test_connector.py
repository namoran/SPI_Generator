# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml

from openpyxl.chart.shapes import (
    GraphicalProperties,
    Transform2D,
)
from openpyxl.chart.text import RichText
from ..geometry import (
    PositiveSize2D,
    Point2D,
    PresetGeometry2D,
    ShapeStyle,
    StyleMatrixReference,
    FontReference,
    GeomGuideList,
)
from ..colors import SchemeColor
from ..text import (
    RichTextProperties,
    ListStyle,
    Paragraph,
    ParagraphProperties,
    CharacterProperties,
)
from ..properties import NonVisualDrawingProps, NonVisualDrawingShapeProps


@pytest.fixture
def ConnectorShape():
    from ..connector import ConnectorShape
    return ConnectorShape


class TestConnectorShape:


    @pytest.mark.xfail
    def test_ctor(self, ConnectorShape):
        fut = ConnectorShape()
        xml = tostring(fut.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ConnectorShape):
        src = """
        <cxnSp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" macro="">
            <nvCxnSpPr>
                <cNvPr id="3" name="Straight Arrow Connector 2">
                </cNvPr>
                <cNvCxnSpPr/>
            </nvCxnSpPr>
            <spPr>
                <a:xfrm flipH="1" flipV="1">
                    <a:off x="3321050" y="3829050"/>
                    <a:ext cx="165100" cy="368300"/>
                </a:xfrm>
                <a:prstGeom prst="straightConnector1">
                    <a:avLst/>
                </a:prstGeom>
                <a:ln>
                    <a:tailEnd type="triangle"/>
                </a:ln>
            </spPr>
        </cxnSp>
        """
        node = fromstring(src)
        cnx = ConnectorShape.from_tree(node)
        assert cnx.nvCxnSpPr.cNvPr.id == 3


@pytest.fixture
def ShapeMeta():
    from ..connector import ShapeMeta
    return ShapeMeta


class TestShapeMeta:


    @pytest.mark.xfail
    def test_ctor(self, ShapeMeta):
        meta = ShapeMeta(cNvPr=NonVisualDrawingProps(), cNvSpPr=NonVisualDrawingShapeProps())
        xml = tostring(meta.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.xfail
    def test_from_xml(self, ShapeMeta):
        src = """
        <root />
        """
        node = fromstring(src)
        meta = ShapeMeta.from_tree(node)
        assert meta == ShapeMeta()


@pytest.fixture
def Shape():
    from ..connector import Shape
    return Shape


class TestShape:


    def test_ctor(self, Shape):
        props = GraphicalProperties(
            xfrm=Transform2D(off=Point2D(x=1767840, y=1341120),
                             ext=PositiveSize2D(cx=1539240, cy=281940)),
            prstGeom=PresetGeometry2D(prst="roundRect", avLst=GeomGuideList()))
        props.ln = None
        ln = StyleMatrixReference(idx=2, schemeClr=SchemeColor(val="accent1", shade=50000))
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        style = ShapeStyle(
            lnRef=ln,
            fillRef=fill,
            effectRef=effect,
            fontRef=font
        )
        body = RichTextProperties(vertOverflow="clip", horzOverflow="clip", rtlCol=False, anchor="t")
        p = Paragraph(endParaRPr=CharacterProperties(lang="en-US", sz="1100"),
                      pPr=ParagraphProperties(algn="l"))
        p.r = []
        text = RichText(bodyPr=body, lstStyle=ListStyle(), p=[p])
        shape = Shape(spPr=props, style=style, txBody=text, macro="[0]!RoundedRectangle1_Click")
        xml = tostring(shape.to_tree())
        expected = """
        <sp macro="[0]!RoundedRectangle1_Click" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <spPr>
          <a:xfrm>
            <a:off x="1767840" y="1341120"/>
            <a:ext cx="1539240" cy="281940"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect">
            <a:avLst/>
          </a:prstGeom>
        </spPr>
        <style>
          <a:lnRef idx="2">
            <a:schemeClr val="accent1">
              <a:shade val="50000"/>
            </a:schemeClr>
          </a:lnRef>
          <a:fillRef idx="1">
            <a:schemeClr val="accent1"/>
          </a:fillRef>
          <a:effectRef idx="0">
            <a:schemeClr val="accent1"/>
          </a:effectRef>
          <a:fontRef idx="minor">
            <a:schemeClr val="lt1"/>
          </a:fontRef>
        </style>
        <txBody>
          <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="l"/>
            <a:endParaRPr lang="en-US" sz="1100"/>
          </a:p>
        </txBody>
        </sp>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Shape, ShapeMeta):
        src = """
        <sp macro="[0]!RoundedRectangle1_Click" textlink="" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <nvSpPr>
          <cNvPr id="2" name="Rounded Rectangle 1"/>
          <cNvSpPr/>
        </nvSpPr>
        <spPr>
          <a:xfrm>
            <a:off x="1767840" y="1341120"/>
            <a:ext cx="1539240" cy="281940"/>
          </a:xfrm>
          <a:prstGeom prst="roundRect">
            <a:avLst/>
          </a:prstGeom>
        </spPr>
        <style>
          <a:lnRef idx="2">
            <a:schemeClr val="accent1">
              <a:shade val="50000"/>
            </a:schemeClr>
          </a:lnRef>
          <a:fillRef idx="1">
            <a:schemeClr val="accent1"/>
          </a:fillRef>
          <a:effectRef idx="0">
            <a:schemeClr val="accent1"/>
          </a:effectRef>
          <a:fontRef idx="minor">
            <a:schemeClr val="lt1"/>
          </a:fontRef>
        </style>
        <txBody>
          <a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="t"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="l"/>
            <a:endParaRPr lang="en-US" sz="1100"/>
          </a:p>
        </txBody>
        </sp>
        """
        node = fromstring(src)
        shape = Shape.from_tree(node)
        meta = ShapeMeta(cNvPr=NonVisualDrawingProps(id=2, name="Rounded Rectangle 1"),
                         cNvSpPr=NonVisualDrawingShapeProps()
                         )
        props = GraphicalProperties(
            xfrm=Transform2D(off=Point2D(x=1767840, y=1341120),
                             ext=PositiveSize2D(cx=1539240, cy=281940)),
            prstGeom=PresetGeometry2D(prst="roundRect", avLst=GeomGuideList()))
        ln = StyleMatrixReference(idx=2, schemeClr=SchemeColor(val="accent1", shade=50000))
        fill = StyleMatrixReference(idx=1, schemeClr=SchemeColor(val="accent1"))
        effect = StyleMatrixReference(idx=0, schemeClr=SchemeColor(val="accent1"))
        font = FontReference(idx="minor", schemeClr=SchemeColor(val="lt1"))
        style = ShapeStyle(
            lnRef=ln,
            fillRef=fill,
            effectRef=effect,
            fontRef=font
        )
        body = RichTextProperties(vertOverflow="clip", horzOverflow="clip", rtlCol=False, anchor="t")
        p = Paragraph(endParaRPr=CharacterProperties(lang="en-US", sz="1100"),
                      pPr=ParagraphProperties(algn="l"))

        text = RichText(bodyPr=body, lstStyle=ListStyle(), p=[p])
        shape2 = Shape(nvSpPr=meta, spPr=props, style=style, txBody=text, macro="[0]!RoundedRectangle1_Click")
        assert shape.meta == shape2.meta
        assert shape.spPr == shape2.spPr
        assert shape.style == shape2.style
        assert shape.txBody == shape2.txBody
        assert shape.macro == shape2.macro
