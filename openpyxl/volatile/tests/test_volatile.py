# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def VolTopicRef():
    from ..volatile import VolTopicRef
    return VolTopicRef


class TestVolTopicRef:


    def test_ctor(self, VolTopicRef):
        ref = VolTopicRef(r="A1", s=3)
        xml = tostring(ref.to_tree())
        expected = """
        <tr r="A1" s="3" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, VolTopicRef):
        src = """
        <tr r="A1" s="1" />
        """
        node = fromstring(src)
        ref = VolTopicRef.from_tree(node)
        assert ref == VolTopicRef(r="A1", s=1)


@pytest.fixture
def VolMain():
    from ..volatile import VolMain
    return VolMain


class TestVolMain:


    def test_ctor(self, VolMain):
        main = VolMain(first="ThisDataModel")
        xml = tostring(main.to_tree())
        expected = """
        <main first="ThisDataModel" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



    def test_from_xml(self, VolMain):
        from ..volatile import VolTopic
        src = """
        <main first="ThisWorkbookDataModel">
            <tp t="e">
                <v>#N/A</v>
            </tp>
        </main>
        """
        node = fromstring(src)
        main = VolMain.from_tree(node)
        assert main == VolMain(first="ThisWorkbookDataModel", tp=[VolTopic(t="e", v="#N/A")])


@pytest.fixture
def VolType():
    from ..volatile import VolType
    return VolType


class TestVolType:


    def test_ctor(self, VolType):
        from ..volatile import VolMain, VolTopic
        typ = VolType(main=[VolMain(first="teststring", tp=[VolTopic(t="s", v='aaa: 4447')])], type="realTimeData")
        xml = tostring(typ.to_tree())
        expected = """
        <volType type="realTimeData">
            <main first="teststring">
                <tp t="s">
                <v>aaa: 4447</v>
                </tp>
            </main>
        </volType>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, VolType):
        src = """
        <volType type="olapFunctions">
            <main first="ThisWorkbookDataModel">
                <tp t="e">
                    <v>#N/A</v>
                    <stp>1</stp>
                    <tr r="A1" s="1" />
                </tp>
            </main>
        </volType>
        """
        node = fromstring(src)
        typ = VolType.from_tree(node)
        assert typ.type == "olapFunctions"
        assert typ.main[0].first == "ThisWorkbookDataModel"


@pytest.fixture
def VolTopic():
    from ..volatile import VolTopic
    return VolTopic


class TestVolTopic:


    def test_ctor(self, VolTopic):
        topic = VolTopic(t="s", v='aaa: 4447')
        xml = tostring(topic.to_tree())
        expected = """
        <tp t="s">
            <v>aaa: 4447</v>
        </tp>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, VolTopic):
        from ..volatile import VolTopicRef
        src = """
            <tp t="e">
                <v>#N/A</v>
                <stp>1</stp>
                <tr r="A1" s="1"></tr>
            </tp>
        """
        node = fromstring(src)
        topic = VolTopic.from_tree(node)
        assert topic == VolTopic(t="e", v="#N/A", stp="1", tr=[VolTopicRef(r="A1", s=1)])


@pytest.fixture
def VolTypes():
    from ..volatile import VolTypesList
    return VolTypesList


class TestVolTypes:


    def test_ctor(self, VolTypes):
        from ..volatile import VolMain, VolTopic, VolType

        typ = VolTypes(volType=[VolType(main=[VolMain(first="teststring", tp=[VolTopic(t="s", v='aaa: 4447')])], type="realTimeData")])
        xml = tostring(typ.to_tree())
        expected = """
        <volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <volType type="realTimeData">
                <main first="teststring">
                    <tp t="s">
                    <v>aaa: 4447</v>
                    </tp>
                </main>
            </volType>
        </volTypes>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, VolTypes):
        src = """
        <volTypes xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <volType type="olapFunctions">
                <main first="ThisWorkbookDataModel">
                    <tp t="e">
                        <v>#N/A</v>
                        <stp>1</stp>
                        <tr r="A1" s="1" />
                    </tp>
                </main>
            </volType>
        </volTypes>
        """
        node = fromstring(src)
        typ = VolTypes.from_tree(node)
        assert typ.volType[0].type == "olapFunctions"
        assert typ.volType[0].main[0].first == "ThisWorkbookDataModel"
