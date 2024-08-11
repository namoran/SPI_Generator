# Copyright (c) 2010-2024 openpyxl
import pytest

from openpyxl.xml.functions import fromstring, tostring
from openpyxl.tests.helper import compare_xml


@pytest.fixture
def Connection():
    from ..connections import Connection
    return Connection


class TestConnection:


    def test_ctor(self, Connection):
        connection = Connection(id=3, refreshedVersion=8, background=True, keepAlive=True)
        xml = tostring(connection.to_tree())
        expected = """
        <connection id="3" keepAlive="1" refreshedVersion="8" background="1"
         interval="0" reconnectionMethod="1" minRefreshableVersion="0" savePassword="0"
         new="0" deleted="0" onlyUseConnectionFile="0" refreshOnLoad="0"
         saveData="0" credentials="integrated"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Connection):
        src = """
        <connection id="2" keepAlive="1"
        name="Query - Table1" description="Connection to the 'Table2' query in the workbook."
        type="5" refreshedVersion="8" background="1" saveData="1" />
        """
        node = fromstring(src)
        connection = Connection.from_tree(node)
        assert connection.saveData == True
        assert connection.name == "Query - Table1"
        assert connection.type_descriptions[connection.type] == "OLE DB-based source"


    @pytest.mark.parametrize("type, result", [
        (8, True),
        (102, False),
    ]
                        )
    def test_known_type(self, Connection, type, result):
        connection = Connection(id=3, refreshedVersion=8, background=True, type=type)
        assert connection.is_known_connection is result


@pytest.fixture
def DbPr():
    from ..connections import DbPr
    return DbPr


class TestDbPr:


    def test_ctor(self, DbPr):
        db_props = DbPr(connection="Data Model Connection", command="Model", commandType=True)
        xml = tostring(db_props.to_tree())
        expected = """
        <dbPr connection="Data Model Connection" command="Model" commandType="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, DbPr):
        src = """
        <dbPr connection="Provider=Microsoft.Mashup.OleDb"
            command="SELECT * FROM [Table2]" />
        """
        node = fromstring(src)
        db_props = DbPr.from_tree(node)
        assert db_props.command == "SELECT * FROM [Table2]"


@pytest.fixture
def OlapPr():
    from ..connections import OlapPr
    return OlapPr


class TestOlapPr:


    def test_ctor(self, OlapPr):
        olap_props = OlapPr(sendLocale=True, rowDrillCount=1000)
        xml = tostring(olap_props.to_tree())
        expected = """
        <olapPr local="0" sendLocale="1" rowDrillCount="1000" localRefresh="1" serverFill="1" serverNumberFormat="1" serverFont="1" serverFontColor="1" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OlapPr):
        src = """
        <olapPr serverFill="1" localRefresh="1" serverFontColor="1"/>
        """
        node = fromstring(src)
        olap_props = OlapPr.from_tree(node)
        assert olap_props == OlapPr(
            serverFill=True,
            localRefresh=True,
            serverFontColor=True
            )


@pytest.fixture
def TextField():
    from ..connections import TextField
    return TextField


class TestTextField:


    def test_ctor(self, TextField):
        text = TextField(type="text")
        xml = tostring(text.to_tree())
        expected = """
            <textField type="text" position="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TextField):
        from ..connections import TextField
        src = """
            <textField type="DMY" position="2" />
        """
        node = fromstring(src)
        text = TextField.from_tree(node)
        assert text == TextField(type="DMY", position=2)


@pytest.fixture
def TextPr():
    from ..connections import TextPr
    return TextPr


class TestTextPr:


    def test_ctor(self, TextPr):
        text_props = TextPr(prompt=False, codePage=437, sourceFile="C:\\Desktop\\text data.txt", delimiter="|")
        xml = tostring(text_props.to_tree())
        expected = """
        <textPr prompt="0" codePage="437" sourceFile="C:\\Desktop\\text data.txt" delimiter="|"
         fileType="win" firstRow="1" delimited="1" decimal="." thousands="," tab="1" qualifier="doubleQuote"
          space="0" comma="0" semicolon="0" consecutive="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TextPr):
        from ..connections import TextField
        src = """
        <textPr space="1" firstRow="13" sourceFile="C:\\Desktop\\text data.txt" delimiter="|">
            <textFields count="1">
                <textField />
            </textFields>
        </textPr>
        """
        node = fromstring(src)
        text_props = TextPr.from_tree(node)
        assert text_props == TextPr(space=True, firstRow=13, sourceFile="C:\\Desktop\\text data.txt", delimiter="|", textFields=[TextField()])


@pytest.fixture
def TableMissing():
    from ..connections import TableMissing
    return TableMissing


class TestTableMissing:


    def test_ctor(self, TableMissing):
        missing_table = TableMissing()
        xml = tostring(missing_table.to_tree())
        expected = """
        <tableMissing />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableMissing):
        src = """
        <tableMissing />
        """
        node = fromstring(src)
        missing_table = TableMissing.from_tree(node)
        assert missing_table == TableMissing()


@pytest.fixture
def Parameter():
    from ..connections import Parameter
    return Parameter


class TestParameter:


    def test_ctor(self, Parameter):
        param = Parameter(name="TestName", boolean=True, sqlType=4)
        xml = tostring(param.to_tree())
        expected = """
        <parameter name="TestName" boolean="1" sqlType="4" parameterType="prompt" refreshOnChange="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Parameter):
        src = """
        <parameter name="user" refreshOnChange="1" parameterType="cell" cell="Sheet1!$C$1"/>
        """
        node = fromstring(src)
        param = Parameter.from_tree(node)
        assert param == Parameter(name="user", refreshOnChange=True, parameterType="cell", cell="Sheet1!$C$1")


@pytest.fixture
def WebPr():
    from ..connections import WebPr
    return WebPr


class TestWebPr:


    def test_ctor(self, WebPr):
        web_props = WebPr(xml=True, firstRow=True, htmlTables=True)
        xml = tostring(web_props.to_tree())
        expected = """
        <webPr xml="1" firstRow="1" htmlTables="1" sourceData="0" parsePre="0" consecutive="0" xl97="0" textDates="0" xl2000="0"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, WebPr):
        src = """
        <webPr sourceData="1" parsePre="1"
        consecutive="1" url="http://ServerName/" htmlTables="1" />
        """
        node = fromstring(src)
        web_props = WebPr.from_tree(node)
        assert web_props == WebPr(sourceData=True, parsePre=True, consecutive=True, url="http://ServerName/", htmlTables=True)


@pytest.fixture
def Tables():
    from ..connections import Tables
    return Tables


class TestTables:


    def test_ctor(self, Tables):
        tables = Tables(x=3)
        xml = tostring(tables.to_tree())
        expected = """
        <tables>
            <x v="3" />
        </tables>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Tables):
        src = """
        <tables count="1">
            <s v="test" />
        </tables>
        """
        node = fromstring(src)
        tables = Tables.from_tree(node)
        assert tables == Tables(s="test", count=1)


@pytest.fixture
def ConnectionList():
    from ..connections import ConnectionList
    return ConnectionList


class TestConnectionList:


    def test_ctor(self, ConnectionList, Connection):
        connections = ConnectionList(connection=[Connection(id=1, refreshedVersion=4)])
        xml = tostring(connections.to_tree())
        expected = """
        <connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
            <connection id="1" refreshedVersion="4" background="0" credentials="integrated" deleted="0"
            interval="0" keepAlive="0" minRefreshableVersion="0" new="0" onlyUseConnectionFile="0" reconnectionMethod="1" refreshOnLoad="0"
            saveData="0" savePassword="0"/>
        </connections>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, ConnectionList, datadir, recwarn):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = ConnectionList.from_tree(node)
        assert len(connections.connection) == 2
        assert connections.connection[0].id == 2
        assert connections.connection[1].saveData == True

        w = recwarn.pop()
        assert issubclass(w.category, UserWarning)


    def test_access_by_id(self, ConnectionList, datadir):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = ConnectionList.from_tree(node)
        conn = connections[2]
        assert conn.name == "ThisWorkbookDataModel"


    def test_connection_not_found(self, ConnectionList, datadir):
        datadir.chdir()
        with open("connections.xml", "rb") as src:
            node = fromstring(src.read())
        connections = ConnectionList.from_tree(node)
        with pytest.raises(IndexError):
            connections[7]


    def test_caches(self, ConnectionList, Connection):
        from openpyxl.pivot.cache import CacheSource, CacheDefinition, CacheFieldList
        cxn1 = Connection(id=5, refreshedVersion=4)
        cxn2 = Connection(id=7, refreshedVersion=4)
        cxn2._cache = CacheDefinition(cacheSource=CacheSource(type="external"),
                                      cacheFields=CacheFieldList(),
                                      )
        cxns = ConnectionList(connection=[cxn1, cxn2])

        assert cxns.caches == [cxn2._cache]
        assert cxn2.id == 2
        assert cxn2._cache.cacheSource.connectionId == 2
