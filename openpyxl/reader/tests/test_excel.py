# Copyright (c) 2010-2024 openpyxl
import os
from io import BytesIO
from shutil import copyfile
from tempfile import NamedTemporaryFile
from zipfile import BadZipfile, ZipFile

from openpyxl.packaging.manifest import Manifest, Override
from openpyxl.packaging.relationship import Relationship
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.xml.functions import fromstring
from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    XLSM,
    XLSX,
    XLTM,
    XLTX,
)
from openpyxl.workbook import Workbook

import pytest


@pytest.fixture
def load_workbook():
    from ..excel import load_workbook
    return load_workbook


def test_read_empty_file(datadir, load_workbook):
    datadir.chdir()
    with pytest.raises(BadZipfile):
        load_workbook('null_file.xlsx')


def test_load_workbook_from_fileobj(datadir, load_workbook):
    """ can a workbook be loaded from a file object without exceptions
    This tests for regressions of
    https://bitbucket.org/openpyxl/openpyxl/issue/433
    """
    datadir.chdir()
    with open('empty_with_no_properties.xlsx', 'rb') as f:
        load_workbook(f)


@pytest.mark.parametrize('wb_type, wb_name', [
    (ct, name) for ct in [XLSX, XLSM, XLTX, XLTM]
               for name in ['/' + ARC_WORKBOOK, '/xl/spqr.xml']
])
def test_find_standard_workbook_part(datadir, wb_type, wb_name):
    from ..excel import _find_workbook_part

    src = """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Override ContentType="{0}"
          PartName="{1}"/>
        </Types>
        """.format(wb_type, wb_name)
    node = fromstring(src)
    package = Manifest.from_tree(node)

    assert _find_workbook_part(package) == Override(wb_name, wb_type)


def test_no_workbook():
    from ..excel import _find_workbook_part

    with pytest.raises(IOError):
        part = _find_workbook_part(Manifest())


def test_overwritten_default():
    from ..excel import _find_workbook_part

    src = """
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    </Types>
    """
    node = fromstring(src)
    package = Manifest.from_tree(node)

    assert _find_workbook_part(package) == Override("/xl/workbook.xml", XLSX)


@pytest.mark.parametrize("extension",
                         ['.xlsb', '.xls', 'no-format']
                         )
def test_invalid_file_extension(extension, load_workbook):
    tmp = NamedTemporaryFile(suffix=extension)
    with pytest.raises(InvalidFileException):
        load_workbook(filename=tmp.name)


def test_style_assignment(datadir, load_workbook):
    datadir.chdir()

    wb = load_workbook("complex-styles.xlsx")
    assert len(wb._alignments) == 9
    assert len(wb._fills) == 6
    assert len(wb._fonts) == 8
    # 7 + 4 borders, because the top-left cell of a merg cell gets
    # a new border and the old ones are not deleted.
    assert len(wb._borders) == 11
    assert len(wb._number_formats) == 0
    assert len(wb._protections) == 1


@pytest.mark.parametrize("ro", [False, True])
def test_close_read(datadir, load_workbook, ro):
    datadir.chdir()

    wb = load_workbook("complex-styles.xlsx", read_only=ro)
    assert hasattr(wb, '_archive') is ro

    wb.close()

    if ro:
        assert wb._archive.fp is None


@pytest.mark.parametrize("wo", [False, True])
def test_close_write(wo):
    from openpyxl.workbook import Workbook
    wb = Workbook(write_only=wo)
    wb.close()


def test_read_stringio(load_workbook):
    filelike = BytesIO(b"certainly not a valid XSLX content")
    # Test invalid file-like objects are detected and not handled as regular files
    with pytest.raises(BadZipfile):
        load_workbook(filelike)


def test_load_workbook_with_vba(datadir, load_workbook):
    datadir.chdir()

    test_file = 'form_controls.xlsm'
    # open the workbook directly from the file
    wb = load_workbook(test_file)
    assert wb._vba is not None


def test_no_external_links(datadir, load_workbook):
    datadir.chdir()

    wb = load_workbook("bug137.xlsx", keep_links=False)
    assert wb._external_links == []


def test_file_closes(datadir, load_workbook):
    """Test whether workbook file is closed correctly after loading"""
    datadir.chdir()
    filename = "empty_with_no_properties-copy.xlsx"
    # create a copy that can be deleted later
    copyfile("empty_with_no_properties.xlsx", filename)

    load_workbook(filename)
    # remove would fail if the file is not closed correctly after loading
    os.remove(filename)


from ..excel import ExcelReader


class TestExcelReader:

    def test_ctor(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        assert reader.valid_files == [
            '[Content_Types].xml',
            '_rels/.rels',
            'xl/_rels/workbook.xml.rels',
            'xl/workbook.xml',
            'xl/sharedStrings.xml',
            'xl/theme/theme1.xml',
            'xl/styles.xml',
            'xl/worksheets/sheet1.xml',
            'docProps/thumbnail.jpeg',
            'docProps/core.xml',
            'docProps/app.xml'
        ]


    def test_read_manifest(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        assert reader.package is not None


    def test_read_strings(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_strings()
        assert reader.shared_strings != []


    def test_read_workbook(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        assert reader.wb is not None

    def test_read_workbook_theme(self, datadir):
        datadir.chdir()
        reader = ExcelReader("complex-styles.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        reader.read_theme()
        assert reader.wb.loaded_theme is not None

    @pytest.mark.parametrize("read_only", [False, True])
    def test_read_workbook_hidden(self, datadir, read_only):
        datadir.chdir()
        reader = ExcelReader("hidden_sheets.xlsx", read_only=read_only)
        reader.read()
        assert reader.wb.sheetnames == ["Sheet", "Hidden", "VeryHidden"]
        hidden = reader.wb.worksheets[1]
        assert hidden.sheet_state == "hidden"
        very_hidden = reader.wb.worksheets[2]
        assert very_hidden.sheet_state == "veryHidden"


    def test_read_chartsheet(self, datadir):
        datadir.chdir()
        reader = ExcelReader("contains_chartsheets.xlsx")
        reader.read_manifest()
        reader.read_workbook()

        rel = Relationship(Target="xl/chartsheets/sheet1.xml", type="chartsheet")

        class Sheet:
            pass

        sheet = Sheet()
        sheet.name = "chart"

        reader.read_chartsheet(sheet, rel)
        assert reader.wb['chart'].title == "chart"


    def test_read_volatile_deps(self, datadir):
        datadir.chdir()
        reader = ExcelReader("sample_with_volatile_deps_and_connection.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        reader.read_volatile_deps()
        # Test Parse
        assert reader.wb._volatile_deps is not None
        assert len(reader.wb._volatile_deps.volType) == 1


    def test_read_connections(self, datadir):
        datadir.chdir()
        reader = ExcelReader("sample_with_volatile_deps_and_connection.xlsx")
        reader.read_manifest()
        reader.read_workbook()
        reader.read_connections()
        # Test Parse
        assert reader.wb._connections is not None
        assert len(reader.wb._connections.connection) == 2
        # Check cache got assigned
        assert reader.wb._connections[3]._cache is not None
        assert reader.wb._connections[2]._cache is None


@pytest.fixture
def WorksheetProcessor():
    from .. excel import WorksheetProcessor
    return WorksheetProcessor


@pytest.fixture
def controls(datadir):
    datadir.chdir()
    with open("form_controls.xml", "rb") as src:
        xml = fromstring(src.read())
    return ControlList.from_tree(xml)


from openpyxl.worksheet.controls import (
    ControlList,
    FormControl,
    ActiveXControl,
)

class TestWorksheetProcessor:


    def test_find_children(self, datadir, WorksheetProcessor):
        datadir.chdir()
        archive = ZipFile("legacy_drawing.xlsm")
        wb = Workbook()
        ws = wb.create_sheet()
        processor = WorksheetProcessor(ws, archive)
        processor.find_children("xl/worksheets/sheet1.xml")
        assert len(processor.rels.vmlDrawing) == 1
        archive.close()


    @pytest.mark.xfail
    def test_get_controls(self, datadir, WorksheetProcessor, controls):
        datadir.chdir()
        archive = ZipFile("form_controls.xlsm")
        wb = Workbook()
        ws = wb.create_sheet()
        ws.controls  = controls

        processor = WorksheetProcessor(ws, archive)
        processor.find_children("xl/worksheets/sheet1.xml")
        assert len(processor.rels.ctrlProp) == 2
        assert len(processor.rels.control) == 5
        processor.get_controls()
        assert isinstance(ws.controls.control[-1].shape, FormControl)

        archive.close()


    def test_get_activex(self, datadir, WorksheetProcessor, load_workbook):
        datadir.chdir()
        archive = ZipFile("form_controls.xlsm")
        wb = load_workbook("form_controls.xlsm")
        ws = wb.active

        processor = WorksheetProcessor(ws, archive)
        processor.find_children("xl/worksheets/sheet1.xml")
        processor.get_activex()
        ctrl = ws.controls.control[0].shape
        assert isinstance(ctrl, ActiveXControl)
        assert ctrl.bin[:10] == b"@2\x05\xd7i\xce\xcd\x11\xa7w"

        embedded = []
        for ctrl in ws.controls.control:
            prop = ctrl.controlPr
            if prop.id:
                embedded.append(prop.image)

        assert len(embedded) == 3
        assert embedded[0].Target == "xl/media/image1.emf"
        assert embedded[0].blob._data()[:10]  == b"\x01\x00\x00\x00l\x00\x00\x00\x00\x00"

        archive.close()


    def test_get_comments(self, datadir, WorksheetProcessor):
        datadir.chdir()
        archive = ZipFile("legacy_drawing.xlsm")
        wb = Workbook()
        ws = wb.create_sheet()

        processor = WorksheetProcessor(ws, archive)
        processor.find_children("xl/worksheets/sheet1.xml")
        processor.get_comments()

        assert ws._cells != {} # make sure sheet is not empty
        comment = ws["B5"].comment
        assert comment.author == "Author"


    def test_get_legacy(self, datadir, WorksheetProcessor):
        datadir.chdir()
        archive = ZipFile("form_controls.xlsm")
        wb = Workbook()
        ws = wb.create_sheet()
        ws.legacy_drawing = "rId3"

        processor = WorksheetProcessor(ws, archive)
        processor.find_children("xl/worksheets/sheet1.xml")
        processor.get_legacy()

        drawing = ws.legacy_drawing
        assert drawing.path == "/xl/drawings/vmlDrawing0.vml"
        rel = drawing.children[0]
        assert rel.target == "xl/media/image3.emf"
        assert rel.blob._data()[:10] == b"\x01\x00\x00\x00l\x00\x00\x00\x00\x00"
