# Copyright (c) 2010-2024 openpyxl

import pytest

from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import (
    fromstring,
    tostring,
    Element,
)

from ..comments import Comment
from ..shape_writer import (
    ShapeWriter,
    vmlns,
    excelns,
    _shape_factory,
)

def test_shape(datadir):

    datadir.chdir()
    with open('size+comments.vml', 'rb') as existing:
        expected = existing.read()

    coord = "D3"
    comment = Comment("a comment", "an author")

    shape = _shape_factory(coord, comment)
    xml = tostring(shape)

    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_shape_with_custom_size(datadir):

    datadir.chdir()
    with open('size+comments.vml', 'rb') as existing:
        expected = existing.read()
        # Change our source document for this test
        expected = expected.replace(b'width:144px;', b'width:80px;')
        expected = expected.replace(b'height:79px;', b'height:20px;')

    coord = "D3"
    comment = Comment("a comment", "an author", 20, 80)
    shape = _shape_factory(coord, comment)
    xml = tostring(shape)

    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def create_comments():

    comment1 = Comment("text", "author")
    comment2 = Comment("text2", "author2")
    comment3 = Comment("text3", "author3")

    coords = ["B2", "C7", "D9"]
    comments = [comment1, comment2, comment3]

    return (pair for pair in zip(coords, comments))


def test_merge_comments_vml(datadir, create_comments):
    datadir.chdir()
    cw = ShapeWriter(create_comments)

    with open('control+comments.vml', 'rb') as existing:
        content = fromstring(cw.write(fromstring(existing.read())))
    assert len(content.findall('{%s}shape' % vmlns)) == 5
    assert len(content.findall('{%s}shapetype' % vmlns)) == 2


def test_write_comments_vml(datadir, create_comments):
    datadir.chdir()
    cw = ShapeWriter(create_comments)

    content = cw.write(Element("xml"))
    with open('commentsDrawing1.vml', 'rb') as expected:
        correct = fromstring(expected.read())
    check = fromstring(content)
    correct_ids = []
    correct_coords = []
    check_ids = []
    check_coords = []

    for i in correct.findall("{%s}shape" % vmlns):
        correct_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        correct_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    for i in check.findall("{%s}shape" % vmlns):
        check_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        check_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    assert set(correct_coords) == set(check_coords)
    assert set(correct_ids) == set(check_ids)
    diff = compare_xml(tostring(correct), tostring(check))
    assert diff is None, diff

