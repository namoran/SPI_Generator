# Copyright (c) 2010-2024 openpyxl
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    NoneSet,
    Set,
    Integer,
)
from openpyxl.descriptors.sequence import Sequence
from openpyxl.descriptors.nested import NestedText
from openpyxl.descriptors.excel import ExtensionList

from openpyxl.xml.constants import SHEET_MAIN_NS


class VolTopicRef(Serialisable):

    tagname = "tr"

    r = String()
    s = Integer()


    def __init__(self,
                 r=None,
                 s=None,
                ):
        self.r = r
        self.s = s


class VolTopic(Serialisable):

    tagname = "tp"

    v = NestedText(expected_type=str,)
    stp = NestedText(expected_type=str, allow_none=True)
    tr = Sequence(expected_type=VolTopicRef,)
    t = NoneSet(values=(["b", "n", "e", "s"]))

    __elements__ = ('v', 'stp', 'tr',)


    def __init__(self,
                 v=None,
                 stp=None,
                 tr=(),
                 t="n",
                ):
        self.v = v
        self.stp = stp
        self.tr = tr
        self.t = t


class VolMain(Serialisable):

    tagname = "main"

    tp = Sequence(expected_type=VolTopic,)
    first = String()


    def __init__(self,
                 tp=(),
                 first=None,
                ):
        self.tp = tp
        self.first = first


class VolType(Serialisable):

    tagname = "volType"

    main = Sequence(expected_type=VolMain)
    type = Set(values=(['realTimeData', 'olapFunctions']))

    __elements__ = ('main',)


    def __init__(self,
                 main=(),
                 type=None,
                ):
        self.main = main
        self.type = type


class VolTypesList(Serialisable):

    tagname = "volTypes"
    _path = "/xl/volatileDependencies.xml"
    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.volatileDependencies+xml"

    volType = Sequence(expected_type=VolType,)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('volType', 'extLst')


    def __init__(self,
                 volType=(),
                 extLst=None
                ):
        self.volType = volType
        self.extLst = extLst


    def to_tree(self, tagname=None, idx=None, namespace=None):
        tree = super().to_tree(tagname, idx, namespace)
        tree.set("xmlns", SHEET_MAIN_NS)
        return tree


    @property
    def path(self):
        return self._path
