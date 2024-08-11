# Copyright (c) 2010-2024 openpyxl

from openpyxl.xml.functions import (
    tostring,
    fromstring,
)

from openpyxl.utils import (
    coordinate_to_tuple,
)

vmlns = "urn:schemas-microsoft-com:vml"
officens = "urn:schemas-microsoft-com:office:office"
excelns = "urn:schemas-microsoft-com:office:excel"

nsmap = {"o": officens,
         "v": vmlns,
         "x": excelns,
         }

VML_ROOT = f"""<xml xmlns:v="{vmlns}"  xmlns:o="{officens}"  xmlns:x="{excelns}" />"""

class ShapeWriter:
    """
    Create VML for comments
    """

    vml = None
    vml_path = None


    def __init__(self, comments):
        self.comments = comments


    def add_comment_shapetype(self):
        xml = """
<xml
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:v="urn:schemas-microsoft-com:vml">
   <o:shapelayout v:ext="edit">
    <o:idmap data="1" v:ext="edit"/>
  </o:shapelayout>
  <v:shapetype coordsize="21600,21600" id="_x0000_t202" path="m,l,21600r21600,l21600,xe" o:spt="202">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype></xml>"""
        tree = fromstring(xml)
        return [el for el in tree]


    def write(self, root):

        if root is None:
            root = VML_ROOT

        if not hasattr(root, "findall"):
            root = fromstring(root)

        # Remove any existing comment shapes
        comments = root.findall("v:shape[@type='#_x0000_t202']", nsmap)
        for c in comments:
            root.remove(c)

        # check whether comments shape type already exists
        shape_types = root.find("v:shapetype[@id='_x0000_t202']", nsmap)
        if shape_types is None:
            shape_types = self.add_comment_shapetype()
            root.extend(shape_types)

        for idx, (coord, comment) in enumerate(self.comments, 1026):
            shape = _shape_factory(coord, comment)
            shape.set("id", f"_x0000_s{idx:04d}")
            root.append(shape)

        return tostring(root)


def _shape_factory(coord, comment):
    """
    Add comment size and position to a shape node
    Row and column index need to be reduced by one
    """
    row, column = coordinate_to_tuple(coord)

    style = {"position": "absolute",
             "margin-left": "59.25pt",
             "margin-top": "1.5pt",
             "width": f"{comment.width}px",
             "height": f"{comment.height}px",
             "z-index": 1,
             "visibility": "hidden",
             }
    style = ";".join(f"{k}:{v}" for k, v in style.items())

    xml = f"""
    <v:shape
        xmlns:v="urn:schemas-microsoft-com:vml"
        xmlns:x="urn:schemas-microsoft-com:office:excel"
        xmlns:o="urn:schemas-microsoft-com:office:office"
        fillcolor="#ffffe1"
        style="{style}"
        type="#_x0000_t202"
        o:insetmode="auto">
      <v:fill color2="#ffffe1"/>
      <v:shadow color="black" obscured="t"/>
      <v:path o:connecttype="none"/>
      <v:textbox style="mso-direction-alt:auto">
        <div style="text-align:left"/>
      </v:textbox>
      <x:ClientData ObjectType="Note">
        <x:MoveWithCells/>
        <x:SizeWithCells/>
        <x:AutoFill>False</x:AutoFill>
        <x:Row>{row - 1}</x:Row>
        <x:Column>{column - 1}</x:Column>
      </x:ClientData>
    </v:shape>
    """
    tree = fromstring(xml)

    return tree
