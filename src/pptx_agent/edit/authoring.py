"""New shapes, written the way PowerPoint writes them.

Everything here was measured: shapes, text boxes, connectors and a table inserted through
PowerPoint's own AppleScript (``make new shape``/``text box``/``connector``/``shape table``,
PowerPoint for Mac 16), saved, and the XML read back.  A new shape is

* ``p:sp`` with ``p:cNvSpPr`` empty, the frame and ``a:prstGeom`` with its ``a:avLst``;
* a ``p:style`` referring to the theme -- ``lnRef idx=2`` (``accent1`` shaded 15%),
  ``fillRef idx=1``, ``effectRef idx=0`` (both ``accent1``), ``fontRef minor`` (``lt1``) --
  so its fill, outline and text colour follow the theme as PowerPoint's do;
* a text body ``<a:bodyPr rtlCol="0" anchor="ctr"/>`` (default insets: none written),
  ``a:lstStyle``, and one centred paragraph ``<a:pPr algn="ctr"/>`` ending in
  ``<a:endParaRPr lang=.../>``.

A text box is ``txBox="1"``, ``a:noFill`` and no ``p:style`` (so no outline either), and
``<a:bodyPr vert="horz" wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr>`` with a plain
left-aligned paragraph.  (PowerPoint's scripted text box leaves ``wrap`` to its default,
which is ``square``; a text box drawn in the UI writes it, and so does this.)

A connector is ``p:cxnSp`` with ``<a:cxnSpLocks/>`` and ``a:stCxn``/``a:endCxn`` when
attached, and the style ``lnRef idx=2``, ``fillRef idx=0``, ``effectRef idx=1`` (all
``accent1``), ``fontRef minor`` (``tx1``).

A table is a ``p:graphicFrame`` with ``<a:graphicFrameLocks noGrp="1"/>``, its frame in
``p:xfrm``, and ``a:tbl`` whose ``a:tblPr firstRow="1" bandRow="1"`` names the table style
``{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`` (Medium Style 2 - Accent 1, PowerPoint's
default); columns and rows share the frame evenly and carry ``a16:colId``/``a16:rowId``;
every cell is an empty paragraph and an empty ``a:tcPr``.

Names are PowerPoint's English ones, numbered one below the shape's ``cNvPr@id``
("Rectangle 3" is id 4), from PowerPoint's own name table.  Two things PowerPoint adds are
left out on purpose: ``a16:creationId`` and ``p14:modId``, which are random GUIDs and
numbers -- a new shape gets this library's durable id stamp instead, and output stays
deterministic.  The text language is the presentation's default text language, where
PowerPoint uses its editing language.
"""

from __future__ import annotations

from lxml import etree

from ..oxml.xml import NAMESPACES, PA_ID_EXT_URI, Element, find, make, qn

#: PowerPoint's default table style, Medium Style 2 - Accent 1 (measured).
DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
#: Used when the presentation names no default text language.
FALLBACK_LANGUAGE = "en-US"

#: ``a:ext`` URIs PowerPoint keys table column and row ids with.
COLUMN_ID_URI = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"
ROW_ID_URI = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"

#: Preset -> the English name PowerPoint gives a new shape of it (its ``strman_f401`` name
#: table).  Presets PowerPoint cannot insert from its gallery have no entry and are named
#: "Shape", PowerPoint's generic name.
SHAPE_NAMES: dict[str, str] = {
    "line": "Straight Connector", "triangle": "Isosceles Triangle",
    "rtTriangle": "Right Triangle", "rect": "Rectangle", "diamond": "Diamond",
    "parallelogram": "Parallelogram", "trapezoid": "Trapezoid",
    "pentagon": "Regular Pentagon", "hexagon": "Hexagon", "heptagon": "Heptagon",
    "octagon": "Octagon", "decagon": "Decagon", "dodecagon": "Dodecagon",
    "star4": "4-Point Star", "star5": "5-Point Star", "star6": "6-Point Star",
    "star7": "7-Point Star", "star8": "8-Point Star", "star10": "10-Point Star",
    "star12": "12-Point Star", "star16": "16-Point Star", "star24": "24-Point Star",
    "star32": "32-Point Star", "roundRect": "Rounded Rectangle",
    "round1Rect": "Round Single Corner Rectangle",
    "round2SameRect": "Round Same Side Corner Rectangle",
    "round2DiagRect": "Round Diagonal Corner Rectangle",
    "snipRoundRect": "Snip and Round Single Corner Rectangle",
    "snip1Rect": "Snip Single Corner Rectangle", "snip2SameRect": "Snip Same Side Corner Rectangle",
    "snip2DiagRect": "Snip Diagonal Corner Rectangle", "plaque": "Plaque", "ellipse": "Oval",
    "teardrop": "Teardrop", "homePlate": "Pentagon", "chevron": "Chevron", "pie": "Pie",
    "blockArc": "Block Arc", "donut": "Donut", "noSmoking": '"No" Symbol',
    "rightArrow": "Right Arrow", "leftArrow": "Left Arrow", "upArrow": "Up Arrow",
    "downArrow": "Down Arrow", "stripedRightArrow": "Striped Right Arrow",
    "notchedRightArrow": "Notched Right Arrow", "bentUpArrow": "Bent-Up Arrow",
    "leftRightArrow": "Left-Right Arrow", "upDownArrow": "Up-Down Arrow",
    "leftUpArrow": "Left-Up Arrow", "leftRightUpArrow": "Left-Right-Up Arrow",
    "quadArrow": "Quad Arrow", "leftArrowCallout": "Left Arrow Callout",
    "rightArrowCallout": "Right Arrow Callout", "upArrowCallout": "Up Arrow Callout",
    "downArrowCallout": "Down Arrow Callout", "leftRightArrowCallout": "Left-Right Arrow Callout",
    "upDownArrowCallout": "Up-Down Arrow Callout", "quadArrowCallout": "Quad Arrow Callout",
    "bentArrow": "Bent Arrow", "uturnArrow": "U-Turn Arrow", "circularArrow": "Circular Arrow",
    "curvedRightArrow": "Curved Right Arrow", "curvedLeftArrow": "Curved Left Arrow",
    "curvedUpArrow": "Curved Up Arrow", "curvedDownArrow": "Curved Down Arrow", "cube": "Cube",
    "can": "Can", "lightningBolt": "Lightning Bolt", "heart": "Heart", "sun": "Sun",
    "moon": "Moon", "smileyFace": "Smiley Face", "irregularSeal1": "Explosion 1",
    "irregularSeal2": "Explosion 2", "foldedCorner": "Folded Corner", "bevel": "Bevel",
    "frame": "Frame", "halfFrame": "Half Frame", "corner": "L-Shape",
    "diagStripe": "Diagonal Stripe", "chord": "Chord", "arc": "Arc",
    "leftBracket": "Left Bracket", "rightBracket": "Right Bracket", "leftBrace": "Left Brace",
    "rightBrace": "Right Brace", "bracketPair": "Double Bracket", "bracePair": "Double Brace",
    "straightConnector1": "Straight Arrow Connector", "bentConnector2": "Elbow Connector",
    "bentConnector3": "Elbow Connector", "bentConnector4": "Elbow Connector",
    "bentConnector5": "Elbow Connector", "curvedConnector2": "Curved Connector",
    "curvedConnector3": "Curved Connector", "curvedConnector4": "Curved Connector",
    "curvedConnector5": "Curved Connector", "callout1": "Line Callout 1 (No Border)",
    "callout2": "Line Callout 2 (No Border)", "callout3": "Line Callout 3 (No Border)",
    "accentCallout1": "Line Callout 1 (Accent Bar)", "accentCallout2": "Line Callout 2 (Accent Bar)",
    "accentCallout3": "Line Callout 3 (Accent Bar)", "borderCallout1": "Line Callout 1",
    "borderCallout2": "Line Callout 2", "borderCallout3": "Line Callout 3",
    "accentBorderCallout1": "Line Callout 1 (Border and Accent Bar)",
    "accentBorderCallout2": "Line Callout 2 (Border and Accent Bar)",
    "accentBorderCallout3": "Line Callout 3 (Border and Accent Bar)",
    "wedgeRectCallout": "Rectangular Callout", "wedgeRoundRectCallout": "Rounded Rectangular Callout",
    "wedgeEllipseCallout": "Oval Callout", "cloudCallout": "Cloud Callout", "cloud": "Cloud",
    "ribbon": "Down Ribbon", "ribbon2": "Up Ribbon", "ellipseRibbon": "Curved Down Ribbon",
    "ellipseRibbon2": "Curved Up Ribbon", "verticalScroll": "Vertical Scroll",
    "horizontalScroll": "Horizontal Scroll", "wave": "Wave", "doubleWave": "Double Wave",
    "plus": "Cross", "flowChartProcess": "Flowchart: Process",
    "flowChartDecision": "Flowchart: Decision", "flowChartInputOutput": "Flowchart: Data",
    "flowChartPredefinedProcess": "Flowchart: Predefined Process",
    "flowChartInternalStorage": "Flowchart: Internal Storage",
    "flowChartDocument": "Flowchart: Document", "flowChartMultidocument": "Flowchart: Multidocument",
    "flowChartTerminator": "Flowchart: Terminator", "flowChartPreparation": "Flowchart: Preparation",
    "flowChartManualInput": "Flowchart: Manual Input",
    "flowChartManualOperation": "Flowchart: Manual Operation",
    "flowChartConnector": "Flowchart: Connector", "flowChartPunchedCard": "Flowchart: Card",
    "flowChartPunchedTape": "Flowchart: Punched Tape",
    "flowChartSummingJunction": "Flowchart: Summing Junction", "flowChartOr": "Flowchart: Or",
    "flowChartCollate": "Flowchart: Collate", "flowChartSort": "Flowchart: Sort",
    "flowChartExtract": "Flowchart: Extract", "flowChartMerge": "Flowchart: Merge",
    "flowChartOnlineStorage": "Flowchart: Stored Data",
    "flowChartMagneticTape": "Flowchart: Sequential Access Storage",
    "flowChartMagneticDisk": "Flowchart: Magnetic Disk",
    "flowChartMagneticDrum": "Flowchart: Direct Access Storage",
    "flowChartDisplay": "Flowchart: Display", "flowChartDelay": "Flowchart: Delay",
    "flowChartAlternateProcess": "Flowchart: Alternate Process",
    "flowChartOffpageConnector": "Flowchart: Off-page Connector",
    "actionButtonBlank": "Action Button: Custom", "actionButtonHome": "Action Button: Home",
    "actionButtonHelp": "Action Button: Help", "actionButtonInformation": "Action Button: Information",
    "actionButtonForwardNext": "Action Button: Forward or Next",
    "actionButtonBackPrevious": "Action Button: Back or Previous",
    "actionButtonEnd": "Action Button: End", "actionButtonBeginning": "Action Button: Beginning",
    "actionButtonReturn": "Action Button: Return", "actionButtonDocument": "Action Button: Document",
    "actionButtonSound": "Action Button: Sound", "actionButtonMovie": "Action Button: Movie",
    "mathPlus": "Plus", "mathMinus": "Minus", "mathMultiply": "Multiply",
    "mathDivide": "Division", "mathEqual": "Equal", "mathNotEqual": "Not Equal",
}
GENERIC_NAME = "Shape"
TEXT_BOX_NAME = "TextBox"
TABLE_NAME = "Table"
CHART_NAME = "Chart"


def default_name(base: str, identifier: int) -> str:
    """``"Rectangle 3"`` for a shape with ``cNvPr@id`` 4: PowerPoint numbers one below."""
    return f"{base} {identifier - 1}"


def default_language(package) -> str:
    """The presentation's default text language (``defaultTextStyle``), else en-US."""
    try:
        root = package.tree(package.presentation_part())
    except ValueError:
        return FALLBACK_LANGUAGE
    if root is None:
        return FALLBACK_LANGUAGE
    for path in ("p:defaultTextStyle/a:defPPr/a:defRPr", "p:defaultTextStyle/a:lvl1pPr/a:defRPr"):
        node = find(root, path)
        if node is not None and node.get("lang"):
            return node.get("lang")
    return FALLBACK_LANGUAGE


# -- pieces --------------------------------------------------------------------------------


def _stamped_cnv_pr(identifier: int, name: str, local: str) -> Element:
    """``p:cNvPr`` with this library's durable id already in place, prefix and all."""
    properties = make("p:cNvPr", id=str(identifier), name=name)
    ext_list = etree.SubElement(properties, qn("a:extLst"))
    ext = etree.SubElement(ext_list, qn("a:ext"))
    ext.set("uri", PA_ID_EXT_URI)
    value = etree.SubElement(ext, qn("pa:id"), nsmap={"pa": NAMESPACES["pa"]})
    value.set("val", local)
    return properties


def _xfrm(x: int, y: int, cx: int, cy: int, *, tag: str = "a:xfrm", rot: int = 0,
          flip_h: bool = False, flip_v: bool = False) -> Element:
    xfrm = make(tag)
    if rot:
        xfrm.set("rot", str(int(rot)))
    if flip_h:
        xfrm.set("flipH", "1")
    if flip_v:
        xfrm.set("flipV", "1")
    etree.SubElement(xfrm, qn("a:off"), x=str(int(x)), y=str(int(y)))
    etree.SubElement(xfrm, qn("a:ext"), cx=str(int(cx)), cy=str(int(cy)))
    return xfrm


def _geometry(preset: str, adjustments=()) -> Element:
    geometry = make("a:prstGeom", prst=preset)
    av_list = etree.SubElement(geometry, qn("a:avLst"))
    for name, value in adjustments:
        etree.SubElement(av_list, qn("a:gd"), name=name, fmla=f"val {int(value)}")
    return geometry


def _scheme(parent: Element, tag: str, value: str, idx: str, shade: str | None = None) -> None:
    reference = etree.SubElement(parent, qn(tag), idx=idx)
    color = etree.SubElement(reference, qn("a:schemeClr"), val=value)
    if shade is not None:
        etree.SubElement(color, qn("a:shade"), val=shade)


def _style(line: tuple, fill: str, effect: str, font_color: str) -> Element:
    style = make("p:style")
    _scheme(style, "a:lnRef", "accent1", *line)
    _scheme(style, "a:fillRef", "accent1", fill)
    _scheme(style, "a:effectRef", "accent1", effect)
    font = etree.SubElement(style, qn("a:fontRef"), idx="minor")
    etree.SubElement(font, qn("a:schemeClr"), val=font_color)
    return style


#: ``autofit`` -> the ``a:bodyPr`` child PowerPoint writes for it.
AUTOFIT_ELEMENTS = {"none": "a:noAutofit", "normal": "a:normAutofit", "shape": "a:spAutoFit"}


def _text_body(body_attributes: dict, language: str, *, centred: bool,
               autofit: str | None = None, tag: str = "p:txBody") -> Element:
    body = make(tag)
    properties = etree.SubElement(body, qn("a:bodyPr"))
    for name, value in body_attributes.items():
        properties.set(name, value)
    if autofit is not None:
        etree.SubElement(properties, qn(AUTOFIT_ELEMENTS[autofit]))
    etree.SubElement(body, qn("a:lstStyle"))
    paragraph = etree.SubElement(body, qn("a:p"))
    if centred:
        etree.SubElement(paragraph, qn("a:pPr"), algn="ctr")
    etree.SubElement(paragraph, qn("a:endParaRPr"), lang=language)
    return body


# -- whole shapes --------------------------------------------------------------------------


def autoshape(identifier: int, name: str, local: str, preset: str, x: int, y: int, cx: int,
              cy: int, language: str, adjustments=()) -> Element:
    shape = make("p:sp")
    nv = etree.SubElement(shape, qn("p:nvSpPr"))
    nv.append(_stamped_cnv_pr(identifier, name, local))
    etree.SubElement(nv, qn("p:cNvSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    properties = etree.SubElement(shape, qn("p:spPr"))
    properties.append(_xfrm(x, y, cx, cy))
    properties.append(_geometry(preset, adjustments))
    shape.append(_style(("2", "15000"), "1", "0", "lt1"))
    shape.append(_text_body({"rtlCol": "0", "anchor": "ctr"}, language, centred=True))
    return shape


def text_box(identifier: int, name: str, local: str, x: int, y: int, cx: int, cy: int,
             language: str, autofit: str = "shape") -> Element:
    shape = make("p:sp")
    nv = etree.SubElement(shape, qn("p:nvSpPr"))
    nv.append(_stamped_cnv_pr(identifier, name, local))
    etree.SubElement(nv, qn("p:cNvSpPr"), txBox="1")
    etree.SubElement(nv, qn("p:nvPr"))
    properties = etree.SubElement(shape, qn("p:spPr"))
    properties.append(_xfrm(x, y, cx, cy))
    properties.append(_geometry("rect"))
    etree.SubElement(properties, qn("a:noFill"))
    shape.append(_text_body({"vert": "horz", "wrap": "square", "rtlCol": "0"}, language,
                            centred=False, autofit=autofit))
    return shape


def connector(identifier: int, name: str, local: str, frame) -> Element:
    shape = make("p:cxnSp")
    nv = etree.SubElement(shape, qn("p:nvCxnSpPr"))
    nv.append(_stamped_cnv_pr(identifier, name, local))
    etree.SubElement(nv, qn("p:cNvCxnSpPr"))
    etree.SubElement(nv, qn("p:nvPr"))
    properties = etree.SubElement(shape, qn("p:spPr"))
    properties.append(_xfrm(frame.x, frame.y, frame.cx, frame.cy, rot=frame.rot,
                            flip_h=frame.flip_h, flip_v=frame.flip_v))
    properties.append(_geometry(frame.preset, frame.adjustments))
    shape.append(_style(("2",), "0", "1", "tx1"))
    return shape


def _even(total: int, count: int) -> list[int]:
    """``count`` sizes summing to exactly ``total``, as even as integers allow."""
    return [round((i + 1) * total / count) - round(i * total / count) for i in range(count)]


def _ext_id(parent: Element, uri: str, tag: str, value: int) -> None:
    ext_list = etree.SubElement(parent, qn("a:extLst"))
    ext = etree.SubElement(ext_list, qn("a:ext"), uri=uri)
    etree.SubElement(ext, qn(tag), nsmap={"a16": NAMESPACES["a16"]}, val=str(value))


def chart_frame(identifier: int, name: str, local: str, x: int, y: int, cx: int, cy: int,
                graphic: Element) -> Element:
    """A slide's ``p:graphicFrame`` around a chart's ``a:graphic``, as PowerPoint writes a
    chart it inserts (measured: no frame locks, an empty ``p:nvPr``)."""
    frame = make("p:graphicFrame")
    nv = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    nv.append(_stamped_cnv_pr(identifier, name, local))
    etree.SubElement(nv, qn("p:cNvGraphicFramePr"))
    etree.SubElement(nv, qn("p:nvPr"))
    frame.append(_xfrm(x, y, cx, cy, tag="p:xfrm"))
    frame.append(graphic)
    return frame


def table(identifier: int, name: str, local: str, rows: int, columns: int, x: int, y: int,
          cx: int, cy: int, language: str, style: str) -> Element:
    frame = make("p:graphicFrame")
    nv = etree.SubElement(frame, qn("p:nvGraphicFramePr"))
    nv.append(_stamped_cnv_pr(identifier, name, local))
    locks = etree.SubElement(nv, qn("p:cNvGraphicFramePr"))
    etree.SubElement(locks, qn("a:graphicFrameLocks"), noGrp="1")
    etree.SubElement(nv, qn("p:nvPr"))
    frame.append(_xfrm(x, y, cx, cy, tag="p:xfrm"))
    graphic = etree.SubElement(frame, qn("a:graphic"))
    data = etree.SubElement(graphic, qn("a:graphicData"),
                            uri="http://schemas.openxmlformats.org/drawingml/2006/table")
    tbl = etree.SubElement(data, qn("a:tbl"))
    properties = etree.SubElement(tbl, qn("a:tblPr"), firstRow="1", bandRow="1")
    if style:
        etree.SubElement(properties, qn("a:tableStyleId")).text = style
    grid = etree.SubElement(tbl, qn("a:tblGrid"))
    # PowerPoint's own ids are random 32-bit numbers; these only need to be unique within
    # the table, and the row and column edits number on from the largest.
    for index, width in enumerate(_even(cx, columns)):
        column = etree.SubElement(grid, qn("a:gridCol"), w=str(width))
        _ext_id(column, COLUMN_ID_URI, "a16:colId", 10001 + index)
    for index, height in enumerate(_even(cy, rows)):
        row = etree.SubElement(tbl, qn("a:tr"), h=str(height))
        for _ in range(columns):
            cell = etree.SubElement(row, qn("a:tc"))
            cell.append(_text_body({}, language, centred=False, tag="a:txBody"))
            etree.SubElement(cell, qn("a:tcPr"))
        _ext_id(row, ROW_ID_URI, "a16:rowId", 10001 + index)
    return frame
