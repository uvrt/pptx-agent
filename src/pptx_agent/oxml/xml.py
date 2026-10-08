"""PresentationML and DrawingML on top of the format-neutral XML core.

The generic machinery -- qualified names, schema-ordered insertion, removal that keeps
whitespace -- lives in :mod:`ooxml_edit.xml` and is re-exported here, so existing imports
keep working.  This module adds what is specific to a ``.pptx``: the ``p:``/``a:`` vocabulary,
the child sequences of the elements this library inserts into, and the shape-tree walk.

**Child order is schema-enforced.**  DrawingML elements are ``xsd:sequence``, not ``xsd:all``:
``a:off`` must precede ``a:ext`` inside ``a:xfrm``, and PowerPoint refuses to open a file that
gets it wrong -- with a repair prompt rather than an error message that says why.  The tables
below are what :func:`insert_in_order` consults.
"""

from __future__ import annotations

from typing import Iterator

from ooxml_edit.xml import (  # noqa: F401  (re-exported)
    CHILD_ORDER,
    NAMESPACES,
    Element,
    _prefixed_name,
    append_in_order,
    child_elements,
    declared_prefixes,
    find,
    findall,
    get_bool,
    get_int,
    insert_in_order,
    iter_descendants,
    local_name,
    make,
    parse_xml,
    prefixed_name,
    qn,
    register_child_order,
    register_namespaces,
    remove,
    replace_choice,
    serialize,
    set_attr,
    set_int,
    subelement,
)

#: The namespaces a slide can reference.  Unknown namespaces in a document are preserved
#: regardless, since nodes are never detached from their tree.
PRESENTATION_NAMESPACES: dict[str, str] = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "dsp": "http://schemas.microsoft.com/office/drawing/2008/diagram",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "a16": "http://schemas.microsoft.com/office/drawing/2014/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    #: pptx-agent's own extension namespace, used to stamp durable shape ids.
    "pa": "https://github.com/uvrt/pptx-agent/2026/ids",
}
register_namespaces(PRESENTATION_NAMESPACES)

# DrawingML text, charts and diagrams: their namespaces (the same URIs as above, plus ``x:``)
# and the child orders of runs, breaks, fields, paragraphs and every chart and diagram
# element are ooxml-edit's charts subpackage's, registered once there for any document
# format.  The table below holds what is the slide's alone.
import ooxml_edit.charts.namespaces  # noqa: E402,F401  (registers a:p, a:r, c:*, dgm:*...)

#: URI of the ``a:ext`` we write into ``p:cNvPr`` to make a shape id permanent.  Office keys
#: extensions by URI and ignores ones it does not recognise, so this is a safe place to live.
PA_ID_EXT_URI = "{4D9F3A21-7C6B-4E52-9E1A-PPTXAGENTID01}"

#: URI PowerPoint itself uses for ``a16:creationId`` -- a durable per-shape id we prefer when
#: the authoring application already provided one.
CREATION_ID_EXT_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"


#: Members of DrawingML's ``EG_FillProperties`` choice -- at most one may be present.
FILL_TAGS: tuple[str, ...] = (
    "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill",
)
#: The same group without ``a:blipFill``/``a:grpFill``, as ``CT_LineProperties`` allows it.
LINE_FILL_TAGS: tuple[str, ...] = ("a:noFill", "a:solidFill", "a:gradFill", "a:pattFill")
#: Members of ``EG_ColorChoice``.
COLOR_TAGS: tuple[str, ...] = (
    "a:scrgbClr", "a:srgbClr", "a:hslClr", "a:sysClr", "a:schemeClr", "a:prstClr",
)
#: Members of ``EG_EffectProperties``.
EFFECT_TAGS: tuple[str, ...] = ("a:effectLst", "a:effectDag")
#: Everything a shape tree may hold between its group properties and its extLst.
SHAPE_TREE_MEMBERS: tuple[str, ...] = (
    "p:sp", "p:grpSp", "p:graphicFrame", "p:cxnSp", "p:pic", "p:contentPart",
    "mc:AlternateContent",
)

#: ``CT_LineProperties`` -- shared by a shape outline and the six table-cell borders.
_LINE_ORDER = (LINE_FILL_TAGS, ("a:prstDash", "a:custDash"), ("a:round", "a:bevel", "a:miter"),
               "a:headEnd", "a:tailEnd", "a:extLst")
#: ``CT_TextCharacterProperties`` -- run properties and their paragraph-level defaults.
_CHARACTER_ORDER = ("a:ln", FILL_TAGS, EFFECT_TAGS, "a:highlight", ("a:uLnTx", "a:uLn"),
                    ("a:uFillTx", "a:uFill"), "a:latin", "a:ea", "a:cs", "a:sym",
                    "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")
#: ``CT_TextParagraphProperties``.
_PARAGRAPH_ORDER = ("a:lnSpc", "a:spcBef", "a:spcAft", ("a:buClrTx", "a:buClr"),
                    ("a:buSzTx", "a:buSzPct", "a:buSzPts"), ("a:buFontTx", "a:buFont"),
                    ("a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip"), "a:tabLst",
                    "a:defRPr", "a:extLst")
#: ``CT_ShapeProperties``.
_SHAPE_PROPERTIES_ORDER = ("a:xfrm", ("a:custGeom", "a:prstGeom"), FILL_TAGS, "a:ln",
                           EFFECT_TAGS, "a:scene3d", "a:sp3d", "a:extLst")
#: ``CT_GroupShape`` -- the slide's shape tree and every group.
_GROUP_SHAPE_ORDER = ("p:nvGrpSpPr", "p:grpSpPr", SHAPE_TREE_MEMBERS, "p:extLst")

#: Parent tag -> the schema's child sequence; a tuple step is a repeating choice whose members
#: share a rank (see :func:`insert_in_order`).  Only parents this project inserts into need an
#: entry; :func:`insert_in_order` appends when a parent is absent from the table.
PRESENTATION_CHILD_ORDER: dict[str, tuple] = {
    "a:xfrm": ("a:off", "a:ext", "a:chOff", "a:chExt"),
    "p:spPr": _SHAPE_PROPERTIES_ORDER,
    "a:spPr": _SHAPE_PROPERTIES_ORDER,
    "p:nvSpPr": ("p:cNvPr", "p:cNvSpPr", "p:nvPr"),
    "p:nvPicPr": ("p:cNvPr", "p:cNvPicPr", "p:nvPr"),
    "p:nvCxnSpPr": ("p:cNvPr", "p:cNvCxnSpPr", "p:nvPr"),
    "p:nvGrpSpPr": ("p:cNvPr", "p:cNvGrpSpPr", "p:nvPr"),
    "p:nvGraphicFramePr": ("p:cNvPr", "p:cNvGraphicFramePr", "p:nvPr"),
    "p:cNvCxnSpPr": ("a:cxnSpLocks", "a:stCxn", "a:endCxn", "a:extLst"),
    "p:sp": ("p:nvSpPr", "p:spPr", "p:style", "p:txBody", "p:extLst"),
    "p:pic": ("p:nvPicPr", "p:blipFill", "p:spPr", "p:style", "p:extLst"),
    "p:cxnSp": ("p:nvCxnSpPr", "p:spPr", "p:style", "p:extLst"),
    "p:grpSp": _GROUP_SHAPE_ORDER,
    "p:spTree": _GROUP_SHAPE_ORDER,
    "p:graphicFrame": ("p:nvGraphicFramePr", "p:xfrm", "a:graphic", "p:extLst"),
    "p:grpSpPr": ("a:xfrm", FILL_TAGS, EFFECT_TAGS, "a:scene3d", "a:extLst"),
    "a:cNvPr": ("a:hlinkClick", "a:hlinkHover", "a:extLst"),
    "p:cNvPr": ("a:hlinkClick", "a:hlinkHover", "a:extLst"),
    "a:bodyPr": ("a:prstTxWarp", ("a:noAutofit", "a:normAutofit", "a:spAutoFit"), "a:scene3d",
                 ("a:sp3d", "a:flatTx"), "a:extLst"),
    "a:rPr": _CHARACTER_ORDER,
    "a:defRPr": _CHARACTER_ORDER,
    "a:endParaRPr": _CHARACTER_ORDER,
    "a:pPr": _PARAGRAPH_ORDER,
    # a:p, a:r, a:br and a:fld: see ooxml_edit.charts.namespaces.TEXT_CHILD_ORDER.
    "p:txBody": ("a:bodyPr", "a:lstStyle", "a:p"),
    "a:txBody": ("a:bodyPr", "a:lstStyle", "a:p"),
    "a:ln": _LINE_ORDER,
    "a:lnL": _LINE_ORDER,
    "a:lnR": _LINE_ORDER,
    "a:lnT": _LINE_ORDER,
    "a:lnB": _LINE_ORDER,
    "a:lnTlToBr": _LINE_ORDER,
    "a:lnBlToTr": _LINE_ORDER,
    "a:gradFill": ("a:gsLst", ("a:lin", "a:path"), "a:tileRect"),
    "a:blipFill": ("a:blip", "a:srcRect", ("a:tile", "a:stretch")),
    "p:blipFill": ("a:blip", "a:srcRect", ("a:tile", "a:stretch")),
    "a:tbl": ("a:tblPr", "a:tblGrid", "a:tr"),
    "a:tblPr": (FILL_TAGS, EFFECT_TAGS, ("a:tableStyle", "a:tableStyleId"), "a:extLst"),
    "a:tr": ("a:tc", "a:extLst"),
    "a:tc": ("a:txBody", "a:tcPr", "a:extLst"),
    # The presentation's lists, which slide-structure edits insert into.
    "p:presentation": ("p:sldMasterIdLst", "p:notesMasterIdLst", "p:handoutMasterIdLst",
                       "p:sldIdLst", "p:sldSz", "p:notesSz", "p:smartTags",
                       "p:embeddedFontLst", "p:custShowLst", "p:photoAlbum", "p:custDataLst",
                       "p:kinsoku", "p:defaultTextStyle", "p:modifyVerifier", "p:extLst"),
    "p:sldIdLst": ("p:sldId", "p:extLst"),
    "p14:section": ("p14:sldIdLst", "p14:extLst"),
    "p:sld": ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:extLst"),
    "p:cSld": ("p:bg", "p:spTree", "p:custDataLst", "p:controls", "p:extLst"),
    "a:tcPr": ("a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:lnTlToBr", "a:lnBlToTr", "a:cell3D",
               FILL_TAGS, "a:headers", "a:extLst"),
}
register_child_order(PRESENTATION_CHILD_ORDER)


# -- the shape tree ------------------------------------------------------------------------

SHAPE_TAGS: tuple[str, ...] = ("p:sp", "p:pic", "p:cxnSp", "p:grpSp", "p:graphicFrame")
_SHAPE_TAGS_RESOLVED = frozenset(qn(tag) for tag in SHAPE_TAGS)


def walk_shape_tree(container: Element) -> Iterator[tuple[Element, tuple[int, ...]]]:
    """Yield ``(shape element, index path)`` for every shape under a ``spTree``/``grpSp``.

    Groups are descended into, and the index path records the position at each level, so a
    shape nested two groups deep yields something like ``(3, 0, 1)``.  The path is what makes
    a shape addressable when ``cNvPr@id`` turns out not to be unique -- which happens in the
    wild, see :mod:`pptx_agent.edit.ids`.
    """
    yield from _walk(container, ())


def _walk(container: Element, prefix: tuple[int, ...]) -> Iterator[tuple[Element, tuple[int, ...]]]:
    index = 0
    for child in container:
        if child.tag not in _SHAPE_TAGS_RESOLVED:
            continue
        path = prefix + (index,)
        yield child, path
        if child.tag == qn("p:grpSp"):
            yield from _walk(child, path)
        index += 1


#: ``ST_ShapeType`` -- every preset geometry ``a:prstGeom@prst`` may name.  Writing any other
#: value is a schema violation PowerPoint answers with a repair prompt.
PRESET_GEOMETRIES: frozenset[str] = frozenset("""
    line lineInv triangle rtTriangle rect diamond parallelogram trapezoid
    nonIsoscelesTrapezoid pentagon hexagon heptagon octagon decagon dodecagon star4 star5
    star6 star7 star8 star10 star12 star16 star24 star32 roundRect round1Rect round2SameRect
    round2DiagRect snipRoundRect snip1Rect snip2SameRect snip2DiagRect plaque ellipse teardrop
    homePlate chevron pieWedge pie blockArc donut noSmoking rightArrow leftArrow upArrow
    downArrow stripedRightArrow notchedRightArrow bentUpArrow leftRightArrow upDownArrow
    leftUpArrow leftRightUpArrow quadArrow leftArrowCallout rightArrowCallout upArrowCallout
    downArrowCallout leftRightArrowCallout upDownArrowCallout quadArrowCallout bentArrow
    uturnArrow circularArrow leftCircularArrow leftRightCircularArrow curvedRightArrow
    curvedLeftArrow curvedUpArrow curvedDownArrow swooshArrow cube can lightningBolt heart sun
    moon smileyFace irregularSeal1 irregularSeal2 foldedCorner bevel frame halfFrame corner
    diagStripe chord arc leftBracket rightBracket leftBrace rightBrace bracketPair bracePair
    straightConnector1 bentConnector2 bentConnector3 bentConnector4 bentConnector5
    curvedConnector2 curvedConnector3 curvedConnector4 curvedConnector5 callout1 callout2
    callout3 accentCallout1 accentCallout2 accentCallout3 borderCallout1 borderCallout2
    borderCallout3 accentBorderCallout1 accentBorderCallout2 accentBorderCallout3
    wedgeRectCallout wedgeRoundRectCallout wedgeEllipseCallout cloudCallout cloud ribbon
    ribbon2 ellipseRibbon ellipseRibbon2 leftRightRibbon verticalScroll horizontalScroll wave
    doubleWave plus flowChartProcess flowChartDecision flowChartInputOutput
    flowChartPredefinedProcess flowChartInternalStorage flowChartDocument
    flowChartMultidocument flowChartTerminator flowChartPreparation flowChartManualInput
    flowChartManualOperation flowChartConnector flowChartPunchedCard flowChartPunchedTape
    flowChartSummingJunction flowChartOr flowChartCollate flowChartSort flowChartExtract
    flowChartMerge flowChartOfflineStorage flowChartOnlineStorage flowChartMagneticTape
    flowChartMagneticDisk flowChartMagneticDrum flowChartDisplay flowChartDelay
    flowChartAlternateProcess flowChartOffpageConnector actionButtonBlank actionButtonHome
    actionButtonHelp actionButtonInformation actionButtonForwardNext actionButtonBackPrevious
    actionButtonEnd actionButtonBeginning actionButtonReturn actionButtonDocument
    actionButtonSound actionButtonMovie gear6 gear9 funnel mathPlus mathMinus mathMultiply
    mathDivide mathEqual mathNotEqual cornerTabs squareTabs plaqueTabs chartX chartStar
    chartPlus
""".split())
