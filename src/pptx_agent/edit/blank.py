"""A new presentation from nothing, written the way PowerPoint writes a new one.

Every part is authored here from the values in the tables below; nothing is copied from
PowerPoint's own template.  The values were measured: a new blank presentation made by
PowerPoint for Mac 16 through its AppleScript (``make new presentation``), saved, and its
XML read back -- at 16:9, and again after switching it to 4:3 and to a banner size, to see
how PowerPoint lays the master and layouts out for another size.

What PowerPoint writes, and so what is written here:

* ``presentation.xml`` with one master, a 16:9 slide size of 12,192,000 x 6,858,000 EMU
  (no ``type``), a notes size of 6,858,000 x 9,144,000, and a ``defaultTextStyle`` of nine
  levels at 18 pt, ``marL`` 457,200 per level, in the theme's minor font and ``tx1``;
* ``presProps.xml``, ``viewProps.xml`` and ``tableStyles.xml`` (default style Medium Style 2
  - Accent 1);
* the Office theme: its twelve colours, Aptos Display / Aptos with their per-script
  fallbacks, and its fill, line, effect and background style lists;
* one master: title at (838200, 365125) 10515600 x 1325563, body at (838200, 1825625)
  10515600 x 4351338, date, footer and slide number along the bottom; its ``titleStyle``
  (44 pt, major font, 90% line spacing), ``bodyStyle`` (28/24/20/18 pt, bullet ``•`` in
  Arial, ``marL`` 228,600 + 457,200 per level, ``indent`` -228,600, 10 pt before the first
  level and 5 pt before the others) and ``otherStyle``;
* the eleven layouts PowerPoint offers -- Title Slide, Title and Content, Section Header,
  Two Content, Comparison, Title Only, Blank, Content with Caption, Picture with Caption,
  Title and Vertical Text, Vertical Title and Text -- with their placeholder types, indexes,
  sizes, positions and list styles;
* ``docProps/core.xml`` and ``docProps/app.xml``; and no slides.

**Other sizes.**  Switched to another size, PowerPoint scales every frame in the master and
the layouts axis by axis (x and width by the width ratio, y and height by the height ratio)
and every text metric -- font sizes, ``marL``, ``indent``, ``defTabSz``, ``spcPts`` -- by the
smaller of the two ratios, rounding to the nearest unit; insets, line spacing percentages and
the presentation's ``defaultTextStyle`` stay.  At 4:3 that is a ratio of 0.75 across and 1
down: 44 pt titles become 33 pt.  :func:`build` lays any size out that way.

Left out on purpose: random ids (``a16:creationId``, ``p14:creationId``) and the thumbnail,
which PowerPoint adds when it saves; the theme-family extension; and the localised names
(PowerPoint names parts in its UI language -- "Kantoorthema", "Titeldia" -- where this
writes the English ones).
"""

from __future__ import annotations

import io
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime

from lxml import etree

from ..oxml.xml import NAMESPACES, Element, qn, serialize

#: Named slide sizes, EMU.
SLIDE_SIZES: dict[str, tuple[int, int]] = {
    "16:9": (12192000, 6858000),
    "4:3": (9144000, 6858000),
}
#: ``p:sldSz@type`` for the sizes that have one; the 16:9 widescreen size has none.
_SIZE_TYPES = {
    (9144000, 6858000): "screen4x3",
    (9144000, 5143500): "screen16x9",
    (9144000, 5715000): "screen16x10",
    (7315200, 914400): "banner",
}
#: ``PresentationFormat`` in ``app.xml``, in English.
_SIZE_FORMATS = {
    (12192000, 6858000): "Widescreen",
    (9144000, 6858000): "On-screen Show (4:3)",
    (9144000, 5143500): "On-screen Show (16:9)",
    (9144000, 5715000): "On-screen Show (16:10)",
    (7315200, 914400): "Banner",
}
#: What the master and layouts are laid out for; other sizes are scaled from it.
BASE_SIZE = SLIDE_SIZES["16:9"]
NOTES_SIZE = (6858000, 9144000)
#: The smallest and largest slide sides the schema allows (``ST_SlideSizeCoordinate``).
MIN_SIDE, MAX_SIDE = 914400, 51206400

THEME_NAME = "Office Theme"
DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

CT = "application/vnd.openxmlformats-officedocument."
CT_PRESENTATION = CT + "presentationml.presentation.main+xml"
CT_TEMPLATE = CT + "presentationml.template.main+xml"
CT_MASTER = CT + "presentationml.slideMaster+xml"
CT_LAYOUT = CT + "presentationml.slideLayout+xml"
CT_THEME = CT + "theme+xml"
CT_PRES_PROPS = CT + "presentationml.presProps+xml"
CT_VIEW_PROPS = CT + "presentationml.viewProps+xml"
CT_TABLE_STYLES = CT + "presentationml.tableStyles+xml"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"
CT_APP = CT + "extended-properties+xml"
CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP = _R + "extended-properties"
REL_DOCUMENT = _R + "officeDocument"
REL_MASTER = _R + "slideMaster"
REL_LAYOUT = _R + "slideLayout"
REL_THEME = _R + "theme"
REL_PRES_PROPS = _R + "presProps"
REL_VIEW_PROPS = _R + "viewProps"
REL_TABLE_STYLES = _R + "tableStyles"

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
_P15 = "http://schemas.microsoft.com/office/powerpoint/2012/main"

# -- the theme -------------------------------------------------------------------------------

#: The Office theme's colours (PowerPoint 2023 onwards): name -> sRGB, or a system colour.
THEME_COLORS: tuple[tuple[str, str], ...] = (
    ("dk1", "sys:windowText:000000"), ("lt1", "sys:window:FFFFFF"),
    ("dk2", "0E2841"), ("lt2", "E8E8E8"),
    ("accent1", "156082"), ("accent2", "E97132"), ("accent3", "196B24"),
    ("accent4", "0F9ED5"), ("accent5", "A02B93"), ("accent6", "4EA72E"),
    ("hlink", "467886"), ("folHlink", "96607D"),
)
MAJOR_FONT, MINOR_FONT = "Aptos Display", "Aptos"
_APTOS_PANOSE = "02110004020202020204"
#: Per-script fallback faces of the Office font scheme: script -> (major, minor).
_SCRIPT_FONTS: tuple[tuple[str, str, str], ...] = (
    ("Jpan", "游ゴシック Light", "游ゴシック"), ("Hang", "맑은 고딕", "맑은 고딕"),
    ("Hans", "等线 Light", "等线"), ("Hant", "新細明體", "新細明體"),
    ("Arab", "Times New Roman", "Arial"), ("Hebr", "Times New Roman", "Arial"),
    ("Thai", "Angsana New", "Cordia New"), ("Ethi", "Nyala", "Nyala"),
    ("Beng", "Vrinda", "Vrinda"), ("Gujr", "Shruti", "Shruti"),
    ("Khmr", "MoolBoran", "DaunPenh"), ("Knda", "Tunga", "Tunga"),
    ("Guru", "Raavi", "Raavi"), ("Cans", "Euphemia", "Euphemia"),
    ("Cher", "Plantagenet Cherokee", "Plantagenet Cherokee"),
    ("Yiii", "Microsoft Yi Baiti", "Microsoft Yi Baiti"),
    ("Tibt", "Microsoft Himalaya", "Microsoft Himalaya"), ("Thaa", "MV Boli", "MV Boli"),
    ("Deva", "Mangal", "Mangal"), ("Telu", "Gautami", "Gautami"), ("Taml", "Latha", "Latha"),
    ("Syrc", "Estrangelo Edessa", "Estrangelo Edessa"), ("Orya", "Kalinga", "Kalinga"),
    ("Mlym", "Kartika", "Kartika"), ("Laoo", "DokChampa", "DokChampa"),
    ("Sinh", "Iskoola Pota", "Iskoola Pota"), ("Mong", "Mongolian Baiti", "Mongolian Baiti"),
    ("Viet", "Times New Roman", "Arial"), ("Uigh", "Microsoft Uighur", "Microsoft Uighur"),
    ("Geor", "Sylfaen", "Sylfaen"), ("Armn", "Arial", "Arial"),
    ("Bugi", "Leelawadee UI", "Leelawadee UI"),
    ("Bopo", "Microsoft JhengHei", "Microsoft JhengHei"),
    ("Java", "Javanese Text", "Javanese Text"), ("Lisu", "Segoe UI", "Segoe UI"),
    ("Mymr", "Myanmar Text", "Myanmar Text"), ("Nkoo", "Ebrima", "Ebrima"),
    ("Olck", "Nirmala UI", "Nirmala UI"), ("Osma", "Ebrima", "Ebrima"),
    ("Phag", "Phagspa", "Phagspa"), ("Syrn", "Estrangelo Edessa", "Estrangelo Edessa"),
    ("Syrj", "Estrangelo Edessa", "Estrangelo Edessa"),
    ("Syre", "Estrangelo Edessa", "Estrangelo Edessa"), ("Sora", "Nirmala UI", "Nirmala UI"),
    ("Tale", "Microsoft Tai Le", "Microsoft Tai Le"),
    ("Talu", "Microsoft New Tai Lue", "Microsoft New Tai Lue"), ("Tfng", "Ebrima", "Ebrima"),
)
#: The fonts a new deck uses, as ``app.xml`` lists them: the theme's two and the bullets'.
FONTS_USED = ("Aptos", "Aptos Display", "Arial")
BULLET_FONT = ("Arial", "020B0604020202020204", "34", "0")  # typeface, panose, pitch, charset

# -- the master and layouts, at 16:9 ---------------------------------------------------------

#: Placeholder frames on the master: type -> (x, y, cx, cy).
MASTER_FRAMES = {
    "title": (838200, 365125, 10515600, 1325563),
    "body": (838200, 1825625, 10515600, 4351338),
    "dt": (838200, 6356350, 2743200, 365125),
    "ftr": (4038600, 6356350, 4114800, 365125),
    "sldNum": (8610600, 6356350, 2743200, 365125),
}
#: ``bodyStyle``: (font size, space before in 1/100 pt) per level; ``marL`` 228,600 plus
#: 457,200 per level, ``indent`` -228,600.
BODY_LEVELS = ((2800, 1000), (2400, 500), (2000, 500), (1800, 500), (1800, 500),
               (1800, 500), (1800, 500), (1800, 500), (1800, 500))
TITLE_SIZE = 4400
OTHER_SIZE = 1800
FOOTER_SIZE = 1200
LEVEL_STEP = 457200
TAB_SIZE = 914400

#: Prompt texts, in English.
TITLE_PROMPT = "Click to edit Master title style"
SUBTITLE_PROMPT = "Click to edit Master subtitle style"
TEXT_PROMPTS = ("Click to edit Master text styles", "Second level", "Third level",
                "Fourth level", "Fifth level")


@dataclass(frozen=True)
class Placeholder:
    """One placeholder on a layout."""

    #: ``p:ph@type``; ``None`` for a content (object) placeholder.
    type: str | None
    idx: int | None
    #: The English base of its name: "Title", "Content Placeholder"...
    name: str
    frame: tuple[int, int, int, int] | None = None
    size: str | None = None          # p:ph@sz: "half", "quarter"
    orient: str | None = None        # p:ph@orient: "vert"
    anchor: str | None = None        # a:bodyPr@anchor
    vert: str | None = None          # a:bodyPr@vert
    #: Font sizes per level for the lstStyle, 1/100 pt (fewer than nine: the rest repeat
    #: the last), or ``None`` for no list style.
    sizes: tuple[int, ...] | None = None
    #: The list style sets marL/indent 0 and no bullet per level (subtitles, captions).
    flat: bool = False
    #: lvl1pPr algn in the list style (and every level when ``flat``).
    align: str | None = None
    bold: bool = False
    tinted: bool = False             # tx1 at 82% tint
    #: The prompt: "title", "subtitle", "text" (one level), "levels" (five), or None.
    prompt: str | None = None


def _footers() -> tuple[Placeholder, ...]:
    return (Placeholder("dt", 10, "Date Placeholder", size="half"),
            Placeholder("ftr", 11, "Footer Placeholder", size="quarter"),
            Placeholder("sldNum", 12, "Slide Number Placeholder", size="quarter"))


def _title(**fields) -> Placeholder:
    return Placeholder("title", None, "Title", prompt="title", **fields)


_CAPTION_TITLE = _title(frame=(839788, 457200, 3932237, 1600200), anchor="b", sizes=(3200,))
_CAPTION_TEXT = Placeholder("body", 2, "Text Placeholder", (839788, 2057400, 3932237, 3811588),
                            size="half", sizes=(1600, 1400, 1200, 1000), flat=True,
                            prompt="text")
_COMPARISON_HEAD = dict(anchor="b", sizes=(2400, 2000, 1800, 1600), flat=True, bold=True,
                        prompt="text")

#: The layouts, in PowerPoint's order: (name, type, placeholders).
LAYOUTS: tuple[tuple[str, str, tuple[Placeholder, ...]], ...] = (
    ("Title Slide", "title", (
        Placeholder("ctrTitle", None, "Title", (1524000, 1122363, 9144000, 2387600),
                    anchor="b", sizes=(6000,), align="ctr", prompt="title"),
        Placeholder("subTitle", 1, "Subtitle", (1524000, 3602038, 9144000, 1655762),
                    sizes=(2400, 2000, 1800, 1600), flat=True, align="ctr",
                    prompt="subtitle"),
    ) + _footers()),
    ("Title and Content", "obj", (
        _title(),
        Placeholder(None, 1, "Content Placeholder", prompt="levels"),
    ) + _footers()),
    ("Section Header", "secHead", (
        _title(frame=(831850, 1709738, 10515600, 2852737), anchor="b", sizes=(6000,)),
        Placeholder("body", 1, "Text Placeholder", (831850, 4589463, 10515600, 1500187),
                    sizes=(2400, 2000, 1800, 1600), flat=True, tinted=True, prompt="text"),
    ) + _footers()),
    ("Two Content", "twoObj", (
        _title(),
        Placeholder(None, 1, "Content Placeholder", (838200, 1825625, 5181600, 4351338),
                    size="half", prompt="levels"),
        Placeholder(None, 2, "Content Placeholder", (6172200, 1825625, 5181600, 4351338),
                    size="half", prompt="levels"),
    ) + _footers()),
    ("Comparison", "twoTxTwoObj", (
        _title(frame=(839788, 365125, 10515600, 1325563)),
        Placeholder("body", 1, "Text Placeholder", (839788, 1681163, 5157787, 823912),
                    **_COMPARISON_HEAD),
        Placeholder(None, 2, "Content Placeholder", (839788, 2505075, 5157787, 3684588),
                    size="half", prompt="levels"),
        Placeholder("body", 3, "Text Placeholder", (6172200, 1681163, 5183188, 823912),
                    size="quarter", **_COMPARISON_HEAD),
        Placeholder(None, 4, "Content Placeholder", (6172200, 2505075, 5183188, 3684588),
                    size="quarter", prompt="levels"),
    ) + _footers()),
    ("Title Only", "titleOnly", (_title(),) + _footers()),
    ("Blank", "blank", _footers()),
    ("Content with Caption", "objTx", (
        _CAPTION_TITLE,
        Placeholder(None, 1, "Content Placeholder", (5183188, 987425, 6172200, 4873625),
                    sizes=(3200, 2800, 2400, 2000), prompt="levels"),
        _CAPTION_TEXT,
    ) + _footers()),
    ("Picture with Caption", "picTx", (
        _CAPTION_TITLE,
        Placeholder("pic", 1, "Picture Placeholder", (5183188, 987425, 6172200, 4873625),
                    sizes=(3200, 2800, 2400, 2000), flat=True),
        _CAPTION_TEXT,
    ) + _footers()),
    ("Title and Vertical Text", "vertTx", (
        _title(),
        Placeholder("body", 1, "Vertical Text Placeholder", orient="vert", vert="eaVert",
                    prompt="levels"),
    ) + _footers()),
    ("Vertical Title and Text", "vertTitleAndTx", (
        Placeholder("title", None, "Vertical Title", (8724900, 365125, 2628900, 5811838),
                    orient="vert", vert="eaVert", prompt="title"),
        Placeholder("body", 1, "Vertical Text Placeholder", (838200, 365125, 7734300, 5811838),
                    orient="vert", vert="eaVert", prompt="levels"),
    ) + _footers()),
)

#: The master's placeholders.
MASTER_PLACEHOLDERS = (
    Placeholder("title", None, "Title Placeholder", MASTER_FRAMES["title"], anchor="ctr",
                prompt="title"),
    Placeholder("body", 1, "Text Placeholder", MASTER_FRAMES["body"], prompt="levels"),
    Placeholder("dt", 2, "Date Placeholder", MASTER_FRAMES["dt"], size="half", anchor="ctr",
                sizes=(FOOTER_SIZE,), align="l", tinted=True),
    Placeholder("ftr", 3, "Footer Placeholder", MASTER_FRAMES["ftr"], size="quarter",
                anchor="ctr", sizes=(FOOTER_SIZE,), align="ctr", tinted=True),
    Placeholder("sldNum", 4, "Slide Number Placeholder", MASTER_FRAMES["sldNum"],
                size="quarter", anchor="ctr", sizes=(FOOTER_SIZE,), align="r", tinted=True),
)

#: ``p:sldMasterId`` and the first ``p:sldLayoutId``: PowerPoint's ids for them.
MASTER_ID = 2147483648

#: Field ids for the date and slide-number fields, shared by the master and every layout.
_FIELD_NAMESPACE = uuid.UUID("7d9c1f3e-5a2b-4c8d-9e6f-0a1b2c3d4e5f")


def _field_id(name: str) -> str:
    return "{%s}" % str(uuid.uuid5(_FIELD_NAMESPACE, name)).upper()


DATE_FIELD_ID = _field_id("datetimeFigureOut")
NUMBER_FIELD_ID = _field_id("slidenum")


# -- building ---------------------------------------------------------------------------------


class _Scale:
    """How a size is laid out from the 16:9 one: frames per axis, text by the smaller ratio."""

    def __init__(self, width: int, height: int) -> None:
        self.sx = width / BASE_SIZE[0]
        self.sy = height / BASE_SIZE[1]
        self.text = min(self.sx, self.sy)

    def frame(self, frame: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
        x, y, cx, cy = frame
        return round(x * self.sx), round(y * self.sy), round(cx * self.sx), round(cy * self.sy)

    def metric(self, value: int) -> int:
        return round(value * self.text)


def _root(tag: str, prefixes: tuple[str, ...], **attributes: str) -> Element:
    element = etree.Element(qn(tag), nsmap={p: NAMESPACES[p] for p in prefixes})
    for name, value in attributes.items():
        element.set(qn(name.replace("__", ":")) if "__" in name else name, value)
    return element


def _sub(parent: Element, tag: str, **attributes: str) -> Element:
    element = etree.SubElement(parent, qn(tag))
    for name, value in attributes.items():
        element.set(qn(name.replace("__", ":")) if "__" in name else name, value)
    return element


def _scheme(parent: Element, value: str, tint: int | None = None) -> None:
    color = _sub(_sub(parent, "a:solidFill"), "a:schemeClr", val=value)
    if tint is not None:
        _sub(color, "a:tint", val=str(tint))


def _font_refs(parent: Element, major: bool) -> None:
    kind = "mj" if major else "mn"
    for tag in ("a:latin", "a:ea", "a:cs"):
        _sub(parent, tag, typeface=f"+{kind}-{tag[2:] if tag != 'a:latin' else 'lt'}")


def _level(parent: Element, level: int, *, marL: int | None, indent: int | None,
           size: int, major: bool = False, space_before: int | None = None,
           line_spacing: bool = False, bullet: bool | None = None, tab: int,
           scale: _Scale) -> Element:
    properties = _sub(parent, f"a:lvl{level}pPr")
    if marL is not None:
        properties.set("marL", str(scale.metric(marL)))
    if indent is not None:
        properties.set("indent", str(scale.metric(indent)))
    for name, value in (("algn", "l"), ("defTabSz", str(scale.metric(tab))), ("rtl", "0"),
                        ("eaLnBrk", "1"), ("latinLnBrk", "0"), ("hangingPunct", "1")):
        properties.set(name, value)
    if line_spacing:
        _sub(_sub(properties, "a:lnSpc"), "a:spcPct", val="90000")
    if space_before is not None:
        spacing = _sub(properties, "a:spcBef")
        if space_before:
            _sub(spacing, "a:spcPts", val=str(scale.metric(space_before)))
        else:
            _sub(spacing, "a:spcPct", val="0")
    if bullet is False:
        _sub(properties, "a:buNone")
    elif bullet:
        typeface, panose, pitch, charset = BULLET_FONT
        _sub(properties, "a:buFont", typeface=typeface, panose=panose, pitchFamily=pitch,
             charset=charset)
        _sub(properties, "a:buChar", char="•")
    run = _sub(properties, "a:defRPr", sz=str(scale.metric(size)), kern="1200")
    _scheme(run, "tx1")
    _font_refs(run, major)
    return properties


def _plain_levels(parent: Element, language: str, scale: _Scale, *,
                  text_scale: bool = True) -> None:
    """``defaultTextStyle``/``otherStyle``: a language and nine plain 18 pt levels."""
    _sub(_sub(parent, "a:defPPr"), "a:defRPr", lang=language)
    metrics = scale if text_scale else _Scale(*BASE_SIZE)
    for level in range(1, 10):
        _level(parent, level, marL=(level - 1) * LEVEL_STEP, indent=None, size=OTHER_SIZE,
               tab=TAB_SIZE, scale=metrics)


def _write(element: Element) -> bytes:
    return serialize(element)


def presentation_xml(width: int, height: int, language: str, master_rel: str) -> bytes:
    root = _root("p:presentation", ("a", "r", "p"), saveSubsetFonts="1",
                 autoCompressPictures="0")
    _sub(_sub(root, "p:sldMasterIdLst"), "p:sldMasterId", id=str(MASTER_ID), r__id=master_rel)
    size = _sub(root, "p:sldSz", cx=str(width), cy=str(height))
    if (width, height) in _SIZE_TYPES:
        size.set("type", _SIZE_TYPES[(width, height)])
    _sub(root, "p:notesSz", cx=str(NOTES_SIZE[0]), cy=str(NOTES_SIZE[1]))
    # The presentation's default text style is not scaled with the slide size (measured).
    _plain_levels(_sub(root, "p:defaultTextStyle"), language, _Scale(*BASE_SIZE))
    return _write(root)


def pres_props_xml() -> bytes:
    root = _root("p:presentationPr", ("a", "r", "p"))
    extensions = _sub(root, "p:extLst")
    for uri, prefix, namespace, tag, value in (
            ("{E76CE94A-603C-4142-B9EB-6D1370010A27}", "p14", _P14, "discardImageEditData", "0"),
            ("{D31A062A-798A-4329-ABDD-BBA856620510}", "p14", _P14, "defaultImageDpi", "32767"),
            ("{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}", "p15", _P15, "chartTrackingRefBased",
             "1")):
        extension = _sub(extensions, "p:ext", uri=uri)
        node = etree.SubElement(extension, "{%s}%s" % (namespace, tag), nsmap={prefix: namespace})
        node.set("val", value)
    return _write(root)


def view_props_xml() -> bytes:
    root = _root("p:viewPr", ("a", "r", "p"))
    normal = _sub(root, "p:normalViewPr")
    _sub(normal, "p:restoredLeft", sz="15611")
    _sub(normal, "p:restoredTop", sz="94658")
    common = _sub(_sub(root, "p:slideViewPr"), "p:cSldViewPr", snapToGrid="0")
    view = _sub(common, "p:cViewPr", varScale="1")
    scale = _sub(view, "p:scale")
    _sub(scale, "a:sx", n="100", d="100")
    _sub(scale, "a:sy", n="100", d="100")
    _sub(view, "p:origin", x="0", y="0")
    _sub(common, "p:guideLst")
    notes = _sub(_sub(root, "p:notesTextViewPr"), "p:cViewPr")
    scale = _sub(notes, "p:scale")
    _sub(scale, "a:sx", n="1", d="1")
    _sub(scale, "a:sy", n="1", d="1")
    _sub(notes, "p:origin", x="0", y="0")
    _sub(root, "p:gridSpacing", cx="72008", cy="72008")
    return _write(root)


def table_styles_xml() -> bytes:
    return _write(_root("a:tblStyleLst", ("a",), **{"def": DEFAULT_TABLE_STYLE}))


def theme_xml() -> bytes:
    root = _root("a:theme", ("a",), name=THEME_NAME)
    elements = _sub(root, "a:themeElements")
    colors = _sub(elements, "a:clrScheme", name="Office")
    for name, value in THEME_COLORS:
        holder = _sub(colors, f"a:{name}")
        if value.startswith("sys:"):
            _, system, last = value.split(":")
            _sub(holder, "a:sysClr", val=system, lastClr=last)
        else:
            _sub(holder, "a:srgbClr", val=value)
    fonts = _sub(elements, "a:fontScheme", name="Office")
    for tag, face, column in (("a:majorFont", MAJOR_FONT, 1), ("a:minorFont", MINOR_FONT, 2)):
        holder = _sub(fonts, tag)
        _sub(holder, "a:latin", typeface=face, panose=_APTOS_PANOSE)
        _sub(holder, "a:ea", typeface="")
        _sub(holder, "a:cs", typeface="")
        for entry in _SCRIPT_FONTS:
            _sub(holder, "a:font", script=entry[0], typeface=entry[column])
    formats = _sub(elements, "a:fmtScheme", name="Office")
    fills = _sub(formats, "a:fillStyleLst")
    _scheme(fills, "phClr")
    _gradient(fills, ((0, (("lumMod", 110000), ("satMod", 105000), ("tint", 67000))),
                      (50000, (("lumMod", 105000), ("satMod", 103000), ("tint", 73000))),
                      (100000, (("lumMod", 105000), ("satMod", 109000), ("tint", 81000)))))
    _gradient(fills, ((0, (("satMod", 103000), ("lumMod", 102000), ("tint", 94000))),
                      (50000, (("satMod", 110000), ("lumMod", 100000), ("shade", 100000))),
                      (100000, (("lumMod", 99000), ("satMod", 120000), ("shade", 78000)))))
    lines = _sub(formats, "a:lnStyleLst")
    for width in ("12700", "19050", "25400"):
        line = _sub(lines, "a:ln", w=width, cap="flat", cmpd="sng", algn="ctr")
        _scheme(line, "phClr")
        _sub(line, "a:prstDash", val="solid")
        _sub(line, "a:miter", lim="800000")
    effects = _sub(formats, "a:effectStyleLst")
    _sub(_sub(effects, "a:effectStyle"), "a:effectLst")
    _sub(_sub(effects, "a:effectStyle"), "a:effectLst")
    shadow = _sub(_sub(_sub(effects, "a:effectStyle"), "a:effectLst"), "a:outerShdw",
                  blurRad="57150", dist="19050", dir="5400000", algn="ctr", rotWithShape="0")
    _sub(_sub(shadow, "a:srgbClr", val="000000"), "a:alpha", val="63000")
    backgrounds = _sub(formats, "a:bgFillStyleLst")
    _scheme(backgrounds, "phClr")
    color = _sub(_sub(backgrounds, "a:solidFill"), "a:schemeClr", val="phClr")
    _sub(color, "a:tint", val="95000")
    _sub(color, "a:satMod", val="170000")
    _gradient(backgrounds, (
        (0, (("tint", 93000), ("satMod", 150000), ("shade", 98000), ("lumMod", 102000))),
        (50000, (("tint", 98000), ("satMod", 130000), ("shade", 90000), ("lumMod", 103000))),
        (100000, (("shade", 63000), ("satMod", 120000)))))
    defaults = _sub(root, "a:objectDefaults")
    line_default = _sub(defaults, "a:lnDef")
    for tag in ("a:spPr", "a:bodyPr", "a:lstStyle"):
        _sub(line_default, tag)
    style = _sub(line_default, "a:style")
    for tag, index, color in (("a:lnRef", "2", "accent1"), ("a:fillRef", "0", "accent1"),
                              ("a:effectRef", "1", "accent1"), ("a:fontRef", "minor", "tx1")):
        _sub(_sub(style, tag, idx=index), "a:schemeClr", val=color)
    _sub(root, "a:extraClrSchemeLst")
    return _write(root)


def _gradient(parent: Element, stops) -> None:
    gradient = _sub(parent, "a:gradFill", rotWithShape="1")
    stop_list = _sub(gradient, "a:gsLst")
    for position, modifiers in stops:
        color = _sub(_sub(stop_list, "a:gs", pos=str(position)), "a:schemeClr", val="phClr")
        for name, value in modifiers:
            _sub(color, f"a:{name}", val=str(value))
    _sub(gradient, "a:lin", ang="5400000", scaled="0")


def _shape_tree(parent: Element) -> Element:
    tree = _sub(parent, "p:spTree")
    nv = _sub(tree, "p:nvGrpSpPr")
    _sub(nv, "p:cNvPr", id="1", name="")
    _sub(nv, "p:cNvGrpSpPr")
    _sub(nv, "p:nvPr")
    xfrm = _sub(_sub(tree, "p:grpSpPr"), "a:xfrm")
    _sub(xfrm, "a:off", x="0", y="0")
    _sub(xfrm, "a:ext", cx="0", cy="0")
    _sub(xfrm, "a:chOff", x="0", y="0")
    _sub(xfrm, "a:chExt", cx="0", cy="0")
    return tree


def _placeholder(tree: Element, identifier: int, spec: Placeholder, scale: _Scale,
                 language: str, date_text: str, *, on_master: bool) -> None:
    shape = _sub(tree, "p:sp")
    nv = _sub(shape, "p:nvSpPr")
    _sub(nv, "p:cNvPr", id=str(identifier), name=f"{spec.name} {identifier - 1}")
    _sub(_sub(nv, "p:cNvSpPr"), "a:spLocks", noGrp="1")
    placeholder = _sub(_sub(nv, "p:nvPr"), "p:ph")
    for name, value in (("type", spec.type), ("orient", spec.orient), ("sz", spec.size),
                        ("idx", None if spec.idx is None else str(spec.idx))):
        if value is not None:
            placeholder.set(name, value)
    properties = _sub(shape, "p:spPr")
    if spec.frame is not None:
        x, y, cx, cy = scale.frame(spec.frame)
        xfrm = _sub(properties, "a:xfrm")
        _sub(xfrm, "a:off", x=str(x), y=str(y))
        _sub(xfrm, "a:ext", cx=str(cx), cy=str(cy))
        if on_master:
            _sub(_sub(properties, "a:prstGeom", prst="rect"), "a:avLst")
    body = _sub(shape, "p:txBody")
    body_properties = _sub(body, "a:bodyPr")
    if on_master:
        if spec.vert:
            body_properties.set("vert", spec.vert)
        else:
            body_properties.set("vert", "horz")
        for name in ("lIns", "tIns", "rIns", "bIns"):
            body_properties.set(name, "91440" if name[0] in "lr" else "45720")
        body_properties.set("rtlCol", "0")
        if spec.anchor:
            body_properties.set("anchor", spec.anchor)
        if spec.type in ("title", "body"):
            _sub(body_properties, "a:normAutofit")
    else:
        if spec.vert:
            body_properties.set("vert", spec.vert)
        if spec.anchor:
            body_properties.set("anchor", spec.anchor)
    _list_style(_sub(body, "a:lstStyle"), spec, scale)
    _prompt(body, spec, language, date_text)


def _list_style(styles: Element, spec: Placeholder, scale: _Scale) -> None:
    if spec.sizes is None:
        return
    levels = 9 if (spec.flat or len(spec.sizes) > 1) else 1
    for level in range(1, levels + 1):
        size = spec.sizes[min(level - 1, len(spec.sizes) - 1)]
        properties = _sub(styles, f"a:lvl{level}pPr")
        if spec.flat:
            properties.set("marL", str(scale.metric((level - 1) * LEVEL_STEP)))
            properties.set("indent", "0")
        if spec.align and (spec.flat or level == 1):
            properties.set("algn", spec.align)
        if spec.flat:
            _sub(properties, "a:buNone")
        run = _sub(properties, "a:defRPr", sz=str(scale.metric(size)))
        if spec.bold:
            run.set("b", "1")
        if spec.tinted:
            _scheme(run, "tx1", tint=82000)


def _prompt(body: Element, spec: Placeholder, language: str, date_text: str) -> None:
    if spec.type == "dt":
        paragraph = _sub(body, "a:p")
        field = _sub(paragraph, "a:fld", id=DATE_FIELD_ID, type="datetimeFigureOut")
        _sub(field, "a:rPr", lang=language)
        _sub(field, "a:t").text = date_text
        _sub(paragraph, "a:endParaRPr", lang=language)
        return
    if spec.type == "sldNum":
        paragraph = _sub(body, "a:p")
        field = _sub(paragraph, "a:fld", id=NUMBER_FIELD_ID, type="slidenum")
        _sub(field, "a:rPr", lang=language)
        _sub(field, "a:t").text = "‹#›"
        _sub(paragraph, "a:endParaRPr", lang=language)
        return
    if spec.prompt is None:
        _sub(_sub(body, "a:p"), "a:endParaRPr", lang=language)
        return
    texts = {"title": (TITLE_PROMPT,), "subtitle": (SUBTITLE_PROMPT,),
             "text": TEXT_PROMPTS[:1], "levels": TEXT_PROMPTS}[spec.prompt]
    for level, text in enumerate(texts):
        paragraph = _sub(body, "a:p")
        if spec.prompt in ("text", "levels"):
            _sub(paragraph, "a:pPr", lvl=str(level))
        run = _sub(paragraph, "a:r")
        _sub(run, "a:rPr", lang=language)
        _sub(run, "a:t").text = text
        if level == len(texts) - 1 and spec.prompt != "text":
            _sub(paragraph, "a:endParaRPr", lang=language)


def master_xml(scale: _Scale, language: str, date_text: str, layout_rels: list[str]) -> bytes:
    root = _root("p:sldMaster", ("a", "r", "p"))
    common = _sub(root, "p:cSld")
    _sub(_sub(_sub(common, "p:bg"), "p:bgRef", idx="1001"), "a:schemeClr", val="bg1")
    tree = _shape_tree(common)
    for identifier, spec in enumerate(MASTER_PLACEHOLDERS, start=2):
        _placeholder(tree, identifier, spec, scale, language, date_text, on_master=True)
    _sub(root, "p:clrMap", bg1="lt1", tx1="dk1", bg2="lt2", tx2="dk2", accent1="accent1",
         accent2="accent2", accent3="accent3", accent4="accent4", accent5="accent5",
         accent6="accent6", hlink="hlink", folHlink="folHlink")
    layouts = _sub(root, "p:sldLayoutIdLst")
    for number, rel_id in enumerate(layout_rels, start=1):
        _sub(layouts, "p:sldLayoutId", id=str(MASTER_ID + number), r__id=rel_id)
    styles = _sub(root, "p:txStyles")
    title = _sub(styles, "p:titleStyle")
    _level(title, 1, marL=None, indent=None, size=TITLE_SIZE, major=True, space_before=0,
           line_spacing=True, bullet=False, tab=TAB_SIZE, scale=scale)
    body = _sub(styles, "p:bodyStyle")
    for level, (size, before) in enumerate(BODY_LEVELS, start=1):
        _level(body, level, marL=228600 + (level - 1) * LEVEL_STEP, indent=-228600, size=size,
               space_before=before, line_spacing=True, bullet=True, tab=TAB_SIZE, scale=scale)
    _plain_levels(_sub(styles, "p:otherStyle"), language, scale)
    return _write(root)


def layout_xml(name: str, kind: str, placeholders: tuple[Placeholder, ...], scale: _Scale,
               language: str, date_text: str) -> bytes:
    root = _root("p:sldLayout", ("a", "r", "p"), type=kind, preserve="1")
    common = _sub(root, "p:cSld", name=name)
    tree = _shape_tree(common)
    for identifier, spec in enumerate(placeholders, start=2):
        _placeholder(tree, identifier, spec, scale, language, date_text, on_master=False)
    _sub(_sub(root, "p:clrMapOvr"), "a:masterClrMapping")
    return _write(root)


def rels_xml(relationships: list[tuple[str, str, str]]) -> bytes:
    root = etree.Element("{%s}Relationships" % _RELS_NS, nsmap={None: _RELS_NS})
    for rel_id, rel_type, target in relationships:
        etree.SubElement(root, "{%s}Relationship" % _RELS_NS, Id=rel_id, Type=rel_type,
                         Target=target)
    return _write(root)


def content_types_xml(overrides: list[tuple[str, str]]) -> bytes:
    root = etree.Element("{%s}Types" % _CT_NS, nsmap={None: _CT_NS})
    for extension, content_type in (("rels", CT_RELS), ("xml", "application/xml")):
        etree.SubElement(root, "{%s}Default" % _CT_NS, Extension=extension,
                         ContentType=content_type)
    for part, content_type in overrides:
        etree.SubElement(root, "{%s}Override" % _CT_NS, PartName="/" + part,
                         ContentType=content_type)
    return _write(root)


def slide_size(size) -> tuple[int, int]:
    """``"16:9"``, ``"4:3"`` or ``(cx, cy)`` in EMU, checked against the schema's range."""
    if isinstance(size, str):
        if size not in SLIDE_SIZES:
            raise ValueError(f"slide size {size!r} is not one of {sorted(SLIDE_SIZES)} or "
                             f"(width, height) in EMU")
        return SLIDE_SIZES[size]
    try:
        width, height = (int(v) for v in size)
    except (TypeError, ValueError):
        raise ValueError(f"slide size must be '16:9', '4:3' or (width, height), not "
                         f"{size!r}") from None
    for value in (width, height):
        if not MIN_SIDE <= value <= MAX_SIDE:
            raise ValueError(f"a slide side must be {MIN_SIDE}..{MAX_SIDE} EMU, not {value}")
    return width, height


def size_format(width: int, height: int) -> str:
    return _SIZE_FORMATS.get((width, height), "Custom")


def build(width: int, height: int, *, language: str, created: datetime,
          title: str | None = None, author: str | None = None) -> bytes:
    """The whole package, as bytes: a presentation with a master, a theme and every
    layout, and no slides."""
    from .properties import app_xml, core_xml

    scale = _Scale(width, height)
    date_text = f"{created.month}/{created.day}/{created.year}"
    parts: list[tuple[str, bytes]] = []
    overrides: list[tuple[str, str]] = []

    layout_rels = [f"rId{n}" for n in range(1, len(LAYOUTS) + 1)]
    theme_rel = f"rId{len(LAYOUTS) + 1}"
    presentation = [("rId1", REL_MASTER, "slideMasters/slideMaster1.xml"),
                    ("rId2", REL_PRES_PROPS, "presProps.xml"),
                    ("rId3", REL_VIEW_PROPS, "viewProps.xml"),
                    ("rId4", REL_THEME, "theme/theme1.xml"),
                    ("rId5", REL_TABLE_STYLES, "tableStyles.xml")]
    package_rels = [("rId1", REL_DOCUMENT, "ppt/presentation.xml"),
                    ("rId2", REL_CORE, "docProps/core.xml"),
                    ("rId3", REL_APP, "docProps/app.xml")]

    parts.append(("_rels/.rels", rels_xml(package_rels)))
    parts.append(("ppt/_rels/presentation.xml.rels", rels_xml(presentation)))
    parts.append(("ppt/presentation.xml", presentation_xml(width, height, language, "rId1")))
    overrides.append(("ppt/presentation.xml", CT_PRESENTATION))
    parts.append(("ppt/slideMasters/_rels/slideMaster1.xml.rels", rels_xml(
        [(rel, REL_LAYOUT, f"../slideLayouts/slideLayout{n}.xml")
         for n, rel in enumerate(layout_rels, start=1)]
        + [(theme_rel, REL_THEME, "../theme/theme1.xml")])))
    parts.append(("ppt/slideMasters/slideMaster1.xml",
                  master_xml(scale, language, date_text, layout_rels)))
    overrides.append(("ppt/slideMasters/slideMaster1.xml", CT_MASTER))
    for number, (name, kind, placeholders) in enumerate(LAYOUTS, start=1):
        path = f"ppt/slideLayouts/slideLayout{number}.xml"
        parts.append((f"ppt/slideLayouts/_rels/slideLayout{number}.xml.rels", rels_xml(
            [("rId1", REL_MASTER, "../slideMasters/slideMaster1.xml")])))
        parts.append((path, layout_xml(name, kind, placeholders, scale, language, date_text)))
        overrides.append((path, CT_LAYOUT))
    parts.append(("ppt/theme/theme1.xml", theme_xml()))
    overrides.append(("ppt/theme/theme1.xml", CT_THEME))
    for path, data, content_type in (("ppt/presProps.xml", pres_props_xml(), CT_PRES_PROPS),
                                     ("ppt/viewProps.xml", view_props_xml(), CT_VIEW_PROPS),
                                     ("ppt/tableStyles.xml", table_styles_xml(),
                                      CT_TABLE_STYLES)):
        parts.append((path, data))
        overrides.append((path, content_type))
    parts.append(("docProps/core.xml", core_xml(title=title, author=author, created=created)))
    overrides.append(("docProps/core.xml", CT_CORE))
    parts.append(("docProps/app.xml", app_xml(format_name=size_format(width, height),
                                              fonts=FONTS_USED, themes=(THEME_NAME,),
                                              titles=())))
    overrides.append(("docProps/app.xml", CT_APP))

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for path, data in [("[Content_Types].xml", content_types_xml(overrides))] + parts:
            entry = zipfile.ZipInfo(path, date_time=(1980, 1, 1, 0, 0, 0))
            entry.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(entry, data)
    return buffer.getvalue()


__all__ = ["BASE_SIZE", "LAYOUTS", "SLIDE_SIZES", "build", "size_format", "slide_size"]
