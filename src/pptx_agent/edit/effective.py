"""Formatting as PowerPoint draws it: inherited values resolved, autofit applied.

A run that says nothing about its size takes it from its paragraph, its shape's list style,
the layout and master placeholders' list styles, the master's text style and the
presentation's default, in that order (:mod:`.inherit`); a typeface may then be the
theme's (``+mn-lt``).  Its colour comes the same way (the ``a:solidFill`` of ``a:rPr`` and
each ``a:defRPr``, ECMA-376 Part 1, 21.1.2.3), with two more layers between a shape's own
list style and the master's ``otherStyle`` for text that is not in a placeholder: a
shape's ``p:style/a:fontRef`` colour (20.1.4.1.17) and a table cell's ``a:tcTxStyle``
(20.1.4.2.30) from its table style's regions -- where pptx2svg measured them to sit.  A
hyperlink's text with no colour of its own is the theme's ``hlink``.  A text body's autofit comes the same way, from its own ``a:bodyPr``
outwards, and only ``normAutofit`` carries a ``fontScale``.

**PowerPoint draws exactly the stored scale.**  ``normAutofit`` without a ``fontScale``
draws the text at full size even when it overflows: PowerPoint computes the scale when it
edits the text, stores it, and on opening a file draws what is stored.  So an effective
size here is the inherited size times the stored scale -- never a scale this library
guesses.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass

from ..oxml.xml import FILL_TAGS, Element, find, get_int, local_name, qn
from . import inherit

AUTOFIT_TAGS = {"noAutofit": "none", "normAutofit": "normal", "spAutoFit": "shape"}
AUTOFIT_NAMES = {name: tag for tag, name in AUTOFIT_TAGS.items()}
_BULLET_TAGS = ("buNone", "buChar", "buAutoNum", "buBlip")
_ALIGNMENT = {"l": "left", "ctr": "center", "r": "right", "just": "justify",
              "dist": "distributed", "justLow": "justify_low", "thaiDist": "thai_distributed"}


@dataclass(frozen=True)
class Autofit:
    """A text body's autofit as it applies: ``mode`` ``"none"``, ``"normal"`` (shrink text
    on overflow) or ``"shape"`` (resize shape to fit text); ``font_scale`` and
    ``line_spacing_reduction`` the stored ``normAutofit`` values (1.0 and 0.0 when none is
    stored -- and then PowerPoint draws the text full size)."""

    mode: str
    font_scale: float
    line_spacing_reduction: float


def body_chain(host) -> list[Element | None]:
    chain = getattr(host, "_body_chain", None)
    if chain is not None:
        return chain()
    body = host._text_body(False)
    return [None if body is None else body.find(qn("a:bodyPr"))]


def autofit(host) -> Autofit:
    """Merged outward-in, as pptx2svg merges ``a:bodyPr``: the nearest body that names an
    autofit decides the mode, the nearest ``normAutofit`` that states a scale the scale."""
    mode = None
    scale = None
    reduction = None
    for properties in body_chain(host):
        if properties is None:
            continue
        for child in properties:
            name = local_name(child) if isinstance(child.tag, str) else ""
            if name not in AUTOFIT_TAGS:
                continue
            if mode is None:
                mode = AUTOFIT_TAGS[name]
            if name == "normAutofit":
                if scale is None and child.get("fontScale") is not None:
                    scale = int(child.get("fontScale")) / 100000
                if reduction is None and child.get("lnSpcReduction") is not None:
                    reduction = int(child.get("lnSpcReduction")) / 100000
    mode = mode or "none"
    if mode != "normal":
        return Autofit(mode, 1.0, 0.0)
    return Autofit(mode, 1.0 if scale is None else scale, 0.0 if reduction is None else reduction)


def list_styles(host) -> list[Element | None]:
    chain = getattr(host, "_list_styles", None)
    if chain is not None:
        return chain()
    body = host._text_body(False)
    return [None if body is None else body.find(qn("a:lstStyle"))]


def _level(paragraph: Element) -> int:
    properties = paragraph.find(qn("a:pPr"))
    return max(0, min(8, get_int(properties, "lvl", 0) or 0)) if properties is not None else 0


def paragraph_sources(host, paragraph: Element) -> list[Element]:
    """The paragraph's own ``a:pPr`` and every list-style level it inherits, nearest first."""
    own = paragraph.find(qn("a:pPr"))
    sources = [own] if own is not None else []
    return sources + inherit.level_properties(list_styles(host), _level(paragraph))


def _run_properties(host, paragraph: Element, run: Element | None) -> list[Element]:
    """``a:rPr`` of the run, then each ``a:defRPr`` it inherits, nearest first."""
    out = []
    if run is not None:
        properties = run.find(qn("a:rPr")) if local_name(run) != "endParaRPr" else run
        if properties is not None:
            out.append(properties)
    for source in paragraph_sources(host, paragraph):
        node = source.find(qn("a:defRPr"))
        if node is not None:
            out.append(node)
    return out


def size(host, paragraph: Element, run: Element | None, *, scaled: bool = True) -> float:
    """A run's size in points -- with the stored autofit scale applied when ``scaled``."""
    points = inherit.DEFAULT_SIZE
    for properties in _run_properties(host, paragraph, run):
        if properties.get("sz"):
            points = int(properties.get("sz")) / 100
            break
    if scaled:
        points *= autofit(host).font_scale
    return round(points, 2)


def typeface(host, paragraph: Element, run: Element | None, script: str = "latin") -> str | None:
    """A run's typeface, with the theme's ``+mn-lt``/``+mj-lt`` resolved; the theme's minor
    face when nothing on the way names one -- Arial, in a shape that is not a placeholder."""
    from .theme import resolve_font, theme_for

    named = None
    for properties in _run_properties(host, paragraph, run):
        font = properties.find(qn(f"a:{script}"))
        if font is not None and font.get("typeface"):
            named = font.get("typeface")
            break
    if named is None and script == "latin" and getattr(host, "_plain_shape", False):
        # A shape that is not a placeholder, with no face named anywhere on the way, is
        # drawn in PowerPoint's application default, not the theme's (pptx2svg measured
        # it: ArialMT in the PDF, Arial's advances).
        return "Arial"
    theme = theme_for(host._theme_source()) if hasattr(host, "_theme_source") else None
    if theme is None:
        return named
    if named is None:
        return theme.fonts.minor if script == "latin" else theme.fonts.minor_east_asian
    return resolve_font(named, theme)


_FILLS = frozenset(qn(tag) for tag in FILL_TAGS)
_RUN_PROPERTIES = ("rPr", "endParaRPr", "defRPr")
#: What a run is drawn in when nothing on the way states a colour: black, as pptx2svg draws
#: a run with none (a master's text styles normally state ``tx1``, so this is rare).
DEFAULT_COLOR = "000000"
#: Unstated: the label :func:`color_source` gives :data:`DEFAULT_COLOR`.
UNSTATED = "nothing on the way (drawn black)"


def _color_layers(host, paragraph: Element, run: Element | None) -> list[tuple[Element, str | None]]:
    """Every element that may state the run's colour, nearest first, with a label for
    those :func:`describe` cannot name from where they sit: the run's ``a:rPr``, the
    paragraph's and each list-style level's ``a:defRPr``, and -- after the shape's own list
    style, before the master's -- the host's text defaults (``fontRef``, ``tcTxStyle``)."""
    layers: list[tuple[Element, str | None]] = []
    if run is not None:
        own = run.find(qn("a:rPr")) if local_name(run) != "endParaRPr" else run
        if own is not None:
            layers.append((own, None))
    properties = paragraph.find(qn("a:pPr"))
    node = properties.find(qn("a:defRPr")) if properties is not None else None
    if node is not None:
        layers.append((node, None))
    defaults = getattr(host, "_text_defaults", None)
    extras = list(defaults()) if defaults is not None else []
    level = _level(paragraph)
    for index, style in enumerate(list_styles(host)):
        if index == 1:
            layers += extras
            extras = []
        node = style.find(qn(f"a:lvl{level + 1}pPr")) if style is not None else None
        node = node.find(qn("a:defRPr")) if node is not None else None
        if node is not None:
            layers.append((node, None))
    return layers + extras


def color_source(host, paragraph: Element, run: Element | None):
    """``(colour as stated, where it comes from)``: the nearest stated colour, unresolved
    (``Color("tx1")``), and a label such as ``"the master's bodyStyle"``.  The colour is
    ``None`` with a label when the nearest fill is not solid (``"... (gradFill)"``), and
    with :data:`UNSTATED` when nothing states one."""
    from .color import Color

    for holder, label in _color_layers(host, paragraph, run):
        if local_name(holder) in _RUN_PROPERTIES:
            fill = next((child for child in holder if child.tag in _FILLS), None)
            if fill is None:
                if holder.find(qn("a:hlinkClick")) is not None \
                        and local_name(holder) == "rPr":
                    return Color("scheme", "hlink"), "the hyperlink (theme hlink)"
                continue
            where = label or describe(holder)
            if local_name(fill) != "solidFill":
                return None, f"{where} ({local_name(fill)})"
            return Color.from_element(fill), where
        stated = Color.from_element(holder)
        if stated is not None:
            return stated, label or describe(holder)
    return None, UNSTATED


def color(host, paragraph: Element, run: Element | None):
    """A run's colour as drawn, resolved (a :class:`~pptx_agent.Color` of kind ``"rgb"``):
    the nearest stated colour (:func:`color_source`), a theme colour through the colour
    map and theme its slide uses, modifiers applied as :meth:`Color.resolve
    <pptx_agent.Color.resolve>` applies them; black when nothing states one; ``None`` when
    the text is not drawn in one solid colour (a gradient, no fill)."""
    from .color import Color

    stated, label = color_source(host, paragraph, run)
    if stated is None:
        return Color.rgb(DEFAULT_COLOR) if label == UNSTATED else None
    if stated.kind == "rgb" and not stated.transforms:
        return Color.rgb(stated.value)
    source = host._theme_source() if hasattr(host, "_theme_source") else None
    if source is None:
        return None
    from .theme import theme_for

    try:
        return Color.rgb(theme_for(source).resolve(stated))
    except ValueError:
        return None


def size_source(host, paragraph: Element, run: Element | None) -> str:
    """Where a run's :func:`size` comes from: ``"the run"``, ``"the paragraph"``,
    ``"layout 'Title and Content' body placeholder"``, ``"the master's bodyStyle"``...,
    with the stored autofit scale when it shrinks the text (``"... at 90% autofit"``)."""
    where = next((describe(properties) for properties in _run_properties(host, paragraph, run)
                  if properties.get("sz")), f"PowerPoint's default {inherit.DEFAULT_SIZE:g} pt")
    scale = autofit(host).font_scale
    return where if scale == 1.0 else f"{where} at {round(scale * 100):g}% autofit"


#: The kinds of text a minimum size is set for (``deck_format(min_text_size={...})``):
#: ``body`` the text of shapes, text boxes and placeholders (groups looked into), ``table``
#: table cells' text, ``chart`` a chart's title, axis titles, axis labels, data labels and
#: legend (:mod:`.chart_text`).
TEXT_KINDS = ("body", "table", "chart")


def text_minimums(value) -> dict[str, float]:
    """A minimum text size as ``{kind: points}``: a number is every kind's; a mapping names
    kinds of :data:`TEXT_KINDS` (a kind it leaves out, or gives ``None``, is not checked).
    ``None`` is no minimum (``{}``).  Raises ``ValueError`` for an unknown kind or a size
    outside 1 to 400 points."""
    if value is None:
        return {}
    if isinstance(value, Mapping):
        unknown = [kind for kind in value if kind not in TEXT_KINDS]
        if unknown:
            raise ValueError(f"min_text_size has no kind {unknown[0]!r}; the kinds are "
                             f"{', '.join(TEXT_KINDS)} (a kind left out is not checked)")
        if not value:
            raise ValueError("min_text_size is an empty mapping; name body, table or chart, "
                             "or give None for no minimum")
        return {kind: _points(points) for kind, points in value.items() if points is not None}
    points = _points(value)
    return {kind: points for kind in TEXT_KINDS}


def _points(value) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)) \
            or not 1 <= float(value) <= 400:
        raise ValueError(f"min_text_size {value!r}: points between 1 and 400, a mapping of "
                         "body, table and chart to points, or None for no minimum")
    return float(value)


@dataclass(frozen=True)
class SmallText:
    """Text drawn below a minimum size (:meth:`Slide.text_below
    <pptx_agent.Slide.text_below>`): the shape, table cell or chart's ``address``, the
    effective ``size`` in points and ``source``, where that size comes from
    (:func:`size_source`); its ``kind`` (:data:`TEXT_KINDS`), the ``shape`` that holds it (a
    cell's table, else the address) and, for a chart, the ``part`` of its text
    (``"axis_labels"``..., :data:`.chart_text.PARTS`)."""

    address: str
    size: float
    source: str
    kind: str = "body"
    shape: str | None = None
    part: str | None = None

    def to_json(self) -> dict:
        return {"shape": self.address, "size": self.size, "from": self.source}


def text_below(frames, minimum: float, kind: str = "body") -> list[SmallText]:
    """Each ``(address, text frame)``'s text drawn below ``minimum`` points, once per
    address, size and source, smallest first within an address.  A frame may be given as
    ``(address, text frame, shape)`` to name the shape holding it (a cell's table)."""
    out: list[SmallText] = []
    for entry in frames:
        address, frame = entry[0], entry[1]
        holder = entry[2] if len(entry) > 2 else address
        host = frame._host()
        found: dict[tuple[float, str], SmallText] = {}
        for paragraph in frame.paragraphs:
            element = paragraph._element()
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                points = size(host, element, run._element())
                if points >= minimum:
                    continue
                source = size_source(host, element, run._element())
                found.setdefault((points, source),
                                 SmallText(address, points, source, kind, holder))
        out += sorted(found.values(), key=lambda item: item.size)
    return out


def chart_text_below(address: str, shape, minimum: float) -> list[SmallText]:
    """A chart shape's text drawn below ``minimum`` points, once per part, size and source,
    smallest first within a part (:func:`.chart_text.chart_text`)."""
    from .chart_text import PARTS, chart_text

    found: dict[tuple[str, float, str], SmallText] = {}
    for text in chart_text(shape):
        if text.size < minimum:
            found.setdefault((text.part, text.size, text.source),
                             SmallText(address, text.size, text.source, "chart", address,
                                       text.part))
    return sorted(found.values(), key=lambda item: (PARTS.index(item.part), item.size))


def small_text_json(found: list[SmallText], minimums: dict[str, float]) -> list[dict]:
    """``check``'s ``small_text``: a shape's text one entry per size and source, as
    :meth:`SmallText.to_json`; a **table one entry** -- ``{"table": address, "cells": how
    many cells, "size": the smallest, "min", "from": where the smallest comes from}``; a
    **chart one entry** -- ``{"chart": address, "min", "parts": [{"part", "size", "from"}]}``,
    each part's smallest size, the part as ``edit_chart``'s ``set_text_style`` names it."""
    out: list[dict] = []
    groups: dict[tuple[str, str], dict] = {}
    for item in found:
        if item.kind == "body":
            out.append(item.to_json())
            continue
        key = (item.kind, item.shape or item.address)
        entry = groups.get(key)
        if item.kind == "table":
            if entry is None:
                entry = groups[key] = {"table": key[1], "cells": set(), "size": item.size,
                                       "min": minimums.get("table"), "from": item.source}
                out.append(entry)
            entry["cells"].add(item.address)
            if item.size < entry["size"]:
                entry.update(size=item.size, **{"from": item.source})
        else:
            if entry is None:
                entry = groups[key] = {"chart": key[1], "min": minimums.get("chart"),
                                       "parts": []}
                out.append(entry)
            if not any(part["part"] == item.part for part in entry["parts"]):
                entry["parts"].append({"part": item.part, "size": item.size,
                                       "from": item.source})
    for entry in out:
        if isinstance(entry.get("cells"), set):
            entry["cells"] = len(entry["cells"])
    return out


_STYLES = ("titleStyle", "bodyStyle", "otherStyle", "notesStyle")
_SHEETS = {"sld": "slide", "sldLayout": "layout", "sldMaster": "master",
           "notes": "notes page", "notesMaster": "notes master"}


def describe(properties: Element) -> str:
    """Where a run-properties element (``a:rPr``, ``a:defRPr``) sits, in words: ``"the
    run"``, ``"the paragraph"``, ``"the shape's list style"``, ``"the cell's list
    style"``, ``"layout 'Title and Content' body placeholder"``, ``"the master's
    bodyStyle"``, ``"the presentation's default text style"`` -- with ``", level 2"`` for a
    level below the first."""
    if local_name(properties) in ("rPr", "endParaRPr"):
        return "the run"
    holder = properties.getparent()
    holder_name = local_name(holder) if holder is not None else ""
    if holder_name == "pPr":
        return "the paragraph"
    level = holder_name[3:-3] if holder_name.startswith("lvl") else ""
    suffix = f", level {level}" if level not in ("", "1") else ""
    root = properties.getroottree().getroot()
    part = local_name(root)
    if part == "presentation":
        return "the presentation's default text style" + suffix
    ancestors = [local_name(node) for node in properties.iterancestors()]
    sheet = _SHEETS.get(part, part)
    data = find(root, "p:cSld")
    name = data.get("name") if data is not None else None
    named = f"{sheet} {name!r}" if name else f"the {sheet}'s"
    style = next((node for node in ancestors if node in _STYLES), None)
    if style is not None:
        return f"{named} {style}{suffix}"
    if "tc" in ancestors:
        return "the cell's list style" + suffix
    shape = next((node for node in properties.iterancestors()
                  if local_name(node) in ("sp", "graphicFrame")), None)
    placeholder = inherit.placeholder_of(shape) if shape is not None else None
    if part == "sld" or placeholder is None:
        return "the shape's list style" + suffix
    return f"{named} {placeholder.get('type') or 'content'} placeholder{suffix}"


@dataclass(frozen=True)
class EffectiveParagraph:
    """A paragraph's formatting as it is drawn, inherited values resolved.

    ``size`` (points, autofit's stored scale applied) and ``font`` are the first run's --
    or the end-of-paragraph mark's, for an empty paragraph; ``alignment`` is ``"left"``,
    ``"center"``...; ``bullet`` the :class:`~pptx_agent.Bullet` drawn, ``None`` for none;
    ``margin_left`` and ``indent`` EMU; ``line_spacing`` a multiple of single spacing or
    ``line_spacing_points`` exact points; ``space_before`` and ``space_after`` EMU (as
    :attr:`Paragraph.space_before <pptx_agent.Paragraph.space_before>`), or
    ``space_before_lines`` and ``space_after_lines`` fractions of a line when the style
    gives them as a percentage.

    For example::

        paragraph.effective.size                    # 28.0
        paragraph.effective.bullet.char             # '•'
    """

    level: int
    size: float
    font: str | None
    alignment: str
    bullet: object
    margin_left: int
    indent: int
    line_spacing: float | None
    line_spacing_points: float | None
    space_before: int | None
    space_before_lines: float | None
    space_after: int | None
    space_after_lines: float | None
    #: The first run's colour as drawn, resolved (:attr:`Run.effective_color
    #: <pptx_agent.Run.effective_color>`).
    color: object = None


def paragraph(host, element: Element) -> EffectiveParagraph:
    from .text import Bullet

    sources = paragraph_sources(host, element)
    runs = [node for node in element if local_name(node) in ("r", "fld")]
    first = next((node for node in runs if (node.findtext(qn("a:t")) or "")), None)
    if first is None:
        first = element.find(qn("a:endParaRPr"))
        if first is None and runs:
            first = runs[0]

    def attribute(name: str, default=None):
        for source in sources:
            if source.get(name) is not None:
                return source.get(name)
        return default

    def spacing(tag: str) -> tuple[float | None, float | None]:
        for source in sources:
            node = source.find(qn(f"a:{tag}"))
            if node is None:
                continue
            points = node.find(qn("a:spcPts"))
            if points is not None:
                return int(points.get("val", "0")) / 100, None
            percent = node.find(qn("a:spcPct"))
            if percent is not None:
                return None, int(percent.get("val", "0")) / 100000
        return None, None

    bullet = None
    for source in sources:
        node = next((child for child in source if isinstance(child.tag, str)
                     and local_name(child) in _BULLET_TAGS), None)
        if node is None:
            continue
        name = local_name(node)
        if name == "buChar":
            bullet = Bullet("char", char=node.get("char"))
        elif name == "buAutoNum":
            bullet = Bullet("number", scheme=node.get("type"),
                            start_at=get_int(node, "startAt"))
        elif name == "buBlip":
            bullet = Bullet("picture")
        break
    line_points, line_multiple = spacing("lnSpc")
    before_points, before_lines = spacing("spcBef")
    after_points, after_lines = spacing("spcAft")
    # spcPts is hundredths of a point; 127 EMU each.
    before_points = None if before_points is None else round(before_points * 12700)
    after_points = None if after_points is None else round(after_points * 12700)
    return EffectiveParagraph(
        level=_level(element),
        size=size(host, element, first),
        font=typeface(host, element, first),
        alignment=_ALIGNMENT.get(attribute("algn", "l"), attribute("algn", "l")),
        bullet=bullet,
        margin_left=int(attribute("marL", 0)),
        indent=int(attribute("indent", 0)),
        line_spacing=line_multiple,
        line_spacing_points=line_points,
        space_before=before_points,
        space_before_lines=before_lines,
        space_after=after_points,
        space_after_lines=after_lines,
        color=color(host, element, first),
    )


def write_autofit(body: Element, mode: str | None, font_scale: float | None = None,
                  keep_scale: bool = False) -> None:
    """Set the body's own autofit: ``mode`` one of ``"none"``, ``"normal"``, ``"shape"``, or
    ``None`` to inherit again; ``font_scale`` for ``"normal"``."""
    from ..oxml.xml import make, replace_choice, subelement

    properties = subelement(body, "a:bodyPr")
    tags = tuple(f"a:{tag}" for tag in AUTOFIT_TAGS)
    if mode is None:
        replace_choice(properties, tags, None)
        return
    if mode not in AUTOFIT_NAMES:
        raise ValueError(f"autofit is one of {sorted(AUTOFIT_NAMES)} or None, not {mode!r}")
    old = next((child for child in properties if isinstance(child.tag, str)
                and local_name(child) == "normAutofit"), None)
    node = make(f"a:{AUTOFIT_NAMES[mode]}")
    if mode == "normal":
        if font_scale is not None:
            if not 0 < font_scale <= 1:
                raise ValueError("font_scale is a fraction in (0, 1]")
            if round(font_scale * 100000) != 100000:
                node.set("fontScale", str(round(font_scale * 100000)))
            if old is not None and old.get("lnSpcReduction") and keep_scale:
                node.set("lnSpcReduction", old.get("lnSpcReduction"))
        elif keep_scale and old is not None:
            for name in ("fontScale", "lnSpcReduction"):
                if old.get(name):
                    node.set(name, old.get(name))
    replace_choice(properties, tags, node)


__all__ = ["DEFAULT_COLOR", "TEXT_KINDS", "UNSTATED", "Autofit", "EffectiveParagraph", "SmallText",
           "autofit", "chart_text_below", "color", "color_source", "describe", "list_styles",
           "paragraph", "paragraph_sources", "size", "size_source", "small_text_json",
           "text_below", "text_minimums", "typeface", "write_autofit"]
