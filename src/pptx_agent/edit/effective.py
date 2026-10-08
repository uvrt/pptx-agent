"""Formatting as PowerPoint draws it: inherited values resolved, autofit applied.

A run that says nothing about its size takes it from its paragraph, its shape's list style,
the layout and master placeholders' list styles, the master's text style and the
presentation's default, in that order (:mod:`.inherit`); a typeface may then be the
theme's (``+mn-lt``).  A text body's autofit comes the same way, from its own ``a:bodyPr``
outwards, and only ``normAutofit`` carries a ``fontScale``.

**PowerPoint draws exactly the stored scale.**  ``normAutofit`` without a ``fontScale``
draws the text at full size even when it overflows: PowerPoint computes the scale when it
edits the text, stores it, and on opening a file draws what is stored.  So an effective
size here is the inherited size times the stored scale -- never a scale this library
guesses.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..oxml.xml import Element, get_int, local_name, qn
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


__all__ = ["Autofit", "EffectiveParagraph", "autofit", "list_styles", "paragraph",
           "paragraph_sources", "size", "typeface", "write_autofit"]
