"""Fills and outlines: reading them, and writing them back as DrawingML.

Both are written into whichever properties element owns them -- a shape's ``p:spPr``, a table
cell's ``a:tcPr``, a run's ``a:rPr`` -- through the same functions, so the rules hold
everywhere:

* A fill is one member of ``EG_FillProperties`` (``noFill``/``solidFill``/``gradFill``/
  ``blipFill``/``pattFill``/``grpFill``).  Setting one removes whichever was there, and the new
  element goes to its schema position -- after the geometry, before ``a:ln``.
* Removing an explicit fill (``None``) means "inherit": from the shape style, the layout
  placeholder, the table style.  That is different from ``"none"``, which is an explicit
  ``a:noFill``.
* Colours are :class:`~pptx_agent.edit.color.Color`, so theme colours stay theme colours.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Iterable, Sequence

from ..oxml.xml import (
    FILL_TAGS,
    LINE_FILL_TAGS,
    Element,
    get_int,
    local_name,
    make,
    qn,
    remove,
    replace_choice,
    set_attr,
    set_int,
    subelement,
)
from .color import Color

ANGLE_UNIT = 60000
POSITION_UNIT = 100000


@dataclass(frozen=True)
class GradientStop:
    """One stop of a gradient fill: where (0.0 to 1.0) and what colour.

    For example::

        [stop.color for stop in shape.fill.stops]
    """

    #: 0.0 .. 1.0 along the gradient.
    position: float
    color: Color


@dataclass(frozen=True)
class Fill:
    """An explicit fill, as read from the file.

    ``kind`` is ``"none"``, ``"solid"``, ``"gradient"``, ``"image"``, ``"pattern"`` or
    ``"group"``.  Only the fields that apply to the kind are set.

    For example::

        shape.fill.kind, shape.fill.color             # ('solid', Color('accent1'))
    """

    kind: str
    color: Color | None = None
    stops: tuple[GradientStop, ...] = ()
    #: Linear gradient angle in degrees, clockwise from left-to-right.
    angle: float | None = None
    #: ``"circle"``, ``"rect"`` or ``"shape"`` for a path gradient.
    path: str | None = None
    #: Relationship id of an image fill's picture.
    image_rel_id: str | None = None
    #: Part name of an image fill's picture, when resolvable.
    image_part: str | None = None
    #: Preset of a pattern fill, e.g. ``"pct50"``.
    pattern: str | None = None


def read_fill(container: Element | None, choices: Iterable[str] = FILL_TAGS) -> Fill | None:
    """The explicit fill on a properties element, or ``None`` when it inherits."""
    if container is None:
        return None
    wanted = {qn(tag) for tag in choices}
    node = next((child for child in container if child.tag in wanted), None)
    if node is None:
        return None
    name = local_name(node)
    if name == "noFill":
        return Fill("none")
    if name == "solidFill":
        return Fill("solid", color=Color.from_element(node))
    if name == "grpFill":
        return Fill("group")
    if name == "pattFill":
        return Fill("pattern", pattern=node.get("prst"),
                    color=Color.from_element(node.find(qn("a:fgClr"))))
    if name == "gradFill":
        stops = []
        gs_list = node.find(qn("a:gsLst"))
        for stop in [] if gs_list is None else gs_list.findall(qn("a:gs")):
            color = Color.from_element(stop)
            if color is not None:
                stops.append(GradientStop((get_int(stop, "pos", 0) or 0) / POSITION_UNIT, color))
        linear = node.find(qn("a:lin"))
        path = node.find(qn("a:path"))
        return Fill(
            "gradient",
            stops=tuple(stops),
            angle=None if linear is None else (get_int(linear, "ang", 0) or 0) / ANGLE_UNIT,
            path=None if path is None else path.get("path"),
        )
    if name == "blipFill":
        blip = node.find(qn("a:blip"))
        return Fill("image", image_rel_id=None if blip is None else blip.get(qn("r:embed")))
    return Fill(name)


# -- building fills ------------------------------------------------------------------------


def solid_fill(color: str | Color) -> Element:
    element = make("a:solidFill")
    element.append(Color.parse(color).to_element())
    return element


def no_fill() -> Element:
    return make("a:noFill")


def gradient_fill(
    stops: Sequence[tuple[float, str | Color]] | Sequence[str | Color],
    *,
    angle: float | None = 90.0,
    path: str | None = None,
) -> Element:
    """A ``gradFill``.

    ``stops`` are ``(position, colour)`` pairs with positions 0..1, or bare colours spread
    evenly.  ``angle`` is in degrees (0 = left to right, 90 = top to bottom) for a linear
    gradient; pass ``path="circle"``/``"rect"``/``"shape"`` instead for a radial one.
    """
    normalised = _normalise_stops(stops)
    element = make("a:gradFill", rotWithShape="1")
    gs_list = make("a:gsLst")
    for position, color in normalised:
        stop = make("a:gs", pos=str(round(position * POSITION_UNIT)))
        stop.append(color.to_element())
        gs_list.append(stop)
    element.append(gs_list)
    if path is not None:
        if path not in {"circle", "rect", "shape"}:
            raise ValueError("path must be 'circle', 'rect' or 'shape'")
        shade = make("a:path", path=path)
        shade.append(make("a:fillToRect", l="50000", t="50000", r="50000", b="50000"))
        element.append(shade)
    elif angle is not None:
        element.append(make("a:lin", ang=str(round(angle * ANGLE_UNIT) % (360 * ANGLE_UNIT)),
                            scaled="0"))
    return element


def image_fill(rel_id: str, *, stretch: bool = True) -> Element:
    element = make("a:blipFill", rotWithShape="1")
    element.append(make("a:blip", r__embed=rel_id))
    if stretch:
        stretch_node = make("a:stretch")
        stretch_node.append(make("a:fillRect"))
        element.append(stretch_node)
    return element


def fill_element(spec: "str | Color | Fill | None") -> Element | None:
    """Turn a user-facing fill spec into an element.

    ``"none"`` -> ``a:noFill``; a colour string or :class:`Color` -> ``a:solidFill``; ``None``
    -> nothing (inherit).  Gradient and image fills have their own setters, since they take
    more than one argument.
    """
    if spec is None:
        return None
    if isinstance(spec, str) and spec.strip().lower() == "none":
        return no_fill()
    if isinstance(spec, (str, Color)):
        return solid_fill(spec)
    raise TypeError(f"cannot use {spec!r} as a fill; use set_gradient_fill/set_image_fill")


def write_fill(container: Element, element: Element | None, choices: Iterable[str] = FILL_TAGS) -> None:
    replace_choice(container, choices, element)


def _normalise_stops(stops) -> list[tuple[float, Color]]:
    items = list(stops)
    if len(items) < 2:
        raise ValueError("a gradient needs at least two stops")
    if all(isinstance(item, (str, Color)) for item in items):
        last = len(items) - 1
        return [(index / last, Color.parse(item)) for index, item in enumerate(items)]
    result = []
    for item in items:
        position, color = item
        if not 0.0 <= float(position) <= 1.0:
            raise ValueError(f"gradient stop position {position} is outside 0..1")
        result.append((float(position), Color.parse(color)))
    return result


# -- outlines ------------------------------------------------------------------------------

#: ``ST_PresetLineDashVal``.
DASH_STYLES = frozenset({
    "solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot", "sysDash",
    "sysDot", "sysDashDot", "sysDashDotDot",
})
#: ``ST_LineEndType``.
ARROWHEADS = frozenset({"none", "triangle", "stealth", "diamond", "oval", "arrow"})
#: ``ST_LineEndWidth`` / ``ST_LineEndLength``.
ARROWHEAD_SIZES = frozenset({"sm", "med", "lg"})
#: ``ST_LineCap``.
LINE_CAPS = frozenset({"rnd", "sq", "flat"})


@dataclass(frozen=True)
class Arrowhead:
    """The decoration at one end of a line: ``type`` (``triangle``, ``stealth``,
    ``oval``...), ``width`` and ``length`` (``sm``, ``med``, ``lg``).

    For example::

        arrow.line.tail                              # Arrowhead(type='triangle', ...)
    """

    type: str
    width: str | None = None
    length: str | None = None


class LineFormat:
    """An outline: ``a:ln`` on a shape, or one of ``a:lnL``/``lnR``/``lnT``/``lnB`` on a cell.

    The line element is located afresh on every call through ``locate`` -- the facade holds
    no XML, so it stays valid across undo.  ``before_change`` is the owning shape's undo
    checkpoint.  Widths are EMU (12,700 per point).

    For example::

        shape.line.width = 19050; shape.line.color = "accent2"; shape.line.dash = "dash"
    """

    def __init__(
        self,
        locate: Callable[[bool], Element | None],
        before_change: Callable[[], None],
        after_change: Callable[[], None],
    ) -> None:
        self._locate = locate
        self._before_change = before_change
        self._after_change = after_change

    # -- reading ---------------------------------------------------------------------------

    @property
    def exists(self) -> bool:
        """Whether the element carries an explicit outline at all.

        For example::

            shape.line.exists
        """
        return self._locate(False) is not None

    @property
    def width(self) -> int | None:
        """Width in EMU (12,700 per point), or ``None`` when inherited. Settable.

        For example::

            shape.line.width = 12700                      # 1 pt
        """
        return get_int(self._locate(False), "w")

    @property
    def fill(self) -> Fill | None:
        """The line's explicit fill, or ``None`` when inherited.

        For example::

            shape.line.fill.kind
        """
        return read_fill(self._locate(False), LINE_FILL_TAGS)

    @property
    def color(self) -> Color | None:
        """The line's solid colour, or ``None``. Settable, in theme colours too.

        For example::

            shape.line.color = "accent1 lumMod=75%"
        """
        fill = self.fill
        return fill.color if fill is not None and fill.kind == "solid" else None

    @property
    def visible(self) -> bool | None:
        """``False`` for an explicit ``a:noFill`` outline, ``None`` when inherited.

        For example::

            shape.line.visible = False
        """
        fill = self.fill
        if fill is None:
            return None
        return fill.kind != "none"

    @property
    def dash(self) -> str | None:
        """The dash preset (``solid``, ``dash``, ``sysDot``...), or ``None``. Settable.

        For example::

            shape.line.dash = "dash"
        """
        line = self._locate(False)
        node = None if line is None else line.find(qn("a:prstDash"))
        return None if node is None else node.get("val")

    @property
    def cap(self) -> str | None:
        """The line cap: ``rnd``, ``sq`` or ``flat``. Settable.

        For example::

            shape.line.cap = "rnd"
        """
        line = self._locate(False)
        return None if line is None else line.get("cap")

    @property
    def head(self) -> Arrowhead | None:
        """The arrowhead at the line's **start** -- where a connector begins, its
        ``begin_connection`` -- or ``None``.  Settable with a type name.  OOXML's name
        (``a:headEnd``); :attr:`start` is the same thing.

        For example::

            arrow.line.head = "oval"
        """
        return self._arrowhead("a:headEnd")

    @property
    def tail(self) -> Arrowhead | None:
        """The arrowhead at the line's **end** -- where a connector ends, its
        ``end_connection`` -- or ``None``.  Settable with a type name.  An arrow pointing
        at the shape it ends on is a ``tail``; :attr:`end` is the same thing.

        For example::

            arrow.line.tail = "triangle"
        """
        return self._arrowhead("a:tailEnd")

    @property
    def start(self) -> Arrowhead | None:
        """The arrowhead where the line starts (its begin): :attr:`head` by its plain name.

        For example::

            arrow.line.start = "oval"            # the same as arrow.line.head = "oval"
        """
        return self.head

    @start.setter
    def start(self, value: "Arrowhead | str | None") -> None:
        self.head = value

    @property
    def end(self) -> Arrowhead | None:
        """The arrowhead where the line ends -- the arrow's point for a connector from A
        to B: :attr:`tail` by its plain name.

        For example::

            arrow.line.end = "triangle"          # points at the end_connection shape
        """
        return self.tail

    @end.setter
    def end(self, value: "Arrowhead | str | None") -> None:
        self.tail = value

    # -- writing ---------------------------------------------------------------------------

    @width.setter
    def width(self, emu: int | None) -> None:
        if emu is not None and not 0 <= int(emu) <= 20116800:
            raise ValueError("line width must be between 0 and 20116800 EMU (1584 pt)")
        with self._change() as line:
            set_int(line, "w", None if emu is None else int(emu))

    @color.setter
    def color(self, value: "str | Color | None") -> None:
        with self._change() as line:
            replace_choice(line, LINE_FILL_TAGS, None if value is None else solid_fill(value))

    @visible.setter
    def visible(self, value: bool | None) -> None:
        """``False`` writes ``a:noFill``; ``None`` removes the explicit outline fill."""
        with self._change() as line:
            current = read_fill(line, LINE_FILL_TAGS)
            if value is None:
                replace_choice(line, LINE_FILL_TAGS, None)
            elif not value:
                replace_choice(line, LINE_FILL_TAGS, no_fill())
            elif current is not None and current.kind == "none":
                # Turning a hidden line back on needs a colour; tx1 is what PowerPoint uses.
                replace_choice(line, LINE_FILL_TAGS, solid_fill("tx1"))

    @dash.setter
    def dash(self, value: str | None) -> None:
        if value is not None and value not in DASH_STYLES:
            raise ValueError(f"dash must be one of {sorted(DASH_STYLES)}")
        with self._change() as line:
            replace_choice(line, ("a:prstDash", "a:custDash"),
                           None if value is None else make("a:prstDash", val=value))

    @cap.setter
    def cap(self, value: str | None) -> None:
        if value is not None and value not in LINE_CAPS:
            raise ValueError(f"cap must be one of {sorted(LINE_CAPS)}")
        with self._change() as line:
            set_attr(line, "cap", value)

    @head.setter
    def head(self, value: "Arrowhead | str | None") -> None:
        self._set_arrowhead("a:headEnd", value)

    @tail.setter
    def tail(self, value: "Arrowhead | str | None") -> None:
        self._set_arrowhead("a:tailEnd", value)

    def set_arrowhead(self, end: str, type: str | None, width: str | None = None,
                      length: str | None = None) -> "LineFormat":
        """``end`` is ``"head"`` or ``"start"`` (where the line begins) or ``"tail"`` or
        ``"end"`` (where it ends -- an arrow's point); returns the outline.

        For example::

            arrow.line.set_arrowhead("end", "triangle", "lg", "lg")
        """
        names = {"head": "head", "start": "head", "tail": "tail", "end": "tail"}
        if end not in names:
            raise ValueError("end must be 'start' ('head') or 'end' ('tail')")
        value = None if type is None else Arrowhead(type, width, length)
        self._set_arrowhead(f"a:{names[end]}End", value)
        return self

    def clear(self) -> "LineFormat":
        """Remove the explicit outline, so it inherits again; returns the outline.

        For example::

            shape.line.clear()
        """
        line = self._locate(False)
        if line is None:
            return self
        self._before_change()
        remove(line)
        self._after_change()
        return self

    # -- internals -------------------------------------------------------------------------

    def _arrowhead(self, tag: str) -> Arrowhead | None:
        line = self._locate(False)
        node = None if line is None else line.find(qn(tag))
        if node is None:
            return None
        return Arrowhead(node.get("type", "none"), node.get("w"), node.get("len"))

    def _set_arrowhead(self, tag: str, value: "Arrowhead | str | None") -> None:
        if isinstance(value, str):
            value = Arrowhead(value)
        if value is not None:
            if value.type not in ARROWHEADS:
                raise ValueError(f"arrowhead must be one of {sorted(ARROWHEADS)}")
            for size in (value.width, value.length):
                if size is not None and size not in ARROWHEAD_SIZES:
                    raise ValueError(f"arrowhead sizes must be one of {sorted(ARROWHEAD_SIZES)}")
        with self._change() as line:
            replace_choice(line, (tag,), None)
            if value is not None:
                node = subelement(line, tag)
                node.set("type", value.type)
                set_attr(node, "w", value.width)
                set_attr(node, "len", value.length)

    def _change(self):
        return _LineChange(self)

    def __repr__(self) -> str:
        return f"<LineFormat width={self.width} color={self.color} dash={self.dash}>"


class _LineChange:
    def __init__(self, line_format: LineFormat) -> None:
        self._format = line_format

    def __enter__(self) -> Element:
        self._format._before_change()
        line = self._format._locate(True)
        assert line is not None
        return line

    def __exit__(self, *exc) -> None:
        if exc[0] is None:
            self._format._after_change()
