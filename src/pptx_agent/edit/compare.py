"""A run of like slides compared, shape by shape: measured, never judged (round 7, L3).

An agent that built a run of like slides -- or had several agents build them -- knows which
slides belong together; :func:`compare_slides` measures where they differ.  It matches the
slides' corresponding shapes, takes for each property the value most of them have (the
majority) and lists the shapes that differ from it: a *near-miss* when the difference is
within the tolerance (``within``; a colour within :data:`NEAR_CHANNEL` per channel), which
is seldom intended, and apart from it the larger differences, which may well be.  Whether
any of it is intended is the caller's to decide.

**Matching**, deterministic, top-level shapes only (a group is compared as one shape):

1. *Placeholders* by their layout placeholder: type and idx (a centred title as a title,
   an untyped one as ``obj``).
2. *Names*: the shapes left whose name is unique on its slide and found on at least two of
   the slides, of the same kind.  Name groups that are each on some of the slides only --
   a slide rebuilt with other names -- are then joined by kind and position as in 3, the
   closest pair first, when no slide has a shape in both (round 8, M4).
3. *Kind and position*: the rest, slide by slide in the order given, each joining the
   nearest group of its kind (the box of that group's first shape) that has no shape on the
   slide yet -- a name group too -- nearest by the sum of the differences of x, y, w and h,
   the closest pair first, when the boxes overlap or are at most :data:`POSITION_REACH`
   apart; else it starts a group of its own.  The kind is as ``ppt_read_slides`` "compact"
   gives it (a graphic frame as ``table``, ``chart`` or ``smartart``).

A name group whose shapes have different names -- one joined it by position -- reports
the name as its first difference: ``name``, the majority's, and the others' under
``other``.  A position group's names (PowerPoint's ``Rectangle 3``, ``Rectangle 7``) are not
compared.

A group on fewer slides than all is reported: ``missing_on`` the slides without it when it
is on more than half of them, ``only`` its shapes when it is on half or fewer.

**Properties** (:data:`PROPERTIES`): the box (``x``, ``y``, ``w``, ``h`` in points, two
decimals), the text's size per paragraph level (the size most of the level's characters
are drawn at, :func:`.effective.size`, autofit scale included), its colour (resolved to
hex, the colour most characters are drawn in) and typeface, the shape's fill (resolved
hex, ``none``, the kind of a gradient or picture fill, ``default`` when the slide states
none) and its outline's colour and width.  Bold, italic and the like are not compared.
A shape without text, a fill or an outline has no value for those properties and is left
out of their comparison.

**Text layout** (round 14, S1), of a shape with text, as drawn -- the paragraph's own
``a:pPr``, its shape's list style, the layout and master placeholders', the master's text
style and the presentation's default (:func:`.effective.paragraph_layout`), the frame's
``a:bodyPr`` from the shape outwards (:class:`~pptx_agent.TextFrame`): per paragraph level,
like the size, ``align`` (``ppt_format_text``'s names), ``line_spacing`` (a multiple of
single, ``"18pt"`` when exact), ``space_before`` and ``space_after`` (points, ``"20%"`` when
a share of a line) and ``bullet`` (``none``, ``char •``, ``autonum arabicPeriod``,
``picture``) -- each the level's first paragraph with text's, ``"mixed"`` when another of
the level differs; and the frame's ``anchor``, ``insets`` (left, top, right, bottom, points,
within the tolerance), ``autofit`` (``none``, ``normal``, ``shape``) and ``direction``
(``a:bodyPr@vert``, ``horz`` when nothing says).  Paragraphs without text are left out.
Every difference -- not only a near-miss: a heading centred on some slides of a run is
seldom a design -- gets ``ppt_format_text`` items giving the majority's value (a level's
paragraphs, or the shape when all its paragraphs are of the level); none where that tool
cannot write the value: the direction, an exact line spacing, a share of a line, a bullet
other than ``•`` or numbering other than ``arabicPeriod``, ``distributed`` alignment or
anchor, and a ``mixed`` majority.

**Tables** (round 11, P3) are compared inside, as PowerPoint draws them (pptx2svg's table
layout, :attr:`Table.drawn_row_heights <pptx_agent.edit.table.Table.drawn_row_heights>`):
with the Header Row flag on, the header row's height (``header_h``), fill (``header_fill``,
the fill most of its cells are drawn with, the table style's included) and text size
(``header_size``); the body rows' height (``row_h``: the height most of them are drawn at,
the least of those tied), each value listing the body row ``rows`` counts of its tables as
context -- a table of three rows and one of six agree when their rows do; the column
``widths`` and the cells' ``insets`` (left, top, right, bottom, the most common), each
compared element by element within the tolerance; and the ``borders`` drawn (the most
common line on each side, ``"1 #FFFFFF"`` when all four agree).  A table's own box height
is not compared: it is its rows'.  Every differing table -- not only a near-miss, since a
stretched header row is seldom a design -- gets a fix: ``ppt_edit_table`` ``calls`` (each
one call's arguments but ``doc``: ``set_heights`` with ``rows``, ``set_widths``,
``format_cells``) or ``ppt_format_text`` items (``/row0``'s size, the table's insets);
widths across tables of other column counts, and borders, get none.

Where a value comes from (``from``): ``slide`` when the slide states it, else ``layout``,
``master`` or ``default`` -- for the size and colour of placeholder text and a
placeholder's box.  A near-miss has a ``fix`` where one call gives the shapes the
majority's value: ``ppt_set_shape`` for a box, fill or outline, ``ppt_format_text`` for
text -- ``size: 0`` where the majority's size is the layout's and the shape's layout gives
the same (:meth:`TextFrame.clear_size <pptx_agent.TextFrame.clear_size>`).

For example::

    comparison = deck.compare_slides(["s:257", "s:258", "s:259", "s:260"])
    comparison.to_json()["shapes"][0]
    # {'shape': '257.2', 'name': 'Title 1', 'differ': [{'property': 'y', 'value': 40,
    #  'on': 3, 'near': [{'value': 41, 'shapes': ['259.2']}],
    #  'fix': {'tool': 'ppt_set_shape', 'items': [{'target': '259.2', 'y': 40}]}}]}
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from ..oxml.xml import local_name
from . import effective
from .errors import ArgumentError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

EMU_PER_POINT = 12700
#: The default near-miss tolerance, points: boxes, text sizes and outline widths.
WITHIN = 2.0
#: Two colours are a near-miss when no channel differs by more than this (of 255).
NEAR_CHANNEL = 8
#: How far apart (the sum of the x, y, w and h differences, points) two boxes that do not
#: overlap may still be matched by position.
POSITION_REACH = 36.0
#: What is compared, in the order reported.
PROPERTIES = ("x", "y", "w", "h", "size", "color", "font", "align", "line_spacing",
              "space_before", "space_after", "bullet", "anchor", "autofit", "direction",
              "fill", "line", "line_width", "header_h", "header_fill", "header_size", "row_h",
              "widths", "insets", "borders")
_NUMERIC = frozenset(("x", "y", "w", "h", "size", "line_width", "header_h", "header_size",
                      "row_h", "space_before", "space_after"))
#: A text's layout, per paragraph level (round 14, S1): each level's first paragraph's
#: effective value, ``"mixed"`` when that level's paragraphs differ.
_PARAGRAPH = ("align", "line_spacing", "space_before", "space_after", "bullet")
#: A text frame's layout, as it applies (its own, else its layout's and master's).
_FRAME = ("anchor", "autofit", "direction")
#: Compared element by element, each within the tolerance: a table's column widths and its
#: cells' insets (left, top, right, bottom), points.
_SERIES = frozenset(("widths", "insets"))
_COLOURS = frozenset(("color", "fill", "line", "header_fill"))
_SHEETS = {"sld": "slide", "sldLayout": "layout", "sldMaster": "master"}


def _points(emu: float) -> float:
    return round(float(emu) / EMU_PER_POINT, 2)


def _exact(emu: float) -> float:
    """EMU as points to four decimals: a fix's length, which gives back the EMU it came
    from (the comparison's two decimals would not)."""
    return round(float(emu) / EMU_PER_POINT, 4)


def _number(value):
    if isinstance(value, tuple):
        return [_number(item) for item in value]
    return int(value) if isinstance(value, float) and value.is_integer() else value


def kind_of(shape: "Shape") -> str:
    """A shape's kind as ``ppt_read_slides`` "compact" gives it."""
    kind = shape.kind
    if kind == "graphic_frame":
        if shape.has_table:
            return "table"
        if shape.has_chart:
            return "chart"
        if shape.has_diagram:
            return "smartart"
    return kind


def _sheet(element) -> str:
    """Which part an element sits in: ``slide``, ``layout``, ``master`` or ``default``."""
    if element is None:
        return "default"
    return _SHEETS.get(local_name(element.getroottree().getroot()), "default")


# -- measuring ----------------------------------------------------------------------------


@dataclass
class _Value:
    value: Any
    source: str | None = None        #: slide, layout, master, default (where it matters)
    fix: Any = None                  #: what a fix item sets to give a shape this value
    rows: int | None = None          #: a table's body rows, the context of its row height


def _hex(color, shape) -> str | None:
    try:
        resolved = color.resolve(shape)
    except Exception:  # noqa: BLE001 -- a colour that does not resolve is not compared
        return None
    return f"#{resolved.lstrip('#').upper()}" if resolved else None


def _color_text(color) -> str:
    from ..tools.common import color_text

    return color_text(color)


def _text_values(shape: "Shape", out: dict) -> None:
    frame = shape.text_frame
    if frame is None:
        return
    host = frame._host()
    placeholder = shape.placeholder is not None
    sizes: dict[int, Counter] = {}
    size_source: dict[tuple[int, float], str] = {}
    colours: Counter = Counter()
    colour_fix: dict[str, Any] = {}
    colour_source: dict[str, str] = {}
    faces: Counter = Counter()
    for paragraph in frame.paragraphs:
        element = paragraph._element()
        level = paragraph.level
        for run in paragraph.runs:
            text = run.text
            weight = len(text.strip())
            if not weight:
                continue
            node = run._element()
            points = effective.size(host, element, node)
            sizes.setdefault(level, Counter())[points] += weight
            if placeholder and (level, points) not in size_source:
                size_source[(level, points)] = _size_sheet(host, element, node)
            colour = effective.color(host, element, node)
            if colour is not None:
                key = f"#{colour.value.upper()}"
                colours[key] += weight
                if key not in colour_fix:
                    stated, _ = effective.color_source(host, element, node)
                    colour_fix[key] = _color_text(stated) if stated is not None else key
                    if placeholder:
                        colour_source[key] = _colour_sheet(host, element, node)
            face = effective.typeface(host, element, node)
            if face:
                faces[face] += weight
    for level, counter in sorted(sizes.items()):
        points = _dominant(counter)
        out[("size", level)] = _Value(points, size_source.get((level, points)))
    if sizes:                            # text to lay out: an empty shape has no layout
        _layout_values(frame, host, out)
    if colours:
        key = _dominant(colours)
        out[("color", 0)] = _Value(key, colour_source.get(key), colour_fix[key])
    if faces:
        out[("font", 0)] = _Value(_dominant(faces))


def _layout_values(frame, host, out: dict) -> None:
    """The text's layout as drawn (:func:`.effective.paragraph_layout`, the frame's inherited
    attributes): per level, the paragraph properties of :data:`_PARAGRAPH` -- the first
    paragraph with text's, ``"mixed"`` when another of the level differs -- and the frame's
    anchor, insets, autofit and text direction.  Each value's ``fix`` is the
    ``ppt_format_text`` item keys that give it, ``None`` where no tool sets it."""
    levels: dict[int, dict[str, list]] = {}
    for paragraph in frame.paragraphs:
        if not paragraph.text.strip():
            continue
        drawn = effective.paragraph_layout(host, paragraph._element())
        values = levels.setdefault(drawn["level"], {})
        for name, value in _paragraph_layout(drawn).items():
            values.setdefault(name, []).append(value)
    for level, values in sorted(levels.items()):
        for name in _PARAGRAPH:
            seen = values[name]
            first = seen[0]
            if any(value != first for value in seen[1:]):
                out[(name, level)] = _Value("mixed")
            else:
                out[(name, level)] = _Value(first[0], fix=first[1])
    anchor = frame.anchor
    out[("anchor", 0)] = _Value(anchor, fix={"anchor": anchor} if anchor in _ANCHORS else None)
    out[("insets", 0)] = _Value(tuple(_points(side) for side in frame.insets))
    autofit = frame.autofit
    out[("autofit", 0)] = _Value(autofit, fix={"autofit": autofit})
    # a:bodyPr@vert (ECMA-376 Part 1, 21.1.2.1.1; "horz" when nothing says): no tool sets it.
    out[("direction", 0)] = _Value(frame._frame_attribute("vert") or "horz")


#: The values ``ppt_format_text`` writes: ``align`` and ``anchor``.
_ALIGNS = ("left", "center", "right", "justify")
_ANCHORS = ("top", "middle", "bottom")


def _paragraph_layout(drawn: dict) -> dict[str, tuple]:
    """A paragraph's :data:`_PARAGRAPH` values, each ``(value, fix)``: line spacing a
    multiple of single (``"18pt"`` when exact), space before and after points (``"20%"``
    of a line when the style gives a share), the bullet ``none``, ``char •``, ``autonum
    arabicPeriod`` or ``picture``."""
    align = drawn["alignment"]
    out = {"align": (align, {"align": align} if align in _ALIGNS else None)}
    exact, multiple = drawn["line_spacing_points"], drawn["line_spacing"]
    if exact is not None:
        out["line_spacing"] = (f"{_number(exact)}pt", None)
    else:
        multiple = 1 if multiple is None else _number(round(multiple, 2))
        out["line_spacing"] = (multiple, {"line_spacing": multiple}
                               if 0.1 <= multiple <= 10 else None)   # the tool's range
    for name in ("space_before", "space_after"):
        lines = drawn[f"{name}_lines"]
        if lines is not None:
            out[name] = (f"{_number(round(lines * 100, 2))}%", None)
        else:
            points = _number(_points(drawn[name] or 0))
            out[name] = (points, {name: points} if points <= 1000 else None)
    bullet = drawn["bullet"]
    if bullet is None:
        out["bullet"] = ("none", {"bullet": "none"})
    elif bullet.kind == "char":
        out["bullet"] = (f"char {bullet.char}",
                         {"bullet": "bullet"} if bullet.char == "\u2022" else None)
    elif bullet.kind == "number":
        out["bullet"] = (f"autonum {bullet.scheme}",
                         {"bullet": "number"} if bullet.scheme == "arabicPeriod" else None)
    else:
        out["bullet"] = (bullet.kind, None)
    return out


def _dominant(counter: Counter):
    """The value with the most characters; the first met of those tied."""
    best = max(counter.values())
    return next(value for value, count in counter.items() if count == best)


def _size_sheet(host, paragraph, run) -> str:
    if effective.slide_size(host, paragraph, run) is not None:
        return "slide"
    for properties in effective._run_properties(host, paragraph, run):
        if properties.get("sz"):
            return _sheet(properties)
    return "default"


def _colour_sheet(host, paragraph, run) -> str:
    for holder, _label in effective._color_layers(host, paragraph, run):
        if local_name(holder) in effective._RUN_PROPERTIES:
            if any(child.tag in effective._FILLS for child in holder):
                return _sheet(holder)
        elif holder is not None:
            from .color import Color

            if Color.from_element(holder) is not None:
                return _sheet(holder)
    return "default"


def measure(shape: "Shape", cache: dict | None = None) -> dict[tuple[str, int], _Value]:
    """A shape's values, ``{(property, level): value}`` -- level 0 but for the text size.
    ``cache`` keeps each slide's resolved table layout between calls (a table's)."""
    out: dict[tuple[str, int], _Value] = {}
    bounds = shape.bounds
    table = kind_of(shape) == "table"
    if bounds is not None:
        source = None
        if shape.placeholder is not None:
            source = "slide" if shape.has_explicit_transform else "layout"
        for name, emu in zip(("x", "y", "w", "h"), bounds):
            if not (table and name == "h"):      # a table's height is its rows'
                out[(name, 0)] = _Value(_points(emu), source)
    if table:
        _table_values(shape, out, {} if cache is None else cache)
    if shape.kind == "shape":
        _text_values(shape, out)
        fill = shape.effective_fill
        if fill is None:
            out[("fill", 0)] = _Value("default", fix="inherit")
        elif fill.kind == "solid" and fill.color is not None:
            value = _hex(fill.color, shape)
            if value is not None:
                out[("fill", 0)] = _Value(value, fix=_color_text(fill.color))
        else:
            out[("fill", 0)] = _Value(fill.kind, fix="none" if fill.kind == "none" else None)
    if shape.kind in ("shape", "connector", "picture"):
        line = shape.line
        if line is not None and line.exists:
            if line.visible is False:
                out[("line", 0)] = _Value("none", fix="none")
            elif line.color is not None:
                value = _hex(line.color, shape)
                if value is not None:
                    out[("line", 0)] = _Value(value, fix=_color_text(line.color))
            if line.width is not None and line.visible is not False:
                out[("line_width", 0)] = _Value(_points(line.width))
    return out


def _table_values(shape: "Shape", out: dict, cache: dict) -> None:
    """A table's values: its header row's drawn height, fill and text size (when the Header
    Row flag is on), its body rows' height -- the height most of them are drawn at, the
    least of those tied -- with their count as context, its column widths, its cells' insets
    and borders (the most common of each)."""
    table = shape.table
    try:
        from .fit import _Measured, table_heights

        slide = shape._slide
        measured = cache.get(slide.slide_id)
        if measured is None:
            measured = cache[slide.slide_id] = _Measured(slide)
        heights = table_heights(shape, measured)
        resolved = getattr(measured.element(shape)[0], "table", None)
    except ImportError:                       # pragma: no cover - the extras-less install
        heights, resolved = table.row_heights, None
    header = 1 if table.flags.get("first_row") and len(heights) > 1 else 0
    if header:
        out[("header_h", 0)] = _Value(_points(heights[0]), fix=_exact(heights[0]))
        fill = _header_fill(resolved)
        if fill is not None:
            out[("header_fill", 0)] = _Value(fill, fix=fill)
        sizes: Counter = Counter()
        for column in range(table.columns):
            cell = table.cell(0, column)
            if cell.is_spanned:
                continue
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    weight = len(run.text.strip())
                    if weight:
                        sizes[effective.size(cell, paragraph._element(), run._element())] \
                            += weight
        if sizes:
            out[("header_size", 0)] = _Value(_dominant(sizes))
    body = [_points(height) for height in heights[header:]]
    if body:
        counts = Counter(body)
        most = max(counts.values())
        typical = min(v for v, n in counts.items() if n == most)
        exact = next(height for height in heights[header:] if _points(height) == typical)
        out[("row_h", 0)] = _Value(typical, fix=_exact(exact), rows=len(body))
    out[("widths", 0)] = _Value(tuple(_points(width) for width in table.column_widths),
                                fix=[_number(_exact(width)) for width in table.column_widths])
    insets = Counter(tuple(_points(side) for side in cell.text_frame.insets)
                     for cell in table.iter_cells() if not cell.is_spanned)
    if insets:
        out[("insets", 0)] = _Value(_dominant(insets))
    borders = _borders(resolved)
    if borders is not None:
        out[("borders", 0)] = _Value(borders)


def _paint(fill) -> str | None:
    """A resolved fill as compared: ``#RRGGBB``, ``none``, or its kind."""
    if fill is None:
        return "none"
    kind = getattr(fill, "type", None)
    if kind == "solid":
        return f"#{fill.color.hex.lstrip('#').upper()}"
    return kind


def _header_fill(resolved) -> str | None:
    """The fill most of the header row's cells are drawn with (pptx2svg's resolved table:
    the cell's own, else the table style's)."""
    if resolved is None or not resolved.rows:
        return None
    fills = Counter(_paint(cell.fill) for cell in resolved.rows[0].cells
                    if not (cell.h_merge or cell.v_merge))
    return _dominant(fills) if fills else None


def _borders(resolved) -> str | None:
    """The cells' borders as drawn, the most common line on each side: ``"1 #FFFFFF"``
    (points and colour) when every side has the same, else ``"top 1 #FFFFFF; bottom
    none; ..."``."""
    if resolved is None:
        return None
    sides = ("top", "bottom", "left", "right")
    lines = {side: Counter() for side in sides}
    for row in resolved.rows:
        for cell in row.cells:
            if cell.h_merge or cell.v_merge:
                continue
            for side in sides:
                line = getattr(cell.borders, side, None) if cell.borders else None
                paint = None if line is None else _paint(line.fill)
                lines[side]["none" if paint in (None, "none") or not line.width
                            else f"{_number(_points(line.width))} {paint}"] += 1
    common = {side: _dominant(counter) for side, counter in lines.items() if counter}
    if not common:
        return None
    if len(set(common.values())) == 1:
        return next(iter(common.values()))
    return "; ".join(f"{side} {value}" for side, value in common.items())


# -- matching -----------------------------------------------------------------------------


@dataclass
class ShapeGroup:
    """Corresponding shapes on the compared slides: ``shapes`` by slide id, matched ``by``
    ``placeholder``, ``name`` or ``position``."""

    by: str
    kind: str
    shapes: dict[int, "Shape"] = field(default_factory=dict)

    @property
    def first(self) -> "Shape":
        return next(iter(self.shapes.values()))


def _placeholder_key(shape: "Shape"):
    ph = shape.placeholder
    kind = ph.type or "obj"
    return ("title" if kind == "ctrTitle" else kind, ph.idx)


def _distance(a, b) -> float:
    return sum(abs(float(p) - float(q)) for p, q in zip(a, b)) / EMU_PER_POINT


def _overlap(a, b) -> bool:
    return (a[0] < b[0] + b[2] and b[0] < a[0] + a[2]
            and a[1] < b[1] + b[3] and b[1] < a[1] + a[3])


def match(slides: list["Slide"]) -> list[ShapeGroup]:
    """The slides' shapes in groups of corresponding shapes (the module's matching rules),
    ordered by the first slide each is on, then back to front there."""
    groups: list[ShapeGroup] = []
    order: dict[int, tuple[int, int]] = {}             # id(group) -> (slide, z) first seen
    left: dict[int, list[tuple[int, "Shape"]]] = {}
    for slide in slides:
        left[slide.slide_id] = [(z, shape) for z, shape in enumerate(slide.shapes)
                                if shape.parent_group is None]

    def add(group: ShapeGroup, slide_index: int, z: int) -> None:
        groups.append(group)
        order[id(group)] = (slide_index, z)

    # 1. placeholders by type and idx
    by_key: dict[Any, ShapeGroup] = {}
    for index, slide in enumerate(slides):
        taken = set()
        rest = []
        for z, shape in left[slide.slide_id]:
            if shape.placeholder is None:
                rest.append((z, shape))
                continue
            key = _placeholder_key(shape)
            if key in taken:
                rest.append((z, shape))
                continue
            taken.add(key)
            group = by_key.get(key)
            if group is None:
                group = by_key[key] = ShapeGroup("placeholder", kind_of(shape))
                add(group, index, z)
            group.shapes[slide.slide_id] = shape
        left[slide.slide_id] = rest
    # 2. names unique on their slide, on two slides or more, of one kind
    named: dict[tuple[str, str], list[int]] = {}
    for slide in slides:
        names = Counter(shape.name for _, shape in left[slide.slide_id] if shape.name)
        for _, shape in left[slide.slide_id]:
            if shape.name and names[shape.name] == 1:
                named.setdefault((shape.name, kind_of(shape)), []).append(slide.slide_id)
    by_name: dict[tuple[str, str], ShapeGroup] = {}
    for index, slide in enumerate(slides):
        rest = []
        for z, shape in left[slide.slide_id]:
            key = (shape.name, kind_of(shape))
            if not shape.name or len(named.get(key, ())) < 2 \
                    or slide.slide_id not in named[key]:
                rest.append((z, shape))
                continue
            group = by_name.get(key)
            if group is None:
                group = by_name[key] = ShapeGroup("name", key[1])
                add(group, index, z)
            group.shapes[slide.slide_id] = shape
        left[slide.slide_id] = rest
    # 2b. name groups on some slides each, joined by kind and position (renamed shapes)
    total = len(slides)
    partial = [group for group in by_name.values() if len(group.shapes) < total]
    joins = []
    for a_number, a in enumerate(partial):
        for b_number in range(a_number + 1, len(partial)):
            b = partial[b_number]
            box, other = a.first.bounds, b.first.bounds
            if a.kind != b.kind or box is None or other is None:
                continue
            distance = _distance(box, other)
            if _overlap(box, other) or distance <= POSITION_REACH:
                joins.append((distance, a_number, b_number))
    joins.sort()
    gone: set[int] = set()
    for distance, a_number, b_number in joins:
        a, b = partial[a_number], partial[b_number]
        if a_number in gone or b_number in gone or set(a.shapes) & set(b.shapes):
            continue
        # The group first met keeps its place (its shapes go back in slide order below).
        keep, drop = (a, b) if order[id(a)] <= order[id(b)] else (b, a)
        keep.shapes.update(drop.shapes)
        gone.add(b_number if drop is b else a_number)
        groups.remove(drop)
        del order[id(drop)]
    # 3. kind and nearest box; a name group without a shape on the slide takes one too
    named_groups = [group for group in groups if group.by == "name"]
    placed: list[ShapeGroup] = []
    for index, slide in enumerate(slides):
        pairs = []
        pool = [*named_groups, *placed]
        for z, shape in left[slide.slide_id]:
            box = shape.bounds
            if box is None:
                continue
            for number, group in enumerate(pool):
                other = group.first.bounds
                if group.kind != kind_of(shape) or other is None \
                        or slide.slide_id in group.shapes:
                    continue
                distance = _distance(box, other)
                if _overlap(box, other) or distance <= POSITION_REACH:
                    pairs.append((distance, z, number, shape))
        pairs.sort(key=lambda item: item[:3])
        used_shapes, used_groups = set(), set()
        for distance, z, number, shape in pairs:
            if z in used_shapes or number in used_groups:
                continue
            used_shapes.add(z)
            used_groups.add(number)
            pool[number].shapes[slide.slide_id] = shape
        for z, shape in left[slide.slide_id]:
            if z not in used_shapes:
                group = ShapeGroup("position", kind_of(shape), {slide.slide_id: shape})
                placed.append(group)
                add(group, index, z)
    position = {slide.slide_id: index for index, slide in enumerate(slides)}
    for group in named_groups:          # a shape joined by position: back in slide order
        group.shapes = dict(sorted(group.shapes.items(), key=lambda item: position[item[0]]))
    groups.sort(key=lambda group: order[id(group)])
    return groups


# -- comparing ----------------------------------------------------------------------------


def _near(name: str, a, b, within: float) -> bool:
    if name in _SERIES:
        return len(a) == len(b) and all(abs(p - q) <= within + 1e-9 for p, q in zip(a, b))
    if name in _NUMERIC:                # a share of a line ("20%") or "mixed" is no number
        return isinstance(a, (int, float)) and isinstance(b, (int, float)) \
            and abs(a - b) <= within + 1e-9
    if name in _COLOURS and isinstance(a, str) and isinstance(b, str) \
            and len(a) == len(b) == 7 and a.startswith("#") and b.startswith("#"):
        return all(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) <= NEAR_CHANNEL
                   for i in (1, 3, 5))
    return False


def _fix_item(name: str, level: int, shape: "Shape", value: _Value, majority: _Value):
    """The fix items giving ``shape`` the majority's value, or ``None``."""
    target = shape.id
    if name in _TABLE_FIX:
        return _TABLE_FIX[name](shape, majority)
    if name in ("x", "y", "w", "h"):
        return [{"target": target, name: _number(majority.value)}]
    if name == "fill":
        return None if majority.fix is None else [{"target": target, "fill": majority.fix}]
    if name == "line":
        return None if majority.fix is None else \
            [{"target": target, "line": {"color": majority.fix}}]
    if name == "line_width":
        return [{"target": target, "line": {"width": _number(majority.value)}}]
    if name == "color":
        return [{"target": target, "color": majority.fix or majority.value}]
    if name in _FRAME:
        return None if majority.fix is None else [{"target": target, **majority.fix}]
    if name == "size":
        frame = shape.text_frame
        host = frame._host()
        paragraphs = [p for p in frame.paragraphs if p.level == level]
        if majority.source not in (None, "slide") and value.source == "slide" and paragraphs \
                and effective.template_size(host, paragraphs[0]._element()) == majority.value \
                and {o.level for o in effective.size_overrides(target, shape)} <= {level}:
            return [{"target": target, "size": 0}]        # no other level's size goes too
        return _level_items(shape, level, {"size": _number(majority.value)})
    if name in _PARAGRAPH:
        return None if majority.fix is None else _level_items(shape, level, majority.fix)
    return None


def _level_items(shape: "Shape", level: int, keys: dict) -> list[dict]:
    """``ppt_format_text`` items giving a level's paragraphs ``keys``: the shape when all
    its paragraphs are of the level, else each of the level's paragraphs with text."""
    frame = shape.text_frame
    if all(p.level == level for p in frame.paragraphs):
        return [{"target": shape.id, **keys}]
    return [{"target": p.address, **keys} for p in frame.paragraphs
            if p.level == level and p.text.strip()]


def _row_fix(shape: "Shape", majority: _Value, rows: list[int]):
    value = _number(majority.fix if majority.fix is not None else majority.value)
    return [{"target": shape.id, "action": "set_heights", "rows": rows,
             "sizes": [value] * len(rows)}]


def _body_rows(shape: "Shape") -> list[int]:
    table = shape.table
    first = 1 if table.flags.get("first_row") and table.rows > 1 else 0
    return list(range(first, table.rows))


def _insets_fix(shape: "Shape", majority: _Value):
    return [{"target": shape.id, "insets": dict(zip(("left", "top", "right", "bottom"),
                                                    (_number(v) for v in majority.value)))}]


#: A table's fixes: the ``ppt_edit_table`` calls (``calls``, each one call's arguments but
#: ``doc``) or ``ppt_format_text`` items giving it the majority's value.
_TABLE_FIX = {
    "header_h": lambda shape, majority: _row_fix(shape, majority, [0]),
    "row_h": lambda shape, majority: _row_fix(shape, majority, _body_rows(shape)),
    "widths": lambda shape, majority: [{"target": shape.id, "action": "set_widths",
                                        "sizes": majority.fix}],
    "header_fill": lambda shape, majority: [{"target": shape.id, "action": "format_cells",
                                             "rows": [0], "fill": majority.fix}],
    "header_size": lambda shape, majority: [{"target": f"{shape.id}/row0",
                                             "size": _number(majority.value)}],
    "insets": _insets_fix,
}

_FIX_TOOL = {"size": "ppt_format_text", "color": "ppt_format_text",
             "header_size": "ppt_format_text", "insets": "ppt_format_text",
             **{name: "ppt_format_text" for name in (*_PARAGRAPH, *_FRAME)},
             "header_h": "ppt_edit_table", "row_h": "ppt_edit_table",
             "widths": "ppt_edit_table", "header_fill": "ppt_edit_table"}
#: Tools that take one target a call: their fixes are ``calls``, not ``items``.
_ONE_TARGET = frozenset(("ppt_edit_table",))
#: Every difference fixed, not only a near-miss: a table's, and the text's layout -- a
#: heading centred on some slides of a run is seldom a design (round 14, S1).
_ALWAYS_FIX = frozenset((*_TABLE_FIX, *_PARAGRAPH, *_FRAME))


@dataclass
class Comparison:
    """What :func:`compare_slides` found: the ``slides`` compared (ids), the ``groups`` of
    corresponding shapes, and the ``within`` tolerance (points)."""

    slides: list[int]
    groups: list[ShapeGroup]
    within: float = WITHIN
    _layouts: dict = field(default_factory=dict, repr=False, compare=False)

    def to_json(self) -> dict[str, Any]:
        """``{"slides", "matched", "agree", "shapes": [...], "only": [...]}``: ``matched``
        how many groups of corresponding shapes, ``agree`` how many of them show no
        difference; ``shapes`` one entry per group that differs or is missing on some
        slides -- ``{"shape": first address, "name", "by": "position" (when matched so),
        "differ": [...], "missing_on": [...]}``; ``only`` the groups on half the slides or
        fewer, ``{"shape", "name", "shapes": [addresses]}``.  Each ``differ`` entry is
        ``{"property", "level" (a size's or paragraph property's, past level 0), "value":
        the majority's, "on": how many have it, "from", "near": [{"value", "from",
        "shapes"}], "other": [...], "fix": {"tool", "items"}}`` -- ``fix`` for the
        near-misses, for every text-layout difference and for every differing table
        (``{"tool": "ppt_edit_table", "calls": [...]}``, a call per entry); ``from`` given
        where the sources differ; ``rows`` a table's body row counts beside its ``row_h``
        (the majority's, and each value's).  A name group whose names differ has
        ``{"property": "name", "value", "on", "other"}`` first."""
        entries, only = self.entries()
        return {"slides": [f"s:{sid}" for sid in self.slides], "matched": len(self.groups),
                "agree": len(self.groups) - len(entries) - len(only),
                "shapes": entries, "only": only}

    def entries(self) -> tuple[list[dict], list[dict]]:
        total = len(self.slides)
        entries: list[dict] = []
        only: list[dict] = []
        for group in self.groups:
            head = {"shape": group.first.id}
            if group.first.name:
                head["name"] = group.first.name
            if group.by == "position":
                head["by"] = "position"
            count = len(group.shapes)
            if count * 2 <= total and count < total:
                only.append({**head, "shapes": [s.id for s in group.shapes.values()]})
                continue
            differ = self._differences(group)
            entry = dict(head)
            if differ:
                entry["differ"] = differ
            if count < total:
                entry["missing_on"] = [f"s:{sid}" for sid in self.slides
                                       if sid not in group.shapes]
            if differ or count < total:
                entries.append(entry)
        return entries, only

    def _differences(self, group: ShapeGroup) -> list[dict]:
        measured = {sid: measure(shape, self._layouts) for sid, shape in group.shapes.items()}
        keys: list[tuple[str, int]] = []
        for values in measured.values():
            for key in values:
                if key not in keys:
                    keys.append(key)
        keys.sort(key=lambda key: (PROPERTIES.index(key[0]), key[1]))
        out = []
        if group.by == "name":
            out += _name_difference(group)
        for name, level in keys:
            holders = [(sid, values[(name, level)]) for sid, values in measured.items()
                       if (name, level) in values]
            if len(holders) < 2:
                continue
            counts = Counter(value.value for _, value in holders)
            if len(counts) == 1:
                continue
            best = max(counts.values())
            top = next(value.value for _, value in holders if counts[value.value] == best)
            majority_values = [value for _, value in holders if value.value == top]
            sources = Counter(value.source for value in majority_values)
            majority_source = sources.most_common(1)[0][0]
            majority = next(value for value in majority_values if value.source == majority_source)
            entry: dict[str, Any] = {"property": name}
            if level:                        # a text size's or paragraph's level
                entry["level"] = level
            entry["value"] = _number(top)
            entry["on"] = best
            if name == "row_h":
                entry["rows"] = sorted({value.rows for value in majority_values})
            near: dict[tuple, list] = {}
            other: dict[tuple, list] = {}
            counts_of: dict[tuple, set] = {}
            fixes: list = []
            for sid, value in holders:
                if value.value == top:
                    continue
                shape = group.shapes[sid]
                bucket = near if _near(name, value.value, top, self.within) else other
                bucket.setdefault((value.value, value.source), []).append(shape.id)
                if name == "row_h":
                    counts_of.setdefault((value.value, value.source), set()).add(value.rows)
                # A table's rows, widths, insets and header: every one that differs is fixed
                # (a stretched header is no design), but not across column counts.
                always = name in _ALWAYS_FIX and not (
                    name in _SERIES and len(value.value) != len(majority.value))
                if bucket is near or always:
                    items = _fix_item(name, level, shape, value, majority)
                    if items is None:
                        fixes = None
                    elif fixes is not None:
                        fixes += items
            shown_from = any(source != majority_source for _, source in [*near, *other])
            if shown_from and majority_source is not None:
                entry["from"] = majority_source
            for label, bucket in (("near", near), ("other", other)):
                if bucket:
                    entry[label] = [
                        {"value": _number(value),
                         **({"from": source} if shown_from and source is not None else {}),
                         **({"rows": sorted(counts_of[(value, source)])}
                            if (value, source) in counts_of else {}),
                         "shapes": shapes}
                        for (value, source), shapes in bucket.items()]
            if fixes and (near or name in _ALWAYS_FIX):
                tool = _FIX_TOOL.get(name, "ppt_set_shape")
                entry["fix"] = {"tool": tool,
                                "calls" if tool in _ONE_TARGET else "items": fixes}
            out.append(entry)
        return out


def _name_difference(group: ShapeGroup) -> list[dict]:
    """A name group's names when they differ: ``{"property": "name", "value": the
    majority's, "on", "other": [{"value", "shapes"}]}`` -- no fix, a name being no
    near-miss."""
    names = [(shape.name, shape.id) for shape in group.shapes.values()]
    counts = Counter(name for name, _ in names)
    if len(counts) < 2:
        return []
    top = _dominant(counts)
    other: dict[str, list[str]] = {}
    for name, address in names:
        if name != top:
            other.setdefault(name, []).append(address)
    return [{"property": "name", "value": top, "on": counts[top],
             "other": [{"value": name, "shapes": shapes} for name, shapes in other.items()]}]


def compare_slides(slides: list["Slide"], *, within: float = WITHIN) -> Comparison:
    """``slides``' corresponding shapes and where they differ (the module's rules);
    ``within`` the near-miss tolerance in points.  Raises ``ValueError`` for fewer than two
    slides or a slide given twice."""
    if len(slides) < 2:
        raise ArgumentError("compare at least two slides")
    ids = [slide.slide_id for slide in slides]
    if len(set(ids)) != len(ids):
        repeated = next(sid for sid in ids if ids.count(sid) > 1)
        raise ArgumentError(f"s:{repeated} is given twice; name each slide once")
    return Comparison(ids, match(slides), float(within))


__all__ = ["Comparison", "NEAR_CHANNEL", "POSITION_REACH", "PROPERTIES", "ShapeGroup",
           "WITHIN", "compare_slides", "kind_of", "match", "measure"]
