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
PROPERTIES = ("x", "y", "w", "h", "size", "color", "font", "fill", "line", "line_width")
_NUMERIC = frozenset(("x", "y", "w", "h", "size", "line_width"))
_COLOURS = frozenset(("color", "fill", "line"))
_SHEETS = {"sld": "slide", "sldLayout": "layout", "sldMaster": "master"}


def _points(emu: float) -> float:
    return round(float(emu) / EMU_PER_POINT, 2)


def _number(value):
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
    if colours:
        key = _dominant(colours)
        out[("color", 0)] = _Value(key, colour_source.get(key), colour_fix[key])
    if faces:
        out[("font", 0)] = _Value(_dominant(faces))


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


def measure(shape: "Shape") -> dict[tuple[str, int], _Value]:
    """A shape's values, ``{(property, level): value}`` -- level 0 but for the text size."""
    out: dict[tuple[str, int], _Value] = {}
    bounds = shape.bounds
    if bounds is not None:
        source = None
        if shape.placeholder is not None:
            source = "slide" if shape.has_explicit_transform else "layout"
        for name, emu in zip(("x", "y", "w", "h"), bounds):
            out[(name, 0)] = _Value(_points(emu), source)
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
    if name in _NUMERIC:
        return abs(a - b) <= within + 1e-9
    if name in _COLOURS and isinstance(a, str) and isinstance(b, str) \
            and len(a) == len(b) == 7 and a.startswith("#") and b.startswith("#"):
        return all(abs(int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) <= NEAR_CHANNEL
                   for i in (1, 3, 5))
    return False


def _fix_item(name: str, level: int, shape: "Shape", value: _Value, majority: _Value):
    """The fix items giving ``shape`` the majority's value, or ``None``."""
    target = shape.id
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
    if name == "size":
        frame = shape.text_frame
        host = frame._host()
        paragraphs = [p for p in frame.paragraphs if p.level == level]
        if majority.source not in (None, "slide") and value.source == "slide" and paragraphs \
                and effective.template_size(host, paragraphs[0]._element()) == majority.value \
                and {o.level for o in effective.size_overrides(target, shape)} <= {level}:
            return [{"target": target, "size": 0}]        # no other level's size goes too
        if all(p.level == level for p in frame.paragraphs):
            return [{"target": target, "size": _number(majority.value)}]
        return [{"target": p.address, "size": _number(majority.value)} for p in paragraphs
                if p.text.strip()]
    return None


_FIX_TOOL = {"size": "ppt_format_text", "color": "ppt_format_text"}


@dataclass
class Comparison:
    """What :func:`compare_slides` found: the ``slides`` compared (ids), the ``groups`` of
    corresponding shapes, and the ``within`` tolerance (points)."""

    slides: list[int]
    groups: list[ShapeGroup]
    within: float = WITHIN

    def to_json(self) -> dict[str, Any]:
        """``{"slides", "matched", "agree", "shapes": [...], "only": [...]}``: ``matched``
        how many groups of corresponding shapes, ``agree`` how many of them show no
        difference; ``shapes`` one entry per group that differs or is missing on some
        slides -- ``{"shape": first address, "name", "by": "position" (when matched so),
        "differ": [...], "missing_on": [...]}``; ``only`` the groups on half the slides or
        fewer, ``{"shape", "name", "shapes": [addresses]}``.  Each ``differ`` entry is
        ``{"property", "level" (sizes past level 0), "value": the majority's, "on": how many
        have it, "from", "near": [{"value", "from", "shapes"}], "other": [...], "fix":
        {"tool", "items"}}`` -- ``fix`` for the near-misses only; ``from`` given where the
        sources differ.  A name group whose names differ has ``{"property": "name",
        "value", "on", "other"}`` first."""
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
        measured = {sid: measure(shape) for sid, shape in group.shapes.items()}
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
            if name == "size" and level:
                entry["level"] = level
            entry["value"] = _number(top)
            entry["on"] = best
            near: dict[tuple, list] = {}
            other: dict[tuple, list] = {}
            fixes: list = []
            for sid, value in holders:
                if value.value == top:
                    continue
                shape = group.shapes[sid]
                bucket = near if _near(name, value.value, top, self.within) else other
                bucket.setdefault((value.value, value.source), []).append(shape.id)
                if bucket is near:
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
                         "shapes": shapes}
                        for (value, source), shapes in bucket.items()]
            if near and fixes:
                entry["fix"] = {"tool": _FIX_TOOL.get(name, "ppt_set_shape"), "items": fixes}
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
        raise ValueError("compare at least two slides")
    ids = [slide.slide_id for slide in slides]
    if len(set(ids)) != len(ids):
        repeated = next(sid for sid in ids if ids.count(sid) > 1)
        raise ValueError(f"s:{repeated} is given twice; name each slide once")
    return Comparison(ids, match(slides), float(within))


__all__ = ["Comparison", "NEAR_CHANNEL", "POSITION_REACH", "PROPERTIES", "ShapeGroup",
           "WITHIN", "compare_slides", "kind_of", "match", "measure"]
