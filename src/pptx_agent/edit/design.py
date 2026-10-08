"""Design facts of a slide: what a reader sees, measured, never judged (LP24, LP15).

:func:`design_facts` reports the palette in use, sets of like shapes and the colours they
carry (and whether a legend-like group is there), the largest empty regions of the content
area, alignment lines and near-misses, the shape vocabulary, the text sizes, and the lines
that cross text with their z-order.  :func:`slide_facts` reports the problem facts a
``check`` adds to overflows and collisions: colours that are not theme colours, and lines
of text close to wrapping.

There are no rules here and no thresholds of taste.  The only parameters are tolerances:
``within`` (how far apart two edges may be and still count as a near-miss), the hue
tolerance that groups colours into one family, the size tolerance of "like" shapes, and the
smallest empty region worth reporting.  What the facts *mean* -- whether four hues on four
like boxes is a rainbow, whether an empty band is dead space -- is the application's to
decide (its prompt, or a critique pass over these facts).

Lengths are EMU here; :meth:`DesignFacts.to_json` gives them in points.

For example::

    facts = slide.design_facts()
    for group in facts.color_groups:
        if len(group.accent_hues) > 2 and group.legend is None:
            print("an app's rule fires on", group.shapes)
"""

from __future__ import annotations

import colorsys
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

EMU_PER_POINT = 12700

#: The facts :func:`design_facts` can report, in order.
KINDS = ("palette", "groups", "empty", "alignment", "vocabulary", "text_sizes", "z_order")

#: Theme slots that are neutrals: text and background colours, not accents.
NEUTRAL_SLOTS = {"tx1", "tx2", "bg1", "bg2", "dk1", "dk2", "lt1", "lt2"}
ACCENT_SLOTS = {f"accent{n}" for n in range(1, 7)} | {"hlink", "folHlink"}

#: Presets whose corners are rounded by an adjustment (the first one is the radius).
ROUNDED = {"roundRect", "round1Rect", "round2SameRect", "round2DiagRect", "snipRoundRect",
           "flowChartAlternateProcess", "wedgeRoundRectCallout"}
_LINE_PRESETS = {"line", "straightConnector1", "bentConnector2", "bentConnector3",
                 "bentConnector4", "bentConnector5", "curvedConnector2", "curvedConnector3",
                 "curvedConnector4", "curvedConnector5"}
#: Equal within this (EMU): rounding of the points a tool wrote.
_SAME = 0.05 * EMU_PER_POINT


def _pt(value: float | None) -> float | None:
    return None if value is None else round(float(value) / EMU_PER_POINT, 2)


def _box_json(box) -> dict[str, float]:
    x0, y0, x1, y1 = box
    return {"x": _pt(x0), "y": _pt(y0), "w": _pt(x1 - x0), "h": _pt(y1 - y0)}


# -- colours ---------------------------------------------------------------------------------


_PERCENT = {"lumMod", "lumOff", "tint", "shade", "alpha", "alphaMod", "alphaOff", "satMod",
            "satOff", "sat", "lum", "red", "green", "blue", "redMod", "greenMod", "blueMod",
            "redOff", "greenOff", "blueOff"}


def color_text(color) -> str:
    """A colour as the tools write one: ``accent2 lumMod=75%``, ``#1F4E79``."""
    parts = [f"#{color.value}" if color.kind == "rgb" else color.value]
    for name, value in color.transforms:
        if value is None:
            parts.append(name)
        elif name in _PERCENT:
            parts.append(f"{name}={value / 1000:g}%")
        else:
            parts.append(f"{name}={value}")
    return " ".join(parts)


@dataclass
class ColorUse:
    """One colour as written, where it is used and how often."""

    color: str                 #: as the tools write it: ``accent2 lumMod=75%``, ``#1F4E79``
    hex: str | None            #: what it looks like
    theme: bool                #: a theme (scheme) colour
    count: int = 0
    uses: dict[str, int] = field(default_factory=dict)   #: fill / line / text -> count
    shapes: list[str] = field(default_factory=list)

    def to_json(self) -> dict[str, Any]:
        data: dict[str, Any] = {"color": self.color, "hex": self.hex, "count": self.count,
                                "uses": dict(self.uses)}
        if not self.theme:
            data["theme"] = False
        data["shapes"] = self.shapes[:12]
        return data


@dataclass
class HueFamily:
    """Colours of one hue: a theme slot and its tints and shades, or RGB colours whose hue
    is within the tolerance.  ``accent`` is true for a chromatic family (an accent slot, or a
    saturated RGB colour); the neutrals (text, background, greys) are not."""

    family: str                #: the slot (``accent2``), ``gray``, or ``hue 210``
    accent: bool
    theme: bool
    variants: list[ColorUse] = field(default_factory=list)

    @property
    def count(self) -> int:
        return sum(v.count for v in self.variants)

    def to_json(self) -> dict[str, Any]:
        return {"family": self.family, "accent": self.accent, "theme": self.theme,
                "count": self.count, "variants": [v.to_json() for v in self.variants]}


def _hls(hex_value: str | None) -> tuple[float, float, float] | None:
    if not hex_value:
        return None
    value = hex_value.lstrip("#")
    try:
        r, g, b = (int(value[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return None
    return colorsys.rgb_to_hls(r, g, b)


def _family(color, hex_value: str | None, hue_tolerance: float) -> tuple[str, bool, bool]:
    """``(family, accent, theme)`` of a colour."""
    if color.kind == "scheme":
        slot = color.value
        return slot, slot in ACCENT_SLOTS, True
    hls = _hls(hex_value)
    if hls is None:
        return f"{color.kind} {color.value}", False, False
    hue, light, sat = hls
    if sat < 0.15 or light < 0.08 or light > 0.95:
        return "gray", False, False
    step = max(1.0, hue_tolerance)
    bucket = int(round(hue * 360 / step) * step) % 360
    return f"hue {bucket}", True, False


def _resolve(color, shape) -> str | None:
    try:
        return color.resolve(shape)
    except Exception:  # noqa: BLE001 -- a colour that does not resolve is reported as is
        return None


def _colors_of(shape: "Shape") -> list[tuple[str, Any]]:
    """``(use, Color)`` for what a shape states: its fill (as drawn), outline, text runs."""
    out: list[tuple[str, Any]] = []
    if shape.kind == "shape":
        fill = shape.effective_fill
        if fill is not None and fill.kind == "solid" and fill.color is not None:
            out.append(("fill", fill.color))
        elif fill is not None and fill.kind == "gradient":
            for stop in fill.stops:
                color = getattr(stop, "color", None)
                if color is not None:
                    out.append(("fill", color))
    if shape.kind in ("shape", "connector", "picture"):
        line = shape.line
        if line is not None and line.exists and line.visible is not False and line.color is not None:
            out.append(("line", line.color))
    if shape.kind == "shape":
        frame = shape.text_frame
        if frame is not None:
            for paragraph in frame.paragraphs:
                for run in paragraph.runs:
                    if run.text.strip() and run.color is not None:
                        out.append(("text", run.color))
    return out


# -- shapes ----------------------------------------------------------------------------------


@dataclass
class _Leaf:
    shape: "Shape"
    order: int
    box: tuple[float, float, float, float]      # drawn, slide EMU: x0, y0, x1, y1
    text: str
    line: bool                                   # a line or connector

    @property
    def id(self) -> str:
        return self.shape.id

    @property
    def width(self) -> float:
        return self.box[2] - self.box[0]

    @property
    def height(self) -> float:
        return self.box[3] - self.box[1]


def _leaves(slide: "Slide", region=None) -> list[_Leaf]:
    from .fit import _flatten

    out = []
    for order, shape in enumerate(_flatten(slide.shapes)):
        bounds = shape.drawn_bounds
        if bounds is None:
            continue
        left, top, width, height = bounds
        box = (float(left), float(top), float(left + width), float(top + height))
        if region is not None:
            cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
            if not (region[0] <= cx <= region[2] and region[1] <= cy <= region[3]):
                continue
        text = (shape.text or "").strip() if shape.kind == "shape" else ""
        line = shape.kind == "connector" or (shape.kind == "shape"
                                             and shape.preset in _LINE_PRESETS)
        out.append(_Leaf(shape, order, box, text, line))
    return out


def _is_title(shape: "Shape") -> bool:
    placeholder = shape.placeholder
    return placeholder is not None and placeholder.type in ("title", "ctrTitle")


def _solid_fill(shape: "Shape"):
    if shape.kind != "shape":
        return None
    fill = shape.effective_fill
    if fill is not None and fill.kind == "solid" and fill.color is not None:
        return fill.color
    return None


# -- the facts -------------------------------------------------------------------------------


@dataclass
class Legend:
    """A legend-like group: small swatches, each beside short text (or holding it)."""

    shapes: list[str]                       #: swatches and their texts
    entries: list[dict[str, str]]           #: ``{"swatch", "text", "color"}``
    covers: list[str] = field(default_factory=list)   #: a colour set's fills it shows
    misses: list[str] = field(default_factory=list)   #: a colour set's fills it does not

    def to_json(self) -> dict[str, Any]:
        return {"shapes": self.shapes[:16], "entries": self.entries[:12],
                "covers": self.covers, "misses": self.misses}


@dataclass
class ColorGroup:
    """A set of like shapes (one preset, a shared width or height within the tolerance) and
    the colours their fills carry."""

    preset: str
    shapes: list[str]
    size: dict[str, float]                  #: the typical width and height, EMU
    fills: list[str]                        #: distinct fills, as written
    hues: list[str]                         #: distinct hue families of the fills
    accent_hues: list[str]                  #: the chromatic ones among them
    legend: Legend | None = None

    @property
    def color_coded(self) -> bool:
        return len(self.fills) > 1

    def to_json(self) -> dict[str, Any]:
        return {"preset": self.preset, "count": len(self.shapes), "shapes": self.shapes[:20],
                "size": {k: _pt(v) for k, v in self.size.items()}, "fills": self.fills,
                "hues": self.hues, "accent_hues": self.accent_hues,
                "colour_coded": self.color_coded,
                "legend": self.legend.to_json() if self.legend is not None else None}


@dataclass
class EmptyRegion:
    box: tuple[float, float, float, float]  #: x0, y0, x1, y1 EMU
    share: float                            #: of the area searched, 0-1

    def to_json(self) -> dict[str, Any]:
        return {"box": _box_json(self.box), "share": round(self.share, 3)}


@dataclass
class AlignLine:
    edge: str                               #: left, center, right, top, middle, bottom
    at: float                               #: EMU
    shapes: list[str]

    def to_json(self) -> dict[str, Any]:
        return {"edge": self.edge, "at": _pt(self.at), "shapes": self.shapes[:16],
                "count": len(self.shapes)}


@dataclass
class NearMiss:
    edge: str
    shape: str
    other: str
    offset: float                           #: shape's edge minus other's, EMU

    def to_json(self) -> dict[str, Any]:
        return {"edge": self.edge, "shape": self.shape, "other": self.other,
                "offset": _pt(self.offset)}


@dataclass
class LineOverText:
    """A line or connector crossing the text of a text-bearing shape, and which is in front:
    ``in_front`` is true when the line is drawn after (over) the shape."""

    line: str
    shape: str
    crossing: float                         #: length of line inside the text area, EMU
    in_front: bool
    line_z: int
    shape_z: int
    shape_opaque: bool

    def to_json(self) -> dict[str, Any]:
        return {"line": self.line, "shape": self.shape, "crossing": _pt(self.crossing),
                "line_in_front": self.in_front, "line_z": self.line_z,
                "shape_z": self.shape_z, "shape_opaque": self.shape_opaque}


@dataclass
class DesignFacts:
    """What :func:`design_facts` measured; a field is ``None`` when not asked for."""

    slide: int
    area: tuple[float, float, float, float]
    palette: list[HueFamily] | None = None
    non_theme: list[ColorUse] | None = None
    color_groups: list[ColorGroup] | None = None
    legends: list[Legend] | None = None
    empty: list[EmptyRegion] | None = None
    panels: list[dict[str, Any]] | None = None
    alignment: dict[str, Any] | None = None
    vocabulary: dict[str, Any] | None = None
    text_sizes: list[dict[str, Any]] | None = None
    z_order: list[LineOverText] | None = None

    def to_json(self) -> dict[str, Any]:
        """The facts in points, lists capped (each with its total), for a tool result."""
        out: dict[str, Any] = {"slide": f"s:{self.slide}", "area": _box_json(self.area)}
        if self.palette is not None:
            out["palette"] = {
                "families": [f.to_json() for f in self.palette[:12]],
                "accent_families": sum(1 for f in self.palette if f.accent),
                "non_theme": [u.to_json() for u in (self.non_theme or [])[:12]]}
        if self.color_groups is not None:
            groups = self.color_groups
            out["groups"] = {"sets": [g.to_json() for g in groups[:10]], "total": len(groups),
                             "legends": [l.to_json() for l in (self.legends or [])[:4]]}
        if self.empty is not None:
            out["empty"] = {"regions": [r.to_json() for r in self.empty],
                            "panels": (self.panels or [])[:10]}
        if self.alignment is not None:
            a = self.alignment
            out["alignment"] = {
                "lines": [line.to_json() for line in a["lines"][:12]],
                "lines_total": len(a["lines"]),
                "near_misses": [miss.to_json() for miss in a["near_misses"][:16]],
                "near_misses_total": len(a["near_misses"]), "within": _pt(a["within"])}
        if self.vocabulary is not None:
            v = self.vocabulary
            out["vocabulary"] = {
                "presets": v["presets"],
                "corners": {"square": v["corners"]["square"],
                            "rounded": [{"radius": _pt(r), "count": c}
                                        for r, c in v["corners"]["rounded"]]},
                "dashes": v["dashes"], "connectors": v["connectors"]}
        if self.text_sizes is not None:
            out["text_sizes"] = [{"size": s["size"], "count": s["count"],
                                  "shapes": s["shapes"][:12]} for s in self.text_sizes]
        if self.z_order is not None:
            out["lines_over_text"] = [z.to_json() for z in self.z_order[:20]]
            if len(self.z_order) > 20:
                out["lines_over_text_total"] = len(self.z_order)
        return out


def design_facts(slide: "Slide", *, region=None, within: float = 2 * EMU_PER_POINT,
                 include: Iterable[str] | None = None, hue_tolerance: float = 20.0,
                 size_tolerance: float = 0.1, min_empty: float = 36 * EMU_PER_POINT,
                 regions: int = 4) -> DesignFacts:
    """The design facts of ``slide``, or of the shapes centred in ``region``
    (``(left, top, width, height)`` EMU; default the slide's content area for the empty
    regions and every shape for the rest).

    * ``within``: EMU; edges and centres closer than this, but not equal, are near-misses.
    * ``hue_tolerance``: degrees; RGB colours whose hue rounds to one step are one family.
    * ``size_tolerance``: a fraction; shapes whose width or height agree within it (and
      share a preset) are like shapes.
    * ``min_empty``: EMU; an empty region narrower or lower than this is not reported.
    * ``include``: some of :data:`KINDS`; all by default.

    For example::

        facts = slide.design_facts(within=Pt(3))
        facts.to_json()["palette"]["accent_families"]
    """
    wanted = set(include or KINDS)
    unknown = wanted - set(KINDS)
    if unknown:
        raise ValueError(f"unknown design facts {sorted(unknown)}; one of {list(KINDS)}")
    if region is not None:
        left, top, width, height = region
        box = (float(left), float(top), float(left + width), float(top + height))
    else:
        left, top, width, height = slide.content_area
        box = (float(left), float(top), float(left + width), float(top + height))
    leaves = _leaves(slide, box if region is not None else None)
    facts = DesignFacts(slide.slide_id, box)
    if "palette" in wanted:
        facts.palette, facts.non_theme = _palette(leaves, hue_tolerance)
    if "groups" in wanted:
        facts.color_groups, facts.legends = _groups(leaves, hue_tolerance, size_tolerance)
    if "empty" in wanted:
        facts.empty, facts.panels = _empty(leaves, box, min_empty, regions, _visible(slide))
    if "alignment" in wanted:
        facts.alignment = _alignment(leaves, within)
    if "vocabulary" in wanted:
        facts.vocabulary = _vocabulary(leaves)
    if "text_sizes" in wanted:
        facts.text_sizes = _text_sizes(leaves)
    if "z_order" in wanted:
        facts.z_order = _lines_over_text(slide, {leaf.id for leaf in leaves})
    return facts


# -- palette ---------------------------------------------------------------------------------


def _palette(leaves: list[_Leaf], hue_tolerance: float):
    uses: dict[str, ColorUse] = {}
    families: dict[str, HueFamily] = {}
    for leaf in leaves:
        for use, color in _colors_of(leaf.shape):
            key = color_text(color)
            entry = uses.get(key)
            if entry is None:
                hex_value = _resolve(color, leaf.shape)
                entry = uses[key] = ColorUse(key, hex_value, color.kind == "scheme")
                name, accent, theme = _family(color, hex_value, hue_tolerance)
                family = families.setdefault(name, HueFamily(name, accent, theme))
                family.variants.append(entry)
            entry.count += 1
            entry.uses[use] = entry.uses.get(use, 0) + 1
            if leaf.id not in entry.shapes:
                entry.shapes.append(leaf.id)
    ordered = sorted(families.values(), key=lambda f: (-f.count, f.family))
    for family in ordered:
        family.variants.sort(key=lambda v: -v.count)
    non_theme = sorted((u for u in uses.values() if not u.theme), key=lambda u: -u.count)
    return ordered, non_theme


# -- colour-coded sets and legends -------------------------------------------------------------


def _close(a: float, b: float, tolerance: float) -> bool:
    return abs(a - b) <= tolerance * max(abs(a), abs(b), 1.0)


def _legends(leaves: list[_Leaf], hue_tolerance: float) -> list[Legend]:
    """Small filled shapes each beside short text (or holding it), two or more of them in a
    row or a column: what reads as a legend."""
    texts = [leaf for leaf in leaves if leaf.text and not leaf.line
             and len(leaf.text) <= 40 and "\n" not in leaf.text]
    entries = []
    for leaf in leaves:
        if leaf.line or leaf.shape.kind != "shape":
            continue
        color = _solid_fill(leaf.shape)
        if color is None:
            continue
        side = max(leaf.width, leaf.height)
        if side > 40 * EMU_PER_POINT:
            continue
        label = None
        if leaf.text and len(leaf.text) <= 40:
            label = leaf
        else:
            best = None
            for text in texts:
                if text is leaf:
                    continue
                gap_x = max(text.box[0] - leaf.box[2], leaf.box[0] - text.box[2], 0)
                cy_leaf = (leaf.box[1] + leaf.box[3]) / 2
                overlaps_y = text.box[1] - _SAME <= cy_leaf <= text.box[3] + _SAME
                if overlaps_y and gap_x <= 24 * EMU_PER_POINT and text.box[0] >= leaf.box[0]:
                    if best is None or gap_x < best[0]:
                        best = (gap_x, text)
            label = best[1] if best else None
        if label is not None:
            entries.append((leaf, label, color_text(color)))
    if len(entries) < 2:
        return []
    # Entries in one row (shared middle) or one column (shared left), similar swatches.
    legends: list[Legend] = []
    used: set[int] = set()
    tol = 4 * EMU_PER_POINT
    for i, (swatch, label, color) in enumerate(entries):
        if i in used:
            continue
        members = [i]
        for j in range(i + 1, len(entries)):
            if j in used:
                continue
            other = entries[j][0]
            same_size = _close(swatch.width, other.width, 0.25) and _close(
                swatch.height, other.height, 0.25)
            column = abs(swatch.box[0] - other.box[0]) <= tol
            row = abs((swatch.box[1] + swatch.box[3]) / 2 - (other.box[1] + other.box[3]) / 2) <= tol
            if same_size and (column or row):
                members.append(j)
        if len(members) >= 2:
            used.update(members)
            shapes, rows = [], []
            for k in members:
                s, t, c = entries[k]
                shapes += [s.id] if s is t else [s.id, t.id]
                rows.append({"swatch": s.id, "text": t.text, "color": c})
            legends.append(Legend(shapes, rows))
    return legends


def _groups(leaves: list[_Leaf], hue_tolerance: float, size_tolerance: float):
    legends = _legends(leaves, hue_tolerance)
    in_legend = {shape for legend in legends for shape in legend.shapes}
    candidates = [leaf for leaf in leaves if leaf.shape.kind == "shape" and not leaf.line
                  and not _is_title(leaf.shape) and leaf.id not in in_legend
                  and _solid_fill(leaf.shape) is not None]
    parent = list(range(len(candidates)))

    def root(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, a in enumerate(candidates):
        for j in range(i + 1, len(candidates)):
            b = candidates[j]
            if a.shape.preset != b.shape.preset:
                continue
            if _close(a.width, b.width, size_tolerance) and _close(a.height, b.height, size_tolerance):
                parent[root(j)] = root(i)
            elif _close(a.height, b.height, size_tolerance) and bool(a.text) == bool(b.text):
                # Bars: one height, widths as long as their data.
                parent[root(j)] = root(i)
    sets: dict[int, list[_Leaf]] = {}
    for i, leaf in enumerate(candidates):
        sets.setdefault(root(i), []).append(leaf)
    groups = []
    for members in sets.values():
        if len(members) < 2:
            continue
        fills: list[str] = []
        hues: list[str] = []
        accents: list[str] = []
        for leaf in members:
            color = _solid_fill(leaf.shape)
            key = color_text(color)
            if key not in fills:
                fills.append(key)
            name, accent, _ = _family(color, _resolve(color, leaf.shape), hue_tolerance)
            if name not in hues:
                hues.append(name)
                if accent:
                    accents.append(name)
        widths = sorted(leaf.width for leaf in members)
        heights = sorted(leaf.height for leaf in members)
        group = ColorGroup(members[0].shape.preset or "custom", [leaf.id for leaf in members],
                           {"w": widths[len(widths) // 2], "h": heights[len(heights) // 2]},
                           fills, hues, accents)
        best = None
        for legend in legends:
            shown = {entry["color"] for entry in legend.entries}
            covers = [fill for fill in fills if fill in shown]
            if covers and (best is None or len(covers) > len(best[1])):
                best = (legend, covers)
        if best is not None:
            legend, covers = best
            group.legend = Legend(legend.shapes, legend.entries, covers,
                                  [fill for fill in fills if fill not in covers])
        groups.append(group)
    groups.sort(key=lambda g: (-len(g.shapes), g.preset))
    return groups, legends


# -- empty regions ---------------------------------------------------------------------------


def _visible(slide: "Slide") -> dict[str, tuple[float, float, float, float]]:
    """Where each shape shows something: its drawn box, or -- for text with neither fill
    nor outline -- the text itself (a frame taller than its text leaves the rest empty)."""
    from .fit import _items

    try:
        items = _items(slide, [None], {})
    except Exception:  # noqa: BLE001 -- no measurement: the boxes stand in
        return {}
    return {item.shape.id: (item.region if item.text and item.region is not None
                            else item.box) for item in items}


def _empty(leaves: list[_Leaf], area, min_size: float, count: int, visible=None):
    """The largest empty rectangles of ``area`` among what the shapes show, greedily: the
    largest, then the largest that does not overlap it, and so on."""
    visible = visible or {}
    ax0, ay0, ax1, ay1 = area
    total = max((ax1 - ax0) * (ay1 - ay0), 1.0)
    obstacles = []
    panels = []
    for leaf in leaves:
        if _is_title(leaf.shape):
            continue
        shown = visible.get(leaf.id, leaf.box)
        x0, y0 = max(shown[0], ax0), max(shown[1], ay0)
        x1, y1 = min(shown[2], ax1), min(shown[3], ay1)
        if x1 <= x0 and y1 <= y0:
            continue
        if x1 < x0 or y1 < y0:
            continue
        if (x1 - x0) * (y1 - y0) >= 0.9 * total:
            continue                          # a background, not content
        fill = _solid_fill(leaf.shape)
        if not leaf.text and fill is not None and (x1 - x0) * (y1 - y0) >= 0.04 * total:
            hex_value = _resolve(fill, leaf.shape)
            hls = _hls(hex_value)
            panels.append({"shape": leaf.id, "fill": color_text(fill), "hex": hex_value,
                           "lightness": round(hls[1], 3) if hls else None,
                           "box": _box_json((x0, y0, x1, y1))})
        # A line is an obstacle as thin as it is drawn.
        obstacles.append((x0, y0, max(x1, x0 + 1), max(y1, y0 + 1)))
    found: list[EmptyRegion] = []
    blocked = list(obstacles)
    for _ in range(count):
        best = _largest_empty(blocked, area)
        if best is None:
            break
        x0, y0, x1, y1 = best
        if x1 - x0 < min_size or y1 - y0 < min_size:
            break
        found.append(EmptyRegion(best, (x1 - x0) * (y1 - y0) / total))
        blocked.append(best)
    return found, panels


def _largest_empty(obstacles, area):
    ax0, ay0, ax1, ay1 = area
    xs = sorted({ax0, ax1, *(v for o in obstacles for v in (o[0], o[2]) if ax0 < v < ax1)})
    ys = sorted({ay0, ay1, *(v for o in obstacles for v in (o[1], o[3]) if ay0 < v < ay1)})
    nx, ny = len(xs) - 1, len(ys) - 1
    if nx <= 0 or ny <= 0:
        return None
    x_index = {v: i for i, v in enumerate(xs)}
    y_index = {v: i for i, v in enumerate(ys)}
    import bisect

    full = [[False] * nx for _ in range(ny)]
    for o in obstacles:
        i0 = bisect.bisect_right(xs, o[0]) - 1
        i1 = bisect.bisect_left(xs, o[2])
        j0 = bisect.bisect_right(ys, o[1]) - 1
        j1 = bisect.bisect_left(ys, o[3])
        for j in range(max(j0, 0), min(j1, ny)):
            row = full[j]
            for i in range(max(i0, 0), min(i1, nx)):
                row[i] = True
    del x_index, y_index
    widths = [xs[i + 1] - xs[i] for i in range(nx)]
    prefix = [0.0]
    for w in widths:
        prefix.append(prefix[-1] + w)
    heights = [0.0] * nx
    starts = [0] * nx
    best = None
    best_area = 0.0
    for j in range(ny):
        h = ys[j + 1] - ys[j]
        for i in range(nx):
            if full[j][i]:
                heights[i] = 0.0
            else:
                if heights[i] == 0.0:
                    starts[i] = j
                heights[i] += h
        # Largest rectangle in a histogram with weighted columns.
        stack: list[int] = []
        for i in range(nx + 1):
            current = heights[i] if i < nx else 0.0
            while stack and heights[stack[-1]] >= current:
                top = stack.pop()
                height = heights[top]
                left = stack[-1] + 1 if stack else 0
                width = prefix[i] - prefix[left]
                area_ = width * height
                if height > 0 and area_ > best_area:
                    best_area = area_
                    best = (xs[left], ys[j + 1] - height, xs[i], ys[j + 1])
            stack.append(i)
    return best


# -- alignment -------------------------------------------------------------------------------

_EDGES = (("left", 0, lambda b: b[0]), ("center", 0, lambda b: (b[0] + b[2]) / 2),
          ("right", 0, lambda b: b[2]), ("top", 1, lambda b: b[1]),
          ("middle", 1, lambda b: (b[1] + b[3]) / 2), ("bottom", 1, lambda b: b[3]))


def _alignment(leaves: list[_Leaf], within: float) -> dict[str, Any]:
    shapes = [leaf for leaf in leaves if not leaf.line and not _is_title(leaf.shape)
              and leaf.width > 0 and leaf.height > 0]
    lines: list[AlignLine] = []
    misses: list[NearMiss] = []
    for edge, _, value in _EDGES:
        points = sorted((value(leaf.box), leaf.id) for leaf in shapes)
        clusters: list[list[tuple[float, str]]] = []
        for point in points:
            if clusters and point[0] - clusters[-1][-1][0] <= _SAME:
                clusters[-1].append(point)
            else:
                clusters.append([point])
        for cluster in clusters:
            if len(cluster) >= 2:
                lines.append(AlignLine(edge, cluster[0][0], [shape for _, shape in cluster]))
        # Near-misses: neighbouring clusters closer than `within`, not equal.
        for first, second in zip(clusters, clusters[1:]):
            gap = second[0][0] - first[-1][0]
            if _SAME < gap <= within:
                # The smaller cluster is the one off the line.
                off, line = (second, first) if len(second) <= len(first) else (first, second)
                for at, shape in off:
                    misses.append(NearMiss(edge, shape, line[0][1], at - line[0][0]))
    lines.sort(key=lambda line: (-len(line.shapes), line.edge, line.at))
    return {"lines": lines, "near_misses": misses, "within": within}


# -- vocabulary and text sizes -------------------------------------------------------------------


def _vocabulary(leaves: list[_Leaf]) -> dict[str, Any]:
    presets: dict[str, int] = {}
    square = 0
    rounded: dict[float, int] = {}
    dashes: dict[str, int] = {}
    connectors: dict[str, int] = {}
    for leaf in leaves:
        shape = leaf.shape
        if shape.kind == "connector":
            kind = shape.preset or "straightConnector1"
            connectors[kind] = connectors.get(kind, 0) + 1
        elif shape.kind == "shape":
            preset = shape.preset or "custom"
            if shape.placeholder is not None:
                preset = "placeholder"
            presets[preset] = presets.get(preset, 0) + 1
            if preset in ROUNDED:
                try:
                    adj = shape.adjustments[0]
                except Exception:  # noqa: BLE001 -- a preset without the usual adjustment
                    adj = 16667
                radius = round(adj / 100000 * min(leaf.width, leaf.height) / EMU_PER_POINT
                               * 2) / 2 * EMU_PER_POINT
                rounded[radius] = rounded.get(radius, 0) + 1
            elif preset in ("rect", "flowChartProcess", "snip1Rect") and not leaf.line:
                square += 1
        if shape.kind in ("shape", "connector"):
            line = shape.line
            if line is not None and line.exists and line.visible is not False:
                dash = line.dash or "solid"
                if leaf.line or line.width or line.color is not None:
                    dashes[dash] = dashes.get(dash, 0) + 1
    return {"presets": dict(sorted(presets.items(), key=lambda kv: -kv[1])),
            "corners": {"square": square, "rounded": sorted(rounded.items())},
            "dashes": dashes, "connectors": connectors}


def _text_sizes(leaves: list[_Leaf]) -> list[dict[str, Any]]:
    sizes: dict[float, dict[str, Any]] = {}
    for leaf in leaves:
        if not leaf.text:
            continue
        frame = leaf.shape.text_frame
        if frame is None:
            continue
        for paragraph in frame.paragraphs:
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                size = round(float(run.effective_size), 1)
                entry = sizes.setdefault(size, {"size": size, "count": 0, "shapes": []})
                entry["count"] += 1
                if leaf.id not in entry["shapes"]:
                    entry["shapes"].append(leaf.id)
    return sorted(sizes.values(), key=lambda e: -e["size"])


# -- lines over text -------------------------------------------------------------------------


def _lines_over_text(slide: "Slide", keep: set[str]) -> list[LineOverText]:
    """Every line crossing a text-bearing shape's text, in front of it or behind it."""
    from .fit import _CROSSING, _items
    from .geometry import length_inside

    items = _items(slide, [None], {})
    texts = [item for item in items if item.text and item.region is not None]
    out = []
    for line in items:
        if line.route is None or line.shape.id not in keep:
            continue
        for item in texts:
            if item is line or item.shape.id not in keep:
                continue
            region = item.area if item.opaque and item.area is not None else item.region
            crossed = length_inside(line.route, region)
            if crossed > _CROSSING:
                out.append(LineOverText(line.shape.id, item.shape.id, crossed,
                                        line.order > item.order, line.order, item.order,
                                        item.opaque))
    return out


# -- the problem facts (LP15) ------------------------------------------------------------------


def slide_facts(slide: "Slide") -> dict[str, Any]:
    """The problem facts a ``check`` adds to overflows and collisions, in points: fills,
    outlines and text colours that are not theme colours (with their shapes), and the wrap
    margin of every title and text shape whose line is close to wrapping.

    For example::

        slide.facts()["non_theme_colors"]
    """
    from .fit import slide_report

    leaves = _leaves(slide)
    colors: dict[tuple[str, str], dict[str, Any]] = {}
    for leaf in leaves:
        for use, color in _colors_of(leaf.shape):
            if color.kind == "scheme":
                continue
            key = (color_text(color), use)
            entry = colors.setdefault(key, {"color": key[0], "hex": _resolve(color, leaf.shape),
                                            "use": use, "shapes": []})
            if leaf.id not in entry["shapes"]:
                entry["shapes"].append(leaf.id)
    _, fits = slide_report(slide)
    wrap = []
    for leaf in leaves:
        fit = fits.get(leaf.id)
        if fit is None:
            continue
        if _is_title(leaf.shape) or fit.near_wrap:
            wrap.append({"shape": leaf.id, "title": _is_title(leaf.shape),
                         "margin": _pt(fit.margin_to_wrap), "near_wrap": bool(fit.near_wrap),
                         "lines": list(fit.lines)})
    return {"non_theme_colors": list(colors.values()), "wrap_margins": wrap}


__all__ = ["ColorGroup", "ColorUse", "DesignFacts", "EmptyRegion", "HueFamily", "KINDS",
           "Legend", "LineOverText", "NearMiss", "AlignLine", "design_facts", "slide_facts"]
