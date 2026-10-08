"""Layout facts about the shapes an edit touched, each with the exact change that resolves it.

After an edit, :func:`layout_facts` looks at the shapes it touched and reports four kinds of
fact, measured against the slide's own other shapes, never against a rule of taste:

* ``near_alignment``: a box (a shape with a fill or an outline) whose edge or centre is 1-3 pt
  off a line that two or more of its neighbours in the same column or row are exactly on,
  one of them a box like it, while it is exactly on none of theirs;
* ``uneven_gap``: in a row or column of like boxes, one gap that differs (by 1-12 pt, or up
  to half the gap) from the gap the others agree on -- four or more boxes, or three boxes
  with text -- where moving the touched box alone makes them agree;
* ``text_size``: a paragraph of a touched box whose size differs from the size every like
  box (same geometry, size and fill, as many paragraphs) has there -- unless the touched
  box states its size and the others inherit theirs (a size just chosen, mid-build);
* ``label_distance``: a label (a one-line text box without fill) more than twice as far from
  its marker (a small box without text; labels and markers matched one to one, nearest
  first) as the other labels of its kind (one size and height) are from theirs, and 8 pt
  more -- when at most a fifth of them are that far, and no line ends nearer.

"Like" is the same preset with width and height within 12%.  Each fact names the shape (or
paragraph) to change and the change -- a position in points for ``ppt_set_shape``, or a size
for ``ppt_format_text`` -- that makes it agree with the others.  Whether to make it is the
caller's judgement: a difference may be intended.  The tolerances were tuned on the trial
outputs for precision (the tool roadmap has the numbers): on 85 trial decks they report 3
facts, and none on the fixture decks with every shape touched.

For example::

    for fact in layout_facts(slide, ["257.12"]):
        print(fact.to_json())
        # {'kind': 'near_alignment', 'fact': '257.12: top edge 2.0 pt below the line 257.9,
        #  257.10, 257.11 are on', 'fix': 'ppt_set_shape 257.12 y=150'}
"""

from __future__ import annotations

from dataclasses import dataclass, field
from statistics import median
from typing import TYPE_CHECKING, Any, Iterable

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

EMU_PER_POINT = 12700
_PT = EMU_PER_POINT
#: Equal within this: the rounding of points a tool wrote.
_SAME = 0.1 * _PT

#: Edges and centres this close to a line, but not on it, are near-misses.
ALIGN_MIN = 1 * _PT
ALIGN_WITHIN = 3 * _PT
#: Neighbours farther apart than this (the gap between them) are not checked for alignment.
ALIGN_REACH = 36 * _PT
#: Gaps within this of each other agree.
GAP_SAME = 1.0 * _PT
#: A gap that differs from the others by more than this share of theirs, and more than
#: ``GAP_CHOICE``, is a choice.
GAP_SHARE = 0.5
GAP_CHOICE = 12 * _PT
#: Like shapes: width and height within this fraction of each other.
LIKE = 0.12
#: A label is "far" when its gap to its anchor exceeds the slide's median label gap by this
#: factor and by at least ``LABEL_EXTRA``.
LABEL_FACTOR = 2.0
LABEL_EXTRA = 8 * _PT
#: ... and only when at most this share of the set is that far (else it is not one set of
#: labels beside their markers).
LABEL_OUTLIERS = 0.2
#: The largest anchor (a marker, bubble or diamond), either side, and the largest label box.
ANCHOR_MAX = 60 * _PT
LABEL_MAX_HEIGHT = 40 * _PT

_LINES = {"line", "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4",
          "bentConnector5", "curvedConnector2", "curvedConnector3", "curvedConnector4",
          "curvedConnector5"}


@dataclass
class LayoutFact:
    """One fact: what, about which shapes, and the change that resolves it."""

    kind: str
    shape: str                      #: the shape the fix changes
    other: str | None               #: the shape (or the set's first) it is measured against
    amount: float                   #: how far off, EMU (text_size: points)
    fix: dict[str, float]           #: ``x``/``y`` (pt, for ppt_set_shape) or ``size``
    text: str
    shapes: list[str] = field(default_factory=list)

    def to_json(self) -> dict[str, Any]:
        tool = "ppt_format_text" if "size" in self.fix else "ppt_set_shape"
        args = " ".join(f"{key}={value:g}" for key, value in self.fix.items())
        return {"kind": self.kind, "fact": self.text,
                "fix": f"{tool} {self.shape} {args}"}


@dataclass
class _Box:
    shape: "Shape"
    id: str
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    preset: str | None
    placeholder: bool
    filled: bool

    @property
    def w(self) -> float:
        return self.x1 - self.x0

    @property
    def h(self) -> float:
        return self.y1 - self.y0

    def edge(self, name: str) -> float:
        return {"left": self.x0, "center": (self.x0 + self.x1) / 2, "right": self.x1,
                "top": self.y0, "middle": (self.y0 + self.y1) / 2, "bottom": self.y1}[name]


def _pt(value: float) -> float:
    return round(value / _PT, 2)


def _boxes(slide: "Slide") -> list[_Box]:
    from .fit import _flatten

    out = []
    for shape in _flatten(slide.shapes):
        if shape.kind not in ("shape", "picture", "graphic_frame"):
            continue
        preset = shape.preset if shape.kind == "shape" else shape.kind
        if preset in _LINES:
            continue
        bounds = shape.drawn_bounds
        if bounds is None:
            continue
        left, top, width, height = (float(v) for v in bounds)
        if width <= 0 or height <= 0:
            continue
        text = (shape.text or "").strip() if shape.kind == "shape" else ""
        filled = False
        if shape.kind == "shape":
            fill = shape.effective_fill
            filled = fill is not None and fill.kind not in (None, "none")
            line = shape.line
            filled = filled or bool(line is not None and line.exists and line.visible is not False
                                    and (line.width or line.color is not None))
        out.append(_Box(shape, shape.id, left, top, left + width, top + height, text,
                        preset, shape.placeholder is not None, filled))
    return out


def _like(a: _Box, b: _Box) -> bool:
    if a.preset != b.preset:
        return False

    def close(p: float, q: float) -> bool:
        return abs(p - q) <= LIKE * max(p, q)

    return close(a.w, b.w) and close(a.h, b.h)


def _overlap(a0: float, a1: float, b0: float, b1: float) -> float:
    return min(a1, b1) - max(a0, b0)


# -- near alignment ----------------------------------------------------------------------------


def _near_alignment(touched: list[_Box], boxes: list[_Box]) -> list[LayoutFact]:
    """A touched box whose edge or centre is within ``ALIGN_WITHIN`` of a line that two or
    more of its neighbours in the same column (or row) are exactly on -- one of them a box
    like it -- while it is exactly on none of theirs."""
    facts = []
    for s in touched:
        # Boxes with a fill or an outline: a plain text box goes where its text belongs.
        if s.placeholder or not s.filled:
            continue
        for axis, edges in ((0, ("left", "center", "right")), (1, ("top", "middle", "bottom"))):
            # A column (axis 0: edges are x) holds shapes above and below each other; a row
            # (axis 1) shapes side by side.
            neighbours = []
            for t in boxes:
                if t is s or t.placeholder:
                    continue
                if axis == 0:
                    across = _overlap(s.x0, s.x1, t.x0, t.x1)
                    gap = max(t.y0 - s.y1, s.y0 - t.y1)
                else:
                    across = _overlap(s.y0, s.y1, t.y0, t.y1)
                    gap = max(t.x0 - s.x1, s.x0 - t.x1)
                if across <= 0 or gap < 0:
                    continue
                neighbours.append((t, gap))
            # The column (or row) it is in, at any distance, with a neighbour within reach.
            if len(neighbours) < 2 or min(gap for _, gap in neighbours) > ALIGN_REACH:
                continue
            neighbours = [t for t, _ in neighbours]
            if any(abs(s.edge(edge) - t.edge(edge)) <= _SAME
                   for edge in edges for t in neighbours):
                continue                             # already on a line with a neighbour
            best = None
            for edge in edges:
                for t in neighbours:
                    delta = s.edge(edge) - t.edge(edge)
                    if not ALIGN_MIN <= abs(delta) <= ALIGN_WITHIN:
                        continue
                    on_line = [u for u in neighbours
                               if abs(u.edge(edge) - t.edge(edge)) <= _SAME]
                    if len(on_line) < 2 or not any(_like(s, u) and u.filled for u in on_line):
                        continue                     # not a line of shapes like this one
                    if best is None or abs(delta) < abs(best[2]):
                        best = (edge, t, delta, on_line)
            if best is None:
                continue
            edge, t, delta, on_line = best
            key = "x" if axis == 0 else "y"
            target = _pt((s.x0 if axis == 0 else s.y0) - delta)
            side = ("right of" if delta > 0 else "left of") if axis == 0 else \
                ("below" if delta > 0 else "above")
            what = {"center": "centre", "middle": "middle"}.get(edge, f"{edge} edge")
            facts.append(LayoutFact(
                "near_alignment", s.id, t.id, abs(delta), {key: target},
                f"{s.id}: {what} {_pt(abs(delta))} pt {side} the line "
                f"{', '.join(u.id for u in on_line[:3])} are on", [s.id] + [u.id for u in on_line]))
    return facts


# -- uneven gaps -------------------------------------------------------------------------------


def _lines_of(boxes: list[_Box], axis: int) -> list[list[_Box]]:
    """Rows (axis 0: left to right) or columns (axis 1) of like shapes, three or more."""
    # Boxes with a fill or an outline: a plain label's place follows what it labels.
    candidates = [b for b in boxes if not b.placeholder and b.filled]
    seen: set[str] = set()
    out = []
    for b in candidates:
        if b.id in seen:
            continue
        members = [c for c in candidates if _like(b, c) and (
            abs(c.edge("middle") - b.edge("middle")) <= _SAME * 5 if axis == 0
            else abs(c.edge("center") - b.edge("center")) <= _SAME * 5)]
        if len(members) < 3:
            continue
        members.sort(key=lambda c: c.x0 if axis == 0 else c.y0)
        gaps = [(m.x0 - p.x1) if axis == 0 else (m.y0 - p.y1)
                for p, m in zip(members, members[1:])]
        if any(gap < 0 for gap in gaps):
            continue
        seen.update(m.id for m in members)
        out.append(members)
    return out


def _uneven_gaps(touched_ids: set[str], boxes: list[_Box]) -> list[LayoutFact]:
    facts = []
    for axis in (0, 1):
        for members in _lines_of(boxes, axis):
            gaps = [(m.x0 - p.x1) if axis == 0 else (m.y0 - p.y1)
                    for p, m in zip(members, members[1:])]
            if len(gaps) == 2:
                facts += _two_gaps(touched_ids, members, gaps, axis)
                continue
            # The gap most others agree with; none: data positions, not a spacing.
            agree = max(gaps, key=lambda g: sum(abs(g - o) <= GAP_SAME for o in gaps))
            agreeing = sum(abs(agree - o) <= GAP_SAME for o in gaps)
            if agreeing < 2 or agreeing == len(gaps) or agreeing * 2 < len(gaps):
                continue
            off = [i for i, g in enumerate(gaps) if abs(g - agree) > GAP_SAME]
            if any(abs(gaps[i] - agree) > max(GAP_SHARE * agree, GAP_CHOICE) for i in off):
                continue                              # a different gap on purpose
            fix = None
            if len(off) == 1 and off[0] == len(gaps) - 1:
                fix = (members[-1], agree - gaps[-1])                  # the last one
            elif len(off) == 1 and off[0] == 0:
                fix = (members[0], gaps[0] - agree)                    # the first one
            elif len(off) == 2 and off[1] == off[0] + 1 and \
                    abs((gaps[off[0]] - agree) + (gaps[off[1]] - agree)) <= GAP_SAME:
                fix = (members[off[1]], agree - gaps[off[0]])          # one between
            if fix is None:
                continue
            mover, shift = fix
            if mover.id not in touched_ids:
                continue
            key = "x" if axis == 0 else "y"
            target = _pt((mover.x0 if axis == 0 else mover.y0) + shift)
            worst = max(abs(gaps[i] - agree) for i in off)
            facts.append(LayoutFact(
                "uneven_gap", mover.id, members[0].id, worst, {key: target},
                f"{mover.id}: gap {_pt(gaps[off[0]])} pt where its {len(members)} like shapes "
                f"have {_pt(agree)} pt", [m.id for m in members]))
    return facts


def _two_gaps(touched_ids: set[str], members: list[_Box], gaps: list[float],
              axis: int) -> list[LayoutFact]:
    """Three like boxes with text (a sequence, not data points) whose two gaps differ a
    little: the touched one moves to make them equal."""
    if not all(m.text for m in members):
        return []
    difference = gaps[1] - gaps[0]
    if not GAP_SAME < abs(difference) <= max(GAP_SHARE * min(gaps), GAP_CHOICE):
        return []
    first, middle, last = members
    if last.id in touched_ids:
        mover, shift, other = last, -difference, gaps[0]
    elif first.id in touched_ids:
        mover, shift, other = first, difference, gaps[1]
    elif middle.id in touched_ids:
        mover, shift, other = middle, difference / 2, (gaps[0] + gaps[1]) / 2
    else:
        return []
    key = "x" if axis == 0 else "y"
    target = _pt((mover.x0 if axis == 0 else mover.y0) + shift)
    return [LayoutFact(
        "uneven_gap", mover.id, first.id, abs(difference), {key: target},
        f"{mover.id}: the gaps in its row of 3 like shapes are {_pt(gaps[0])} and "
        f"{_pt(gaps[1])} pt; this makes both {_pt(other)}", [m.id for m in members])]


# -- text sizes --------------------------------------------------------------------------------


def _sizes_of(box: _Box) -> list[float | None] | None:
    """Each paragraph's text size, pt (the size most of its characters have); None for no
    text, or text that autofit may shrink (resizing the shape to its text is fine)."""
    frame = box.shape.text_frame if box.shape.kind == "shape" else None
    if frame is None or not box.text or frame.autofit == "normal":
        return None
    out: list[float | None] = []
    for paragraph in frame.paragraphs:
        counts: dict[float, int] = {}
        for run in paragraph.runs:
            if run.text.strip():
                size = round(float(run.effective_size), 1)
                counts[size] = counts.get(size, 0) + len(run.text.strip())
        out.append(max(counts, key=counts.get) if counts else None)
    return out


def _explicit_size(box: _Box) -> bool:
    """Whether some run of the shape states its own size (rather than inheriting one)."""
    frame = box.shape.text_frame if box.shape.kind == "shape" else None
    return frame is not None and any(run.size is not None for paragraph in frame.paragraphs
                                     for run in paragraph.runs if run.text.strip())


def _fill_of(box: _Box) -> str | None:
    fill = box.shape.effective_fill if box.shape.kind == "shape" else None
    color = getattr(fill, "color", None) if fill is not None else None
    if color is None:
        return None if fill is None else str(fill.kind)
    from .design import color_text

    return color_text(color)


def _text_sizes(touched_ids: set[str], boxes: list[_Box]) -> list[LayoutFact]:
    """A touched shape with a paragraph whose size differs from the size every like shape
    (same geometry, size and fill, as many paragraphs) has in that paragraph."""
    facts = []
    # Shapes with a fill or an outline: a plain text box's size goes with its role, which
    # its geometry does not tell.
    texted = [b for b in boxes if b.text and not b.placeholder and b.filled]
    for s in texted:
        if s.id not in touched_ids:
            continue
        mine = _sizes_of(s)
        if mine is None:
            continue
        others = []
        for b in texted:
            if b is s or not _like(s, b) or _fill_of(b) != _fill_of(s):
                continue
            theirs = _sizes_of(b)
            if theirs is not None and len(theirs) == len(mine):
                others.append((b, theirs))
        if len(others) < 2:
            continue
        if _explicit_size(s) and not any(_explicit_size(b) for b, _ in others):
            continue                                  # a size just set, the others inherited
        for index, size in enumerate(mine):
            column = [theirs[index] for _, theirs in others]
            if size is None or None in column or len(set(column)) != 1:
                continue                              # the others do not agree
            common = column[0]
            if abs(size - common) < 0.5:
                continue
            target = s.id if len(mine) == 1 else f"{s.id}/p{index}"
            facts.append(LayoutFact(
                "text_size", target, others[0][0].id, abs(size - common), {"size": common},
                f"{target}: text {size:g} pt where its {len(others)} like shapes have "
                f"{common:g} pt", [s.id] + [b.id for b, _ in others]))
            break
    return facts


# -- labels and anchors ------------------------------------------------------------------------


def _edge_gap(a: _Box, b: _Box) -> float:
    dx = max(b.x0 - a.x1, a.x0 - b.x1, 0)
    dy = max(b.y0 - a.y1, a.y0 - b.y1, 0)
    return (dx * dx + dy * dy) ** 0.5


def _label_size(box: _Box) -> tuple[float, float] | None:
    sizes = _sizes_of(box)
    if not sizes or sizes[0] is None:
        return None
    return sizes[0], round(box.h / _PT)


def _label_distances(touched_ids: set[str], boxes: list[_Box],
                     lines: list[tuple[float, float, float, float]]) -> list[LayoutFact]:
    """A touched label much farther from its anchor than the slide's other labels of its kind
    (one text size and height) are from theirs."""
    anchors = [b for b in boxes if not b.text and not b.placeholder and b.shape.kind == "shape"
               and b.filled and b.w <= ANCHOR_MAX and b.h <= ANCHOR_MAX]
    labels = [b for b in boxes if b.text and not b.placeholder and b.shape.kind == "shape"
              and not b.filled and b.h <= LABEL_MAX_HEIGHT and "\n" not in b.text]
    texts = [b for b in boxes if b.text and not b.placeholder]
    if len(anchors) < 3 or len(labels) < 3:
        return []
    # Labels are compared with labels of their own kind: one text size and height.
    kinds: dict[tuple, list[_Box]] = {}
    for label in labels:
        kind = _label_size(label)
        if kind is not None:
            kinds.setdefault(kind, []).append(label)
    facts = []
    for group in kinds.values():
        if len(group) >= 3 and any(label.id in touched_ids for label in group):
            facts += _far_labels(touched_ids, group, anchors, texts, lines)
    return facts


def _far_labels(touched_ids, labels, anchors, texts, lines) -> list[LayoutFact]:
    # Labels and markers matched one to one, nearest first: a marker's label is the nearest
    # label not already nearer to another marker.
    candidates = sorted(((_edge_gap(label, anchor), i, j) for i, label in enumerate(labels)
                         for j, anchor in enumerate(anchors)), key=lambda t: t[0])
    used_labels: set[int] = set()
    used_anchors: set[int] = set()
    pairs = []
    for gap, i, j in candidates:
        if i in used_labels or j in used_anchors:
            continue
        used_labels.add(i)
        used_anchors.add(j)
        pairs.append((labels[i], anchors[j], gap))
    if len(pairs) < 3:
        return []
    usual = median(gap for _, _, gap in pairs)
    far = [gap for _, _, gap in pairs if gap > usual * LABEL_FACTOR and gap - usual >= LABEL_EXTRA]
    if len(far) > max(1, LABEL_OUTLIERS * len(pairs)):
        return []                                     # not one set of labels beside markers
    facts = []
    for label, anchor, gap in pairs:
        if label.id not in touched_ids:
            continue
        if gap <= usual * LABEL_FACTOR or gap - usual < LABEL_EXTRA:
            continue
        if min(anchors, key=lambda a: _edge_gap(label, a)) is not anchor:
            continue                                  # its nearest marker has another label
        if any(_gap_to_point(label, end) < gap for end in lines):
            continue                                  # it may label the end of a line
        # Move the label straight toward its anchor until the gap is the usual one.
        lx, ly = label.edge("center"), label.edge("middle")
        ax, ay = anchor.edge("center"), anchor.edge("middle")
        dx, dy = ax - lx, ay - ly
        distance = (dx * dx + dy * dy) ** 0.5 or 1.0
        step = gap - usual
        nx, ny = label.x0 + dx / distance * step, label.y0 + dy / distance * step
        facts.append(LayoutFact(
            "label_distance", label.id, anchor.id, gap - usual, {"x": _pt(nx), "y": _pt(ny)},
            f"{label.id}: {_pt(gap)} pt from its marker {anchor.id}; the other "
            f"{len(pairs) - 1} labels like it are about {_pt(usual)} pt from theirs",
            [label.id, anchor.id]))
    return facts


def _line_ends(slide: "Slide") -> list[tuple[float, float]]:
    """The end points of every line and connector on the slide."""
    from .fit import _flatten

    out = []
    for shape in _flatten(slide.shapes):
        if shape.kind == "connector" or (shape.kind == "shape" and shape.preset in _LINES):
            bounds = shape.drawn_bounds
            if bounds is None:
                continue
            left, top, width, height = (float(v) for v in bounds)
            flip_h, flip_v = bool(shape.flip_h), bool(shape.flip_v)
            out.append((left + (width if flip_h else 0), top + (height if flip_v else 0)))
            out.append((left + (0 if flip_h else width), top + (0 if flip_v else height)))
    return out


def _gap_to_point(label: _Box, point: tuple[float, float]) -> float:
    dx = max(point[0] - label.x1, label.x0 - point[0], 0)
    dy = max(point[1] - label.y1, label.y0 - point[1], 0)
    return (dx * dx + dy * dy) ** 0.5


# -- all of them -------------------------------------------------------------------------------


def layout_facts(slide: "Slide", touched: Iterable[str], *, limit: int = 5) -> list[LayoutFact]:
    """The layout facts about the shapes ``touched`` (addresses) on ``slide``, largest first,
    at most ``limit``.  A touched group counts its members as touched."""
    boxes = _boxes(slide)
    wanted = set(touched)
    touched_ids = {b.id for b in boxes if b.id in wanted or any(
        g.id in wanted for g in _ancestors(b.shape))}
    if not touched_ids:
        return []
    touched_boxes = [b for b in boxes if b.id in touched_ids]
    facts = (_label_distances(touched_ids, boxes, _line_ends(slide))
             + _text_sizes(touched_ids, boxes)
             + _uneven_gaps(touched_ids, boxes) + _near_alignment(touched_boxes, boxes))
    # One fact per shape (the first kind found, in the order above), largest first.
    out: list[LayoutFact] = []
    for fact in facts:
        if all(f.shape != fact.shape for f in out):
            out.append(fact)
    order = {"label_distance": 0, "text_size": 1, "uneven_gap": 2, "near_alignment": 3}
    out.sort(key=lambda f: (order[f.kind], -f.amount))
    return out[:limit]


def _ancestors(shape: "Shape"):
    group = shape.parent_group
    while group is not None:
        yield group
        group = group.parent_group


__all__ = ["LayoutFact", "layout_facts"]
