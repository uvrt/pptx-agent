"""PowerPoint's Align and Distribute, and the shared moves under the layout helpers (LP1).

Every function takes shapes (a group moves as one) and works on what is **drawn**: a
rotated shape by its turned box, a connector by its route (:attr:`Shape.drawn_bounds`),
which is also what PowerPoint aligns.  Lengths are EMU.  Each call is one undo step.

* :func:`align` -- line shapes up on an edge or centre: of the selection's bounds, the
  slide, a box (the content area), or the first shape (a key object);
* :func:`distribute` -- even gaps along one axis: between the outermost shapes, across
  the slide or a box, or a fixed gap from the first;
* :func:`place` -- put a shape's drawn box at a position, optionally at a new size.

For example::

    from pptx_agent.edit.arrange import align, distribute

    align([chevron, column], "center", to="first")      # centre the column under it
    distribute(boxes, "horizontal")                     # equal gaps, outer two stay put
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterable, Sequence

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape

#: The edges and centres :func:`align` lines up.
EDGES = ("left", "center", "right", "top", "middle", "bottom")
#: What a call aligns or distributes relative to.
RELATIVE = ("selection", "slide", "box", "first")

Box = tuple[int, int, int, int]


def drawn_box(shape: "Shape") -> Box:
    """``(left, top, width, height)`` of what is drawn, slide EMU; a ``ValueError`` names a
    shape with no position."""
    bounds = shape.drawn_bounds
    if bounds is None:
        raise ValueError(f"{shape.id} has no position to arrange")
    return bounds


def outermost(shapes: Sequence["Shape"]) -> list["Shape"]:
    """``shapes`` without any that sit inside another one listed (a group's member moves
    with its group), in the order given, each once."""
    seen: set[str] = set()
    ids = {shape.id for shape in shapes}
    out = []
    for shape in shapes:
        if shape.id in seen:
            continue
        seen.add(shape.id)
        parent = shape.parent_group
        nested = False
        while parent is not None:
            if parent.id in ids:
                nested = True
                break
            parent = parent.parent_group
        if not nested:
            out.append(shape)
    return out


def move_to(shape: "Shape", left: float | None = None, top: float | None = None) -> bool:
    """Move ``shape`` so its drawn box starts at ``left``/``top`` (slide EMU; ``None``
    keeps that coordinate).  True when it moved."""
    x, y, _, _ = drawn_box(shape)
    dx = 0 if left is None else round(left - x)
    dy = 0 if top is None else round(top - y)
    if dx or dy:
        shape.move_by(dx, dy, space="slide")
        return True
    return False


def resize(shape: "Shape", width: float | None = None, height: float | None = None) -> None:
    """Give ``shape`` a drawn ``width``/``height`` (slide EMU): its frame's, or for a shape
    turned a quarter, the other side's.  Other turns are refused (a turned box's size
    depends on both sides)."""
    if width is None and height is None:
        return
    if shape.kind == "connector":
        raise ValueError(f"{shape.id} is a connector: its ends set its size")
    turn = round(shape.rotation or 0) % 180
    if turn not in (0, 90) or abs((shape.rotation or 0) - round(shape.rotation or 0)) > 1e-6:
        raise ValueError(f"{shape.id} is turned {shape.rotation:g} degrees: only an unturned "
                         "shape, or one turned a quarter, can be resized to a box")
    if turn == 90:
        width, height = height, width
    scale_x, scale_y = shape._parent_scale()
    with shape._slide.document.batch():
        if width is not None:
            shape.width = max(1, round(width / scale_x))
        if height is not None:
            shape.height = max(1, round(height / scale_y))


def place(shape: "Shape", left: float, top: float, width: float | None = None,
          height: float | None = None) -> None:
    """Put ``shape``'s drawn box at ``(left, top)``, at ``width``/``height`` when given."""
    with shape._slide.document.batch():
        resize(shape, width, height)
        move_to(shape, left, top)


def selection_box(shapes: Iterable["Shape"]) -> Box:
    """The box around the drawn boxes of ``shapes``."""
    boxes = [drawn_box(shape) for shape in shapes]
    if not boxes:
        raise ValueError("no shapes")
    left = min(b[0] for b in boxes)
    top = min(b[1] for b in boxes)
    right = max(b[0] + b[2] for b in boxes)
    bottom = max(b[1] + b[3] for b in boxes)
    return left, top, right - left, bottom - top


def _reference(shapes: Sequence["Shape"], to: str, box: Box | None) -> Box:
    if to == "selection":
        return selection_box(shapes)
    if to == "slide":
        width, height = shapes[0]._slide.document.slide_size
        return 0, 0, width, height
    if to == "box":
        if box is None:
            raise ValueError("to='box' needs a box")
        return box
    if to == "first":
        return drawn_box(shapes[0])
    raise ValueError(f"to must be one of {', '.join(RELATIVE)}, not {to!r}")


def align(shapes: Sequence["Shape"], edge: str, *, to: str = "selection",
          box: Box | None = None) -> list["Shape"]:
    """Line ``shapes`` up on ``edge`` (left, center, right, top, middle, bottom) of the
    selection's bounds, the slide, ``box`` (``to="box"``), or the first shape
    (``to="first"``, which stays put).  Returns the shapes that moved.

    For example::

        align(labels, "left")                    # PowerPoint's Align Left
        align([chevron, column], "center", to="first")
    """
    if edge not in EDGES:
        raise ValueError(f"edge must be one of {', '.join(EDGES)}, not {edge!r}")
    shapes = outermost(shapes)
    if not shapes:
        return []
    if to == "selection" and len(shapes) < 2:
        raise ValueError("aligning to the selection needs two shapes or more; give "
                         "to='slide' or to='box' to align one")
    ref = _reference(shapes, to, box)
    moved = []
    movers = shapes[1:] if to == "first" else shapes
    with shapes[0]._slide.document.batch():
        for shape in movers:
            x, y, w, h = drawn_box(shape)
            left = top = None
            if edge == "left":
                left = ref[0]
            elif edge == "center":
                left = ref[0] + (ref[2] - w) / 2
            elif edge == "right":
                left = ref[0] + ref[2] - w
            elif edge == "top":
                top = ref[1]
            elif edge == "middle":
                top = ref[1] + (ref[3] - h) / 2
            else:
                top = ref[1] + ref[3] - h
            if move_to(shape, left, top):
                moved.append(shape)
    return moved


def distribute(shapes: Sequence["Shape"], axis: str, *, to: str = "selection",
               box: Box | None = None, gap: float | None = None) -> list["Shape"]:
    """Space ``shapes`` evenly along ``axis`` (horizontal or vertical), in the order they
    stand (left to right, top to bottom), as PowerPoint does.

    * ``to="selection"``: the outermost two stay put and the gaps between all are equal;
    * ``to="slide"``/``"box"``: the first starts at the slide's (box's) near edge, the last
      ends at its far edge, equal gaps between;
    * ``gap`` (EMU): a fixed gap instead, from the first shape where it stands (or from
      the slide's or box's near edge).

    Equal gaps that would be negative (the shapes are wider than the room) are refused.
    Returns the shapes that moved.
    """
    if axis not in ("horizontal", "vertical"):
        raise ValueError(f"axis must be horizontal or vertical, not {axis!r}")
    if to == "first":
        raise ValueError("distribute spreads over the selection, the slide or a box")
    shapes = outermost(shapes)
    if len(shapes) < 2 or (len(shapes) < 3 and to == "selection" and gap is None):
        raise ValueError("distributing needs three shapes or more (two with to='slide', "
                         "to='box' or a gap)")
    i = 0 if axis == "horizontal" else 1
    boxes = {shape.id: drawn_box(shape) for shape in shapes}
    ordered = sorted(shapes, key=lambda s: (boxes[s.id][i], boxes[s.id][1 - i]))
    sizes = [boxes[s.id][i + 2] for s in ordered]
    if to == "selection":
        start = boxes[ordered[0].id][i]
        last = boxes[ordered[-1].id]
        end = last[i] + last[i + 2]
    else:
        ref = _reference(shapes, to, box)
        start, end = ref[i], ref[i] + ref[i + 2]
    if gap is None:
        room = end - start - sum(sizes)
        step = room / (len(ordered) - 1)
        if step < 0:
            raise ValueError(f"the shapes are {sum(sizes) / 12700:.2f} pt long and the room "
                             f"is {(end - start) / 12700:.2f} pt: they cannot be spread "
                             "without overlapping")
    else:
        step = gap
    moved = []
    position = float(start)
    with shapes[0]._slide.document.batch():
        for shape, size in zip(ordered, sizes):
            if axis == "horizontal":
                changed = move_to(shape, left=position)
            else:
                changed = move_to(shape, top=position)
            if changed:
                moved.append(shape)
            position += size + step
    return moved


__all__ = ["EDGES", "RELATIVE", "align", "distribute", "drawn_box", "move_to", "outermost",
           "place", "resize", "selection_box"]
