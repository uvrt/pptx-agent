"""Stable shape addressing.

An agent holds a reference to a shape across several edits -- "move the chart right, then make
its title bold".  Addressing by position cannot support that: pptx-svg addresses shapes as
``(slideIdx, shapeIdx)`` and its own documentation concedes that every reorder invalidates the
reference, so each mutator has to hand back a fresh index.  Ids here are designed so a
reference taken before an edit still resolves afterwards.

**The slide half** is ``p:sldId/@id`` from ``sldIdLst`` -- a deck-unique integer that is
unaffected by reordering, unlike the slide's position.

**The shape half** comes from the first of these that is available:

1. a pptx-agent stamp in ``p:cNvPr/a:extLst``, written the first time a shape is mutated;
2. ``a16:creationId``, PowerPoint's own durable per-shape id, when the authoring application
   wrote one;
3. ``p:cNvPr/@id``.

Case 3 needs a caveat that the obvious implementation gets wrong: **``cNvPr@id`` is not
reliably unique within a slide.**  In ``real-financial-report.pptx`` slide 2 a ``p:sp`` named
"Text 1" and a ``p:graphicFrame`` named "Table 0" both carry ``id="3"`` (slides 3 and 4 are
likewise affected).  When a part contains duplicates, the shape-tree index path is appended --
``256.3#0-2`` -- which is unique by construction.

That path component is the only volatile part of an id, and stamping exists to retire it: the
first mutation of a shape freezes whatever id it currently has into ``extLst``, so later
structural edits, and PowerPoint's own id renumbering, cannot move it.  Office preserves
extension elements it does not recognise, which is what makes the stamp durable.
"""

from __future__ import annotations

from dataclasses import dataclass

from ooxml_edit.stamp import ExtensionStamp

from ..oxml.xml import (
    CREATION_ID_EXT_URI,
    PA_ID_EXT_URI,
    Element,
    find,
    local_name,
    qn,
    walk_shape_tree,
)

#: Shape tag -> the non-visual-properties container holding its ``p:cNvPr``.
_NV_CONTAINER: dict[str, str] = {
    "sp": "p:nvSpPr",
    "pic": "p:nvPicPr",
    "cxnSp": "p:nvCxnSpPr",
    "grpSp": "p:nvGrpSpPr",
    "graphicFrame": "p:nvGraphicFramePr",
}


@dataclass(frozen=True, order=True)
class ShapeId:
    """A deck-wide shape address, rendered as ``"<sldId>.<local>"``.

    For example::

        ShapeId.parse("257.3#5")                       # ShapeId(slide_id=257, local='3#5')
    """

    slide_id: int
    local: str

    def __str__(self) -> str:
        return f"{self.slide_id}.{self.local}"

    @classmethod
    def parse(cls, text: str | "ShapeId") -> "ShapeId":
        """A shape id from its text, ``<sldId>.<local>``.

        For example::

            ShapeId.parse("256.29").slide_id   # 256
        """
        if isinstance(text, ShapeId):
            return text
        slide, separator, local = text.partition(".")
        if not separator or not local:
            raise ValueError(f"malformed shape id {text!r}; expected '<sldId>.<shape>'")
        try:
            return cls(int(slide), local)
        except ValueError:
            raise ValueError(f"malformed shape id {text!r}; slide part must be an integer") from None


def cnv_pr(shape: Element) -> Element | None:
    """The ``p:cNvPr`` of a shape element, whichever flavour of shape it is."""
    container = _NV_CONTAINER.get(local_name(shape))
    if container is None:
        return None
    return find(shape, f"{container}/p:cNvPr")


class SlideShapeIndex:
    """Every shape in one slide part, addressable by id.

    Built once per slide-part parse.  Rebuild after a structural change -- adding, removing or
    reordering shapes -- since index paths shift.
    """

    def __init__(self, slide_id: int, sp_tree: Element) -> None:
        self.slide_id = slide_id
        self._by_id: dict[str, Element] = {}
        self._for_element: dict[int, str] = {}

        shapes = list(walk_shape_tree(sp_tree))

        # Find duplicated cNvPr ids up front: they decide whether the cheap id is usable.
        seen: dict[str, int] = {}
        for shape, _ in shapes:
            properties = cnv_pr(shape)
            raw = properties.get("id") if properties is not None else None
            if raw is not None:
                seen[raw] = seen.get(raw, 0) + 1
        duplicated = {raw for raw, count in seen.items() if count > 1}

        for shape, path in shapes:
            local = _local_id(shape, path, duplicated)
            if local is None:
                continue
            # Defensive: a stamp or creationId could in principle collide with a derived id.
            if local in self._by_id:
                local = f"{local}#{'-'.join(str(step) for step in path)}"
            self._by_id[local] = shape
            self._for_element[id(shape)] = local

    def __len__(self) -> int:
        return len(self._by_id)

    def __iter__(self):
        for local, shape in self._by_id.items():
            yield ShapeId(self.slide_id, local), shape

    def get(self, shape_id: ShapeId | str) -> Element | None:
        parsed = ShapeId.parse(shape_id)
        if parsed.slide_id != self.slide_id:
            return None
        return self._by_id.get(parsed.local)

    def id_of(self, shape: Element) -> ShapeId | None:
        local = self._for_element.get(id(shape))
        return None if local is None else ShapeId(self.slide_id, local)


def _local_id(shape: Element, path: tuple[int, ...], duplicated: set[str]) -> str | None:
    properties = cnv_pr(shape)
    if properties is None:
        return None

    stamped = read_stamp(properties)
    if stamped:
        return stamped

    creation = _creation_id(properties)
    if creation:
        return f"c{creation}"

    raw = properties.get("id")
    if raw is None:
        return "#" + "-".join(str(step) for step in path)
    if raw in duplicated:
        return f"{raw}#{'-'.join(str(step) for step in path)}"
    return raw


def _creation_id(properties: Element) -> str | None:
    """``a16:creationId/@id`` if the authoring application wrote one."""
    for extension in _extensions(properties):
        if extension.get("uri") != CREATION_ID_EXT_URI:
            continue
        node = extension.find(qn("a16:creationId"))
        value = node.get("id") if node is not None else None
        if value:
            return value.strip("{}")
    return None


#: Where a shape's durable id lives: ``p:cNvPr/a:extLst/a:ext[@uri]/pa:id/@val``.
SHAPE_STAMP = ExtensionStamp(ext_list="a:extLst", ext="a:ext", uri=PA_ID_EXT_URI, value="pa:id")


def read_stamp(properties: Element) -> str | None:
    """The durable id pptx-agent wrote into ``p:cNvPr``, if any."""
    return SHAPE_STAMP.read(properties)


def write_stamp(properties: Element, local: str) -> None:
    """Freeze ``local`` as the shape's permanent id.

    Called on a shape's first mutation, never on open -- so a deck that is only read stays
    byte-identical, and the stamp records the id the shape *already had* rather than minting a
    new one.  Re-stamping is a no-op.
    """
    SHAPE_STAMP.write(properties, local)


def _extensions(properties: Element) -> list[Element]:
    ext_list = properties.find(qn("a:extLst"))
    return [] if ext_list is None else list(ext_list.findall(qn("a:ext")))
