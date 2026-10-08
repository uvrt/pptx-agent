"""Slide structure: the presentation's slide list, and the parts behind each slide.

Everything here changes the package rather than one part, and the rule that makes it safe is
the one :mod:`ooxml_edit.opc` enforces for parts: nothing is removed until it is proved
unreferenced across the whole package (:meth:`~ooxml_edit.opc.OpcPackage.reap`).

A slide is listed in up to three places in ``presentation.xml``, and every one must be kept
consistent or PowerPoint offers to repair the file:

* ``p:sldIdLst/p:sldId`` -- the order, and the ``sldId`` this library's shape ids start with;
* ``p14:sectionLst`` -- sections partition the slide list *in order*, by ``sldId``;
* ``p:custShowLst`` -- custom shows list slides by relationship id.
"""

from __future__ import annotations

import zlib
from dataclasses import dataclass, field
from typing import NamedTuple

from ..oxml.package import CT_SLIDE, PresentationPackage, REL_SLIDE, REL_SLIDE_LAYOUT
from ..oxml.xml import (
    NAMESPACES,
    Element,
    find,
    findall,
    local_name,
    make,
    qn,
    remove,
    serialize,
    subelement,
)

#: ``p:ext`` URI of PowerPoint 2010's sections.
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
#: ``p:ext`` URI of a slide's ``p14:creationId``.
SLIDE_CREATION_ID_EXT_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
#: The lowest ``sldId`` the schema allows.
MIN_SLIDE_ID = 256
MAX_SLIDE_ID = 2147483647
#: Placeholders PowerPoint leaves on the layout when it makes a slide: they are switched on
#: per deck under Insert > Header & Footer, not copied.
_NOT_COPIED = {"dt", "ftr", "sldNum", "hdr"}
_HYPERLINK_TAGS = {"hlinkClick", "hlinkHover", "hlinkMouseOver"}


class Placeholder(NamedTuple):
    """What makes a shape a placeholder: its ``type`` and ``idx`` (``p:ph``).

    ``type`` is ``"title"``, ``"ctrTitle"``, ``"subTitle"``, ``"body"``, ``"pic"``,
    ``"tbl"``, ``"chart"``, ``"dt"``, ``"ftr"``, ``"sldNum"``... or ``None``, which the file
    format reads as ``"obj"`` (a content placeholder that takes text, a table, a chart or a
    picture).  ``idx`` is what a slide placeholder matches its layout placeholder by;
    ``None`` for the title, which matches by type.  A plain tuple too:
    ``shape.placeholder == ("title", None)`` holds.

    For example::

        kind, idx = shape.placeholder                # Placeholder(type='body', idx=1)
        shape.placeholder.type in ("title", "ctrTitle")
    """

    type: str | None
    idx: int | None


class _Frame:
    """``left``/``top``/``width``/``height`` read from ``bounds``, as a :class:`Shape` has
    them -- and a shape's ``bounds`` is the same tuple."""

    bounds: tuple[int, int, int, int] | None

    @property
    def left(self) -> int | None:
        """``bounds[0]``, EMU; ``None`` when it states no frame.

        For example::

            ph.left, ph.top, ph.width, ph.height
        """
        return None if self.bounds is None else self.bounds[0]

    @property
    def top(self) -> int | None:
        """``bounds[1]``, EMU.

        For example::

            ph.top
        """
        return None if self.bounds is None else self.bounds[1]

    @property
    def width(self) -> int | None:
        """``bounds[2]``, EMU.

        For example::

            ph.width
        """
        return None if self.bounds is None else self.bounds[2]

    @property
    def height(self) -> int | None:
        """``bounds[3]``, EMU.

        For example::

            ph.height
        """
        return None if self.bounds is None else self.bounds[3]


@dataclass(frozen=True)
class LayoutPlaceholder(_Frame):
    """A placeholder a layout offers -- what a slide made from it gets, and where.

    ``placeholder`` is its :class:`Placeholder` (``type``, ``idx``), ``name`` its name on
    the layout, ``bounds`` ``(left, top, width, height)`` in EMU -- its own, or inherited
    from the master when the layout states none; ``left``, ``top``, ``width`` and
    ``height`` too, one by one, as a :class:`Shape` has them -- and ``sizes`` the default
    text size, in points, of each of the nine outline levels, through the master's text
    styles.

    For example::

        for ph in deck.layout("Title Only").placeholders:
            print(ph.type, ph.idx, ph.name, ph.bounds, ph.sizes[0])
    """

    placeholder: Placeholder
    name: str
    bounds: tuple[int, int, int, int] | None
    sizes: tuple[float, ...]

    @property
    def type(self) -> str | None:
        """The placeholder type (``None`` is ``"obj"``).

        For example::

            [ph.type for ph in layout.placeholders]      # ['title', 'body', ...]
        """
        return self.placeholder.type

    @property
    def idx(self) -> int | None:
        """The placeholder index a slide's placeholder matches by.

        For example::

            [ph.idx for ph in layout.placeholders]       # [None, 1, 2]
        """
        return self.placeholder.idx

    @property
    def bottom(self) -> int | None:
        """Where it ends, EMU from the top: ``top + height``.

        For example::

            below_title = deck.layout("Title Only").placeholders[0].bottom
        """
        return None if self.bounds is None else self.bounds[1] + self.bounds[3]


@dataclass(frozen=True)
class LayoutShape(_Frame):
    """A shape on a layout, read-only: a layout is edited in PowerPoint's slide master view,
    not here.  ``id`` is its ``cNvPr@id``; ``kind`` as :attr:`Shape.kind`; ``bounds`` its
    own frame, if it states one, and ``left``/``top``/``width``/``height`` the same one
    by one -- as a :class:`Shape` has both.

    For example::

        [(shape.kind, shape.name, shape.placeholder) for shape in layout.shapes]
    """

    id: str
    kind: str
    name: str
    placeholder: Placeholder | None
    bounds: tuple[int, int, int, int] | None
    text: str


@dataclass(frozen=True)
class Layout:
    """A slide layout a new slide can be made from, and what it offers.

    For example::

        layout = deck.layout("Title and Content")
        print([(ph.type, ph.idx, ph.bounds) for ph in layout.placeholders])
        deck.add_slide(layout)
    """

    name: str
    part_path: str
    master_part: str
    _package: object = field(default=None, compare=False, hash=False, repr=False)

    @property
    def placeholders(self) -> list[LayoutPlaceholder]:
        """The placeholders a slide made from this layout gets -- and the date, footer,
        slide number and header ones it leaves on the layout, as PowerPoint does -- in the
        layout's order, each with its type, idx, name, bounds and default text sizes.

        For example::

            title = next(ph for ph in layout.placeholders if ph.type in ("title", "ctrTitle"))
            top_of_content = title.bottom
        """
        from . import inherit

        package = self._require_package()
        out = []
        for shape in self._tree():
            ph = inherit.placeholder_of(shape)
            if ph is None:
                continue
            kind = ph.get("type")
            idx = ph.get("idx")
            placeholder = Placeholder(kind, int(idx) if idx and idx.isdigit() else None)
            bounds = inherit.frame_of(shape)
            if bounds is None:
                bounds = inherit.frame_of(inherit.master_placeholder(
                    package, self.master_part, inherit.master_kind(kind)))
            chain = inherit.list_styles(
                package, self.master_part, None, shape, kind, is_placeholder=True)
            sizes = tuple(inherit.level_size(chain, level) for level in range(9))
            name = _name_of(shape)
            out.append(LayoutPlaceholder(placeholder, name, bounds, sizes))
        return out

    @property
    def shapes(self) -> list[LayoutShape]:
        """Every shape on the layout, back to front, read-only (:class:`LayoutShape`):
        placeholders and the layout's own decoration.

        For example::

            [shape.name for shape in deck.layout("Title Slide").shapes]
        """
        from . import inherit

        out = []
        for shape in self._tree():
            ph = inherit.placeholder_of(shape)
            placeholder = None
            if ph is not None:
                idx = ph.get("idx")
                placeholder = Placeholder(ph.get("type"),
                                          int(idx) if idx and idx.isdigit() else None)
            properties = shape.find(f".//{qn('p:cNvPr')}")
            texts = ["".join(t.text or "" for t in paragraph.iter(qn("a:t")))
                     for paragraph in shape.findall(f"{qn('p:txBody')}/{qn('a:p')}")]
            out.append(LayoutShape(
                id=(properties.get("id") if properties is not None else None) or "",
                kind=_KINDS.get(local_name(shape), "unknown"), name=_name_of(shape),
                placeholder=placeholder, bounds=inherit.frame_of(shape),
                text="\n".join(texts)))
        return out

    def _require_package(self):
        if self._package is None:
            raise ValueError("this Layout was made by hand; take it from deck.layouts")
        return self._package

    def _tree(self) -> list[Element]:
        root = self._require_package().tree(self.part_path)
        tree = find(root, "p:cSld/p:spTree") if root is not None else None
        return [] if tree is None else [shape for shape in tree if local_name(shape) in _KINDS]

    def __repr__(self) -> str:
        return f"<Layout {self.name!r} {self.part_path}>"


_KINDS = {"sp": "shape", "pic": "picture", "cxnSp": "connector", "grpSp": "group",
          "graphicFrame": "graphic_frame"}


def _name_of(shape: Element) -> str:
    properties = shape.find(f".//{qn('p:cNvPr')}")
    return (properties.get("name") if properties is not None else None) or ""


# -- reading -----------------------------------------------------------------------------


def layouts(package: PresentationPackage) -> list[Layout]:
    """Every layout, master by master, in each master's own order."""
    presentation = package.presentation_part()
    root = package.tree(presentation)
    found: list[Layout] = []
    for master_id in findall(root, "p:sldMasterIdLst/p:sldMasterId"):
        master = package.related_part(presentation, master_id.get(qn("r:id")))
        master_root = package.tree(master) if master else None
        if master_root is None:
            continue
        for layout_id in findall(master_root, "p:sldLayoutIdLst/p:sldLayoutId"):
            part = package.related_part(master, layout_id.get(qn("r:id")))
            layout_root = package.tree(part) if part else None
            if layout_root is None:
                continue
            common = find(layout_root, "p:cSld")
            name = (common.get("name") if common is not None else None) or ""
            found.append(Layout(name, part, master, package))
    return found


def slide_layout(package: PresentationPackage, slide_part: str) -> str | None:
    related = package.related_parts_of_type(slide_part, REL_SLIDE_LAYOUT)
    return related[0] if related else None


# -- the slide list ----------------------------------------------------------------------


def _presentation(package: PresentationPackage) -> tuple[str, Element]:
    path = package.presentation_part()
    root = package.tree(path)
    if root is None:
        raise ValueError("the presentation part is missing")
    return path, root


def _entries(root: Element) -> list[Element]:
    return findall(root, "p:sldIdLst/p:sldId")


def _entry(root: Element, slide_id: int) -> Element:
    for node in _entries(root):
        if node.get("id") == str(slide_id):
            return node
    raise KeyError(f"no slide with id {slide_id}")


def new_slide_id(root: Element) -> int:
    """One more than the largest ``sldId`` in use (sections included), and at least 256."""
    used = [int(node.get("id")) for node in _entries(root) if (node.get("id") or "").isdigit()]
    used += [int(node.get("id")) for section in _section_lists(root)
             for node in section if (node.get("id") or "").isdigit()]
    candidate = max(used + [MIN_SLIDE_ID - 1]) + 1
    if candidate > MAX_SLIDE_ID:
        taken = set(used)
        candidate = next(i for i in range(MIN_SLIDE_ID, MAX_SLIDE_ID + 1) if i not in taken)
    return candidate


def insert_entry(package: PresentationPackage, slide_part: str, index: int | None) -> int:
    """List ``slide_part`` at ``index`` (``None``: last); returns its new ``sldId``."""
    path, root = _presentation(package)
    rel_id = package.add_relationship(path, REL_SLIDE, slide_part)
    slide_id = new_slide_id(root)
    node = make("p:sldId", id=str(slide_id), r__id=rel_id)
    _place(root, node, index)
    package.mark_dirty(path)
    return slide_id


def move_entry(package: PresentationPackage, slide_id: int, index: int) -> None:
    path, root = _presentation(package)
    node = _entry(root, slide_id)
    remove(node)
    node.tail = None
    _sections_remove(root, slide_id)
    _place(root, node, index)
    package.mark_dirty(path)


def remove_entry(package: PresentationPackage, slide_id: int) -> str:
    """Take a slide out of the slide list, its section and every custom show.

    Returns the presentation's relationship id for it, which is now unreferenced.
    """
    path, root = _presentation(package)
    node = _entry(root, slide_id)
    rel_id = node.get(qn("r:id")) or ""
    remove(node)
    _sections_remove(root, slide_id)
    for show_slide in list(root.iter(qn("p:sld"))):
        if show_slide.get(qn("r:id")) == rel_id:
            remove(show_slide)
    package.mark_dirty(path)
    return rel_id


def _place(root: Element, node: Element, index: int | None) -> None:
    entries = _entries(root)
    position = len(entries) if index is None else index
    if not 0 <= position <= len(entries):
        raise IndexError(f"slide index {index} out of range 0..{len(entries)}")
    if position < len(entries):
        entries[position].addprevious(node)
    elif entries:
        entries[-1].addnext(node)
    else:
        subelement(root, "p:sldIdLst").insert(0, node)
    predecessor = int(entries[position - 1].get("id")) if position > 0 else None
    _sections_insert(root, int(node.get("id")), predecessor)


# -- sections ----------------------------------------------------------------------------


def _section_lists(root: Element) -> list[Element]:
    lists = []
    for extension in findall(root, "p:extLst/p:ext"):
        if extension.get("uri") != SECTION_EXT_URI:
            continue
        for section in extension.iter(qn("p14:section")):
            ids = section.find(qn("p14:sldIdLst"))
            if ids is None:
                ids = subelement(section, "p14:sldIdLst")
            lists.append(ids)
    return lists


def sections(root: Element) -> list[tuple[str, list[int]]]:
    """``(section name, [sldId...])`` in order -- for tests and diagnostics."""
    result = []
    for ids in _section_lists(root):
        section = ids.getparent()
        result.append((section.get("name") or "",
                       [int(node.get("id")) for node in ids if node.get("id")]))
    return result


def _sections_remove(root: Element, slide_id: int) -> None:
    for ids in _section_lists(root):
        for node in list(ids):
            if node.get("id") == str(slide_id):
                remove(node)


def _sections_insert(root: Element, slide_id: int, predecessor: int | None) -> None:
    """Keep sections a partition of the slide list: the slide joins its predecessor's section
    right after it, or opens the first section when it is first."""
    lists = _section_lists(root)
    if not lists:
        return
    node = make("p14:sldId", id=str(slide_id))
    if predecessor is not None:
        for ids in lists:
            for existing in ids:
                if existing.get("id") == str(predecessor):
                    existing.addnext(node)
                    return
    lists[0].insert(0, node)


# -- slide parts -------------------------------------------------------------------------


def unlink_references(package: PresentationPackage, slide_part: str, keep: set[str]) -> None:
    """Remove every hyperlink in the package that jumps to ``slide_part``.

    PowerPoint does the same when the target slide is deleted: the link goes, the text or
    shape that carried it stays.  ``keep`` are parts that will go with the slide (its notes)
    and need no unlinking.  A reference this library cannot remove safely is refused rather
    than left dangling.
    """
    r_namespace = "{%s}" % NAMESPACES["r"]
    for source, relationship in package.relationship_sources(slide_part):
        if source in keep or relationship.is_external:
            continue
        root = package.tree(source)
        if root is not None:
            holders = [node for node in root.iter()
                       if any(name.startswith(r_namespace) and value == relationship.id
                              for name, value in node.attrib.items())]
            for node in holders:
                if local_name(node) not in _HYPERLINK_TAGS:
                    raise ValueError(
                        f"{source} refers to {slide_part} through <{local_name(node)}>, which "
                        "cannot be unlinked safely")
            for node in holders:
                remove(node)
            if holders:
                package.mark_dirty(source)
        package.remove_relationship(source, relationship.id)


def reachable_from(package: PresentationPackage, part: str) -> set[str]:
    seen = {part}
    pending = [part]
    while pending:
        for rel in package.relationships(pending.pop()).values():
            if rel.target_part is not None and rel.target_part not in seen \
                    and package.has_part(rel.target_part):
                seen.add(rel.target_part)
                pending.append(rel.target_part)
    return seen


def renew_creation_id(package: PresentationPackage, slide_part: str, salt: int) -> None:
    """Give a copied slide a ``p14:creationId`` of its own, if the original carried one."""
    root = package.tree(slide_part)
    for extension in findall(root, "p:extLst/p:ext"):
        if extension.get("uri") != SLIDE_CREATION_ID_EXT_URI:
            continue
        node = extension.find(qn("p14:creationId"))
        if node is None:
            continue
        used = set()
        for _, path in package.slide_parts():
            other = package.tree(path)
            for existing in other.iter(qn("p14:creationId")) if other is not None else ():
                used.add(existing.get("val"))
        value = zlib.crc32(f"{slide_part}:{salt}".encode()) or 1
        while str(value) in used:
            value = (value * 2654435761 + 1) % 4294967296 or 1
        node.set("val", str(value))
        package.mark_dirty(slide_part)


def slide_from_layout(package: PresentationPackage, layout_part: str) -> bytes:
    """A new slide's XML: the layout's placeholders, empty, inheriting everything, in the
    deck's default language."""
    layout = package.tree(layout_part)
    tree = find(layout, "p:cSld/p:spTree")
    if tree is None:
        raise ValueError(f"{layout_part} has no shape tree")
    slide = make_root("p:sld", {prefix: NAMESPACES[prefix] for prefix in ("a", "r", "p")})
    shapes = subelement(subelement(slide, "p:cSld"), "p:spTree")
    group_nv = subelement(shapes, "p:nvGrpSpPr")
    group_nv.append(make("p:cNvPr", id="1", name=""))
    group_nv.append(make("p:cNvGrpSpPr"))
    group_nv.append(make("p:nvPr"))
    group = subelement(shapes, "p:grpSpPr")
    xfrm = subelement(group, "a:xfrm")
    for tag, attributes in (("a:off", {"x": "0", "y": "0"}), ("a:ext", {"cx": "0", "cy": "0"}),
                            ("a:chOff", {"x": "0", "y": "0"}),
                            ("a:chExt", {"cx": "0", "cy": "0"})):
        xfrm.append(make(tag, **attributes))

    from .authoring import default_language

    language = default_language(package)
    next_id = 2
    for shape in tree:
        if local_name(shape) != "sp":
            continue
        placeholder = find(shape, "p:nvSpPr/p:nvPr/p:ph")
        if placeholder is None or placeholder.get("type") in _NOT_COPIED:
            continue
        source = find(shape, "p:nvSpPr/p:cNvPr")
        sp = make("p:sp")
        nv = make("p:nvSpPr")
        nv.append(make("p:cNvPr", id=str(next_id),
                       name=(source.get("name") if source is not None else None)
                       or f"Placeholder {next_id - 1}"))
        locks = make("p:cNvSpPr")
        locks.append(make("a:spLocks", noGrp="1"))
        nv.append(locks)
        nv_pr = make("p:nvPr")
        copied = make("p:ph")
        for attribute in ("type", "orient", "sz", "idx"):
            if placeholder.get(attribute) is not None:
                copied.set(attribute, placeholder.get(attribute))
        nv_pr.append(copied)
        nv.append(nv_pr)
        sp.append(nv)
        sp.append(make("p:spPr"))
        body = make("p:txBody")
        body.append(make("a:bodyPr"))
        body.append(make("a:lstStyle"))
        paragraph = make("a:p")
        paragraph.append(make("a:endParaRPr", lang=language))
        body.append(paragraph)
        sp.append(body)
        shapes.append(sp)
        next_id += 1

    color_map = subelement(slide, "p:clrMapOvr")
    color_map.append(make("a:masterClrMapping"))
    return serialize(slide)


def make_root(tag: str, namespaces: dict[str, str]) -> Element:
    from lxml import etree

    return etree.Element(qn(tag), nsmap=namespaces)


__all__ = [
    "CT_SLIDE",
    "Layout",
    "LayoutPlaceholder",
    "LayoutShape",
    "Placeholder",
    "insert_entry",
    "layouts",
    "move_entry",
    "remove_entry",
    "renew_creation_id",
    "sections",
    "slide_from_layout",
    "slide_layout",
    "unlink_references",
]
