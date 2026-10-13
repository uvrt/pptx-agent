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
    """A read-only summary of a shape on a layout: ``id`` (its ``cNvPr@id``), ``kind``,
    ``name``, ``placeholder``, ``bounds`` and ``text``.  No longer what
    :attr:`Layout.shapes` returns -- that is the editable :class:`~pptx_agent.Shape` now
    -- but kept for code that builds one.

    For example::

        LayoutShape("2", "shape", "Title 1", Placeholder("title", None), None, "Title")
    """

    id: str
    kind: str
    name: str
    placeholder: Placeholder | None
    bounds: tuple[int, int, int, int] | None
    text: str


#: :meth:`Layout.placeholder` and :meth:`Master.placeholder`'s names -> the ``p:ph`` types
#: they match (``None``: a content placeholder, ``"obj"``).
PLACEHOLDER_NAMES: dict[str, tuple] = {
    "title": ("title", "ctrTitle"), "body": ("body", None, "obj"),
    "content": (None, "obj", "body"), "subtitle": ("subTitle",), "date": ("dt",),
    "footer": ("ftr",), "slide_number": ("sldNum",), "header": ("hdr",),
    "picture": ("pic",), "chart": ("chart",), "table": ("tbl",),
}


class _TemplatePart:
    """What a layout and a master share: their shapes -- the same editable
    :class:`~pptx_agent.Shape` a slide has, addressed ``L<n>.<id>`` / ``M<n>.<id>`` -- and
    adding shapes, pictures, connectors and tables to them."""

    part_path: str
    _document: object
    _key: str

    @property
    def id(self) -> str:
        """The address of this layout (``"L2"``) or master (``"M0"``): its
        ``sldLayoutId`` (``sldMasterId``) less 2147483648, as durable as a slide's id --
        PowerPoint keeps it when it saves.  Its shapes are ``"L2.<n>"``.

        For example::

            deck.layout("Title and Content").id          # 'L2'
        """
        return self._key

    def _sheet(self):
        if self._document is None:
            raise ValueError(f"this {type(self).__name__} was made by hand; take it from "
                             f"deck.{type(self).__name__.lower()}s")
        return self._document._template_sheet(self.part_path)

    @property
    def shapes(self) -> list:
        """Every shape, back to front (groups' members after their group), as editable
        :class:`~pptx_agent.Shape` objects: text, formatting, fill, position, delete,
        replace_image -- what a slide's shapes do.  An edit here shows on every slide
        built on it (:attr:`slides`).

        For example::

            [(shape.id, shape.name, shape.text) for shape in deck.masters[0].shapes]
        """
        return self._sheet().shapes

    def shape(self, shape_id: str):
        """A shape by its address, ``"L2.5"`` (or the local part, ``"5"``).

        For example::

            deck.layout("Title and Content").shape("L2.3").set_prompt("Typ hier")
        """
        text = str(shape_id)
        if "." not in text:
            text = f"{self._key}.{text}"
        return self._sheet().shape(text)

    def placeholder(self, which: "str | int"):
        """The placeholder ``which`` names: ``"title"``, ``"body"`` (or ``"content"``: a
        content placeholder), ``"subtitle"``, ``"date"``, ``"footer"``, ``"slide_number"``,
        ``"picture"``, a ``p:ph`` type (``"ctrTitle"``) -- the first such, back to front --
        or an ``idx`` (``1``).

        For example::

            deck.layout("Title and Content").placeholder("body").set_prompt("Typ hier")
        """
        types = PLACEHOLDER_NAMES.get(which) if isinstance(which, str) else None
        for shape in self.shapes:
            ph = shape.placeholder
            if ph is None:
                continue
            if isinstance(which, int) and not isinstance(which, bool):
                if ph.idx == which:
                    return shape
            elif types is not None and ph.type in types:
                return shape
            elif ph.type == which:
                return shape
        have = [(shape.id, shape.placeholder.type or "content", shape.placeholder.idx)
                for shape in self.shapes if shape.placeholder is not None]
        raise KeyError(f"no placeholder {which!r} on {self._key}; it has {have}")

    @property
    def slides(self) -> list:
        """The slides an edit here shows on: a layout's slides, or every slide of the
        master's layouts, in presentation order.

        For example::

            [slide.slide_id for slide in deck.layout("Title Slide").slides]
        """
        return self._document._slides_on(self.part_path)

    def add_shape(self, *args, **kwargs):
        """Add an autoshape here, as :meth:`Slide.add_shape <pptx_agent.Slide.add_shape>`
        does; every slide built on it draws it, behind the slide's own shapes.

        For example::

            deck.masters[0].add_shape("rect", 0, 0, 12192000, 228600, fill="accent1")
        """
        return self._sheet().add_shape(*args, **kwargs)

    def add_textbox(self, *args, **kwargs):
        """Add a text box here, as :meth:`Slide.add_textbox <pptx_agent.Slide.add_textbox>`
        does -- a legal line, a classification mark.

        For example::

            deck.masters[0].add_textbox(457200, 6400800, 7000000, 300000,
                                        text="© 2027 Example BV")
        """
        return self._sheet().add_textbox(*args, **kwargs)

    def add_picture(self, *args, **kwargs):
        """Add a picture here (SVG too), as :meth:`Slide.add_picture
        <pptx_agent.Slide.add_picture>` does -- a logo every slide shows.

        For example::

            deck.masters[0].add_picture("logo.svg", left=10972800, top=228600, width=914400)
        """
        return self._sheet().add_picture(*args, **kwargs)

    def add_connector(self, *args, **kwargs):
        """Add a connector here, as :meth:`Slide.add_connector
        <pptx_agent.Slide.add_connector>` does -- a rule under the title.

        For example::

            layout.add_connector("straight", (457200, 1371600), (11734800, 1371600))
        """
        return self._sheet().add_connector(*args, **kwargs)

    def add_table(self, *args, **kwargs):
        """Add a table here, as :meth:`Slide.add_table <pptx_agent.Slide.add_table>` does.

        For example::

            layout.add_table(2, 2, 457200, 1600200, 4572000, 914400)
        """
        return self._sheet().add_table(*args, **kwargs)


@dataclass(frozen=True)
class Layout(_TemplatePart):
    """A slide layout: what a new slide is made from, what it offers, and -- since round 3
    -- an editable part of its own, as PowerPoint's Slide Master view edits it.

    ``id`` is its address (``"L2"``); :attr:`shapes` are editable :class:`Shape` objects
    (``"L2.3"``), and every slide built on the layout (:attr:`slides`) shows an edit.
    :meth:`rename`, :meth:`duplicate` and :meth:`delete` change the layout list.

    For example::

        layout = deck.layout("Title and Content")
        print([(ph.type, ph.idx, ph.bounds) for ph in layout.placeholders])
        layout.placeholder("body").set_prompt("Typ hier uw bevinding")
        deck.add_slide(layout)
    """

    name: str
    part_path: str
    master_part: str
    _package: object = field(default=None, compare=False, hash=False, repr=False)
    _document: object = field(default=None, compare=False, hash=False, repr=False)
    _key: str = field(default="", compare=False, hash=False, repr=False)

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
    def master(self) -> "Master":
        """The master this layout belongs to.

        For example::

            deck.layout("Title Slide").master.id         # 'M0'
        """
        return next(master for master in self._document.masters
                    if master.part_path == self.master_part)

    @property
    def type(self) -> str | None:
        """The layout's ``type`` (``"title"``, ``"obj"``, ``"twoObj"``...); ``None`` for a
        custom layout.  A slide on a ``"title"`` layout is a title slide, which
        :meth:`Document.set_footer` leaves out by default, as PowerPoint does.

        For example::

            deck.layout("Title Slide").type              # 'title'
        """
        root = self._require_package().tree(self.part_path)
        return None if root is None else root.get("type")

    def rename(self, name: str) -> "Layout":
        """Rename the layout (``cSld@name``, what Slide Master view shows); returns it as
        renamed.  See :meth:`Document.rename_layout`.

        For example::

            deck.layout("Title and Content").rename("Inhoud")
        """
        return self._document.rename_layout(self, name)

    def duplicate(self, name: str | None = None) -> "Layout":
        """A copy right after this one, as Slide Master view's Duplicate Layout makes it.
        See :meth:`Document.duplicate_layout`.

        For example::

            dutch = deck.layout("Title and Content").duplicate("Inhoud NL")
        """
        return self._document.duplicate_layout(self, name=name)

    def delete(self, *, reassign_to: "Layout | str | None" = None) -> None:
        """Delete this layout.  See :meth:`Document.delete_layout`: one in use is refused
        unless ``reassign_to`` names the layout its slides move to.

        For example::

            deck.layout("Vertical Title and Text").delete()
        """
        self._document.delete_layout(self, reassign_to=reassign_to)

    def _require_package(self):
        if self._package is None:
            raise ValueError("this Layout was made by hand; take it from deck.layouts")
        return self._package

    def _tree(self) -> list[Element]:
        root = self._require_package().tree(self.part_path)
        tree = find(root, "p:cSld/p:spTree") if root is not None else None
        return [] if tree is None else [shape for shape in tree if local_name(shape) in _KINDS]

    def __repr__(self) -> str:
        return f"<Layout {self._key or '?'} {self.name!r} {self.part_path}>"


@dataclass(frozen=True)
class Master(_TemplatePart):
    """A slide master: the shapes every slide of its layouts shows (a logo, a legal line,
    the footer placeholders), its text styles, its theme -- editable as Slide Master view
    edits it, its shapes the same :class:`Shape` objects a slide has (``"M0.5"``).

    ``name`` is the master's own (``cSld@name``) or, when it has none, its theme's, as
    PowerPoint shows it; :attr:`layouts` are its layouts in their order.

    For example::

        master = deck.masters[0]
        legal = next(s for s in master.shapes if "2026" in s.text)
        legal.set_text(legal.text.replace("2026", "2027"))
    """

    name: str
    part_path: str
    _document: object = field(default=None, compare=False, hash=False, repr=False)
    _key: str = field(default="", compare=False, hash=False, repr=False)

    @property
    def layouts(self) -> list[Layout]:
        """This master's layouts, in its order.

        For example::

            [layout.name for layout in deck.masters[0].layouts]
        """
        return [layout for layout in self._document.layouts
                if layout.master_part == self.part_path]

    def add_layout(self, name: str = "Custom Layout", *, after: "Layout | str | None" = None
                   ) -> Layout:
        """A new layout on this master, as Slide Master view's Insert Layout makes it.
        See :meth:`Document.add_layout`.

        For example::

            legal = deck.masters[0].add_layout("Legal")
        """
        return self._document.add_layout(name, master=self, after=after)

    def __repr__(self) -> str:
        return f"<Master {self._key or '?'} {self.name!r} {self.part_path}>"


_KINDS = {"sp": "shape", "pic": "picture", "cxnSp": "connector", "grpSp": "group",
          "graphicFrame": "graphic_frame"}


def _name_of(shape: Element) -> str:
    properties = shape.find(f".//{qn('p:cNvPr')}")
    return (properties.get("name") if properties is not None else None) or ""


# -- reading -----------------------------------------------------------------------------


def template_key(prefix: str, node: Element, fallback: int) -> str:
    """``L<n>`` / ``M<n>``: the ``sldLayoutId``/``sldMasterId`` less 2147483648 (a file
    whose id is missing or below the schema's minimum gets ``fallback`` added to it)."""
    from .ids import TEMPLATE_ID_BASE

    raw = node.get("id") or ""
    value = int(raw) if raw.isdigit() else 0
    return f"{prefix}{value - TEMPLATE_ID_BASE if value >= TEMPLATE_ID_BASE else value + fallback}"


def masters(package: PresentationPackage, document=None) -> list[Master]:
    """Every slide master, in ``sldMasterIdLst`` order."""
    presentation = package.presentation_part()
    root = package.tree(presentation)
    found: list[Master] = []
    for node in findall(root, "p:sldMasterIdLst/p:sldMasterId"):
        part = package.related_part(presentation, node.get(qn("r:id")))
        master_root = package.tree(part) if part else None
        if master_root is None:
            continue
        common = find(master_root, "p:cSld")
        name = (common.get("name") if common is not None else None) or _theme_name(package, part)
        found.append(Master(name, part, document, template_key("M", node, 1 << 20)))
    return found


def _theme_name(package: PresentationPackage, master_part: str) -> str:
    from ..oxml.package import REL_THEME

    themes = package.related_parts_of_type(master_part, REL_THEME)
    root = package.tree(themes[0]) if themes else None
    return (root.get("name") if root is not None else None) or ""


def layouts(package: PresentationPackage, document=None) -> list[Layout]:
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
            found.append(Layout(name, part, master, package, document,
                                template_key("L", layout_id, 1 << 21)))
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


def slide_from_layout(package: PresentationPackage, layout_part: str,
                      language: str | None = None) -> bytes:
    """A new slide's XML: the layout's placeholders, empty, inheriting everything, in the
    deck's default language (``language`` where the deck states none)."""
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

    language = default_language(package, language)
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
    "Master",
    "masters",
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
