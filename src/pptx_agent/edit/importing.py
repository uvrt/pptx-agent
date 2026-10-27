"""A slide of another open deck imported whole (:meth:`Document.import_slide`, round 17, P1).

:meth:`Document.duplicate_slide` copies a slide within its deck; this copies one into
another deck -- a sample slide kept in a deck of samples, say, made from the same template.

What the copy keeps and what it changes:

* the slide part with everything it relates to: pictures and other media (imported once by
  content: a picture the target already holds, byte for byte, is shared), charts with their
  embedded workbooks, SmartArt parts, tags, external hyperlinks, and its notes page (on the
  target's notes master, added when the target has none); every part gets a new name;
* the layout is the target's layout **of the same name**, on the master of the same name
  (when no master of the target has that name, the layout of the name on any master).  A
  target without one refuses, listing its layouts -- unless ``bring_layout``: the source's
  layout is then copied onto the target's master of the same name; when the target has no
  master of that name, the source's master comes too (with its theme, listing only that
  layout), because a layout draws from its master -- its placeholders' places and text
  styles, its background, its theme -- and on another master would not look as it did;
* shape names and ``cNvPr`` ids stay (so a shape's address is the new slide's id with the
  same shape half: ``300.4`` becomes ``412.4``); the slide gets a new ``sldId``;
* theme colours stay theme references.  When the target slide's theme is not the source
  slide's (other colours or fonts), the copy takes the target's: nothing is converted, and
  :attr:`ImportResult.theme_differs` says so;
* not copied: comments (the review is of the original, as for a duplicate), links that jump
  to another slide of the source deck (that slide is not in the target; the link is taken
  off its text or shape and the shape reported), slide numbers written by reference
  (``{{slide:s:262}}``: they name the source's slides; the text stays).

The whole import is one undo step on the target deck; the source is only read.
"""

from __future__ import annotations

import posixpath
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Callable

from ooxml_edit.opc import RELS_NS, rels_path_for
from ooxml_edit.xml import parse_xml, serialize

from ..oxml.package import (
    CT_SLIDE,
    REL_NOTES_MASTER,
    REL_NOTES_SLIDE,
    REL_SLIDE,
    REL_SLIDE_LAYOUT,
    REL_SLIDE_MASTER,
)
from ..oxml.xml import find, findall, make, qn, remove
from .errors import ArgumentError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, Slide
    from .slides import Layout

_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#: Comments are the original's review (legacy and modern): never copied.
_COMMENTS = frozenset({
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
    "http://schemas.microsoft.com/office/2018/10/relationships/comments",
})
_CT_MASTER = "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"


@dataclass
class ImportResult:
    """What :meth:`Document.import_slide` made."""

    #: The new slide in the target deck.
    slide: "Slide"
    #: The target layout it is on, and that layout's master's name.
    layout: "Layout"
    master: str
    #: Every shape of the source slide (groups' members included): source id -> new id.
    mapping: dict[str, str] = field(default_factory=dict)
    #: What ``bring_layout`` copied into the target: ``"layout"`` and, when the target had
    #: no master of the name, ``"master"`` too.
    brought: list[str] = field(default_factory=list)
    #: Whether the target slide's theme (colours or fonts) is not the source slide's.
    theme_differs: bool = False
    #: The theme colours the slide uses that the target draws differently:
    #: ``{"color": "accent1", "was": "#134074", "now": "#4472C4"}``.
    theme_changes: list[dict] = field(default_factory=list)
    #: Shapes (new ids) whose link to another slide of the source deck was taken off.
    dropped_links: list[str] = field(default_factory=list)


def import_slide(document: "Document", source: "Slide", index: int | None = None, *,
                 bring_layout: bool = False, notes: str | None = None) -> ImportResult:
    """Copy ``source`` (a slide of another open deck) into ``document`` at 0-based
    ``index`` (default last); ``notes``, when given, replaces the copy's speaker notes (as
    :meth:`Document.duplicate_slide` does).  See the module docstring."""
    if source.document is document:
        raise ArgumentError("import_slide copies from another deck; within one deck, "
                            "duplicate the slide")
    source_layout = source.layout
    if source_layout is None:
        raise ArgumentError(f"s:{source.slide_id} has no layout to map")
    master_name = source_layout.master.name
    layout = find_layout(document, source_layout.name, master_name)
    if layout is None and not bring_layout:
        raise ArgumentError(
            f"the target deck has no layout {source_layout.name!r} (master {master_name!r}); "
            "its layouts: " + ", ".join(f"{lay.name!r}" for lay in document.layouts)
            + ". Pass bring_layout to copy the source's layout in, or set_layout after a "
            "duplicate in the source deck")
    source_package, package = source.document.package, document.package
    imported: dict[str, str] = {}
    brought: list[str] = []
    with document.batch():
        document.history.checkpoint()
        if layout is None:
            layout = _bring_layout(document, source_layout, imported, brought)
        part = package.unused_part_name("ppt/slides/slide{n}.xml")
        package.add_part(part, source_package.read(source.part_path), CT_SLIDE, override=True)
        jumps: list[str] = []

        def route(relationship):
            kind = relationship.type
            if kind == REL_SLIDE_LAYOUT:
                return layout.part_path
            if kind in _COMMENTS:
                return None
            if kind == REL_SLIDE:
                jumps.append(relationship.id)
                return None
            if kind == REL_NOTES_SLIDE:
                return _import_notes(document, source_package, relationship.target_part,
                                     part, imported)
            return _import_part(source_package, package, relationship.target_part, imported)

        _copy_relationships(source_package, source.part_path, package, part, route)
        dropped = _drop_jumps(package, part, jumps)
        from . import slides as _slides
        from .comments import drop_comments

        drop_comments(package, part)
        slide_id = document._new_slide_entry(part, index)
        _slides.renew_creation_id(package, part, slide_id)
        document._reset_caches()
        slide = document.slide(slide_id)
        from .slidenumbers import drop

        drop([shape for shape in slide.shapes if shape.parent_group is None])
        document._copy_notes(slide, notes)
        mapping = {old.id: new.id for old, new in zip(source.shapes, slide.shapes)}
        result = ImportResult(slide, document.layout(layout.part_path), layout.master.name,
                              mapping, brought)
        result.dropped_links = [shape.id for shape in slide.shapes
                                if any(element is shape._element for element in dropped)]
        before, after = source.theme, slide.theme
        result.theme_differs = (dict(before.colors) != dict(after.colors)
                                or (before.fonts.major, before.fonts.minor)
                                != (after.fonts.major, after.fonts.minor))
        if result.theme_differs:
            from .copying import _theme_changes

            result.theme_changes = _theme_changes(package.tree(part), source, slide)
    return result


def find_layout(document: "Document", name: str, master: str) -> "Layout | None":
    """``document``'s layout named ``name`` on its master named ``master``; when no master
    has that name, the first layout named ``name``; else ``None``."""
    layouts = document.layouts
    if any(found.name == master for found in document.masters):
        return next((lay for lay in layouts if lay.name == name and lay.master.name == master),
                    None)
    return next((lay for lay in layouts if lay.name == name), None)


# -- parts ---------------------------------------------------------------------------------


def _import_part(source_package, package, path: str | None, imported: dict[str, str]):
    from .copying import _import

    if path is None or not source_package.has_part(path):
        return path
    return _import(source_package, package, path, imported)


def _copy_relationships(source_package, source_part: str, package, new_part: str,
                        route: Callable) -> None:
    """``new_part``'s relationships: ``source_part``'s, each internal one re-pointed where
    ``route(relationship)`` says (``None``: left out).  Ids are kept, so the part's XML
    needs no rewriting; external ones are copied as they are."""
    root = source_package.tree(rels_path_for(source_part))
    if root is None:
        return
    copied = parse_xml(serialize(root))
    relationships = source_package.relationships(source_part)
    base = posixpath.dirname(new_part)
    for node in list(copied.iter("{%s}Relationship" % RELS_NS)):
        relationship = relationships.get(node.get("Id"))
        if relationship is None or relationship.is_external:
            continue
        target = route(relationship)
        if target is None:
            node.getparent().remove(node)
        else:
            node.set("Target", posixpath.relpath(target, base or "."))
    package.add_part(rels_path_for(new_part), serialize(copied),
                     source_package.content_type(rels_path_for(source_part)))


def _drop_jumps(package, part: str, rel_ids: list[str]) -> list:
    """Take the links naming ``rel_ids`` (jumps to the source's slides) off; the shapes
    that had one (their elements)."""
    if not rel_ids:
        return []
    root = package.tree(part)
    owners = []
    for node in list(root.iter()):
        if not isinstance(node.tag, str) or node.get(f"{{{_R}}}id") not in rel_ids:
            continue
        owner = node.getparent()
        while owner is not None and owner.getparent() is not None \
                and owner.getparent().tag not in (qn("p:spTree"), qn("p:grpSp")):
            owner = owner.getparent()
        if owner is not None and all(owner is not seen for seen in owners):
            owners.append(owner)
        remove(node)
    package.mark_dirty(part)
    return owners


def _import_notes(document: "Document", source_package, notes: str | None, slide_part: str,
                  imported: dict[str, str]) -> str | None:
    """The source's notes page copied for the new slide, on the target's notes master."""
    from .notes import CT_NOTES_SLIDE, _ensure_master, _language

    if notes is None or not source_package.has_part(notes):
        return None
    package = document.package
    master = _ensure_master(package, _language(document), document.slide_size)
    part = package.unused_part_name("ppt/notesSlides/notesSlide{n}.xml")
    package.add_part(part, source_package.read(notes), CT_NOTES_SLIDE, override=True)

    def route(relationship):
        if relationship.type == REL_NOTES_MASTER:
            return master
        if relationship.type == REL_SLIDE:
            return slide_part
        return _import_part(source_package, package, relationship.target_part, imported)

    _copy_relationships(source_package, notes, package, part, route)
    return part


def _bring_layout(document: "Document", layout: "Layout", imported: dict[str, str],
                  brought: list[str]) -> "Layout":
    """``layout`` (the source's) copied onto the target's master of its master's name --
    with that master, when the target has none of the name."""
    from .templates import CT_SLIDE_LAYOUT, _list_layout

    source_package, package = layout._document.package, document.package
    master_name = layout.master.name
    master = next((found.part_path for found in document.masters
                   if found.name == master_name), None)
    if master is None:
        master = _bring_master(document, source_package, layout.master_part, imported)
        brought.append("master")
    part = package.unused_part_name("ppt/slideLayouts/slideLayout{n}.xml")
    package.add_part(part, source_package.read(layout.part_path), CT_SLIDE_LAYOUT,
                     override=True)

    def route(relationship):
        if relationship.type == REL_SLIDE_MASTER:
            return master
        return _import_part(source_package, package, relationship.target_part, imported)

    _copy_relationships(source_package, layout.part_path, package, part, route)
    _list_layout(document, master, part, None)
    brought.append("layout")
    document._reset_caches()
    return document.layout(part)


def _bring_master(document: "Document", source_package, source_master: str,
                  imported: dict[str, str]) -> str:
    """The source's master copied in, its theme with it, listing none of its layouts yet
    (the one brought is listed next); returns its part."""
    from .templates import _next_template_id

    package = document.package
    part = package.unused_part_name("ppt/slideMasters/slideMaster{n}.xml")
    root = parse_xml(source_package.read(source_master))
    for entry in findall(root, "p:sldLayoutIdLst/p:sldLayoutId"):
        remove(entry)
    package.add_part(part, serialize(root), _CT_MASTER, override=True)

    def route(relationship):
        if relationship.type == REL_SLIDE_LAYOUT:
            return None
        return _import_part(source_package, package, relationship.target_part, imported)

    _copy_relationships(source_package, source_master, package, part, route)
    presentation = package.presentation_part()
    rel_id = package.add_relationship(presentation, REL_SLIDE_MASTER, part)
    entries = find(package.tree(presentation), "p:sldMasterIdLst")
    entries.append(make("p:sldMasterId", id=str(_next_template_id(package)), r__id=rel_id))
    package.mark_dirty(presentation)
    return part


__all__ = ["ImportResult", "find_layout", "import_slide"]
