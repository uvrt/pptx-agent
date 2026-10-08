"""Copying shapes between slides and decks (:meth:`Slide.copy_shapes`).

:meth:`Shape.duplicate` copies one shape into its own container on its own slide.  This
module copies a *set* of shapes -- loose shapes, groups, connectors between them, pictures,
charts, SmartArt -- to any slide: the same one, another one, or a slide of another open
:class:`~pptx_agent.Document`.

What a copy keeps and what it changes:

* every shape in it gets a fresh ``cNvPr@id``, group members included, and loses any
  stamped id or ``creationId`` -- it is a different shape;
* a connector glued to a shape inside the set is glued to that shape's copy; glue to a
  shape outside the set is dropped (the connector keeps its drawn position) and reported;
* what a shape relates to is carried along: media and hyperlinks are shared within a deck
  and imported (once, by content) into another; charts with their embedded workbooks and
  SmartArt diagrams are per-shape state and copied, as :meth:`Document.duplicate_slide`
  copies them; a link to another slide cannot go to another deck and is refused;
* a placeholder becomes a plain shape, with the geometry it inherited written out and each
  run's effective size stated;
* a shape copied out of a group lands on the slide at the place it was drawn;
* theme colours stay theme references.  When the target slide's theme gives one of them
  another colour, the result says so (:attr:`CopyResult.theme_changes`) -- a fact, nothing
  is recoloured.

The whole copy is one undo step on the target deck; the source is only read.
"""

from __future__ import annotations

import copy
import posixpath
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterable, Sequence

from lxml import etree
from ooxml_edit.opc import RELS_NS, numbered_template, rels_path_for

from ..oxml.package import (
    REL_NOTES_MASTER,
    REL_NOTES_SLIDE,
    REL_SLIDE,
    REL_SLIDE_LAYOUT,
    REL_SLIDE_MASTER,
    SHARED_ON_DUPLICATE,
)
from ..oxml.xml import NAMESPACES, append_in_order, local_name, make, qn, remove
from ooxml_edit.xml import parse_xml, serialize

if TYPE_CHECKING:
    from .document import Shape, Slide

#: Relationships a shape never owns: the slide's own structure.
_STRUCTURAL = frozenset({REL_SLIDE_LAYOUT, REL_SLIDE_MASTER, REL_NOTES_SLIDE, REL_NOTES_MASTER,
                         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"})
_SHAPE_TAGS = ("sp", "pic", "cxnSp", "grpSp", "graphicFrame", "contentPart")
_XFRM_PARENT = {"sp": "p:spPr", "pic": "p:spPr", "cxnSp": "p:spPr", "grpSp": "p:grpSpPr",
                "graphicFrame": None}
_R = NAMESPACES["r"]


@dataclass
class CopyResult:
    """What :meth:`Slide.copy_shapes` made."""

    #: The copies of the given shapes, top level on the target slide, in the order given.
    shapes: list["Shape"] = field(default_factory=list)
    #: Every copied shape, nested members included: source address -> new address.
    mapping: dict[str, str] = field(default_factory=dict)
    #: Connector ends that were glued outside the set and are now loose:
    #: ``{"connector": new address, "end": "begin"|"end", "was_glued_to": raw cNvPr id}``.
    dropped_glue: list[dict] = field(default_factory=list)
    #: Theme colours the copy uses that the target slide's theme draws differently:
    #: ``{"color": "accent2", "was": "#134074", "now": "#ED7D31"}``.
    theme_changes: list[dict] = field(default_factory=list)
    #: Source placeholders that became plain shapes.
    plain_placeholders: list[str] = field(default_factory=list)


def copy_shapes(shapes: "Sequence[Shape]", to_slide: "Slide | None" = None, *,
                at: tuple[int, int] | None = None, dx: int = 0, dy: int = 0) -> CopyResult:
    """Copy ``shapes`` (all on one slide) to ``to_slide`` (default: their own slide).

    ``at`` is where the top-left corner of the set's bounds goes (slide EMU); otherwise the
    copy is offset by ``(dx, dy)`` from where the originals are.  See the module docstring.
    """
    if not shapes:
        raise ValueError("copy_shapes needs at least one shape")
    source = shapes[0]._slide
    for shape in shapes:
        if shape._slide.document is not source.document or shape._slide.slide_id != source.slide_id:
            raise ValueError(f"{shape.id}: every shape to copy must be on one slide "
                             f"(s:{source.slide_id})")
    chosen = _top_level(shapes)
    target = to_slide if to_slide is not None else source
    bounds = [shape.slide_bounds for shape in chosen]
    for shape, box in zip(chosen, bounds):
        if box is None:
            raise ValueError(f"{shape.id} has no position to copy from")
    if at is not None:
        dx = int(at[0]) - min(box[0] for box in bounds)
        dy = int(at[1]) - min(box[1] for box in bounds)

    result = CopyResult()
    document = target.document
    source_package = source.document.package
    package = document.package
    with document.batch():
        document.history.checkpoint()
        target._watch()
        holder = etree.Element("holder")
        sources: list[list] = []  # per chosen shape: its shape-tag elements, in order
        for shape in chosen:
            clone = copy.deepcopy(shape._element)
            _strip(clone)
            holder.append(clone)
            sources.append(_shape_elements(shape._element))
        inside = {node.get("id") for elements in sources for element in elements
                  for node in [_cnv(element)] if node is not None}
        loose = [(glue, glue.getparent().getparent().getparent()) for tag in ("a:stCxn", "a:endCxn")
                 for glue in holder.iter(qn(tag)) if glue.get("id") not in inside]
        from .document import renumber_copy

        renumber_copy(holder, target._next_shape_ids)
        for glue, connector in loose:
            result.dropped_glue.append({
                "connector": connector, "end": "begin" if local_name(glue) == "stCxn" else "end",
                "was_glued_to": glue.get("id")})
            remove(glue)
        _carry_relationships(holder, source_package, source.part_path, package,
                             target.part_path)
        result.theme_changes = _theme_changes(holder, source, target)
        clones = list(holder)
        for shape, clone, box in zip(chosen, clones, bounds):
            placeholder = _drop_placeholder(clone)
            if placeholder:
                result.plain_placeholders.append(shape.id)
            if placeholder or shape.parent_group is not None or not _has_offset(clone):
                _write_frame(clone, box)
            _translate(clone, dx, dy)
        tree = target._sp_tree()
        for clone in clones:
            append_in_order(tree, clone)
        target._invalidate()
        for shape, clone, elements in zip(chosen, clones, sources):
            copied = target._wrap(clone)
            result.shapes.append(copied)
            for old, new in zip(elements, _shape_elements(clone)):
                result.mapping[source._wrap(old).id] = target._wrap(new).id
            if shape.id in result.plain_placeholders:
                _state_sizes(shape, copied)
        for entry in result.dropped_glue:
            entry["connector"] = target._wrap(entry["connector"]).id
        target._settle()
    return result


# -- the set -------------------------------------------------------------------------------


def _top_level(shapes: "Sequence[Shape]") -> "list[Shape]":
    """The given shapes without repeats and without members of a group also given, in
    document order."""
    elements = {id(shape._element): shape for shape in shapes}
    kept = []
    for shape in shapes:
        ancestor = shape._element.getparent()
        covered = False
        while ancestor is not None:
            if id(ancestor) in elements:
                covered = True
                break
            ancestor = ancestor.getparent()
        if not covered and all(other._element is not shape._element for other in kept):
            kept.append(shape)
    order = {id(node): n for n, node in enumerate(shapes[0]._slide._sp_tree().iter())}
    return sorted(kept, key=lambda shape: order.get(id(shape._element), 0))


def _cnv(element):
    for child in element:
        if isinstance(child.tag, str) and local_name(child).startswith("nv"):
            return child.find(qn("p:cNvPr"))
    return None


def _shape_elements(root) -> list:
    """``root`` and every shape inside it, in document order."""
    return [node for node in root.iter()
            if isinstance(node.tag, str) and local_name(node) in _SHAPE_TAGS
            and _cnv(node) is not None]


def _strip(clone) -> None:
    from .document import _strip_identity

    _strip_identity(clone)


# -- relationships ---------------------------------------------------------------------------


def _carry_relationships(holder, source_package, source_part, package, target_part) -> None:
    """Re-point every relationship the copies use at the target slide part."""
    carried: dict[str, str] = {}
    imported: dict[str, str] = {}

    def carry(rel_id: str) -> str:
        if rel_id not in carried:
            carried[rel_id] = _carry(rel_id, source_package, source_part, package, target_part,
                                     imported)
        return carried[rel_id]

    for node in holder.iter():
        if not isinstance(node.tag, str):
            continue
        for name, value in list(node.attrib.items()):
            if name.startswith("{%s}" % _R) and value:
                node.set(name, carry(value))
        if local_name(node) == "dataModelExt" and node.get("relId"):
            node.set("relId", carry(node.get("relId")))


def _carry(rel_id, source_package, source_part, package, target_part, imported) -> str:
    relationship = source_package.relationships(source_part).get(rel_id)
    if relationship is None:
        raise ValueError(f"the shapes use {rel_id}, which {source_part} does not define")
    if relationship.is_external:
        return package.add_external_relationship(target_part, relationship.type,
                                                 relationship.target)
    if relationship.type in _STRUCTURAL:
        raise ValueError(f"a shape cannot carry a {relationship.type.rsplit('/', 1)[-1]} "
                         "relationship")
    part = relationship.target_part
    if package is source_package:
        if relationship.type not in SHARED_ON_DUPLICATE:
            part = package.copy_part(part, share=lambda rel: rel.type in SHARED_ON_DUPLICATE)
    else:
        if relationship.type == REL_SLIDE:
            raise ValueError("a link to another slide cannot be copied into another deck; "
                             "remove the link first")
        part = _import(source_package, package, part, imported)
    return package.add_relationship(target_part, relationship.type, part)


def _import(source, package, path: str, imported: dict[str, str]) -> str:
    """Copy part ``path`` of ``source`` into ``package``, with what it relates to."""
    if path in imported:
        return imported[path]
    data = source.read(path)
    if data is None:
        raise ValueError(f"the shapes relate to {path}, which is missing")
    content_type = source.content_type(path)
    directory = posixpath.dirname(path)
    if path.startswith("ppt/media/"):
        existing = package.find_part_with_bytes(data, directory)
        if existing is not None and package.content_type(existing) == content_type:
            imported[path] = existing
            return existing
    new = package.unused_part_name(numbered_template(path))
    imported[path] = new
    package.add_part(new, data, content_type, override=source._has_override(path))
    rels = source.tree(rels_path_for(path))
    if rels is not None:
        copied = parse_xml(serialize(rels))
        relationships = source.relationships(path)
        base = posixpath.dirname(new)
        for node in copied.iter("{%s}Relationship" % RELS_NS):
            relationship = relationships.get(node.get("Id"))
            if relationship is None or relationship.is_external or relationship.target_part is None:
                continue
            if relationship.type in _STRUCTURAL or relationship.type == REL_SLIDE:
                raise ValueError(f"{path} relates to {relationship.target_part}, which cannot "
                                 "be copied into another deck")
            if not source.has_part(relationship.target_part):
                continue
            moved = _import(source, package, relationship.target_part, imported)
            node.set("Target", posixpath.relpath(moved, base or "."))
        package.add_part(rels_path_for(new), serialize(copied),
                         source.content_type(rels_path_for(path)))
    return new


# -- geometry ---------------------------------------------------------------------------------


def _xfrm(element, create: bool):
    tag = local_name(element)
    if tag == "graphicFrame":
        holder = element
        name = "p:xfrm"
    else:
        holder = element.find(qn(_XFRM_PARENT.get(tag) or "p:spPr"))
        if holder is None:
            if not create:
                return None
            holder = make(_XFRM_PARENT.get(tag) or "p:spPr")
            append_in_order(element, holder)
        name = "a:xfrm"
    node = holder.find(qn(name))
    if node is None and create:
        node = make(name)
        holder.insert(0, node)
    return node


def _has_offset(element) -> bool:
    node = _xfrm(element, create=False)
    return node is not None and node.find(qn("a:off")) is not None \
        and node.find(qn("a:ext")) is not None


def _write_frame(element, box) -> None:
    """Give a copy the frame it was drawn with: ``box`` on the slide (EMU)."""
    node = _xfrm(element, create=True)
    left, top, width, height = (int(v) for v in box)
    offset = node.find(qn("a:off"))
    if offset is None:
        offset = make("a:off")
        node.insert(0, offset)
    offset.set("x", str(left))
    offset.set("y", str(top))
    extent = node.find(qn("a:ext"))
    if extent is None:
        extent = make("a:ext")
        offset.addnext(extent)
    extent.set("cx", str(width))
    extent.set("cy", str(height))


def _translate(element, dx: int, dy: int) -> None:
    if not dx and not dy:
        return
    offset = _xfrm(element, create=True).find(qn("a:off"))
    offset.set("x", str(int(offset.get("x", "0")) + int(dx)))
    offset.set("y", str(int(offset.get("y", "0")) + int(dy)))


# -- placeholders and theme ---------------------------------------------------------------------


def _drop_placeholder(element) -> bool:
    found = False
    for ph in list(element.iter(qn("p:ph"))):
        remove(ph)
        found = True
    return found


def _state_sizes(source: "Shape", copied: "Shape") -> None:
    """A placeholder's copy no longer inherits its sizes: state each run's."""
    before, after = source.text_frame, copied.text_frame
    if before is None or after is None:
        return
    for old, new in zip(before.paragraphs, after.paragraphs):
        for run_old, run_new in zip(old.runs, new.runs):
            if run_new.size is None:
                run_new.size = run_old.effective_size


def _theme_changes(holder, source: "Slide", target: "Slide") -> list[dict]:
    names = sorted({node.get("val") for node in holder.iter(qn("a:schemeClr")) if node.get("val")})
    if not names or (source.document is target.document and source.slide_id == target.slide_id):
        return []
    before, after = source.theme.colors, target.theme.colors
    changes = []
    for name in names:
        if name in before and name in after and before[name].upper() != after[name].upper():
            changes.append({"color": name, "was": before[name], "now": after[name]})
    return changes


__all__ = ["CopyResult", "copy_shapes"]
