"""Source addresses kept in the deck: where a table's or a chart's values came from.

A tool that places values from a source (``ppt_add_table`` and ``ppt_edit_table`` with
``from``, ``edit_chart`` with ``from``) says where each came from in its result; the
:mod:`ooxml_edit.provenance` *record* of that is also written into the deck, so it can be
read back after a save, in another session, or after PowerPoint has saved the file.

**Where.**  In the graphic frame's tags part, the place PowerPoint keeps an add-in's tags:
``p:nvGraphicFramePr/p:nvPr/p:custDataLst/p:tags r:id`` relates the frame, through its
slide, to a ``ppt/tags/tagN.xml`` part
(``application/vnd.openxmlformats-officedocument.presentationml.tags+xml``) holding::

    <p:tagLst xmlns:a="..." xmlns:r="..." xmlns:p="...">
      <p:tag name="UVRT_PROVENANCE" val="{&quot;v&quot;:1,&quot;source&quot;:...}"/>
    </p:tagLst>

``val`` is :func:`ooxml_edit.provenance.dumps` of the record.  PowerPoint for Mac kept
such a tag, on the right frame and with its value byte for byte, through open, a text edit,
a table cell edit and a chart resize, and save (measured 2026-10-14); it renumbers the
relationship ids and re-serialises the part around the tag.  The name is upper case because
PowerPoint's object model upper-cases tag names.  Other tags on the frame stay, as they
are; a tags part another frame shares is copied before it is written, so a record never
lands on two frames.  :data:`TAG_NAME` is the one tag this module writes.

**Text facts.**  A value a text tool wrote by reference (``{{d:x/key/col}}``) is kept in
the same way on the shape it is in -- any shape with text, through its own ``p:nvPr`` -- or,
for a table cell, on the table's frame: a :data:`FACTS_TAG` tag beside the
:data:`TAG_NAME` one (see "text facts" below).  That PowerPoint keeps such a tag on a
``p:sp`` as it does on a graphic frame is not measured: the element is the same
(``CT_ApplicationNonVisualDrawingProps``' ``custDataLst`` in ECMA-376's ``pml.xsd``).

**Copies and deletes.**  The tags part is per-shape state: a duplicated slide, a duplicated
shape and shapes copied to another slide or deck get their own copy of it
(:data:`~pptx_agent.oxml.package.SHARED_ON_DUPLICATE` does not share tags); deleting the
shape or the slide releases the relationship and reaps the part.

:func:`store` writes or drops a frame's record; :func:`read` reads one;
:func:`records` reads the deck's (:meth:`Document.provenance
<pptx_agent.Document.provenance>`).  The table rules -- what a cell write, a row or a
column insert or delete does to a record -- are the pure functions after them
(:func:`after_cells`, :func:`after_row_insert`, ...): a record never names a cell its value
did not come to, so a change it cannot follow drops the entry, or the record.
"""

from __future__ import annotations

import re
import warnings
from typing import TYPE_CHECKING, Any, Iterable, Mapping

from ooxml_edit import provenance as _records

from ..oxml.xml import NAMESPACES, append_in_order, find, make, qn, remove, serialize

if TYPE_CHECKING:
    from .document import Document, Shape

#: The tag that holds the record.  Upper case: PowerPoint upper-cases tag names.
TAG_NAME = "UVRT_PROVENANCE"
#: The tag that holds a shape's or a table's text facts: ``{"v":1,"facts":[...]}``.
FACTS_TAG = "UVRT_TEXT_FACTS"
#: The relationship from a slide to a tags part, and the part's content type.
REL_TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
CT_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"

_R_ID = "{%s}id" % NAMESPACES["r"]


class ProvenanceWarning(UserWarning):
    """A tag that should hold a record does not: it is skipped, never trusted."""


# -- the frame's tags part ----------------------------------------------------------------------


def _frame(shape: "Shape") -> None:
    if shape._element.tag != qn("p:graphicFrame"):
        raise ValueError(f"{shape.id} is a {shape.kind}: a source address is kept on a "
                         "table's or a chart's graphic frame")


def _nv_pr(shape: "Shape"):
    """The shape's ``p:nvPr`` (in ``p:nvSpPr``, ``p:nvGraphicFramePr``, ...)."""
    return shape._element.find("./*/p:nvPr", NAMESPACES)


def _tags_element(shape: "Shape"):
    nv_pr = _nv_pr(shape)
    return None if nv_pr is None else find(nv_pr, "p:custDataLst/p:tags")


def _tags_part(shape: "Shape") -> str | None:
    """The tags part the frame relates to (``None``: none, or a dangling id)."""
    element = _tags_element(shape)
    if element is None or not element.get(_R_ID):
        return None
    package = shape._slide.document.package
    part = package.related_part(shape._slide.part_path, element.get(_R_ID))
    return part if part is not None and package.has_part(part) else None


def read(shape: "Shape", name: str = TAG_NAME) -> dict[str, Any] | None:
    """The record kept on ``shape``'s frame (``name``: :data:`FACTS_TAG`, its text facts,
    on any shape), or ``None``.  An unreadable one is ``None``, with a
    :class:`ProvenanceWarning`."""
    if name == TAG_NAME and shape._element.tag != qn("p:graphicFrame"):
        return None
    part = _tags_part(shape)
    if part is None:
        return None
    tag = _our_tag(shape._slide.document.package.tree(part), name)
    if tag is None:
        return None
    try:
        return _records.loads(tag.get("val") or "")
    except _records.RecordError as error:
        warnings.warn(f"{shape.id}: the {name} tag holds no source record ({error}); "
                      "it is skipped", ProvenanceWarning, stacklevel=2)
        return None


def _our_tag(root, name: str = TAG_NAME):
    if root is None:
        return None
    for tag in root.iter(qn("p:tag")):
        if (tag.get("name") or "").upper() == name:
            return tag
    return None


def store(shape: "Shape", found: Mapping[str, Any] | None, name: str = TAG_NAME) -> bool:
    """Keep ``found`` on ``shape``'s graphic frame, replacing the record it had; ``None``
    drops it -- and the tags part and its relationship, once no tag is left.  Other tags
    stay.  Part of the document's current undo step.  Whether anything changed.
    ``name``: :data:`FACTS_TAG` keeps text facts, on any shape."""
    if name == TAG_NAME:
        _frame(shape)
    slide = shape._slide
    package = slide.document.package
    if found is None:
        part = _tags_part(shape)
        tag = _our_tag(package.tree(part), name) if part is not None else None
        if tag is None:
            return False
        shape._before_change()
        part = _own_part(shape)                  # a shared part: this frame's copy only
        root = package.tree(part)
        remove(_our_tag(root, name))
        package.mark_dirty(part)
        if root.find(qn("p:tag")) is None:
            _unlink(shape)
        shape._after_change()
        return True
    text = _records.dumps(found)
    part = _own_part(shape)
    if part is not None:
        tag = _our_tag(package.tree(part), name)
        if tag is not None and tag.get("val") == text:
            return False
    shape._before_change()
    if part is None:
        part = _new_part(shape)
    root = package.tree(part)
    tag = _our_tag(root, name)
    if tag is None:
        tag = make("p:tag")
        root.append(tag)
    tag.set("name", name)
    tag.set("val", text)
    package.mark_dirty(part)
    shape._after_change()
    return True


def _own_part(shape: "Shape") -> str | None:
    """The frame's tags part, copied first when anything else relates to it too -- another
    slide, or another frame of this one -- so a write reaches this frame alone."""
    part = _tags_part(shape)
    if part is None:
        return None
    slide = shape._slide
    package = slide.document.package
    rel_id = _tags_element(shape).get(_R_ID)
    sources = package.relationship_sources(part)
    users = [node for node in slide._sp_tree().iter(qn("p:tags")) if node.get(_R_ID) == rel_id]
    if len(sources) <= 1 and len(users) <= 1:
        return part
    shape._before_change()
    copied = package.copy_part(part, share=lambda rel: True)
    _tags_element(shape).set(_R_ID, package.add_relationship(slide.part_path, REL_TAGS,
                                                             copied))
    slide._touch()
    return copied


def _new_part(shape: "Shape") -> str:
    """A new, empty tags part, related from the slide and named in the frame."""
    slide = shape._slide
    package = slide.document.package
    part = package.unused_part_name("ppt/tags/tag{n}.xml")
    package.add_part(part, serialize(_tag_list()), CT_TAGS, override=True)
    rel_id = package.add_relationship(slide.part_path, REL_TAGS, part)
    holder = _tags_element(shape)
    if holder is None:
        nv_pr = _nv_pr(shape)
        if nv_pr is None:
            raise ValueError(f"{shape.id}: the shape has no p:nvPr")
        custom = find(nv_pr, "p:custDataLst")
        if custom is None:
            custom = append_in_order(nv_pr, make("p:custDataLst"))
        holder = make("p:tags")
        custom.append(holder)          # CT_CustomerDataList: custData*, then tags
    holder.set(_R_ID, rel_id)
    slide._touch()
    return part


def _tag_list():
    """An empty ``p:tagLst`` declaring the ``a``, ``r`` and ``p`` prefixes, as PowerPoint
    writes one."""
    from lxml import etree

    return etree.Element(qn("p:tagLst"), nsmap={"a": NAMESPACES["a"], "r": NAMESPACES["r"],
                                                "p": NAMESPACES["p"]})


def _unlink(shape: "Shape") -> None:
    """Take the frame's ``p:tags`` (and an emptied ``p:custDataLst``) away; the slide's
    settle releases the relationship and reaps the part."""
    element = _tags_element(shape)
    custom = element.getparent()
    remove(element)
    if len(custom) == 0:
        remove(custom)


# -- text facts ---------------------------------------------------------------------------------
#
# A value written into a shape's or a table cell's text by reference (``{{d:x/key/col}}``,
# :mod:`ooxml_edit.tools.textfacts`) is a *fact*: an :mod:`ooxml_edit.provenance` fact --
# ``source``, ``text``, ``document``, ``version``, ``sources`` and its place, ``p`` the
# paragraph and ``at`` the character in it, in the text ``text_frame.text`` reads (a line
# break is one character, ``\v``) -- and, in a table, ``cell`` ``[row, col]``.  A shape's
# facts are one record, ``{"v":1,"facts":[...]}``, in the :data:`FACTS_TAG` tag of its tags
# part, beside a table's or a chart's own record when it has one.  A fact is kept while its
# characters are its text: an edit keeps the ones it does not touch, moved with the text
# (:func:`ooxml_edit.provenance.facts_after_edit`), and reading lists only the ones still
# where they say (:func:`ooxml_edit.provenance.facts_found`), whoever edited the text.


def text_of(shape: "Shape", cell: tuple[int, int] | None = None) -> str | None:
    """The text a fact's place counts in: the shape's, or a table cell's; ``None`` when
    there is none."""
    if cell is not None:
        table = shape.table
        if table is None or not (cell[0] < table.rows and cell[1] < table.columns):
            return None
        return table.cell(*cell).text_frame.text
    frame = shape.text_frame
    return None if frame is None else frame.text


def facts(shape: "Shape") -> list[dict[str, Any]]:
    """Every fact the shape keeps, as kept (not checked against its text)."""
    found = read(shape, FACTS_TAG)
    listed = found.get("facts") if found else None
    return [dict(fact) for fact in listed if isinstance(fact, Mapping)] \
        if isinstance(listed, list) else []


def _cell_key(fact: Mapping[str, Any]) -> tuple[int, int] | None:
    cell = fact.get("cell")
    if isinstance(cell, list) and len(cell) == 2 and all(isinstance(n, int) for n in cell):
        return cell[0], cell[1]
    return None


def store_facts(shape: "Shape", found: Iterable[Mapping[str, Any]]) -> bool:
    """Keep ``found`` as the shape's facts, replacing what it had; none drops the tag."""
    found = [dict(fact) for fact in found]
    return store(shape, {"v": _records.VERSION, "facts": found} if found else None,
                 FACTS_TAG)


def facts_written(shape: "Shape", cell: tuple[int, int] | None, before: str, after: str,
                  new: Iterable[Mapping[str, Any]] = ()) -> bool:
    """The text of ``shape`` (or its ``cell``) went from ``before`` to ``after``: its facts
    the edit did not touch stay, moved; ``new`` (facts with their ``p`` and ``at`` in
    ``after``) are added where their text is.  Whether anything changed."""
    kept = facts(shape)
    others = [fact for fact in kept if _cell_key(fact) != cell]
    mine = [{k: v for k, v in fact.items() if k != "cell"} for fact in kept
            if _cell_key(fact) == cell]
    if not mine and not new:
        return False
    survivors = _records.facts_after_edit(mine, before, after)
    added = _records.facts_found(list(new), after)
    places = {(fact["p"], fact["at"]) for fact in added}
    merged = sorted([fact for fact in survivors if (fact["p"], fact["at"]) not in places]
                    + added, key=lambda fact: (fact["p"], fact["at"]))
    if cell is not None:
        merged = [{**fact, "cell": [cell[0], cell[1]]} for fact in merged]
    return store_facts(shape, others + merged)


def cell_facts_after(shape: "Shape", rule: str, index: int | None = None,
                     cells: Iterable[tuple[int, int]] = ()) -> bool:
    """A table's facts after a row or column was inserted (``insert_row``,
    ``insert_column``: the cells at or after ``index`` move one on) or deleted
    (``delete_row``, ``delete_column``: its cells' facts go, the ones after move back), or
    after ``cells`` were rewritten some other way (``cells``: their facts go)."""
    kept = facts(shape)
    if not kept:
        return False
    gone = set(cells)
    out = []
    for fact in kept:
        key = _cell_key(fact)
        if key is None:
            out.append(fact)
            continue
        row, col = key
        if key in gone:
            continue
        if rule in ("insert_row", "delete_row"):
            if rule == "delete_row" and row == index:
                continue
            if index is not None and row >= index + (rule == "delete_row"):
                row += 1 if rule == "insert_row" else -1
        elif rule in ("insert_column", "delete_column"):
            if rule == "delete_column" and col == index:
                continue
            if index is not None and col >= index + (rule == "delete_column"):
                col += 1 if rule == "insert_column" else -1
        out.append({**fact, "cell": [row, col]})
    return store_facts(shape, out) if out != kept else False


def facts_found(shape: "Shape") -> list[dict[str, Any]]:
    """The shape's facts whose text is still where they say."""
    out = []
    by_text: dict[Any, list[dict[str, Any]]] = {}
    for fact in facts(shape):
        by_text.setdefault(_cell_key(fact), []).append(fact)
    for cell, listed in by_text.items():
        text = text_of(shape, cell)
        if text is not None:
            out += _records.facts_found(listed, text)
    return out


# -- the deck's records -------------------------------------------------------------------------


def records(document: "Document") -> dict[str, dict[str, Any]]:
    """``{shape address: record}`` for every graphic frame on the deck's slides that keeps
    one -- in its tags part, or, for a chart frame without the tag, in its chart part
    (:func:`ooxml_edit.charts.provenance`) -- and every shape or table with text facts
    still true (``facts``, beside a table's own record).  Unreadable records are skipped
    with a :class:`ProvenanceWarning`."""
    out: dict[str, dict[str, Any]] = {}
    for slide in document.slides:
        for shape in slide.shapes:
            found = None
            if shape._element.tag == qn("p:graphicFrame"):
                found = read(shape)
                if found is None:
                    found = _chart_record(shape)
            if _tags_part(shape) is not None:
                listed = facts_found(shape)
                if listed:
                    found = {**(found or {"v": _records.VERSION}), "facts": listed}
            if found is not None:
                out[shape.id] = found
    return out


def _chart_record(shape: "Shape") -> dict[str, Any] | None:
    chart = shape._element.find(".//" + qn("c:chart"))
    if chart is None or not chart.get(_R_ID):
        return None
    package = shape._slide.document.package
    part = package.related_part(shape._slide.part_path, chart.get(_R_ID))
    if part is None or not package.has_part(part):
        return None
    from ooxml_edit.charts import provenance as chart_provenance

    try:
        return chart_provenance(package, part)
    except _records.RecordError as error:
        warnings.warn(f"{shape.id}: its chart part holds no source record ({error}); it is "
                      "skipped", ProvenanceWarning, stacklevel=2)
        return None


# -- table rules --------------------------------------------------------------------------------
#
# A table record (ppt_add_table, ppt_edit_table's set_cells from) says: the header row
# (``header_row``, absent without one), the first placed source row's table row
# (``first_row``; the n-th source row placed is table row first_row + n - 1: source row n
# when every row is placed, as a table continued on the next slides places some), and per
# placed column its table column (``col``) and source rows (``rows``, runs from 1).  The
# block is the rows from the header (or first_row) to the last placed row, in the
# entries' columns.  The rules keep
# that true or drop what it no longer says: a stale address is worse than none.


def _block_rows(found: Mapping[str, Any]) -> tuple[int, int] | None:
    """The block's first and last table row, ``None`` when the record does not say."""
    first = found.get("first_row")
    if not isinstance(first, int):
        return None
    top = found.get("header_row")
    top = top if isinstance(top, int) else first
    last = first - 1
    for entry in found.get("columns") or []:
        last = max(last, first + len(_numbers(entry.get("rows"))) - 1)
    return min(top, first), max(last, top)


def _numbers(runs: Any) -> list[int]:
    out: list[int] = []
    for part in str(runs or "").split(","):
        match = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+)\s*)?", part)
        if match:
            start = int(match.group(1))
            out += list(range(start, int(match.group(2) or start) + 1))
    return out


def _with_columns(found: Mapping[str, Any], entries: list[dict[str, Any]]):
    if not entries:
        return None
    return {**found, "columns": entries}


def after_cells(found: Mapping[str, Any] | None,
                cells: Iterable[tuple[int, int]]) -> dict[str, Any] | None:
    """The record after the text of ``cells`` (``(row, col)``, from 0) was written without
    a source: a written cell in the block takes its column's entry away; no entry left, no
    record.  A record the rules cannot read is dropped."""
    if found is None:
        return None
    span = _block_rows(found)
    if span is None:
        return None
    touched = {col for row, col in cells if span[0] <= row <= span[1]}
    if not touched:
        return dict(found)
    return _with_columns(found, [dict(entry) for entry in found.get("columns") or []
                                 if entry.get("col") not in touched])


def after_row_insert(found: Mapping[str, Any] | None, index: int) -> dict[str, Any] | None:
    """A row inserted at ``index``: above the block (or at its top) the block moves down
    one; inside it the record is dropped (its rows are no longer one run); below it nothing
    changes."""
    return _rows_moved(found, index, +1)


def after_row_delete(found: Mapping[str, Any] | None, index: int) -> dict[str, Any] | None:
    """Row ``index`` deleted: above the block, the block moves up one; inside it (the
    header included) the record is dropped; below it nothing changes."""
    return _rows_moved(found, index, -1)


def _rows_moved(found, index: int, step: int):
    if found is None:
        return None
    span = _block_rows(found)
    if span is None:
        return None
    top, last = span
    if index > last:
        return dict(found)
    if index < top or (step > 0 and index == top):
        out = dict(found)
        for key in ("header_row", "first_row"):
            if isinstance(out.get(key), int):
                out[key] += step
        return out
    return None


def after_column_insert(found: Mapping[str, Any] | None, index: int) -> dict[str, Any] | None:
    """A column inserted at ``index``: the entries at or right of it move right one."""
    if found is None:
        return None
    return _with_columns(found, [{**entry, "col": entry["col"] + 1}
                                 if isinstance(entry.get("col"), int) and entry["col"] >= index
                                 else dict(entry) for entry in found.get("columns") or []])


def after_column_delete(found: Mapping[str, Any] | None, index: int) -> dict[str, Any] | None:
    """Column ``index`` deleted: its entry goes, the ones right of it move left one; no
    entry left, no record."""
    if found is None:
        return None
    entries = []
    for entry in found.get("columns") or []:
        col = entry.get("col")
        if col == index:
            continue
        entries.append({**entry, "col": col - 1} if isinstance(col, int) and col > index
                       else dict(entry))
    return _with_columns(found, entries)


def cell_of(address: str) -> tuple[str, int, int] | None:
    """``(frame address, row, col)`` for a cell address (``256.7/cell1,2``, or anything
    inside it), ``None`` for anything else."""
    match = re.match(r"^([^/]+)/cell(\d+),(\d+)(?:/|$)", str(address))
    if match is None:
        return None
    return match.group(1), int(match.group(2)), int(match.group(3))


__all__ = ["CT_TAGS", "FACTS_TAG", "REL_TAGS", "TAG_NAME", "ProvenanceWarning", "after_cells",
           "after_column_delete", "after_column_insert", "after_row_delete",
           "after_row_insert", "cell_facts_after", "cell_of", "facts", "facts_found",
           "facts_written", "read", "records", "store", "store_facts", "text_of"]
