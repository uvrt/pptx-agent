"""Slide numbers written by reference, kept current as slides are added, deleted and moved.

``{{slide:s:262}}`` in a text tool's text (``{{slide:$intro}}`` with a ref) is written as
that slide's number now, by the tool layer (:mod:`ooxml_edit.tools.textfacts`' own
references, ``DocumentFormat.own_facts``), and kept on its shape as a text fact like a
dataset value (:mod:`.provenance`, the ``UVRT_TEXT_FACTS`` tag): ``source`` ``slide:s:262``,
``slide`` the id, ``text`` the number written, ``p`` and ``at`` its place.

**The number** is the one the slide's own slide-number field shows: its position in the
slide list plus ``p:presentation@firstSlideNum`` minus one (1 when absent), so 1, 2, 3 ...
in a deck that states none.  Measured for the field (pptx2svg's ``resolve/fields.py``,
PowerPoint 16 for Mac, 2026-10-10: 5, 6, ... 12 for eight slides with
``firstSlideNum="5"``).  Hidden slides count: the field numbers by position in the slide
list, hidden or not, so a hidden slide keeps its number and the ones after it keep theirs.

**Kept current.**  At the end of every outermost :meth:`Document.batch
<pptx_agent.Document.batch>` -- every tool call, and every library call that adds,
deletes, moves or duplicates slides, since those run in one -- the deck compares its
slide order (and ``firstSlideNum``) with the one it last numbered for, and when it
changed rewrites every slide fact whose number is no longer its slide's (:func:`keep_current`),
in the same undo step: undo puts the order and the numbers back together.  Saving through
:meth:`Document.save <pptx_agent.Document.save>`, :meth:`save_as_template
<pptx_agent.Document.save_as_template>` or the ``save_document`` tool rewrites them whatever
the order did (:meth:`Document.update_slide_numbers <pptx_agent.Document.update_slide_numbers>`),
which also catches a deck PowerPoint reordered since.  The text is rewritten the way a text
tool writes it -- the characters around it keep their formatting -- and the deck's other
facts in that text move with it.

**A slide that is gone.**  A fact whose slide was deleted keeps its last number, and is
marked ``"gone": true`` so that a slide given the same id later (a reopened deck may hand
out a deleted top id again) is never taken for it.  ``check`` and ``save_document`` list it
under ``missing_slides`` ("slide s:262 no longer exists"); rewrite the text, or delete it.

**The limit.**  PowerPoint has no field that shows *another* slide's number, so the number
is plain text to PowerPoint: a deck reordered in PowerPoint shows stale numbers until this
library next opens it and saves it (or changes its slide order).  A fact whose text was
edited is dropped as every fact is (:func:`ooxml_edit.provenance.facts_found`).  Speaker
notes and SmartArt nodes take the number but keep no fact, so it is not kept current there.
Shapes copied to another deck lose their slide facts (:func:`drop`): the ids name the source
deck's slides.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, Iterable, Mapping

from ..oxml.xml import qn
from . import provenance as _provenance
from . import slides as _slides

if TYPE_CHECKING:
    from .document import Document, Shape, Slide

#: The prefix of a slide-number reference, and of its fact's ``source``.
PREFIX = "slide:"
#: What ``check`` and ``save_document`` say of a reference whose slide is gone.
GONE = "slide s:{id} no longer exists: the text keeps its last number, {text}"


def first_number(document: "Document") -> int:
    """``p:presentation@firstSlideNum``: the first slide's number, 1 when absent."""
    _, root = _slides._presentation(document.package)
    value = root.get("firstSlideNum")
    return int(value) if value is not None and value.lstrip("-").isdigit() else 1


def order(document: "Document") -> tuple[int, tuple[int, ...]]:
    """What the numbers follow: the first number and the slide ids in order."""
    _, root = _slides._presentation(document.package)
    ids = tuple(int(node.get("id")) for node in _slides._entries(root)
                if (node.get("id") or "").isdigit())
    return first_number(document), ids


def numbers(document: "Document") -> dict[int, int]:
    """``{sldId: the number its slide-number field shows}``."""
    first, ids = order(document)
    return {sid: first + n for n, sid in enumerate(ids)}


def is_slide_fact(fact: Mapping[str, Any]) -> bool:
    return str(fact.get("source", "")).startswith(PREFIX) and \
        isinstance(fact.get("slide"), int) and not isinstance(fact.get("slide"), bool)


def _hosts(document: "Document") -> "list[Slide]":
    """Every slide, layout and master: where a shape with text facts may be."""
    out: list = list(document.slides)
    for template in [*document.masters, *document.layouts]:
        out.append(document._template_sheet(template.part_path))
    return out


def tagged(host: "Slide") -> "list[Shape]":
    """The shapes on ``host`` that relate to a tags part (``p:nvPr/p:custDataLst/p:tags``)."""
    out = []
    for node in host._sp_tree().iter(qn("p:tags")):
        custom = node.getparent()
        nv_pr = custom.getparent() if custom is not None else None
        holder = nv_pr.getparent() if nv_pr is not None else None
        element = holder.getparent() if holder is not None else None
        if element is not None and custom.tag == qn("p:custDataLst") and \
                nv_pr.tag == qn("p:nvPr"):
            out.append(host._wrap(element))
    return out


def _address(shape: "Shape", cell: tuple[int, int] | None) -> str:
    return shape.id if cell is None else f"{shape.id}/cell{cell[0]},{cell[1]}"


def _cell(fact: Mapping[str, Any]) -> tuple[int, int] | None:
    return _provenance._cell_key(fact)


def keep_current(document: "Document") -> list[dict[str, Any]]:
    """Rewrite every slide fact whose number is not its slide's now; mark the ones whose
    slide is gone.  Each change: ``{"shape", "slide", "text", "was"}``, or
    ``{"shape", "slide", "text", "gone": true}`` for a slide newly found gone.  Run inside
    the deck's undo step (:meth:`Document.batch`)."""
    now = numbers(document)
    changes: list[dict[str, Any]] = []
    for host in _hosts(document):
        for shape in tagged(host):
            by_text: dict[Any, list[dict[str, Any]]] = {}
            for fact in _provenance.facts_found(shape):
                by_text.setdefault(_cell(fact), []).append(fact)
            for cell, listed in by_text.items():
                changes += _renumber(shape, cell, listed, now)
    return changes


def _renumber(shape: "Shape", cell: tuple[int, int] | None, listed: list[dict[str, Any]],
              now: Mapping[int, int]) -> list[dict[str, Any]]:
    from ooxml_edit import provenance as _records

    before = _provenance.text_of(shape, cell)
    if before is None:
        return []
    placed = sorted(((_records.offset(before, fact["p"], fact["at"]), fact) for fact in listed),
                    key=lambda item: item[0])
    changes: list[dict[str, Any]] = []
    pieces: list[str] = []
    new: list[dict[str, Any]] = []
    at = shift = 0
    for start, fact in placed:
        fact = {k: v for k, v in fact.items() if k != "cell"}
        text = fact["text"]
        if is_slide_fact(fact) and not fact.get("gone"):
            number = now.get(fact["slide"])
            if number is None:
                fact["gone"] = True
                changes.append({"shape": _address(shape, cell), "slide": f"s:{fact['slide']}",
                                "text": text, "gone": True})
            elif str(number) != text:
                changes.append({"shape": _address(shape, cell), "slide": f"s:{fact['slide']}",
                                "text": str(number), "was": text})
                text = str(number)
        if start < at:                         # overlapping facts: the text says which
            return []
        pieces.append(before[at:start])
        pieces.append(text)
        at = start + len(fact["text"])
        p, place = _records.place("".join(pieces), start + shift)
        new.append({**fact, "text": text, "p": p, "at": place})
        shift += len(text) - len(fact["text"])
    if not changes:
        return []
    after = "".join(pieces) + before[at:]
    if after != before:
        frame = shape.text_frame if cell is None else shape.table.cell(*cell).text_frame
        frame.set_text(after)
        after = _provenance.text_of(shape, cell)
    _provenance.facts_written(shape, cell, before, after, new)
    return changes


def missing(document: "Document", slide_ids: Iterable[int] | None = None
            ) -> list[dict[str, str]]:
    """The slide facts on ``slide_ids`` (default: every slide, layout and master) whose
    slide is not in the deck: ``{"shape", "slide", "text", "message"}``."""
    now = numbers(document)
    wanted = None if slide_ids is None else set(slide_ids)
    hosts = _hosts(document) if wanted is None else \
        [slide for slide in document.slides if slide.slide_id in wanted]
    out = []
    for host in hosts:
        for shape in tagged(host):
            for fact in _provenance.facts_found(shape):
                if is_slide_fact(fact) and (fact.get("gone") or fact["slide"] not in now):
                    out.append({"shape": _address(shape, _cell(fact)),
                                "slide": f"s:{fact['slide']}", "text": fact["text"],
                                "message": GONE.format(id=fact["slide"], text=fact["text"])})
    return out


def drop(shapes: "Iterable[Shape]") -> None:
    """Take the slide facts off ``shapes`` (and their members): copied to another deck,
    the ids they name are another deck's.  The text stays as written."""
    for shape in shapes:
        for member in [shape, *_members(shape)]:
            kept = _provenance.facts(member)
            if any(is_slide_fact(fact) for fact in kept):
                _provenance.store_facts(member, [fact for fact in kept
                                                 if not is_slide_fact(fact)])


def _members(shape: "Shape") -> "list[Shape]":
    out = []
    for child in shape.children:
        out += [child, *_members(child)]
    return out


__all__ = ["GONE", "PREFIX", "drop", "first_number", "is_slide_fact", "keep_current",
           "missing", "numbers", "order", "tagged"]
