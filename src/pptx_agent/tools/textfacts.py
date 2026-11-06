"""Text facts in the deck tools: a value written by reference is kept on its shape.

The dispatcher writes ``{{d:by_topic/door/mean}}`` as the value before a text tool runs
(:mod:`ooxml_edit.tools.textfacts`; ``call.facts`` says where in each argument).  The tool
then says where that is in the deck: :class:`Spot` is the text a target writes into -- a
shape's, or a table cell's (the table's frame keeps the facts) -- and
:meth:`Spot.written` keeps the new facts and what the edit left of the old ones
(:func:`pptx_agent.edit.provenance.facts_written`), in the call's undo step.  Speaker notes
and SmartArt nodes take the values but keep no facts.

A slide's number by reference, ``{{slide:s:262}}`` (``{{slide:$intro}}``), is the deck's own
reference (``DocumentFormat.own_facts``): :func:`slide_number` writes it, as the slide's
number field shows it, and the deck keeps it current (:mod:`pptx_agent.edit.slidenumbers`).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Mapping

from ooxml_edit import provenance as _records

from ooxml_edit.tools import NumberText, ToolError

from ..edit import provenance as _provenance
from ..edit import slidenumbers as _slidenumbers

#: A text address: a shape, a table cell, a paragraph, a run.
_ADDRESS = re.compile(r"^([^/]+)(?:/cell(\d+),(\d+))?(?:/p(\d+)(?:/r(\d+))?)?$")


@dataclass
class Spot:
    """Where a target's text is: ``shape`` (a text shape, or a table's frame with
    ``cell``), and the paragraph (``p``) or run (``r``) the target is, if it is one."""

    shape: Any
    cell: tuple[int, int] | None = None
    p: int | None = None
    r: int | None = None

    def text(self) -> str | None:
        return _provenance.text_of(self.shape, self.cell)

    def base(self, after: str) -> int | None:
        """Where the target's own text starts in ``after`` (the whole text): 0 for a shape
        or cell, a paragraph's start; ``None`` for a run, whose start is not kept."""
        if self.r is not None:
            return None
        if self.p is None:
            return 0
        return _records.offset(after, self.p, 0)

    def written(self, before: str | None, placed: Iterable[tuple[int, Any]] = ()) -> bool:
        """The target's text was written (``before`` the whole text it had): keep the
        facts ``placed`` (``(offset in the target's text, Fact)``) and what the edit left
        of the old ones."""
        after = self.text()
        if before is None or after is None:
            return False
        placed = list(placed)
        base = self.base(after) if placed else None
        new = []
        if base is not None:
            for where, fact in placed:
                p, at = _records.place(after, base + where)
                new.append({**fact.record, "p": p, "at": at})
        return _provenance.facts_written(self.shape, self.cell, before, after, new)


def spot(deck, address: Any) -> Spot | None:
    """The :class:`Spot` of a text address (``256.5``, ``256.5/p1``, ``256.7/cell1,2``,
    ``L2.5``), or ``None`` -- speaker notes, a SmartArt node, a shape without text."""
    match = _ADDRESS.match(str(address or ""))
    if match is None:
        return None
    try:
        shape = deck.shape(match.group(1))
    except (KeyError, IndexError, ValueError, LookupError):
        return None
    cell = (int(match.group(2)), int(match.group(3))) if match.group(2) else None
    if cell is None and shape.text_frame is None:
        return None
    if cell is not None and getattr(shape, "table", None) is None:
        return None
    p = int(match.group(4)) if match.group(4) else None
    r = int(match.group(5)) if match.group(5) else None
    return Spot(shape, cell, p, r)


def placed(call, field: str, item: Mapping[str, Any]) -> list[tuple[int, Any]]:
    """The facts written into an item's text (``field``: ``items[0]``), each with its
    offset in the text the item writes: ``text`` as it is, ``paragraphs`` joined by
    ``\\n``, each paragraph its runs' texts one after another."""
    out = [(fact.start, fact) for fact in call.facts_at(f"{field}.text")]
    if "paragraphs" in item and isinstance(item["paragraphs"], list):
        start = 0
        for j, paragraph in enumerate(item["paragraphs"]):
            runs = (paragraph.get("runs") or []) if isinstance(paragraph, Mapping) else []
            at = start
            for k, run in enumerate(runs):
                text = run.get("text", "") if isinstance(run, Mapping) else ""
                out += [(at + fact.start, fact)
                        for fact in call.facts_at(f"{field}.paragraphs[{j}].runs[{k}].text")]
                at += len(text)
            start = at + 1
    return out


def cell_facts(call, field: str) -> list[tuple[int, Any]]:
    """The facts written into one cell's text argument (``data[1][2]``,
    ``cells[0].text``)."""
    return [(fact.start, fact) for fact in call.facts_at(field)]


def slide_number(deck, name: str) -> tuple[str, dict[str, Any]]:
    """What ``{{slide:<name>}}`` writes: the number of the slide ``name`` (``s:262``, or the
    address a ``$ref`` named) as its slide-number field shows it, and the fact's ``slide``
    (its id).  A bare number, a shape or a slide that is not there is refused, listing the
    slides."""
    text = str(name).strip()
    now = _slidenumbers.numbers(deck)
    first, ids = _slidenumbers.order(deck)
    options = [f"s:{sid}" for sid in ids][:50]
    head = text[2:] if text.startswith("s:") else None
    if head is None or not head.isdigit():
        said = f"{{{{slide:{text}}}}} names no slide: write {{{{slide:s:<id>}}}} or a $ref to a slide"
        if text.isdigit() and 1 <= int(text) - first + 1 <= len(ids):
            sid = ids[int(text) - first]
            said = (f"{{{{slide:{text}}}}}: a slide reference names the slide, not its number "
                    f"(that is what it writes); slide {text} is s:{sid}: write "
                    f"{{{{slide:s:{sid}}}}}")
        elif "." in text:
            said = (f"{{{{slide:{text}}}}}: {text} is a shape; a slide reference names a "
                    f"slide: {{{{slide:s:{text.split('.')[0]}}}}}")
        raise ToolError("invalid_arguments", said, valid_options=options)
    sid = int(head)
    if sid not in now:
        raise ToolError("not_found", f"{{{{slide:s:{sid}}}}}: the deck has no slide s:{sid}",
                        valid_options=options)
    return str(now[sid]), {"slide": sid}


__all__ = ["Spot", "cell_facts", "placed", "slide_number", "spot"]


# -- numbers no value holds (round 36) ------------------------------------------------------------

#: Placeholders whose text is the slide's furniture, never a figure: footer, date, number.
FURNITURE = ("ftr", "dt", "sldNum")


def number_texts(call, scope: Mapping[str, Any]) -> list[NumberText]:
    """``DocumentFormat.number_texts`` for a deck (``check``'s ``include: ["numbers"]``,
    :mod:`ooxml_edit.tools.numbers`): each shape's text and each table cell's on the
    slides in ``scope`` (``slides``; every slide without), the spans written by reference
    marked (the shape's or table's kept facts that are still true), in slide order.  Left
    out: a table or chart placed with ``from`` (its frame keeps the record), charts and
    SmartArt (their text is no paragraph of the slide), footers, dates and slide numbers,
    speaker notes."""
    from .common import numbers as slide_numbers

    if scope.get("pages") is not None:
        raise ToolError("invalid_arguments", "a deck has slides, not pages: give slides",
                        field="pages")
    if scope.get("range") is not None:
        raise ToolError("invalid_arguments", "check range names a document's blocks (docx); "
                        "leave it out here: slides say which slides", field="range")
    deck = call.document
    out: list[NumberText] = []
    for number in slide_numbers(deck, scope.get("slides")):
        slide = deck.slides[number - 1]
        where = f"s:{slide.slide_id}"
        for shape in slide.shapes:
            placeholder = shape.placeholder
            if placeholder is not None and placeholder.type in FURNITURE:
                continue
            table = shape.table if shape.kind == "graphic_frame" else None
            if table is not None:
                if _provenance.read(shape) is not None:
                    continue                          # placed with from: it has its record
                kept = _provenance.facts_found(shape)
                for row in range(table.rows):
                    for column in range(table.columns):
                        text = _provenance.text_of(shape, (row, column)) or ""
                        if text.strip():
                            spans = _spans(text, [fact for fact in kept
                                                  if fact.get("cell") == [row, column]])
                            out.append(NumberText(f"{shape.id}/cell{row},{column}", text, spans,
                                                  {"slides": [where]}, {"slide": where}))
                continue
            if shape.text_frame is None:
                continue
            text = _provenance.text_of(shape) or ""
            if not text.strip():
                continue
            spans = _spans(text, [fact for fact in _provenance.facts_found(shape)
                                  if "cell" not in fact])
            out.append(NumberText(shape.id, text, spans, {"slides": [where]}, {"slide": where}))
    return out


def _spans(text: str, facts: Iterable[Mapping[str, Any]]) -> tuple[tuple[int, int], ...]:
    """Where in ``text`` the facts' values are: ``(start, end)`` each."""
    out = []
    for fact in facts:
        start = _records.offset(text, fact.get("p"), fact.get("at"))
        if start is not None and isinstance(fact.get("text"), str):
            out.append((start, start + len(fact["text"])))
    return tuple(out)
