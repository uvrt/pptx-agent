"""Text facts in the deck tools: a value written by reference is kept on its shape.

The dispatcher writes ``{{d:by_topic/door/mean}}`` as the value before a text tool runs
(:mod:`ooxml_edit.tools.textfacts`; ``call.facts`` says where in each argument).  The tool
then says where that is in the deck: :class:`Spot` is the text a target writes into -- a
shape's, or a table cell's (the table's frame keeps the facts) -- and
:meth:`Spot.written` keeps the new facts and what the edit left of the old ones
(:func:`pptx_agent.edit.provenance.facts_written`), in the call's undo step.  Speaker notes
and SmartArt nodes take the values but keep no facts.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable, Mapping

from ooxml_edit import provenance as _records

from ..edit import provenance as _provenance

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


__all__ = ["Spot", "cell_facts", "placed", "spot"]
