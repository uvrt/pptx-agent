"""Empty paragraphs left between text (production feedback round 18, Q4): ``check``'s
``empty_paragraphs``, with ``include: ["layout"]``.

Blank lines used for spacing inside a text box or a table cell drift from a sample's
spacing, and nothing else reports them: they fit, and they draw nothing.  Every run of them
with text on both sides, in one text body, is listed in that body's one entry::

    {"kind": "empty_paragraphs", "at": "256.7/cell1,2", "paragraphs": [1],
     "fix": {"tool": "ppt_set_text", "items": [{"target": "256.7/cell1,2",
                                                "text": "Title\\nBody text"}]}}

``at`` is the shape or cell as the text tools address it, ``paragraphs`` 0-based as
``256.5/p1`` counts them.  A paragraph is empty when it shows nothing: no text but
whitespace, and no line break, field, equation or anything else beside its runs.  Leading
and trailing empties (an anchor's padding, a placeholder's one paragraph) are not listed.
``fix`` writes the text back without them: :meth:`TextFrame.set_text
<pptx_agent.TextFrame.set_text>` diffs paragraphs first, so every other paragraph is left
as it is, formatting and all.  A cell merged into another holds no text of its own and is
passed over."""

from __future__ import annotations

from typing import Any

from ..oxml.xml import qn
from .text import _paragraph_text

#: What an empty paragraph may hold: its properties, its end mark and runs of whitespace.
_ALLOWED = (qn("a:pPr"), qn("a:endParaRPr"))
_RUN, _TEXT = qn("a:r"), qn("a:t")


def empty_paragraphs(slide) -> list[dict[str, Any]]:
    """The slide's runs of empty paragraphs between text, its shapes in document order (a
    group's members after it), a table's cells row by row."""
    found: list[dict[str, Any]] = []
    for shape in slide.shapes:
        if shape.has_table:
            for cell in shape.table.iter_cells():
                if not cell.is_spanned:
                    found += _runs(cell.address, cell.text_frame)
        elif shape.text_frame is not None:
            found += _runs(shape.id, shape.text_frame)
    return found


def _empty(paragraph) -> bool:
    for child in paragraph:
        if not isinstance(child.tag, str) or child.tag in _ALLOWED:
            continue
        if child.tag != _RUN:
            return False                       # a break, a field, an equation...
        for node in child:
            if node.tag == _TEXT and (node.text or "").strip():
                return False
    return True


def _runs(address: str, frame) -> list[dict[str, Any]]:
    """The text body's entry, every run in it listed (one entry, so its fix is the one
    text): ``[]`` where it has none."""
    paragraphs = frame._paragraph_elements()
    empty = [_empty(p) for p in paragraphs]
    removed: list[int] = []
    k = 0
    while k < len(empty):
        if not empty[k]:
            k += 1
            continue
        start = k
        while k < len(empty) and empty[k]:
            k += 1
        if start and k < len(empty):
            removed += range(start, k)
    if not removed:
        return []
    text = "\n".join(_paragraph_text(p) for n, p in enumerate(paragraphs) if n not in removed)
    return [{"kind": "empty_paragraphs", "at": address, "paragraphs": removed,
             "fix": {"tool": "ppt_set_text", "items": [{"target": address, "text": text}]}}]
