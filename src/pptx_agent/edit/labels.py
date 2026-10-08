"""Finding a row, a column, a series or a category by its label -- strictly.

A wrong index writes a plausible number into the wrong place without a word, which was the
end-to-end pilot's main correctness worry (ROADMAP.md, "Usability").  A label either names
exactly one thing or the lookup raises :class:`LabelError`, listing what there is.

Matching is on the text with its whitespace collapsed (line breaks and runs of spaces are
one space, the ends trimmed).  Only when nothing matches that way is the comparison
repeated under Unicode compatibility (NFKC) and case folding -- so ``"売上高(億円)"`` finds
``"売上高（億円）"`` and ``"revenue"`` finds ``"Revenue"`` -- and that second pass must be
unique too.
"""

from __future__ import annotations

import unicodedata
from typing import Iterable, TypeVar

T = TypeVar("T")


class LabelError(KeyError):
    """A label that names nothing, or more than one thing; the message lists the
    candidates.  A :class:`KeyError`, as a missing series name always was.

    The same facts are attributes, for a program to act on: ``label`` (what was asked
    for), ``what`` (``"row"``, ``"column"``, ``"series"``, ``"category"``, ``"title"``),
    ``candidates`` (every label there is, whitespace collapsed, in order) and ``matches``
    (when the label named more than one thing: their positions or objects; else empty)::

        try:
            table.cell_by_label("Revenue", "Q5")
        except pptx_agent.LabelError as error:
            print(error)             # ... no column 'Q5'; the columns are 'Item', 'Q1', ...
            print(error.candidates)  # ['Item', 'Q1', 'Q2', 'Q3', 'Q4']
    """

    def __init__(self, message: str = "", *, label=None, what: str | None = None,
                 candidates=(), matches=()) -> None:
        super().__init__(message)
        self.label = label
        self.what = what
        self.candidates: list[str] = list(candidates)
        self.matches: list = list(matches)

    def __str__(self) -> str:
        return str(self.args[0]) if self.args else ""


def collapse(text) -> str:
    """``text`` with its whitespace collapsed, for comparison."""
    return " ".join(str(text).split())


def _loose(text) -> str:
    return unicodedata.normalize("NFKC", collapse(text)).casefold()


def _plural(what: str) -> str:
    return {"series": "series", "category": "categories"}.get(what, what + "s")


def find_label(label, candidates: Iterable[tuple[T, object]], *, what: str,
               where: str) -> T:
    """The one ``key`` among ``(key, text)`` pairs whose text is ``label``.

    ``what`` names the kind of thing (``"row"``, ``"series"``) and ``where`` the container,
    for the message.
    """
    pairs = list(candidates)
    labels = [collapse(text) for _, text in pairs if text not in (None, "")]
    for normalise in (collapse, _loose):
        wanted = normalise(label)
        hits = [key for key, text in pairs if text is not None and normalise(text) == wanted]
        if len(hits) == 1:
            return hits[0]
        if len(hits) > 1:
            raise LabelError(f"{where}: {label!r} names {len(hits)} {_plural(what)} "
                             f"({', '.join(map(str, hits))}); give the position instead, or "
                             f"make the labels unique", label=label, what=what,
                             candidates=labels, matches=hits)
    listed = ", ".join(repr(text) for text in labels)
    raise LabelError(f"{where}: no {what} {label!r}; the {_plural(what)} are "
                     f"{listed or 'unlabelled'}", label=label, what=what, candidates=labels)


__all__ = ["LabelError", "collapse", "find_label"]
