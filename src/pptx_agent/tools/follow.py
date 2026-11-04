"""``ppt_follow_sample``: copies of a sample slide made to follow its revised design in one
call (round 31, R8) -- :mod:`pptx_agent.edit.follow` at the tool boundary.

What it takes, as the other slide tools do: the sample as ``s:300``, a ``$ref``, a slide
number or a shape's address on it (``300.4``: its slide, with a warning); ``sample_doc`` for
a sample in another open deck; ``copies`` as slide numbers, ids or refs, none (or ``[]``,
``"all"``) for every slide made from the sample.  A copy that is the sample is skipped with
a warning.  Empty strings and ``false`` where a strict decoder fills in fields are the
defaults.  A copy already following is ``changed: false``, never an error.  Every
``not_followed`` entry has a ``kind`` (:data:`~pptx_agent.edit.follow.KINDS`): ``own``,
``changed_by_copy``, ``structure``; ``report: "actionable"`` leaves the ``own`` ones out,
counted per copy (``own``).  The reply keeps to a size budget: past it the changed shapes
become counts per copy, then the ``own`` entries, and copies that still do not fit are
listed under ``left_out`` with the call that shows them.
"""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, string, tool

from ..edit.errors import ArgumentError
from ..edit.follow import OWN, REMOVE, find_copies, follow_sample
from .common import KIND, library_errors, numbers, slide as slide_at
from .slides import SLIDES

#: The reply's budget, characters of its data, when the session's limit is larger.
BUDGET = 12_000
#: What ``report`` takes: every ``not_followed`` entry, or those without kind ``own``.
REPORT = ("all", "actionable")


def _documents(arguments) -> list[str]:
    """The deck changed, and the sample's deck when it is another (locked, only read)."""
    docs = [arguments["doc"]] if isinstance(arguments.get("doc"), str) else []
    other = arguments.get("sample_doc")
    if isinstance(other, str) and other and other not in docs:
        docs.append(other)
    return docs


def _slide_of(deck, address, field: str, warnings: list[str]):
    """A slide from what another slide tool takes: ``s:300``, ``300``, a position (an
    integer, or ``"#3"``), or a shape's address on it (``300.4``, with a warning)."""
    if isinstance(address, int):
        return deck.slides[numbers(deck, [address], field=field)[0] - 1]
    text = str(address).strip()
    if text.startswith("#") and text[1:].isdigit():
        return deck.slides[numbers(deck, [int(text[1:])], field=field)[0] - 1]
    head = text.split("/")[0]
    if "." in head:
        slide_part = head.split(".")[0]
        found = slide_at(deck, slide_part, field=field)
        warnings.append(f"{field}: {text} is a shape; its slide s:{found.slide_id} is meant")
        return found
    return slide_at(deck, text, field=field)


def _copies(deck, copies, sample, warnings: list[str]):
    if copies in (None, "", "all", "found") or copies == []:
        if copies == []:
            warnings.append("copies: [] read as every slide made from the sample")
        return None
    if isinstance(copies, (str, int)):
        copies = [copies]
    chosen = []
    for index, item in enumerate(copies):
        if item in ("", None):
            continue
        slide = _slide_of(deck, item, f"copies[{index}]", warnings)
        if slide.document is sample.document and slide.slide_id == sample.slide_id:
            warnings.append(f"copies[{index}]: s:{sample.slide_id} is the sample; skipped")
            continue
        if all(slide.slide_id != other.slide_id for other in chosen):
            chosen.append(slide)
    if not chosen:
        raise ToolError("invalid_arguments", f"no copy named but the sample s:{sample.slide_id}: "
                        "list its copies, e.g. [\"s:301\", \"s:302\"], or leave copies out to "
                        "follow every slide made from it", field="copies",
                        valid_options=_candidates(deck, sample))
    return chosen


def _candidates(deck, sample) -> list[str]:
    """Slides a caller may mean: the copies found, those near, then the sample's layout's."""
    out: list[str] = []
    try:
        found = find_copies(sample, deck)
        out += found.slides + [entry["slide"] for entry in found.near]
    except ArgumentError:
        pass
    layout = sample.layout
    for slide in deck.slides:
        address = f"s:{slide.slide_id}"
        if slide.layout is not None and layout is not None and \
                slide.layout.name == layout.name and address not in out and \
                not (slide.document is sample.document and slide.slide_id == sample.slide_id):
            out.append(address)
    return out[:50]


def _actionable(entry: dict) -> bool:
    """Whether a copy's entry has something to handle: a ``not_followed`` entry other than
    the copy's own content, or an overflow."""
    return bool(entry.get("overflows")) or any(
        item.get("kind") != OWN for item in entry.get("not_followed", ()))


def _without_own(data: dict) -> int:
    """``data``'s ``not_followed`` entries of kind ``own`` taken out, counted per copy as
    ``own``; how many were taken out."""
    total = 0
    for entry in data["copies"]:
        items = entry.get("not_followed")
        if not items:
            continue
        kept = [item for item in items if item.get("kind") != OWN]
        count = len(items) - len(kept)
        if not count:
            continue
        total += count
        if kept:
            entry["not_followed"] = kept
        else:
            del entry["not_followed"]
        entry["own"] = count
    return total


def _fit(data: dict, budget: int, brief_json, own: bool = True) -> tuple[dict, list[str]]:
    """``data`` within ``budget`` characters: first the changed shapes as counts, then the
    ``own`` entries as counts, then copies left out (listed), never cut silently."""
    notes: list[str] = []
    if len(json.dumps(data, ensure_ascii=False)) <= budget:
        return data, notes
    data = brief_json
    notes.append("changed shapes are counted per copy to keep the reply small; name a copy "
                 "in copies with preview: true to see its shapes")
    if own and len(json.dumps(data, ensure_ascii=False)) > budget and _without_own(data):
        notes.append("not_followed entries of kind own (the copy's own content, kept) are "
                     "counted per copy under own; name a copy in copies with preview: true "
                     "to list them")
    entries = data["copies"]
    kept: list[dict] = []
    size = len(json.dumps({**data, "copies": []}, ensure_ascii=False)) + 200
    left: list[str] = []
    # copies needing attention first
    ordered = sorted(entries, key=lambda e: not _actionable(e))
    for entry in ordered:
        cost = len(json.dumps(entry, ensure_ascii=False)) + 2
        if size + cost <= budget:
            kept.append(entry)
            size += cost
        else:
            left.append(entry["slide"])
    if left:
        data["left_out"] = left
        notes.append(f"{len(left)} copies did not fit and are under left_out: call with "
                     "copies naming some of them and preview: true to read theirs")
    order = {entry["slide"]: n for n, entry in enumerate(entries)}
    data["copies"] = sorted(kept, key=lambda e: order[e["slide"]])
    return data, notes


@tool("ppt_follow_sample",
      "Make copies of a sample slide follow its revised design: box, shapes added or "
      "removed, fill, outline, tables, text formatting; each copy keeps its text, values "
      "and pictures. Per copy: what changed, what could not follow and why, new "
      "overflows. One undo step.",
      {"doc": string("Document id: the copies' deck."),
       "sample": string("The sample slide, s:300 or $ref.", or_integer=1),
       "copies": array(string(or_integer=1), "Default every slide made from it (its "
                       "layout, its shapes' ids and names).",
                       optional=True),
       "sample_doc": string("The sample's deck when another, d2.", optional=True),
       "preview": boolean("Report only, change nothing.", optional=True),
       "remove": string("Copy shapes the sample lacks: auto (those most copies have) "
                        "deleted, all, or none.", enum=list(REMOVE), aliases={"": "auto"},
                        optional=True),
       "report": string("actionable: not_followed without kind own.",
                        enum=list(REPORT), aliases={"": "all"}, optional=True)},
      kind=KIND, group=SLIDES, mutates=True, refs=("sample", "copies[]"),
      documents=_documents)
@library_errors
def ppt_follow_sample(call, doc, sample, copies=None, sample_doc=None, preview=None,
                      remove=None, report=None):
    deck = call.document
    warnings: list[str] = []
    where = sample_doc or doc
    entry = call.entries.get(where)
    if entry is None or entry.kind != call.entry.kind:
        raise ToolError("invalid_arguments", f"{where} is not an open deck", field="sample_doc",
                        valid_options=[name for name, e in call.entries.items()
                                       if e.kind == call.entry.kind][:50])
    source = _slide_of(entry.document, sample, "sample", warnings)
    chosen = _copies(deck, copies, source, warnings)
    try:
        following = follow_sample(source, chosen, deck=deck, preview=bool(preview),
                                  remove=remove or "auto")
    except ArgumentError as error:
        raise ToolError("not_found" if chosen is None else "invalid_arguments", str(error),
                        field="copies" if chosen is None else "sample",
                        valid_options=_candidates(deck, source)) from None
    if not preview:
        for item in following.copies:
            if item.is_changed:
                call.touch(int(item.slide[2:]))
    budget = min(BUDGET, call.limits.max_result_chars - 2000)
    full, brief = following.to_json(), following.to_json(brief=True)
    actionable = report == "actionable"
    if actionable:
        _without_own(full)
        _without_own(brief)
    data, notes = _fit(full, budget, brief, own=not actionable)
    warnings += notes
    followed = data["followed"]
    attention = sum(1 for c in brief["copies"] if _actionable(c))
    verb = "would follow" if preview else "followed"
    summary = (f"{len(following.copies)} cop{'y' if len(following.copies) == 1 else 'ies'} of "
               f"{following.sample}: {followed} {verb}, {data['unchanged']} already "
               f"follow{'ing' if preview else ''}"
               + (f"; {attention} with something to handle (not_followed, overflows)"
                  if attention else ""))
    changed: list[Any] = [] if preview else [c.slide for c in following.copies if c.is_changed]
    return Result(summary=summary, data=data, warnings=warnings, changed=changed)


TOOLS = [ppt_follow_sample]
