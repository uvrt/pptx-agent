"""The shared tools' PowerPoint handlers: new, save, find, replace, render, check, properties.

Their definitions are :mod:`ooxml_edit.tools.shared`'s, one per tool for every format; this
module adds what a deck does.  ``open_document``, ``list_documents``, ``close_document``,
``undo`` and ``batch`` need nothing format-specific and come with the shared module.
"""

from __future__ import annotations

import re
from typing import Any

from ooxml_edit.tools import Result, ToolError, page_list, shared

from .common import KIND, facts, library_errors, numbers, slide_id_of, validate_delta
from .render import png_size, render_slides

#: Slides a render may hold, and the default width (1280 px: 1,196 Claude tokens at 16:9).
MAX_RENDER = 4
DEFAULT_WIDTH = 1280


def _deck_only(field: str, use: str) -> ToolError:
    return ToolError("invalid_arguments", f"{field} is for documents; a deck takes {use}",
                     field=field, valid_options=[use])


@shared.handler("new_document", kind=KIND)
@library_errors
def new_document(call, kind, template_blob=None, size=None, title=None, author=None,
                 name=None):
    from ..edit.document import Document

    if size is not None and size not in ("16:9", "4:3"):
        raise ToolError("invalid_arguments", f"a deck's size is 16:9 or 4:3, not {size!r}",
                        field="size", valid_options=["16:9", "4:3"])
    template = call.blob(template_blob).data if template_blob else None
    deck = Document.new(size=size, template=template, title=title, author=author,
                        created=call.now())
    doc_id = call.session.adopt(deck, KIND, name or "new.pptx",
                                source=template_blob or "new")
    width, height = deck.slide_size
    return Result(summary=f"Made {doc_id}: a new deck", created=[doc_id],
                  data={"doc": doc_id, "kind": KIND, "slides": 0,
                        "size": {"w": round(width / 12700, 2), "h": round(height / 12700, 2)},
                        "layouts": [layout.name for layout in deck.layouts]})


@shared.handler("save_document", kind=KIND)
@library_errors
def save_document(call, doc, name, format):
    deck = call.document
    if format == "pptx":
        data = deck.to_bytes(template=False)
    elif format == "potx":
        data = deck.to_bytes(template=True)
    elif format == "outline":
        data = deck.to_outline().encode("utf-8")
    else:
        raise ToolError("invalid_arguments", f"a deck saves as pptx, potx or outline, not "
                        f"{format}", field="format", valid_options=["pptx", "potx", "outline"])
    described = call.output(name, format, data)
    left = unresolved(deck)
    summary = f"Saved {name} ({format}, {len(data):,} bytes)"
    if left:
        counts = ", ".join(f"{len(left[k])} {k.replace('_', ' ')}" for k in UNRESOLVED
                           if left.get(k))
        summary += f"; still in the deck: {counts} (unresolved, listed)"
        described = {**described, "unresolved": left}
    return Result(summary=summary, data=described)


#: The fit and collision facts ``save_document`` reports when any are left.
UNRESOLVED = ("overflows", "collisions", "off_slide")
#: At most this many of each kind are listed.
UNRESOLVED_LISTED = 20


def unresolved(deck) -> dict[str, Any]:
    """The fit and collision facts still in the deck at save time, every slide: what
    ``check`` would report.  Facts, not a refusal: the save has happened."""
    found = facts(deck, [slide.slide_id for slide in deck.slides])
    left: dict[str, Any] = {}
    for kind in UNRESOLVED:
        entries = found.get(kind) or []
        if entries:
            left[kind] = entries[:UNRESOLVED_LISTED]
            if len(entries) > UNRESOLVED_LISTED:
                left[f"{kind}_total"] = len(entries)
    return left


def _pattern(text: str, regex: bool):
    if not regex:
        return re.compile(re.escape(text))
    try:
        return re.compile(text)
    except re.error as exc:
        raise ToolError("invalid_arguments", f"not a regular expression: {exc}",
                        field="text") from None


def _context(text: str, start: int, end: int, width: int = 40) -> str:
    left = max(0, start - width)
    right = min(len(text), end + width)
    return (("…" if left else "") + text[left:right] + ("…" if right < len(text) else "")
            ).replace("\v", "↵")


@shared.handler("find_text", kind=KIND)
@library_errors
def find_text(call, doc, text, regex=False, slides=None, range=None, stories=None,
              cursor=None):
    if range is not None:
        raise _deck_only("range", "slides")
    if stories is not None:
        raise _deck_only("stories", "slides")
    deck = call.document
    pattern = _pattern(text, regex)
    matches = []
    for block in deck.outline_blocks(numbers(deck, slides)):
        for found in pattern.finditer(block.text):
            if found.end() == found.start():
                continue
            matches.append({"address": block.address, "kind": block.kind, "slide": block.slide,
                            "context": _context(block.text, found.start(), found.end())})
    page, total, next_cursor = page_list(matches, cursor=cursor, limit=call.limits.max_list_items)
    return Result(summary=f"{total} match(es) for {text!r}", data=page, total=total,
                  next_cursor=next_cursor)


@shared.handler("replace_text", kind=KIND)
@library_errors
def replace_text(call, doc, find, replace, expect, regex=False, slides=None, range=None,
                 stories=None):
    if range is not None:
        raise _deck_only("range", "slides")
    if stories is not None:
        raise _deck_only("stories", "slides")
    deck = call.document
    pattern = _pattern(find, regex)
    hits = []
    for block in deck.outline_blocks(numbers(deck, slides)):
        count = sum(1 for m in pattern.finditer(block.text) if m.end() > m.start())
        if count:
            hits.append((block, count))
    total = sum(count for _, count in hits)
    if expect == "one" and total != 1:
        if total == 0:
            raise ToolError("not_found", f"{find!r} is not in the deck", field="find")
        raise ToolError("ambiguous", f"{find!r} occurs {total} times; expect=one needs "
                        "exactly one: narrow it, or use expect=all",
                        valid_options=[f"{block.address}: {_context(block.text, *_span(pattern, block.text))}"
                                       for block, _ in hits][:50], field="find")
    changed = []
    for block, _ in hits:
        target = deck.resolve(block.address)
        new = pattern.sub(replace, block.text) if regex else block.text.replace(find, replace)
        target.text = new
        changed.append(block.address)
        sid = slide_id_of(block.address)
        if sid is not None:
            call.touch(sid)
    return Result(summary=f"Replaced {total} occurrence(s) in {len(changed)} place(s)",
                  changed=changed, data={"count": total})


def _span(pattern, text):
    found = pattern.search(text)
    return (found.start(), found.end()) if found else (0, 0)


@shared.handler("render", kind=KIND)
@library_errors
def render(call, doc, slides=None, pages=None, width=None):
    if pages is not None:
        raise _deck_only("pages", "slides")
    deck = call.document
    if not slides:
        raise ToolError("invalid_arguments", f"name the slides to render, at most "
                        f"{MAX_RENDER}, e.g. [1]", field="slides",
                        valid_options=list(range(1, len(deck.slides) + 1))[:50])
    if len(slides) > MAX_RENDER:
        raise ToolError("limit", f"at most {MAX_RENDER} slides per render", field="slides")
    wanted = numbers(deck, slides)
    width = width or DEFAULT_WIDTH
    entry = call.entry
    ids = {n: deck.slides[n - 1].slide_id for n in wanted}
    missing = [n for n in wanted
               if entry.render_cache.get((entry.version, ids[n], width)) is None]
    if missing:
        images = call.run(render_slides, deck.to_bytes(), missing, width)
        for number, png in zip(missing, images):
            entry.render_cache.put((entry.version, ids[number], width), png)
    described = []
    for number in wanted:
        png = entry.render_cache.get((entry.version, ids[number], width))
        w, h = png_size(png)
        image = call.image(png, w, h, label=f"slide {number} (s:{ids[number]})")
        described.append(image.describe())
    return Result(summary=f"Rendered slide(s) {', '.join(map(str, wanted))}",
                  data={"cached": [n for n in wanted if n not in missing]})


#: What ``check`` can report on a deck.
DECK_FACTS = ("fit", "collisions", "validate")


@shared.handler("check", kind=KIND)
@library_errors
def check(call, doc, slides=None, pages=None, include=None, boxes=False):
    if pages is not None:
        raise _deck_only("pages", "slides")
    deck = call.document
    wanted = list(include or DECK_FACTS)
    notes = []
    for item in wanted:
        if item in ("reflow", "fields"):
            notes.append(f"{item}: documents only")
        elif item == "app":
            notes.append("app: no critique pass is registered for decks")
    chosen = numbers(deck, slides)
    ids = [deck.slides[n - 1].slide_id for n in chosen]
    entry = call.entry
    key = ("check", tuple(ids), tuple(sorted(set(wanted))), bool(boxes))
    data = entry.check_cache.get((entry.version, key))
    if data is None:
        data = {"slides": [f"s:{sid}" for sid in ids]}
        data.update(facts(deck, ids, boxes=boxes, include=wanted))
        if "facts" in wanted or "design" in wanted:
            data.update(slide_facts_json(deck, ids, problems="facts" in wanted))
        if "validate" in wanted:
            problems = deck.validate()
            data["validate"] = {"problems": [str(p) for p in problems],
                                **validate_delta(entry, problems)}
        entry.check_cache.put((entry.version, key), data)
    if notes:
        data = {**data, "notes": notes}
    count = sum(len(data.get(k) or []) for k in ("overflows", "collisions", "off_slide"))
    new = len((data.get("validate") or {}).get("new") or [])
    return Result(summary=f"{count} fit/collision fact(s), {new} new validation problem(s)",
                  data=data)


def _compact_design(facts) -> dict[str, Any]:
    """The design facts a check carries: the headline of each kind; ppt_design_facts has
    the detail."""
    data = facts.to_json()
    palette = data["palette"]
    groups = data["groups"]
    alignment = data["alignment"]
    return {
        "palette": [{"family": f["family"], "accent": f["accent"], "count": f["count"],
                     **({"theme": False} if not f["theme"] else {})}
                    for f in palette["families"]],
        "sets": [{"preset": g["preset"], "count": g["count"], "shapes": g["shapes"][:8],
                  "fills": g["fills"], "accent_hues": g["accent_hues"],
                  "legend": None if g["legend"] is None else
                  {"covers": g["legend"]["covers"], "misses": g["legend"]["misses"]}}
                 for g in groups["sets"] if g["colour_coded"]],
        "legends": len(groups["legends"]),
        "empty": data["empty"]["regions"],
        "near_misses": alignment["near_misses_total"],
        "text_sizes": [s["size"] for s in data["text_sizes"]],
        "lines_over_text": data["lines_over_text"]}


def slide_facts_json(deck, ids, *, problems: bool) -> dict[str, Any]:
    """``check``'s facts: the problem facts (LP15) with ``problems``, and the design facts
    (LP24) in brief, per slide."""
    per_slide = []
    for sid in ids:
        target = deck.slide(f"s:{sid}")
        entry: dict[str, Any] = {"slide": f"s:{sid}"}
        if problems:
            entry.update(target.facts())
        entry["design"] = _compact_design(target.design_facts())
        per_slide.append(entry)
    return {"facts": per_slide}


@shared.handler("set_properties", kind=KIND)
@library_errors
def set_properties(call, doc, title=None, author=None, language=None, subject=None):
    deck = call.document
    if subject is not None:
        raise ToolError("invalid_arguments", "a deck has no subject property here; set "
                        "title, author or language", field="subject",
                        valid_options=["title", "author", "language"])
    if title is not None:
        deck.title = title
    if author is not None:
        deck.author = author
    if language is not None:
        deck.language = language
    return Result(summary="Set the deck's properties", changed=[doc],
                  data={"title": deck.title, "author": deck.author, "language": deck.language})


TOOLS = [new_document, save_document, find_text, replace_text, render, check, set_properties]
