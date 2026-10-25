"""The shared tools' PowerPoint handlers: new, save, find, replace, render, check, properties.

Their definitions are :mod:`ooxml_edit.tools.shared`'s, one per tool for every format; this
module adds what a deck does.  ``open_document``, ``close_document``, ``read_blob``,
``undo`` and ``batch`` need nothing format-specific and come with the shared module.
"""

from __future__ import annotations

import re
from typing import Any

from ooxml_edit.tools import Result, ToolError, collect_warnings, page_list, shared, tiles

from .common import (KIND, facts, library_errors, numbers, slide_id_of, targets,
                     template_key_of, template_options, validate_delta)
from ..edit.document import font_dirs_of, locale_of
from ..edit.effective import text_minimums
from ooxml_edit.tools import render as pipeline
from ooxml_edit.tools.render import png_size

from .render import RENDERER
from . import textfacts
from .tables import forget_cells

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
                        created=call.now(), locale=call.locale)
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
    # Slide numbers written by reference, as the slides are now: PowerPoint may have moved
    # them since the deck was opened (pptx_agent.edit.slidenumbers).
    renumbered = deck.update_slide_numbers() if format in ("pptx", "potx", "outline") else []
    deck.slide_numbers_kept = []
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
    if renumbered:
        described = {**described, "slide_numbers": renumbered[:UNRESOLVED_LISTED]}
    return Result(summary=summary, data=described)


#: The fit, collision, empty-placeholder and small-text facts ``save_document`` reports
#: when any are left.
UNRESOLVED = ("overflows", "collisions", "off_slide", "empty_placeholders", "small_text",
              "chart_title_font", "missing_slides")
#: At most this many of each kind are listed.
UNRESOLVED_LISTED = 20


def unresolved(deck) -> dict[str, Any]:
    """The fit, collision and empty-placeholder facts still in the deck at save time,
    every slide: what ``check`` would report.  Facts, not a refusal: the save has
    happened."""
    found = facts(deck, [slide.slide_id for slide in deck.slides])
    left: dict[str, Any] = {}
    for kind in UNRESOLVED:
        entries = found.get(kind) or []
        if entries:
            left[kind] = entries[:UNRESOLVED_LISTED]
            if len(entries) > UNRESOLVED_LISTED:
                left[f"{kind}_total"] = len(entries)
    if left.get("chart_title_font"):
        # Its entries are a chart and a call; the note says why the call is wanted.
        left["chart_title_font_note"] = found["chart_title_font_note"]
    return left


def _pattern(text: str, regex: bool):
    if not regex:
        return re.compile(re.escape(text))
    try:
        return re.compile(text)
    except re.error as exc:
        raise ToolError("invalid_arguments", f"not a regular expression: {exc}",
                        field="text") from None


#: ``stories`` on a deck: where find_text and replace_text look.  ``body`` (the default) is
#: the slides; ``layouts``, ``masters`` and ``all`` add the text that shows on every slide
#: built on them (a legal line, a footer year, a placeholder's prompt).
_DECK_STORIES = {None: "slides", "body": "slides", "layouts": "layouts", "masters": "masters",
                 "all": "all"}


def _scope(stories: str | None) -> str:
    if stories not in _DECK_STORIES:
        raise ToolError("invalid_arguments", f"stories is {stories!r}; a deck takes "
                        "body (its slides), layouts, masters or all", field="stories",
                        valid_options=["body", "layouts", "masters", "all"])
    return _DECK_STORIES[stories]


def _blocks(deck, slides, stories):
    scope = _scope(stories)
    if scope in ("layouts", "masters") and slides is not None:
        raise ToolError("invalid_arguments", f"slides narrows the slides; stories={stories} "
                        "searches no slides", field="slides")
    chosen = numbers(deck, slides) if scope in ("slides", "all") else None
    return deck.outline_blocks(chosen, scope=scope)


def _match(block, start: int, end: int) -> dict[str, Any]:
    data = {"address": block.address, "kind": block.kind}
    if block.part is not None:
        data["on"] = block.part
    else:
        data["slide"] = block.slide
    data["context"] = _context(block.text, start, end)
    return data


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
    deck = call.document
    pattern = _pattern(text, regex)
    matches = []
    for block in _blocks(deck, slides, stories):
        for found in pattern.finditer(block.text):
            if found.end() == found.start():
                continue
            matches.append(_match(block, found.start(), found.end()))
    page, total, next_cursor = page_list(matches, cursor=cursor, limit=call.limits.max_list_items)
    return Result(summary=f"{total} match(es) for {text!r}", data=page, total=total,
                  next_cursor=next_cursor)


@shared.handler("replace_text", kind=KIND)
@library_errors
def replace_text(call, doc, find, replace, expect, regex=False, slides=None, range=None,
                 stories=None):
    if range is not None:
        raise _deck_only("range", "slides")
    deck = call.document
    pattern = _pattern(find, regex)
    hits = []
    for block in _blocks(deck, slides, stories):
        count = sum(1 for m in pattern.finditer(block.text) if m.end() > m.start())
        if count:
            hits.append((block, count))
    total = sum(count for _, count in hits)
    if expect == "one" and total != 1:
        if total == 0:
            elsewhere = [] if _scope(stories) == "all" else [
                block.address for block in deck.outline_blocks(scope="all")
                if block.part is not None and pattern.search(block.text)]
            if elsewhere:
                raise ToolError("not_found", f"{find!r} is not on the slides but on a layout or "
                                f"master ({', '.join(elsewhere[:5])}): stories=all reaches it",
                                field="stories", valid_options=["all", "layouts", "masters"])
            raise ToolError("not_found", f"{find!r} is not in the deck", field="find")
        raise ToolError("ambiguous", f"{find!r} occurs {total} times; expect=one needs "
                        "exactly one: narrow it, or use expect=all",
                        valid_options=[f"{block.address}: {_context(block.text, *_span(pattern, block.text))}"
                                       for block, _ in hits][:50], field="find")
    changed = []
    written = call.facts_at("replace")
    for block, _ in hits:
        target = deck.resolve(block.address)
        where = textfacts.spot(deck, block.address)
        before = where.text() if where is not None else None
        new = pattern.sub(replace, block.text) if regex else block.text.replace(find, replace)
        target.text = new
        if where is not None:            # the values replace wrote by reference, and the rest
            where.written(before, [(start + fact.start, fact) for start in
                                   _replaced_at(pattern, block.text, replace, regex)
                                   for fact in written] if written else ())
        changed.append(block.address)
        sid = slide_id_of(block.address)
        if sid is not None:
            call.touch(sid)
        key = template_key_of(block.address)
        if key is not None:
            call.touch(key)
    forget_cells(call, changed)          # a table's source record drops what was replaced
    result = Result(summary=f"Replaced {total} occurrence(s) in {len(changed)} place(s)",
                    changed=changed, data={"count": total})
    if not total and _scope(stories) != "all":
        elsewhere = [block.address for block in deck.outline_blocks(scope="all")
                     if block.part is not None and pattern.search(block.text)]
        if elsewhere:
            result.warnings.append(f"{find!r} is on a layout or master too "
                                   f"({', '.join(elsewhere[:5])}): stories=all reaches it")
    return result


def _replaced_at(pattern, text: str, replace: str, regex: bool) -> list[int]:
    """Where each replacement starts in the new text; none for a regex replacement with
    groups (``\\1``), whose length is not the replacement's."""
    if regex and "\\" in replace:
        return []
    starts, delta = [], 0
    for match in pattern.finditer(text):
        if not regex and match.end() == match.start():
            continue
        starts.append(match.start() + delta)
        delta += len(replace) - (match.end() - match.start())
    return starts


def _span(pattern, text):
    found = pattern.search(text)
    return (found.start(), found.end()) if found else (0, 0)


@shared.handler("render", kind=KIND)
@library_errors
def render(call, doc, slides=None, pages=None, width=None, sheet=None):
    if pages is not None:
        raise _deck_only("pages", "slides")
    deck = call.document
    if not slides:
        raise ToolError("invalid_arguments", f"name the slides to render, at most "
                        f"{MAX_RENDER}, e.g. [1], or a layout, L2; with sheet, up to "
                        f"{tiles.MAX_TILES}", field="slides",
                        valid_options=(list(range(1, len(deck.slides) + 1))
                                       or template_options(deck))[:50])
    if sheet is not None:
        return _render_sheet(call, deck, slides, width, sheet)
    if len(slides) > MAX_RENDER:
        raise ToolError("limit", f"at most {MAX_RENDER} slides per render; sheet: {{}} tiles "
                        f"up to {tiles.MAX_TILES} into one image", field="slides")
    wanted = targets(deck, slides)
    width = width or DEFAULT_WIDTH
    call.check_images(len(wanted))       # before drawing: a refused call costs nothing
    drawn, ids = draw(call, deck, wanted, width)
    described = []
    for number in wanted:
        png = drawn[number][0]
        w, h = png_size(png)
        image = call.image(png, w, h, label=_render_name(deck, number))
        described.append(image.describe())
    unrendered = _left_out(wanted, ids, drawn)
    data: dict[str, Any] = {"cached": [n for n in wanted if drawn[n][2]]}
    if any(isinstance(n, str) for n in wanted):
        data["template_note"] = TEMPLATE_NOTE
    if unrendered:
        data["missing_glyphs"] = unrendered
        data["missing_glyphs_note"] = MISSING_GLYPHS_NOTE
    shown = [", ".join(str(n) for n in wanted if not isinstance(n, str))]
    shown = [f"slide(s) {shown[0]}"] if shown[0] else []
    shown += [f"{n} ({'layout' if n.startswith('L') else 'master'})"
              for n in wanted if isinstance(n, str)]
    return Result(summary=f"Rendered {', '.join(shown)}", data=data)


def draw(call, deck, items, width) -> tuple[dict, dict]:
    """``{item: (png, missing, cached)}`` for slides and template keys ``items`` drawn
    ``width`` px wide, and ``{item: id}`` (a slide's id, a layout's or master's key).

    ooxml-edit's render pipeline draws them (:func:`ooxml_edit.tools.render.draw`, with the
    deck's :data:`.render.RENDERER`): an image is kept by its content in the toolbox's
    store, this version's key of each view on the document, so a repeat is answered there;
    the rest go to the workers at once, in chunks, told which images the store holds -- a
    view an edit left as it was, or the same slide in another document, is keyed, found and
    not drawn again -- and handed the faces a worker decoded from the deck's embedded
    fonts, which no worker decodes again.  ``cached`` says an image was not drawn by this
    call."""
    entry = call.entry
    # A slide by its id, a layout or master by its key (L2, M0): what the keys are kept by.
    ids = {item: item if isinstance(item, str) else deck.slides[item - 1].slide_id
           for item in items}
    # A layout or master goes to the worker by its part path.
    parts = {item: deck._template_by_key(item).part_path for item in dict.fromkeys(items)
             if isinstance(item, str)}
    # The font folders, resolved here (the session's, else OOXML_FONT_DIRS) and handed to
    # the worker, which sees neither; they key the cache, so renders with different
    # folders do not mix.  The locale likewise: the session's, handed over and keying.
    fonts, locale = font_dirs_of(deck), locale_of(deck)
    # The faces a worker decoded from the deck's embedded fonts, for the others.
    decoded = {key: found[0] for key in entry.font_keys.keys()
               if (found := call.decoded_fonts.peek(key)) is not None}
    # No worker has decoded this deck's embedded faces yet: one chunk first, alone, and its
    # faces go with the others, so no worker decodes them a second time.
    first_alone = not entry.font_keys.keys() and _embeds_fonts(deck)

    def on_facts(facts, alone):
        found = facts["fonts"]
        for key, font in found.items():
            call.decoded_fonts.put(key, font)
            entry.font_keys.put(key, True)
        if alone:
            if not found:
                entry.font_keys.put(NO_FACES, True)   # it embeds none a slide uses
            return {**decoded, **found}
        return None

    drawn, _ = pipeline.draw(call, RENDERER, items, width, ids=ids, fonts=fonts, locale=locale,
                             wire=parts, state=decoded, first_alone=first_alone,
                             on_facts=on_facts)
    return drawn, ids


#: Kept in an entry's ``font_keys`` once a deck's render decoded no embedded face.
NO_FACES = "none"


def _embeds_fonts(deck) -> bool:
    """Whether the deck lists embedded fonts (``p:embeddedFontLst``)."""
    root = deck.package.tree(deck.package.presentation_part())
    return root is not None and any(
        child.tag.endswith("}embeddedFontLst") and len(child) for child in root)


def _left_out(wanted, ids, drawn) -> list[dict]:
    """The ``missing_glyphs`` facts of the views drawn: each face and script once, on the
    first view that has it, as pptx2svg reports a render's."""
    out, seen = [], set()
    for item in wanted:
        address = item if isinstance(item, str) else f"s:{ids[item]}"
        for fact in drawn[item][1]:
            if (fact["face"], fact["script"]) in seen:
                continue
            seen.add((fact["face"], fact["script"]))
            out.append({"slide": address, **fact})
    return out


def _render_sheet(call, deck, slides, width, sheet) -> Result:
    """``render`` with ``sheet``: the slides tiled into one image, in order, each labelled
    with its number and id (``3 · s:258``; a layout ``L2 Title Only``) -- a contact sheet
    (:mod:`ooxml_edit.tools.tiles`).  ``sheet.width`` is each tile's; ``width`` stands in
    for it when the sheet gives none.  The tiles are the images ``render`` keeps
    (:func:`draw`: those it holds are not drawn again, the rest over the workers at once);
    the sheet is cached whole by version, one entry, so a sheet of 30 does not push the
    single renders out.  ``cached_tiles`` lists the tiles not drawn for it.  Tiles too
    wide or too tall for the image's edge are drawn narrower (:func:`tiles.fit`), and
    ``sheet.shrunk`` says from what width and why."""
    limits = call.limits
    wanted = targets(deck, slides)
    tiling = tiles.tiling_of(sheet, len(wanted), name="sheet", width=width, what="slides",
                             max_edge=limits.max_image_edge)
    cx, cy = deck.slide_size
    # Narrower tiles where the sheet would be too tall: a slide's shape is known.
    tiling = tiles.fit(tiling, [cy / cx] * len(wanted), name="sheet", what="slides",
                       max_edge=limits.max_image_edge)
    call.check_images(1)                 # before drawing: a refused call costs nothing
    entry, version = call.entry, call.entry.version
    ids = {item: item if isinstance(item, str) else deck.slides[item - 1].slide_id
           for item in wanted}
    fonts, locale = font_dirs_of(deck), locale_of(deck)
    key = ("sheet", version, tuple(ids[n] for n in wanted), tiling.columns, tiling.width,
           fonts, locale)
    made = entry.render_cache.get(key)
    cached = made is not None
    tiles_cached = None
    if made is None:
        drawn, _ = draw(call, deck, wanted, tiling.width)
        views = [(drawn[n][0], *png_size(drawn[n][0])) for n in wanted]
        fitted = tiles.fit(tiling, [h / w for _, w, h in views], name="sheet", what="slides",
                           max_edge=limits.max_image_edge)
        if fitted.width != tiling.width:     # drawn a pixel taller than the slide's shape
            tiling = fitted
            drawn, _ = draw(call, deck, wanted, tiling.width)
        labels = {n: _sheet_label(deck, n) for n in wanted}
        tiles_cached = [n for n in wanted if drawn[n][2]]
        # The snapshot (the tiles, ids, labels) is taken; the worker draws from it.
        png, w, h = pipeline.draw_tiles(call, RENDERER, [(drawn[n][0], labels[n]) for n in wanted],
                                        tiling, name="sheet", what="slides", fonts=fonts)
        made = (png, w, h, {n: drawn[n][1] for n in wanted}, tiling)
        entry.render_cache.put(key, made)
    png, w, h, lost, tiling = made
    shown = [n if isinstance(n, str) else f"s:{ids[n]}" for n in wanted]
    call.image(png, w, h, label=f"sheet of {len(wanted)}: {', '.join(shown[:3])}"
               f"{', ...' if len(shown) > 3 else ''}")
    data: dict[str, Any] = {"sheet": {"columns": tiling.columns, "width": tiling.width,
                                      "slides": shown},
                            "cached": cached}
    if tiling.shrunk:
        data["sheet"]["shrunk"] = tiling.shrunk
    if tiles_cached:
        data["cached_tiles"] = tiles_cached
    if any(isinstance(n, str) for n in wanted):
        data["template_note"] = TEMPLATE_NOTE
    unrendered = _left_out(wanted, ids, {n: (None, lost[n], None) for n in wanted})
    if unrendered:
        data["missing_glyphs"] = unrendered
        data["missing_glyphs_note"] = MISSING_GLYPHS_NOTE
    return Result(summary=f"Rendered {len(wanted)} slide(s) as one sheet, {tiling.columns} "
                  f"column(s) of {tiling.width} px{tiles.narrowed_note(tiling)}, labelled with "
                  "number and id", data=data)


def _sheet_label(deck, item) -> str:
    """A tile's label: ``3 · s:258``, `` · hidden`` after a hidden slide's; ``L2 Title Only``."""
    if isinstance(item, str):
        return f"{item} {deck._template_by_key(item).name}"
    slide = deck.slides[item - 1]
    root = deck.package.tree(slide.part_path)
    hidden = root is not None and root.get("show") in ("0", "false")
    return f"{item} · s:{slide.slide_id}{' · hidden' if hidden else ''}"


def _render_name(deck, item) -> str:
    """An image's label: ``slide 2 (s:257)``, ``layout L2 'Title and Content'``."""
    if not isinstance(item, str):
        return f"slide {item} (s:{deck.slides[item - 1].slide_id})"
    template = deck._template_by_key(item)
    kind = "layout" if item.startswith("L") else "master"
    return f"{kind} {item} {template.name!r}"


#: What a layout's or master's image is, for the model reading it.
TEMPLATE_NOTE = ("a layout or master is drawn as a new slide on it starts, its placeholders "
                 "as dashed boxes tagged with type and idx (a wireframe, not PowerPoint's "
                 "drawing); describe lists the layouts")


#: What a ``missing_glyphs`` fact means for the model reading the image.
MISSING_GLYPHS_NOTE = ("the image leaves this text out: the renderer has no font for it "
                       "(pptx2svg-fonts not installed?). The deck is unchanged and "
                       "PowerPoint draws it; do not edit the text because the image lacks it")


#: What ``check`` can report on a deck.
DECK_FACTS = ("fit", "collisions", "validate")


@shared.handler("check", kind=KIND)
@library_errors
def check(call, doc, slides=None, pages=None, include=None, boxes=False, spacing=None,
          contrast=None):
    if pages is not None:
        raise _deck_only("pages", "slides")
    below = None if spacing is None else float(spacing)
    minimum = None if contrast is None else float(contrast)
    deck = call.document
    wanted = list(include or DECK_FACTS)
    notes = []
    for item in wanted:
        if item in ("reflow", "fields", "layout"):
            notes.append(f"{item}: documents only")
        elif item == "app":
            notes.append("app: no critique pass is registered for decks")
    chosen = numbers(deck, slides)
    ids = [deck.slides[n - 1].slide_id for n in chosen]
    entry = call.entry
    version = entry.version
    key = ("check", tuple(ids), tuple(sorted(set(wanted))), bool(boxes), font_dirs_of(deck),
           locale_of(deck), tuple(text_minimums(deck.min_text_size).items()), below, minimum)
    data = entry.check_cache.get((version, key))
    if data is None:
        target = deck
        if call.may_unlock:
            # Measuring is the slow part: it runs on a copy of this version, opened from the
            # bytes taken under the lock, with the lock let go -- so an edit elsewhere in the
            # deck does not wait for it.  Inside a batch it runs on the deck, under the lock.
            data_bytes, font_dirs, locale = entry.snapshot(), deck.font_dirs, deck.locale
        with call.unlocked() as released:
            if released:
                with collect_warnings():          # said when the deck was opened, not now
                    target = type(deck).open(data_bytes)
                target.font_dirs = font_dirs
                target.locale = locale
                target.min_text_size = deck.min_text_size
            data = {"slides": [f"s:{sid}" for sid in ids]}
            data.update(facts(target, ids, boxes=boxes, include=wanted))
            if "facts" in wanted or "design" in wanted:
                data.update(slide_facts_json(target, ids, problems="facts" in wanted))
            if below is not None:
                data.update(spacing_json(target, ids, below))
            if minimum is not None:
                data.update(contrast_json(target, ids, minimum))
            problems = target.validate() if "validate" in wanted else None
        if problems is not None:
            data["validate"] = {"problems": [str(p) for p in problems],
                                **validate_delta(entry, problems)}
        entry.check_cache.put((version, key), data)
    if notes:
        data = {**data, "notes": notes}
    count = sum(len(data.get(k) or []) for k in ("overflows", "collisions", "off_slide",
                                                  "small_text", "chart_title_font",
                                                  "missing_slides"))
    new = len((data.get("validate") or {}).get("new") or [])
    tight = ""
    if below is not None:
        tight = (f", {len(data['spacing']) + data.get('spacing_left_out', 0)} spacing(s) "
                 f"below {below:g} pt")
    if minimum is not None:
        tight += (f", {len(data['contrast']) + data.get('contrast_left_out', 0)} text(s) "
                  f"below {minimum:g}:1 contrast")
    return Result(summary=f"{count} fit/collision fact(s), {new} new validation problem(s)"
                  f"{tight}", data=data)


#: The spacings ``check`` lists per slide, nearest first; ``spacing_left_out`` counts the rest.
SPACING_PER_SLIDE = 12

#: Said with ``spacing``: what an entry measures.
SPACING_NOTE = ("spacing: measured, nearest first, pt: between two neighbours on an axis; "
                "inside a filled or outlined shape, the content's padding to an edge; "
                "clearance to a footer, slide_number or date, or past the layout's "
                "content_area (gap < 0). Text as drawn, not its box. Overlaps are "
                "collisions; backgrounds, lines and connectors are left out. Cite the gap")


def spacing_json(deck, ids, below: float) -> dict[str, Any]:
    """``check``'s ``spacing``: each slide's tightest spacings below ``below`` pt
    (:meth:`Slide.spacing <pptx_agent.Slide.spacing>`), nearest first, at most
    :data:`SPACING_PER_SLIDE` a slide."""
    found, left_out, drawn = [], 0, True
    for sid in ids:
        report = deck.slide(f"s:{sid}").spacing(below=round(below * 12700))
        drawn = drawn and report.drawn
        found += [entry.to_json() for entry in report.entries[:SPACING_PER_SLIDE]]
        left_out += max(0, len(report.entries) - SPACING_PER_SLIDE)
    out: dict[str, Any] = {"spacing": found}
    if left_out:
        out["spacing_left_out"] = left_out
    if found:
        out["spacing_note"] = SPACING_NOTE
    if not drawn:                              # pragma: no cover - without pptx2svg
        out["spacing_boxes"] = "text measured by its box: pptx2svg is not installed"
    return out


#: The texts ``check`` lists per slide for ``contrast``, lowest first; ``contrast_left_out``
#: counts the rest.
CONTRAST_PER_SLIDE = 12

#: Said with ``contrast``: what an entry measures.
CONTRAST_NOTE = ("contrast: measured, lowest first: WCAG 2 ratio of the text's colour as drawn "
                 "(source: where it comes from) to what is behind it (over: its own fill, a "
                 "shape beneath, or the background; a gradient's worst stop; alpha blended); "
                 "large text (18 pt, 14 pt bold) held to 2/3 of the threshold; behind "
                 "unknown: a picture, table or chart there, not sampled")


def contrast_json(deck, ids, minimum: float) -> dict[str, Any]:
    """``check``'s ``contrast``: each slide's text below ``minimum``:1 contrast with what is
    behind it (:meth:`Slide.contrast <pptx_agent.Slide.contrast>`), lowest first, at most
    :data:`CONTRAST_PER_SLIDE` a slide."""
    found, left_out = [], 0
    for sid in ids:
        report = deck.slide(f"s:{sid}").contrast(minimum=minimum)
        found += [entry.to_json() for entry in report.entries[:CONTRAST_PER_SLIDE]]
        left_out += max(0, len(report.entries) - CONTRAST_PER_SLIDE)
    out: dict[str, Any] = {"contrast": found}
    if left_out:
        out["contrast_left_out"] = left_out
    if found:
        out["contrast_note"] = CONTRAST_NOTE
    return out


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
