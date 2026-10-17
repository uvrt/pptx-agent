"""Slide tools: ``ppt_add_slide``, ``ppt_draft_slides``, ``ppt_manage_slides``,
``ppt_set_theme``.

Positions are slide numbers, from 1, as PowerPoint numbers them; a slide's lasting address
is its id, ``s:256``.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (REVIEW, Result, ToolError, array, boolean, integer, number, obj,
                               string, tool)

from ..edit.theme import MAPPED, SLOTS
from .common import (KIND, box_emu, box_pt, id_hint, library_errors, numbers, resolve,
                     slide as slide_at, template_key_of, template_options)

SLIDES = "ppt_slides"
_KEY = string("Retry key: a repeat with the same key makes nothing new.", optional=True)


def _order(deck) -> list[str]:
    return [f"s:{s.slide_id}" for s in deck.slides]


def _layout(deck, name: str):
    try:
        return deck.layout(name)
    except (KeyError, LookupError, ValueError):
        raise ToolError("not_found", f"no layout {name!r}", field="layout",
                        valid_options=[layout.name for layout in deck.layouts]) from None


def _body_lines(body: str) -> list[tuple[str, int]]:
    """``body`` lines as (text, level): a leading ``- `` is dropped, two spaces of indent
    before it are one level."""
    lines = []
    for raw in body.split("\n"):
        stripped = raw.lstrip(" ")
        level = (len(raw) - len(stripped)) // 2
        if stripped.startswith(("- ", "* ")):
            stripped = stripped[2:]
        lines.append((stripped, min(level, 8)))
    return lines


def _position(deck, at: int | None, field: str = "at") -> int | None:
    if at is None:
        return None
    if not 1 <= at <= len(deck.slides) + 1:
        raise ToolError("invalid_arguments", f"position {at}; from 1 to {len(deck.slides) + 1}"
                        f"{id_hint(deck, at)}", field=field)
    return at - 1


@tool("ppt_add_slide",
      "Add a slide from a layout, optionally with title, body text and notes. Returns its "
      "id and placeholders.",
      {"doc": string("Document id."),
       "layout": string("Layout name, from describe, e.g. Title and Content."),
       "at": integer("Position from 1. Default last.", minimum=1, maximum=10000,
                     optional=True),
       "title": string("Title text.", optional=True),
       "body": string("Body text: a paragraph per line, two spaces of indent per level.",
                      optional=True),
       "notes": string("Speaker notes.", optional=True),
       "ref": string("Ref name for the slide.", optional=True),
       "key": _KEY},
      kind=KIND, group=SLIDES, mutates=True)
@library_errors
def ppt_add_slide(call, doc, layout, at=None, title=None, body=None, notes=None, ref=None,
                  key=None):
    deck = call.document
    added = deck.add_slide(_layout(deck, layout), index=_position(deck, at))
    if title is not None:
        if added.title is None:
            raise ToolError("invalid_arguments", f"layout {layout!r} has no title placeholder",
                            field="title")
        added.title = title
    if body is not None:
        holder = next((shape for shape in added.shapes if shape.placeholder is not None
                       and shape.placeholder.type in (None, "body", "obj")), None)
        if holder is None:
            raise ToolError("invalid_arguments", f"layout {layout!r} has no body placeholder",
                            field="body")
        lines = _body_lines(body)
        holder.set_text("\n".join(text for text, _ in lines))
        for paragraph, (_, level) in zip(holder.text_frame.paragraphs, lines):
            if level:
                paragraph.level = level
    if notes is not None:
        added.notes = notes
    address = f"s:{added.slide_id}"
    if ref:
        call.define_ref(ref, address)
    call.touch(added.slide_id)
    placeholders = [{"id": shape.id, "type": shape.placeholder.type or "content",
                     "box": box_pt(shape.bounds)}
                    for shape in added.shapes if shape.placeholder is not None]
    for entry, shape in zip(placeholders, [s for s in added.shapes if s.placeholder]):
        if (shape.placeholder.idx or 0) > 12:   # one added to the layout: say which
            entry.update(idx=shape.placeholder.idx, name=shape.name)
    return Result(summary=f"Added slide {address} ({layout})", created=[address],
                  data={"slide": address, "n": added.index + 1, "placeholders": placeholders,
                        "content_area": box_pt(added.content_area)})


@tool("ppt_draft_slides",
      "Draft whole slides from Markdown in the deck's layouts: # starts a slide (its "
      "title), lists fill the body, tables become tables, ![alt](b2) a picture blob, Notes: "
      "speaker notes.",
      {"doc": string("Document id."),
       "markdown": string("The outline."),
       "at": integer("First new slide's position from 1. Default last.", minimum=1,
                     maximum=10000, optional=True)},
      kind=KIND, group=SLIDES, mutates=True)
@library_errors
def ppt_draft_slides(call, doc, markdown, at=None):
    deck = call.document
    if len(markdown) > call.limits.max_markdown_chars:
        raise ToolError("limit", f"at most {call.limits.max_markdown_chars:,} characters",
                        field="markdown")
    if markdown.count("\n#") + markdown.startswith("#") > 200:
        raise ToolError("limit", "at most 200 slides per call", field="markdown")
    images = {handle: blob.data for handle, blob in call.session.blobs.items()
              if blob.mime.startswith("image/")}
    drafted = deck.insert_outline(markdown, at=_position(deck, at), images=images)
    for added in drafted:
        call.touch(added.slide_id)
    created = [f"s:{s.slide_id}" for s in drafted]
    return Result(summary=f"Drafted {len(drafted)} slide(s)", created=created,
                  data={"slides": [{"id": f"s:{s.slide_id}", "n": s.index + 1,
                                    "title": s.title,
                                    "layout": s.layout.name if s.layout else None}
                                   for s in drafted]})


@tool("ppt_manage_slides",
      "Duplicate, move or delete a slide, find one by its title, delete its empty "
      "placeholders (not a title or one with a fill), or set_layout as PowerPoint's Layout "
      "menu. Returns the slide order.",
      {"doc": string("Document id."),
       "action": string("What to do.", enum=["duplicate", "move", "delete", "find_by_title",
                                             "delete_empty_placeholders", "set_layout"]),
       "slide": string("The slide, s:256 or $ref.", optional=True),
       "to": integer("move, duplicate: new position from 1. Default after it.", minimum=1,
                     maximum=10000, optional=True),
       "notes": string("duplicate: the copy's notes.", optional=True),
       "title": string("find_by_title: the title.", optional=True),
       "ref": string("duplicate: ref name for the copy.", optional=True),
       "layout": string("set_layout: L2 or its name.", optional=True)},
      kind=KIND, group=SLIDES, mutates=True, refs=("slide",),
      action_roles={"find_by_title": REVIEW})
@library_errors
def ppt_manage_slides(call, doc, action, slide=None, to=None, notes=None, title=None,
                      ref=None, layout=None):
    deck = call.document
    if action == "find_by_title":
        if title is None:
            raise ToolError("invalid_arguments", "find_by_title needs title", field="title")
        found = deck.slide_titled(title)
        return Result(summary=f"{title!r} is s:{found.slide_id}",
                      data={"slide": f"s:{found.slide_id}", "n": found.index + 1,
                            "order": _order(deck)})
    if slide is None:
        raise ToolError("invalid_arguments", f"{action} needs slide", field="slide")
    target = slide_at(deck, slide)
    if action == "duplicate":
        index = _position(deck, to, "to")
        copy = deck.duplicate_slide(target, index=index, notes=notes)
        address = f"s:{copy.slide_id}"
        if ref:
            call.define_ref(ref, address)
        call.touch(copy.slide_id)
        return Result(summary=f"Duplicated s:{target.slide_id} as {address}",
                      created=[address], data={"slide": address, "order": _order(deck)})
    if action == "move":
        if to is None:
            raise ToolError("invalid_arguments", "move needs to", field="to")
        if not 1 <= to <= len(deck.slides):
            raise ToolError("invalid_arguments", f"position {to}; from 1 to {len(deck.slides)}"
                            f"{id_hint(deck, to)}", field="to")
        deck.move_slide(target, to - 1)
        return Result(summary=f"Moved s:{target.slide_id} to {to}",
                      changed=[f"s:{target.slide_id}"], data={"order": _order(deck)})
    if action == "delete_empty_placeholders":
        deleted = target.delete_empty_placeholders()
        call.touch(target.slide_id)
        kept = target.empty_placeholders()
        titles = [shape.id for shape in kept if shape.placeholder.type in ("title", "ctrTitle")]
        boxes = [shape.id for shape in kept if shape.id not in titles]
        summary = f"Deleted {len(deleted)} empty placeholder(s) on s:{target.slide_id}"
        if titles:
            summary += f"; kept the empty title {', '.join(titles)}: fill it"
        if boxes:
            summary += (f"; kept {', '.join(boxes)}: its fill or outline is drawn; delete it "
                        "with ppt_arrange if the box should go")
        data = {"slide": f"s:{target.slide_id}", "deleted": deleted}
        if titles:
            data["empty_title"] = titles
        if boxes:
            data["kept_drawn"] = boxes
        return Result(summary=summary, removed=deleted, data=data)
    if action == "set_layout":
        return _set_layout(call, deck, target, layout)
    if action == "delete":
        address = f"s:{target.slide_id}"
        deck.delete_slide(target)
        return Result(summary=f"Deleted {address}", removed=[address],
                      data={"order": _order(deck)})
    raise ToolError("invalid_arguments", f"unknown action {action!r}", field="action")


@tool("ppt_set_theme",
      "Set the theme's colours and fonts; every theme reference in the deck follows. "
      "Returns the theme and its colour roles.",
      {"doc": string("Document id."),
       "colors": array(obj({"slot": string(enum=list(SLOTS + MAPPED)),
                            "hex": string("#RRGGBB.")}),
                       "Theme colour slots to set.", optional=True),
       "fonts": obj({"major": string("Headings.", optional=True),
                     "minor": string("Body.", optional=True)}, "Theme fonts.",
                    optional=True)},
      kind=KIND, group=SLIDES, mutates=True)
@library_errors
def ppt_set_theme(call, doc, colors=None, fonts=None):
    deck = call.document
    if not colors and not fonts:
        raise ToolError("invalid_arguments", "give colors or fonts", field="colors")
    theme = deck.theme
    if colors:
        theme = theme.set_colors({entry["slot"]: entry["hex"] for entry in colors})
    if fonts:
        theme = theme.set_fonts(major=fonts.get("major"), minor=fonts.get("minor"))
    for target in deck.slides:
        call.touch(target.slide_id)
    roles = theme.roles
    return Result(summary="Set the theme", changed=["theme"],
                  data={"colors": dict(theme.colors),
                        "fonts": {"major": theme.fonts.major, "minor": theme.fonts.minor},
                        "roles": {"primary": roles.primary, "secondary": roles.secondary,
                                  "neutral": roles.neutral, "highlight": roles.highlight}})


def _set_layout(call, deck, target, layout) -> Result:
    """ppt_manage_slides set_layout: the slide on another layout as PowerPoint's Layout menu
    puts it (measured, round 7, A), and what had no place there."""
    chosen = _template_layout(deck, layout, "layout")
    change = target.set_layout(chosen)
    call.touch(target.slide_id)
    address = f"s:{target.slide_id}"
    data: dict[str, Any] = {"slide": address, "layout": f"{chosen.id} ({chosen.name})",
                            "placeholders": [{"shape": shape, "to": place}
                                             for shape, place in change.matched]}
    if change.added:
        data["added"] = change.added
    if change.removed:
        data["removed_empty"] = change.removed
    if change.orphans:
        data["orphans"] = [{"shape": shape, "why": why} for shape, why in change.orphans]
    if change.boxes_dropped:
        data["boxes_dropped"] = [{"shape": shape, "box": box_pt(box)}
                                 for shape, box in change.boxes_dropped]
    notes = []
    if change.boxes_dropped:
        notes.append("placeholders with a box of their own now take the layout's, as in "
                     "PowerPoint (their text and formatting stay): set a box again with "
                     "ppt_set_shape if it was meant")
    if change.orphans:
        notes.append("orphans keep their content where it was drawn, snapped to the nearer "
                     "side, as PowerPoint keeps them; move it into a placeholder "
                     "(ppt_set_text) and delete the orphan, or keep it")
    if notes:
        data["note"] = "; ".join(notes)
    summary = (f"Put {address} on {chosen.id} ({chosen.name}): {len(change.matched)} "
               f"placeholder(s) kept, {len(change.added)} added")
    if change.orphans:
        summary += (f"; {len(change.orphans)} orphan(s): "
                    + ", ".join(f"{shape} ({why})" for shape, why in change.orphans))
    return Result(summary=summary, changed=[address], created=change.added,
                  removed=change.removed, data=data)


def _layouts_json(deck) -> list[dict[str, Any]]:
    return [{"id": layout.id, "name": layout.name, "slides": len(layout.slides)}
            for layout in deck.layouts]


def _template_layout(deck, address: str | None, field: str):
    if not address:
        raise ToolError("invalid_arguments", f"give {field}: a layout, L2 or its name",
                        field=field, valid_options=template_options(deck))
    try:
        return deck.layout(str(address).split(".")[0] if template_key_of(address) else address)
    except KeyError:
        raise ToolError("not_found", f"no layout {address!r}", field=field,
                        valid_options=[f"{lay.id} ({lay.name})" for lay in deck.layouts][:50]
                        ) from None


@tool("ppt_edit_master",
      "Slide masters and layouts: a layout placeholder's prompt text, Insert > Header & "
      "Footer, adding, duplicating, renaming or deleting layouts and adding their "
      "placeholders. Their shapes' text, "
      "fill and box: ppt_set_text, ppt_set_shape (L2.5, M0.7); new shapes: the add tools "
      "with slide L2 or M0. Returns the slides affected.",
      {"doc": string("Document id."),
       "action": string("What to do.", enum=["set_prompt", "set_footer", "add_layout",
                                             "duplicate_layout", "rename_layout",
                                             "delete_layout", "add_placeholder"]),
       "target": string("set_prompt: the layout placeholder, L2.3. Layout actions: the "
                        "layout, L2 or its name (add_layout: the one to follow; "
                        "add_placeholder: or L2.7 to add to slides).",
                        optional=True),
       "text": string("set_prompt: the prompt, \\n for the next level. set_footer: the "
                      "footer; empty: none. Other actions: the name.",
                      optional=True),
       "type": string("add_placeholder.", enum=["body", "content", "pic", "title"],
                      optional=True),
       "box": obj({"x": number(minimum=-10000, maximum=10000),
                   "y": number(minimum=-10000, maximum=10000),
                   "w": number(minimum=0, maximum=10000),
                   "h": number(minimum=0, maximum=10000)},
                  "add_placeholder: points.", optional=True),
       "reassign_to": string("delete_layout: the layout its slides move to.", optional=True),
       "slide_number": boolean("set_footer. Default true.", optional=True),
       "date": string("set_footer: auto (updates), fixed text, or leave out for none.",
                      optional=True),
       "slides": array(string(or_integer=1), "set_footer: only these slides (Apply), numbers "
                       "or ids (s:256). Default all (Apply to All: masters and layouts too). "
                       "add_placeholder: also add it to these.",
                       optional=True),
       "on_title_slide": boolean("set_footer. Default false.", optional=True)},
      kind=KIND, group=SLIDES, mutates=True, refs=("target", "slides[]"))
@library_errors
def ppt_edit_master(call, doc, action, target=None, text=None, reassign_to=None,
                    slide_number=None, date=None, slides=None, on_title_slide=None,
                    type=None, box=None):
    deck = call.document
    if action == "set_prompt":
        if text is None:
            raise ToolError("invalid_arguments", "set_prompt needs text", field="text")
        shape = resolve(deck, target or "", field="target")
        shape.set_prompt(text)
        key = template_key_of(shape.id)
        call.touch(key, shape.id)
        return Result(summary=f"Set the prompt of {shape.id}", changed=[shape.id],
                      data={"prompt": shape.prompt,
                            "note": "a prompt shows in an empty placeholder while editing, "
                                    "never in a show or an export"})
    if action == "set_footer":
        chosen = numbers(deck, slides) if slides is not None else None
        today = call.session.now().date()
        changed = deck.set_footer(text, slide_number=True if slide_number is None
                                  else slide_number, date=date, slides=chosen,
                                  skip_title=not on_title_slide, today=today)
        for sid in changed:
            call.touch(sid)
        if slides is None:
            for master in deck.masters:
                call.touch(master.id)
        scope = "all slides, masters and layouts" if slides is None else "the slides given"
        return Result(summary=f"Set the header and footer on {scope}; {len(changed)} slide(s) "
                      "changed", changed=[f"s:{sid}" for sid in changed],
                      data={"slides": [f"s:{sid}" for sid in changed]})
    if action == "add_layout":
        after = _template_layout(deck, target, "target") if target else None
        layout = deck.add_layout(text or "Custom Layout", after=after)
        call.touch(layout.id)
        return Result(summary=f"Added layout {layout.id} {layout.name!r}", created=[layout.id],
                      data={"layout": layout.id, "shapes": [
                          {"id": shape.id, "placeholder": shape.placeholder.type or "content"}
                          for shape in layout.shapes if shape.placeholder is not None],
                          "layouts": _layouts_json(deck)})
    layout = _template_layout(deck, target, "target")
    if action == "add_placeholder":
        existing = resolve(deck, target, field="target") if "." in str(target) else None
        return _add_placeholder(call, deck, layout, type, box, text, slides, existing)
    if action == "duplicate_layout":
        copy = deck.duplicate_layout(layout, name=text)
        call.touch(copy.id)
        return Result(summary=f"Duplicated {layout.id} as {copy.id} {copy.name!r}",
                      created=[copy.id], data={"layout": copy.id,
                                               "layouts": _layouts_json(deck)})
    if action == "rename_layout":
        if not text:
            raise ToolError("invalid_arguments", "rename_layout needs text, the new name",
                            field="text")
        renamed = deck.rename_layout(layout, text)
        return Result(summary=f"Renamed {layout.id} to {renamed.name!r}", changed=[layout.id],
                      data={"layouts": _layouts_json(deck)})
    if action == "delete_layout":
        to = _template_layout(deck, reassign_to, "reassign_to") if reassign_to else None
        moved = deck.delete_layout(layout, reassign_to=to)
        for sid in moved:
            call.touch(sid)
        data: dict[str, Any] = {"layouts": _layouts_json(deck)}
        if moved:
            data["moved"] = [f"s:{sid}" for sid in moved]
        return Result(summary=f"Deleted layout {layout.id} {layout.name!r}",
                      removed=[layout.id], data=data)
    raise ToolError("invalid_arguments", f"unknown action {action!r}", field="action")


_ADDABLE = ["body", "content", "pic", "title"]


def _add_placeholder(call, deck, layout, kind, box, name, slides, existing=None) -> Result:
    """ppt_edit_master add_placeholder: the layout's new placeholder, and an empty copy on
    the ``slides`` asked for (PowerPoint gives one only to slides made afterwards); with
    ``existing`` (a layout placeholder, ``L9.6``) only the copies."""
    if existing is not None:
        kind = existing.placeholder.type if existing.placeholder is not None else None
        if existing.placeholder is None or kind in ("dt", "ftr", "sldNum", "hdr"):
            raise ToolError("invalid_arguments", f"{existing.id} is not a layout placeholder "
                            "a slide fills", field="target")
        kind = {None: "content", "obj": "content", "ctrTitle": "title"}.get(kind, kind)
        if not slides:
            raise ToolError("invalid_arguments", f"{existing.id} is on {layout.id} already; "
                            "give slides to add it to", field="slides",
                            valid_options=[f"s:{s.slide_id}" for s in layout.slides][:50])
    elif kind is None:
        raise ToolError("invalid_arguments", "add_placeholder needs type", field="type",
                        valid_options=_ADDABLE)
    if kind not in _ADDABLE:
        raise ToolError("invalid_arguments", f"a layout placeholder of type {kind!r} is not "
                        "added here", field="type", valid_options=_ADDABLE)
    if box is None and existing is None:
        raise ToolError("invalid_arguments", "add_placeholder needs box, in points",
                        field="box")
    on_it = {slide.slide_id for slide in layout.slides}
    chosen = [deck.slides[n - 1] for n in numbers(deck, slides)] if slides else []
    strays = [f"s:{s.slide_id}" for s in chosen if s.slide_id not in on_it]
    if strays:
        raise ToolError("invalid_arguments", f"{', '.join(strays)} not on {layout.id}; a "
                        "slide takes only its own layout's placeholders", field="slides",
                        valid_options=[f"s:{sid}" for sid in sorted(on_it)][:50])
    if existing is None and kind == "title" and any(shape.placeholder.type in ("title", "ctrTitle")
                               for shape in layout.shapes if shape.placeholder):
        raise ToolError("invalid_arguments", f"{layout.id} has a title placeholder already; "
                        "add body or content", field="type", valid_options=_ADDABLE[:3])
    if existing is None:
        x, y, w, h = box_emu(box)
        if w <= 0 or h <= 0:
            raise ToolError("invalid_arguments", "box needs w and h over 0", field="box")
    with deck.batch():
        shape = existing or layout.add_placeholder(kind, x, y, w, h, name=name)
        idx = shape.placeholder.idx
        added = [s.add_layout_placeholder(idx if idx is not None else "title")
                 for s in chosen]
    call.touch(layout.id, shape.id)
    for copy in added:
        call.touch(copy._slide.slide_id, copy.id)
    def has_it(slide) -> bool:
        return any(s.placeholder is not None and (
            s.placeholder.idx == idx if idx is not None
            else s.placeholder.type in ("title", "ctrTitle")) for s in slide.shapes)

    without = [f"s:{s.slide_id}" for s in layout.slides if not has_it(s)]
    data: dict[str, Any] = {"placeholder": shape.id, "idx": idx, "type": kind,
                            "name": shape.name, "box": box_pt(shape.bounds)}
    if added:
        data["added_to"] = [copy.id for copy in added]
    if without:
        data["slides_without"] = without
    which = f"idx {idx}" if idx is not None else "the title"
    data["note"] = (f"a slide made on {layout.id} gets it empty ({which}): ppt_add_slide lists "
                    "its id; fill it with ppt_set_text"
                    + (" or ppt_add_picture target" if kind == "pic" else "")
                    + (f". Slides already on it lack it, as in PowerPoint: add_placeholder "
                       f"with target {shape.id} and slides adds it to them" if without else ""))
    made = [copy.id for copy in added]
    if existing is not None:
        return Result(summary=f"Added {shape.id} ({which}) to {len(made)} slide(s)",
                      created=made, data=data)
    return Result(summary=f"Added {kind} placeholder {shape.id} ({which}) to {layout.id}",
                  created=[shape.id, *made], data=data)


TOOLS = [ppt_add_slide, ppt_draft_slides, ppt_manage_slides, ppt_set_theme, ppt_edit_master]
