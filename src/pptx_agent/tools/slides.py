"""Slide tools: ``ppt_add_slide``, ``ppt_draft_slides``, ``ppt_manage_slides``,
``ppt_set_theme``.

Positions are slide numbers, from 1, as PowerPoint numbers them; a slide's lasting address
is its id, ``s:256``.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, integer, obj, string, tool

from ..edit.theme import MAPPED, SLOTS
from .common import KIND, box_pt, library_errors, slide as slide_at

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
        raise ToolError("invalid_arguments", f"position {at}; from 1 to {len(deck.slides) + 1}",
                        field=field)
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
      "Duplicate, move or delete a slide, or find one by its title. Returns the slide order.",
      {"doc": string("Document id."),
       "action": string("What to do.", enum=["duplicate", "move", "delete", "find_by_title"]),
       "slide": string("The slide, s:256 or $ref.", optional=True),
       "to": integer("move, duplicate: new position from 1. Default after it.", minimum=1,
                     maximum=10000, optional=True),
       "notes": string("duplicate: the copy's notes.", optional=True),
       "title": string("find_by_title: the title.", optional=True),
       "ref": string("duplicate: ref name for the copy.", optional=True)},
      kind=KIND, group=SLIDES, mutates=True, refs=("slide",))
@library_errors
def ppt_manage_slides(call, doc, action, slide=None, to=None, notes=None, title=None,
                      ref=None):
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
            raise ToolError("invalid_arguments", f"position {to}; from 1 to {len(deck.slides)}",
                            field="to")
        deck.move_slide(target, to - 1)
        return Result(summary=f"Moved s:{target.slide_id} to {to}",
                      changed=[f"s:{target.slide_id}"], data={"order": _order(deck)})
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


TOOLS = [ppt_add_slide, ppt_draft_slides, ppt_manage_slides, ppt_set_theme]
