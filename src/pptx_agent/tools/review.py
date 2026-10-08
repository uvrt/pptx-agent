"""``ppt_comments``: review comments on slides and shapes, as ``word_comments`` is for Word.

PowerPoint's modern comments (:mod:`pptx_agent.edit.comments`): a thread on a slide or a
shape, with replies, resolved or open.  The actions are ``word_comments``' -- list, add,
reply, resolve, reopen, edit, delete -- many at once in ``items``; a comment's address is
``c:`` and eight hex digits.  Comments never change what a slide shows, so no layout facts
come back; ``list`` only reads.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import Result, ToolError, array, boolean, obj, string, tool
from ooxml_edit.tools.results import page_list

from .common import KIND, library_errors, slide_id_of

REVIEW = "ppt_review"

_ITEM = obj({
    "target": string("add: the slide (s:256) or shape (256.5) it is on.", optional=True),
    "comment": string("The comment (c:...), except for add.", optional=True),
    "text": string("add, reply, edit: the text; \\n between paragraphs.", optional=True),
    "resolve": boolean("reply: also resolve the thread.", optional=True),
    "ref": string("add, reply: ref name for the new comment.", optional=True),
})


def _short(text: str, limit: int) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= limit else text[:limit - 1] + "…"


def thread_json(thread) -> dict[str, Any]:
    data = {"id": thread.address, "on": thread.target, "author": thread.author,
            "done": thread.done, "text": _short(thread.text, 300)}
    if thread.replies:
        data["replies"] = [{"id": r.address, "author": r.author, "text": _short(r.text, 200)}
                           for r in thread.replies]
    return data


def _need(value, field: str):
    if value in (None, ""):
        raise ToolError("invalid_arguments", f"{field} is required for this action",
                        field=field)
    return value


@tool("ppt_comments",
      "Review comment threads on slides and shapes: list, add, reply, resolve, reopen, edit, "
      "delete, many at once. Comments do not change the slides.",
      {"doc": string("Document id."),
       "action": string("What to do.", enum=["list", "add", "reply", "resolve", "reopen",
                                             "edit", "delete"]),
       "items": array(_ITEM, "The comments to act on (not list).", optional=True),
       "author": string("add, reply: the author's name.", optional=True),
       "slides": array(string(), "list: only these slides (s:256). Default all.",
                       optional=True),
       "open_only": boolean("list: only unresolved threads.", optional=True),
       "cursor": string("next_cursor of the previous page.", optional=True)},
      kind=KIND, group=REVIEW, mutates=True, refs=("items[].target", "items[].comment"),
      reads=lambda arguments: arguments.get("action") == "list")
@library_errors
def ppt_comments(call, doc, action, items=None, author=None, slides=None, open_only=False,
                 cursor=None):
    deck = call.document
    if action == "list":
        chosen = None
        if slides:
            from .common import slide as slide_at

            chosen = [slide_at(deck, s, field="slides") for s in slides]
        threads = [t for t in (deck.comments() if chosen is None else
                               [t for s in chosen for t in deck.comments(s)])
                   if not (open_only and t.done)]
        shown, total, next_cursor = page_list([thread_json(t) for t in threads], cursor=cursor,
                                              limit=call.limits.max_list_items)
        result = Result(summary=f"{total} thread(s)", data=shown, next_cursor=next_cursor)
        result.total = total
        return result
    if not items:
        raise ToolError("invalid_arguments", f"{action} needs items", field="items")
    now = call.now()
    made, changed, removed = [], [], []
    for index, item in enumerate(items):
        where = f"items[{index}]"
        if action == "add":
            target = _need(item.get("target"), f"{where}.target")
            text = _need(item.get("text"), f"{where}.text")
            thread = deck.add_comment(_target(deck, target, f"{where}.target"), text,
                                      author=_need(author, "author"), date=now)
            made.append(thread.address)
            if item.get("ref"):
                call.define_ref(item["ref"], thread.address)
            continue
        identifier = _need(item.get("comment"), f"{where}.comment")
        _comment(deck, identifier, f"{where}.comment")
        if action == "reply":
            reply = deck.reply_to_comment(identifier, _need(item.get("text"), f"{where}.text"),
                                          author=_need(author, "author"), date=now)
            made.append(reply.address)
            if item.get("ref"):
                call.define_ref(item["ref"], reply.address)
            if item.get("resolve"):
                deck.resolve_comment(identifier)
                changed.append(identifier)
        elif action == "resolve":
            changed.append(deck.resolve_comment(identifier).address)
        elif action == "reopen":
            changed.append(deck.reopen_comment(identifier).address)
        elif action == "edit":
            deck.edit_comment(identifier, _need(item.get("text"), f"{where}.text"))
            changed.append(identifier)
        else:
            deck.delete_comment(identifier)
            removed.append(identifier)
    return Result(summary=f"{action}: {len(items)} comment(s)", created=made, changed=changed,
                  removed=removed)


def _target(deck, target: str, field: str):
    text = str(target)
    if text.startswith("s:") or text.isdigit():
        from .common import slide as slide_at

        return slide_at(deck, text, field=field)
    from .common import shape

    if slide_id_of(text) is None:
        raise ToolError("invalid_arguments", f"{target!r} is not a slide or shape address",
                        field=field)
    return shape(deck, text, field=field)


def _comment(deck, identifier: str, field: str) -> None:
    try:
        deck.comment(identifier)
    except KeyError:
        raise ToolError("not_found", f"no comment {identifier!r}", field=field,
                        valid_options=[t.address for t in deck.comments()][:50]) from None


TOOLS = [ppt_comments]
