"""Markdown in and out: the dialect, a normalised AST, and inline text written so it reads
back as meant.

The dialect is docx-agent's (its ``markdown/parse.py``): CommonMark plus GFM tables and
strikethrough, parsed by markdown-it-py, which is used behind this module only.  Footnotes,
which docx-agent also reads, are left out: a slide has nowhere to put one, so ``[^1]`` stays
text.

**The normalised AST** is nested tuples, equal when two outlines mean the same thing:

* blocks -- ``("comment", text)`` for an HTML block that is only comments (one per
  comment), ``("html", text)``, ``("p", inline, first_line)``, ``("h", level, inline)``,
  ``("code", text)``, ``("quote", blocks)``, ``("list", ordered, start, items)`` (an item
  is a tuple of blocks), ``("table", header, rows)``, ``("hr",)``;
* inline -- ``("t", text, marks, href)`` with adjacent alike texts merged (``marks`` a
  frozenset of ``strong``, ``em``, ``strike``, ``code``), ``("br",)`` (a hard break, or
  ``<br>``), ``("html", text)``, ``("img", alt, src, title)``.

``first_line`` is the paragraph's first source line, stripped: the speaker-notes marker is
recognised on it, so an escaped ``Notes\\:`` is text.  What normalisation forgets, because
Markdown does not keep it: list markers and delimiters, emphasis delimiters, whitespace at
the ends of a block or a line, and a paragraph that is empty.

The writing half (:func:`inline_markdown`) escapes text so it reads back as text, keeps
delimiters off whitespace, and checks itself: a block's inline Markdown is parsed back and,
when CommonMark's flanking rules would read it otherwise, strikethrough and then emphasis
are dropped from that block -- the text never changes.  docx-agent's renderer does the same.
"""

from __future__ import annotations

import re
from functools import lru_cache

from markdown_it import MarkdownIt

MARKS = ("strike", "strong", "em", "code")
_DELIMITERS = {"strong": "**", "em": "*", "strike": "~~"}
_COMMENT = re.compile(r"<!--(.*?)-->", re.DOTALL)
_BREAK = re.compile(r"<br\s*/?>", re.IGNORECASE)
_ALWAYS = frozenset("\\`*_[]<~")
_ENTITY = re.compile(r"&(?:#[0-9]{1,7}|#[xX][0-9a-fA-F]{1,6}|[A-Za-z][A-Za-z0-9]{1,31});")
#: What starts a block at the beginning of a line -- and only that: CommonMark wants
#: whitespace (or the end of the line) after a list marker or an ATX heading's hashes, and a
#: thematic break or setext underline is the whole line.  ``11.9%`` and ``+0.6pt`` are text
#: as they stand.  The end of a text item counts as the end of the line, because the next
#: item may start with a space.
_LINE_START = re.compile(r"(>|#+(?=[ \t]|$)|[+-](?=[ \t]|$)|-[- \t]*$|=[= \t]*$)"
                         r"|(\d{1,9})([.)])(?=[ \t]|$)")


@lru_cache(maxsize=1)
def parser() -> MarkdownIt:
    """CommonMark, GFM tables and strikethrough."""
    return MarkdownIt("commonmark").enable(["table", "strikethrough"])


# -- Markdown -> AST -------------------------------------------------------------------------


def parse(text: str) -> tuple:
    """``text``'s normalised AST (see the module docstring)."""
    lines = text.splitlines()
    tokens = parser().parse(text)
    blocks, _ = _blocks(tokens, 0, None, lines)
    return tuple(blocks)


def _blocks(tokens, i: int, until: str | None, lines: list[str]) -> tuple[list, int]:
    out: list = []
    while i < len(tokens):
        token = tokens[i]
        kind = token.type
        if until is not None and kind == until:
            return out, i + 1
        if kind == "html_block":
            out.extend(_html_block(token.content))
            i += 1
        elif kind == "paragraph_open":
            inline = _inline(tokens[i + 1].children or [])
            first = lines[token.map[0]].strip() if token.map else ""
            if inline:
                out.append(("p", inline, first))
            i += 3
        elif kind == "heading_open":
            out.append(("h", int(token.tag[1]), _inline(tokens[i + 1].children or [])))
            i += 3
        elif kind in ("fence", "code_block"):
            out.append(("code", token.content))
            i += 1
        elif kind == "blockquote_open":
            inner, i = _blocks(tokens, i + 1, "blockquote_close", lines)
            out.append(("quote", tuple(inner)))
        elif kind in ("bullet_list_open", "ordered_list_open"):
            ordered = kind == "ordered_list_open"
            start = int(token.attrGet("start") or 1) if ordered else 1
            close = kind.replace("open", "close")
            items = []
            i += 1
            while tokens[i].type != close:
                inner, i = _blocks(tokens, i + 1, "list_item_close", lines)
                items.append(tuple(inner))
            out.append(("list", ordered, start, tuple(items)))
            i += 1
        elif kind == "table_open":
            table, i = _table(tokens, i + 1)
            out.append(table)
        elif kind == "hr":
            out.append(("hr",))
            i += 1
        else:
            i += 1
    return out, i


def _html_block(content: str) -> list:
    stripped = content.strip()
    comments = _COMMENT.findall(stripped)
    if comments and _COMMENT.sub("", stripped).strip() == "":
        return [("comment", c.strip()) for c in comments]
    return [("html", content.rstrip("\n"))]


def _table(tokens, i: int) -> tuple[tuple, int]:
    header: list = []
    rows: list = []
    current: list | None = None
    while tokens[i].type != "table_close":
        token = tokens[i]
        if token.type == "tr_open":
            current = []
        elif token.type == "tr_close":
            (rows if header else header).append(tuple(current))
            current = None
        elif token.type in ("th_open", "td_open"):
            current.append(_inline(tokens[i + 1].children or []))
            i += 2
        i += 1
    return ("table", header[0] if header else (), tuple(rows)), i + 1


def _inline(children) -> tuple:
    items: list = []
    marks: list[str] = []
    hrefs: list[str] = []
    for token in children:
        kind = token.type
        href = hrefs[-1] if hrefs else None
        if kind == "text":
            items.append(("t", token.content, frozenset(marks), href))
        elif kind == "code_inline":
            items.append(("t", token.content, frozenset(marks + ["code"]), href))
        elif kind == "softbreak":
            items.append(("t", "\n", frozenset(marks), href))
        elif kind == "hardbreak":
            items.append(("br",))
        elif kind == "html_inline":
            items.append(("br",) if _BREAK.fullmatch(token.content.strip()) else ("html", token.content))
        elif kind == "strong_open":
            marks.append("strong")
        elif kind == "em_open":
            marks.append("em")
        elif kind == "s_open":
            marks.append("strike")
        elif kind in ("strong_close", "em_close", "s_close"):
            if marks:
                marks.pop()
        elif kind == "link_open":
            hrefs.append(token.attrGet("href"))
        elif kind == "link_close":
            if hrefs:
                hrefs.pop()
        elif kind == "image":
            alt = "".join(c.content for c in token.children or [] if c.type in ("text", "code_inline"))
            items.append(("img", alt, token.attrGet("src") or "", token.attrGet("title")))
    return normalise_inline(items)


def normalise_inline(items: list) -> tuple:
    """Merge alike texts, drop empty ones, and forget whitespace at the ends of the block
    and around hard breaks (Markdown keeps none of it).  Tabs count as spaces; a soft line
    break is a space."""
    merged: list = []
    for item in items:
        if item[0] == "t":
            text = item[1].replace("\t", " ").replace("\n", " ")
            if merged and merged[-1][0] == "t" and merged[-1][2:] == item[2:]:
                merged[-1] = ("t", merged[-1][1] + text) + item[2:]
                continue
            merged.append(("t", text) + item[2:])
        else:
            merged.append(item)
    for k, item in enumerate(merged):
        if item[0] != "t":
            continue
        text = item[1]
        if k == 0 or merged[k - 1][0] == "br":
            text = text.lstrip(" ")
        if k == len(merged) - 1 or merged[k + 1][0] == "br":
            text = text.rstrip(" ")
        merged[k] = ("t", text) + item[2:]
    final: list = []
    for item in merged:
        if item[0] == "t" and not item[1]:
            continue
        if final and item[0] == "t" and final[-1][0] == "t" and final[-1][2:] == item[2:]:
            final[-1] = ("t", final[-1][1] + item[1]) + item[2:]
        else:
            final.append(item)
    while final and final[-1][0] == "br":
        final.pop()
    while final and final[0][0] == "br":
        final.pop(0)
    return tuple(final)


def plain(inline: tuple) -> str:
    """The characters of an inline sequence: text, and a line break as ``\\n``."""
    return "".join(i[1] if i[0] == "t" else "\n" if i[0] == "br" else "" for i in inline)


# -- AST -> Markdown (inline) ----------------------------------------------------------------


def inline_markdown(items: list, *, line_start: bool = True, table: bool = False,
                    heading: bool = False) -> str:
    """``items`` (``("t", text, marks, href)`` and ``("br",)``) as Markdown that reads back
    as ``normalise_inline(items)``, dropping marks the dialect cannot place (see above).

    In a table cell or a heading -- one line each -- a break is written ``<br>``."""
    items = list(normalise_inline(items))
    one_line = table or heading
    for dropped in ((), ("strike",), ("strike", "strong", "em")):
        if dropped:
            items = [("t", i[1], i[2] - set(dropped), i[3]) if i[0] == "t" else i for i in items]
        text = _write_inline(items, line_start=line_start, table=table, one_line=one_line)
        if not any(i[0] == "t" and i[2] - {"code"} for i in items):
            return text
        if _reads_back(text, items, one_line=one_line):
            return text
    return text


def _reads_back(text: str, items: list, *, one_line: bool) -> bool:
    tokens = parser().parseInline(text)
    children = tokens[0].children if tokens else []
    return _inline(children or []) == normalise_inline(items)


def _write_inline(items: list, *, line_start: bool, table: bool, one_line: bool) -> str:
    out: list[str] = []
    open_marks: list[str] = []
    href: str | None = None
    at_start = line_start

    def close_to(keep: frozenset) -> None:
        while open_marks and not set(open_marks) <= keep:
            out.append(_DELIMITERS[open_marks.pop()])

    def close_link() -> None:
        nonlocal href
        if href is not None:
            close_to(frozenset())
            out.append(f"]({destination(href)})")
            href = None

    spaced = _spaced(items)
    for k, item in enumerate(spaced):
        item_href = item[3] if item[0] == "t" else None
        if item_href != href:
            close_link()
            if item_href is not None:
                out.append("[")
                href = item_href
        if item[0] == "t":
            marks = frozenset(item[2]) - {"code"}
            close_to(marks)
            for mark in sorted(marks - set(open_marks), key=_order(spaced, k)):
                out.append(_DELIMITERS[mark])
                open_marks.append(mark)
            if "code" in item[2]:
                out.append(_code(item[1]))
            else:
                text = escape(item[1], table=table)
                if at_start:
                    text = _escape_line_start(text)
                out.append(text)
            at_start = at_start and not item[1].strip()
            continue
        close_to(frozenset())
        if item[0] == "br":
            close_link()
            out.append("<br>" if one_line else "\\\n")
            at_start = not one_line
    close_link()
    close_to(frozenset())
    return "".join(out)


def _order(items: list, k: int):
    """Marks opened together nest by how long they last: the longest outermost."""
    def lasts(mark: str) -> tuple:
        n = k
        while n < len(items) and items[n][0] == "t" and mark in items[n][2]:
            n += 1
        return (-n, MARKS.index(mark))
    return lasts


def _spaced(items: list) -> list:
    """Whitespace moved out of marked text, so no delimiter touches it."""
    out: list = []
    for item in items:
        if item[0] == "t" and item[2] - {"code"} and item[1]:
            core = item[1].strip(" ")
            if not core:
                out.append(("t", item[1], item[2] & {"code"}, item[3]))
                continue
            lead = item[1][: len(item[1]) - len(item[1].lstrip(" "))]
            trail = item[1][len(item[1].rstrip(" ")):]
            if lead:
                out.append(("t", lead, item[2] & {"code"}, item[3]))
            out.append(("t", core, item[2], item[3]))
            if trail:
                out.append(("t", trail, item[2] & {"code"}, item[3]))
        else:
            out.append(item)
    return out


def _longest_run(text: str, char: str) -> int:
    runs = re.findall(re.escape(char) + "+", text)
    return max((len(r) for r in runs), default=0)


def _code(text: str) -> str:
    fence = "`" * (_longest_run(text, "`") + 1)
    pad = text.startswith("`") or text.endswith("`") or (
        text.startswith(" ") and text.endswith(" ") and text.strip(" ") != "")
    return f"{fence} {text} {fence}" if pad else f"{fence}{text}{fence}"


def escape(text: str, *, table: bool = False) -> str:
    """Text that Markdown reads back as itself (inline)."""
    out = []
    for k, char in enumerate(text):
        if char in _ALWAYS or (table and char == "|"):
            out.append("\\" + char)
        elif char == "&" and _ENTITY.match(text, k):
            out.append("\\&")
        elif char == "\n":
            out.append(" ")
        else:
            out.append(char)
    return "".join(out)


def _escape_line_start(text: str) -> str:
    stripped = text.lstrip(" \t")
    match = _LINE_START.match(stripped)
    if match is None:
        return stripped
    if match.group(1):
        return "\\" + stripped
    return match.group(2) + "\\" + match.group(3) + stripped[match.end():]


def destination(url: str) -> str:
    if re.search(r"[\s()<>]", url):
        return "<" + url.replace("<", "%3C").replace(">", "%3E") + ">"
    return url


def comment(text: str) -> str:
    """An HTML comment holding ``text``, which cannot end it early (docx-agent's form)."""
    return "<!-- " + comment_text(text) + " -->"


def comment_text(text: str) -> str:
    return text.replace("--", "- -").replace(">", "&gt;")


__all__ = ["comment", "comment_text", "destination", "escape", "inline_markdown",
           "normalise_inline", "parse", "parser", "plain"]
