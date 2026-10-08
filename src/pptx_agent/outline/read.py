"""``to_outline``: the deck as a compact Markdown reading view (see :mod:`pptx_agent.outline`).

Nothing here writes: the trees are only read, the ids are the ones the deck has now, and
no shape is stamped.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING

from ..edit import notes as _notes
from ..oxml.package import REL_SLIDE_LAYOUT, REL_SLIDE_MASTER
from ..oxml.xml import Element, find, local_name, qn
from .markdown import comment, destination, escape, inline_markdown

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, Shape, Slide

#: Typefaces read back as inline code; :data:`CODE_FONT` is what the drafting half writes.
CODE_FONT = "Consolas"
MONOSPACE = frozenset({"consolas", "courier new", "courier", "menlo", "monaco",
                       "source code pro", "cascadia code", "cascadia mono", "lucida console",
                       "sf mono", "roboto mono", "dejavu sans mono", "fira code", "fira mono",
                       "jetbrains mono", "ibm plex mono", "andale mono", "noto sans mono"})
TITLE_TYPES = ("title", "ctrTitle")
#: Placeholder types whose list style comes from the master's ``bodyStyle``.
_BODY_TYPES = frozenset({"body", "subTitle", "obj"})
#: Placeholder type -> the master placeholder it inherits from (as Slide._inherited_xfrm).
_MASTER_KIND = {"title": "title", "ctrTitle": "title", "dt": "dt", "ftr": "ftr",
                "sldNum": "sldNum", "hdr": "hdr"}
_BULLETS = {qn("a:buNone"): None, qn("a:buChar"): "bullet", qn("a:buBlip"): "bullet",
            qn("a:buAutoNum"): "number"}
NOTES_MARKER = "Notes:"
_SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")


# -- the outline -----------------------------------------------------------------------------


def to_outline(document: "Document", slides=None, *, ids: bool = True, notes: bool = True) -> str:
    """See :meth:`pptx_agent.edit.document.Document.to_outline`."""
    chosen = select(document, slides)
    out: list[str] = []
    for slide in chosen:
        out += _slide(document, slide, ids=ids, notes=notes)
    while out and out[-1] == "":
        out.pop()
    return "\n".join(out) + "\n" if out else ""


@dataclass(frozen=True)
class TextBlock:
    r"""One piece of a deck's text, raw -- not Markdown, nothing escaped.

    ``address`` is what :meth:`Document.resolve` takes; ``text`` is what that object's
    ``text`` setter (``set_text``) takes back unchanged: paragraphs joined by ``"\n"``, a
    line break as ``"\v"`` -- the object's ``text_frame.text``, and its ``text`` but for a
    shape's line breaks, which ``Shape.text`` reads as ``"\n"``.  ``kind`` is
    ``"title"``, ``"text"``, ``"cell"``, ``"smartart"`` (a diagram node, ``<id>/node<k>``)
    or ``"notes"`` (the slide's speaker notes, ``<sldId>/notes``, a :class:`TextFrame`);
    ``slide`` is the 1-based slide number::

        for block in deck.outline_blocks():
            print(block.address, block.text)      # 256.29 11.9%
    """

    address: str
    kind: str
    text: str
    slide: int


def outline_blocks(document: "Document", slides=None) -> list[TextBlock]:
    """See :meth:`pptx_agent.edit.document.Document.outline_blocks`."""
    out: list[TextBlock] = []
    for slide in select(document, slides):
        number = slide.index + 1
        shapes = _flatten(slide.shapes)
        title = next((s for s in shapes if s.kind == "shape" and s.placeholder is not None
                      and s.placeholder[0] in TITLE_TYPES), None)
        ordered = ([title] if title is not None else []) + [s for s in shapes if s is not title]
        for shape in ordered:
            kind = local_name(shape._element)
            if kind == "sp":
                text = shape.text_frame.text
                if text:
                    out.append(TextBlock(shape.id, "title" if shape is title else "text", text,
                                         number))
            elif kind == "graphicFrame" and shape.has_table:
                table = shape.table
                for cell in table.iter_cells():
                    if cell.is_spanned:
                        continue
                    text = cell.text
                    if text:
                        out.append(TextBlock(cell.address, "cell", text, number))
            elif kind == "graphicFrame" and shape.has_diagram:
                for k, node in enumerate(shape.diagram.nodes):
                    if node.text:
                        out.append(TextBlock(f"{shape.id}/node{k}", "smartart", node.text,
                                             number))
        notes = slide.notes
        if notes:
            out.append(TextBlock(f"{slide.slide_id}/notes", "notes", notes, number))
    return out


def find_text(document: "Document", needle, slides=None) -> list[TextBlock]:
    """See :meth:`pptx_agent.edit.document.Document.find_text`."""
    if isinstance(needle, re.Pattern):
        return [b for b in outline_blocks(document, slides) if needle.search(b.text)]
    if not isinstance(needle, str) or not needle:
        raise TypeError("find_text takes a non-empty string or a compiled re.Pattern")
    return [b for b in outline_blocks(document, slides) if needle in b.text]


def select(document: "Document", slides) -> list["Slide"]:
    """``None``: every slide.  Otherwise slides, 1-based slide numbers (as ``render_svg``
    takes -- slide 1 is ``deck.slides[0]``), :class:`Slide` objects or ``"s:<sldId>"``
    ids, in the order given."""
    if slides is None:
        return list(document.slides)
    if isinstance(slides, (int, str)) or hasattr(slides, "slide_id"):
        slides = [slides]
    every = document.slides
    chosen = []
    for item in slides:
        if hasattr(item, "slide_id"):
            chosen.append(document.slide(item.slide_id))
        elif isinstance(item, str):
            if not item.startswith("s:"):
                raise ValueError(f"a slide is a 1-based number, a Slide or 's:<sldId>', "
                                 f"not {item!r}")
            chosen.append(document.slide(item))
        elif isinstance(item, int) and not isinstance(item, bool):
            if not 1 <= item <= len(every):
                hint = (" -- slide numbers here are 1-based, so slide 1 is deck.slides[0]; "
                        "pass the Slide itself or 's:<sldId>' to leave no doubt")
                if item >= 256 and any(s.slide_id == item for s in every):
                    hint = f" -- {item} is an sldId: pass 's:{item}'"
                raise IndexError(f"slide number {item} out of range 1..{len(every)}{hint}")
            chosen.append(every[item - 1])
        else:
            raise TypeError(f"not a slide: {item!r}; pass a 1-based number, a Slide or "
                            "'s:<sldId>'")
    return chosen


def _slide(document: "Document", slide: "Slide", *, ids: bool, notes: bool) -> list[str]:
    package = document.package
    root = package.tree(slide.part_path)
    layout = slide.layout
    words = [f"s:{slide.slide_id}"] if ids else []
    if root is not None and root.get("show") in ("0", "false"):
        words.append("hidden")
    if layout is not None:
        words.append(f"layout: {layout.name}")
    lines: list[str] = []
    if ids and words:
        lines.append(comment(" ".join(words)))

    shapes = _flatten(slide.shapes)
    title = next((s for s in shapes if s.kind == "shape" and s.placeholder is not None
                  and s.placeholder[0] in TITLE_TYPES), None)
    heading = "#"
    if title is not None:
        inline = _joined(_paragraph_items(document, slide.part_path, title._element))
        text = inline_markdown(inline, line_start=False, heading=True)
        if text.endswith("#"):
            text = text[:-1] + "\\#"
        heading = ("# " + text).rstrip()
        if ids:
            heading += " " + comment(title.id)
    lines += [heading, ""]

    first = True
    for shape in shapes:
        if shape is title:
            continue
        block = _shape_block(document, slide, shape, ids=ids)
        if block is None:
            continue
        label, body = block
        if ids:
            lines.append(comment(label))
        elif not first:
            lines += ["---", ""]
        if body:
            lines += body
        lines.append("")
        first = False

    if notes:
        notes_lines = _notes_block(document, slide)
        if notes_lines:
            if ids:
                lines.append(comment(f"{slide.slide_id}/notes"))
            lines += [NOTES_MARKER, ""] + notes_lines + [""]
    return lines


def _flatten(shapes: list["Shape"]) -> list["Shape"]:
    out = []
    for shape in shapes:
        if shape.kind == "group":
            out += _flatten(shape.children)
        else:
            out.append(shape)
    return out


# -- shapes ----------------------------------------------------------------------------------


def placeholder_label(placeholder: tuple) -> str:
    kind, idx = placeholder
    return f"ph: {kind or 'obj'}" + ("" if idx is None else f" {idx}")


def _shape_block(document: "Document", slide: "Slide", shape: "Shape", *,
                 ids: bool) -> tuple[str, list[str]] | None:
    """``(comment label, lines)`` for a shape, or ``None`` for one the outline leaves out
    (connectors, shapes without text, OLE objects)."""
    element = shape._element
    kind = local_name(element)
    placeholder = shape.placeholder
    prefix = [shape.id] if ids else []
    if kind == "sp":
        lines = _text_lines(document, slide, shape)
        if not lines:
            if placeholder is None:
                return None
            return " ".join(prefix + [placeholder_label(placeholder), "empty"]), []
        words = prefix + ([placeholder_label(placeholder)] if placeholder is not None else [])
        return " ".join(words), lines
    if kind == "pic":
        return " ".join(prefix + ["picture"] + ([placeholder_label(placeholder)]
                                                if placeholder is not None else [])), \
            [_picture(document, slide, shape)]
    if kind == "graphicFrame":
        if shape.has_table:
            words = prefix + ["table"] + ([placeholder_label(placeholder)]
                                          if placeholder is not None else [])
            return " ".join(words), _table(document, slide, shape)
        if shape.has_chart:
            return " ".join(prefix + ["chart"]), [chart_summary(shape)]
        if shape.has_diagram:
            nodes = shape.diagram.model.get("nodes", [])
            return " ".join(prefix + ["smartart"]), _smartart(nodes)
    return None


def _picture(document: "Document", slide: "Slide", shape: "Shape") -> str:
    properties = find(shape._element, "p:nvPicPr/p:cNvPr")
    alt = (properties.get("descr") or "") if properties is not None else ""
    source = shape.image_part or ""
    if not source:
        blip = shape._blip()
        link = blip.get(qn("r:link")) if blip is not None else None
        relationship = document.package.relationships(slide.part_path).get(link or "")
        source = relationship.target if relationship is not None else ""
    return f"![{escape(' '.join(alt.split()))}]({destination(source)})"


def chart_summary(shape: "Shape") -> str:
    """``[chart: column; title: ...; series: A, B; categories: Q1, Q2]``."""
    chart = shape.chart
    root = chart._root()
    types = []
    for name in chart.chart_types:
        if name == "bar":
            direction = root.find(f".//{qn('c:barDir')}")
            name = "column" if direction is not None and direction.get("val") == "col" else "bar"
        if name not in types:
            types.append(name)
    words = [" + ".join(types) or "chart"]
    if chart.title:
        words.append("title: " + _summary_text(chart.title))
    names = [_summary_text(s.name or "") for s in chart.series]
    if names:
        words.append("series: " + ", ".join(names))
    categories = [_summary_text("" if c is None else str(c)) for c in chart.categories]
    if categories:
        words.append("categories: " + ", ".join(categories))
    return "[chart: " + "; ".join(words) + "]"


def _summary_text(text: str) -> str:
    return escape(" ".join(str(text).split()))


def _smartart(nodes: list[dict]) -> list[str]:
    lines = []
    depth = 0
    for k, node in enumerate(nodes):
        level = int(node.get("lvl") or 0)
        level = min(level, depth + 1) if k else 0
        depth = level
        text = inline_markdown([("t", node.get("t") or "", frozenset(), None)])
        lines.append("  " * level + "- " + text if text else "  " * level + "-")
    return lines


def _table(document: "Document", slide: "Slide", shape: "Shape") -> list[str]:
    tbl = shape._element.find(f".//{qn('a:tbl')}")
    rows = []
    for row in tbl.findall(qn("a:tr")):
        cells = []
        for cell in row.findall(qn("a:tc")):
            merged = cell.get("hMerge") in ("1", "true") or cell.get("vMerge") in ("1", "true")
            body = cell.find(qn("a:txBody"))
            if merged or body is None:
                cells.append("")
                continue
            items = _joined([inline for _, _, _, inline in
                             paragraphs(document, slide.part_path, body, [])])
            cells.append(inline_markdown(items, line_start=False, table=True))
        rows.append(cells)
    if not rows:
        return []
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]

    def line(cells: list[str]) -> str:
        return "| " + " | ".join(cells) + " |"

    return [line(rows[0]), "| " + " | ".join(["---"] * width) + " |"] + [line(r) for r in rows[1:]]


# -- text ------------------------------------------------------------------------------------


def _joined(paragraph_items: list[tuple]) -> list:
    """Paragraphs as one line of inline items, separated by breaks (a heading, a cell)."""
    out: list = []
    for inline in paragraph_items:
        if not inline:
            continue
        if out:
            out.append(("br",))
        out += list(inline)
    return out


def _paragraph_items(document: "Document", part: str, element: Element) -> list[tuple]:
    body = element.find(qn("p:txBody"))
    if body is None:
        return []
    return [inline for _, _, _, inline in paragraphs(document, part, body, [])]


def _text_lines(document: "Document", slide: "Slide", shape: "Shape") -> list[str]:
    body = shape._element.find(qn("p:txBody"))
    if body is None:
        return []
    found = paragraphs(document, slide.part_path, body, style_chain(document, slide, shape))
    return text_block(found)


def text_block(found: list[tuple]) -> list[str]:
    """Paragraphs ``(kind, level, start, inline)`` as Markdown lines: lists nested by level,
    other paragraphs separated by blank lines."""
    lines: list[str] = []
    stack: list[tuple[bool, int, int]] = []  # (ordered, content column, next number)
    for kind, level, start, inline in found:
        if not inline:
            continue
        if kind == "p":
            if lines:
                lines.append("")
            stack = []
            rendered = _indented(inline_markdown(list(inline)), "", "")
            # A paragraph that says only "Notes:" would start the speaker notes.
            if rendered[0].strip().casefold() == NOTES_MARKER.casefold():
                rendered[0] = rendered[0].replace(":", "\\:")
            lines += rendered
            continue
        ordered = kind == "number"
        level = min(level, len(stack))
        if stack and level < len(stack):
            del stack[level + 1:]
            if stack[level][0] != ordered:
                del stack[level:]
        if not stack and lines and lines[-1] != "":
            lines.append("")
        column = stack[level - 1][1] if level > 0 and len(stack) >= level else 0
        if level < len(stack):
            number = stack[level][2]
        else:
            number = start
        marker = f"{number}. " if ordered else "- "
        text = inline_markdown(list(inline))
        lines += _indented(text, " " * column + marker, " " * (column + len(marker)))
        entry = (ordered, column + len(marker), number + 1)
        if level < len(stack):
            stack[level] = entry
        else:
            stack.append(entry)
    return lines


def _indented(text: str, first: str, rest: str) -> list[str]:
    parts = text.split("\n")
    return [(first + parts[0]).rstrip()] + [(rest + p).rstrip() for p in parts[1:]]


def paragraphs(document: "Document", part: str, body: Element, chain: list) -> list[tuple]:
    """``(kind, level, start, inline)`` per ``a:p``: ``kind`` ``"p"``, ``"bullet"`` or
    ``"number"`` by the bullet the paragraph shows, through ``chain`` (list-style elements
    to consult, nearest first)."""
    relationships = document.package.relationships(part)
    out = []
    for paragraph in body.findall(qn("a:p")):
        properties = paragraph.find(qn("a:pPr"))
        level = int(properties.get("lvl") or 0) if properties is not None else 0
        level = max(0, min(level, 8))
        kind, start = _bullet(properties, level, chain)
        out.append((kind, level, start, tuple(_runs(paragraph, relationships))))
    return out


def _bullet(properties: Element | None, level: int, chain: list) -> tuple[str, int]:
    sources = [properties] + [style.find(qn(f"a:lvl{level + 1}pPr")) for style in chain
                              if style is not None]
    for source in sources:
        if source is None:
            continue
        for child in source:
            if child.tag in _BULLETS:
                kind = _BULLETS[child.tag]
                if kind is None:
                    return "p", 1
                start = int(child.get("startAt") or 1) if kind == "number" else 1
                return kind, start
    return "p", 1


def _runs(paragraph: Element, relationships: dict) -> list:
    items: list = []
    for node in paragraph:
        tag = local_name(node)
        if tag == "br":
            items.append(("br",))
        elif tag in ("r", "fld"):
            text_node = node.find(qn("a:t"))
            text = (text_node.text or "") if text_node is not None else ""
            if not text:
                continue
            properties = node.find(qn("a:rPr"))
            items.append(("t", text.replace("\v", "\n")) + _marks(properties, relationships))
    return items


def _marks(properties: Element | None, relationships: dict) -> tuple[frozenset, str | None]:
    if properties is None:
        return frozenset(), None
    marks = set()
    if properties.get("b") in ("1", "true"):
        marks.add("strong")
    if properties.get("i") in ("1", "true"):
        marks.add("em")
    if properties.get("strike") in ("sngStrike", "dblStrike"):
        marks.add("strike")
    latin = properties.find(qn("a:latin"))
    if latin is not None and (latin.get("typeface") or "").casefold() in MONOSPACE:
        marks.add("code")
    href = None
    link = properties.find(qn("a:hlinkClick"))
    if link is not None and not link.get("action"):
        relationship = relationships.get(link.get(qn("r:id")) or "")
        if relationship is not None and relationship.is_external and _SCHEME.match(
                relationship.target or ""):
            href = relationship.target
    return frozenset(marks), href


# -- inheritance -----------------------------------------------------------------------------


def style_chain(document: "Document", slide: "Slide", shape: "Shape") -> list:
    """The list styles a slide shape's paragraphs inherit from, nearest first: its own, its
    layout and master placeholders', the master's text style for its kind, and the
    presentation's default -- as pptx2svg resolves them."""
    package = document.package
    chain = [shape._element.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")]
    placeholder = shape.placeholder
    layouts = package.related_parts_of_type(slide.part_path, REL_SLIDE_LAYOUT)
    masters = package.related_parts_of_type(layouts[0], REL_SLIDE_MASTER) if layouts else []
    master = package.tree(masters[0]) if masters else None
    if placeholder is not None:
        layout_shape = slide._layout_placeholder(placeholder)
        if layout_shape is not None:
            chain.append(layout_shape.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"))
            node = find(layout_shape, "p:nvSpPr/p:nvPr/p:ph")
            kind = node.get("type") if node is not None else placeholder[0]
        else:
            kind = placeholder[0]
        master_kind = _MASTER_KIND.get(kind, "body")
        tree = find(master, "p:cSld/p:spTree") if master is not None else None
        for candidate in [] if tree is None else tree:
            ph = candidate.find(f".//{qn('p:ph')}")
            if ph is not None and _MASTER_KIND.get(ph.get("type"), "body") == master_kind:
                chain.append(candidate.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"))
                break
        style = "titleStyle" if kind in TITLE_TYPES else "bodyStyle" \
            if (kind or "obj") in _BODY_TYPES else "otherStyle"
    else:
        style = "otherStyle"
    if master is not None:
        chain.append(find(master, f"p:txStyles/p:{style}"))
    presentation = package.tree(package.presentation_part())
    if presentation is not None:
        chain.append(find(presentation, "p:defaultTextStyle"))
    return chain


def notes_chain(document: "Document", body_shape: Element) -> list:
    package = document.package
    chain = [body_shape.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}")]
    master = _notes.notes_master(package)
    if master is not None:
        placeholder = _notes.master_placeholder(package, master, "body")
        if placeholder is not None:
            chain.append(placeholder.find(f"{qn('p:txBody')}/{qn('a:lstStyle')}"))
        root = package.tree(master)
        chain.append(find(root, "p:notesStyle") if root is not None else None)
    return chain


def _notes_block(document: "Document", slide: "Slide") -> list[str]:
    package = document.package
    part = _notes.notes_part(package, slide.part_path)
    if part is None:
        return []
    body_shape = _notes.notes_body(package, part)
    body = body_shape.find(qn("p:txBody")) if body_shape is not None else None
    if body is None:
        return []
    return text_block(paragraphs(document, part, body, notes_chain(document, body_shape)))


__all__ = ["CODE_FONT", "MONOSPACE", "NOTES_MARKER", "TextBlock", "chart_summary", "find_text",
           "outline_blocks", "paragraphs", "placeholder_label", "select", "style_chain",
           "text_block", "to_outline"]
