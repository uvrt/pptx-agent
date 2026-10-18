"""Reading a deck: the shared ``describe``'s deck handler (the deck at a glance) and
``ppt_read_slides``."""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import (CORE, Result, ToolError, array, integer, page_list, page_text,
                              shared, string, tool)

from ..edit.units import EMU_PER_POINT
from .common import (KIND, box_pt, color_json, color_text, library_errors, place, pt, targets)


@shared.handler("describe", kind=KIND)
@library_errors
def describe(call, doc):
    """The shared ``describe``'s deck handler: the deck at a glance."""
    deck = call.document
    width, height = deck.slide_size
    slides = []
    for number, slide in enumerate(deck.slides, 1):
        layout = slide.layout
        slides.append({"id": f"s:{slide.slide_id}", "n": number, "title": slide.title,
                       "layout": layout.name if layout is not None else None,
                       "shapes": len([s for s in slide.shapes if s.parent_group is None]),
                       "content_area": box_pt(slide.content_area),
                       "notes": bool(slide.has_notes)})
    theme = deck.theme
    roles = theme.roles
    ramps = {slot: [t.color for t in tints] for slot, tints in theme.ramps.items()}
    layouts = []
    several = len(deck.masters) > 1
    for layout in deck.layouts:
        shapes = layout.shapes
        placeholders = []
        for p, shape in zip(layout.placeholders,
                            [s for s in shapes if s.placeholder is not None
                             and s.parent_group is None]):
            entry = {"id": shape.id, "type": p.type or "content", "idx": p.idx,
                     "box": box_pt(p.bounds)}
            if shape.prompt is not None:
                entry["prompt"] = shape.prompt
            placeholders.append(entry)
        item = {"id": layout.id, "name": layout.name, "placeholders": placeholders}
        if several:
            item["master"] = layout.master.id
        text = _own_text(shapes)
        if text:
            item["text"] = text
        layouts.append(item)
    masters = []
    for master in deck.masters:
        item = {"id": master.id, "name": master.name}
        text = _own_text(master.shapes)
        if text:
            item["text"] = text
        footer = {}
        for kind, label in (("ftr", "footer"), ("dt", "date")):
            shape = next((s for s in master.shapes if s.placeholder is not None
                          and s.placeholder.type == kind), None)
            if shape is not None and shape.text:
                footer[label] = {"id": shape.id, "text": shape.text}
        if footer:
            item["footer"] = footer
        masters.append(item)
    data = {
        "doc": doc, "size": {"w": pt(width), "h": pt(height)}, "slides": slides,
        "theme": {"colors": dict(theme.colors), "fonts": {"major": theme.fonts.major,
                                                         "minor": theme.fonts.minor},
                  "roles": {"primary": roles.primary, "secondary": roles.secondary,
                            "neutral": roles.neutral, "highlight": roles.highlight},
                  "tints": ramps, "table_styles": _table_styles(deck)},
        "masters": masters,
        "layouts": layouts,
        "properties": {"title": deck.title, "author": deck.author, "language": deck.language},
        "baseline_problems": [str(p) for p in call.entry.baseline_problems],
    }
    threads = deck.comments()
    if threads:
        data["comments"] = {"threads": len(threads),
                            "open": sum(1 for thread in threads if not thread.done),
                            "read": "ppt_comments list"}
    return Result(summary=f"{len(slides)} slide(s), {len(layouts)} layout(s)", data=data)


def _own_text(shapes) -> list[dict[str, str]]:
    """A layout's or master's text that is not a placeholder -- a legal line, a
    classification mark -- which every slide built on it shows."""
    return [{"id": shape.id, "text": shape.text} for shape in shapes
            if shape.placeholder is None and shape.kind == "shape" and shape.text.strip()]


def _table_styles(deck) -> dict[str, Any]:
    """The deck's table styles in brief: its default, its own (a template's), and how the
    built-in names go."""
    from ..edit.tablestyles import BUILTIN_NAMES

    styles = deck.table_styles
    default = next((s["name"] or s["id"] for s in styles if s.get("default")), None)
    return {"default": default, "own": [s["name"] or s["id"] for s in styles
                                        if not s["builtin"]],
            "builtin": BUILTIN_NAMES}


def _run_colors(paragraphs) -> list[str]:
    """The explicit colours of a text's runs, each once, in order."""
    seen: list[str] = []
    for paragraph in paragraphs:
        for run in paragraph.runs:
            color = run.color
            if color is not None and run.text:
                text = color_text(color)
                if text not in seen:
                    seen.append(text)
    return seen


def _shape_json(shape, z: int, *, idx: bool = False) -> dict[str, Any]:
    """A shape's geometry entry; with ``idx`` (a read with a layout or master in it) a
    placeholder's ``idx`` too, the number a slide's placeholder inherits by."""
    data: dict[str, Any] = {"id": shape.id, "kind": shape.kind, "z": z}
    if shape.name:
        data["name"] = shape.name
    data["box"] = box_pt(shape.bounds)
    drawn = shape.drawn_bounds
    if drawn is not None and shape.bounds is not None and tuple(drawn) != tuple(shape.bounds):
        data["drawn"] = box_pt(drawn)
    if shape.rotation:
        data["rotation"] = shape.rotation
    parent = shape.parent_group
    if parent is not None:
        data["group"] = parent.id
    if shape.placeholder is not None:
        data["placeholder"] = shape.placeholder.type or "content"
        if idx and shape.placeholder.idx is not None:
            data["idx"] = shape.placeholder.idx
    if shape.kind == "shape":
        data["preset"] = shape.preset
        fill = shape.effective_fill
        if fill is not None:
            data["fill"] = (color_json(fill.color, shape) if fill.kind == "solid"
                            else {"kind": fill.kind})
        frame = shape.text_frame
        text = frame.text if frame is not None else ""
        if text:
            data["text"] = text
            sizes = sorted({run.effective_size for paragraph in frame.paragraphs
                            for run in paragraph.runs if run.text})
            if sizes:
                data["sizes"] = sizes
            colors = _run_colors(frame.paragraphs)
            if colors:
                data["text_colors"] = colors
            data["frame"] = {"insets": [pt(v) for v in frame.insets], "anchor": frame.anchor,
                             "autofit": frame.autofit, "wrap": frame.wrap}
    if shape.kind in ("shape", "connector", "picture"):
        line = shape.line
        if line is not None and line.exists:
            outline: dict[str, Any] = {}
            if line.color is not None:
                outline.update(color_json(line.color, shape) or {})
            if line.width is not None:
                outline["width"] = pt(line.width)
            if line.dash:
                outline["dash"] = line.dash
            if line.visible is False:
                outline = {"visible": False}
            if outline:
                data["line"] = outline
    if shape.kind == "connector":
        for end, connection in (("from", shape.begin_connection), ("to", shape.end_connection)):
            if connection is not None:
                data[end] = {"shape": connection[0].id, "site": connection[1]}
    elif shape.kind == "group":
        data["children"] = [child.id for child in shape.children]
    elif shape.kind == "picture":
        size = shape.image_size
        if size is not None:
            data["image"] = {"px": [size.width, size.height]}
        if shape.vector is not None:
            data.setdefault("image", {})["svg"] = True
    elif shape.kind == "graphic_frame":
        if shape.has_table:
            table = shape.table
            data["table"] = {"rows": table.rows, "columns": table.columns}
            fills: dict[str, list[str]] = {}
            texts: dict[str, list[str]] = {}
            for row in range(table.rows):
                for column in range(table.columns):
                    cell = table.cell(row, column)
                    fill = cell.fill
                    if fill is not None and fill.kind == "solid" and fill.color is not None:
                        fills.setdefault(color_text(fill.color), []).append(f"{row},{column}")
                    for color in _run_colors(cell.text_frame.paragraphs):
                        texts.setdefault(color, []).append(f"{row},{column}")
            if fills:
                data["cell_fills"] = fills
            if texts:
                data["cell_text_colors"] = texts
        elif shape.has_chart:
            data["chart"] = True
        elif shape.has_diagram:
            data["smartart"] = True
    return data


#: The columns of ``detail: "compact"``, one row per shape.
COMPACT_COLUMNS = ["id", "kind", "box", "fill", "size", "text", "group"]
#: The compact columns when a layout or master is read: a placeholder's type and idx too
#: (``"title"``, ``"content 1"``, ``"dt 10"``) -- a slide's read with it as well, so its
#: placeholders show which of the layout's they inherit from.
TEMPLATE_COLUMNS = [*COMPACT_COLUMNS, "placeholder"]
#: Words of a text the compact view shows before it cuts.
COMPACT_WORDS = 8


def _words(text: str) -> str:
    """The first words of a text, and how many it has when they are not all."""
    words = text.split()
    if len(words) <= COMPACT_WORDS:
        return " ".join(words)
    return f"{' '.join(words[:COMPACT_WORDS])}… ({len(words)} words)"


def _compact_row(shape) -> list[Any]:
    """One shape in brief, the columns of :data:`COMPACT_COLUMNS`: the kind as geometry
    has it (a graphic frame as ``table``, ``chart`` or ``smartart``), the box in whole
    points, the fill drawn (hex where it resolves; a gradient or picture fill by its
    kind), the text's size (the effective size most of its characters are drawn at,
    inherited and autofit-scaled, as ``geometry``'s ``sizes``) and its first words -- a
    table's size and first cell, a chart's type and series, a group's member count."""
    kind, fill, text, size = shape.kind, None, None, None
    if kind == "shape":
        drawn = shape.effective_fill
        if drawn is not None and drawn.kind == "solid" and drawn.color is not None:
            fill = color_json(drawn.color, shape)
            fill = fill.get("hex") or fill["color"]
        elif drawn is not None and drawn.kind != "none":
            fill = drawn.kind
        frame = shape.text_frame
        text = _words(frame.text) if frame is not None else None
        if text:
            size = _dominant_size(frame)
    elif kind == "group":
        text = f"{len(shape.children)} shapes"
    elif kind == "graphic_frame":
        if shape.has_table:
            kind, table = "table", shape.table
            text = f"{table.rows}×{table.columns}: {_words(table.cell(0, 0).text_frame.text)}"
        elif shape.has_chart:
            kind, chart = "chart", shape.chart
            kind_of = "dot" if chart.is_dot else chart.chart_type
            text = f"{kind_of or 'chart'}, {len(chart.series_names)} series"
        elif shape.has_diagram:
            kind = "smartart"
    bounds = shape.bounds
    box = (None if bounds is None
           else [round(float(v) / EMU_PER_POINT) for v in bounds])
    parent = shape.parent_group
    return [shape.id, kind, box, fill, size, text or None,
            parent.id if parent is not None else None]


def _placeholder_label(shape) -> str | None:
    """A layout's or master's placeholder as the compact view labels it: its type
    (``content`` when it has none) and its idx, ``"content 1"``."""
    placeholder = shape.placeholder
    if placeholder is None:
        return None
    kind = placeholder.type or "content"
    return kind if placeholder.idx is None else f"{kind} {placeholder.idx}"


def _dominant_size(frame) -> float | None:
    """The effective size most of a text's characters are drawn at; the first of those
    tied."""
    weights: dict[float, int] = {}
    for paragraph in frame.paragraphs:
        for run in paragraph.runs:
            weight = len(run.text.strip())
            if weight:
                size = run.effective_size
                weights[size] = weights.get(size, 0) + weight
    if not weights:
        return None
    best = max(weights.values())
    size = next(size for size, weight in weights.items() if weight == best)
    return int(size) if float(size).is_integer() else size


@tool("ppt_read_slides",
      "Read slides with shape ids. outline: Markdown, each block under its shape id "
      "(escaped; never paste it back as text). geometry: each shape's kind, box, z-order, "
      "fill, line, text and font sizes. compact: id, kind, box, fill, text size, first "
      "words per shape. svg: an SVG per slide, data-id per shape.",
      {"doc": string("Document id."),
       "slides": array(string(or_integer=1), "Slide numbers from 1, or ids (s:256, $ref); "
                       "layouts L2, masters M0. Default all.", optional=True),
       "detail": string("What to show.", enum=["outline", "geometry", "compact", "svg"]),
       "cursor": string("next_cursor of the previous page.", optional=True)},
      kind=KIND, group=CORE, refs=("slides[]",))
@library_errors
def ppt_read_slides(call, doc, detail, slides=None, cursor=None):
    """Slides -- and layouts and masters by their keys (``L2``, ``M0``, round 8, M3), read
    as slides are: their shapes' rows under the addresses the other tools take
    (``L2.3``), a placeholder's type and idx with them, inherited sizes resolved.  Not in
    ``outline`` (Markdown of slides) or ``svg`` (the agent view is of slides);
    ``describe`` lists a layout's placeholders and text, ``render`` draws it."""
    deck = call.document
    chosen = list(range(1, len(deck.slides) + 1)) if slides is None else targets(deck, slides)
    templates = [item for item in chosen if isinstance(item, str)]
    if templates and detail in ("outline", "svg"):
        raise ToolError("refused", f"{detail} reads slides only, not {', '.join(templates)}: "
                        "read a layout or master with detail compact or geometry, describe "
                        "(its placeholders and text) or render (its picture)",
                        field="detail", valid_options=["compact", "geometry"])
    if detail == "outline":
        text = deck.to_outline(slides=chosen)
        page, next_cursor = page_text(text, cursor=cursor, limit=call.limits.max_result_chars - 2000)
        return Result(summary=f"Outline of slide(s) {_span(chosen)}", data=page,
                      next_cursor=next_cursor)
    if detail == "svg":
        return _svg(call, deck, chosen, cursor)
    sheets = [_sheet_of(deck, item) for item in chosen]
    if detail == "compact":
        columns = TEMPLATE_COLUMNS if templates else COMPACT_COLUMNS
        rows = [_compact_row(shape) + ([_placeholder_label(shape)] if templates else [])
                for sheet in sheets for shape in sheet.shapes]
        page, total, next_cursor = page_list(rows, cursor=cursor, limit=300)
        if len(json.dumps(page, ensure_ascii=False)) > call.limits.max_result_chars - 2000:
            page, total, next_cursor = page_list(rows, cursor=cursor, limit=120)
        return Result(summary=f"Shapes of slide(s) {_span(chosen)} in brief: {total}",
                      data={"columns": columns, "shapes": page},
                      total=total, next_cursor=next_cursor)
    shapes = []
    for item, sheet in zip(chosen, sheets):
        where = item if isinstance(item, str) else place(sheet)
        for z, shape in enumerate(sheet.shapes):
            shapes.append({"slide": where, **_shape_json(shape, z, idx=bool(templates))})
    page, total, next_cursor = page_list(shapes, cursor=cursor, limit=60)
    if len(json.dumps(page)) > call.limits.max_result_chars - 2000:
        page, total, next_cursor = page_list(shapes, cursor=cursor, limit=25)
    return Result(summary=f"Geometry of slide(s) {_span(chosen)}: {total} shape(s)", data=page,
                  total=total, next_cursor=next_cursor)


def _sheet_of(deck, item):
    """The slide a number names, or the layout or master a key (``L2``, ``M0``) names."""
    return deck._template_by_key(item) if isinstance(item, str) else deck.slides[item - 1]


def _svg(call, deck, chosen: list[int], cursor: str | None) -> Result:
    """The agent view of each slide (pptx2svg, LR3), paged by slide within the result cap."""
    entry = call.entry
    budget = call.limits.max_result_chars - 2000
    start = int(cursor[1:]) if cursor and cursor[1:].isdigit() else 0
    if start >= len(chosen) and chosen:
        raise ToolError("invalid_arguments", f"cursor {cursor!r} is past the end",
                        field="cursor")
    pages, used = [], 0
    for number in chosen[start:]:
        target = deck.slides[number - 1]
        key = ("agent-svg", target.slide_id)
        try:
            svg = entry.cached(entry.render_cache, key, lambda: target.render_svg(agent=True))
        except RuntimeError as exc:
            raise ToolError("refused", f"the SVG view is not available: {exc}",
                            valid_options=["outline", "geometry", "compact"]) from None
        if pages and used + len(svg) > budget:
            break
        pages.append({"slide": f"s:{target.slide_id}", "n": number, "chars": len(svg),
                      "svg": svg})
        used += len(svg)
    end = start + len(pages)
    next_cursor = f"c{end}" if end < len(chosen) else None
    return Result(summary=f"SVG view of slide(s) {_span(chosen[start:end])}", data=pages,
                  total=len(chosen), next_cursor=next_cursor)


def _span(chosen: "list[int | str]") -> str:
    if len(chosen) > 3 and all(isinstance(item, int) for item in chosen) \
            and chosen == list(range(chosen[0], chosen[-1] + 1)):
        return f"{chosen[0]}-{chosen[-1]}"
    return ", ".join(map(str, chosen))


TOOLS = [describe, ppt_read_slides]
