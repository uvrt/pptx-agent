"""Reading a deck: the shared ``describe``'s deck handler (the deck at a glance) and
``ppt_read_slides``."""

from __future__ import annotations

import json
from typing import Any

from ooxml_edit.tools import (CORE, Result, ToolError, array, integer, page_list, page_text,
                              shared, string, tool)

from .common import (KIND, box_pt, color_json, color_text, library_errors, numbers, pt)


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
    for layout in deck.layouts:
        layouts.append({"name": layout.name, "placeholders": [
            {"type": p.type or "content", "idx": p.idx, "box": box_pt(p.bounds)}
            for p in layout.placeholders]})
    data = {
        "doc": doc, "size": {"w": pt(width), "h": pt(height)}, "slides": slides,
        "theme": {"colors": dict(theme.colors), "fonts": {"major": theme.fonts.major,
                                                         "minor": theme.fonts.minor},
                  "roles": {"primary": roles.primary, "secondary": roles.secondary,
                            "neutral": roles.neutral, "highlight": roles.highlight},
                  "tints": ramps},
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


def _shape_json(shape, z: int) -> dict[str, Any]:
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


@tool("ppt_read_slides",
      "Read slides with shape ids. outline: Markdown, each block under its shape id "
      "(escaped; never paste it back as text). geometry: each shape's kind, box, z-order, "
      "fill, line, text and font sizes. svg: a compact SVG per slide, data-id per shape.",
      {"doc": string("Document id."),
       "slides": array(integer(minimum=1), "Slide numbers from 1. Default all.",
                       optional=True),
       "detail": string("What to show.", enum=["outline", "geometry", "svg"]),
       "cursor": string("next_cursor of the previous page.", optional=True)},
      kind=KIND, group=CORE)
@library_errors
def ppt_read_slides(call, doc, detail, slides=None, cursor=None):
    deck = call.document
    chosen = numbers(deck, slides)
    if detail == "outline":
        text = deck.to_outline(slides=chosen)
        page, next_cursor = page_text(text, cursor=cursor, limit=call.limits.max_result_chars - 2000)
        return Result(summary=f"Outline of slide(s) {_span(chosen)}", data=page,
                      next_cursor=next_cursor)
    if detail == "svg":
        return _svg(call, deck, chosen, cursor)
    shapes = []
    for number in chosen:
        target = deck.slides[number - 1]
        for z, shape in enumerate(target.shapes):
            shapes.append({"slide": f"s:{target.slide_id}", **_shape_json(shape, z)})
    page, total, next_cursor = page_list(shapes, cursor=cursor, limit=60)
    if len(json.dumps(page)) > call.limits.max_result_chars - 2000:
        page, total, next_cursor = page_list(shapes, cursor=cursor, limit=25)
    return Result(summary=f"Geometry of slide(s) {_span(chosen)}: {total} shape(s)", data=page,
                  total=total, next_cursor=next_cursor)


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
                            valid_options=["outline", "geometry"]) from None
        if pages and used + len(svg) > budget:
            break
        pages.append({"slide": f"s:{target.slide_id}", "n": number, "chars": len(svg),
                      "svg": svg})
        used += len(svg)
    end = start + len(pages)
    next_cursor = f"c{end}" if end < len(chosen) else None
    return Result(summary=f"SVG view of slide(s) {_span(chosen[start:end])}", data=pages,
                  total=len(chosen), next_cursor=next_cursor)


def _span(chosen: list[int]) -> str:
    if len(chosen) > 3 and chosen == list(range(chosen[0], chosen[-1] + 1)):
        return f"{chosen[0]}-{chosen[-1]}"
    return ", ".join(map(str, chosen))


TOOLS = [describe, ppt_read_slides]
