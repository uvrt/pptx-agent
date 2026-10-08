"""Text tools: ``ppt_set_text`` (speaker notes too, ``256/notes``) and ``ppt_format_text``.

Measuring text before building is ``ppt_add_shape`` with ``measure: true`` (:mod:`.shapes`)."""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (CORE, Result, ToolError, array, boolean, integer, number, obj,
                              string, tool)

from ..edit.text import Paragraph, Run, TextFrame
from .common import (ALIGN, ANCHOR, AUTOFIT, BULLET, KIND, emu, fit_json, insets_param,
                     library_errors, resolve, text_params, touch_address, write_text)

TEXT_GROUP = "ppt_text"
#: Up to this many items, a call returns each target's own fit; beyond it (and in a batch)
#: the checks' overflows, which measure each slide once, say what does not fit.
MAX_ITEM_FITS = 3

#: Presets ppt_add_shape names directly (and measures with); any other is "other".
COMMON_PRESETS = [
    "textbox", "rect", "roundRect", "ellipse", "triangle", "diamond", "parallelogram",
    "pentagon", "hexagon", "chevron", "homePlate", "rightArrow", "leftArrow", "upArrow",
    "downArrow", "flowChartProcess", "flowChartDecision", "flowChartTerminator",
    "flowChartAlternateProcess", "wedgeRoundRectCallout", "snip1Rect", "round2SameRect",
    "can", "star5", "line", "other"]


def preset_of(item: dict[str, Any], *, field: str = "preset") -> str:
    """The preset an item names: one of :data:`COMMON_PRESETS`, or ``other`` with
    ``preset_name``; an unknown name is ``invalid_arguments`` listing the presets."""
    from ..edit.presets import PRESETS

    preset = item.get("preset", "rect")
    if preset == "other":
        name = item.get("preset_name")
        if not name or name not in PRESETS:
            raise ToolError("invalid_arguments",
                            f"preset_name {name!r} is not a preset geometry", field="preset_name",
                            valid_options=sorted(PRESETS))
        return name
    return preset


def _fit_of(target) -> dict[str, Any] | None:
    shape = getattr(target, "kind", None) and target
    if shape is not None and shape.kind == "shape":
        try:
            return fit_json(shape.text_fit())
        except Exception:  # noqa: BLE001 -- a fit that cannot be measured is left out
            return None
    return None


@tool("ppt_set_text",
      "Replace the text of shapes, table cells or speaker notes (256/notes): text keeps "
      "the formatting, paragraphs (the text spec) states it. Returns each target's fit.",
      {"doc": string("Document id."),
       "items": array(obj({"target": string("Shape 256.5, cell 256.7/cell1,2, notes "
                                            "256/notes, or $ref."),
                           **text_params(described=False)}),
                      "Targets and their text.", min_items=1)},
      kind=KIND, group=CORE, mutates=True, refs=("items[].target",))
@library_errors
def ppt_set_text(call, doc, items):
    deck = call.document
    fits = {}
    for index, item in enumerate(items):
        if "text" not in item and "paragraphs" not in item:
            raise ToolError("invalid_arguments", "give text or paragraphs",
                            field=f"items[{index}].text", valid_options=["text", "paragraphs"])
        target = resolve(deck, item["target"], field=f"items[{index}].target")
        write_text(target, item, field=f"items[{index}]")
        touch_address(call, item["target"])
    if len(items) <= MAX_ITEM_FITS and not call.in_batch:
        for item in items:
            fit = _fit_of(deck.resolve(item["target"]))
            if fit is not None:
                fits[item["target"]] = fit
    return Result(summary=f"Set the text of {len(items)} target(s)",
                  changed=[item["target"] for item in items], data={"fits": fits})


_RUN_KEYS = ("bold", "italic", "underline", "strike", "size", "font", "color", "hyperlink")
_PARAGRAPH_KEYS = ("align", "level", "bullet", "space_before", "space_after", "line_spacing")
_FRAME_KEYS = ("autofit", "font_scale", "insets", "anchor", "wrap")


@tool("ppt_format_text",
      "Format existing text: runs (bold, size, colour...), paragraphs (alignment, bullets, "
      "spacing) and the frame (insets, anchor, autofit). A shape formats all its text; "
      "256.5/p1 one paragraph, 256.5/p1/r0 one run.",
      {"doc": string("Document id."),
       "items": array(obj({
           "target": string("Shape, cell, notes, paragraph (/p1), run (/p1/r0) or $ref."),
           "bold": boolean(optional=True),
           "italic": boolean(optional=True),
           "underline": boolean(optional=True),
           "strike": boolean(optional=True),
           "size": number(minimum=1, maximum=400, optional=True),
           "font": string(optional=True),
           "color": string(optional=True),
           "hyperlink": string("URL; empty removes it.", optional=True),
           "align": string(enum=ALIGN, optional=True),
           "level": integer("0-8.", minimum=0, maximum=8, optional=True),
           "bullet": string(enum=BULLET, optional=True),
           "space_before": number(minimum=0, maximum=1000, optional=True),
           "space_after": number(minimum=0, maximum=1000, optional=True),
           "line_spacing": number("Times single.", minimum=0.1, maximum=10, optional=True),
           "autofit": string(enum=AUTOFIT, optional=True),
           "font_scale": number("Stored shrink, 0.25-1, as PowerPoint draws it.",
                                minimum=0.25, maximum=1, optional=True),
           "insets": insets_param(),
           "anchor": string(enum=ANCHOR, optional=True),
           "wrap": boolean(optional=True)}),
           "Targets and their formatting.", min_items=1)},
      kind=KIND, group=TEXT_GROUP, mutates=True, refs=("items[].target",))
@library_errors
def ppt_format_text(call, doc, items):
    from ..edit.units import Pt

    deck = call.document
    for index, item in enumerate(items):
        field = f"items[{index}]"
        target = resolve(deck, item["target"], field=f"{field}.target")
        frame = target if isinstance(target, TextFrame) else getattr(target, "text_frame", None)
        if isinstance(target, Run):
            paragraphs, runs = [], [target]
        elif isinstance(target, Paragraph):
            paragraphs, runs = [target], list(target.runs)
        elif frame is not None:
            paragraphs = list(frame.paragraphs)
            runs = [run for paragraph in paragraphs for run in paragraph.runs]
        else:
            raise ToolError("invalid_arguments", f"{item['target']} holds no text",
                            field=f"{field}.target")
        run_format = {key: item[key] for key in _RUN_KEYS if key in item}
        if "hyperlink" in run_format and run_format["hyperlink"] == "":
            run_format["hyperlink"] = None
        for run in runs:
            if run_format:
                run.format(**run_format)
        for paragraph in paragraphs:
            if "align" in item:
                paragraph.alignment = item["align"]
            if "level" in item:
                paragraph.level = item["level"]
            if "space_before" in item:
                paragraph.space_before = Pt(item["space_before"])
            if "space_after" in item:
                paragraph.space_after = Pt(item["space_after"])
            if "line_spacing" in item:
                paragraph.line_spacing = item["line_spacing"]
            if item.get("bullet") == "bullet":
                paragraph.set_bullet()
            elif item.get("bullet") == "number":
                paragraph.set_numbering()
            elif item.get("bullet") == "none":
                paragraph.set_no_bullet()
        frame_keys = [key for key in _FRAME_KEYS if key in item]
        if frame_keys:
            if isinstance(target, (Run, Paragraph)) or frame is None:
                raise ToolError("invalid_arguments", "frame properties (insets, anchor, "
                                "autofit, font_scale, wrap) go on a shape or cell",
                                field=f"{field}.{frame_keys[0]}")
            if "insets" in item:
                frame.insets = tuple(emu(item["insets"][s])
                                     for s in ("left", "top", "right", "bottom"))
            if "anchor" in item:
                frame.anchor = item["anchor"]
            if "wrap" in item:
                frame.wrap = item["wrap"]
            if "autofit" in item:
                frame.autofit = item["autofit"]
            if "font_scale" in item:
                frame.font_scale = item["font_scale"]
        touch_address(call, item["target"])
    return Result(summary=f"Formatted {len(items)} target(s)",
                  changed=[item["target"] for item in items])


TOOLS = [ppt_set_text, ppt_format_text]
