"""``ppt_design_facts``: measurable design facts of a slide, with addresses, no verdicts;
``ppt_compare_slides``: a run of like slides compared shape by shape (round 7, L3).

What the facts mean -- a rainbow, dead space, a missing legend -- is for the application's
own guidance or critique pass to say; the tool only measures (LP24).
"""

from __future__ import annotations

import json

from ooxml_edit.tools import Result, ToolError, array, number, obj, page_list, string, tool

from ..edit.design import KINDS
from .common import KIND, box_emu, emu, library_errors, numbers, slide as slide_at
from .shapes import GRAPHICS
from .slides import SLIDES


@tool("ppt_design_facts",
      "Measured design facts of a slide, no verdicts: palette and theme roles, sets of like "
      "shapes with their colours and legend, empty regions, alignment near-misses, shape "
      "vocabulary, text sizes, lines over text.",
      {"doc": string("Document id."),
       "slide": string("Slide, s:256 or $ref."),
       "region": obj({"x": number(minimum=-10000, maximum=10000),
                      "y": number(minimum=-10000, maximum=10000),
                      "w": number(minimum=0, maximum=10000),
                      "h": number(minimum=0, maximum=10000)},
                     "Only shapes centred here. Default the slide (content area for empty "
                     "regions).", optional=True),
       "within": number("Near-miss tolerance. Default 2.", minimum=0, maximum=100,
                        optional=True),
       "include": array(string(enum=list(KINDS)), "Facts to report. Default all.",
                        optional=True)},
      kind=KIND, group=GRAPHICS, refs=("slide",))
@library_errors
def ppt_design_facts(call, doc, slide, region=None, within=None, include=None):
    target = slide_at(call.document, slide)
    facts = target.design_facts(region=box_emu(region) if region else None,
                                within=emu(2 if within is None else within),
                                include=include or None)
    data = facts.to_json()
    groups = len(facts.color_groups or [])
    return Result(summary=f"Design facts of s:{target.slide_id}"
                  + (f": {groups} set(s) of like shapes" if facts.color_groups is not None
                     else ""), data=data)


#: Shapes (groups of corresponding shapes that differ) in one page of a comparison.
COMPARE_PAGE = 40


@tool("ppt_compare_slides",
      "Compare a run of like slides shape by shape: where box, text size, colour, font, "
      "fill or outline differ from the majority; near-misses with a fix. Measures only.",
      {"doc": string("Document id."),
       "slides": array(string(or_integer=1), "Slide numbers from 1, or ids (s:256, $ref); "
                       "at least two."),
       "cursor": string("next_cursor of the previous page.", optional=True)},
      kind=KIND, group=SLIDES, refs=("slides[]",), review_core=True)
@library_errors
def ppt_compare_slides(call, doc, slides, cursor=None):
    """The library's :meth:`Document.compare_slides <pptx_agent.Document.compare_slides>`
    with its 2 pt tolerance (the library takes another), paged by shape: ``only`` (shapes
    on half the slides or fewer) comes with the first page."""
    from ..edit.compare import compare_slides

    deck = call.document
    chosen = numbers(deck, slides)
    if len(chosen) < 2:                          # numbers() lists a slide once
        raise ToolError("invalid_arguments", "compare at least two slides: list the run of "
                        "like slides, e.g. [\"s:257\", \"s:258\", \"s:259\"]", field="slides")
    targets = [deck.slides[n - 1] for n in chosen]
    data = compare_slides(targets).to_json()
    few = len(data["only"])
    shapes, total, next_cursor = page_list(data["shapes"], cursor=cursor, limit=COMPARE_PAGE)
    if len(json.dumps(shapes, ensure_ascii=False)) > call.limits.max_result_chars - 2000:
        shapes, total, next_cursor = page_list(data["shapes"], cursor=cursor,
                                               limit=COMPARE_PAGE // 4)
    data["shapes"] = shapes
    if cursor:
        data.pop("only")
    near = sum(1 for entry in shapes for item in entry.get("differ", ()) if "near" in item)
    return Result(summary=f"Compared {len(targets)} slides: {data['matched']} shape(s) "
                  f"matched, {total} differ or are missing somewhere ({near} near-miss "
                  f"propert{'y' if near == 1 else 'ies'} on this page), "
                  f"{few} on few slides",
                  data=data, total=total, next_cursor=next_cursor)


TOOLS = [ppt_design_facts, ppt_compare_slides]
