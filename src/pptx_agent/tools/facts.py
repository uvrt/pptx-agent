"""``ppt_design_facts``: measurable design facts of a slide, with addresses, no verdicts.

What the facts mean -- a rainbow, dead space, a missing legend -- is for the application's
own guidance or critique pass to say; the tool only measures (LP24).
"""

from __future__ import annotations

from ooxml_edit.tools import Result, array, number, obj, string, tool

from ..edit.design import KINDS
from .common import KIND, box_emu, emu, library_errors, slide as slide_at
from .shapes import GRAPHICS


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


TOOLS = [ppt_design_facts]
