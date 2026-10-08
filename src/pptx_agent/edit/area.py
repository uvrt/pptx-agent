"""Where a slide's content goes: :meth:`Slide.content_area <pptx_agent.Slide.content_area>`.

A fact read from the slide's layout, not a design rule: the box the layout itself gives its
body, or -- for a layout without one -- the band between the title and the footer the
layout draws.  First written for the SVG-authoring spike, where agents in both arms
praised it; graphics tools place into it by default.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:  # pragma: no cover
    from .document import Slide

_INCH = 914400


def content_area(slide: "Slide") -> tuple[int, int, int, int]:
    """The slide's content area, EMU ``(left, top, width, height)``.

    * The union of the layout's body and content placeholders, when it has any.
    * Otherwise the title's left and right edges, from 0.15 in below the title (or below a
      rule drawn under it) down to 0.08 in above the highest layout shape in the bottom
      fifth of the slide (a footer band, a wordmark), or 0.4 in above the slide's foot.
    * Without a title, 0.4 in in from each side, from 1.2 in down.
    """
    width, height = slide.document.slide_size
    layout = slide.layout
    placeholders = layout.placeholders if layout is not None else []
    bodies = [p for p in placeholders
              if p.type in (None, "body", "obj") and p.bounds and p.bounds[2] > 0
              and p.bounds[3] > 0]
    if bodies:
        left = min(p.bounds[0] for p in bodies)
        top = min(p.bounds[1] for p in bodies)
        right = max(p.bounds[0] + p.bounds[2] for p in bodies)
        bottom = max(p.bounds[1] + p.bounds[3] for p in bodies)
        return left, top, right - left, bottom - top
    title = next((p for p in placeholders if p.type in ("title", "ctrTitle") and p.bounds),
                 None)
    if title is not None:
        left, right = title.bounds[0], title.bounds[0] + title.bounds[2]
        top = title.bounds[1] + title.bounds[3]
    else:
        left, right, top = int(0.4 * _INCH), width - int(0.4 * _INCH), int(1.2 * _INCH)
    held = {tuple(p.bounds) for p in placeholders if p.bounds}
    bottom = height - int(0.4 * _INCH)
    for shape in (layout.shapes if layout is not None else []):
        bounds = shape.bounds
        if not bounds or tuple(bounds) in held:
            continue
        if bounds[3] == 0 and top - _INCH // 2 <= bounds[1] < height / 2:   # a rule under the title
            top = max(top, bounds[1])
        elif bounds[1] > height * 0.8:
            bottom = min(bottom, bounds[1] - int(0.08 * _INCH))
    top += int(0.15 * _INCH)
    return left, top, right - left, max(0, bottom - top)


__all__ = ["content_area"]
