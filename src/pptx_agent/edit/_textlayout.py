"""The one module that reaches into ooxml-common's text layout.

:func:`~pptx_agent.edit.fit.measure` (``text_fit``), :func:`~pptx_agent.edit.fit.overflows`
and :func:`~pptx_agent.edit.fit.measure_text` lay text out with the measurement pptx2svg
draws with -- advance widths, kerning, line breaking, line heights and paragraph spacing --
through ooxml-common's public :mod:`~ooxml_common.drawingml.textmeasure`, ``as_drawn``:
each paragraph wrapped inside its ``marL`` and broken at its line breaks, as the renderer
and PowerPoint lay it out.  Kerned as PowerPoint kerns: a static face's legacy ``kern``
table, not its ``GPOS`` pairs (``ooxml_common.drawingml.rules.POWERPOINT.kerning``).

Everything here is public ooxml-common: where a shape's text area sits in its frame is
:class:`~ooxml_common.drawingml.textmeasure.TextBodyMeasure`'s ``area_left`` and
``area_top`` (ooxml-common 0.5), which replaced the last private call.
"""

from __future__ import annotations

# -- the switch: everything pptx-agent uses from ooxml-common's text layout ------------------
from ooxml_common.drawingml import scene as model
from ooxml_common.drawingml.context import RenderContext
from ooxml_common.drawingml.rules import POWERPOINT
from ooxml_common.drawingml.textmeasure import (  # noqa: F401  (re-exported)
    MeasuredLine,
    MeasuredParagraph,
    TextBodyMeasure,
    measure_shape_text,
    measure_text_body,
)
from ooxml_common.text.measure import DefaultTextMeasurer
from ooxml_common.units import PX_PER_PT, emu_to_px, px_to_emu
# ---------------------------------------------------------------------------------------------

#: PowerPoint's default text-frame insets, EMU: left, top, right, bottom.
DEFAULT_INSETS = (91440, 45720, 91440, 45720)


def context(embedded_metrics=None) -> RenderContext:
    """A measuring context, PowerPoint's rules and kerning; ``embedded_metrics`` are a
    deck's own embedded faces."""
    return RenderContext(measurer=DefaultTextMeasurer(embedded_metrics, kerning=POWERPOINT.kerning),
                         rules=POWERPOINT)


__all__ = ["DEFAULT_INSETS", "MeasuredLine", "MeasuredParagraph", "PX_PER_PT", "TextBodyMeasure",
           "context", "emu_to_px", "measure_shape_text", "measure_text_body",
           "model", "px_to_emu"]
