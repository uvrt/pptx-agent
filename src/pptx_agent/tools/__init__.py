"""PowerPoint tools for an agent: the deck side of :mod:`ooxml_edit.tools`.

An application puts these in a toolbox and a model edits decks only through them -- no
code, no files, no XML::

    from ooxml_edit.tools import Toolbox
    from pptx_agent.tools import FORMAT, GROUPS, TOOLS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=fixed_clock)
    d1 = session.open(deck_bytes, name="q3-review.pptx")
    tools = toolbox.definitions("anthropic", groups="core")
    result = toolbox.dispatch(session, "describe", {"doc": "d1"})

What is here (the tool layer's roadmap numbers them):

* the shared tools' deck handlers (:mod:`.shared_tools`) and the session tools every format
  shares: ``open_document``, ``new_document``, ``save_document``, ``list_documents``,
  ``close_document``, ``undo``, ``find_text``, ``replace_text``, ``render``, ``check``,
  ``set_properties`` and ``batch``;
* the shared ``describe``'s deck handler and ``ppt_read_slides`` (:mod:`.read`);
* text: ``ppt_set_text`` (speaker notes too) and ``ppt_format_text`` (table cells too);
* shapes: ``ppt_add_shape``, ``ppt_set_shape``, ``ppt_add_picture``, ``ppt_add_connector``,
  ``ppt_arrange``; ``ppt_add_shape`` with ``measure: true`` measures text before building;
  tables: ``ppt_add_table``, ``ppt_edit_table`` (fill and borders too);
* slides: ``ppt_add_slide``, ``ppt_draft_slides``, ``ppt_manage_slides``, ``ppt_set_theme``;
* charts and SmartArt, through the shared ``edit_chart`` (``action: "read"`` reads one) and
  ``edit_smartart`` (:mod:`.charts`);
* review: ``ppt_comments``, PowerPoint's comment threads on slides and shapes (:mod:`.review`);
* design facts: ``ppt_design_facts``, measured facts of one slide (:mod:`.facts`).

Positions are the model's own, in points: the layout tools (``ppt_layout``, ``ppt_scale``,
``ppt_copy``, ``ppt_align``, ``place``) and the experimental ``ppt_draw`` were removed after
trial 3 -- no model called them in about forty runs.  Their library calls stay
(:mod:`pptx_agent.edit.layout`, :mod:`~pptx_agent.edit.arrange`,
:mod:`~pptx_agent.edit.scales`, :meth:`Slide.copy_shapes`, :mod:`~pptx_agent.edit.svgprofile`).

Lengths are points at this boundary (12,700 EMU each).  Every changing call returns
``checks``: text that does not fit, collisions and shapes off the slide on the slides it
touched, new validation problems, and layout facts about the shapes the call touched
(near-alignment, uneven gaps, outlier text sizes, labels far from their anchors), each with
the exact change that would resolve it -- facts, not verdicts.

Loading.  ``toolbox.definitions("anthropic")`` sends every tool, the non-core ones deferred
behind tool search; ``groups=["ppt_graphics", "ppt_objects"]`` loads the named groups
instead, for a model without tool search.  ``toolbox.system_prompt(extra=...)`` gives the
shared fragment, :data:`PROMPT`, then the application's own guidance.

House rules belong to the application, written against the facts the tools report.  An
example (documentation only; nothing like it ships)::

    HOUSE_RULES = '''House rules (Acme):
    - Before saving a slide with a graphic, call ppt_design_facts on it.
    - If a set of like shapes uses more than two accent hues and has no legend, recolour
      it to tints of one accent, or add a legend if the colours carry meaning.
    - Leave no empty region larger than a quarter of the content area.
    - Body text is at least 12 pt.'''
    system = toolbox.system_prompt(extra=HOUSE_RULES)
"""

from __future__ import annotations

import io
import zipfile

from ooxml_edit.tools import CORE, DocumentFormat, ToolGroup, shared

from ..edit.chart import ChartDataError, ChartDataWarning
from ..edit.deck import TemplateOpenedWarning
from ..edit.labels import LabelError
from ..edit.text import MarkdownEscapeWarning
from ..edit.units import UnitWarning
from ..fullstate import FullStateError
from ..outline import OutlineWarning
from . import charts, facts, read, review, shapes, shared_tools, slides, tables, text
from .common import KIND, checks, pt

PROMPT = """\
PowerPoint decks
- describe first: slides with ids and content areas, theme colours and fonts, layouts \
and their placeholders. Then ppt_read_slides for the slides the task touches.
- Addresses: s:256 a slide (its id; it never changes; slide numbers 1, 2... are \
positions). 256.5 a shape (256.5#2 when the deck numbered two shapes alike), 256.5/p1 a \
paragraph, 256.5/p1/r0 a run, 256.7/cell1,2 a table cell (row, column from 0), 256/notes \
the speaker notes. The text tools take all of these.
- Positions and sizes are points from the slide's top-left corner.
- Text: text keeps the formatting of what it replaces; paragraphs (the text spec) states \
runs, bullets and frame. To size a box, measure first: ppt_add_shape with measure true \
and the spec and preset (or like) the shape will get adds nothing and returns box_height; \
or set fit_height when adding.
- checks name shapes: overflows, collisions, off_slide. within_allowance is not an \
overflow. A line behind an opaque shape is not a collision: send a marker line behind \
bars rather than splitting it. checks.layout: shapes you touched that almost line up, \
space or size like their neighbours but not quite, each with the exact fix; apply it \
unless the difference is intended.
- Before saving: check (fit, collisions, validate) and resolve the overflows and \
collisions the task does not intend; save_document lists any left. For slides with \
graphics, ppt_design_facts reports what the application's guidance may ask about.
"""

#: Claude's strict mode fits only some tools (20 tools, 24 optional parameters and, measured,
#: about 32 free-text strings per request): the writing tools the reference transcripts
#: (p1-p8, o1, m1) call most, while the optional parameters last.  ppt_set_text alone would
#: take 20 of the 24, so it is left to the dispatcher's validation, as are the option-heavy
#: shape tools.  edit_chart joined them in T4: with its add action it has 14 optional
#: parameters, more than half the request's.  ppt_align left the list with the layout tools
#: (post-T4).
STRICT_FIRST = ("ppt_add_slide", "ppt_manage_slides", "ppt_draft_slides")

_PRESENTATION = b"ppt/presentation.xml"


def detect(data: bytes, name: str) -> bool:
    """A PowerPoint package: a ZIP holding ``ppt/presentation.xml``."""
    if not data.startswith(b"PK"):
        return False
    if name.lower().endswith((".pptx", ".potx", ".pptm", ".potm")):
        return True
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return "ppt/presentation.xml" in archive.namelist()
    except zipfile.BadZipFile:
        return False


def _open(data: bytes):
    from ..edit.document import Document

    return Document.open(data)


def _summary(deck) -> dict:
    width, height = deck.slide_size
    return {"slides": len(deck.slides), "title": deck.title,
            "size": {"w": pt(width), "h": pt(height)},
            "next": "call describe"}


FORMAT = DocumentFormat(
    kind=KIND, open=_open, detect=detect, problems=lambda deck: deck.validate(),
    warnings=(TemplateOpenedWarning, MarkdownEscapeWarning, ChartDataWarning, OutlineWarning,
              UnitWarning),
    prompt=PROMPT,
    errors={LabelError: ("label_not_found", lambda exc: exc.candidates[:50]),
            ChartDataError: "refused", FullStateError: "refused"},
    summary=_summary, checks=checks, strict_first=STRICT_FIRST)

GROUPS = [*shared.GROUPS,
          ToolGroup(text.TEXT_GROUP, "Formatting text: runs, paragraphs, frames, table cells."),
          ToolGroup(shapes.GRAPHICS, "Shapes (and measuring text for them), connectors, z-order, "
                    "groups, duplicates; a slide's design facts."),
          ToolGroup(shapes.OBJECTS, "Pictures and tables, their fill and borders."),
          ToolGroup(review.REVIEW, "Review comments on slides and shapes: threads, replies, "
                    "resolving."),
          ToolGroup(slides.SLIDES, "Adding, drafting, moving and deleting slides; the theme.")]

#: Every PowerPoint tool, the shared ones included, in the order they are offered.
TOOLS = [
    *[tool for tool in shared.SESSION_TOOLS if tool.name != "batch"],
    *shared_tools.TOOLS,
    *read.TOOLS,
    text.ppt_set_text, shapes.ppt_set_shape,
    next(tool for tool in shared.SESSION_TOOLS if tool.name == "batch"),
    text.ppt_format_text,
    shapes.ppt_add_shape, shapes.ppt_add_connector, shapes.ppt_arrange,
    shapes.ppt_add_picture, *tables.TOOLS,
    *slides.TOOLS,
    *charts.TOOLS,
    *facts.TOOLS,
    *review.TOOLS,
]

__all__ = ["FORMAT", "GROUPS", "KIND", "PROMPT", "STRICT_FIRST", "TOOLS", "detect"]
