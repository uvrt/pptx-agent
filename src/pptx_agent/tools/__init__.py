"""PowerPoint tools for an agent: the deck side of :mod:`ooxml_edit.tools`.

An application puts these in a toolbox and a model edits decks only through them -- no
code, no files, no XML::

    from ooxml_edit.tools import Toolbox
    from pptx_agent.tools import FORMAT, GROUPS, TOOLS

    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=fixed_clock)
    d1 = session.open(deck_bytes, name="q3-review.pptx")
    tools = toolbox.definitions("anthropic", groups="core")
    openai_tools = toolbox.definitions("openai-responses")   # or "openai-chat"
    result = toolbox.dispatch(session, "describe", {"doc": "d1"})

How to run them with each provider (Claude, OpenAI Responses, Chat Completions):
``README.md`` in this package.

What is here (the tool layer's roadmap numbers them):

* the shared tools' deck handlers (:mod:`.shared_tools`) and the session tools every format
  shares: ``open_document``, ``new_document``, ``save_document``, ``close_document``,
  ``read_blob``, ``undo``, ``find_text``, ``replace_text``, ``render``, ``check``,
  ``set_properties`` and ``batch``;
* the shared ``describe``'s deck handler and ``ppt_read_slides`` (:mod:`.read`);
* text: ``ppt_set_text`` (speaker notes too) and ``ppt_format_text`` (table cells too);
* shapes: ``ppt_add_shape``, ``ppt_set_shape`` (``fill`` and ``line`` on an SVG picture
  recolour it, as PowerPoint's Graphics Fill and Outline), ``ppt_add_picture`` (SVG pictures too; an application's
  asset catalogue adds ``asset`` and ``ppt_find_asset``, :func:`deck_format`),
  ``ppt_add_connector``, ``ppt_arrange`` (align and distribute too); ``ppt_add_shape``
  with ``measure: true`` measures text before building; tables: ``ppt_add_table`` (with
  ``measure: true``, row heights and pages before placing), ``ppt_edit_table`` (fill and
  borders too);
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
Align and distribute came back in round 3 as ``ppt_arrange`` actions, for house rules that
require equal spacing: one call that guarantees it, which the application's rules name.

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

import dataclasses
import io
import zipfile
from collections.abc import Mapping

from ooxml_edit.opc import rels_path_for
from ooxml_edit.tools import CORE, DocumentFormat, ToolGroup, shared

from ..edit.chart import ChartDataError, ChartDataWarning
from ..edit.deck import TemplateOpenedWarning
from ..edit.labels import LabelError
from ..edit.provenance import ProvenanceWarning
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
the speaker notes. L2.5, M0.7: a layout's or master's shape; an edit there shows on every \
slide built on it (checks.affects). The text tools take all of these.
- Positions and sizes are points from the slide's top-left corner.
- Text: text keeps the formatting of what it replaces; paragraphs (the text spec) states \
runs, bullets and frame. To size a box, measure first: ppt_add_shape with measure true \
and the spec and preset (or like) the shape will get adds nothing and returns box_height; \
or set fit_height when adding.
- A figure from a dataset or table goes into text as {{d:<id>/<key>/<column>}} or \
{{t:<table>/<key>/<column>}}, :0.0 after it to round: the tool writes the value in the \
locale and keeps its source.
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
#: (post-T4).  ppt_arrange, strict until round 3, has eight optional parameters with align
#: and distribute and no longer fits beside the others; ppt_add_table takes its place.
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


def _configure(deck, session) -> None:
    """A deck joining a session renders and measures with the session's font folders
    (``Toolbox(font_dirs=...)``; ``None``: ``OOXML_FONT_DIRS``) and is drawn and edited in
    its locale (``Toolbox(locale=...)``; en-US unless given)."""
    deck.font_dirs = session.font_dirs
    deck.locale = session.locale


def _summary(deck) -> dict:
    width, height = deck.slide_size
    return {"slides": len(deck.slides), "title": deck.title,
            "size": {"w": pt(width), "h": pt(height)},
            "next": "call describe"}


#: Relationships a slide's undo scope does not follow: to what other slides share (layouts,
#: masters, themes), to other slides, and back up to the presentation.
_SHARED_RELATIONSHIPS = ("/slideLayout", "/slideMaster", "/slide", "/notesMaster", "/theme",
                         "/handoutMaster", "/presentation", "/commentAuthors")


def _slide_scope(deck, scope: str) -> list[str]:
    """``undo``'s ``scope`` on a deck: the slide (``256``, ``s:256``) and what it owns -- its
    relationships, notes, charts and their workbooks, diagrams, comments, media."""
    from .common import slide, slide_id_of

    number = slide_id_of(str(scope))
    target = slide(deck, str(number) if number is not None else str(scope), field="scope")
    own = deck.package.reachable_parts(
        target.part_path, lambda rel: not rel.type.endswith(_SHARED_RELATIONSHIPS))
    # Its layout and master too, so the latest change to what the slide shows is found --
    # and refused as entangled when it was a layout's or master's (_template_parts).
    layout, master = target._template_parts()
    return own + [part for path in (layout, master) if path
                  for part in (path, rels_path_for(path))]


def _template_parts(deck) -> list[str]:
    """``undo``'s shared parts on a deck: every layout and master (and their
    relationships), which every slide on them shows -- a change to one is no one slide's
    to take back (``DocumentFormat.undo_shared``)."""
    return [part for template in [*deck.masters, *deck.layouts]
            for part in (template.part_path, rels_path_for(template.part_path))]


FORMAT = DocumentFormat(
    kind=KIND, open=_open, detect=detect, problems=lambda deck: deck.validate(),
    warnings=(TemplateOpenedWarning, MarkdownEscapeWarning, ChartDataWarning, OutlineWarning,
              UnitWarning, ProvenanceWarning),
    prompt=PROMPT,
    errors={LabelError: ("label_not_found", lambda exc: exc.candidates[:50]),
            ChartDataError: "refused", FullStateError: "refused"},
    summary=_summary, checks=checks, strict_first=STRICT_FIRST,
    undo_scope=_slide_scope, undo_shared=_template_parts,
    restored=lambda deck: deck._reset_caches(),
    configure=_configure,
    store_provenance=tables.store_provenance, read_provenance=tables.read_provenance)

def deck_format(*, min_text_size: "float | dict[str, float] | None" = None,
                assets=None) -> DocumentFormat:
    """:data:`FORMAT` with the application's own settings -- never tool arguments, so the
    definitions the model is sent stay the same unless a setting adds a tool:

    ``min_text_size`` (points): ``check``, the checks after every change and
    ``save_document`` report ``small_text``, the text drawn smaller, with its effective
    size and where the size comes from -- a placeholder that inherits 9 pt from its layout
    is reported as ``"from": "layout 'Title and Content' body placeholder"``
    (:meth:`Slide.text_below <pptx_agent.Slide.text_below>`).  One number is the minimum
    for all text; a mapping sets it per kind -- ``body`` (shapes, text boxes and
    placeholders), ``table`` (table cells), ``chart`` (a chart's title, axis titles, axis
    labels, data labels and legend) -- and a kind it leaves out is not checked::

        toolbox = Toolbox(TOOLS, formats=[deck_format(min_text_size=12)], groups=GROUPS)
        deck_format(min_text_size={"body": 12, "table": 10, "chart": 10})

    A shape's entry is one per size and source, ``{"shape", "size", "from"}``; a table is
    one entry, ``{"table": "257.4", "cells": 30, "size": 9.0, "min": 10, "from": ...}`` --
    how many cells are below, the smallest size and where it comes from; a chart is one
    entry, ``{"chart": "257.5", "min": 10, "parts": [{"part": "axis_labels", "size": 8.0,
    "from": ...}]}`` -- each part's smallest, named as ``edit_chart``'s
    ``set_text_style`` takes it.

    ``assets``: the application's asset catalogue -- brand icons, logos -- any object with
    ``search(query, limit)`` and ``get(id) -> (bytes, MIME type)``
    (:class:`ooxml_edit.tools.AssetCatalogue`), such as
    :class:`ooxml_edit.tools.FolderCatalogue` over a folder and a JSON or CSV index.  With
    one the toolbox offers ``ppt_find_asset`` (deferred, in ``ppt_objects``: ids, names and
    keywords for words) and ``ppt_add_picture`` takes ``asset``, an id, in place of
    ``image`` -- on ``target`` too, to replace a picture's image.  Without one neither is
    sent, so a toolbox without a catalogue pays nothing for it::

        from ooxml_edit.tools import FolderCatalogue

        icons = FolderCatalogue("brand-icons")               # brand-icons/index.json
        toolbox = Toolbox(TOOLS, formats=[deck_format(assets=icons)], groups=GROUPS)
    """
    from ..edit.effective import text_minimums

    text_minimums(min_text_size)                 # refused here, not at the first check
    if isinstance(min_text_size, Mapping):
        min_text_size = {kind: float(points) for kind, points in min_text_size.items()
                         if points is not None}
    elif min_text_size is not None:
        min_text_size = float(min_text_size)
    if assets is not None and not (callable(getattr(assets, "search", None))
                                   and callable(getattr(assets, "get", None))):
        raise ValueError("assets is a catalogue: an object with search(query, limit) and "
                         "get(id), such as ooxml_edit.tools.FolderCatalogue")
    tools = (() if assets is None
             else (shapes.picture_tool(assets), shapes.find_asset_tool(assets)))

    def configure(deck, session) -> None:
        _configure(deck, session)
        deck.min_text_size = (dict(min_text_size) if isinstance(min_text_size, dict)
                              else min_text_size)

    return dataclasses.replace(FORMAT, configure=configure, tools=tools)


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

__all__ = ["FORMAT", "GROUPS", "KIND", "PROMPT", "STRICT_FIRST", "TOOLS", "deck_format",
           "detect"]
