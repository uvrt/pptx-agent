"""pptx-agent -- an AI-editable PowerPoint layer.

Layers:

* :mod:`pptx_agent.core` is format-neutral: the OPC package, XML helpers with schema-ordered
  insertion, undo history and id stamping -- the part a docx editor would share.
* :mod:`pptx_agent.oxml` adds PresentationML: the ``p:``/``a:`` vocabulary and child orders,
  and where the slides are.  Nodes are live views, never owned copies, so anything this
  library does not model survives a read/edit/write cycle untouched.
* :mod:`pptx_agent.edit` is the semantic API: slides (add, delete, duplicate, reorder),
  shapes, pictures, text frames, paragraphs, runs and hyperlinks, fills and outlines, tables,
  groups, charts (cache and embedded workbook kept in step), SmartArt, stable ids, undo --
  and new shapes built from nothing: autoshapes, text boxes, connectors that stay attached
  to their shapes, and tables, written as PowerPoint writes them -- and new decks, from
  nothing (PowerPoint's Office theme, master and eleven layouts, at any size) or from a
  template, with their metadata.
* Rendering is delegated to `pptx2svg <https://github.com/uvrt/pptx2svg>`_, which turns the
  current state of the document into SVG or PNG for an agent to look at.
* :mod:`pptx_agent.fullstate` makes that SVG carry the slide itself: ``data-ooxml-*``
  attributes read from the unresolved document, and each shape's OOXML, so an SVG edited
  elsewhere can be applied back with :meth:`Document.apply_svg`.

docs/USAGE.md in the repository has a tested example for each common task.  In short::

    from pptx_agent import Document

    deck = Document.open("results.pptx")
    print(deck.to_outline())                    # a Markdown reading view, with shape ids
    for block in deck.find_text("4,285"):       # raw text and addresses, to edit from
        deck.resolve(block.address).text = block.text.replace("4,285", "4,310")
    table = deck.shape("257.3#5").table
    table.cell_by_label("営業利益", "当期実績").text = "520億円"   # strict: LabelError
    deck.shape("257.25").chart.series["売上高（億円）"].set_value("Q3", 4310)
    problems = deck.validate()                  # what PowerPoint would repair, if anything
    deck.save("restated.pptx")

The outline's text is escaped Markdown: never paste it into ``set_text`` (which warns with
:class:`MarkdownEscapeWarning` when it looks that way).  More::

    deck = Document.open("deck.pptx")
    title = deck.slides[0].shapes[0]
    title.set_text("Q3 results")
    title.move_by(dx=914400)          # one inch right

    run = deck.resolve(f"{title.id}/p0/r0")
    run.format(bold=True, color="accent1")      # a theme colour stays a theme colour
    title.fill = "accent1 lumMod=20% lumOff=80%"
    title.line.width = 12700
    run.set_hyperlink("https://example.org")

    copy = deck.slides[0].duplicate()           # notes and charts copied, media shared
    added = deck.add_slide("Title and Content", index=1)
    added.add_picture("chart.png", left=914400, top=914400, width=4 * 914400)
    deck.move_slide(copy, 0)
    deck.delete_slide(deck.slides[-1])          # with its notes, charts, orphaned media

    chart = deck.shape("257.5").chart           # cache and embedded workbook, together
    chart.series[0].set_value(2, 4285)
    chart.add_category("Q4", [4400, 530])       # cells, formulas and the table follow
    chart.set_title("Revenue")
    deck.shape("258.3").diagram.set_text(1, "Ship it")   # SmartArt node text

    slide = deck.slides[1]                      # shapes from nothing, styled by the theme
    box = slide.add_shape("roundRect", 914400, 914400, 1828800, 914400, text="Plan",
                          adjustments={"adj": 30000})
    goal = slide.add_shape("ellipse", 5486400, 2743200, 1371600, 1371600, text="Ship")
    arrow = slide.add_connector("elbow", (box, 3), (goal, 2), line={"tail": "triangle"})
    goal.move_by(dy=-914400)                    # the connector is re-routed, still attached
    slide.add_textbox(914400, 4572000, 3657600, 369332, "Notes")
    slide.add_table(3, 4, 914400, 5029200, 7315200, 1097280).table.cell(0, 0).text = "Q1"

    fresh = Document.new(size="16:9", title="Plan", author="Ada")   # or template="brand.potx"
    cover = fresh.add_slide("Title Slide")      # placeholders inherit from layout and master
    cover.shapes[0].set_text("Plan")
    fresh.save("plan.pptx")                     # fresh.save_as_template("plan.potx")

    png = deck.slides[0].render_png(width=1280)
    svg = deck.slides[0].render_svg(full_state=True)  # edit it anywhere, then:
    deck.apply_svg(svg)                         # typed edits, raw XML, new shapes; one undo
    deck.save("edited.pptx")

Lengths are EMU throughout: 914,400 per inch, 12,700 per point; ``Pt(6)`` is 6 points as
EMU (76,200).  Font sizes alone are in points.
"""

from __future__ import annotations

from .edit.chart import Chart, ChartDataError, ChartDataWarning, Series, SeriesList
from .edit.color import Color
from .edit.creating import Adjustments
from .edit.deck import TemplateOpenedWarning
from .edit.diagram import Diagram, DiagramNode
from .edit.effective import EffectiveParagraph
from .edit.document import EMU_PER_INCH, EMU_PER_POINT, Document, ImageSize, Shape, Slide
from .edit.fill import Arrowhead, Fill, GradientStop, LineFormat
from .edit.fit import (WRAP_MARGIN, Overflow, RowsFit, TextFit, TextMeasure, fit_box,
                       measure_text)
from .edit.ids import ShapeId
from .edit.labels import LabelError
from .edit.slides import Layout, LayoutPlaceholder, LayoutShape, Placeholder
from .edit.table import Region, Table, TableCell
from .edit.theme import Theme, ThemeFonts, ThemeRoles, Tint
from .edit.presets import PRESETS
from .edit.textspec import ParagraphSpec, RunSpec, TextSpec
from .edit.units import Cm, Inches, Pt, UnitWarning, to_pt
from .edit.text import Bullet, Hyperlink, MarkdownEscapeWarning, Paragraph, Run, TextFrame
from .fullstate import ApplyReport, FullStateError, NotFullStateSvg
from .outline import OutlineWarning
from .outline.read import TextBlock
from .oxml.package import OoxmlPackage, PresentationPackage

__version__ = "0.0.1"

__all__ = [
    "Adjustments",
    "ApplyReport",
    "Arrowhead",
    "Bullet",
    "Chart",
    "ChartDataError",
    "ChartDataWarning",
    "Cm",
    "Color",
    "Diagram",
    "DiagramNode",
    "Document",
    "EffectiveParagraph",
    "EMU_PER_INCH",
    "EMU_PER_POINT",
    "Fill",
    "FullStateError",
    "GradientStop",
    "Hyperlink",
    "ImageSize",
    "Inches",
    "LabelError",
    "Layout",
    "LayoutPlaceholder",
    "LayoutShape",
    "LineFormat",
    "MarkdownEscapeWarning",
    "NotFullStateSvg",
    "OoxmlPackage",
    "OutlineWarning",
    "Overflow",
    "Paragraph",
    "ParagraphSpec",
    "Placeholder",
    "PRESETS",
    "PresentationPackage",
    "Pt",
    "Region",
    "RowsFit",
    "Run",
    "RunSpec",
    "Series",
    "SeriesList",
    "Shape",
    "ShapeId",
    "Slide",
    "Table",
    "TableCell",
    "TextBlock",
    "TextFit",
    "TextMeasure",
    "TextSpec",
    "TemplateOpenedWarning",
    "TextFrame",
    "Theme",
    "ThemeFonts",
    "ThemeRoles",
    "Tint",
    "UnitWarning",
    "WRAP_MARGIN",
    "fit_box",
    "measure_text",
    "to_pt",
]
