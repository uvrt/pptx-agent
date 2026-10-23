"""The size a chart's text is drawn at: its title, axis titles, axis labels, data labels and
legend, read from the chart part.

The cascade is the one pptx2svg draws charts with (``ooxml_common.chart.layout``), so a
size here is the size a render shows:

* **Axis labels, the legend and data labels** take the innermost ``c:txPr`` that states a
  size -- a point's ``c:dLbl``, its series' ``c:dLbls``, the chart group's ``c:dLbls`` for
  data labels; the axis' or the legend's own -- then the chart space's ``c:txPr``
  (ECMA-376 Part 1, 21.2.2.216), then :data:`~ooxml_common.chart.layout.DEFAULT_CHART_FONT_PT`,
  10 pt, measured in PowerPoint for an axis and a legend with no ``c:txPr`` anywhere.  A
  data label's own rich text (``c:dLbl/c:tx/c:rich``) states its runs' sizes over that.
* **The chart title and axis titles** are DrawingML text (``c:rich``) and resolve as a
  plain text box's does: the run, the paragraph, the rich text's list style, then the
  master's ``otherStyle`` and the presentation's default text style, then 18 pt -- measured
  by pptx2svg: an unstyled title is drawn at 18 pt, not at the chart's 10.  A title's own
  ``c:txPr`` (21.2.2.217; what ooxml-edit's ``set_text_style`` writes for a title with no
  rich text) is read after the rich text's paragraph.  The chart space's ``c:txPr`` is not
  applied to a rich title, as pptx2svg does not apply it.
* **An automatic chart title** -- a ``c:title`` with no ``c:tx``, which PowerPoint shows
  as the series' name -- takes its own ``c:txPr``'s size, else **1.2 times** the chart
  space's ``c:txPr`` size rounded to the nearest point, else 18 pt
  (:func:`ooxml_common.chart.layout.powerpoint_auto_title_size`, as pptx2svg draws it).
  Measured on PowerPoint 16 for Mac, 2026-10-16: a chart-wide 8, 10, 11 and 24 pt drew it
  at 10, 12, 13 and 29 pt, bold; none at 18 pt; an unsized title ``c:txPr`` changed
  nothing.  A rich title with no size anywhere stays 18 pt, whatever the chart-wide size.

**A rich chart title that states no font** -- no paragraph ``a:defRPr`` and no face on its
runs (``ooxml_common.chart.layout.title_states_no_font``) -- is drawn by PowerPoint in
**Arial**, regular, at the run's size or 18 pt, not in the theme's font a text box gets,
whatever the chart-wide ``c:txPr`` names (measured on PowerPoint 16 for Mac, 2026-10-16,
as pptx2svg draws it).  :func:`unstated_title_fonts` lists each such chart with the one
``edit_chart`` call that states it: ``set_text_style``, part ``title``, font ``+mn-lt`` (the
theme's body font, a reference rather than a name), ``bold`` false and, where the title
states no size, the size ``edit_chart add`` writes for a new chart's title (ooxml-edit's
``POWERPOINT_LOOK``, 18.62 pt).  ``bold`` false is needed: the call writes a paragraph
``a:defRPr``, and a title with one is drawn **bold** unless ``b="0"`` is stated (measured:
``+mn-lt`` and 18.62 pt without ``b`` came out Aptos bold).  A title that has paragraph
defaults is drawn in the theme's face already and is not listed.  Nothing is changed by
itself.

What is checked is what PowerPoint shows: a title the chart states (its text, or the
automatic one), axes not deleted and with tick labels, data labels that show something, a
legend that is there.
"""

from __future__ import annotations

from dataclasses import dataclass
from xml.etree import ElementTree

from lxml import etree
from ooxml_common.chart.layout import (DEFAULT_CHART_FONT_PT, powerpoint_auto_title_size,
                                       title_states_no_font)
from ooxml_common.chart.read import AXIS_ELEMENTS, parse_chart_space
from ooxml_common.drawingml.read_text import parse_text_body
from ooxml_common.xmlutil import child, local_name
from ooxml_edit.charts import POWERPOINT_LOOK

from ..oxml.xml import qn
from . import inherit

#: The parts of a chart's text, named as ``edit_chart``'s ``set_text_style`` names them.
PARTS = ("title", "axis_titles", "axis_labels", "data_labels", "legend")

_CHART = "the chart's text properties"
_AXES = {"catAx": "category", "valAx": "value", "dateAx": "date", "serAx": "series"}
_DEFAULT = f"the chart default of {DEFAULT_CHART_FONT_PT:g} pt"
#: What PowerPoint draws a title in when nothing on the way states a size.
TITLE_DEFAULT = inherit.DEFAULT_SIZE


@dataclass(frozen=True)
class ChartText:
    """One piece of a chart's text: its ``part`` (:data:`PARTS`), the ``size`` it is drawn at
    in points and ``source``, where that size comes from."""

    part: str
    size: float
    source: str


#: The theme's body font, as ``edit_chart``'s ``style.font`` names it.
THEME_FONT = "+mn-lt"
#: The size ``edit_chart add`` writes for a new chart's title, in points (ooxml-edit's
#: ``POWERPOINT_LOOK``: 18.62, measured on PowerPoint's new chart).
NEW_TITLE_SIZE = POWERPOINT_LOOK.title / 100


def _read(shape):
    """``(chart space element, SourceChart)`` of a chart shape, or ``None``."""
    from ooxml_edit.charts.host import chart_part

    from .chart import graphic_host

    host = graphic_host(shape)
    part = chart_part(host)
    root = host.package.tree(part) if part else None
    if root is None:
        return None
    # ooxml-common reads the standard library's tree; a copy keeps lxml's comments out.
    space = ElementTree.fromstring(etree.tostring(root))
    chart = parse_chart_space(space)
    return None if chart is None else (space, chart)


def title_font_fix(address: str, shape) -> dict | None:
    """The ``edit_chart`` call that states the font of a chart's shown rich title that
    states no face (see the module's notes), or ``None``."""
    read = _read(shape)
    if read is None:
        return None
    chart = read[1]
    if chart.auto_title_deleted or chart.title is None or not chart.title.plain:
        return None
    if not title_states_no_font(chart):
        return None
    style: dict = {"font": THEME_FONT}
    if not _states_size(chart.title.rich):
        style["size"] = NEW_TITLE_SIZE
    style["bold"] = False
    return {"action": "set_text_style", "target": address, "part": "title", "style": style}


def _states_size(rich) -> bool:
    """Whether a title's runs state a size (its paragraphs state no defaults: see
    :func:`title_font_fix`)."""
    return any(run.properties is not None and run.properties.font_size
               for paragraph in rich.paragraphs for run in paragraph.runs)


def unstated_title_fonts(slide) -> list[dict]:
    """``chart_title_font``: each chart on ``slide`` whose title states no font, with the
    ``edit_chart`` call that states it (``fix``), one entry per chart."""
    from .fit import _flatten

    out = []
    for shape in _flatten(slide.shapes):
        if shape.kind == "graphic_frame" and shape.has_chart:
            fix = title_font_fix(shape.id, shape)
            if fix is not None:
                out.append({"chart": shape.id, "fix": fix})
    return out


def chart_text(shape) -> list[ChartText]:
    """The text a chart shape shows, one :class:`ChartText` per piece, in :data:`PARTS`
    order; ``[]`` for a shape without a readable chart."""
    read = _read(shape)
    if read is None:
        return []
    space, chart = read
    chart_size = _txpr(chart.text_properties)
    styles = _master_styles(shape)
    out: list[ChartText] = []

    def labels(*bodies) -> tuple[float, str]:
        for body, source in bodies:
            points = _txpr(body)
            if points:
                return points, source
        return (chart_size, _CHART) if chart_size else (DEFAULT_CHART_FONT_PT, _DEFAULT)

    node = child(space, "chart")
    if chart.title_stated:
        title = child(node, "title")
        if child(title, "tx") is None:
            out.append(ChartText("title", *_auto_title_size(chart, chart_size)))
        else:
            out += [ChartText("title", *found) for found in
                    _title_sizes(chart.title, child(title, "txPr"), styles, "title")]
    plot_area = child(node, "plotArea")
    axes = [n for n in (plot_area if plot_area is not None else ())
            if local_name(n.tag) in AXIS_ELEMENTS]
    for axis, element in zip(chart.axes, axes):
        if axis.delete or not axis.title_stated:
            continue
        out += [ChartText("axis_titles", *found) for found in
                _title_sizes(axis.title, child(child(element, "title"), "txPr"), styles,
                             "axis title")]
    for axis in chart.axes:
        if not axis.delete and axis.tick_label_position != "none":
            which = _AXES.get(axis.kind, "the")
            out.append(ChartText("axis_labels", *labels(
                (axis.text_properties, f"the {which} axis' text properties"))))
    for plot in chart.plots:
        group = plot.data_labels
        for series in plot.series:
            own = series.data_labels
            if _shown(own, group):
                out.append(ChartText("data_labels", *labels(
                    (own.text_properties if own else None, "the series' data labels"),
                    (group.text_properties if group else None, "the chart's data labels"))))
            for point in (own.overrides.values() if own else ()):
                if not _shown(point, own, group):
                    continue
                inherited = labels(
                    (point.text_properties, "a point's data label"),
                    (own.text_properties, "the series' data labels"),
                    (group.text_properties if group else None, "the chart's data labels"))
                rich = point.text.rich if point.text is not None else None
                runs = [run for paragraph in (rich.paragraphs if rich else ())
                        for run in paragraph.runs if run.text.strip()]
                if not runs:
                    out.append(ChartText("data_labels", *inherited))
                for run in runs:
                    stated = run.properties.font_size if run.properties else None
                    out.append(ChartText("data_labels", *((stated, "a point's data label")
                                                          if stated else inherited)))
    if chart.legend is not None:
        out.append(ChartText("legend", *labels(
            (chart.legend.text_properties, "the legend's text properties"))))
    return out


def _txpr(body) -> float | None:
    """A ``c:txPr``'s size: its first ``a:defRPr@sz``, in points (as pptx2svg reads it)."""
    for paragraph in (body.paragraphs if body is not None else ()):
        properties = paragraph.properties
        if properties is not None and properties.default_run_properties is not None:
            return properties.default_run_properties.font_size or None
    return None


_FLAGS = ("show_value", "show_category_name", "show_series_name", "show_percent",
          "show_bubble_size", "show_range")


def _shown(*sources) -> bool:
    """Whether data labels show anything: no ``c:delete`` on the way, and a show flag whose
    innermost statement is true (pptx2svg's ``_read_labels``)."""
    present = [source for source in sources if source is not None]
    if not present or any(source.delete for source in present):
        return False
    for name in _FLAGS:
        stated = next((getattr(s, name) for s in present if getattr(s, name) is not None), None)
        if stated:
            return True
    return False


def _master_styles(shape) -> list:
    """The list styles a plain text box on the shape's slide falls back to: the
    presentation's default text style (not the master's ``otherStyle``: measured,
    :mod:`ooxml_common.drawingml.text_chain`)."""
    slide = shape._slide
    package = slide.document.package
    master = inherit.master_of(package, inherit.layout_of(package, slide.part_path))
    return inherit.list_styles(package, master, None, None, None, False)[1:]


def _auto_title_size(chart, chart_size: float | None) -> tuple[float, str]:
    """``(size, source)`` of an automatic chart title (see the module's notes)."""
    size = powerpoint_auto_title_size(chart)
    if _txpr(chart.title_text_properties):
        return size, "the title's text properties"
    if chart_size:
        return size, f"1.2 times the chart's text properties ({chart_size:g} pt), rounded"
    return size, f"PowerPoint's default {size:g} pt for an automatic title"


def _title_sizes(title, txpr_node, styles, what: str) -> list[tuple[float, str]]:
    """``(size, source)`` for each run of a title's text (one entry when it has none)."""
    from .effective import describe

    rich = title.rich if title is not None else None
    own = parse_text_body(txpr_node) if txpr_node is not None else None
    paragraphs = rich.paragraphs if rich is not None else []
    runs = [(paragraph, run) for paragraph in paragraphs for run in paragraph.runs
            if run.text.strip()]
    if not runs:
        runs = [(paragraphs[0] if paragraphs else None, None)]
    found: dict[tuple[float, str], None] = {}
    for paragraph, run in runs:
        properties = paragraph.properties if paragraph is not None else None
        level = (properties.level or 0) if properties is not None else 0
        listed = (rich.list_style.levels[level] if rich is not None and rich.list_style
                  and level < len(rich.list_style.levels) else None)
        candidates = [
            (run.properties.font_size if run is not None and run.properties else None,
             f"the {what}'s run"),
            (properties.default_run_properties.font_size
             if properties is not None and properties.default_run_properties else None,
             f"the {what}'s paragraph"),
            (_txpr(own), f"the {what}'s text properties"),
            (listed.default_run_properties.font_size
             if listed is not None and listed.default_run_properties else None,
             f"the {what}'s list style")]
        for node in inherit.level_properties(styles, level):
            defaults = node.find(qn("a:defRPr"))
            if defaults is not None and defaults.get("sz"):
                candidates.append((int(defaults.get("sz")) / 100, describe(defaults)))
        size, source = next(((points, where) for points, where in candidates if points),
                            (TITLE_DEFAULT, f"PowerPoint's default {TITLE_DEFAULT:g} pt"))
        found.setdefault((round(size, 2), source))
    return list(found)


__all__ = ["NEW_TITLE_SIZE", "PARTS", "THEME_FONT", "TITLE_DEFAULT", "ChartText", "chart_text",
           "title_font_fix", "unstated_title_fonts"]
