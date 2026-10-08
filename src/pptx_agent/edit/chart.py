"""Charts: the data, titles and legend of a chart, kept in step with its workbook.

The editing itself is :mod:`ooxml_edit.charts.chart` -- a chart part and its embedded
workbook are the same DrawingML in a deck and in a Word document, so it lives there once.
What is left here is what makes it PowerPoint's: :func:`graphic_host` tells it where a
chart lives on a slide (the graphic frame, the slide part whose relationships name the
chart, the deck's undo step) and what to call things in messages ("PowerPoint's Edit Data",
"the deck"), and :class:`Chart` takes a shape, as it always has.
"""

from __future__ import annotations

from contextlib import contextmanager
from typing import TYPE_CHECKING, Callable, Iterator

from ..oxml import xml as _oxml  # noqa: F401  (the slide's vocabulary is registered first)
from ooxml_edit.charts import chart as _chart
from ooxml_edit.charts.chart import (  # noqa: F401  (re-exported)
    AXIS_TAGS,
    C_NS,
    PLOT_TAGS,
    PLOT_TYPES,
    REL_PACKAGE,
    ChartDataError,
    ChartDataWarning,
    chart_model,
    decode_number,
)
from ooxml_edit.charts.create import POWERPOINT_LOOK
from ooxml_edit.charts.host import GraphicHost, chart_part
from ooxml_edit.xml import qn

from .labels import find_label

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape

#: What PowerPoint writes into the runs of a title it creates.
TITLE_LANG = "en-US"


def graphic_host(shape: "Shape") -> GraphicHost:
    """A slide's graphic frame as ooxml-edit's charts and diagrams see it."""
    slide = shape._slide
    document = slide.document

    @contextmanager
    def edit() -> Iterator[None]:
        with document.batch():
            document.history.checkpoint()
            yield

    return GraphicHost(package=document.package, part=slide.part_path, frame=shape._element,
                       edit=edit, address=shape.id, application="PowerPoint",
                       document="deck", lang=TITLE_LANG, look=POWERPOINT_LOOK)


class Chart(_chart.Chart):
    """The chart in a graphic frame -- fully editable: values, categories, series, titles
    and legend, with the chart's cached data **and** its embedded workbook (what Edit Data
    opens) kept in step, so PowerPoint never shows stale numbers.  Re-resolved from the
    document on every call, like :class:`~pptx_agent.edit.table.Table`, so it survives undo::

        chart = deck.shape("257.25").chart
        chart.series["売上高（億円）"].set_value("Q3", 4310)     # by label, strictly
        chart.series[0].set_value(2, 4310)                     # or by position
        chart.add_category("Q4", [4400, 530]); chart.set_title("Revenue")

    Series are numbered in document order (``chart.series[0]``) and found by name
    (``chart.series["name"]``, :meth:`series_named`); categories by position or by label
    (:meth:`category_index`).  Values are numbers or ``None`` for a blank; category labels
    are text (numbers, for a chart whose categories are numeric or dates).  A label lookup
    that finds nothing or more than one raises :class:`~pptx_agent.LabelError` listing the
    labels.  A chart without an embedded workbook is edited in its cache only, with a
    :class:`~pptx_agent.ChartDataWarning`.
    """

    def __init__(self, resolve: Callable[[], "Shape"]) -> None:
        super().__init__(lambda: graphic_host(resolve()))

    # -- reading (documented here; the work is ooxml-edit's) ----------------------------

    @property
    def chart_type(self) -> str | None:
        """The first plot's type: ``"bar"`` (columns too), ``"line"``, ``"pie"``,
        ``"doughnut"``, ``"area"``, ``"scatter"``, ``"radar"``...::

            chart.chart_type                # 'bar'
        """
        return super().chart_type

    @property
    def chart_types(self) -> list[str]:
        """Every plot's type, in order -- two for a combination chart::

            chart.chart_types               # ['bar', 'line']
        """
        return super().chart_types

    @property
    def categories(self) -> list:
        """The category labels, in order, from the first series that has any::

            chart.categories                # ['Q1', 'Q2', 'Q3']
        """
        return super().categories

    @property
    def title(self) -> str | None:
        """The chart's own title text; ``None`` when it shows none of its own::

            chart.title                     # '四半期業績'
        """
        return super().title

    @property
    def has_legend(self) -> bool:
        """Whether the chart shows a legend; settable (as :meth:`set_legend`)::

            chart.has_legend = False
        """
        return super().has_legend

    @has_legend.setter
    def has_legend(self, value: bool) -> None:
        self.set_legend(value)

    def axis_title(self, axis: str) -> str | None:
        """The ``"category"`` (x) or ``"value"`` (y) axis's title::

            chart.axis_title("value")       # '億円'
        """
        return super().axis_title(axis)

    @property
    def axes(self) -> list[str]:
        """The axes the chart has -- ``["category", "value"]``, or none for a pie::

            chart.axes
        """
        return super().axes

    @property
    def workbook_part(self) -> str | None:
        """The embedded workbook kept in step with the chart, or ``None`` when there is
        none -- then edits change the chart's cache only, with a
        :class:`~pptx_agent.ChartDataWarning`::

            chart.workbook_part             # 'ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx'
        """
        return super().workbook_part

    def workbook_values(self) -> dict | None:
        """What the embedded workbook -- the sheet Edit Data opens -- holds for this chart,
        read cell by cell through each series' formulas, in the shape of :attr:`data`:
        ``{"sheet", "categories", "series": [{"name", "values"}]}``.  Read-only; ``None``
        when the chart has no workbook.  Equal to the cache when the two agree, which every
        edit here keeps true::

            book = chart.workbook_values()
            assert book["series"] == chart.data["series"]

        A series whose formula is not one range on one sheet (a defined name, several
        areas) reads ``None`` for what it cannot reach.
        """
        # The shared reader (ooxml-edit) gives each value with its cell reference; this
        # keeps the shape pptx-agent always returned.
        book = super().workbook_values()
        if book["workbook"] is None:
            return None
        return {"sheet": book["sheet"],
                "categories": (book["categories"] or {}).get("values") or [],
                "series": [{"name": (series["name"] or {}).get("value"),
                            "values": (series["values"] or {}).get("values")}
                           for series in book["series"]]}

    @property
    def data(self) -> dict:
        """Everything above as one JSON-ready dictionary (types, title, axis titles,
        legend, format, categories, series)::

            json.dumps(chart.data, ensure_ascii=False)
        """
        return super().data

    # -- editing (documented here; the work is ooxml-edit's) ----------------------------

    def add_category(self, label=None, values=None, *, index: int | None = None) -> "Chart":
        """Insert a category at ``index`` (default: last), with a value per series -- a list
        in series order or ``{series name: value}``; a series given nothing gets a blank.
        In the workbook the cells below move down, and formulas and a table grow::

            chart.add_category("Q4", {"売上高（億円）": 4400, "営業利益（億円）": 530})
        """
        return super().add_category(label, values, index=index)

    def remove_series(self, which) -> "Chart":
        """Delete a series -- a :class:`Series`, a position or a name -- from the chart and
        its workbook::

            chart.remove_series("営業利益（億円）")
        """
        if isinstance(which, str):
            which = self.series_named(which)
        return super().remove_series(which)

    def set_title(self, text: str | None) -> "Chart":
        """Title the chart; ``None`` removes the title (and keeps an automatic one away)::

            chart.set_title("四半期業績")
        """
        return super().set_title(text)

    def set_axis_title(self, axis: str, text: str | None) -> "Chart":
        """Title the ``"category"`` or ``"value"`` axis; ``None`` removes it::

            chart.set_axis_title("value", "億円")
        """
        return super().set_axis_title(axis, text)

    def set_legend(self, visible: bool, position: str = "r") -> "Chart":
        """Show the legend at ``position`` (``r``, ``l``, ``t``, ``b``, ``tr``), or hide it::

            chart.set_legend(True, "b")
        """
        return super().set_legend(visible, position)

    # -- formatting (documented here; the work is ooxml-edit's) -------------------------

    @property
    def gap_width(self) -> int | None:
        """The space between bar or column clusters, percent of a bar's width; ``None``
        without bars::

            chart.gap_width                 # 219, PowerPoint's for a new column chart
        """
        return super().gap_width

    def set_gap_width(self, percent: int) -> "Chart":
        """The space between clusters, 0-500% of a bar's width; narrower gaps make wider
        bars, with room for data labels::

            chart.set_gap_width(60)
        """
        return super().set_gap_width(percent)

    @property
    def data_labels(self) -> list:
        """Per series, whether its values are shown as labels, and their own number format
        (``None`` when they follow the values')::

            chart.data_labels               # [{'series': 0, 'shown': True, 'format': '0.0'}]
        """
        return super().data_labels

    def set_data_labels(self, visible: bool, *, number_format: str | None = None,
                        series=None) -> "Chart":
        """Show or hide every point's value, on every series or those at ``series``
        (positions), in ``number_format`` (an Excel code) or the values' own; styled as
        PowerPoint styles data labels::

            chart.set_data_labels(True, number_format='"€"#,##0.0"m"')
        """
        return super().set_data_labels(visible, number_format=number_format, series=series)

    # -- series and categories by label --------------------------------------------------

    @property
    def series(self) -> "SeriesList":
        """Every series, in document order: a list, also indexed by name::

            chart.series[0].values          # [3980, 4120, 4285]
            chart.series["営業利益（億円）"]   # the same as series_named(...)
        """
        return SeriesList(self, [Series(self, index) for index in range(len(self._series()))])

    @property
    def series_names(self) -> list:
        """Each series' name, in order (``None`` for one without a name)::

            chart.series_names              # ['売上高（億円）', '営業利益（億円）']
        """
        return [series.name for series in self.series]

    def series_named(self, name: str) -> "Series":
        """The one series called ``name``.  Strict: a name no series has, or two series
        share, raises :class:`~pptx_agent.LabelError` listing the names (whitespace
        collapsed; failing an exact match, full-width forms and case are ignored if that
        finds exactly one)::

            chart.series_named("売上高（億円）").set_value("Q3", 4310)
        """
        index = find_label(name, enumerate(self.series_names), what="series",
                           where=f"{self.address} (chart)")
        return Series(self, index)

    def category_index(self, label) -> int:
        """The position of the one category labelled ``label``; strict, like
        :meth:`series_named`.  A number finds a numeric category::

            chart.category_index("Q3")      # 2
        """
        categories = self.categories
        if isinstance(label, (int, float)) and not isinstance(label, bool):
            hits = [k for k, c in enumerate(categories)
                    if isinstance(c, (int, float)) and c == label]
            if len(hits) == 1:
                return hits[0]
        return find_label(label, enumerate(categories), what="category",
                          where=f"{self.address} (chart)")

    def set_category(self, index, label) -> "Chart":
        """Relabel a category -- by position, or by its current label -- in every series'
        cache and in the workbook::

            chart.set_category("Q3", "Q3 (restated)")
        """
        return super().set_category(_position(self, index), label)

    def remove_category(self, index) -> "Chart":
        """Delete a category -- by position or label -- and its value in every series,
        cache and workbook::

            chart.remove_category("Q1")
        """
        return super().remove_category(_position(self, index))

    def add_series(self, name, values=None, *, index=None) -> "Series":
        """Add a series at ``index`` (default: last), with one value per category, coloured
        as PowerPoint colours a new series; the workbook gets its column::

            chart.add_series("Forecast", [4000, 4200, 4400])
        """
        added = super().add_series(name, values, index=index)
        return Series(self, added.index)

    # -- number formats ------------------------------------------------------------------

    @property
    def number_format(self) -> str | None:
        """The number format the values carry: the ``formatCode`` cached with the first
        series' values -- the source cells' format in the workbook, such as ``General``,
        ``0.0`` or ``#,##0``.  ``None`` when the values cache none (literal values, or
        none cached).

        It says how the *workbook* formats the numbers, not what the slide shows, and
        writing a value never changes it.  Values are stored exactly as given (``12.1``
        stays ``12.1``); what decides the precision on screen is the format of whatever
        displays them -- data labels and the value axis -- which use this one only when
        they are source-linked.  ``General`` shows as many decimals as the number has.
        See :attr:`number_formats` for those::

            chart.number_format             # 'General'
        """
        return super().number_format

    @property
    def number_formats(self) -> dict:
        """What decides how numbers look on the slide: ``{"values", "data_labels",
        "value_axis"}``.  ``values`` is :attr:`number_format`; ``data_labels`` the format
        of the first series' data labels that show values (``None`` when no data label
        shows a value -- then a value's precision changes only a bar's height or a point's
        position, not any text); ``value_axis`` the value axis's (``None`` without one).
        A source-linked format reads as the values' own::

            chart.number_formats
            # {'values': 'General', 'data_labels': None, 'value_axis': 'General'}
        """
        root = self._root()
        values = self.number_format
        return {"values": values, "data_labels": _label_format(root, values),
                "value_axis": _format_of(self._axis(root, "value"), values)}


class SeriesList(list):
    """A chart's series: a list, also indexed by name, as strictly as
    :meth:`Chart.series_named`::

        chart.series[0], chart.series["売上高（億円）"], "売上高（億円）" in chart.series
    """

    def __init__(self, chart: Chart, items) -> None:
        super().__init__(items)
        self._chart = chart

    def __getitem__(self, key):
        if isinstance(key, str):
            return self._chart.series_named(key)
        return super().__getitem__(key)

    def __contains__(self, item) -> bool:
        if isinstance(item, str):
            return item in self._chart.series_names
        return super().__contains__(item)


class Series(_chart.Series):
    """One series of a :class:`Chart`, by position.  Re-resolved on every call::

        series = chart.series["営業利益（億円）"]
        series.set_value("Q3", 520)        # by category label
        series.value("Q3")                 # 520
        series.set_values([465, 488, 520])
    """

    @property
    def name(self) -> str | None:
        """The series name (its legend entry), ``None`` when it has none; settable, and the
        workbook's header cell follows::

            series.name = "売上高（億円・修正後）"
        """
        return super().name

    @name.setter
    def name(self, value: str) -> None:
        self.set_name(value)

    def set_name(self, value: str) -> "Series":
        """Rename the series, in the chart and its workbook::

            series.set_name("売上高（億円・修正後）")
        """
        return super().set_name(value)

    @property
    def values(self) -> list:
        """The values, one per category; ``None`` for a blank::

            series.values                   # [3980, 4120, 4285]
        """
        return super().values

    @property
    def categories(self) -> list:
        """This series' category labels::

            series.categories               # ['Q1', 'Q2', 'Q3']
        """
        return super().categories

    def set_values(self, values) -> "Series":
        """Every point at once, one value per category, in the cache and the workbook::

            series.set_values([42.7, 29.8, 18.9, 8.6])
        """
        return super().set_values(values)

    def value(self, category):
        """The value at a category -- by position (an ``int``) or by label::

            series.value("Q3")              # 4285
        """
        return self.values[_position(self.chart, category)]

    def set_value(self, category, value: "float | int | None") -> "Series":
        """Set one point -- by position (an ``int``) or by category label -- in the
        chart's cache and in its workbook cell.  ``None`` blanks it.  A label must name
        exactly one category (:class:`~pptx_agent.LabelError` otherwise)::

            series.set_value("その他", 369)     # same as set_value(3, 369) here
        """
        return super().set_value(_position(self.chart, category), value)

    def items(self) -> list[tuple]:
        """``(category, value)`` pairs, in order::

            dict(series.items())            # {'Q1': 3980, 'Q2': 4120, 'Q3': 4285}
        """
        return list(zip(self.categories, self.values))


def _position(chart: Chart, category) -> int:
    """A category given by position stays one; anything else is a label."""
    if isinstance(category, int) and not isinstance(category, bool):
        return category
    return chart.category_index(category)


def _flag(owner, tag: str) -> bool | None:
    node = owner.find(qn(tag))
    return None if node is None else node.get("val", "1") in ("1", "true")


def _shows_number(owner) -> bool | None:
    """Whether a ``c:dLbls`` or ``c:dLbl`` shows a value or a percentage; ``None`` when it
    does not say."""
    if owner is None:
        return None
    if _flag(owner, "c:delete"):
        return False
    flags = [_flag(owner, "c:showVal"), _flag(owner, "c:showPercent")]
    if all(flag is None for flag in flags):
        return None
    return any(flags)


def _format_of(owner, values: str | None) -> str | None:
    if owner is None:
        return None
    node = owner.find(qn("c:numFmt"))
    if node is None or node.get("sourceLinked") in ("1", "true"):
        return values
    return node.get("formatCode")


def _label_format(root, values: str | None) -> str | None:
    """The format of the first data label that shows a number: a point's own ``c:dLbl``
    wins over its series' ``c:dLbls``, which wins over its plot's."""
    for series in _chart._series_of(root):
        own = series.element.find(qn("c:dLbls"))
        plot = series.plot.find(qn("c:dLbls")) if series.plot is not None else None
        points = {node.find(qn("c:idx")).get("val"): node
                  for node in (own.findall(qn("c:dLbl")) if own is not None else [])
                  if node.find(qn("c:idx")) is not None}
        for index in range(max(series.values.count, 1)):
            chain = [points.get(str(index)), own, plot]
            for owner in chain:
                shown = _shows_number(owner)
                if shown is None:
                    continue
                if shown:
                    return _format_of(owner, values)
                break
    return None


def chart_part_of(shape: "Shape") -> str | None:
    """The chart part a graphic frame points at, or ``None``."""
    return chart_part(graphic_host(shape))


def workbook_part(package, chart_part: str) -> tuple[str | None, str]:
    """The embedded workbook behind a chart, or ``None`` and why there is none."""
    return _chart.workbook_part(package, chart_part, document="deck")


__all__ = ["Chart", "ChartDataError", "ChartDataWarning", "Series", "SeriesList", "chart_model",
           "chart_part_of", "graphic_host", "workbook_part"]
