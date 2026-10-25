"""The shared chart and SmartArt tools' deck handlers: ``edit_chart`` (``action: "read"``
reads a chart) and ``edit_smartart``.  A chart's cached values and its embedded workbook change together."""

from __future__ import annotations

from typing import Any

from ooxml_edit.charts.actions import (STYLE_ACTIONS, add_style_actions, apply_add_style,
                                       apply_style_action, category_labels, chart_style_json,
                                       resolve_category, resolve_node, resolve_series,
                                       scale_axis)
from ooxml_edit.tools import Result, ToolError, shared
from ooxml_edit.tools.shared import CHART_TYPES

from .common import KIND, box_emu, library_errors, shape as shape_at, slide as slide_at

_LEGEND = {"right": "r", "left": "l", "top": "t", "bottom": "b"}

#: What ``add`` needs, said when it is missing.
_ADD_NEEDS = {"chart_type": "one of " + ", ".join(CHART_TYPES),
              "data": "each series as {name, values}, one value per category",
              "box": "the frame as {x, y, w, h} in pt",
              "categories": "the category labels, in order"}


def _chart(deck, target: str):
    frame = shape_at(deck, target)
    chart = frame.chart if frame.kind == "graphic_frame" else None
    if chart is None:
        raise ToolError("invalid_arguments", f"{target} is a {frame.kind}, not a chart",
                        field="target")
    return frame, chart


def _categories(chart) -> list:
    """The categories a read lists: ISO dates for dates, else the labels as they are."""
    return category_labels(chart) if chart.category_dates is not None else list(chart.categories)


def chart_json(frame, chart) -> dict[str, Any]:
    data = {"address": frame.id, "type": "dot" if chart.is_dot else chart.chart_type,
            "title": chart.title,
            "categories": _categories(chart),
            "series": [{"name": series.name, "values": list(series.values)}
                       for series in chart.series],
            "number_formats": chart.number_formats}
    if chart.category_dates is not None:
        # Dates kept as serials: listed as ISO dates, and how the chart shows them.
        data["category_format"] = chart.category_format
    if chart.gap_width is not None:
        data["gap_width"] = chart.gap_width
    return data


def _add(call, target, chart_type, categories, data, box, width, values, number_format, text,
         position, ref, axis=None, scale=None, gridlines=None, data_labels=None, colors=None,
         style=None, part=None) -> Result:
    """``action: "add"``: a new chart on the slide ``target``, in ``box``, with the look
    its arguments give (``add_style_actions``: the same chart as the separate calls)."""
    if width is not None:
        raise ToolError("invalid_arguments", "width is for documents; on a slide give box",
                        field="width")
    for name, given in (("chart_type", chart_type), ("data", data), ("box", box)):
        if given is None:
            raise ToolError("invalid_arguments", f"add needs {name}: {_ADD_NEEDS[name]}",
                            field=name)
    if chart_type in ("scatter", "bubble"):
        if values is None or categories is not None:
            raise ToolError("invalid_arguments", f"a {chart_type} chart takes its x values in "
                            "values, not categories", field="values")
        labels = list(values)
    else:
        if categories is None:
            raise ToolError("invalid_arguments", "add needs categories: "
                            + _ADD_NEEDS["categories"], field="categories")
        if values is not None:
            raise ToolError("invalid_arguments", "values are a scatter or bubble chart's x "
                            "values; give each series' values in data", field="values")
        labels = list(categories)
    if len(labels) * max(len(data), 1) > 10_000:
        raise ToolError("limit", "at most 10,000 chart values", field="data")
    host = slide_at(call.document, target, field="target")
    x, y, w, h = box_emu(box)
    if position not in (None, "none", "right", "left", "top", "bottom"):
        raise ToolError("invalid_arguments", f"{position} is a label position; a legend is "
                        "right, left, top, bottom or none", field="position")
    look = add_style_actions(chart_type, data, data_labels=data_labels, colors=colors,
                             style=style, part=part)
    legend = None if position == "none" else (position or "default")
    reverse = (scale or {}).get("reverse") and axis == "category"
    rest = {k: v for k, v in (scale or {}).items() if not (reverse and k == "reverse")}
    with call.document.batch():  # the chart and its look: one undo step
        frame = host.add_chart(chart_type, labels, [dict(entry) for entry in data], x, y, w, h,
                               title=text, legend=legend, number_format=number_format,
                               category_order="reversed" if reverse else "normal",
                               gridlines=gridlines != "none")
        chart = call.document.shape(frame.id).chart
        if rest:  # no axis: the chart's one value axis (refused, and undone, if it has more)
            apply_style_action(chart, "set_axis_scale", scale=rest,
                               axis=axis or scale_axis(chart, rest, action="add with scale"))
        if gridlines not in (None, "major", "none"):
            apply_style_action(chart, "set_gridlines", axis="value" if chart_type not in
                               ("scatter", "bubble") else "y", gridlines=gridlines)
        apply_add_style(chart, look)
    if ref:
        call.define_ref(ref, frame.id)
    call.touch(host.slide_id)
    return Result(summary=f"Added a {chart_type} chart {frame.id}", created=[frame.id],
                  data=chart_json(frame, call.document.shape(frame.id).chart))


def _read(frame, chart) -> Result:
    data = chart_json(frame, chart)
    data["legend"] = chart.has_legend
    data["axes"] = {axis: chart.axis_title(axis) for axis in chart.axes}
    data.update(chart_style_json(chart))
    data["workbook"] = chart.workbook_values()
    return Result(summary=f"Chart {frame.id}: {len(data['series'])} series, "
                  f"{len(data['categories'])} categories", data=data)


@shared.handler("edit_chart", kind=KIND)
@library_errors
def edit_chart(call, doc, target, action, chart_type=None, categories=None, data=None, box=None,
               width=None, number_format=None, series=None, category=None, values=None,
               value=None, text=None, axis=None, position=None, ref=None, scale=None,
               gridlines=None, style=None, part=None, labels=None, cells=None, secondary=None,
               data_labels=None, colors=None):
    if action == "add":
        return _add(call, target, chart_type, categories, data, box, width, values,
                    number_format, text, position, ref, axis, scale, gridlines, data_labels,
                    colors, style, part)
    frame, chart = _chart(call.document, target)
    if action == "read":
        return _read(frame, chart)
    if action in STYLE_ACTIONS:
        chosen = None if series is None else resolve_series(chart, series, action).index
        point = (None if category is None or chosen is None
                 else resolve_category(chart, category, action))
        apply_style_action(chart, action, series=chosen, category=point, axis=axis,
                           scale=scale, gridlines=gridlines, style=style, part=part,
                           labels=labels, cells=cells, position=position,
                           chart_type=chart_type, secondary=secondary)
        call.touch(frame._slide.slide_id)
        fresh = call.document.shape(frame.id).chart
        data = chart_json(frame, fresh)
        data.update(chart_style_json(fresh))
        return Result(summary=f"{action} on chart {frame.id}", changed=[frame.id], data=data)

    def need(name, given):
        if given is None:
            raise ToolError("invalid_arguments", f"{action} needs {name}: "
                            + _expected(chart, action, name), field=name)
        return given

    if action == "set_values":
        resolve_series(chart, series, action).set_values(need("values", values))
    elif action == "set_value":
        resolve_series(chart, series, action).set_value(
            resolve_category(chart, category, action), need("value", value))
    elif action == "add_category":
        chart.add_category(need("text", text), list(values) if values is not None else None)
    elif action == "remove_category":
        chart.remove_category(resolve_category(chart, category, action))
    elif action == "rename_category":
        chart.set_category(resolve_category(chart, category, action), need("text", text))
    elif action == "add_series":
        chart.add_series(need("text", text), list(values) if values is not None else None)
    elif action == "remove_series":
        chart.remove_series(resolve_series(chart, series, action).index)
    elif action == "rename_series":
        resolve_series(chart, series, action).set_name(need("text", text))
    elif action == "set_title":
        chart.set_title(text)
    elif action == "set_axis_title":
        chart.set_axis_title(need("axis", axis), text)
    elif action == "set_legend":
        where = position or "right"
        if where == "none":
            chart.set_legend(False)
        else:
            chart.set_legend(True, _LEGEND[where])
    elif action == "show_data_labels":
        chosen = None if series is None else [resolve_series(chart, series, action).index]
        chart.set_data_labels(True, number_format=number_format, series=chosen)
    elif action == "hide_data_labels":
        chosen = None if series is None else [resolve_series(chart, series, action).index]
        chart.set_data_labels(False, series=chosen)
    elif action == "set_gap_width":
        given = need("value", value)
        if given != int(given):
            raise ToolError("invalid_arguments", "a gap width is a whole percentage, 0-500",
                            field="value")
        chart.set_gap_width(int(given))
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id).chart
    return Result(summary=f"{action} on chart {frame.id}", changed=[frame.id],
                  data=chart_json(frame, fresh))


def _expected(chart, action: str, name: str) -> str:
    """What a missing ``name`` of ``action`` could be, on this chart."""
    if name == "axis":
        return " or ".join(chart.axes) or "the chart has no axes"
    if name == "values":
        return f"one number per category ({len(chart.categories)})"
    if name == "value":
        return ("the gap as a whole % of a bar's width, 0-500" if action == "set_gap_width"
                else "the number")
    return {"add_category": "the new category's label", "rename_category": "the new label",
            "add_series": "the new series' name"}.get(action, "the new name")


@shared.handler("edit_smartart", kind=KIND)
@library_errors
def edit_smartart(call, doc, target, action, node=None, text=None):
    frame = shape_at(call.document, target)
    diagram = frame.diagram if frame.kind == "graphic_frame" else None
    if diagram is None:
        raise ToolError("invalid_arguments", f"{target} is not SmartArt", field="target")
    chosen = None if action == "add_node" else resolve_node(diagram, node, action)
    if action == "set_text":
        diagram.set_text(chosen, text or "")
    elif action == "add_node":
        diagram.add_node(text or "")
    elif action == "remove_node":
        diagram.remove_node(chosen)
    elif action == "add_child":
        chosen.add_child(text or "")
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id).diagram
    return Result(summary=f"{action} on SmartArt {frame.id}", changed=[frame.id],
                  data={"nodes": [{"n": i, "text": n.text, "level": n.level}
                                  for i, n in enumerate(fresh.nodes)]})


TOOLS = [edit_chart, edit_smartart]
