"""The shared chart and SmartArt tools' deck handlers: ``edit_chart`` (``action: "read"``
reads a chart) and ``edit_smartart``.  A chart's cached values and its embedded workbook change together."""

from __future__ import annotations

from typing import Any

from ooxml_edit.charts.actions import (DATA_ACTIONS, STYLE_ACTIONS, add_plan,
                                       apply_data_action, apply_smartart_action,
                                       apply_style_action, chart_json as _chart_json,
                                       chart_read_json, chart_style_json, finish_add,
                                       resolve_category, resolve_series, smartart_json)
from ooxml_edit.tools import Result, ToolError, shared

from .common import KIND, box_emu, library_errors, shape as shape_at, slide as slide_at


def _chart(deck, target: str):
    frame = shape_at(deck, target)
    chart = frame.chart if frame.kind == "graphic_frame" else None
    if chart is None:
        raise ToolError("invalid_arguments", f"{target} is a {frame.kind}, not a chart",
                        field="target")
    return frame, chart


def chart_json(frame, chart) -> dict[str, Any]:
    """What an ``edit_chart`` result holds (ooxml-edit's ``chart_json``), the chart by its
    shape address."""
    return _chart_json(chart, "address")


def _add(call, target, chart_type, categories, data, box, width, values, number_format, text,
         position, ref, axis=None, scale=None, gridlines=None, data_labels=None, colors=None,
         style=None, part=None) -> Result:
    """``action: "add"``: a new chart on the slide ``target``, in ``box``, with the look
    its arguments give (ooxml-edit's ``add_plan`` / ``finish_add``: the same chart as the
    separate calls)."""
    if width is not None:
        raise ToolError("invalid_arguments", "width is for documents; on a slide give box",
                        field="width")
    plan = add_plan(chart_type, categories, data, values, position=position, axis=axis,
                    scale=scale, data_labels=data_labels, colors=colors, style=style,
                    part=part, required={"box": (box, "the frame as {x, y, w, h} in pt")})
    host = slide_at(call.document, target, field="target")
    x, y, w, h = box_emu(box)
    with call.document.batch():  # the chart and its look: one undo step
        frame = host.add_chart(chart_type, plan["labels"], [dict(entry) for entry in data],
                               x, y, w, h, title=text, legend=plan["legend"],
                               number_format=number_format,
                               category_order=plan["category_order"],
                               gridlines=gridlines != "none")
        finish_add(call.document.shape(frame.id).chart, chart_type, plan, axis=axis,
                   gridlines=gridlines)
    if ref:
        call.define_ref(ref, frame.id)
    call.touch(host.slide_id)
    return Result(summary=f"Added a {chart_type} chart {frame.id}", created=[frame.id],
                  data=chart_json(frame, call.document.shape(frame.id).chart))


def _read(frame, chart) -> Result:
    data = chart_read_json(chart, "address")
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
    if action not in DATA_ACTIONS:  # pragma: no cover -- the schema's enum
        raise ToolError("invalid_arguments", f"unknown action {action!r}", field="action")
    apply_data_action(chart, action, series=series, category=category, values=values,
                      value=value, text=text, axis=axis, position=position,
                      number_format=number_format)
    call.touch(frame._slide.slide_id)
    fresh = call.document.shape(frame.id).chart
    return Result(summary=f"{action} on chart {frame.id}", changed=[frame.id],
                  data=chart_json(frame, fresh))


@shared.handler("edit_smartart", kind=KIND)
@library_errors
def edit_smartart(call, doc, target, action, node=None, text=None):
    frame = shape_at(call.document, target)
    diagram = frame.diagram if frame.kind == "graphic_frame" else None
    if diagram is None:
        raise ToolError("invalid_arguments", f"{target} is not SmartArt", field="target")
    apply_smartart_action(diagram, action, node=node, text=text)
    call.touch(frame._slide.slide_id)
    return Result(summary=f"{action} on SmartArt {frame.id}", changed=[frame.id],
                  data=smartart_json(call.document.shape(frame.id).diagram))


TOOLS = [edit_chart, edit_smartart]
