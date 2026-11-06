"""Values styled by a scale on a deck (production round 36, R31): the deck side of
``style_by_value`` and ``check``'s ``scale`` (:mod:`ooxml_edit.tools.scales` reads the scale,
chooses each value's step and reports; this module reads the values and writes the steps).

**Targets** (:func:`scale_values`, the deck ``DocumentFormat.scale_values``):

* a shape (``256.5``): its text is the value; the shape is filled and its text coloured;
* a table cell (``256.7/cell1,2``): its text, the cell filled and its text coloured;
* a table's column (``256.7/col1``, from 0): every body cell of it -- the first row left out
  where the table's Header Row is on (``a:tblPr firstRow``), as the table style draws it;
* values written by reference (``{{d:scores/*/score}}``): each value a text tool wrote from a
  matching reference and still where it was written (the shape's kept facts) -- its cell, or
  its shape.

**Writing** (:func:`write`): a deck has no character styles, so each value is formatted
directly -- the shape's or cell's fill (``a:solidFill`` with the step's ``RRGGBB``) and every
run's text colour, as ``ppt_set_shape`` ``fill`` and ``ppt_format_text`` ``color`` write them.
A value a picture or chart frame holds is skipped.

**The step a value has now**: the step whose fill the shape or cell has (explicit, solid, the
step's ``RRGGBB``); ``exact`` where every run's colour is the step's text colour too.
"""

from __future__ import annotations

import re
from typing import Any, Mapping

from ooxml_edit.tools import Result, ScaleValue, ToolError
from ooxml_edit.tools import scales as _scales

from ..edit import provenance as _provenance
from .common import numbers as slide_numbers
from .common import resolve, touch_shape

_CELL = re.compile(r"cell(\d+),(\d+)")
#: A slide's furniture: its footers, dates and slide numbers are no values.
FURNITURE = ("ftr", "dt", "sldNum")


def scale_values(call: Any, scale: _scales.Scale, scope: Mapping[str, Any]) -> list[ScaleValue]:
    """``DocumentFormat.scale_values`` for a deck (the module's notes); ``scope``: ``check``'s
    ``slides``."""
    if scope.get("pages") is not None:
        raise ToolError("invalid_arguments", "a deck has slides, not pages: give slides", field="pages")
    if scope.get("range") is not None:
        raise ToolError("invalid_arguments", "check range names a document's blocks (docx); leave it "
                        "out here: slides say which slides", field="range")
    deck = call.document
    wanted = None
    if scope.get("slides") is not None:
        wanted = {deck.slides[n - 1].slide_id for n in slide_numbers(deck, scope.get("slides"))}
    out: list[ScaleValue] = []
    for target in scale.targets:
        try:
            found = _read_target(deck, scale, target)
        except ToolError as error:
            found = [ScaleValue(target, skip=error.message)]
        for value in found:
            if wanted is not None and value.skip is None and value.handle[0]._slide.slide_id not in wanted:
                continue
            out.append(value)
    return out


def _read_target(deck: Any, scale: _scales.Scale, target: str) -> list[ScaleValue]:
    column = _scales.column_target(target)
    if column is not None:
        return _column(deck, scale, target, *column)
    wanted = _scales.reference_target(target)
    if wanted is not None:
        return _referenced(deck, scale, target, wanted)
    found = resolve(deck, target)
    head, _, rest = target.partition("/")
    match = _CELL.fullmatch(rest) if rest else None
    if match is not None:
        shape = deck.shape(head)
        return [_value(scale, shape, (int(match.group(1)), int(match.group(2))), None)]
    if rest:
        raise ToolError("invalid_arguments", f"{target} is part of a shape's text: a deck's value is a "
                        "shape (256.5) or a cell (256.7/cell1,2), filled whole")
    if getattr(found, "has_table", False):
        raise ToolError("invalid_arguments", f"{target} is a table: name its column ({target}/col1) or "
                        "its cells")
    return [_value(scale, found, None, None)]


def _location(shape: Any, cell: tuple[int, int] | None) -> str:
    return shape.id if cell is None else f"{shape.id}/cell{cell[0]},{cell[1]}"


def _value(scale: _scales.Scale, shape: Any, cell: tuple[int, int] | None, given: str | None) -> ScaleValue:
    address = _location(shape, cell)
    where = {"slide": f"s:{shape._slide.slide_id}"}
    if cell is None:
        placeholder = shape.placeholder
        if placeholder is not None and placeholder.type in FURNITURE:
            return ScaleValue(address, given=given, skip=f"{address} is the slide's {placeholder.type}")
        if shape.text_frame is None or shape.kind != "shape":
            return ScaleValue(address, given=given, skip=f"{address} is a {shape.kind}: it holds no value "
                              "to fill")
        holder = shape
    else:
        holder = shape.table.cell(*cell)
    text = (_provenance.text_of(shape, cell) or "").strip()
    if not text:
        return ScaleValue(address, given=given, skip=f"{address} holds no text")
    fill = holder.fill
    hexed = fill.color.value.upper() if fill is not None and fill.kind == "solid" and fill.color is not None \
        and fill.color.kind == "rgb" else None
    now = next((s.index for s in scale.steps if s.fill == hexed), None)
    exact, differs = True, None
    if now is not None:
        colours = {_hex(run.color) for paragraph in holder.text_frame.paragraphs for run in paragraph.runs
                   if run.text.strip()}
        if colours != {scale.steps[now].text}:
            exact = False
            shown = sorted(c or "inherited" for c in colours)
            differs = f"text colour {', '.join(shown)}"
    return ScaleValue(address, text, given=given, now=now, exact=exact, differs=differs, where=where,
                      handle=(shape, cell))


def _hex(colour: Any) -> str | None:
    return colour.value.upper() if colour is not None and colour.kind == "rgb" else None


def _column(deck: Any, scale: _scales.Scale, target: str, table_id: str, column: int) -> list[ScaleValue]:
    shape = resolve(deck, table_id)
    table = shape.table if getattr(shape, "has_table", False) else None
    if table is None:
        raise ToolError("invalid_arguments", f"{table_id} is not a table: /col names a table's column")
    if column >= table.columns:
        raise ToolError("not_found", f"no col{column}: {table_id} has {table.columns} columns, "
                        f"col0-col{table.columns - 1}")
    first = 1 if table.flags.get("first_row") and table.rows > 1 else 0
    out = []
    for row in range(first, table.rows):
        cell = table.cell(row, column)
        origin = cell.origin if hasattr(cell, "origin") else None
        if origin is not None and (origin.row, origin.column) != (row, column):
            continue                                  # hidden by a merge
        out.append(_value(scale, shape, (row, column), target))
    return out


def _referenced(deck: Any, scale: _scales.Scale, target: str, wanted: str) -> list[ScaleValue]:
    out: list[ScaleValue] = []
    seen: set[str] = set()
    for slide in deck.slides:
        for shape in slide.shapes:
            for fact in _provenance.facts_found(shape):
                source = fact.get("source")
                if not isinstance(source, str) or not _scales.reference_matches(wanted, source):
                    continue
                key = fact.get("cell")
                cell = (key[0], key[1]) if isinstance(key, list) and len(key) == 2 else None
                address = _location(shape, cell)
                if address not in seen:
                    seen.add(address)
                    out.append(_value(scale, shape, cell, target))
    if not out:
        raise ToolError("not_found", f"{target}: no value written by a reference like it is still in the "
                        "deck's text")
    return out


def write(call: Any, scale: _scales.Scale, plan: list[tuple[ScaleValue, _scales.Step]],
          preview: bool) -> Result:
    """Fill each value's shape or cell and colour its text (the module's notes)."""
    changed: list[str] = []
    for value, step in plan:
        shape, cell = value.handle
        if not preview:
            holder = shape if cell is None else shape.table.cell(*cell)
            holder.fill = f"#{step.fill}"
            for paragraph in holder.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.color = f"#{step.text}"
            touch_shape(call, shape)
        if value.target not in changed:
            changed.append(value.target)
    result = Result(changed=[] if preview else changed)
    result.data = {"styles": "none: a deck has no character styles; fill and text colour set directly"}
    return result


def style_by_value(call: Any, doc: str, scale: Any, preview: bool | None = None) -> Result:
    """``style_by_value``'s deck handler (its handler in :mod:`.shared_tools`)."""
    return _scales.apply(call, scale, bool(preview), write)
