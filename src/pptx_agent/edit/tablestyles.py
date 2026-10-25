"""Table styles: the deck's own (``ppt/tableStyles.xml``) and PowerPoint's built-in ones.

A table names its style by GUID in ``a:tblPr/a:tableStyleId``; its Table Design flags --
header row, total row, first and last column, banded rows and columns -- are ``a:tblPr``'s
``firstRow``, ``lastRow``, ``firstCol``, ``lastCol``, ``bandRow`` and ``bandCol``, and say
which of the style's regions apply.  ``ppt/tableStyles.xml`` holds the deck's custom styles
(a brand template's own) and ``def``, the style PowerPoint gives a new table.  The built-in
styles live in PowerPoint: a deck names one by GUID and need not define it (see
:data:`BUILTIN_TABLE_STYLES`).
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from ooxml_common.drawingml import table_styles as _catalogue

from ..oxml.xml import Element, append_in_order, find, findall, make, qn, remove
from .errors import ArgumentError

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

REL_TABLE_STYLES = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                    "tableStyles")

#: PowerPoint's built-in table styles, GUID -> name, as its Table Design gallery names them:
#: the 74 styles of ooxml-common's catalogue (:mod:`ooxml_common.drawingml.table_styles`),
#: measured from PowerPoint 16 (each GUID confirmed by drawing it in PowerPoint: an id
#: PowerPoint does not know draws as no style), in its order.
BUILTIN_TABLE_STYLES: dict[str, str] = {guid: spec["name"] for guid, spec
                                        in _catalogue.BUILTIN_TABLE_STYLES.items()}
#: How the built-in names go, for a reader who wants one without the list of 74.
BUILTIN_NAMES = ("No Style, No Grid; No Style, Table Grid; Themed Style 1|2 - Accent 1-6; "
                 "Light Style 1|2|3, Medium Style 1|2|3|4 and Dark Style 1, each plain or "
                 "- Accent 1-6; Dark Style 2, plain or - Accent 1/Accent 2, 3/4, 5/6")
#: PowerPoint's own default, Medium Style 2 - Accent 1: a new deck's ``def``.
DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

GUID = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
                  r"[0-9A-Fa-f]{12}\}")
_BUILTIN_BY_KEY = {guid.upper(): guid for guid in BUILTIN_TABLE_STYLES}


def _name_key(name: str) -> str:
    """A style name as compared: case, runs of spaces and the dash's kind ignored."""
    text = re.sub(r"\s*[-–—]\s*", " - ", name.strip())
    return re.sub(r"\s+", " ", text).casefold()


def styles_part(document: "Document") -> str | None:
    """The deck's ``tableStyles.xml``, or ``None`` when it has none."""
    package = document.package
    parts = package.related_parts_of_type(package.presentation_part(), REL_TABLE_STYLES)
    return parts[0] if parts else None


def _styles_root(document: "Document") -> Element | None:
    part = styles_part(document)
    return document.package.tree(part) if part is not None else None


def deck_default(document: "Document") -> str | None:
    """``tableStyles.xml``'s ``def``: the style the deck gives a new table, or ``None``."""
    root = _styles_root(document)
    value = root.get("def") if root is not None else None
    return value or None


def table_styles(document: "Document") -> list[dict]:
    """See :attr:`pptx_agent.edit.document.Document.table_styles`."""
    out: list[dict] = []
    seen: set[str] = set()
    root = _styles_root(document)
    default = deck_default(document)
    for node in findall(root, "a:tblStyle") if root is not None else []:
        style_id = node.get("styleId")
        if not style_id or style_id.upper() in seen:
            continue
        seen.add(style_id.upper())
        builtin = style_id.upper() in _BUILTIN_BY_KEY
        name = node.get("styleName") or (BUILTIN_TABLE_STYLES[_BUILTIN_BY_KEY[style_id.upper()]]
                                         if builtin else "")
        out.append(_entry(style_id, name, builtin, default))
    for style_id, name in BUILTIN_TABLE_STYLES.items():
        if style_id.upper() not in seen:
            out.append(_entry(style_id, name, True, default))
    return out


def _entry(style_id: str, name: str, builtin: bool, default: str | None) -> dict:
    entry = {"id": style_id, "name": name, "builtin": builtin}
    if default is not None and style_id.upper() == default.upper():
        entry["default"] = True
    return entry


def describe(document: "Document", style_id: str | None) -> dict | None:
    """``{"id", "name", "builtin"}`` for a style id (the deck's definition first), or
    ``None`` for no id.  An id neither the deck nor PowerPoint has keeps ``name`` ``None``:
    PowerPoint draws such a table with no style."""
    if not style_id:
        return None
    key = style_id.upper()
    for entry in table_styles(document):
        if entry["id"].upper() == key:
            return {"id": entry["id"], "name": entry["name"], "builtin": entry["builtin"]}
    return {"id": style_id, "name": None, "builtin": False}


def resolve(document: "Document", style: str) -> str:
    """The GUID of ``style``: a style id or a name the deck or PowerPoint has (case, spacing
    and the dash's kind do not matter; the deck's own style wins over a built-in of the
    same name).  ``ValueError`` names what there is."""
    if not isinstance(style, str) or not style.strip():
        raise ArgumentError(f"a table style is a name or a {{GUID}}, not {style!r}")
    styles = table_styles(document)
    text = style.strip()
    if GUID.fullmatch(text):
        for entry in styles:
            if entry["id"].upper() == text.upper():
                return entry["id"]
        raise ArgumentError(f"no table style {text} in this deck or built into PowerPoint "
                         "(PowerPoint would draw the table with no style); "
                         "deck.table_styles lists them")
    key = _name_key(text)
    for entry in styles:
        if entry["name"] and _name_key(entry["name"]) == key:
            return entry["id"]
    own = [entry["name"] for entry in styles if not entry["builtin"] and entry["name"]]
    raise ArgumentError(f"no table style named {text!r}; the deck's own: "
                     f"{', '.join(own) if own else 'none'}; built in: {BUILTIN_NAMES}")


#: Table Design flag -> ``a:tblPr`` attribute.
FLAGS = {"first_row": "firstRow", "last_row": "lastRow", "first_column": "firstCol",
         "last_column": "lastCol", "banded_rows": "bandRow", "banded_columns": "bandCol"}


def _band(index: int, first: bool, last: bool, count: int) -> int | None:
    """0 for a first band, 1 for a second, ``None`` outside the banding: a header or total
    row (column) does not band and does not advance it, as pptx2svg measured."""
    start = 1 if first else 0
    end = count - (1 if last else 0)
    if not start <= index < end:
        return None
    return (index - start) % 2


def _builtin_text_style(style_id: str) -> Element | None:
    """A built-in style's text colours as an ``a:tblStyle`` holding only ``a:tcTxStyle``s,
    or ``None`` for a GUID that is not built in."""
    from .color import Color

    key = style_id.upper()
    regions = _catalogue.builtin_text_colors(key)
    if regions is None:
        return None
    style = make("a:tblStyle")
    style.set("styleId", style_id)
    style.set("styleName", BUILTIN_TABLE_STYLES.get(key, style_id))
    for region, (slot, transforms) in regions.items():
        text = make("a:tcTxStyle")
        spec = " ".join([slot] + [f"{kind}={value}" for kind, value in transforms])
        text.append(Color.parse(spec).to_element())
        holder = make(f"a:{region}")
        holder.append(text)
        style.append(holder)
    return style


def cell_text_styles(document: "Document", table: Element, row: int,
                     column: int) -> list[tuple[Element, str]]:
    """The ``a:tcTxStyle`` of each region of the table's style that covers a cell, highest
    precedence first, with a label (``"table style 'Brand' firstRow"``): the corner
    cells, then the header and total rows, the first and last columns, the row and column
    bands, the whole table -- ECMA-376 Part 1, 20.1.4.2, in the order pptx2svg draws them.
    The style is the deck's own definition in ``tableStyles.xml`` when it has one, else
    PowerPoint's built-in style of that GUID, whose text colours were measured
    (:func:`ooxml_common.drawingml.table_styles.builtin_text_colors`: PowerPoint does not
    write a built-in definition into the deck); a GUID neither knows gives none, and PowerPoint draws such a table unstyled."""
    properties = table.find(qn("a:tblPr"))
    style_id = style_id_of(properties) or deck_default(document)
    if not style_id:
        return []
    root = _styles_root(document)
    style = next((node for node in findall(root, "a:tblStyle")
                  if node.get("styleId") == style_id), None) if root is not None else None
    if style is None:
        style = _builtin_text_style(style_id)
        if style is None:
            return []
    flags = read_flags(properties)
    rows = findall(table, "a:tr")
    columns = len(findall(table, "a:tblGrid/a:gridCol"))
    cell = None
    if 0 <= row < len(rows):
        cells = findall(rows[row], "a:tc")
        cell = cells[column] if 0 <= column < len(cells) else None
    row_span = max(1, int(cell.get("rowSpan", "1"))) if cell is not None else 1
    grid_span = max(1, int(cell.get("gridSpan", "1"))) if cell is not None else 1
    first_row = flags["first_row"] and row == 0
    last_row = flags["last_row"] and row + row_span >= len(rows)
    first_col = flags["first_column"] and column == 0
    last_col = flags["last_column"] and column + grid_span >= columns
    regions = ["wholeTbl"]
    if flags["banded_columns"]:
        band = _band(column, flags["first_column"], flags["last_column"], columns)
        if band is not None:
            regions.append("band1V" if band == 0 else "band2V")
    if flags["banded_rows"]:
        band = _band(row, flags["first_row"], flags["last_row"], len(rows))
        if band is not None:
            regions.append("band1H" if band == 0 else "band2H")
    regions += [name for name, on in (("lastCol", last_col), ("firstCol", first_col),
                                      ("lastRow", last_row), ("firstRow", first_row)) if on]
    corner = ("nwCell" if first_row and first_col else "neCell" if first_row and last_col
              else "swCell" if last_row and first_col else "seCell" if last_row and last_col
              else None)
    if corner:
        regions.append(corner)
    name = style.get("styleName") or style_id
    out = []
    for region in reversed(regions):
        node = find(style, f"a:{region}/a:tcTxStyle")
        if node is not None:
            out.append((node, f"table style {name!r} {region}"))
    return out


def read_flags(properties: Element | None) -> dict[str, bool]:
    """The Table Design flags ``a:tblPr`` sets (each off when absent)."""
    return {flag: properties is not None and properties.get(attribute) in ("1", "true")
            for flag, attribute in FLAGS.items()}


def write_flags(properties: Element, changes: dict[str, bool]) -> None:
    """Set flags on ``a:tblPr``: a flag that is on is ``"1"`` and one that is off goes, as
    PowerPoint writes them, the attributes in the schema's order."""
    flags = read_flags(properties)
    flags.update(changes)
    on = {attribute for flag, attribute in FLAGS.items() if flags[flag]}
    others = {name: value for name, value in properties.attrib.items()
              if name not in FLAGS.values()}
    properties.attrib.clear()
    for name in _SCHEMA_ORDER:
        if name in on:
            properties.set(name, "1")
        elif name in others:
            properties.set(name, others.pop(name))
    for name, value in others.items():
        properties.set(name, value)


#: ``CT_TableProperties``' attributes in the schema's order.
_SCHEMA_ORDER = ("rtl", "firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol")


def style_id_of(properties: Element | None) -> str | None:
    """``a:tblPr/a:tableStyleId``'s text, or ``None``."""
    node = find(properties, "a:tableStyleId") if properties is not None else None
    text = (node.text or "").strip() if node is not None else ""
    return text or None


def write_style_id(properties: Element, style_id: str) -> None:
    """Set ``a:tableStyleId``, where the schema puts it; an inline ``a:tableStyle``, its
    alternative, goes."""
    for node in list(properties):
        if node.tag in (qn("a:tableStyle"), qn("a:tableStyleId")):
            remove(node)
    node = make("a:tableStyleId")
    node.text = style_id
    append_in_order(properties, node)


__all__ = ["BUILTIN_TABLE_STYLES", "DEFAULT_TABLE_STYLE", "FLAGS", "resolve", "table_styles"]
