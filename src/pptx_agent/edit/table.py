"""Tables: cells, their text, fill and borders, row and column sizing, and structure.

**The merge model, which is what pptx-svg got wrong.**  An ``a:tbl`` is a full grid: every row
has one ``a:tc`` per ``a:gridCol``, including the cells a merge hides.  A merged region is
spelled the way PowerPoint writes it::

    top-left (the origin)       gridSpan=W rowSpan=H     holds the text
    rest of the top row         rowSpan=H hMerge=1
    rest of the left column     gridSpan=W vMerge=1
    everything else             hMerge=1 vMerge=1

So inserting a row *inside* a region must lengthen it -- every top-row cell's ``rowSpan``
grows, and the new row's cells in that region become ``vMerge`` (plus ``gridSpan``/``hMerge``
by column) -- while inserting at a region's edge must leave it alone.  Deleting the region's
top row hands the origin's text and properties to the row below.  Columns are the mirror
image.  Getting any of this wrong produces a grid whose spans overrun the table, which
PowerPoint "repairs" by discarding the table's formatting.

Lengths are EMU.  Inserting or deleting rows and columns keeps the graphic frame's extent in
step with the grid, the way PowerPoint does.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass
from typing import TYPE_CHECKING, Callable, Iterator

from ..oxml.xml import (
    FILL_TAGS,
    Element,
    append_in_order,
    find,
    get_int,
    make,
    qn,
    remove,
    set_attr,
    set_int,
    subelement,
)
from .color import Color
from .labels import find_label
from .fill import (
    Fill,
    LineFormat,
    fill_element,
    gradient_fill,
    read_fill,
    write_fill,
)
from .text import TextFrame

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape
    from .fit import RowsFit

#: Cell edge -> the ``a:tcPr`` child that draws it.
BORDER_TAGS: dict[str, str] = {
    "left": "a:lnL",
    "right": "a:lnR",
    "top": "a:lnT",
    "bottom": "a:lnB",
    "diagonal_down": "a:lnTlToBr",
    "diagonal_up": "a:lnBlToTr",
}

#: PowerPoint's per-row and per-column ids live in extensions; copies must get new values.
_ROW_ID_EXT = "{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}"
_COL_ID_EXT = "{9D8B030D-6E8A-4147-A177-3AD203B41FA5}"

_SPAN_ATTRIBUTES = ("gridSpan", "rowSpan", "hMerge", "vMerge")


@dataclass(frozen=True, order=True)
class Region:
    """A merged area, inclusive: rows ``top..bottom``, columns ``left..right``.

    For example::

        table.region_at(0, 1)                         # Region(top=0, left=1, bottom=0, right=2)
    """

    top: int
    left: int
    bottom: int
    right: int

    @property
    def height(self) -> int:
        """Rows the region spans.

        For example::

            table.region_at(0, 1).height
        """
        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        """Columns the region spans.

        For example::

            table.region_at(0, 1).width
        """
        return self.right - self.left + 1

    def contains(self, row: int, column: int) -> bool:
        """Whether grid position ``(row, column)`` is inside.

        For example::

            region.contains(0, 2)
        """
        return self.top <= row <= self.bottom and self.left <= column <= self.right


class Table:
    """The ``a:tbl`` inside a graphic frame -- fully editable: cell text and formatting,
    rows and columns, sizes, merges.  Re-resolved from the document on every call::

        table = deck.shape("257.3#5").table
        table.cell_by_label("営業利益", "当期実績").text = "520億円"   # by label, strictly
        table.cell(3, 1).text = "520億円"                              # or by position
        table.insert_row(like=1); table.delete_column(4)

    Rows and columns are 0-based; row 0 is usually the header.
    """

    def __init__(self, resolve: Callable[[], "Shape"]) -> None:
        self._resolve = resolve

    # -- shape ------------------------------------------------------------------------------

    @property
    def address(self) -> str:
        """The table's shape id.

        For example::

            table.address                                # '257.3#5'
        """
        return self._resolve().id

    @property
    def rows(self) -> int:
        """The number of rows.

        For example::

            table.rows                                   # 6
        """
        return len(self._rows())

    @property
    def columns(self) -> int:
        """The number of grid columns.

        For example::

            table.columns                                # 5
        """
        return len(self._grid_columns())

    def __len__(self) -> int:
        return self.rows

    def cell(self, row: int, column: int) -> "TableCell":
        """The cell at ``(row, column)``, 0-based; raises for a position outside the grid.

        For example::

            table.cell(1, 1).text = "4,310億円"
        """
        self._tc(row, column)  # validate now, so a bad address fails where it was made
        return TableCell(self, row, column)

    def iter_cells(self) -> Iterator["TableCell"]:
        """Every grid position, row by row -- merged-over ones too (see :class:`TableCell`)::

            for cell in table.iter_cells():
                print(cell.address, cell.text)
        """
        for row in range(self.rows):
            for column in range(self.columns):
                yield TableCell(self, row, column)

    # -- by label --------------------------------------------------------------------------

    def row_labels(self, column: int = 0) -> list[str]:
        """The text of every row's cell in ``column`` (the label column), top to bottom; a
        position a merge covers reads its region's text::

            table.row_labels()        # ['項目', '売上高', '売上総利益', ...]
        """
        return [self.cell(row, column).origin.text for row in range(self.rows)]

    def column_labels(self, row: int = 0) -> list[str]:
        """The text of every column's cell in ``row`` (the header row), left to right::

            table.column_labels()     # ['項目', '当期実績', '前年同期', '増減額', '増減率']
        """
        return [self.cell(row, column).origin.text for column in range(self.columns)]

    def find_row(self, label: str, *, column: int = 0) -> int:
        """The index of the one row whose cell in ``column`` reads ``label``.

        Strict: a label no row has, or more than one row has, raises
        :class:`~pptx_agent.LabelError` listing the labels -- never a near miss.  Whitespace
        is collapsed; failing an exact match, full-width and half-width forms and case are
        ignored, if that finds exactly one (see :mod:`pptx_agent.edit.labels`)::

            table.find_row("営業利益")              # 3
        """
        return find_label(label, enumerate(self.row_labels(column)), what="row",
                          where=f"{self.address} (labels in column {column})")

    def find_column(self, label: str, *, row: int = 0) -> int:
        """The index of the one column whose cell in ``row`` reads ``label``; as strict as
        :meth:`find_row`::

            table.find_column("当期実績")           # 1
        """
        return find_label(label, enumerate(self.column_labels(row)), what="column",
                          where=f"{self.address} (labels in row {row})")

    def cell_by_label(self, row_label: str, column_label: str, *, header_row: int = 0,
                      label_column: int = 0) -> "TableCell":
        """The cell where the row labelled ``row_label`` (in column ``label_column``) meets
        the column headed ``column_label`` (in row ``header_row``) -- by what the table
        says, not by counting.  Raises :class:`~pptx_agent.LabelError`, listing the labels,
        when either names no row or column or more than one::

            table.cell_by_label("営業利益", "当期実績").text = "520億円"
        """
        return self.cell(self.find_row(row_label, column=label_column),
                         self.find_column(column_label, row=header_row))

    @property
    def column_widths(self) -> list[int]:
        """Each grid column's width, EMU.

        For example::

            table.column_widths
        """
        return [get_int(col, "w", 0) or 0 for col in self._grid_columns()]

    @property
    def row_heights(self) -> list[int]:
        """Each row's height, EMU.

        For example::

            table.row_heights
        """
        return [get_int(row, "h", 0) or 0 for row in self._rows()]

    @property
    def drawn_row_heights(self) -> list[int]:
        """Each row's height as PowerPoint draws it, EMU: :attr:`row_heights` is a minimum,
        and a row whose text needs more grows to fit it, moving the rows below down.
        Measured with pptx2svg's table layout and text measurement; the stored heights
        without pptx2svg.

        For example::

            sum(table.drawn_row_heights)            # the table's height on the slide
        """
        from .fit import table_heights

        return table_heights(self._shape())

    def rows_fitting(self, bottom: int | None = None) -> "RowsFit":
        """How many rows, from the first, fit above ``bottom`` (a y on the slide, EMU; the
        slide's bottom edge by default) as PowerPoint draws them, grown to fit their text:
        a :class:`~pptx_agent.RowsFit` with the ``count``, each row's drawn ``heights``,
        the table's ``top`` and how far it runs ``past``.  To paginate a long table, keep
        ``count`` rows here and move the rest to a table on the next slide.  Needs
        pptx2svg.

        For example::

            rows = table.rows_fitting()
            rows.count, rows.past                   # (6, 4114800): rows 7 on are cut off
        """
        from .fit import rows_fitting

        return rows_fitting(self._shape(), bottom)

    def set_column_width(self, column: int, width: int) -> "Table":
        """Set a grid column's width, EMU; the frame grows or shrinks with it.

        For example::

            table.set_column_width(0, 1828800)
        """
        if int(width) < 0:
            raise ValueError("a column width cannot be negative")
        with self._editing() as table:
            node = self._grid_columns(table)[self._column_index(column, table)]
            delta = int(width) - (get_int(node, "w", 0) or 0)
            set_int(node, "w", int(width))
            self._grow_frame(dx=delta)
        return self

    def set_row_height(self, row: int, height: int) -> "Table":
        """Set a row's height.  PowerPoint treats it as a minimum: text may still grow it.

        For example::

            table.set_row_height(0, 457200)
        """
        if int(height) < 0:
            raise ValueError("a row height cannot be negative")
        with self._editing() as table:
            node = self._rows(table)[self._row_index(row, table)]
            delta = int(height) - (get_int(node, "h", 0) or 0)
            set_int(node, "h", int(height))
            self._grow_frame(dy=delta)
        return self

    # -- merges -----------------------------------------------------------------------------

    @property
    def merged_regions(self) -> list[Region]:
        """Every merged area (spanning more than one cell), in reading order.

        For example::

            table.merged_regions
        """
        return [region for region in self._regions() if region.height > 1 or region.width > 1]

    def region_at(self, row: int, column: int) -> Region:
        """The region covering a grid position -- a 1x1 region for an unmerged cell.

        For example::

            table.region_at(0, 2)
        """
        for region in self._regions():
            if region.contains(row, column):
                return region
        return Region(row, column, row, column)

    def merge(self, top: int, left: int, bottom: int, right: int) -> "TableCell":
        """Merge the rectangle ``(top, left)``-``(bottom, right)``, inclusive.

        Text from the covered cells is appended to the origin's, as PowerPoint does.  A merge
        that would cut through an existing merged region is refused rather than guessed at.

        For example::

            origin = table.merge(0, 1, 0, 2)              # a header across two columns
        """
        with self._editing() as table:
            rows, cols = len(self._rows(table)), len(self._grid_columns(table))
            if not (0 <= top <= bottom < rows and 0 <= left <= right < cols):
                raise IndexError(f"{self.address}: ({top},{left})-({bottom},{right}) is outside "
                                 f"the {rows}x{cols} table")
            if top == bottom and left == right:
                return TableCell(self, top, left)
            target = Region(top, left, bottom, right)
            for region in self._regions(table):
                inside = target.contains(region.top, region.left) and target.contains(
                    region.bottom, region.right)
                overlaps = not (region.bottom < top or region.top > bottom
                                or region.right < left or region.left > right)
                if overlaps and not inside:
                    raise ValueError(f"{self.address}: the merge would cut through the merged "
                                     f"region {region}")
            origin = self._tc(top, left, table)
            for row in range(top, bottom + 1):
                for column in range(left, right + 1):
                    tc = self._tc(row, column, table)
                    if tc is not origin:
                        _move_text(tc, origin)
            _write_region(self._rows(table), target)
        return TableCell(self, top, left)

    def split(self, row: int, column: int) -> "Table":
        """Undo the merge covering ``(row, column)``; its cells become independent again.

        For example::

            table.split(0, 1)
        """
        with self._editing() as table:
            region = self.region_at(row, column)
            rows = self._rows(table)
            for r in range(region.top, region.bottom + 1):
                for c in range(region.left, region.right + 1):
                    _clear_spans(_cells(rows[r])[c])
        return self

    # -- structure --------------------------------------------------------------------------

    def insert_row(self, index: int | None = None, *, like: int | None = None,
                   height: int | None = None) -> "Table":
        """Insert an empty row before row ``index`` (default: at the bottom).

        Cell formatting -- fill, borders, margins, paragraph and run properties -- is copied
        from row ``like``, by default the row above (or below, when inserting at the top).
        A merged region the new row falls *inside* grows to include it.

        For example::

            table.insert_row(like=5)                      # a last row formatted like row 5
        """
        with self._editing() as table:
            rows = self._rows(table)
            count = len(rows)
            position = count if index is None else index
            if not 0 <= position <= count:
                raise IndexError(f"{self.address}: cannot insert a row at {index}")
            template_index = like if like is not None else (position - 1 if position > 0 else 0)
            template = rows[self._row_index(template_index, table)]
            regions = self._regions(table)

            new = copy.deepcopy(template)
            _refresh_ids(new, table, "a16:rowId")
            for tc in _cells(new):
                _blank_cell(tc)
            if height is not None:
                set_int(new, "h", int(height))
            if position < count:
                rows[position].addprevious(new)
            else:
                rows[-1].addnext(new)

            rows = self._rows(table)
            for region in regions:
                if region.top < position <= region.bottom:
                    _write_region(rows, Region(region.top, region.left, region.bottom + 1,
                                               region.right))
            self._grow_frame(dy=get_int(new, "h", 0) or 0)
        return self

    def delete_row(self, index: int) -> "Table":
        """Remove row ``index``, shrinking any merged region it passes through.

        For example::

            table.delete_row(2)
        """
        with self._editing() as table:
            rows = self._rows(table)
            index = self._row_index(index, table)
            if len(rows) == 1:
                raise ValueError(f"{self.address}: a table must keep at least one row")
            regions = self._regions(table)
            for region in regions:
                if region.height == 1 or not region.top <= index <= region.bottom:
                    continue
                if index == region.top:
                    # The origin is going: its text and properties move down one row.
                    _hand_over(_cells(rows[region.top])[region.left],
                               _cells(rows[region.top + 1])[region.left])
                shrunk = Region(region.top, region.left, region.bottom - 1, region.right)
                for c in range(region.left, region.right + 1):
                    for r in range(region.top, region.bottom + 1):
                        _clear_spans(_cells(rows[r])[c])
                remaining = [r for r in range(region.top, region.bottom + 1) if r != index]
                _write_region([rows[r] for r in remaining],
                              Region(0, region.left, shrunk.height - 1, region.right))
            height = get_int(rows[index], "h", 0) or 0
            remove(rows[index])
            self._grow_frame(dy=-height)
        return self

    def insert_column(self, index: int | None = None, *, like: int | None = None,
                      width: int | None = None) -> "Table":
        """Insert an empty column before column ``index`` (default: at the right).

        Formatting is copied from column ``like`` (default: the column to the left, or to
        the right when inserting first).  A merged region the new column falls *inside*
        grows to include it.

        For example::

            table.insert_column(2, like=1)
        """
        with self._editing() as table:
            grid = self._grid_columns(table)
            count = len(grid)
            position = count if index is None else index
            if not 0 <= position <= count:
                raise IndexError(f"{self.address}: cannot insert a column at {index}")
            template_index = self._column_index(
                like if like is not None else (position - 1 if position > 0 else 0), table)
            regions = self._regions(table)

            new_col = copy.deepcopy(grid[template_index])
            _refresh_ids(new_col, table, "a16:colId")
            if width is not None:
                set_int(new_col, "w", int(width))
            if position < count:
                grid[position].addprevious(new_col)
            else:
                grid[-1].addnext(new_col)

            for row in self._rows(table):
                cells = _cells(row)
                new_tc = copy.deepcopy(cells[template_index])
                _blank_cell(new_tc)
                if position < count:
                    cells[position].addprevious(new_tc)
                else:
                    cells[-1].addnext(new_tc)

            rows = self._rows(table)
            for region in regions:
                if region.left < position <= region.right:
                    _write_region(rows, Region(region.top, region.left, region.bottom,
                                               region.right + 1))
            self._grow_frame(dx=get_int(new_col, "w", 0) or 0)
        return self

    def delete_column(self, index: int) -> "Table":
        """Remove column ``index``, shrinking any merged region it passes through.

        For example::

            table.delete_column(4)
        """
        with self._editing() as table:
            grid = self._grid_columns(table)
            index = self._column_index(index, table)
            if len(grid) == 1:
                raise ValueError(f"{self.address}: a table must keep at least one column")
            rows = self._rows(table)
            for region in self._regions(table):
                if region.width == 1 or not region.left <= index <= region.right:
                    continue
                if index == region.left:
                    _hand_over(_cells(rows[region.top])[region.left],
                               _cells(rows[region.top])[region.left + 1])
                for r in range(region.top, region.bottom + 1):
                    for c in range(region.left, region.right + 1):
                        _clear_spans(_cells(rows[r])[c])
                # Rewrite the region over the surviving columns, then drop the doomed one.
                survivors = [c for c in range(region.left, region.right + 1) if c != index]
                _write_region_columns(rows, region.top, region.bottom, survivors)
            width = get_int(grid[index], "w", 0) or 0
            for row in rows:
                remove(_cells(row)[index])
            remove(grid[index])
            self._grow_frame(dx=-width)
        return self

    # -- internals --------------------------------------------------------------------------

    def _shape(self) -> "Shape":
        return self._resolve()

    def _table(self) -> Element:
        table = _table_element(self._shape()._element)
        if table is None:
            raise ValueError(f"{self.address} no longer holds a table")
        return table

    def _rows(self, table: Element | None = None) -> list[Element]:
        return list((table if table is not None else self._table()).findall(qn("a:tr")))

    def _grid_columns(self, table: Element | None = None) -> list[Element]:
        grid = (table if table is not None else self._table()).find(qn("a:tblGrid"))
        return [] if grid is None else list(grid.findall(qn("a:gridCol")))

    def _row_index(self, row: int, table: Element | None = None) -> int:
        count = len(self._rows(table))
        if not 0 <= row < count:
            raise IndexError(f"{self.address}: no row {row} in a {count}-row table")
        return row

    def _column_index(self, column: int, table: Element | None = None) -> int:
        count = len(self._grid_columns(table))
        if not 0 <= column < count:
            raise IndexError(f"{self.address}: no column {column} in a {count}-column table")
        return column

    def _tc(self, row: int, column: int, table: Element | None = None) -> Element:
        rows = self._rows(table)
        self._row_index(row, table)
        cells = _cells(rows[row])
        if not 0 <= column < len(cells):
            raise IndexError(f"{self.address}: no cell ({row},{column})")
        return cells[column]

    def _regions(self, table: Element | None = None) -> list[Region]:
        regions = []
        for r, row in enumerate(self._rows(table)):
            for c, tc in enumerate(_cells(row)):
                if _is_set(tc, "hMerge") or _is_set(tc, "vMerge"):
                    continue
                height = max(get_int(tc, "rowSpan", 1) or 1, 1)
                width = max(get_int(tc, "gridSpan", 1) or 1, 1)
                regions.append(Region(r, c, r + height - 1, c + width - 1))
        return regions

    def _grow_frame(self, *, dx: int = 0, dy: int = 0) -> None:
        if not dx and not dy:
            return
        extent = find(self._shape()._element, "p:xfrm/a:ext")
        if extent is None:
            return
        if dx:
            set_int(extent, "cx", max((get_int(extent, "cx", 0) or 0) + dx, 0))
        if dy:
            set_int(extent, "cy", max((get_int(extent, "cy", 0) or 0) + dy, 0))

    def _editing(self):
        return _TableEdit(self)

    def __repr__(self) -> str:
        return f"<Table {self.address} {self.rows}x{self.columns}>"


class _TableEdit:
    def __init__(self, table: Table) -> None:
        self._table = table
        self._batch = None

    def __enter__(self) -> Element:
        shape = self._table._shape()
        element = self._table._table()
        self._batch = shape._batch()
        self._batch.__enter__()
        shape._before_change()
        return element

    def __exit__(self, *exc):
        if exc[0] is None:
            self._table._shape()._after_change()
        return self._batch.__exit__(*exc)


class TableCell:
    """One grid position, addressed as ``<shape id>/cell<row>,<column>`` (0-based)::

        cell = deck.shape("257.3#5").table.cell(1, 1)     # or table.cell_by_label(...)
        cell.text = "4,310億円"; cell.fill = "accent1 lumMod=20% lumOff=80%"

    Its :attr:`text` reads and writes; :attr:`fill`, :meth:`border` and
    :meth:`set_border` style it.  A position hidden by a merge is still a real ``a:tc`` --
    it just is not drawn.  Its text and fill are invisible; use :attr:`origin` to reach the
    cell that is.
    """

    def __init__(self, table: Table, row: int, column: int) -> None:
        self._table = table
        self.row = row
        self.column = column

    @property
    def address(self) -> str:
        """``<shape id>/cell<row>,<column>``, as :meth:`Document.resolve` takes it.

        For example::

            cell.address                                 # '257.3#5/cell1,1'
        """
        return f"{self._table.address}/cell{self.row},{self.column}"

    # -- merge state -----------------------------------------------------------------------

    @property
    def is_merge_origin(self) -> bool:
        """``True`` for the top-left cell of a merge, which holds its text.

        For example::

            table.cell(0, 1).is_merge_origin
        """
        region = self._table.region_at(self.row, self.column)
        return (region.height > 1 or region.width > 1) and (region.top, region.left) == (
            self.row, self.column)

    @property
    def is_spanned(self) -> bool:
        """``True`` when another cell's merge hides this one.

        For example::

            table.cell(0, 2).is_spanned                  # True under a merge
        """
        region = self._table.region_at(self.row, self.column)
        return (region.top, region.left) != (self.row, self.column)

    @property
    def span(self) -> tuple[int, int]:
        """``(rows, columns)`` of the region this cell belongs to.

        For example::

            cell.span                                    # (1, 2)
        """
        region = self._table.region_at(self.row, self.column)
        return region.height, region.width

    @property
    def origin(self) -> "TableCell":
        """The cell that draws this position: itself, or its merge's top-left cell.

        For example::

            table.cell(0, 2).origin.text = "FY2025"
        """
        region = self._table.region_at(self.row, self.column)
        return TableCell(self._table, region.top, region.left)

    # -- text ------------------------------------------------------------------------------

    @property
    def text_frame(self) -> TextFrame:
        """The cell's paragraphs and runs, as a shape's.

        For example::

            cell.text_frame.paragraph(0).alignment = "right"
        """
        table, row, column = self._table, self.row, self.column
        return TextFrame(lambda: TableCell(table, row, column), self.address)

    @property
    def text(self) -> str:
        r"""The cell's text -- **read and write**.  Reading gives it raw (paragraphs joined
        by ``"\n"``, a line break as ``"\v"``); assigning replaces it the way
        :meth:`Shape.set_text <pptx_agent.Shape.set_text>` does, keeping the formatting of
        every character that survives, so a bold figure stays bold::

            cell = table.cell_by_label("売上高", "当期実績")
            cell.text = "4,310億円"            # same font, size, colour and weight

        On a position a merge covers, the text is invisible: write to :attr:`origin`.
        """
        return self.text_frame.text

    @text.setter
    def text(self, value: str) -> None:
        self.text_frame.set_text(value)

    def _list_styles(self) -> list:
        """A cell is not a placeholder: its own list style, the master's ``otherStyle``,
        the presentation's default."""
        from . import inherit

        shape = self._table._shape()
        package = shape._slide.document.package
        master = inherit.master_of(package, inherit.layout_of(package, shape._slide.part_path))
        own = self._tc().find(f"{qn('a:txBody')}/{qn('a:lstStyle')}")
        return inherit.list_styles(package, master, own, None, None, False)

    def _body_chain(self) -> list:
        return [self._tc().find(f"{qn('a:txBody')}/{qn('a:bodyPr')}")]

    def _theme_source(self):
        return self._table._shape()

    def set_text(self, value: str) -> "TableCell":
        """Set :attr:`text`, the same way; returns the cell.

        For example::

            table.cell_by_label("売上高", "当期実績").set_text("4,310億円").fill = "accent1"
        """
        self.text_frame.set_text(value)
        return self

    @property
    def paragraphs(self):
        """The cell's paragraphs.

        For example::

            [p.text for p in cell.paragraphs]
        """
        return self.text_frame.paragraphs

    # -- fill and borders ------------------------------------------------------------------

    @property
    def fill(self) -> Fill | None:
        """The cell's explicit fill, or ``None`` when the table style decides. Settable.

        For example::

            cell.fill = "accent1 lumMod=20% lumOff=80%"
        """
        return read_fill(self._tc().find(qn("a:tcPr")))

    @fill.setter
    def fill(self, value: "str | Color | None") -> None:
        self._write_fill(fill_element(value))

    def set_gradient_fill(self, stops, *, angle: float | None = 90.0,
                          path: str | None = None) -> "TableCell":
        """A gradient fill, as :meth:`Shape.set_gradient_fill`.

        For example::

            cell.set_gradient_fill([(0, "accent1"), (1, "bg1")])
        """
        self._write_fill(gradient_fill(stops, angle=angle, path=path))
        return self

    def border(self, side: str) -> LineFormat:
        """``"left"``, ``"right"``, ``"top"``, ``"bottom"``, ``"diagonal_down"`` or
        ``"diagonal_up"``.

        Adjacent cells each carry their own copy of a shared edge, and PowerPoint draws the
        later one over the earlier; set both (or use :meth:`set_border`) for a predictable
        result.

        For example::

            cell.border("bottom").width = 12700
        """
        try:
            tag = BORDER_TAGS[side]
        except KeyError:
            raise ValueError(f"side must be one of {sorted(BORDER_TAGS)}") from None
        table, row, column = self._table, self.row, self.column

        def locate(create: bool) -> Element | None:
            tc = table._tc(row, column)
            properties = subelement(tc, "a:tcPr") if create else tc.find(qn("a:tcPr"))
            if properties is None:
                return None
            if create:
                line = properties.find(qn(tag))
                if line is None:
                    line = subelement(properties, tag)
                return line
            return properties.find(qn(tag))

        return LineFormat(
            locate,
            lambda: table._shape()._before_change(),
            lambda: table._shape()._after_change(),
        )

    def set_border(self, side: str, *, width: int | None = None,
                   color: "str | Color | None" = None, dash: str | None = None,
                   shared: bool = True) -> "TableCell":
        """Style one edge.  With ``shared`` the neighbour's matching edge is styled too, so
        the edge looks the same whichever copy PowerPoint draws.

        For example::

            cell.set_border("bottom", width=12700, color="accent1")
        """
        with self._table._shape()._batch():
            targets = [(self, side)]
            if shared:
                neighbour = self._neighbour(side)
                if neighbour is not None:
                    targets.append(neighbour)
            for cell, edge in targets:
                line = cell.border(edge)
                if width is not None:
                    line.width = width
                if color is not None:
                    line.color = color
                if dash is not None:
                    line.dash = dash
        return self

    # -- host protocol ---------------------------------------------------------------------

    def _text_body(self, create: bool) -> Element | None:
        tc = self._tc()
        body = tc.find(qn("a:txBody"))
        if body is None and create:
            body = subelement(tc, "a:txBody")
            subelement(body, "a:bodyPr")
            subelement(body, "a:lstStyle")
            body.append(make("a:p"))
        return body

    def _before_change(self) -> None:
        self._table._shape()._before_change()

    @property
    def _slide(self):
        """The slide the table is on -- where a run's hyperlink relationship belongs."""
        return self._table._shape()._slide

    def _after_change(self) -> None:
        self._table._shape()._after_change()

    def _batch(self):
        return self._table._shape()._batch()

    # -- internals -------------------------------------------------------------------------

    def _tc(self) -> Element:
        return self._table._tc(self.row, self.column)

    def _write_fill(self, element: Element | None) -> None:
        tc = self._tc()
        if element is None and tc.find(qn("a:tcPr")) is None:
            return
        with self._batch():
            self._before_change()
            write_fill(subelement(self._tc(), "a:tcPr"), element, FILL_TAGS)
            self._after_change()

    def _neighbour(self, side: str) -> tuple["TableCell", str] | None:
        offsets = {"left": (0, -1, "right"), "right": (0, 1, "left"),
                   "top": (-1, 0, "bottom"), "bottom": (1, 0, "top")}
        if side not in offsets:
            return None
        dr, dc, opposite = offsets[side]
        row, column = self.row + dr, self.column + dc
        if 0 <= row < self._table.rows and 0 <= column < self._table.columns:
            return TableCell(self._table, row, column), opposite
        return None

    def __repr__(self) -> str:
        return f"<TableCell {self.address} {self.text[:30]!r}>"


# -- XML helpers ---------------------------------------------------------------------------


def _table_element(frame: Element) -> Element | None:
    return find(frame, "a:graphic/a:graphicData/a:tbl")


def _cells(row: Element) -> list[Element]:
    return list(row.findall(qn("a:tc")))


def _is_set(tc: Element, attribute: str) -> bool:
    return (tc.get(attribute) or "").strip().lower() in {"1", "true", "on"}


def _clear_spans(tc: Element) -> None:
    for attribute in _SPAN_ATTRIBUTES:
        set_attr(tc, attribute, None)


def _write_region(rows: list[Element], region: Region) -> None:
    """Spell ``region`` onto the grid the way PowerPoint does (see the module docstring)."""
    for r in range(region.top, region.bottom + 1):
        cells = _cells(rows[r])
        for c in range(region.left, region.right + 1):
            tc = cells[c]
            _clear_spans(tc)
            if r == region.top and region.height > 1:
                tc.set("rowSpan", str(region.height))
            if c == region.left and region.width > 1:
                tc.set("gridSpan", str(region.width))
            if c > region.left:
                tc.set("hMerge", "1")
            if r > region.top:
                tc.set("vMerge", "1")


def _write_region_columns(rows: list[Element], top: int, bottom: int, columns: list[int]) -> None:
    """Write a region over a non-contiguous set of columns (one is about to be deleted)."""
    height, width = bottom - top + 1, len(columns)
    for r in range(top, bottom + 1):
        cells = _cells(rows[r])
        for position, c in enumerate(columns):
            tc = cells[c]
            _clear_spans(tc)
            if r == top and height > 1:
                tc.set("rowSpan", str(height))
            if position == 0 and width > 1:
                tc.set("gridSpan", str(width))
            if position > 0:
                tc.set("hMerge", "1")
            if r > top:
                tc.set("vMerge", "1")


def _blank_cell(tc: Element) -> None:
    """Empty a copied cell: no spans, no text, formatting kept."""
    _clear_spans(tc)
    body = tc.find(qn("a:txBody"))
    if body is None:
        return
    paragraphs = body.findall(qn("a:p"))
    keep = paragraphs[0] if paragraphs else None
    for extra in paragraphs[1:]:
        remove(extra)
    if keep is None:
        body.append(make("a:p"))
        return
    run_properties = None
    for child in list(keep):
        if child.tag in {qn("a:r"), qn("a:fld")} and run_properties is None:
            run_properties = child.find(qn("a:rPr"))
        if child.tag not in {qn("a:pPr"), qn("a:endParaRPr")}:
            keep.remove(child)
    if keep.find(qn("a:endParaRPr")) is None and run_properties is not None:
        end = copy.deepcopy(run_properties)
        end.tag = qn("a:endParaRPr")
        keep.append(end)


def _move_text(source: Element, origin: Element) -> None:
    """Append ``source``'s non-empty paragraphs to ``origin`` and leave ``source`` blank."""
    source_body = source.find(qn("a:txBody"))
    if source_body is None:
        return
    paragraphs = [p for p in source_body.findall(qn("a:p"))
                  if "".join(t.text or "" for t in p.iter(qn("a:t"))).strip()]
    if paragraphs:
        origin_body = origin.find(qn("a:txBody"))
        if origin_body is None:
            origin_body = subelement(origin, "a:txBody")
            subelement(origin_body, "a:bodyPr")
            subelement(origin_body, "a:lstStyle")
        existing = origin_body.findall(qn("a:p"))
        # Replace an empty origin rather than leaving a blank first line.
        if len(existing) == 1 and not "".join(t.text or "" for t in existing[0].iter(qn("a:t"))):
            remove(existing[0])
        for paragraph in paragraphs:
            origin_body.append(copy.deepcopy(paragraph))
    _blank_cell(source)


def _hand_over(origin: Element, successor: Element) -> None:
    """The origin cell is being deleted: give its content to the cell taking its place."""
    for tag in ("a:txBody", "a:tcPr"):
        old = successor.find(qn(tag))
        if old is not None:
            remove(old)
        mine = origin.find(qn(tag))
        if mine is not None:
            append_in_order(successor, copy.deepcopy(mine))


def _refresh_ids(element: Element, table: Element, tag: str) -> None:
    """Give a copied row or column a fresh PowerPoint ``a16:rowId``/``a16:colId``."""
    try:
        name = qn(tag)
    except KeyError:
        return
    used = []
    for node in table.iter(name):
        try:
            used.append(int(node.get("val", "")))
        except ValueError:
            continue
    fresh = max(used, default=10000) + 1
    for node in element.iter(name):
        node.set("val", str(fresh))
        fresh += 1
