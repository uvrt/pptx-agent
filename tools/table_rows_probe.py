"""How tall PowerPoint draws a table's rows: one table per slide, its rows grown by text.

A row's stored height (``a:tr@h``) is a minimum; PowerPoint grows a row whose text needs
more.  ``Table.drawn_row_heights`` predicts the drawn height with pptx2svg's table layout
(``pptx2svg.table_row_heights``) and the text measurement ``text_fit`` uses.  This probe
puts tables that exercise it -- font sizes from 10 to 40 pt, cell margins from none to half
an inch, stored heights taller and shorter than the text, paragraphs, empty rows, Aptos,
Arial, Calibri and Times New Roman, and a 12-row table that runs off the slide -- on one
slide each, and reads the rows PowerPoint drew back from its PDF:

    python tools/table_rows_probe.py build ~/table-rows-probe.pptx cases.json
    (export it to PDF with PowerPoint: tests/oracle.py's export_pdf)
    python tools/table_rows_probe.py read ~/table-rows-probe.pdf cases.json

``read`` takes each row's top and bottom from the fill PowerPoint paints in the table's
first column (the default style fills every cell) and prints them against the prediction.
The measured rows are in ``tests/fixtures/table-rows-probe.json``; tests/test_fit.py holds
the prediction to them.  Needs pypdfium2 to read.
"""

from __future__ import annotations

import json
import sys

EMU_PER_PT = 12700
INCH = 914400

LONG = "A longer cell text that wraps onto two or three lines in this column width"
WORDS = ("Rows grow to fit their text, and every row below moves down with them; the stored "
         "height is only a minimum")


def _cases():
    """(name, column widths, stored row heights, rows of (text, formatting, margins))."""
    out = []
    # 1. The production case: twelve rows of three wrapping 18 pt cells, past the slide.
    out.append(("wrap12", [3657600] * 3, [365760] * 12,
                [[(f"R{r}C{c}: {LONG}", {}, None) for c in (1, 2, 3)] for r in range(1, 13)]))
    # 2. Font sizes, one per row: a short cell beside a wrapping one, and an empty row.
    sizes = (10, 14, 20, 28, 40)
    out.append(("sizes", [1828800, 5486400], [228600] * 6,
                [[(f"{size} pt", {"size": size}, None), (WORDS, {"size": size}, None)]
                 for size in sizes] + [[("", {}, None), ("", {}, None)]]))
    # 3. Cell margins, (left, top, right, bottom) EMU, one set per row, 16 pt.
    margins = [(0, 0, 0, 0), None, (91440, 182880, 91440, 91440), (457200, 457200, 457200, 0),
               (91440, 91440, 91440, 365760)]
    out.append(("margins", [2743200, 4572000], [274320] * 5,
                [[("Margins", {"size": 16}, m), (WORDS, {"size": 16}, m)] for m in margins]))
    # 4. Mixed stored heights: tall rows with little text keep their height, short rows with
    # more text grow; paragraphs and an empty row.
    out.append(("mixed", [2286000, 2286000, 2286000],
                [1097280, 274320, 731520, 182880, 548640, 274320],
                [[("Tall row, short text", {}, None), ("x", {}, None), ("", {}, None)],
                 [(LONG, {"size": 14}, None), ("Two", {}, None), ("One", {}, None)],
                 [("One\nTwo\nThree\nFour", {"size": 14}, None), ("Short", {}, None),
                  ("", {}, None)],
                 [("", {}, None), ("", {}, None), ("", {}, None)],
                 [(WORDS, {"size": 12, "typeface": "Arial"}, None), ("Arial 12", {}, None),
                  ("", {}, None)],
                 [("Last", {"size": 24}, None), (LONG, {"size": 11}, None), ("", {}, None)]]))
    # 5. Faces: the same text in Arial, Calibri and Times New Roman, at two sizes.
    out.append(("faces", [2743200, 2743200, 2743200], [274320] * 2,
                [[(WORDS, {"size": size, "typeface": face}, None)
                  for face in ("Arial", "Calibri", "Times New Roman")] for size in (14, 20)]))
    return out


def build(deck_path: str, cases_path: str) -> None:
    from pptx_agent import Document

    deck = Document.new(size="16:9")
    cases = []
    for name, widths, heights, rows in _cases():
        slide = deck.add_slide("Blank")
        frame = slide.add_table(len(rows), len(widths), 457200, 457200, sum(widths),
                                sum(heights))
        table = frame.table
        for index, width in enumerate(widths):
            table.set_column_width(index, width)
        for index, height in enumerate(heights):
            table.set_row_height(index, height)
        for r, row in enumerate(rows):
            for c, (text, formatting, margins) in enumerate(row):
                cell = table.cell(r, c)
                if text:
                    cell.text = text
                    if formatting:
                        cell.text_frame.format(**formatting)
                if margins is not None:
                    left, top, right, bottom = margins
                    cell.text_frame.set_insets(left=left, top=top, right=right, bottom=bottom)
        cases.append({"page": len(deck.slides), "name": name, "shape": frame.id,
                      "top": frame.top, "left": frame.left, "first_column": widths[0],
                      "stored": table.row_heights, "predicted": table.drawn_row_heights})
    deck.save(deck_path)
    with open(cases_path, "w") as handle:
        json.dump(cases, handle, indent=1)
    print(f"{len(cases)} slides")


def drawn_rows(page, left: float, width: float) -> list[tuple[float, float]]:
    """``(top, bottom)`` in points from the page's top of each fill PowerPoint painted in a
    table's first column (``left``, ``width`` in points)."""
    import pypdfium2.raw as raw

    height = page.get_height()
    rows = set()
    for item in page.get_objects():
        if item.type != raw.FPDF_PAGEOBJ_PATH:
            continue
        x0, y0, x1, y1 = item.get_bounds()
        if abs(x0 - left) < 1.0 and abs((x1 - x0) - width) < 1.0 and y1 - y0 > 1.0:
            rows.add((round(height - y1, 3), round(height - y0, 3)))
    return sorted(rows)


def read(pdf_path: str, cases_path: str) -> None:
    import pypdfium2

    with open(cases_path) as handle:
        cases = json.load(handle)
    document = pypdfium2.PdfDocument(pdf_path)
    for case in cases:
        page = document[case["page"] - 1]
        rows = drawn_rows(page, case["left"] / EMU_PER_PT, case["first_column"] / EMU_PER_PT)
        case["drawn"] = [round((bottom - top) * EMU_PER_PT) for top, bottom in rows]
        case["drawn_top"] = round(rows[0][0] * EMU_PER_PT) if rows else None
        print(case["name"], f"page {page.get_height():.0f} pt tall")
        for index, stored in enumerate(case["stored"]):
            predicted = case["predicted"][index]
            drawn = case["drawn"][index] if index < len(case["drawn"]) else None
            off = "" if drawn is None else f"{(predicted - drawn) / EMU_PER_PT:+7.2f} pt"
            print(f"  row {index + 1:2}  stored {stored / EMU_PER_PT:6.1f}  predicted "
                  f"{predicted / EMU_PER_PT:6.1f}  drawn "
                  f"{'-' if drawn is None else f'{drawn / EMU_PER_PT:6.1f}'}  {off}")
    with open(cases_path, "w") as handle:
        json.dump(cases, handle, indent=1)


if __name__ == "__main__":
    {"build": build, "read": read}[sys.argv[1]](*sys.argv[2:])
