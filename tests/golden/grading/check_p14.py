"""P14: a chart slide on a content layout with no empty placeholder left, and a table in the
template's own table style with its header row on and banding off."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "generated" / "trial"
CSV = FIXTURES / "quarterly-revenue-2026.csv"
INPUT = INPUTS / "p14-brand-chart-table" / "input" / "business-review-2026.pptx"
BRAND = "{6D3B0E52-8C1A-4F7E-B9A4-1E2F3A4B5C6D}"


def body(c, out: Path):
    rows = [line.split(",") for line in CSV.read_text(encoding="utf-8").split("\n") if line]
    quarters = [r[0] for r in rows[1:]]
    want = {name: [float(r[k]) for r in rows[1:]] for k, name in enumerate(rows[0]) if k}
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    titles = [title_of(s) for s in deck.slides]
    c.check("6 slides, the new ones 4th and 5th", len(titles) == 6
            and titles[2] == "Results by region" and titles[5] == "Next steps", titles)
    if len(titles) != 6:
        return
    chart_slide, table_slide = deck.slides[3], deck.slides[4]
    for n, slide in ((4, chart_slide), (5, table_slide)):
        text = title_of(slide) or ""
        title = next((sh for sh in slide.shapes if sh.placeholder
                      and sh.placeholder[0] in ("title", "ctrTitle")), None)
        fit = title.text_fit() if title is not None else None
        c.check(f"slide {n}: a title on one line, as drawn", bool(norm(text)) and "\n" not in text
                and "\v" not in text and fit is not None and tuple(fit.lines) == (1,)
                and not fit.overflows, (text, fit))
    c.check("slide 4 is on Title and Content", chart_slide.layout.name == "Title and Content",
            chart_slide.layout.name)
    frames = [s for s in chart_slide.shapes if s.kind == "graphic_frame" and s.chart is not None]
    c.check("slide 4: one chart", len(frames) == 1, len(frames))
    if frames:
        chart = frames[0].chart
        c.check("quarters as categories, in order", list(chart.categories) == quarters,
                chart.categories)
        got = {s.name: [float(v) for v in s.values] for s in chart.series}
        c.check("one series per region, CSV values exactly", got == want, got)
        area = chart_slide.content_area
        left, top, width, height = frames[0].bounds
        c.check("the chart is in the content area", left >= area[0] - 12700
                and top >= area[1] - 12700 and left + width <= area[0] + area[2] + 12700
                and top + height <= area[1] + area[3] + 12700, (frames[0].bounds, area))
    empty = [s.id for s in chart_slide.empty_placeholders()]
    c.check("slide 4: no empty placeholder (no 'Click to add text')", not empty, empty)
    tables = [s for s in table_slide.shapes if s.kind == "graphic_frame" and s.table is not None]
    c.check("slide 5: one table", len(tables) == 1, len(tables))
    if tables:
        table = tables[0].table
        cells = [[table.cell(r, k).text for k in range(table.columns)] for r in range(table.rows)]
        expected = [["Region", "Revenue 2026 (EUR m)"]] + [
            [name, f"{sum(values):.1f}"] for name, values in want.items()]
        c.check("header and one row per region, totals to one decimal", cells == expected, cells)
        style = table.style or {}
        c.check("the template's own table style", style.get("id") == BRAND
                and not style.get("builtin"), style)
        flags = table.flags
        c.check("header row on, banded rows off", flags["first_row"] and not flags["banded_rows"],
                flags)
    empty = [s.id for s in table_slide.empty_placeholders()]
    c.check("slide 5: no empty placeholder", not empty, empty)
    problems = deck.overflows()
    c.check("nothing overflows, overlaps or leaves the slide", not problems,
            [str(p) for p in problems][:5])
    before = [src.to_outline(slides=[n], ids=False) for n in (1, 2, 3, 4)]
    after = [deck.to_outline(slides=[n], ids=False) for n in (1, 2, 3, 6)]
    c.check("the other slides unchanged", after == before, after)


if __name__ == "__main__":
    main("p14-brand-chart-table", body)
