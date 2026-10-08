"""P13: a radar chart comparing three shortlisted vendors on six criteria from a CSV, with a
legend, theme colours and a one-line takeaway title."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import main, norm, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

FIXTURES = Path(__file__).resolve().parents[2] / "fixtures" / "generated" / "trial"
INPUT = FIXTURES / "wms-vendor-selection.pptx"
CSV = FIXTURES / "wms-vendor-scores.csv"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU = 12700


def body(c, out: Path):
    rows = [line.split(",") for line in CSV.read_text(encoding="utf-8").split("\n") if line]
    criteria = [r[0] for r in rows[1:]]
    want = {name: [float(r[k]) for r in rows[1:]] for k, name in enumerate(rows[0]) if k}
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    titles = [title_of(s) for s in deck.slides]
    c.check("5 slides, the new one 4th", len(titles) == 5 and titles[2] == "Shortlist"
            and titles[4] == "Recommendation", titles)
    if len(titles) != 5:
        return
    slide = deck.slides[3]
    title = next((sh for sh in slide.shapes if sh.placeholder
                  and sh.placeholder[0] in ("title", "ctrTitle")), None)
    text = title.text if title is not None else ""
    c.check("a title on one line", bool(norm(text)) and "\n" not in text and "\v" not in text, text)
    fit = title.text_fit() if title is not None else None
    c.check("the title fits on one line as drawn", fit is not None and tuple(fit.lines) == (1,)
            and not fit.overflows, fit)
    c.check("a takeaway that names a vendor, not a label", any(v in text for v in want) and
            norm(text).lower() not in {"vendor scores", "vendor comparison", "evaluation scores"},
            text)
    frames = [s for s in slide.shapes if s.kind == "graphic_frame" and s.chart is not None]
    c.check("one chart", len(frames) == 1, len(frames))
    if not frames:
        return
    frame = frames[0]
    chart = frame.chart
    root = chart._root()
    radar = root.find(f"{C}chart/{C}plotArea/{C}radarChart")
    c.check("a radar chart", radar is not None and chart.chart_types == ["radar"],
            chart.chart_types)
    if radar is None:
        return
    c.check("criteria as categories, in order", list(chart.categories) == criteria,
            chart.categories)
    got = {s.name: [float(v) for v in s.values] for s in chart.series}
    c.check("one series per vendor, CSV scores exactly", got == want, got)
    book = chart.workbook_values()
    c.check("Edit Data holds the drawn values", book is not None and book["categories"] == criteria
            and {s["name"]: s["values"] for s in book["series"]} == want, book)
    colours = []
    for ser in radar.findall(f"{C}ser"):
        fill = ser.find(f"{C}spPr/{A}ln/{A}solidFill")
        colours.append(None if fill is None else fill[0].tag.split("}")[1])
    c.check("theme colours (schemeClr) on every vendor's line", colours
            and all(f == "schemeClr" for f in colours), colours)
    c.check("a legend", chart.has_legend, chart.has_legend)
    left, top, width, height = slide.content_area
    page_w, page_h = deck.slide_size
    box = (frame.left, frame.top, frame.left + frame.width, frame.top + frame.height)
    c.check("the chart is below the title and on the slide",
            box[1] >= title.top + title.height - EMU and box[3] <= page_h and box[2] <= page_w
            and box[0] >= 0, box)
    c.check("the chart fills most of the content area",
            frame.width * frame.height >= 0.6 * width * height,
            (frame.width, frame.height, width, height))
    c.check("nothing overlaps or leaves the slide", not deck.overflows(slides=[4]),
            deck.overflows(slides=[4]))
    same = all(deck.to_outline(slides=[i]).split("\n", 1)[1]
               == src.to_outline(slides=[j]).split("\n", 1)[1]
               for i, j in ((1, 1), (2, 2), (3, 3), (5, 4)))
    c.check("other slides unchanged", same)


if __name__ == "__main__":
    main("p13-vendor-radar", body)
