"""P15: a maturity gap chart (columns in tx2 with one point highlighted, targets as diamond
markers, a 0-5 axis by 1, no gridlines, 10 pt text) and an effort-against-impact bubble
matrix labelled R1-R6, from two CSV files."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

FOLDER = Path(os.environ.get("GOLDEN_INPUTS", str(INPUTS))) / "p15-maturity-charts" / "input"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
EMU = 12700


def rows(name: str) -> list[list[str]]:
    text = (FOLDER / name).read_text(encoding="utf-8")
    return [line.split(",") for line in text.split("\n") if line]


def takeaway(c, slide, label: str, words) -> None:
    title = next((sh for sh in slide.shapes if sh.placeholder
                  and sh.placeholder[0] in ("title", "ctrTitle")), None)
    text = title.text if title is not None else ""
    fit = title.text_fit() if title is not None else None
    c.check(f"{label}: a one-line title that fits", bool(norm(text)) and "\n" not in text
            and fit is not None and tuple(fit.lines) == (1,) and not fit.overflows, (text, fit))
    c.check(f"{label}: a takeaway, not a label", any(w.lower() in text.lower() for w in words)
            and norm(text).lower() not in {"maturity gap", "maturity scores", "effort vs impact",
                                           "effort versus impact", "recommendations"}, text)
    return title


def frame_of(c, deck, slide, label: str, number: int, title):
    frames = [s for s in slide.shapes if s.kind == "graphic_frame" and s.chart is not None]
    c.check(f"{label}: one chart", len(frames) == 1, len(frames))
    if not frames:
        return None
    frame = frames[0]
    left, top, width, height = slide.content_area
    page_w, page_h = deck.slide_size
    c.check(f"{label}: below the title, on the slide, filling the content area",
            frame.top >= title.top + title.height - EMU and frame.top + frame.height <= page_h
            and frame.left + frame.width <= page_w and frame.left >= 0
            and frame.width * frame.height >= 0.6 * width * height,
            (frame.left, frame.top, frame.width, frame.height))
    c.check(f"{label}: nothing overlaps or leaves the slide", not deck.overflows(slides=[number]),
            deck.overflows(slides=[number]))
    return frame


def body(c, out: Path):
    scores = rows("maturity-scores.csv")[1:]
    capabilities = [r[0] for r in scores]
    current = [float(r[1]) for r in scores]
    target = [float(r[2]) for r in scores]
    recs = rows("recommendations.csv")[1:]
    deck = Document.open(out)
    src = Document.open(FOLDER / "maturity-assessment.pptx")
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    titles = [title_of(s) for s in deck.slides]
    c.check("6 slides, the new ones 4th and 5th", len(titles) == 6 and titles[2] == "Approach"
            and titles[5] == "Next steps", titles)
    if len(titles) != 6:
        return

    # -- the maturity gap ------------------------------------------------------------------
    slide = deck.slides[3]
    title = takeaway(c, slide, "gap", ["target", "gap", "below", "procure", "behind", "short"])
    frame = frame_of(c, deck, slide, "gap", 4, title)
    if frame is not None:
        chart = frame.chart
        root = chart._root()
        c.check("gap: columns and a line plot on the same axes", chart.chart_types == ["bar", "line"]
                and [n.get("val") for n in root.find(f".//{C}barChart").findall(f"{C}axId")]
                == [n.get("val") for n in root.find(f".//{C}lineChart").findall(f"{C}axId")],
                chart.chart_types)
        c.check("gap: capabilities in order", list(chart.categories) == capabilities, chart.categories)
        got = {s.name: [float(v) for v in s.values] for s in chart.series}
        c.check("gap: CSV values exactly", got == {"Current": current, "Target": target}, got)
        book = chart.workbook_values()
        c.check("gap: Edit Data holds the drawn values", book["categories"] == capabilities
                and {s["name"]: s["values"] for s in book["series"]} == got, book)
        scale = chart.axis_scale("value")
        c.check("gap: value axis 0-5 by 1, no gridlines", (scale["minimum"], scale["maximum"],
                scale["major_unit"]) == (0, 5, 1) and not scale["major_gridlines"]
                and not scale["minor_gridlines"], scale)
        look = chart.series_format("Current")
        lowest = current.index(min(current))
        point = look.get("points", {}).get(lowest, {})
        c.check("gap: current in tx2, the lowest highlighted in an accent", look.get("fill") == "tx2"
                and str(point.get("fill", "")).startswith("accent")
                and set(look.get("points", {})) == {lowest}, look)
        marks = chart.series_format("Target")
        c.check("gap: targets as diamond markers with no line", marks.get("marker") == "diamond"
                and marks.get("line") is None and "line" in marks, marks)
        sizes = {n.get("sz") for n in root.iter(f"{A}defRPr", f"{A}rPr") if n.get("sz")}
        c.check("gap: all text 10 pt", sizes == {"1000"}, sizes)

    # -- effort against impact ---------------------------------------------------------------
    slide = deck.slides[4]
    title = takeaway(c, slide, "matrix", ["quick", "win", "R1", "R4", "low effort", "first",
                                          "impact"])
    frame = frame_of(c, deck, slide, "matrix", 5, title)
    if frame is not None:
        chart = frame.chart
        root = chart._root()
        c.check("matrix: a bubble chart", chart.chart_types == ["bubble"], chart.chart_types)
        ser = root.find(f".//{C}bubbleChart/{C}ser")

        def numbers(tag):
            return [float(p.find(f"{C}v").text) for p in ser.find(tag).iter(f"{C}pt")]
        c.check("matrix: effort, impact and priority exactly",
                numbers(f"{C}xVal") == [float(r[2]) for r in recs]
                and numbers(f"{C}yVal") == [float(r[3]) for r in recs]
                and numbers(f"{C}bubbleSize") == [float(r[4]) for r in recs],
                (numbers(f"{C}xVal"), numbers(f"{C}yVal")))
        labels = chart.point_labels(0)
        c.check("matrix: labelled R1-R6", labels == [r[0] for r in recs], labels)
        place = {n.get("val") for n in ser.iter(f"{C}dLblPos")}
        c.check("matrix: labels beside the bubbles", place and place <= {"r", "l"}, place)
        scales = chart.scales
        c.check("matrix: both axes 0-5", all((scales[a]["minimum"], scales[a]["maximum"]) == (0, 5)
                                             for a in ("x", "y")), scales)
        c.check("matrix: axes titled Effort and Impact", (chart.axis_title("category"),
                chart.axis_title("value")) == ("Effort", "Impact"),
                (chart.axis_title("category"), chart.axis_title("value")))

    same = all(deck.to_outline(slides=[i]).split("\n", 1)[1] == src.to_outline(slides=[j]).split("\n", 1)[1]
               for i, j in ((1, 1), (2, 2), (3, 3), (6, 4)))
    c.check("other slides unchanged", same)


if __name__ == "__main__":
    main("p15-maturity-charts", body)
