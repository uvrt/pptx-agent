"""P2: restate Q2 revenue / operating profit / net profit throughout the deck (answer key inside)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p2-quarterly-update/input/harbor-lane-q2-results.pptx"

TEXT = {  # shape id -> expected text after the restatement
    "256.13": "2,165", "256.15": "▲ 10.0%",
    "256.17": "262", "256.19": "▲ 18.0%",
    "256.21": "184", "256.23": "▲ 16.5%",
    "256.29": "12.1%", "256.30": "+0.8pt YoY",
    "257.12": "49.2%", "257.16": "50.4%", "257.20": "50.4%",
    "258.16": ("Road freight: cross-border volumes up 9% on the Rotterdam to Milan corridor\n"
               "Warehousing: two new e-commerce sites opened, in Tilburg and Lyon\n"
               "Air & sea: ocean forwarding margins recovered as rates stabilised\n"
               "Other & adjustments: margin of 12.0%, up from 9.2% a year ago"),
}
UNCHANGED_TEXT = ["256.33", "256.34", "256.37", "256.38", "256.41", "256.42", "257.13", "257.17", "257.21",
                  "256.14", "256.18", "256.22"]
PL = [
    ["Item", "Q2 actual", "Prior year", "Change", "Change %"],
    ["Revenue", "EUR 2,165m", "EUR 1,968m", "+197m", "+10.0%"],
    ["Gross profit", "EUR 642m", "EUR 600m", "+42m", "+7.0%"],
    ["Operating profit", "EUR 262m", "EUR 222m", "+40m", "+18.0%"],
    ["Profit before tax", "EUR 248m", "EUR 215m", "+33m", "+15.3%"],
    ["Net profit", "EUR 184m", "EUR 158m", "+26m", "+16.5%"],
]
SEG = [
    ["Segment", "Revenue", "Share", "Op. profit", "Margin", "Headcount"],
    ["Road freight", "EUR 980m", "45.3%", "EUR 108m", "11.0%", "5,420"],
    ["Warehousing", "EUR 620m", "28.6%", "EUR 82m", "13.2%", "3,180"],
    ["Air & sea", "EUR 390m", "18.0%", "EUR 51m", "13.1%", "960"],
    ["Other & adjustments", "EUR 175m", "8.1%", "EUR 21m", "12.0%", "410"],
    ["Total", "EUR 2,165m", "100.0%", "EUR 262m", "12.1%", "9,970"],
]
CHARTS = {
    "257.25": {"Revenue (EUR m)": [2010, 2075, 2165], "Op. profit (EUR m)": [231, 240, 262]},
    "257.29": {"Operating margin (%)": [11.5, 11.6, 12.1], "Gross margin (%)": [29.9, 29.8, 29.7],
               "Net margin (%)": [8.2, 8.3, 8.5]},
    "258.9": {"Prior year": [923, 568, 347, 130], "Current": [980, 620, 390, 175]},
    "258.13": {"Share of revenue": [45.3, 28.6, 18.0, 8.1]},
}


def grid(t):
    return [[norm(t.cell(r, c).text) for c in range(t.columns)] for r in range(t.rows)]


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    before = {str(p) for p in src.validate()}
    after = {str(p) for p in deck.validate()}
    c.check("validate(): no new problems", after <= before, sorted(after - before))
    c.check("still 3 slides", len(deck.slides) == 3, len(deck.slides))
    for sid, want in TEXT.items():
        try:
            got = deck.shape(sid).text
        except Exception as e:  # noqa: BLE001
            got = f"<{e}>"
        c.check(f"text {sid} = {want[:40]!r}", norm(got) == norm(want), got)
    bad = [(sid, deck.shape(sid).text, src.shape(sid).text) for sid in UNCHANGED_TEXT
           if deck.shape(sid).text != src.shape(sid).text]
    c.check("leave-alone figures unchanged (ROE, EPS, dividend, forecasts)", not bad, bad)
    pl = grid(deck.shape("257.3#5").table)
    for r, (g, w) in enumerate(zip(pl, PL)):
        c.check(f"P&L row {w[0]}", g == w, g)
    seg = grid(deck.shape("258.4#3").table)
    for r, (g, w) in enumerate(zip(seg, SEG)):
        c.check(f"segment row {w[0]}", g == w, g)
    for sid, series in CHARTS.items():
        ch = deck.shape(sid).chart
        for name, want in series.items():
            got = [None if v is None else round(float(v), 4) for v in ch.series[name].values]
            c.check(f"chart {sid} {name}", got == [float(x) for x in want], got)
    # every other text shape unchanged
    changed = []
    for s_new, s_old in zip(deck.slides, src.slides):
        for sh in s_old.shapes:
            if sh.id in TEXT or sh.has_table or sh.has_chart:
                continue
            try:
                if deck.shape(sh.id).text != sh.text:
                    changed.append((sh.id, sh.text[:30], deck.shape(sh.id).text[:30]))
            except Exception as e:  # noqa: BLE001
                changed.append((sh.id, "missing", str(e)))
    c.check("no other text changed", not changed, changed[:6])
    # formatting kept on the edited figures: bold stays bold
    fmt = []
    for sid in ("256.13", "256.17", "256.21", "256.29"):
        r_new = deck.shape(sid).text_frame.paragraph(0).runs
        r_old = src.shape(sid).text_frame.paragraph(0).runs
        if [x.bold for x in r_new] != [x.bold for x in r_old][:len(r_new)] or \
           [x.size for x in r_new] != [x.size for x in r_old][:len(r_new)]:
            fmt.append(sid)
    c.check("figures keep their run formatting", not fmt, fmt)
    # workbook kept in step with the cache (Edit Data)
    import zipfile
    from common import xlsx_grid
    bad = []
    with zipfile.ZipFile(out) as z:
        names = z.namelist()
        for sid, series in CHARTS.items():
            ch = deck.shape(sid).chart
            part = getattr(ch, "workbook_part", None)
            if not part or part not in names:
                bad.append((sid, "no workbook part", part))
                continue
            g = xlsx_grid(z.read(part))
            values = sorted(v for v in g.values() if isinstance(v, float))
            for name, want in series.items():
                for v in want:
                    if not any(abs(v - x) < 1e-6 for x in values):
                        bad.append((sid, name, v))
    c.check("embedded workbooks hold the new values", not bad, bad[:6])


if __name__ == "__main__":
    main("p2-quarterly-update", body)
