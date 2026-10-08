"""P8: Gantt-style engagement plan built from shapes; bar geometry checked against the month headers."""
import calendar
import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402
from design import (IN, box, common_checks, fill_key, find_exact, leaf_shapes, ntext, run_sizes,  # noqa: E402
                    text_shapes)

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p8-engagement-plan/input/halden-proposal-draft.pptx"
MONTHS = [("Nov 2026", 2026, 11), ("Dec 2026", 2026, 12), ("Jan 2027", 2027, 1), ("Feb 2027", 2027, 2),
          ("Mar 2027", 2027, 3)]
PHASES = [("Diagnose", date(2026, 11, 2), date(2026, 11, 20)), ("Design", date(2026, 11, 23), date(2026, 12, 18)),
          ("Negotiate", date(2027, 1, 4), date(2027, 2, 12)), ("Embed", date(2027, 2, 15), date(2027, 3, 5))]
WORKSTREAMS = ["Spend & baseline", "Category strategy", "Negotiations", "Operating model", "Change & capability"]
BARS = [
    ("Spend & baseline", "Spend cube", date(2026, 11, 2), date(2026, 11, 20)),
    ("Spend & baseline", "Savings tracking", date(2027, 1, 4), date(2027, 3, 5)),
    ("Category strategy", "Category deep dives", date(2026, 11, 16), date(2026, 12, 18)),
    ("Category strategy", "Wave-2 preparation", date(2027, 1, 25), date(2027, 2, 26)),
    ("Negotiations", "Wave-1 RFQs", date(2027, 1, 4), date(2027, 1, 22)),
    ("Negotiations", "Supplier negotiations", date(2027, 1, 25), date(2027, 2, 12)),
    ("Operating model", "Target model design", date(2026, 11, 30), date(2026, 12, 18)),
    ("Operating model", "Transition and hand-over", date(2027, 2, 15), date(2027, 3, 5)),
    ("Change & capability", "Category manager coaching", date(2027, 1, 11), date(2027, 2, 26)),
]
MILESTONES = [("Kick-off", date(2026, 11, 2)), ("SteerCo 1: savings hypothesis", date(2026, 11, 20)),
              ("SteerCo 2: wave plan approved", date(2026, 12, 18)),
              ("SteerCo 3: first contracts signed", date(2027, 2, 12)), ("Final SteerCo", date(2027, 3, 5))]
BOARD = date(2027, 1, 22)
FOOT = ("Note: programme weeks exclude the holiday break, 21 Dec 2026 to 1 Jan 2027. "
        "Dates as agreed with the Halden PMO on 9 October 2026.")
BAR_PRESETS = ("rect", "roundRect", "flowChartProcess", "flowChartAlternateProcess", "round1Rect", "round2SameRect",
               "snip1Rect", "snip2SameRect")
MS_PRESETS = ("diamond", "flowChartDecision", "triangle", "flowChartMerge", "flowChartExtract")
TOL = int(0.1 * IN)
T = int(0.02 * IN)


def spread(values):
    return max(values) - min(values) if values else 0


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    slide = common_checks(c, deck, src, 3, (1, 2), "Engagement plan")
    shapes = text_shapes(slide)
    leaves = leaf_shapes(slide)
    every = " ".join(ntext(sh.text) for sh in shapes)

    # ---- month scale
    heads = []
    for label, _, _ in MONTHS:
        cand = find_exact(shapes, label)
        heads.append(cand[0] if len(cand) == 1 else None)
    c.check("five month headers, one shape each, exact text", all(heads),
            [len(find_exact(shapes, m[0])) for m in MONTHS])
    if not all(heads):
        return
    hb = [box(h) for h in heads]
    c.check("month headers in one row", spread([b[1] for b in hb]) <= T, [b[1] for b in hb])
    gaps = [hb[i + 1][0] - hb[i][2] for i in range(4)]
    c.check("month headers left to right, edge to edge (0.03 in)", all(abs(g) <= 0.03 * IN for g in gaps),
            [round(g / IN, 3) for g in gaps])

    def x_at(d: date, part: str) -> float:
        k = next(i for i, (_, y, m) in enumerate(MONTHS) if (y, m) == (d.year, d.month))
        left, right = hb[k][0], hb[k][2]
        days = calendar.monthrange(d.year, d.month)[1]
        frac = {"start": d.day - 1, "end": d.day, "mid": d.day - 0.5}[part] / days
        return left + frac * (right - left)

    # ---- bars
    bars = {}
    for ws, name, a, b in BARS:
        cand = [sh for sh in find_exact(shapes, name) if sh.preset in BAR_PRESETS]
        bars[name] = cand[0] if len(cand) == 1 else None
    c.check("every activity is one rect/roundRect bar with its name as text", all(bars.values()),
            {n: (len(find_exact(shapes, n)), [s.preset for s in find_exact(shapes, n)]) for n in bars
             if not bars[n]})
    errs = {}
    for ws, name, a, b in BARS:
        sh = bars.get(name)
        if sh is None:
            continue
        bx = box(sh)
        dl, dr = bx[0] - x_at(a, "start"), bx[2] - x_at(b, "end")
        if abs(dl) > TOL or abs(dr) > TOL:
            errs[name] = (round(dl / IN, 3), round(dr / IN, 3))
    c.check("bar left/right edges on their dates (0.1 in)", not errs and all(bars.values()), errs)
    have = [s for s in bars.values() if s]
    if have:
        hs = [box(s)[3] - box(s)[1] for s in have]
        c.check("all bars the same height", spread(hs) <= T, hs)
        rows = []
        for ws in WORKSTREAMS:
            mids = [(box(bars[n])[1] + box(bars[n])[3]) / 2 for w, n, _, _ in BARS if w == ws and bars.get(n)]
            rows.append(sum(mids) / len(mids) if mids else None)
            c.check(f"row '{ws}': bars share one row", spread(mids) <= T, mids)
        if None not in rows:
            steps = [rows[i + 1] - rows[i] for i in range(len(rows) - 1)]
            c.check("rows top-to-bottom in workstream order, evenly spaced", min(steps) > 0 and spread(steps) <= 0.03 * IN,
                    [round(s / IN, 3) for s in steps])
            plot_left = min(box(s)[0] for s in have)
            lab = {}
            for ws, mid in zip(WORKSTREAMS, rows):
                cand = find_exact(shapes, ws)
                cand = [s for s in cand if box(s)[2] <= plot_left + T]
                ok = any(abs((box(s)[1] + box(s)[3]) / 2 - mid) <= TOL for s in cand)
                lab[ws] = ok
            c.check("workstream labels left of the plot, centred on their rows", all(lab.values()), lab)
            small = [(n, z) for n, s in bars.items() if s for z in run_sizes(s) if z < 9]
            small += [(s.text, z) for ws in WORKSTREAMS for s in find_exact(shapes, ws) for z in run_sizes(s) if z < 9]
            c.check("bar texts and workstream labels >= 9 pt", not small, small)
        # colour: one per workstream or one per phase
        by_ws = {}
        by_ph = {}
        for ws, name, a, b in BARS:
            if bars.get(name):
                by_ws.setdefault(ws, set()).add(fill_key(bars[name]))
                ph = next(p for p, s, e in PHASES if s <= a <= e)
                by_ph.setdefault(ph, set()).add(fill_key(bars[name]))
        per_ws = all(len(v) == 1 for v in by_ws.values()) and len(set().union(*by_ws.values())) > 1
        per_ph = all(len(v) == 1 for v in by_ph.values()) and len(set().union(*by_ph.values())) > 1
        c.check("one colour per workstream or per phase", per_ws or per_ph,
                {"by_workstream": {k: sorted(v) for k, v in by_ws.items()}})

    # ---- phase bands
    top_bars = min(box(s)[1] for s in have) if have else None
    perr = {}
    pboxes = []
    for name, a, b in PHASES:
        cand = [s for s in find_exact(shapes, name) if top_bars is None or box(s)[3] <= top_bars + T]
        best = None
        for s in cand:
            bx = box(s)
            e = max(abs(bx[0] - x_at(a, "start")), abs(bx[2] - x_at(b, "end")))
            if best is None or e < best[0]:
                best = (e, s)
        if best is None or best[0] > TOL:
            perr[name] = None if best is None else round(best[0] / IN, 3)
        else:
            pboxes.append(box(best[1]))
    c.check("phase bands span their dates (0.1 in), above the bars", not perr, perr)
    if len(pboxes) == 4:
        c.check("phase bands in one row", spread([b[1] for b in pboxes]) <= T, [b[1] for b in pboxes])

    # ---- milestones
    markers = [s for s in leaves if s.kind == "shape" and s.preset in MS_PRESETS]
    used, merr = set(), {}
    for label, d in MILESTONES:
        x = x_at(d, "mid")
        near = sorted((abs((box(s)[0] + box(s)[2]) / 2 - x), s.id) for s in markers if s.id not in used)
        if near and near[0][0] <= TOL:
            used.add(near[0][1])
        else:
            merr[label] = round(near[0][0] / IN, 3) if near else None
    c.check("five milestone markers centred on their dates (0.1 in)", not merr, {"markers": len(markers), **merr})
    missing = [m for m, _ in MILESTONES if ntext(m) not in every]
    c.check("milestone labels present exactly", not missing, missing)

    # ---- board update line
    x = x_at(BOARD, "mid")
    lines = []
    for s in leaves:
        bx = box(s)
        if (s.is_connector or s.preset in ("line", "straightConnector1")) and bx[2] - bx[0] <= T:
            lines.append((abs(bx[0] - x), s, bx))
    lines.sort(key=lambda t: t[0])
    ok = bool(lines) and lines[0][0] <= TOL
    c.check("vertical line at 22 Jan 2027 (0.1 in)", ok, [(round(d / IN, 3), s.id) for d, s, _ in lines[:3]])
    if ok and have and rows and None not in rows:
        _, s, bx = lines[0]
        c.check("board line spans first to last workstream row", bx[1] <= rows[0] and bx[3] >= rows[-1],
                (bx[1], bx[3], rows[0], rows[-1]))
        dash = str(s.line.dash) if s.line is not None else None
        c.check("board line dashed", dash not in (None, "None", "solid"), dash)
    c.check("'Board update' label present", "board update" in every.lower())
    absent = [t for t in WORKSTREAMS + [n for _, n, _, _ in BARS] + [p for p, _, _ in PHASES] if ntext(t) not in every]
    c.check("all workstream/bar/phase text present", not absent, absent)
    c.check("footnote present exactly", ntext(FOOT) in every)


if __name__ == "__main__":
    main("p8-engagement-plan", body)
