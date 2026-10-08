"""P7: phased-approach proposal slide built from shapes (chevrons, activity columns, deliverables)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402
from design import (IN, box, common_checks, fill_key, has_all, ntext, run_sizes,  # noqa: E402
                    text_shapes)

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p7-phased-approach/input/halden-proposal-draft.pptx"
PHASES = [("Diagnose", "Weeks 1–3"), ("Design", "Weeks 4–7"), ("Negotiate", "Weeks 8–13"), ("Embed", "Weeks 14–16")]
ACTIVITIES = [
    ["Build the group spend cube from ERP data for all six plants", "Interview 25 category owners and plant managers",
     "Benchmark prices for the top 15 categories", "Size the savings potential by category"],
    ["Prioritise eight wave-1 categories", "Write category strategies and negotiation plans",
     "Design the target procurement operating model"],
    ["Run RFQs and supplier negotiations for wave 1", "Track savings in a weekly war room",
     "Prepare the wave-2 categories", "Coach category managers on the job"],
    ["Hand over to the new procurement organisation", "Set up procurement KPIs and governance",
     "Agree the wave-2 roadmap"],
]
DELIVERABLES = [
    ["Spend cube and cost baseline", "Savings hypothesis by category"],
    ["Eight category strategies", "Target operating model", "Wave plan approved by the SteerCo"],
    ["Signed supplier agreements", "Savings tracker validated by Finance"],
    ["Savings sign-off", "Procurement KPI dashboard", "12-month roadmap"],
]
SOURCE = ("Source: Halden Group ERP spend data FY2025; Meridian Partners procurement benchmarks. "
          "Timeline indicative, subject to data access in week 1.")
CHEVRONS = ("chevron", "homePlate")
T = int(0.02 * IN)


def spread(values):
    return max(values) - min(values) if values else 0


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)

    def chevron_pair(a, b):
        return a.preset in CHEVRONS and b.preset in CHEVRONS

    slide = common_checks(c, deck, src, 2, (1, 3), "Our approach", extra_ok=chevron_pair)
    shapes = text_shapes(slide)
    every = " ".join(ntext(sh.text) for sh in shapes)

    chev = sorted([sh for sh in shapes if sh.preset in CHEVRONS] +
                  [sh for sh in slide.shapes if sh.kind == "shape" and sh.preset in CHEVRONS and not norm(sh.text)],
                  key=lambda s: box(s)[0])
    chev = list({s.id: s for s in chev}.values())
    chev.sort(key=lambda s: box(s)[0])
    c.check("exactly 4 chevron shapes (chevron/homePlate)", len(chev) == 4, [(s.id, s.preset) for s in chev])
    order = []
    for sh in chev:
        t = ntext(sh.text)
        order.append(next((i for i, (n, w) in enumerate(PHASES) if ntext(n) in t and ntext(w) in t), None))
    c.check("chevron texts: phase name + weeks, in phase order", order == [0, 1, 2, 3],
            [ntext(s.text) for s in chev])
    if len(chev) == 4:
        b = [box(s) for s in chev]
        c.check("chevrons in one row (same top)", spread([x[1] for x in b]) <= T, [x[1] for x in b])
        c.check("chevrons same size", spread([x[2] - x[0] for x in b]) <= T and spread([x[3] - x[1] for x in b]) <= T,
                [(x[2] - x[0], x[3] - x[1]) for x in b])
        steps = [b[i + 1][0] - b[i][0] for i in range(3)]
        c.check("chevrons evenly spaced", spread(steps) <= T, steps)
        c.check("chevron fills are theme colours or inherited", True, [fill_key(s) for s in chev])

    cols, dels = [], []
    for k in range(4):
        cand = [sh for sh in shapes if has_all(sh, ACTIVITIES[k])]
        cols.append(cand[0] if len(cand) == 1 else None)
        dcand = [sh for sh in shapes if has_all(sh, DELIVERABLES[k])]
        dels.append(dcand[0] if len(dcand) == 1 else None)
    c.check("each phase's activities in exactly one text shape", all(cols),
            [len([sh for sh in shapes if has_all(sh, ACTIVITIES[k])]) for k in range(4)])
    c.check("each phase's deliverables in exactly one shape", all(dels),
            [len([sh for sh in shapes if has_all(sh, DELIVERABLES[k])]) for k in range(4)])
    c.check("activity and deliverable shapes are separate", all(a is not d for a, d in zip(cols, dels) if a and d))
    if all(cols):
        cb = [box(s) for s in cols]
        c.check("activity columns equal width", spread([x[2] - x[0] for x in cb]) <= T, [x[2] - x[0] for x in cb])
        gutters = [cb[i + 1][0] - cb[i][2] for i in range(3)]
        c.check("activity columns equal gutters (left to right)", spread(gutters) <= T and min(gutters) >= 0, gutters)
        if len(chev) == 4:
            off = [abs((x[0] + x[2]) / 2 - (y[0] + y[2]) / 2) for x, y in zip(cb, [box(s) for s in chev])]
            c.check("each column centred under its chevron (0.1 in)", max(off) <= 0.1 * IN,
                    [round(o / IN, 3) for o in off])
            below = all(x[1] >= box(chev[0])[3] - T for x in cb)
            c.check("columns below the chevrons", below)
        small = [(norm(s.text)[:20], z) for s in cols for z in run_sizes(s) if z < 10]
        c.check("activity text >= 10 pt", not small, small)
    if all(dels):
        db = [box(s) for s in dels]
        filled = [getattr(s.fill, "kind", None) if s.fill is not None else "inherited" for s in dels]
        c.check("deliverable boxes are filled", all(f not in (None, "none") for f in filled)
                and all(s.kind == "shape" for s in dels), filled)
        c.check("deliverable boxes in one row, same top and height",
                spread([x[1] for x in db]) <= T and spread([x[3] - x[1] for x in db]) <= T,
                [(x[1], x[3] - x[1]) for x in db])
        if all(cols):
            cb = [box(s) for s in cols]
            c.check("deliverable box same width as its column (0.05 in)",
                    max(abs((d[2] - d[0]) - (k[2] - k[0])) for d, k in zip(db, cb)) <= 0.05 * IN,
                    [((d[2] - d[0]) / IN, (k[2] - k[0]) / IN) for d, k in zip(db, cb)])
            c.check("deliverable box under its column", all(d[1] >= k[1] for d, k in zip(db, cb)))
    c.check("'Deliverables' label present", "deliverables" in every.lower())
    missing = [t for t in sum(ACTIVITIES, []) + sum(DELIVERABLES, []) + [n for n, _ in PHASES] + [w for _, w in PHASES]
               if ntext(t) not in every]
    c.check("all activity/deliverable/phase text present exactly", not missing, missing)
    c.check("source line present exactly", ntext(SOURCE) in every)


if __name__ == "__main__":
    main("p7-phased-approach", body)
