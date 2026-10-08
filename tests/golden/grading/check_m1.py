"""M1: 2x2 prioritisation matrix: axes, quadrants, ten bubbles at their scores, legend."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main  # noqa: E402
from design import IN, fill_key, is_title, leaf_shapes, ntext, run_sizes, text_shapes  # noqa: E402
from design2 import centre, common_checks2, dbox, every_text, gap, sbox, spread  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "m1-priority-matrix/input/halden-proposal-draft.pptx"
X_TITLE = "Ease of implementation (0 = hard, 10 = easy)"
Y_TITLE = "Savings potential (0 = low, 10 = high)"
QUADRANTS = {"Quick wins": (1, 1), "Strategic bets": (0, 1), "Fill-ins": (1, 0), "Deprioritise": (0, 0)}
ITEMS = [
    ("Cement & binders", 7.5, 8.5, "Direct materials"),
    ("Steel reinforcement", 3.5, 9.0, "Direct materials"),
    ("Packaging", 8.0, 6.5, "Direct materials"),
    ("Energy contracts", 2.5, 7.0, "Indirect"),
    ("MRO spares", 8.5, 3.5, "Indirect"),
    ("IT & telecoms", 6.0, 2.0, "Indirect"),
    ("Professional services", 9.0, 5.5, "Indirect"),
    ("Inbound freight", 6.5, 7.5, "Logistics"),
    ("Warehousing", 2.0, 3.0, "Logistics"),
    ("Fleet leasing", 4.0, 1.5, "Logistics"),
]
TYPES = ["Direct materials", "Indirect", "Logistics"]
SOURCE = "Source: Meridian Partners diagnostic, scores agreed with category owners in week 3."
AX = int(0.05 * IN)
POS = int(0.1 * IN)
T = int(0.02 * IN)


def lines(slide):
    out = []
    for sh in leaf_shapes(slide):
        if sh.kind == "connector" or sh.preset == "line":
            b = dbox(sh)
            out.append((sh, b))
    return out


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    slide = common_checks2(c, deck, src, 3, (1, 2), "Prioritisation")
    shapes = [sh for sh in text_shapes(slide) if not is_title(sh)]
    every = every_text(slide)

    ls = lines(slide)
    vert = [(s, b) for s, b in ls if b[2] - b[0] <= T and b[3] - b[1] >= 2 * IN]
    horiz = [(s, b) for s, b in ls if b[3] - b[1] <= T and b[2] - b[0] >= 2 * IN]
    c.check("a vertical and a horizontal axis line (>= 2 in long)", vert and horiz,
            {"vertical": len(vert), "horizontal": len(horiz)})
    if not (vert and horiz):
        return
    yaxis = min(vert, key=lambda sb: sb[1][0])[1]       # leftmost vertical
    xaxis = max(horiz, key=lambda sb: sb[1][3])[1]      # lowest horizontal
    x0, x10 = xaxis[0], xaxis[2]
    y0, y10 = xaxis[1], yaxis[1]
    c.check("axes meet at the origin (0.05 in)",
            abs(yaxis[0] - x0) <= AX and abs(yaxis[3] - y0) <= AX,
            {"y-axis x": yaxis[0], "x-axis left": x0, "y-axis bottom": yaxis[3], "x-axis y": y0})

    def at(e, s):
        return x0 + e / 10 * (x10 - x0), y0 - s / 10 * (y0 - y10)

    ovals = [sh for sh in leaf_shapes(slide) if sh.preset == "ellipse"]
    bubbles, misses = [], []
    for name, e, s, _ in ITEMS:
        px, py = at(e, s)
        best = min(ovals, key=lambda o: (centre(sbox(o))[0] - px) ** 2 + (centre(sbox(o))[1] - py) ** 2,
                   default=None)
        if best is None:
            bubbles.append(None)
            misses.append((name, "no oval"))
            continue
        cx, cy = centre(sbox(best))
        d = ((cx - px) ** 2 + (cy - py) ** 2) ** 0.5
        bubbles.append(best if d <= POS else None)
        if d > POS:
            misses.append((name, round(d / IN, 3)))
    c.check("every initiative an oval centred on its scores (0.1 in)", not misses, misses)
    found = [b for b in bubbles if b]
    c.check("ten distinct bubbles", len({b.id for b in found}) == 10, len({b.id for b in found}))
    if found:
        sizes = [(sbox(b)[2] - sbox(b)[0], sbox(b)[3] - sbox(b)[1]) for b in found]
        c.check("bubbles all the same size, at least 0.2 in across",
                spread([w for w, _ in sizes]) <= T and spread([h for _, h in sizes]) <= T
                and min(min(w, h) for w, h in sizes) >= 0.2 * IN, sizes)

    bad_labels = []
    label_shapes = []
    for (name, *_), b in zip(ITEMS, bubbles):
        if b is None:
            bad_labels.append((name, "no bubble"))
            continue
        bc = centre(sbox(b))
        cand = [sh for sh in shapes if ntext(sh.text) == ntext(name)]
        near = [sh for sh in cand if sh.id == b.id or
                ((centre(sbox(sh))[0] - bc[0]) ** 2 + (centre(sbox(sh))[1] - bc[1]) ** 2) ** 0.5 <= IN]
        if not near:
            bad_labels.append((name, [round(((centre(sbox(sh))[0] - bc[0]) ** 2 +
                                              (centre(sbox(sh))[1] - bc[1]) ** 2) ** 0.5 / IN, 2) for sh in cand]))
        label_shapes += near[:1]
    c.check("each label in its bubble or within 1.0 in of it", not bad_labels, bad_labels)

    x5, y5 = (x0 + x10) / 2, (y0 + y10) / 2
    quad_bad, quad_shapes = [], []
    for label, (qx, qy) in QUADRANTS.items():
        cand = [sh for sh in shapes if ntext(sh.text) == ntext(label)]
        if len(cand) != 1:
            quad_bad.append((label, f"{len(cand)} shapes"))
            continue
        b = dbox(cand[0])
        quad_shapes.append(cand[0])
        xl, xr = (x0, x5) if qx == 0 else (x5, x10)
        yt, yb = (y5, y0) if qy == 0 else (y10, y5)
        if not (b[0] >= xl - T and b[2] <= xr + T and b[1] >= yt - T and b[3] <= yb + T):
            quad_bad.append((label, [round(v / IN, 2) for v in b]))
    c.check("quadrant labels each wholly inside their quadrant", not quad_bad, quad_bad)

    fills = {}
    for (name, _, _, kind), b in zip(ITEMS, bubbles):
        if b is not None:
            fills.setdefault(kind, set()).add(fill_key(b))
    one_each = all(len(v) == 1 for v in fills.values()) and len(fills) == 3
    distinct = one_each and len({next(iter(v)) for v in fills.values()}) == 3
    c.check("one fill per spend type, three different fills", distinct, {k: sorted(v) for k, v in fills.items()})
    if distinct:
        bubble_ids = {b.id for b in found}
        leg_bad = []
        for kind in TYPES:
            f = next(iter(fills[kind]))
            texts = [sh for sh in shapes if ntext(sh.text) == ntext(kind)]
            if not texts:
                leg_bad.append((kind, "no legend text"))
                continue
            ok = False
            for t in texts:
                if fill_key(t) == f and getattr(t.fill, "kind", None) == "solid":
                    ok = True
                for sw in leaf_shapes(slide):
                    if sw.id in bubble_ids or sw.id == t.id or sw.kind not in ("shape",):
                        continue
                    if fill_key(sw) == f and gap(sbox(sw), sbox(t)) <= 0.3 * IN:
                        ok = True
            if not ok:
                leg_bad.append((kind, "no swatch within 0.3 in"))
        c.check("legend: each spend type next to a swatch of its fill", not leg_bad, leg_bad)

    small = [(ntext(s.text)[:20], z) for s in label_shapes + quad_shapes for z in run_sizes(s) if z < 9]
    c.check("bubble and quadrant labels >= 9 pt", not small, small)
    texts = [X_TITLE, Y_TITLE, SOURCE] + list(QUADRANTS) + [n for n, *_ in ITEMS] + TYPES
    missing = [t for t in texts if ntext(t) not in every]
    c.check("all text present exactly", not missing, missing)


if __name__ == "__main__":
    main("m1-priority-matrix", body)
