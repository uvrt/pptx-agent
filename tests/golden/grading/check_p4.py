"""P4: order-to-cash flow chart with attached connectors and theme colours."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, all_shapes, main, norm, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p4-process-diagram/input/order-to-cash.pptx"
STEPS = ["Order received", "Credit check", "Pick and pack", "Ship order", "Send invoice", "Payment received",
         "Notify customer"]
EDGES = [("Order received", "Credit check"), ("Credit check", "Pick and pack"), ("Pick and pack", "Ship order"),
         ("Ship order", "Send invoice"), ("Send invoice", "Payment received"), ("Credit check", "Notify customer")]


def colour(fill):
    if fill is None:
        return None
    col = getattr(fill, "color", None)
    return str(col) if col is not None else str(fill)


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    c.check("still 3 slides", len(deck.slides) == 3, len(deck.slides))
    others = all(deck.to_outline(slides=[i]).split("\n", 1)[1] == src.to_outline(slides=[i]).split("\n", 1)[1]
                 for i in (1, 3))
    c.check("slides 1 and 3 unchanged", others)
    slide = deck.slides[1]
    c.check("slide 2 title kept", title_of(slide) == "Order-to-cash process", title_of(slide))
    shapes = all_shapes(slide)
    by = {}
    for sh in shapes:
        t = norm(sh.text) if sh.kind == "shape" and not sh.is_connector else ""
        if t in STEPS:
            by.setdefault(t, sh)
    missing = [s for s in STEPS if s not in by]
    c.check("a shape for every step, text inside", not missing, missing)
    dec = by.get("Credit check")
    c.check("credit check is a diamond", dec is not None and dec.preset in ("diamond", "flowChartDecision"),
            dec and dec.preset)
    rects = [by[s].preset for s in STEPS if s in by and s != "Credit check"]
    c.check("other steps rectangles/rounded rectangles",
            all(p in ("rect", "roundRect", "flowChartProcess", "flowChartAlternateProcess", "snipRect",
                      "flowChartTerminator") for p in rects), rects)
    conns = [sh for sh in shapes if sh.is_connector]
    ids = {sh.id: n for n, sh in by.items()}
    edges, loose = [], []
    for cn in conns:
        b, e = cn.begin_connection, cn.end_connection
        if not b or not e:
            loose.append(cn.id)
            continue
        bid = b[0].id if hasattr(b[0], "id") else str(b[0])
        eid = e[0].id if hasattr(e[0], "id") else str(e[0])
        edges.append((ids.get(bid, bid), ids.get(eid, eid), cn))
    c.check("every connector attached at both ends", conns and not loose, loose)
    pairs = {(a, b) for a, b, _ in edges}
    miss = [e for e in EDGES if e not in pairs]
    rev = [e for e in miss if (e[1], e[0]) in pairs]
    c.check("connectors join the right steps, in flow direction", not miss, {"missing": miss, "reversed": rev})
    heads = []
    for a, b, cn in edges:
        ln = cn.line
        tail = getattr(ln, "tail", None)
        head = getattr(ln, "head", None)
        heads.append((a, b, str(head), str(tail)))
    arrow_ok = all((t not in ("None", "none") and t is not None) or (h not in ("None", "none"))
                   for a, b, h, t in heads)
    c.check("connectors have arrowheads", heads and arrow_ok, heads)
    texts = [norm(sh.text) for sh in shapes if sh.kind == "shape"]
    c.check("'Pass' and 'Fail' labels present", "Pass" in texts and "Fail" in texts, texts)
    fills = {s: colour(by[s].fill) for s in by}
    rgb = {s: f for s, f in fills.items() if f and "#" in f}
    c.check("step fills are theme colours (no hard-coded RGB)", not rgb, rgb)
    ordinary = {fills.get(s) for s in STEPS if s not in ("Credit check", "Notify customer") and s in fills}
    c.check("ordinary steps share one colour", len(ordinary) == 1 and None not in ordinary, ordinary)
    trio = [fills.get("Order received"), fills.get("Credit check"), fills.get("Notify customer")]
    c.check("decision and notify use different colours", len(set(trio)) == 3 and None not in trio, trio)
    lines = {cn.id: str(cn.line.color) for cn in conns if cn.line.color is not None}
    c.check("connector colours not hard-coded RGB", not [v for v in lines.values() if "#" in v], lines)
    W, H = deck.slide_size
    tshape = next((sh for sh in slide.shapes if sh.placeholder and sh.placeholder[0] == "title"), None)
    title_bottom = (tshape.top + tshape.height) if tshape is not None else 0
    boxes = [(n, sh.left, sh.top, sh.left + sh.width, sh.top + sh.height) for n, sh in by.items()]
    out_of = [b[0] for b in boxes if b[1] < 0 or b[2] < title_bottom - 10000 or b[3] > W or b[4] > H]
    c.check("shapes inside the slide and below the title", not out_of, out_of)
    overlaps = []
    for i in range(len(boxes)):
        for j in range(i + 1, len(boxes)):
            a, b = boxes[i], boxes[j]
            if a[1] < b[3] and b[1] < a[3] and a[2] < b[4] and b[2] < a[4]:
                overlaps.append((a[0], b[0]))
    c.check("no step shapes overlap", not overlaps, overlaps)
    main_flow = ["Order received", "Credit check", "Pick and pack", "Ship order", "Send invoice", "Payment received"]
    if all(s in by for s in main_flow):
        xs = [(by[s].top // 914400 * 0, by[s].left) for s in main_flow]
        rows = {}
        for s in main_flow:
            rows.setdefault(round(by[s].top / 457200), []).append(by[s].left)
        lr = all(v == sorted(v) for v in rows.values())
        c.check("main flow reads left to right (per row)", lr, {k: v for k, v in rows.items()})


if __name__ == "__main__":
    main("p4-process-diagram", body)
