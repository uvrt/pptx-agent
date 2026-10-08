"""O1: programme governance org chart built from shapes and glued connectors."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main  # noqa: E402
from design import IN, is_title, leaf_shapes, ntext, run_sizes, text_shapes  # noqa: E402
from design2 import centre, common_checks2, every_text, sbox, spread  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "o1-governance-chart/input/halden-proposal-draft.pptx"
UNITS = {
    "sc": ("Steering Committee", "Group CEO (chair), CFO, COO and Meridian partner; meets monthly"),
    "pl": ("Programme Lead", "Ingrid Solberg, Group Procurement Director"),
    "pmo": ("PMO", "Savings tracking, risks and reporting; meets weekly"),
}
WORKSTREAMS = [
    ("Spend & Baseline", "Lead: Erik Dahl, Finance", ["Spend data analyst", "Plant controllers (6)"]),
    ("Category Strategy", "Lead: Maria Berg, Procurement",
     ["Category managers (8)", "Technical specialists", "Meridian category experts"]),
    ("Negotiations", "Lead: Jonas Lind, Procurement", ["Negotiation team", "Legal counsel"]),
    ("Operating Model & Change", "Lead: Sofie Nyberg, HR",
     ["HR business partner", "Change manager", "Training coordinator"]),
]
FOOTNOTE = "Note: names as proposed by Halden; roles to be confirmed at SteerCo 1."
T = int(0.02 * IN)
COL = int(0.05 * IN)


def holding(shapes, name, detail):
    """Shapes whose text holds the name and the detail (the name as a whole line or prefix)."""
    out = []
    for sh in shapes:
        body = ntext(sh.text)
        if ntext(detail) in body and ntext(name) in body.replace(ntext(detail), ""):
            out.append(sh)
    return out


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    slide = common_checks2(c, deck, src, 2, (1, 3), "Programme governance")
    shapes = [sh for sh in text_shapes(slide) if not is_title(sh)]
    every = every_text(slide)

    found = {}
    for key, (name, detail) in UNITS.items():
        cand = holding(shapes, name, detail)
        found[key] = cand[0] if len(cand) == 1 else None
    c.check("Steering Committee, Programme Lead, PMO: one shape each with name + detail",
            all(found.values()), {k: len(holding(shapes, *UNITS[k])) for k in UNITS})
    ws = []
    for name, lead, _ in WORKSTREAMS:
        cand = holding(shapes, name, lead)
        ws.append(cand[0] if len(cand) == 1 else None)
    c.check("each workstream: one shape with name + lead line", all(ws),
            [len(holding(shapes, n, l)) for n, l, _ in WORKSTREAMS])
    roles = []
    for _, _, rs in WORKSTREAMS:
        row = []
        for r in rs:
            cand = [sh for sh in shapes if ntext(sh.text) == ntext(r)]
            row.append(cand[0] if len(cand) == 1 else None)
        roles.append(row)
    c.check("each role: its own shape holding exactly the role text", all(all(r) for r in roles),
            [[bool(x) for x in row] for row in roles])

    sc, pl, pmo = found["sc"], found["pl"], found["pmo"]
    if sc and pl:
        a, b = sbox(sc), sbox(pl)
        cx = centre(b)[0]
        c.check("Steering Committee above Programme Lead, PL centred within SC's width",
                a[3] <= b[1] + T and a[0] - T <= cx <= a[2] + T, (a, b))
    if pl and pmo:
        a, b = sbox(pl), sbox(pmo)
        beside = b[2] <= a[0] + T or b[0] >= a[2] - T
        cy = centre(b)[1]
        c.check("PMO beside the Programme Lead (left or right, no overlap, centre within PL height)",
                beside and a[1] - T <= cy <= a[3] + T, (a, b))
    if all(ws):
        wb = [sbox(s) for s in ws]
        c.check("workstreams in one row (same top)", spread([x[1] for x in wb]) <= T, [x[1] for x in wb])
        c.check("workstreams same width and height",
                spread([x[2] - x[0] for x in wb]) <= T and spread([x[3] - x[1] for x in wb]) <= T,
                [(x[2] - x[0], x[3] - x[1]) for x in wb])
        c.check("workstreams left to right in the given order", all(wb[i][2] <= wb[i + 1][0] + T for i in range(3)),
                [x[0] for x in wb])
        gaps = [wb[i + 1][0] - wb[i][2] for i in range(3)]
        c.check("workstreams equal gaps", spread(gaps) <= T and min(gaps) >= 0, gaps)
        if pl:
            c.check("workstreams below the Programme Lead", min(x[1] for x in wb) >= sbox(pl)[3] - T)
        if pmo:
            c.check("PMO above the workstream row", sbox(pmo)[3] <= min(x[1] for x in wb) + T)
        if all(all(r) for r in roles):
            bad = []
            for k, row in enumerate(roles):
                col = wb[k]
                prev = col[3] - T
                for r in row:
                    b = sbox(r)
                    ok = b[1] >= prev and b[0] >= col[0] - COL and b[2] <= col[2] + COL
                    if not ok:
                        bad.append((WORKSTREAMS[k][0], ntext(r.text), [round(v / IN, 2) for v in b],
                                    [round(v / IN, 2) for v in col]))
                    prev = b[3] - T
            c.check("roles below their workstream, within its column, in order", not bad, bad)

    # connectors glued at both ends
    pairs = set()
    unglued = 0
    for sh in leaf_shapes(slide):
        if sh.kind != "connector":
            continue
        a, b = sh.begin_connection, sh.end_connection
        if a and b:
            pairs.add(frozenset((a[0].id, b[0].id)))
        else:
            unglued += 1
    need = []
    if sc and pl:
        need.append(("SC-PL", sc, pl))
    if pl and pmo:
        need.append(("PL-PMO", pl, pmo))
    for k, w in enumerate(ws):
        if pl and w:
            need.append((f"PL-WS{k + 1}", pl, w))
    missing = [n for n, a, b in need if frozenset((a.id, b.id)) not in pairs]
    c.check("glued connectors: SC-PL, PL-PMO, PL-each workstream", need and len(need) == 6 and not missing,
            {"missing": missing, "glued": len(pairs), "unglued connectors": unglued})

    texts = [t for pair in UNITS.values() for t in pair] + \
        [t for n, l, rs in WORKSTREAMS for t in [n, l] + rs] + [FOOTNOTE]
    missing = [t for t in texts if ntext(t) not in every]
    c.check("all text present exactly", not missing, missing)
    small_units = [(ntext(s.text)[:20], z) for s in list(found.values()) + ws if s for z in run_sizes(s) if z < 10]
    c.check("unit and workstream text >= 10 pt", not small_units, small_units)
    small_roles = [(ntext(s.text)[:20], z) for row in roles for s in row if s for z in run_sizes(s) if z < 9]
    c.check("role text >= 9 pt", not small_roles, small_roles)


if __name__ == "__main__":
    main("o1-governance-chart", body)
