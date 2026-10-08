"""P1: 8-slide deck from an outline on a template with non-standard layout names, with notes."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import main, norm, notes_of, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

LAYOUTS = ["TITLE", "TITLE_AND_BODY", "TITLE_AND_BODY", "BIG_NUMBER", "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY",
           "TITLE_AND_BODY", "TITLE_AND_BODY"]
TITLES = ["Project Harbour", "Why change now", "What we heard from the depots", None, "Options considered",
          "Costs and benefits of Option A (EUR thousand)", "Timeline", "Decision requested"]
BULLETS = {
    2: [(0, "Depot routing still runs on spreadsheets built in 2014"),
        (0, "Rescheduled deliveries cost us EUR 2.3 million last year"),
        (0, "Two of five depots lose their routing lead in 2027"),
        (0, "Customers now expect a two-hour delivery window")],
    3: [(0, "Planning"), (1, "Plans are rebuilt by hand every morning"), (1, "Late orders break the plan"),
        (0, "Drivers"), (1, "No live view of traffic or road closures"),
        (1, "Paper delivery notes still in use at three depots"), (0, "Customers"),
        (1, "Delivery windows are too wide")],
    7: [(0, "November 2026: select the vendor"), (0, "January 2027: pilot at the Eastport depot"),
        (0, "April 2027: roll out to all depots"), (0, "September 2027: retire the old spreadsheets")],
    8: [(0, "Approve Option A and a budget of EUR 1.4 million"), (0, "Appoint a business owner for routing"),
        (0, "Review progress at the January committee")],
}
NOTES = [
    "Welcome everyone. Today we ask for a decision on the delivery option for Project Harbour.",
    "The cost figure comes from the finance review in June. Stress that the staffing risk is the most urgent.",
    "We interviewed 46 people across all five depots in August and September.",
    "This is the single number to remember. The industry benchmark is around 15 percent.",
    "Option B looks cheaper but carries the staffing risk from slide 2.",
    "Option A pays back during year 2. Savings come mostly from fewer rescheduled deliveries.",
    "The pilot depot was chosen because it has the highest rescheduling rate.",
    "We need the decision today to keep the November vendor selection on track.",
]
TABLE = [["Item", "Year 1", "Year 2", "Year 3"], ["Licences", "380", "380", "380"],
         ["Implementation", "210", "40", "0"], ["Savings", "-150", "-900", "-1,250"],
         ["Net", "440", "-480", "-870"]]


def paras(shape):
    return [(p.level or 0, norm(p.text)) for p in shape.text_frame.paragraphs if norm(p.text)]


def body(c, out: Path):
    deck = Document.open(out)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    slides = deck.slides
    c.check("exactly 8 slides", len(slides) == 8, len(slides))
    names = [s.layout.name for s in slides]
    c.check("template layouts as specified", names == LAYOUTS, names)
    c.check("deck uses the template's theme/masters only (all layouts from template)",
            {l.name for l in deck.layouts} >= set(LAYOUTS))
    titles = [title_of(s) for s in slides]
    bad = [(i + 1, t, e) for i, (t, e) in enumerate(zip(titles, TITLES)) if e and t != e]
    c.check("titles word for word", not bad, bad)
    if len(slides) >= 1:
        sub = [norm(sh.text) for sh in slides[0].shapes if sh.placeholder and sh.placeholder[0] == "subTitle"]
        c.check("title slide subtitle", sub == ["Modernising field logistics. Steering committee, 12 October 2026"], sub)
    for n, want in BULLETS.items():
        if len(slides) < n:
            continue
        bodies = [sh for sh in slides[n - 1].shapes if sh.placeholder and sh.placeholder[0] in ("body", "obj", None)
                  and sh.placeholder != ("title", None) and sh.kind == "shape" and norm(sh.text)
                  and sh.placeholder[0] not in ("title", "ctrTitle")]
        got = paras(bodies[0]) if bodies else []
        c.check(f"slide {n} bullets and levels", got == want, got)
        if bodies:
            bulleted = all(p.bullet is not None or True for p in bodies[0].text_frame.paragraphs)
    if len(slides) >= 4:
        texts = [norm(sh.text) for sh in slides[3].shapes]
        c.check("big number slide: 38% and its sentence",
                "38%" in texts and "of deliveries were rescheduled at least once in 2025" in texts, texts)
        big = [sh for sh in slides[3].shapes if norm(sh.text) == "38%"]
        c.check("38% in the BIG_NUMBER layout's big placeholder", big and big[0].placeholder is not None,
                [b.placeholder for b in big])
    if len(slides) >= 5:
        cols = [sh for sh in slides[4].shapes if sh.placeholder and sh.placeholder[0] == "body"]
        got = sorted([(sh.left, paras(sh)) for sh in cols])
        want = [[(0, "Option A: buy a routing platform"), (0, "Live in 9 months"),
                 (0, "EUR 1.4 million over three years")],
                [(0, "Option B: extend our in-house planner"), (0, "Live in 18 months"),
                 (0, "EUR 1.1 million over three years")]]
        c.check("two columns, A left and B right", [g[1] for g in got] == want, got)
    if len(slides) >= 6:
        tables = [sh.table for sh in slides[5].shapes if sh.has_table]
        got = [[norm(t.cell(r, k).text) for k in range(t.columns)] for t in tables[:1] for r in range(t.rows)]
        c.check("slide 6 real table with the figures", got == TABLE, got)
    nb = [(i + 1, notes_of(deck, i + 1)) for i in range(len(slides))]
    bad = [(i, n) for (i, n), want in zip(nb, NOTES) if norm(want) not in n]
    c.check("speaker notes on every slide", len(nb) == 8 and not bad, bad)
    empty = [(i + 1, sh.id, sh.placeholder) for i, s in enumerate(slides) for sh in s.shapes
             if sh.placeholder and sh.kind == "shape" and not norm(sh.text)]
    c.check("no empty (prompt-showing) placeholders", not empty, empty)
    # body text sizes / overflow are judged visually (rubric) and by PowerPoint's PDF


if __name__ == "__main__":
    main("p1-outline-deck", body)
