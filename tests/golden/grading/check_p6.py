"""P6: delete, move, duplicate (+ new notes), add a Section Header slide, update the agenda."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm, notes_of, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p6-restructure-deck/input/website-relaunch.pptx"
ORDER = ["Website relaunch: project review", "Agenda", "Goals", "Part 2: Delivery", "Budget", "Timeline",
         "Team: design", "Team: engineering", "Risks", "Questions"]
SRC_INDEX = {"Website relaunch: project review": 1, "Goals": 3, "Timeline": 4, "Budget": 5, "Team: design": 6,
             "Risks": 7, "Questions": 9}


def body_lines(slide):
    out = []
    for sh in slide.shapes:
        ph = sh.placeholder
        if ph and ph[0] in ("title", "ctrTitle"):
            continue
        if sh.kind == "shape" and norm(sh.text):
            out += [norm(p.text) for p in sh.text_frame.paragraphs if norm(p.text)]
    return out


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    titles = [title_of(s) for s in deck.slides]
    c.check("final slide order", titles == ORDER, titles)
    c.check("appendix slide gone", not any("Appendix" in t for t in titles) and
            "Traffic 2024" not in deck.to_outline(), titles)
    if "Part 2: Delivery" in titles:
        i = titles.index("Part 2: Delivery")
        s = deck.slides[i]
        c.check("section slide uses the Section Header layout", s.layout.name == "Section Header", s.layout.name)
        empty = [sh.placeholder for sh in s.shapes if sh.placeholder and not norm(sh.text) and sh.kind == "shape"]
        c.check("section slide has no empty placeholder", not empty, empty)
    if "Agenda" in titles:
        lines = body_lines(deck.slides[titles.index("Agenda")])
        c.check("agenda in the new order", lines == ["Goals", "Budget", "Timeline", "Team", "Risks", "Questions"],
                lines)
    if "Team: engineering" in titles:
        i = titles.index("Team: engineering")
        s = deck.slides[i]
        c.check("engineering bullets", body_lines(s) == ["Lead: Priya Nair", "Two backend developers",
                                                          "Two frontend developers"], body_lines(s))
        c.check("engineering slide same layout as design slide",
                s.layout.name == src.slides[5].layout.name, s.layout.name)
        n = notes_of(deck, i + 1)
        c.check("engineering notes say the Utrecht office", "Engineering works from the Utrecht office." in n, n)
        c.check("engineering notes say nothing about design", "design" not in n.lower(), n)
    unchanged = []
    for t, j in SRC_INDEX.items():
        if t not in titles:
            unchanged.append((t, "missing"))
            continue
        i = titles.index(t) + 1
        a = deck.to_outline(slides=[i]).split("\n", 1)[1]
        b = src.to_outline(slides=[j]).split("\n", 1)[1]
        if a != b:
            unchanged.append((t, a[:120], b[:120]))
    c.check("other slides unchanged, notes included", not unchanged, unchanged[:3])
    if "Agenda" in titles:
        n = notes_of(deck, titles.index("Agenda") + 1)
        c.check("agenda notes kept", "Keep the agenda short." in n, n)
    ids_src = {s.slide_id for s in src.slides}
    kept = [s.slide_id for s in deck.slides if title_of(s) in SRC_INDEX]
    c.check("moved slides are the original slides (ids kept), not re-created", all(k in ids_src for k in kept),
            kept)


if __name__ == "__main__":
    main("p6-restructure-deck", body)
