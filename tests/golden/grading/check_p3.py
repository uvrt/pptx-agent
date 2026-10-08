"""P3: split the overcrowded 'Lessons learned' slide into two; notes follow their topics."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm, notes_of, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p3-split-slide/input/pilot-retrospective.pptx"
EMU_PT = 12700


def bullets(slide):
    for sh in slide.shapes:
        if sh.placeholder and sh.placeholder[0] not in ("title", "ctrTitle") and norm(sh.text):
            return sh, [(p.level or 0, norm(p.text)) for p in sh.text_frame.paragraphs if norm(p.text)]
    return None, []


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    s = deck.slides
    c.check("5 slides", len(s) == 5, len(s))
    titles = [title_of(x) for x in s]
    c.check("titles", titles == ["Pilot retrospective", "Agenda", "Lessons learned (1 of 2)",
                                 "Lessons learned (2 of 2)", "Next steps"], titles)
    _, orig = bullets(src.slides[2])
    if len(s) < 5:
        return
    sh3, b3 = bullets(s[2])
    sh4, b4 = bullets(s[3])
    c.check("slide 3: Planning, Delivery, People word for word with levels", b3 == orig[:12], b3)
    c.check("slide 4: Tooling, Customers word for word with levels", b4 == orig[12:], b4)
    lay = src.slides[2].layout.name
    c.check("both slides use the original layout", s[2].layout.name == lay and s[3].layout.name == lay,
            (s[2].layout.name, s[3].layout.name))
    n3, n4 = notes_of(deck, 3), notes_of(deck, 4)
    c.check("slide 3 notes: Planning, Delivery, People only",
            all(k in n3 for k in ("Planning:", "Delivery:", "People:")) and not any(
                k in n3 for k in ("Tooling:", "Customers:")), n3[:300])
    c.check("slide 4 notes: Tooling, Customers only",
            all(k in n4 for k in ("Tooling:", "Customers:")) and not any(
                k in n4 for k in ("Planning:", "Delivery:", "People:")), n4[:300])
    others_ok = all(deck.to_outline(slides=[i]).split("\n", 1)[1] == src.to_outline(slides=[j]).split("\n", 1)[1]
                    for i, j in ((1, 1), (2, 2), (5, 4)))
    c.check("other slides unchanged (outline incl. notes)", others_ok)
    # font sizes: explicit run sizes, where set, must be >= 18 pt; effective size is judged visually
    small = []
    for sh in (sh3, sh4):
        if sh is None:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.size is not None and r.size < 18:
                    small.append((norm(p.text)[:25], r.size))
    c.check("no explicit run size below 18 pt", not small, small[:5])
    # rough fit estimate: lines x 1.2 x size must fit the placeholder height (sizes default to the
    # layout's 28/24 pt for levels 0/1 when not set explicitly)
    for n, sh, b in ((3, sh3, b3), (4, sh4, b4)):
        if sh is None:
            continue
        need = 0.0
        for p in sh.text_frame.paragraphs:
            lvl = p.level or 0
            size = next((r.size for r in p.runs if r.size), None) or (28 if lvl == 0 else 24)
            need += size * 1.08 + (10 if lvl == 0 else 5)   # Office theme: 90% line spacing, spcBef 10/5
        have = sh.height / EMU_PT if sh.height else 0
        c.check(f"slide {n}: text height estimate fits the placeholder (Office metrics; 28/24 pt unless set)",
                need <= have * 1.02, f"need~{need:.0f}pt have {have:.0f}pt")


if __name__ == "__main__":
    main("p3-split-slide", body)
