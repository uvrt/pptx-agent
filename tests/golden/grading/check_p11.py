"""P11: SmartArt updated in place: a four-step process in order, a third goal, a replaced
risk; still SmartArt; nothing else changed."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p11-update-smartart" / "input" / "delivery-process.pptx"


def diagrams(deck):
    out = []
    for slide in deck.slides:
        frames = [s for s in slide.shapes if s.kind == "graphic_frame" and s.diagram is not None]
        out.append(frames[0].diagram if frames else None)
    return out


def body(c, out: Path):
    deck, src = Document.open(out), Document.open(INPUT)
    new = {str(p) for p in deck.validate()} - {str(p) for p in src.validate()}
    c.check("no new validation problems", not new, sorted(new))
    c.check("still two slides", len(deck.slides) == 2, len(deck.slides))
    process, goals = diagrams(deck)
    c.check("slide 1 is still SmartArt", process is not None)
    c.check("slide 2 is still SmartArt", goals is not None)
    if process is None or goals is None:
        return
    steps = [norm(n.text) for n in process.nodes]
    c.check("the process reads Discover, Plan, Build, Ship", steps == ["Discover", "Plan", "Build",
                                                                     "Ship"], steps)
    c.check("the four steps are on one level", len({n.level for n in process.nodes}) == 1,
            [n.level for n in process.nodes])
    tree = [(n.level, norm(n.text)) for n in goals.nodes]
    top = [text for level, text in tree if level == min(lv for lv, _ in tree)]
    c.check("Goals and Risks are still the two headings", top == ["Goals", "Risks"], tree)
    under = {}
    current = None
    for level, text in tree:
        if text in ("Goals", "Risks"):
            current = text
            under[current] = []
        elif current:
            under[current].append(text)
    c.check("three goals, Lower cost third", under.get("Goals") == ["Faster edits", "Fewer prompts",
                                                                    "Lower cost"], under)
    c.check("the risk is Slow reviews", under.get("Risks") == ["Slow reviews"], under)
    texts = [[norm(s.text) for s in slide.shapes if s.kind == "shape"] for slide in deck.slides]
    before = [[norm(s.text) for s in slide.shapes if s.kind == "shape"] for slide in src.slides]
    c.check("titles and other shapes unchanged", texts == before, texts)


if __name__ == "__main__":
    main("p11-update-smartart", body)
