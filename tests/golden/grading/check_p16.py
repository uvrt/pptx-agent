"""P16: the approach slide -- four phases in a row, each an SVG icon centred above its label,
equally spaced across the content area (round 3: SVG pictures and ppt_arrange's align and
distribute)."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm, title_of  # noqa: E402

from pptx_agent import Document  # noqa: E402

FOLDER = Path(os.environ.get("GOLDEN_INPUTS", str(INPUTS))) / "p16-approach-icons" / "input"
PHASES = [("Diagnose", "Baseline the spend", "diagnose.svg"),
          ("Design", "Category strategies", "design.svg"),
          ("Negotiate", "Supplier rounds", "negotiate.svg"),
          ("Embed", "Savings tracking", "embed.svg")]
PT = 12700


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(FOLDER / "halden-proposal-draft.pptx")
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    titles = [title_of(s) for s in deck.slides]
    c.check("still 3 slides", len(titles) == 3, titles)
    slide = deck.slides[1]
    title = next(sh for sh in slide.shapes if sh.placeholder and sh.placeholder[0] == "title")
    fit = title.text_fit()
    c.check("a one-line title that fits", bool(norm(title.text)) and "\n" not in title.text
            and tuple(fit.lines) == (1,) and not fit.overflows, (title.text, fit))
    c.check("a takeaway, not a label", len(norm(title.text).split()) >= 6
            and norm(title.text).lower() not in {"our approach", "approach"}, title.text)

    # slide.shapes lists a group and then its members: the members are what is drawn.
    shapes = [s for s in slide.shapes if s.kind != "group" and s.id != title.id]
    pictures = [s for s in shapes if s.kind == "picture"]
    labels = [s for s in shapes if s.kind != "picture" and norm(s.text)]
    c.check("four pictures, four labels", len(pictures) == 4 and len(labels) == 4,
            (len(pictures), len(labels)))
    if len(pictures) != 4 or len(labels) != 4:
        return
    pictures.sort(key=lambda s: s.drawn_bounds[0])
    labels.sort(key=lambda s: s.drawn_bounds[0])
    icons = [(FOLDER / name).read_bytes() for _, _, name in PHASES]
    c.check("the icons are SVG pictures of the given files, in phase order",
            [p.vector for p in pictures] == icons, [p.vector is not None for p in pictures])
    c.check("each has a PNG fallback", all(p.image_part and p.image_part.endswith(".png")
                                           and p.image_size for p in pictures),
            [p.image_part for p in pictures])
    c.check("the icons are one size", len({(p.width, p.height) for p in pictures}) == 1,
            [(p.width, p.height) for p in pictures])
    texts = [norm(s.text) for s in labels]
    c.check("the labels, left to right", all(name in text and line in text for text, (name, line, _)
                                             in zip(texts, PHASES)), texts)
    pairs = list(zip(pictures, labels))
    c.check("each icon centred above its label",
            all(abs((p.drawn_bounds[0] + p.drawn_bounds[2] / 2)
                    - (t.drawn_bounds[0] + t.drawn_bounds[2] / 2)) <= 0.5 * PT
                and p.drawn_bounds[1] + p.drawn_bounds[3] <= t.drawn_bounds[1]
                for p, t in pairs), [(p.drawn_bounds, t.drawn_bounds) for p, t in pairs])
    steps = []
    for p, t in pairs:
        a, b = p.drawn_bounds, t.drawn_bounds
        left, top = min(a[0], b[0]), min(a[1], b[1])
        steps.append((left, top, max(a[0] + a[2], b[0] + b[2]) - left,
                      max(a[1] + a[3], b[1] + b[3]) - top))
    gaps = [b[0] - (a[0] + a[2]) for a, b in zip(steps, steps[1:])]
    c.check("equal gaps between the steps", gaps and max(gaps) - min(gaps) <= 0.5 * PT
            and min(gaps) > 0, [g / PT for g in gaps])
    area = slide.content_area
    c.check("the steps span the content area", abs(steps[0][0] - area[0]) <= PT
            and abs(steps[-1][0] + steps[-1][2] - (area[0] + area[2])) <= PT,
            ([s[0] / PT for s in steps], area))
    c.check("one row: the steps' tops line up", len({s[1] for s in steps}) == 1
            and len({p.drawn_bounds[1] for p in pictures}) == 1, [s[1] / PT for s in steps])
    c.check("below the title, inside the content area",
            all(s[1] >= title.top + title.height and s[1] + s[3] <= area[1] + area[3] + PT
                for s in steps), (steps, area))
    c.check("nothing overlaps or leaves the slide", not deck.overflows(slides=[2])
            and not slide.collisions(),
            (deck.overflows(slides=[2]), slide.collisions()))
    same = all(deck.to_outline(slides=[i]).split("\n", 1)[1]
               == src.to_outline(slides=[i]).split("\n", 1)[1] for i in (1, 3))
    c.check("other slides unchanged", same)


if __name__ == "__main__":
    main("p16-approach-icons", body)
