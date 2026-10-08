"""P12: the brand applied through the theme: four slots and two fonts set, every shape still
on theme colours and theme fonts, nothing else changed."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, all_shapes, main, norm  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p12-theme-rebrand" / "input" / "patient-review.pptx"
WANT = {"accent1": "#0B6E4F", "accent2": "#E09F3E", "dk2": "#1B2A41", "lt2": "#F4F1EA"}
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def body(c, out: Path):
    deck, src = Document.open(out), Document.open(INPUT)
    new = {str(p) for p in deck.validate()} - {str(p) for p in src.validate()}
    c.check("no new validation problems", not new, sorted(new))
    colors = deck.theme.colors
    c.check("the four slots hold the new colours", all(colors[k].upper() == v for k, v in WANT.items()),
            {k: colors[k] for k in WANT})
    others = {k: v for k, v in colors.items() if k not in WANT and k not in ("bg1", "tx1", "bg2", "tx2")}
    before = {k: v for k, v in src.theme.colors.items() if k in others}
    c.check("the other slots unchanged", others == before, (others, before))
    c.check("headings in Lora, body in Source Sans 3",
            (deck.theme.fonts.major, deck.theme.fonts.minor) == ("Lora", "Source Sans 3"),
            (deck.theme.fonts.major, deck.theme.fonts.minor))
    rgb = []
    fonts = []
    for slide in deck.slides:
        for shape in all_shapes(slide):
            element = shape._element
            rgb += [shape.id for node in element.iter(f"{A}srgbClr")]
            fonts += [(shape.id, node.get("typeface")) for node in element.iter(f"{A}latin")
                      if not (node.get("typeface") or "").startswith("+")]
    c.check("no shape got a hard-coded colour", not rgb, rgb)
    c.check("no shape got a font of its own", not fonts, fonts)
    texts = [[norm(s.text or "") for s in all_shapes(slide) if s.kind == "shape"] for slide in deck.slides]
    was = [[norm(s.text or "") for s in all_shapes(slide) if s.kind == "shape"] for slide in src.slides]
    c.check("the slides' text unchanged", texts == was)
    boxes = [[s.bounds for s in all_shapes(slide)] for slide in deck.slides]
    were = [[s.bounds for s in all_shapes(slide)] for slide in src.slides]
    c.check("the slides' shapes unchanged", boxes == were)


if __name__ == "__main__":
    main("p12-theme-rebrand", body)
