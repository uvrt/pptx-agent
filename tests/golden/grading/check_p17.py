"""P17: a corporate template made ready for a new year and language -- the legal line on the
master, the legal notice on a layout, a layout's prompt, Header & Footer (round 3: master and
layout editing, G5)."""
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402

from pptx_agent import Document  # noqa: E402

FOLDER = Path(os.environ.get("GOLDEN_INPUTS", str(INPUTS))) / "p17-template-new-year" / "input"
LEGAL = "© 2027 Example Corp B.V. Vertrouwelijk. Alle rechten voorbehouden."
NOTICE = ("Dit document is vertrouwelijk en uitsluitend bestemd voor de geadresseerde. Het mag "
          "niet worden gekopieerd of doorgegeven zonder schriftelijke toestemming van Example "
          "Corp B.V.")
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _footer_placeholders(slide):
    return {shape.placeholder.type: shape for shape in slide.shapes
            if shape.placeholder is not None and shape.placeholder.type in ("dt", "ftr", "sldNum")}


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(FOLDER / "example-template.pptx")
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    c.check("the slides' own text unchanged",
            [[norm(s.text) for s in slide.shapes if s.placeholder is None
              or s.placeholder.type not in ("dt", "ftr", "sldNum")] for slide in deck.slides]
            == [[norm(s.text) for s in slide.shapes] for slide in src.slides],
            [slide.title for slide in deck.slides])
    master = deck.masters[0]
    legal = [s for s in master.shapes if s.placeholder is None and norm(s.text)]
    c.check("the master's legal line, in Dutch, for 2027", [norm(s.text) for s in legal] == [LEGAL],
            [s.text for s in legal])
    c.check("the legal line kept its place and size",
            legal and legal[0].bounds == src.masters[0].shapes[-1].bounds
            and legal[0].paragraphs[0].runs[0].size == 8, legal and legal[0].bounds)
    left = [b.address for b in deck.find_text("2026", scope="all")] + \
        [b.address for b in deck.find_text("Ltd", scope="all")] + \
        [b.address for b in deck.find_text("Confidential", scope="all")]
    c.check("no 2026, Ltd. or English legal text anywhere", not left, left)
    names = [layout.name for layout in deck.layouts]
    c.check("the Legal layout is Juridisch now", "Juridisch" in names and "Legal" not in names, names)
    if "Juridisch" in names:
        notice = [s for s in deck.layout("Juridisch").shapes if s.placeholder is None
                  and norm(s.text)]
        c.check("its notice in Dutch", [norm(s.text) for s in notice] == [NOTICE],
                [s.text for s in notice])
        c.check("the legal slide still on it",
                [s.layout.name for s in deck.slides][-1] == "Juridisch",
                [s.layout.name for s in deck.slides])
    body_ph = deck.layout("Title and Content").placeholder("body")
    c.check("the content prompt in Dutch, as a custom prompt",
            body_ph.prompt is not None and norm(body_ph.prompt.split("\n")[0]) ==
            "Typ hier uw bevinding", body_ph.prompt)
    c.check("the other levels' prompts kept", len(body_ph.paragraphs) == 5,
            len(body_ph.paragraphs))
    title_slide, *rest = deck.slides
    c.check("no footer on the title slide", not _footer_placeholders(title_slide),
            list(_footer_placeholders(title_slide)))
    ok, detail = True, []
    for number, slide in enumerate(rest, 2):
        found = _footer_placeholders(slide)
        detail.append((slide.slide_id, sorted(found)))
        number_field = found.get("sldNum") and found["sldNum"]._element.find(f".//{A}fld")
        if sorted(found) != ["ftr", "sldNum"] or number_field is None \
                or number_field.get("type") != "slidenum" \
                or norm(found["ftr"].text) != "Example Corp B.V.":
            ok = False
    c.check("footer and slide number on every other slide, no date", ok, detail)
    presentation = deck.package.tree(deck.package.presentation_part())
    c.check("Don't show on title slide recorded", presentation.get("showSpecialPlsOnTitleSld")
            == "0", presentation.get("showSpecialPlsOnTitleSld"))
    hf = deck.package.tree(master.part_path).find(f"{P}hf")
    c.check("the master's Header & Footer record", hf is not None
            and dict(hf.attrib) == {"hdr": "0", "dt": "0"}, None if hf is None else dict(hf.attrib))
    layout_footers = {layout.name: norm(layout.placeholder("footer").text)
                      for layout in deck.layouts
                      if any(s.placeholder and s.placeholder.type == "ftr" for s in layout.shapes)}
    c.check("every layout's footer says it too", set(layout_footers.values()) == {"Example Corp B.V."},
            layout_footers)
    c.check("the editing language is Dutch", deck.language == "nl-NL", deck.language)
    svg = deck.render_svg([3])[0]
    c.check("slide 3 draws the legal line, the footer and its number",
            "Vertrouwelijk" in svg and "Example Corp B.V." in svg and ">3<" in svg.replace(" ", ""),
            [word for word in ("Vertrouwelijk", "Example Corp B.V.") if word not in svg])


if __name__ == "__main__":
    main("p17", body)
