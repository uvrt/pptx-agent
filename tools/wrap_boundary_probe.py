"""Where PowerPoint wraps a short label: one word per slide, in boxes around its width.

Trial 2's p4 run 2 had a "Pass" label (Aptos, 18 pt) in a 650,000 EMU text box that
``text_fit()`` kept on one line and PowerPoint broke into "Pas / s".  This probe puts one
word per slide in a text box whose text width (the box less its 0.1 in insets) is the
measured width of the word, kerned (K) or not (U), give or take a fraction of a point:

    python tools/wrap_boundary_probe.py build ~/wrap-probe.pptx cases.json
    (export it to PDF with PowerPoint: tests/oracle.py's export_pdf)
    python tools/wrap_boundary_probe.py read ~/wrap-probe.pdf cases.json

``read`` prints each case with the lines PowerPoint drew (from the PDF's text, by
baseline) and the advance of every glyph it placed.  What it showed, on PowerPoint 16 for
Mac (fit.py's module doc has the rule): insets are subtracted exactly; PowerPoint kerns
with the face's legacy ``kern`` table only -- Aptos's ``s s`` class pair, which only its
``GPOS`` has, is drawn unkerned at 12, 18 and 24 pt -- and draws each advance up to about
0.05 pt off the font's.  Needs pypdfium2 to read.
"""

from __future__ import annotations

import ctypes
import json
import sys

FONTS = ("Aptos", "Calibri", "Arial")
SIZES = (12, 18, 24)
WORDS = ("Pass", "Fail", "Total", "AVAWAY", "Review")
INSET = 91440
PX = 9525          # EMU per px at 96 dpi


def build(deck_path: str, cases_path: str) -> None:
    from ooxml_common.text.measure import DefaultTextMeasurer

    from pptx_agent import Document
    from pptx_agent.oxml.xml import qn

    measurer = DefaultTextMeasurer()
    deck = Document.new(size="16:9")
    cases = []

    def add(word, size, font, text_emu, insets, tag, kern=None):
        slide = deck.add_slide("Blank")
        side = INSET if insets else 0
        box = slide.add_textbox(914400, 914400, round(text_emu) + 2 * side, 600000, word)
        run = box.text_frame.paragraph(0).run(0)
        run.format(size=size, typeface=font)
        if not insets:
            box.text_frame.set_insets(left=0, right=0)
        if kern is not None:
            run._element().find(qn("a:rPr")).set("kern", str(kern))
        cases.append(dict(page=len(deck.slides), word=word, size=size, font=font,
                          text_emu=round(text_emu), insets=insets, tag=tag, kern=kern))

    for font in FONTS:
        for size in SIZES:
            for word in WORDS:
                kerned = measurer.measure_text_width(word, size, False, font) * PX
                plain = sum(measurer.measure_text_width(c, size, False, font) for c in word) * PX
                widths = {"K-0.3pt": kerned - 3810, "K+0.05pt": kerned + 635,
                          "K+0.3pt": kerned + 3810, "U-0.05pt": plain - 635,
                          "U+0.05pt": plain + 635, "U+0.3pt": plain + 3810}
                if plain - kerned > 2000:
                    widths["mid"] = (kerned + plain) / 2
                seen = set()
                for tag, width in widths.items():
                    if round(width) not in seen:
                        seen.add(round(width))
                        add(word, size, font, width, True, tag)
    kerned = measurer.measure_text_width("Pass", 18, False, "Aptos") * PX
    plain = sum(measurer.measure_text_width(c, 18, False, "Aptos") for c in "Pass") * PX
    for tag, width in {"K+0.05pt": kerned + 635, "U-0.05pt": plain - 635,
                       "U+0.05pt": plain + 635}.items():
        add("Pass", 18, "Aptos", width, False, "no insets " + tag)
        add("Pass", 18, "Aptos", width, True, "kern=0 " + tag, kern=0)
    deck.save(deck_path)
    with open(cases_path, "w") as handle:
        json.dump(cases, handle, indent=1)
    print(f"{len(cases)} slides")


def read(pdf_path: str, cases_path: str) -> None:
    import pypdfium2
    import pypdfium2.raw as raw

    with open(cases_path) as handle:
        cases = json.load(handle)
    document = pypdfium2.PdfDocument(pdf_path)
    for case in cases:
        text_page = document[case["page"] - 1].get_textpage()
        chars = []
        for index in range(text_page.count_chars()):
            char = chr(raw.FPDFText_GetUnicode(text_page, index))
            x, y = ctypes.c_double(), ctypes.c_double()
            raw.FPDFText_GetCharOrigin(text_page, index, ctypes.byref(x), ctypes.byref(y))
            if char.strip():
                chars.append((char, x.value, round(y.value, 2)))
        baselines = sorted({c[2] for c in chars}, reverse=True)
        lines = ["".join(c[0] for c in chars if c[2] == y) for y in baselines]
        advances = [round(b[1] - a[1], 3) for a, b in zip(chars, chars[1:]) if a[2] == b[2]]
        print(f"{case['font']:8}{case['size']:3} {case['word']:7} {case['tag']:20} "
              f"text width {case['text_emu']:8,}  {' / '.join(lines):12} {advances}")


if __name__ == "__main__":
    {"build": build, "read": read}[sys.argv[1]](*sys.argv[2:])
