"""A finished slide moved onto another layout (production feedback round 7, L2).

The spec is PowerPoint 16 (Mac) changing the layout of 53 made-up slides
(agent-work/feedback-round7/oracle/results.md, A1-A9; outputs/L2-compare.txt,
L2b-compare.txt, L2c-compare.txt; B2 for a picture placeholder).  Each case below rebuilds
one measured slide with the same layouts and checks the shapes PowerPoint wrote, line by
line -- except the names of the placeholders added, which PowerPoint writes in its
interface language ("Tijdelijke aanduiding voor inhoud 3") and this the layout's."""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from layout_decks import build, dump, ph, pic, png, rect
from pptx_agent import Document
from pptx_agent.edit.layoutchange import ORPHAN_IDX
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from pptx_agent.validate import check

PT = 12700
OV = dict(tbox=(100, 40, 500, 60), bbox=(100, 200, 400, 200), tsz=4000, bsz=1200)


def la(n, title=True, body=True, tbox=None, bbox=None, tsz=None, bsz=None, bname=None,
       btype=None, extra=(), pre=()):
    """A slide on LA (title, content idx 1) as the oracle wrote them."""
    shapes = list(pre) + [ph(2, f"S{n} Title", "title", None, tbox,
                             text=f"S{n} title" if title else None, size=tsz)]
    shapes.append(ph(3, bname or f"S{n} Body", btype, 1, bbox,
                     text=f"S{n} body" if body else None, size=bsz))
    return shapes + list(extra)


def lc(n, left=True, right=True, order=("left", "right")):
    """A slide on LC (title, two half-size content placeholders)."""
    sides = {"left": ph(3, f"S{n} Left", None, 1, None, text=f"S{n} left" if left else None,
                        extra=' sz="half"'),
             "right": ph(4, f"S{n} Right", None, 2, None,
                         text=f"S{n} right" if right else None, extra=' sz="half"')}
    return [ph(2, f"S{n} Title", "title", None, None, text=f"S{n} title")] + [
        sides[side] for side in order]


def moved(start, shapes, *targets):
    """The slide built on ``start`` with ``shapes``, after set_layout to each target."""
    deck = build([(start, shapes)])
    slide = deck.slides[0]
    change = None
    for target in targets:
        change = slide.set_layout(target)
    return deck, deck.slides[0], change


PICTURE_PH = '<p:ph type="pic" idx="13"/>'

# (case, start layout, shapes, targets, the shapes PowerPoint wrote)
CASES = [
    ("L2 S1 idx 1 -> idx 2, other shapes kept", "LA",
     la(1, extra=[rect(9, "S1 Rect", (700, 450, 200, 60), "S1 rect")]), ["LB"],
     ["sp id=2 'S1 Title' ph(type=title) tx=Y 'S1 title'",
      "sp id=3 'S1 Body' ph(idx=2) tx=Y 'S1 body'",
      "sp id=9 'S1 Rect' xfrm(700,450,200,60) tx=Y 'S1 rect'"]),
    ("L2 S2 boxes dropped, sizes kept, an empty placeholder added", "LA",
     la(2, **OV, extra=[rect(9, "S2 Rect", (700, 450, 200, 60), "S2 rect")]), ["LC"],
     ["sp id=2 'S2 Title' ph(type=title) sz=['4000'] tx=Y 'S2 title'",
      "sp id=3 'S2 Body' ph(idx=1) sz=['1200'] tx=Y 'S2 body'",
      "sp id=4 'LC Right' ph(idx=2) tx=Y ''",
      "sp id=9 'S2 Rect' xfrm(700,450,200,60) tx=Y 'S2 rect'"]),
    ("L2 S3 an orphan on the right half", "LC", lc(3), ["LA"],
     ["sp id=2 'S3 Title' ph(type=title) tx=Y 'S3 title'",
      "sp id=3 'S3 Left' ph(idx=1) tx=Y 'S3 left'",
      f"sp id=4 'S3 Right' ph(sz=half,idx={ORPHAN_IDX}) xfrm(540,120,420,380) tx=Y "
      "'S3 right'"]),
    ("L2 S4 an orphan on the left half, the empty one gone", "LC", lc(4, right=False), ["LF"],
     ["sp id=2 'S4 Title' ph(type=title) tx=Y 'S4 title'",
      f"sp id=3 'S4 Left' ph(sz=half,idx={ORPHAN_IDX}) xfrm(0,120,420,380) tx=Y 'S4 left'"]),
    ("L2 S5 a picture placeholder takes no text", "LA", la(5), ["LD"],
     ["sp id=2 'S5 Title' ph(type=title) tx=Y 'S5 title'",
      "sp id=4 'LD Picture' ph(type=pic,idx=13) tx=N ''",
      f"sp id=3 'S5 Body' ph(idx={ORPHAN_IDX}) xfrm(0,120,860,380) tx=Y 'S5 body'"]),
    ("L2 S6 an empty picture placeholder goes", "LD",
     [ph(2, "S6 Title", "title", None, None, text="S6 title"),
      ph(3, "S6 Picture", "pic", 13)], ["LA"],
     ["sp id=2 'S6 Title' ph(type=title) tx=Y 'S6 title'",
      "sp id=4 'LA Body' ph(idx=1) tx=Y ''"]),
    ("L2 S7 content into a body placeholder of another idx", "LA", la(7), ["LE"],
     ["sp id=2 'S7 Title' ph(type=title) tx=Y 'S7 title'",
      "sp id=3 'S7 Body' ph(type=body,idx=14) tx=Y 'S7 body'"]),
    ("L2 S8 a title orphan keeps its type", "LA", la(8, **OV), ["LG"],
     ["sp id=3 'S8 Body' ph(idx=1) sz=['1200'] tx=Y 'S8 body'",
      f"sp id=2 'S8 Title' ph(type=title,idx={ORPHAN_IDX}) xfrm(0,40,500,60) sz=['4000'] "
      "tx=Y 'S8 title'"]),
    ("L2 S9 the layout's order", "LA", la(9), ["LI"],
     ["sp id=3 'S9 Body' ph(idx=1) tx=Y 'S9 body'",
      "sp id=2 'S9 Title' ph(type=title) tx=Y 'S9 title'"]),
    ("L2 S10 lowest idx first, never the name", "LA", la(10, bname="Notes Area"), ["LJ"],
     ["sp id=2 'S10 Title' ph(type=title) tx=Y 'S10 title'",
      "sp id=4 'Other' ph(idx=8) tx=Y ''",
      "sp id=3 'Notes Area' ph(idx=7) tx=Y 'S10 body'"]),
    ("L2 S11 title to centred title, content to subtitle", "LA", la(11, **OV), ["LL"],
     ["sp id=2 'S11 Title' ph(type=ctrTitle) sz=['4000'] tx=Y 'S11 title'",
      "sp id=3 'S11 Body' ph(type=subTitle,idx=1) sz=['1200'] tx=Y 'S11 body'"]),
    ("L2 S13 a body placeholder becomes content", "LA", la(13, btype="body"), ["LB"],
     ["sp id=2 'S13 Title' ph(type=title) tx=Y 'S13 title'",
      "sp id=3 'S13 Body' ph(idx=2) tx=Y 'S13 body'"]),
    ("L2 S14 two into one: the lower idx matched", "LC", lc(14), ["LB"],
     ["sp id=2 'S14 Title' ph(type=title) tx=Y 'S14 title'",
      "sp id=3 'S14 Left' ph(idx=2) tx=Y 'S14 left'",
      f"sp id=4 'S14 Right' ph(sz=half,idx={ORPHAN_IDX}) xfrm(540,120,420,380) tx=Y "
      "'S14 right'"]),
    ("L2 S15 idx 1 and 5 to idx 1 and 3", "LA",
     la(15, extra=[ph(5, "S15 Orphan5", None, 5, (600, 300, 300, 100),
                      text="S15 orphan idx5")]), ["LP"],
     ["sp id=2 'S15 Title' ph(type=title) tx=Y 'S15 title'",
      "sp id=3 'S15 Body' ph(idx=1) tx=Y 'S15 body'",
      "sp id=5 'S15 Orphan5' ph(idx=3) tx=Y 'S15 orphan idx5'"]),
    ("L2 S16 content into body", "LA", la(16), ["LK"],
     ["sp id=2 'S16 Title' ph(type=title) tx=Y 'S16 title'",
      "sp id=3 'S16 Body' ph(type=body,idx=1) tx=Y 'S16 body'"]),
    ("L2 S17 content prefers content over a lower-idx body", "LA", la(17), ["LM"],
     ["sp id=2 'S17 Title' ph(type=title) tx=Y 'S17 title'",
      "sp id=4 'LM Text5' ph(type=body,idx=5) tx=Y ''",
      "sp id=3 'S17 Body' ph(idx=6) tx=Y 'S17 body'"]),
    ("L2 S18 a table placeholder takes no text", "LA", la(18), ["LN"],
     ["sp id=2 'S18 Title' ph(type=title) tx=Y 'S18 title'",
      "sp id=4 'LN Table' ph(type=tbl,idx=1) tx=N ''",
      f"sp id=3 'S18 Body' ph(idx={ORPHAN_IDX}) xfrm(0,120,860,380) tx=Y 'S18 body'"]),
    ("L2 S19 an empty title replaced, the next free id", "LA", la(19, title=False), ["LF"],
     ["sp id=4 'LF Title' ph(type=title) tx=Y ''",
      f"sp id=3 'S19 Body' ph(idx={ORPHAN_IDX}) xfrm(0,120,860,380) tx=Y 'S19 body'"]),
    ("L2 S20 the same layout again drops the boxes", "LA",
     la(20, **OV, extra=[rect(9, "S20 Rect", (700, 450, 200, 60), "S20 rect")]), ["LA"],
     ["sp id=2 'S20 Title' ph(type=title) sz=['4000'] tx=Y 'S20 title'",
      "sp id=3 'S20 Body' ph(idx=1) sz=['1200'] tx=Y 'S20 body'",
      "sp id=9 'S20 Rect' xfrm(700,450,200,60) tx=Y 'S20 rect'"]),
    ("L2 S22 two content placeholders again lose sz", "LC", lc(22), ["LA", "LC"],
     ["sp id=2 'S22 Title' ph(type=title) tx=Y 'S22 title'",
      "sp id=3 'S22 Left' ph(idx=1) tx=Y 'S22 left'",
      "sp id=4 'S22 Right' ph(idx=2) tx=Y 'S22 right'"]),
    ("L2 S23 all empty: ids counted before they go", "LA", la(23, title=False, body=False),
     ["LD"],
     ["sp id=4 'LD Title' ph(type=title) tx=Y ''",
      "sp id=5 'LD Picture' ph(type=pic,idx=13) tx=N ''"]),
    ("L2 S24 there and back", "LA", la(24, **OV), ["LC", "LA"],
     ["sp id=2 'S24 Title' ph(type=title) sz=['4000'] tx=Y 'S24 title'",
      "sp id=3 'S24 Body' ph(idx=1) sz=['1200'] tx=Y 'S24 body'"]),
    ("L2b S1 sorted by idx, not tree order", "LC", lc(1, order=("right", "left")), ["LB2"],
     ["sp id=2 'S1 Title' ph(type=title) tx=Y 'S1 title'",
      "sp id=3 'S1 Left' ph(idx=5) tx=Y 'S1 left'",
      f"sp id=4 'S1 Right' ph(sz=half,idx={ORPHAN_IDX}) xfrm(540,120,420,380) tx=Y "
      "'S1 right'"]),
    ("L2b S2 a matching name does not count", "LA", la(2, bname="Notes Area"), ["LJ2"],
     ["sp id=2 'S2 Title' ph(type=title) tx=Y 'S2 title'",
      "sp id=3 'Notes Area' ph(idx=7) tx=Y 'S2 body'",
      "sp id=4 'Notes Area' ph(idx=8) tx=Y ''"]),
    ("L2b S3 the lowest idx, not the first in the tree", "LA", la(3), ["LJ3"],
     ["sp id=2 'S3 Title' ph(type=title) tx=Y 'S3 title'",
      "sp id=4 'Zeta' ph(idx=9) tx=Y ''",
      "sp id=3 'S3 Body' ph(idx=3) tx=Y 'S3 body'"]),
    ("L2b S4 content first to content, then to body", "LC", lc(4), ["LQ"],
     ["sp id=2 'S4 Title' ph(type=title) tx=Y 'S4 title'",
      "sp id=3 'S4 Left' ph(idx=1) tx=Y 'S4 left'",
      "sp id=4 'S4 Right' ph(type=body,idx=2) tx=Y 'S4 right'"]),
    ("L2b S5 a shape before the placeholders stays first", "LA",
     la(5, pre=[rect(9, "S5 Rect first", (700, 450, 200, 60), "S5 rect")]), ["LC"],
     ["sp id=9 'S5 Rect first' xfrm(700,450,200,60) tx=Y 'S5 rect'",
      "sp id=2 'S5 Title' ph(type=title) tx=Y 'S5 title'",
      "sp id=3 'S5 Body' ph(idx=1) tx=Y 'S5 body'",
      "sp id=4 'LC Right' ph(idx=2) tx=Y ''"]),
    ("L2b S6 a picture that is no placeholder is untouched", "LA",
     la(6, extra=[pic(9, "S6 Image", "rId2", (700, 400, 100, 100))]), ["LC"],
     ["sp id=2 'S6 Title' ph(type=title) tx=Y 'S6 title'",
      "sp id=3 'S6 Body' ph(idx=1) tx=Y 'S6 body'",
      "sp id=4 'LC Right' ph(idx=2) tx=Y ''",
      "pic id=9 'S6 Image' xfrm(700,400,100,100) tx=N ''"]),
    ("L2b S8 both orphans on a blank layout", "LA", la(8), ["LR"],
     [f"sp id=2 'S8 Title' ph(type=title,idx={ORPHAN_IDX}) xfrm(0,20,860,80) tx=Y "
      "'S8 title'",
      f"sp id=3 'S8 Body' ph(idx={ORPHAN_IDX}) xfrm(0,120,860,380) tx=Y 'S8 body'"]),
    ("L2b S9 the filled one of two takes the place", "LC", lc(9, left=False), ["LA"],
     ["sp id=2 'S9 Title' ph(type=title) tx=Y 'S9 title'",
      "sp id=4 'S9 Right' ph(idx=1) tx=Y 'S9 right'"]),
    ("L2b S10 a placeholder of an idx the layout lacks still matches", "LA",
     la(10, extra=[ph(5, "S10 Body5", None, 5, (600, 300, 300, 100), text="S10 body5")]),
     ["LC"],
     ["sp id=2 'S10 Title' ph(type=title) tx=Y 'S10 title'",
      "sp id=3 'S10 Body' ph(idx=1) tx=Y 'S10 body'",
      "sp id=5 'S10 Body5' ph(idx=2) tx=Y 'S10 body5'"]),
    ("L2b S11 a picture placeholder's picture into content, own size, centred", "LD",
     [ph(2, "S11 Title", "title", None, None, text="S11 title"),
      pic(3, "S11 Picture", "rId2", placeholder=PICTURE_PH)], ["LA"],
     ["sp id=2 'S11 Title' ph(type=title) tx=Y 'S11 title'",
      "pic id=3 'S11 Picture' ph(idx=1) xfrm(479,309,2,2) tx=N ''"]),
    ("L2b S12 a picture into a picture placeholder, cropped to fill", "LD",
     [ph(2, "S12 Title", "title", None, None, text="S12 title"),
      pic(3, "S12 Picture", "rId2", placeholder=PICTURE_PH)], ["LS"],
     ["sp id=2 'S12 Title' ph(type=title) tx=Y 'S12 title'",
      "pic id=3 'S12 Picture' ph(type=pic,idx=1) tx=N ''"]),
    ("L2b S21 an orphan beside a matched title", "LA", la(21, bbox=(300, 200, 200, 100)),
     ["LF"],
     ["sp id=2 'S21 Title' ph(type=title) tx=Y 'S21 title'",
      f"sp id=3 'S21 Body' ph(idx={ORPHAN_IDX}) xfrm(0,200,200,100) tx=Y 'S21 body'"]),
    ("L2b S22 an orphan right of centre", "LA", la(22, bbox=(500, 200, 200, 100)), ["LF"],
     ["sp id=2 'S22 Title' ph(type=title) tx=Y 'S22 title'",
      f"sp id=3 'S22 Body' ph(idx={ORPHAN_IDX}) xfrm(760,200,200,100) tx=Y 'S22 body'"]),
]

A = lambda n: rect(8, f"S{n} RectA", (700, 450, 200, 60), "A")   # noqa: E731
B = lambda n: rect(9, f"S{n} RectB", (700, 380, 200, 60), "B")   # noqa: E731
T = lambda n: ph(2, f"S{n} Title", "title", None, None, text=f"S{n} title")  # noqa: E731


def body(n, kind=None, extra=""):
    return ph(3, f"S{n} Body", kind, 1, None, text=f"S{n} body", extra=extra)


CASES += [
    ("L2c S1 the block goes where the first placeholder was", "LA",
     [A(1), T(1), B(1), body(1)], ["LI"],
     ["sp id=8 'S1 RectA' xfrm(700,450,200,60) tx=Y 'A'",
      "sp id=3 'S1 Body' ph(idx=1) tx=Y 'S1 body'",
      "sp id=2 'S1 Title' ph(type=title) tx=Y 'S1 title'",
      "sp id=9 'S1 RectB' xfrm(700,380,200,60) tx=Y 'B'"]),
    ("L2c S2 shapes between placeholders move after the block", "LA",
     [T(2), A(2), body(2), B(2)], ["LI"],
     ["sp id=3 'S2 Body' ph(idx=1) tx=Y 'S2 body'",
      "sp id=2 'S2 Title' ph(type=title) tx=Y 'S2 title'",
      "sp id=8 'S2 RectA' xfrm(700,450,200,60) tx=Y 'A'",
      "sp id=9 'S2 RectB' xfrm(700,380,200,60) tx=Y 'B'"]),
    ("L2c S3 a quarter-size placeholder first: idx decides", "LA", [T(3), body(3)], ["LT"],
     ["sp id=2 'S3 Title' ph(type=title) tx=Y 'S3 title'",
      "sp id=3 'S3 Body' ph(idx=1) tx=Y 'S3 body'",
      "sp id=4 'LT Full2' ph(idx=2) tx=Y ''"]),
    ("L2c S4 the slide's sz plays no part and goes", "LA",
     [T(4), body(4, extra=' sz="quarter"')], ["LT"],
     ["sp id=2 'S4 Title' ph(type=title) tx=Y 'S4 title'",
      "sp id=3 'S4 Body' ph(idx=1) tx=Y 'S4 body'",
      "sp id=4 'LT Full2' ph(idx=2) tx=Y ''"]),
    ("L2c S5 body prefers body over a lower-idx content", "LA", [T(5), body(5, "body")],
     ["LU"],
     ["sp id=2 'S5 Title' ph(type=title) tx=Y 'S5 title'",
      "sp id=4 'LU Obj1' ph(idx=1) tx=Y ''",
      "sp id=3 'S5 Body' ph(type=body,idx=2) tx=Y 'S5 body'"]),
    ("L2c S6 content passes a picture placeholder", "LA", [T(6), body(6)], ["LV"],
     ["sp id=2 'S6 Title' ph(type=title) tx=Y 'S6 title'",
      "sp id=4 'LV Pic1' ph(type=pic,idx=1) tx=N ''",
      "sp id=3 'S6 Body' ph(idx=2) tx=Y 'S6 body'"]),
    ("L2c S7 back to the layout's order, other shapes after", "LI",
     [body(7), A(7), T(7)], ["LA"],
     ["sp id=2 'S7 Title' ph(type=title) tx=Y 'S7 title'",
      "sp id=3 'S7 Body' ph(idx=1) tx=Y 'S7 body'",
      "sp id=8 'S7 RectA' xfrm(700,450,200,60) tx=Y 'A'"]),
]

#: L2b S13-S20: an orphan's x snaps to the side its centre is on (centre at half: left).
SNAPS = [((300, 200, 200, 100), 0), ((380, 200, 200, 100), 0), ((500, 200, 200, 100), 760),
         ((0, 200, 200, 100), 0), ((760, 200, 200, 100), 760), ((100, 200, 760, 100), 0),
         ((-50, 200, 200, 100), 0), ((900, 200, 200, 100), 760)]


@pytest.mark.parametrize("case", CASES, ids=[c[0] for c in CASES])
def test_powerpoints_layout_change(case):
    _, start, shapes, targets, expected = case
    deck, slide, _ = moved(start, shapes, *targets)
    assert dump(slide) == expected
    assert slide.layout.name == targets[-1]
    assert check(Document.open(deck.to_bytes()).package) == []


@pytest.mark.parametrize("box,x", SNAPS)
def test_an_orphan_snaps_to_the_side_its_centre_is_on(box, x):
    _, slide, change = moved("LA", la(13, title=False, bbox=box), "LR")
    assert dump(slide) == [
        f"sp id=3 'S13 Body' ph(idx={ORPHAN_IDX}) xfrm({x},{box[1]},{box[2]},{box[3]}) tx=Y "
        "'S13 body'"]
    assert change.orphans == [("256.3", "no content or text placeholder left on L19 'LR'")]
    assert change.removed == ["256.2"]


def test_what_is_kept_on_a_matched_placeholder():
    """L2b S7: fill, outline, bodyPr, paragraph and run formatting, name and id stay; only
    the box goes."""
    styled = ('<p:sp><p:nvSpPr><p:cNvPr id="2" name="S7 Title"/><p:cNvSpPr><a:spLocks '
              'noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr>'
              '<a:xfrm><a:off x="1270000" y="508000"/><a:ext cx="6350000" cy="762000"/></a:xfrm>'
              '<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill><a:ln w="38100"><a:solidFill>'
              '<a:srgbClr val="C00000"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr '
              'anchor="b" lIns="0"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr '
              'algn="r"/><a:r><a:rPr lang="en-US" sz="3000" b="1" dirty="0"><a:solidFill>'
              '<a:srgbClr val="0070C0"/></a:solidFill><a:latin typeface="Courier New"/></a:rPr>'
              '<a:t>S7 title</a:t></a:r></a:p></p:txBody></p:sp>')
    deck, slide, change = moved("LA", [styled, ph(3, "S7 Body", None, 1, None,
                                                  text="S7 body")], "LC")
    from lxml import etree

    title_ = slide.shape("256.2")
    xml = etree.tostring(title_._element).decode()
    expected = styled.replace('<a:xfrm><a:off x="1270000" y="508000"/><a:ext cx="6350000" '
                              'cy="762000"/></a:xfrm>', "")
    assert xml.split(">", 1)[1] == expected.split(">", 1)[1]
    assert change.boxes_dropped == [("256.2", (1270000, 508000, 6350000, 762000))]
    assert title_.bounds == (50 * PT, 20 * PT, 860 * PT, 80 * PT)     # the layout's now


def test_a_picture_placeholder_again_on_its_layout_is_cropped_to_fill():
    """B2: re-applying the layout drops a picture placeholder's own box and writes the crop
    for the layout's box -- empty when the image has the box's shape."""
    image = png(56, 38)                                        # the LD box is 560 x 380 pt
    shapes = [ph(2, "Title", "title", None, None, text="Photo"),
              pic(3, "Photo", "rId2", (500, 150, 400, 300), placeholder=PICTURE_PH)]
    deck = build([("LD", shapes)], media=image)
    slide = deck.slides[0]
    change = slide.set_layout("LD")
    picture = slide.shape("256.3")
    rect_ = picture._element.find(".//{%s}srcRect" % "http://schemas.openxmlformats.org/"
                                  "drawingml/2006/main")
    assert rect_ is not None and dict(rect_.attrib) == {}
    assert change.boxes_dropped == [("256.3", (500 * PT, 150 * PT, 400 * PT, 300 * PT))]
    assert picture.bounds == (200 * PT, 120 * PT, 560 * PT, 380 * PT)


def test_a_picture_with_a_crop_of_its_own_keeps_it_and_says_so():
    shapes = [ph(2, "Title", "title", None, None, text="Photo"),
              pic(3, "Photo", "rId2", placeholder=PICTURE_PH, crop='l="10000" r="5000"')]
    _, slide, change = moved("LD", shapes, "LS")
    rect_ = next(slide.shape("256.3")._element.iter(
        "{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect"))
    assert dict(rect_.attrib) == {"l": "10000", "r": "5000"}
    assert change.crops_kept == ["256.3"]


def test_a_large_picture_into_content_is_made_to_fit():
    image = png(1600, 800)                                     # 800 x 400 pt at 144 dpi
    shapes = [ph(2, "Title", "title", None, None, text="Photo"),
              pic(3, "Photo", "rId2", placeholder=PICTURE_PH)]
    deck = build([("LD", shapes)], media=image)
    slide = deck.slides[0]
    slide.set_layout("LC")                                     # LC Left: 420 x 380 pt
    x, y, w, h = slide.shape("256.3").bounds
    assert (w, h) == (420 * PT, 210 * PT)
    assert (x, y + h / 2) == (50 * PT, (120 + 190) * PT)


def test_only_the_layout_relationship_changes():
    deck = build([("LA", la(6, extra=[pic(9, "S6 Image", "rId2", (700, 400, 100, 100))]))])
    slide = deck.slides[0]
    before = {rid: (rel.type, rel.target_part)
              for rid, rel in deck.package.relationships(slide.part_path).items()}
    slide.set_layout("LC")
    after = {rid: (rel.type, rel.target_part)
             for rid, rel in deck.package.relationships(slide.part_path).items()}
    assert set(after) == set(before)
    assert after["rId2"] == before["rId2"]
    assert after["rId1"][1] == deck.layout("LC").part_path


def test_one_undo_step():
    deck, slide, _ = moved("LA", la(2, **OV))
    before = dump(slide)
    slide.set_layout("LC")
    assert deck.undo()
    slide = deck.slides[0]
    assert dump(slide) == before and slide.layout.name == "LA"


def test_a_placeholder_added_to_the_layout_takes_content():
    """L1 and L2 together: content moves onto a placeholder added with add_placeholder."""
    deck = build([("LA", la(1)), ("LF", [ph(2, "T", "title", None, None, text="T")])])
    added = deck.layout("LF").add_placeholder("body", 50 * PT, 320 * PT, 860 * PT,
                                              100 * PT, name="Note")
    change = deck.slides[0].set_layout("LF")
    assert change.matched[1] == ("256.3", added.id)
    assert dump(deck.slides[0])[1] == "sp id=3 'S1 Body' ph(type=body,idx=13) tx=Y 'S1 body'"


def test_overrides_of_size_remain_after_the_change():
    """A4 keeps the slide's run sizes; L4's check still lists them against the new layout."""
    _, slide, _ = moved("LA", la(2, **OV), "LC")
    found = {o.address: o.set for o in slide.size_overrides()}
    assert found == {"256.2": (40.0,), "256.3": (12.0,)}


# -- the tool ---------------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as box:
        yield box


def test_the_tool_reports_orphans_and_dropped_boxes(toolbox):
    deck = build([("LA", la(8, **OV))])
    session = toolbox.session()
    doc = session.open(deck.to_bytes(), "d.pptx")
    result = toolbox.dispatch(session, "ppt_manage_slides",
                              {"doc": doc, "action": "set_layout", "slide": "s:256",
                               "layout": "LG"})
    assert result.ok, result.error
    data = result.data
    assert data["layout"].endswith("(LG)")
    assert data["orphans"] == [{"shape": "256.2",
                                "why": "no title placeholder left on L8 'LG'"}]
    assert data["boxes_dropped"] == [{"shape": "256.3",
                                      "box": {"x": 100, "y": 200, "w": 400, "h": 200}}]
    assert "PowerPoint" in data["note"]
    missing = toolbox.dispatch(session, "ppt_manage_slides",
                               {"doc": doc, "action": "set_layout", "slide": "s:256",
                                "layout": "Nope"})
    assert not missing.ok and missing.error.code == "not_found"
