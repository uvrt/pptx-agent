"""A new deck (E5): from nothing at any size, and from a template.

The numbers pinned below were measured on PowerPoint for Mac 16: a new blank presentation
made through its AppleScript and saved, at 16:9, switched to 4:3 and to a banner, and a deck
of a slide per layout with text in every placeholder (see :mod:`pptx_agent.edit.blank`).
The deck PowerPoint wrote is not in the repository -- it is PowerPoint's own template --
only what was read from it.

Every new deck goes through the gates every edit goes through: the validity checks (schema
order and orphans included) on every part, save and reopen with nothing lost, undo back to
the new deck byte for byte, the full-state SVG of every slide applied to a fresh copy
changing no byte, and a pptx2svg render that puts each placeholder where the master and
layout do, in their fonts and sizes.
"""

from __future__ import annotations

import io
import os
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest
from lxml import etree

import oracle as oracle_helper
from pptx_agent import Document
from pptx_agent.edit import blank
from pptx_agent.oxml.xml import find, findall, qn
from test_validity import (
    _PNG,
    assert_valid,
    check_content_types,
    check_no_dangling_relationships,
    check_slide_list,
    check_well_formed,
    check_zip,
    orphaned_parts,
    requires_powerpoint,
)

CREATED = datetime(2026, 10, 3, 9, 40, 41, tzinfo=timezone.utc)

#: Layout -> (type, [(ph type, idx, sz, orient), ...]) as PowerPoint wrote them.
MEASURED_LAYOUTS = {
    "Title Slide": ("title", [("ctrTitle", None, None, None), ("subTitle", 1, None, None)]),
    "Title and Content": ("obj", [("title", None, None, None), (None, 1, None, None)]),
    "Section Header": ("secHead", [("title", None, None, None), ("body", 1, None, None)]),
    "Two Content": ("twoObj", [("title", None, None, None), (None, 1, "half", None),
                               (None, 2, "half", None)]),
    "Comparison": ("twoTxTwoObj", [("title", None, None, None), ("body", 1, None, None),
                                   (None, 2, "half", None), ("body", 3, "quarter", None),
                                   (None, 4, "quarter", None)]),
    "Title Only": ("titleOnly", [("title", None, None, None)]),
    "Blank": ("blank", []),
    "Content with Caption": ("objTx", [("title", None, None, None), (None, 1, None, None),
                                       ("body", 2, "half", None)]),
    "Picture with Caption": ("picTx", [("title", None, None, None), ("pic", 1, None, None),
                                       ("body", 2, "half", None)]),
    "Title and Vertical Text": ("vertTx", [("title", None, None, None),
                                           ("body", 1, None, "vert")]),
    "Vertical Title and Text": ("vertTitleAndTx", [("title", None, None, "vert"),
                                                   ("body", 1, None, "vert")]),
}
FOOTERS = [("dt", 10, "half", None), ("ftr", 11, "quarter", None),
           ("sldNum", 12, "quarter", None)]

#: Where each placeholder of a new slide is drawn, 16:9: (layout, ph type) -> frame.
MEASURED_FRAMES = {
    ("Title Slide", "ctrTitle"): (1524000, 1122363, 9144000, 2387600),
    ("Title Slide", "subTitle"): (1524000, 3602038, 9144000, 1655762),
    ("Title and Content", "title"): (838200, 365125, 10515600, 1325563),
    ("Title and Content", None): (838200, 1825625, 10515600, 4351338),
    ("Section Header", "title"): (831850, 1709738, 10515600, 2852737),
    ("Section Header", "body"): (831850, 4589463, 10515600, 1500187),
    ("Comparison", "title"): (839788, 365125, 10515600, 1325563),
    ("Content with Caption", "title"): (839788, 457200, 3932237, 1600200),
    ("Content with Caption", None): (5183188, 987425, 6172200, 4873625),
    ("Picture with Caption", "pic"): (5183188, 987425, 6172200, 4873625),
    ("Vertical Title and Text", "title"): (8724900, 365125, 2628900, 5811838),
    ("Vertical Title and Text", "body"): (838200, 365125, 7734300, 5811838),
}


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _root(data: bytes, name: str):
    return etree.fromstring(_parts(data)[name])


def _placeholders(root) -> list[tuple]:
    found = []
    for ph in root.iter(qn("p:ph")):
        idx = ph.get("idx")
        found.append((ph.get("type"), None if idx is None else int(idx), ph.get("sz"),
                      ph.get("orient")))
    return found


def all_order_violations(data: bytes) -> set:
    """Schema-order violations in every XML part, not only slides."""
    from pptx_agent.core.xml import _ranks, prefixed_name

    found = set()
    for name, raw in _parts(data).items():
        if not name.endswith(".xml") or name == "[Content_Types].xml":
            continue
        for parent in etree.fromstring(raw).iter():
            ranks = _ranks(prefixed_name(parent))
            if not ranks:
                continue
            last = -1
            for child in parent:
                rank = ranks.get(prefixed_name(child), len(ranks) + 1)
                if rank < last:
                    found.add((name, prefixed_name(parent), prefixed_name(child)))
                last = max(last, rank)
    return found


def assert_new_deck_valid(data: bytes) -> None:
    check_well_formed(data)
    check_no_dangling_relationships(data)
    check_content_types(data, data)
    assert orphaned_parts(data) == set()
    check_slide_list(data)
    assert all_order_violations(data) == set()
    check_zip(data)


# ------------------------------------------------------------------------------------------
# From nothing: what PowerPoint makes for File > New
# ------------------------------------------------------------------------------------------


def test_a_new_deck_offers_powerpoints_layouts():
    document = Document.new(created=CREATED)
    assert document.slides == [] and not document.history.can_undo()
    assert [layout.name for layout in document.layouts] == list(MEASURED_LAYOUTS)
    for layout in document.layouts:
        root = document.package.tree(layout.part_path)
        kind, placeholders = MEASURED_LAYOUTS[layout.name]
        assert root.get("type") == kind and root.get("preserve") == "1"
        assert _placeholders(root) == placeholders + FOOTERS, layout.name


def test_the_package_is_the_one_powerpoint_writes():
    data = Document.new(created=CREATED).to_bytes()
    assert sorted(_parts(data)) == sorted(
        ["[Content_Types].xml", "_rels/.rels", "docProps/app.xml", "docProps/core.xml",
         "ppt/_rels/presentation.xml.rels", "ppt/presentation.xml", "ppt/presProps.xml",
         "ppt/viewProps.xml", "ppt/tableStyles.xml", "ppt/theme/theme1.xml",
         "ppt/slideMasters/slideMaster1.xml", "ppt/slideMasters/_rels/slideMaster1.xml.rels"]
        + [f"ppt/slideLayouts/slideLayout{n}.xml" for n in range(1, 12)]
        + [f"ppt/slideLayouts/_rels/slideLayout{n}.xml.rels" for n in range(1, 12)])
    presentation = _root(data, "ppt/presentation.xml")
    size = find(presentation, "p:sldSz")
    assert (size.get("cx"), size.get("cy"), size.get("type")) == ("12192000", "6858000", None)
    notes = find(presentation, "p:notesSz")
    assert (notes.get("cx"), notes.get("cy")) == ("6858000", "9144000")
    assert find(presentation, "p:sldMasterIdLst/p:sldMasterId").get("id") == "2147483648"
    assert find(presentation, "p:sldIdLst") is None
    default = find(presentation, "p:defaultTextStyle")
    assert [find(default, f"a:lvl{n}pPr/a:defRPr").get("sz") for n in range(1, 10)] == \
        ["1800"] * 9
    assert [find(default, f"a:lvl{n}pPr").get("marL") for n in range(1, 10)] == \
        [str(457200 * n) for n in range(9)]
    assert _root(data, "ppt/tableStyles.xml").get("def") == \
        "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"
    master = _root(data, "ppt/slideMasters/slideMaster1.xml")
    assert [node.get("id") for node in findall(master, "p:sldLayoutIdLst/p:sldLayoutId")] == \
        [str(2147483649 + n) for n in range(11)]


def test_the_theme_is_the_office_theme():
    theme = _root(Document.new().to_bytes(), "ppt/theme/theme1.xml")
    colours = {}
    for slot in find(theme, "a:themeElements/a:clrScheme"):
        node = slot[0]
        colours[etree.QName(slot).localname] = node.get("lastClr") or node.get("val")
    assert colours == {"dk1": "000000", "lt1": "FFFFFF", "dk2": "0E2841", "lt2": "E8E8E8",
                       "accent1": "156082", "accent2": "E97132", "accent3": "196B24",
                       "accent4": "0F9ED5", "accent5": "A02B93", "accent6": "4EA72E",
                       "hlink": "467886", "folHlink": "96607D"}
    fonts = find(theme, "a:themeElements/a:fontScheme")
    assert find(fonts, "a:majorFont/a:latin").get("typeface") == "Aptos Display"
    assert find(fonts, "a:minorFont/a:latin").get("typeface") == "Aptos"
    assert len(findall(fonts, "a:majorFont/a:font")) == 47
    styles = find(theme, "a:themeElements/a:fmtScheme")
    assert [len(find(styles, tag)) for tag in ("a:fillStyleLst", "a:lnStyleLst",
                                               "a:effectStyleLst", "a:bgFillStyleLst")] == \
        [3, 3, 3, 3]
    assert [line.get("w") for line in findall(styles, "a:lnStyleLst/a:ln")] == \
        ["12700", "19050", "25400"]


def _styles(master) -> dict:
    styles = {}
    for name in ("titleStyle", "bodyStyle", "otherStyle"):
        holder = find(master, f"p:txStyles/p:{name}")
        levels = []
        for n in range(1, 10):
            level = find(holder, f"a:lvl{n}pPr")
            if level is None:
                break
            spacing = find(level, "a:spcBef/a:spcPts")
            levels.append((find(level, "a:defRPr").get("sz"), level.get("marL"),
                           level.get("indent"), level.get("defTabSz"),
                           None if spacing is None else spacing.get("val")))
        styles[name] = levels
    return styles


def test_the_master_text_styles_are_powerpoints():
    master = _root(Document.new().to_bytes(), "ppt/slideMasters/slideMaster1.xml")
    styles = _styles(master)
    assert styles["titleStyle"] == [("4400", None, None, "914400", None)]
    assert styles["bodyStyle"] == [
        (size, str(228600 + 457200 * n), "-228600", "914400", before) for n, (size, before) in
        enumerate([("2800", "1000"), ("2400", "500"), ("2000", "500")]
                  + [("1800", "500")] * 6)]
    assert styles["otherStyle"] == [("1800", str(457200 * n), None, "914400", None)
                                    for n in range(9)]
    title = find(master, "p:txStyles/p:titleStyle/a:lvl1pPr")
    assert find(title, "a:lnSpc/a:spcPct").get("val") == "90000"
    assert find(title, "a:defRPr/a:latin").get("typeface") == "+mj-lt"
    body = find(master, "p:txStyles/p:bodyStyle/a:lvl1pPr")
    assert find(body, "a:buChar").get("char") == "•"
    assert find(body, "a:buFont").get("typeface") == "Arial"
    frames = {}
    for shape in findall(master, "p:cSld/p:spTree/p:sp"):
        off, ext = find(shape, "p:spPr/a:xfrm/a:off"), find(shape, "p:spPr/a:xfrm/a:ext")
        frames[find(shape, "p:nvSpPr/p:nvPr/p:ph").get("type")] = tuple(
            int(v) for v in (off.get("x"), off.get("y"), ext.get("cx"), ext.get("cy")))
    assert frames == blank.MASTER_FRAMES


def test_placeholders_on_a_new_slide_are_where_powerpoint_draws_them():
    document = Document.new()
    for (name, kind), frame in MEASURED_FRAMES.items():
        slide = document.add_slide(name)
        shape = next(s for s in slide.shapes if s.placeholder[0] == kind)
        assert not shape.has_explicit_transform
        assert (shape.left, shape.top, shape.width, shape.height) == frame, (name, kind)


#: Master and layout frames and text sizes PowerPoint wrote at 4:3 (x by 0.75, y kept).
MEASURED_4X3 = {
    "master": {"title": (628650, 365126, 7886700, 1325563),
               "body": (628650, 1825625, 7886700, 4351338),
               "dt": (628650, 6356351, 2057400, 365125),
               "ftr": (3028950, 6356351, 3086100, 365125),
               "sldNum": (6457950, 6356351, 2057400, 365125)},
    "Comparison": [(629841, 365126, 7886700, 1325563), (629842, 1681163, 3868340, 823912),
                   (629842, 2505075, 3868340, 3684588), (4629150, 1681163, 3887391, 823912),
                   (4629150, 2505075, 3887391, 3684588)],
    "Vertical Title and Text": [(6543675, 365125, 1971675, 5811838),
                                (628650, 365125, 5800725, 5811838)],
}


def _frames(root) -> list[tuple[int, int, int, int]]:
    frames = []
    for shape in findall(root, "p:cSld/p:spTree/p:sp"):
        off, ext = find(shape, "p:spPr/a:xfrm/a:off"), find(shape, "p:spPr/a:xfrm/a:ext")
        if off is not None:
            frames.append(tuple(int(v) for v in (off.get("x"), off.get("y"), ext.get("cx"),
                                                 ext.get("cy"))))
    return frames


def _close(a, b, tolerance: int = 2) -> bool:
    return len(a) == len(b) and all(abs(x - y) <= tolerance for x, y in zip(a, b))


def test_a_4x3_deck_is_laid_out_as_powerpoint_lays_it_out():
    document = Document.new(size="4:3")
    data = document.to_bytes()
    size = find(_root(data, "ppt/presentation.xml"), "p:sldSz")
    assert (size.get("cx"), size.get("cy"), size.get("type")) == ("9144000", "6858000",
                                                                  "screen4x3")
    master = _root(data, "ppt/slideMasters/slideMaster1.xml")
    measured = MEASURED_4X3["master"]
    for frame, kind in zip(_frames(master), ("title", "body", "dt", "ftr", "sldNum")):
        assert _close(frame, measured[kind]), kind
    styles = _styles(master)
    assert [level[0] for level in styles["bodyStyle"]] == \
        ["2100", "1800", "1500"] + ["1350"] * 6
    assert styles["titleStyle"][0][0] == "3300"
    assert styles["bodyStyle"][0][1:] == ("171450", "-171450", "685800", "750")
    assert styles["otherStyle"][1][:2] == ("1350", "342900")
    # The presentation's default text style keeps its 16:9 sizes, as PowerPoint's does.
    default = find(_root(data, "ppt/presentation.xml"), "p:defaultTextStyle/a:lvl1pPr/a:defRPr")
    assert default.get("sz") == "1800"
    for name in ("Comparison", "Vertical Title and Text"):
        root = document.package.tree(document.layout(name).part_path)
        assert all(_close(a, b) for a, b in zip(_frames(root), MEASURED_4X3[name]))
    title = document.package.tree(document.layout("Title Slide").part_path)
    sizes = [node.get("sz") for node in title.iter(qn("a:defRPr"))]
    assert sizes[:3] == ["4500", "1800", "1500"]


def test_any_size_is_laid_out_by_powerpoints_rule():
    """A banner (8 x 1 in) as PowerPoint laid it out: frames x0.6 across, x0.1333 down,
    text by the smaller ratio, rounded."""
    document = Document.new(size=(7315200, 914400))
    master = document.package.tree(document.layouts[0].master_part)
    assert _close(_frames(master)[0], (502920, 48683, 6309360, 176742))
    sizes = sorted({node.get("sz") for node in master.iter(qn("a:defRPr")) if node.get("sz")},
                   key=int)
    assert sizes == ["160", "240", "267", "320", "373", "587"]
    assert find(document.package.tree(document.package.presentation_part()),
                "p:sldSz").get("type") == "banner"


@pytest.mark.parametrize("size", ["16:10", (100, 100), (914400, 60000000), "big", None, 3])
def test_a_slide_size_out_of_range_is_refused(size):
    if size is None:
        assert Document.new(size=None).slide_size == (12192000, 6858000)
        return
    with pytest.raises(ValueError):
        Document.new(size=size)


@pytest.mark.parametrize("size", ["16:9", "4:3", (9144000, 5143500)])
def test_a_new_deck_is_valid(size):
    assert_new_deck_valid(Document.new(size=size).to_bytes())


# ------------------------------------------------------------------------------------------
# The round trip: a whole deck built from nothing
# ------------------------------------------------------------------------------------------


def build(document: Document) -> dict:
    """Slides on several layouts with every placeholder filled, then shapes, connectors, a
    table and a picture; and E5's acceptance slide.  Returns what to read back."""
    from authoring import acceptance_edits

    expected: dict = {"texts": {}}
    names = ("Title Slide", "Title and Content", "Two Content", "Comparison",
             "Picture with Caption", "Section Header", "Blank")
    available = {layout.name for layout in document.layouts}
    layouts = list(names) if set(names) <= available else document.layouts[:7]
    for layout in layouts:
        slide = document.add_slide(layout)
        name = getattr(layout, "name", layout)
        for number, shape in enumerate(slide.shapes):
            if shape.placeholder[0] == "pic":
                continue
            text = f"{name} {number}".strip()
            shape.set_text(text if shape.placeholder[0] in ("title", "ctrTitle")
                           else f"{text}\nsecond line")
            expected["texts"][shape.id] = shape.text
    canvas = document.add_slide("Title Only" if "Title Only" in available else None)
    titles = [s for s in canvas.shapes if s.placeholder and s.placeholder[0] in ("title",
                                                                                  "ctrTitle")]
    if titles:
        titles[0].set_text("Canvas")
    left = canvas.add_shape("roundRect", 914400, 1828800, 2286000, 1143000, text="Left",
                            adjustments={"adj": 25000}, fill="accent2")
    right = canvas.add_shape("ellipse", 7315200, 3200400, 1600200, 1143000, text="Right")
    elbow = canvas.add_connector("elbow", (left, 3), (right, 1), line={"tail": "triangle"})
    table = canvas.add_table(2, 3, 914400, 4572000, 5486400, 741680).table
    table.cell(0, 0).text = "Quarter"
    table.cell(1, 2).text = "4,285"
    picture = canvas.add_picture(_PNG, 9144000, 914400, width=1828800)
    right.move_by(dy=-457200)
    expected.update(canvas=canvas.slide_id, connector=elbow.id, picture=picture.id,
                    table=canvas.shapes[-2].id, left=left.id, right=right.id)
    expected["acceptance"] = acceptance_edits(document)
    document.title = "Built from nothing"
    document.author = "pptx-agent tests"
    return expected


def test_a_deck_built_from_nothing_round_trips():
    document = Document.new(created=CREATED)
    expected = build(document)
    saved = document.to_bytes()
    assert_new_deck_valid(saved)
    reopened = Document.open(saved)
    assert reopened.to_bytes() == saved
    assert len(reopened.slides) == 9
    for identifier, text in expected["texts"].items():
        assert reopened.shape(identifier).text == text
    connector = reopened.shape(expected["connector"])
    assert connector.begin_connection[0].id == expected["left"]
    assert connector.end_connection[0].id == expected["right"]
    assert reopened.shape(expected["table"]).table.cell(1, 2).text == "4,285"
    assert reopened.shape(expected["picture"]).image_part.startswith("ppt/media/")
    assert (reopened.title, reopened.author) == ("Built from nothing", "pptx-agent tests")
    assert reopened.created == CREATED


def test_a_deck_built_from_nothing_survives_its_full_state_svgs():
    pytest.importorskip("pptx2svg")
    document = Document.new(created=CREATED)
    build(document)
    saved = document.to_bytes()
    reopened = Document.open(saved)
    for slide in reopened.slides:
        copy = Document.open(saved)
        report = copy.apply_svg(slide.render_svg(full_state=True))
        assert not report, slide
        assert copy.to_bytes() == saved


def test_creating_the_deck_is_the_base_state():
    document = Document.new(created=CREATED)
    document.history._max_depth = 10_000
    fresh = document.to_bytes()
    assert not document.undo()
    build(document)
    edited = document.to_bytes()
    while document.undo():
        pass
    assert document.to_bytes() == fresh and document.slides == []
    while document.redo():
        pass
    assert document.to_bytes() == edited


def test_new_decks_are_deterministic():
    assert Document.new(created=CREATED).to_bytes() == Document.new(created=CREATED).to_bytes()


# ------------------------------------------------------------------------------------------
# Rendering: placeholders inherit from the master and layout
# ------------------------------------------------------------------------------------------


def _drawn(svg: str, identifier: str) -> tuple[tuple[float, float], list[str], list[str]]:
    root = etree.fromstring(svg.encode())
    group = next(g for g in root.iter("{http://www.w3.org/2000/svg}g")
                 if g.get("data-pptx-id") == identifier)
    x, y = (float(v) for v in re.fullmatch(r"translate\(([-\d.]+),\s*([-\d.]+)\)",
                                           group.get("transform")).groups())
    spans = [s for s in group.iter("{http://www.w3.org/2000/svg}tspan") if (s.text or "").strip()
             and s.text != "•"]
    return (x, y), [s.get("font-size") for s in spans], [s.get("font-family") for s in spans]


@pytest.mark.parametrize("size", ["16:9", "4:3"])
def test_pptx2svg_draws_placeholders_from_the_master_and_layout(size):
    """Each placeholder where its layout or the master puts it, in the theme's fonts at the
    master's or layout's sizes: titles 44 pt Aptos Display (60 pt on the title slide),
    body text 28 pt then 24 pt Aptos -- at 4:3, PowerPoint's 0.75 of each."""
    pytest.importorskip("pptx2svg")
    document = Document.new(size=size)
    scale = 1.0 if size == "16:9" else 0.75
    width = document.slide_size[0]
    px = 1280 / width if size == "16:9" else 960 / width
    cases = [("Title Slide", "ctrTitle", ["60"]), ("Title and Content", "title", ["44"]),
             ("Title and Content", None, ["28", "24"]), ("Section Header", "title", ["60"]),
             ("Comparison", "body", ["24"])]
    for layout, kind, points in cases:
        slide = document.add_slide(layout)
        shape = next(s for s in slide.shapes if s.placeholder[0] == kind)
        # Two lines where the box has room for them; the comparison heads auto-fit two.
        shape.set_text("One\nTwo" if len(points) == 2 else "Heading")
        if len(points) == 2:
            shape.paragraphs[1].level = 1
        options = {} if size == "16:9" else {"width": 960}
        (x, y), sizes, families = _drawn(slide.render_svg(**options), shape.id)
        assert abs(x - shape.left * px) < 0.5 and abs(y - shape.top * px) < 0.5, layout
        want = [round(float(p) * scale * 96 / 72 * (px * 914400 / 96), 2) for p in points]
        assert [round(float(s), 2) for s in sizes][:len(want)] == want, (layout, kind)
        face = "Aptos Display" if kind in ("title", "ctrTitle") else "Aptos"
        assert all(f.split(",")[0].strip("'") == face for f in families), families


# ------------------------------------------------------------------------------------------
# Metadata, language and app.xml
# ------------------------------------------------------------------------------------------


def test_core_properties_are_written_and_edited():
    document = Document.new(title="Plan", author="Ada", created=CREATED)
    core = _root(document.to_bytes(), "docProps/core.xml")
    ns = {"dc": "http://purl.org/dc/elements/1.1/", "dcterms": "http://purl.org/dc/terms/",
          "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"}
    assert core.findtext("dc:title", namespaces=ns) == "Plan"
    assert core.findtext("dc:creator", namespaces=ns) == "Ada"
    assert core.findtext("dcterms:created", namespaces=ns) == "2026-10-03T09:40:41Z"
    assert core.findtext("cp:revision", namespaces=ns) == "1"
    assert (document.title, document.author, document.created, document.modified) == \
        ("Plan", "Ada", CREATED, CREATED)
    before = document.to_bytes()
    later = datetime(2026, 10, 4, 12, 0, tzinfo=timezone.utc)
    document.title = "Plan B"
    document.modified = later
    assert Document.open(document.to_bytes()).title == "Plan B"
    assert Document.open(document.to_bytes()).modified == later
    document.undo()
    document.undo()
    assert document.to_bytes() == before
    with pytest.raises(TypeError):
        document.modified = "yesterday"


def test_a_deck_without_core_properties_gets_them(financial_report):
    document = Document.open(financial_report.read_bytes())
    had = document.package.has_part("docProps/core.xml")
    document.title = "Annual"
    reopened = Document.open(document.to_bytes())
    assert reopened.title == "Annual"
    assert_valid(document.to_bytes(), financial_report.read_bytes())
    if not had:
        document.undo()
        assert not document.package.has_part("docProps/core.xml")


def test_the_language_is_the_decks_and_new_shapes_take_it():
    document = Document.new(language="nl-BE")
    assert document.language == "nl-BE"
    slide = document.add_slide("Blank")
    box = slide.add_textbox(914400, 914400, 914400, 369332, "Hallo")
    assert 'lang="nl-BE"' in etree.tostring(box._element).decode()
    title = document.add_slide("Title Slide").shapes[0]
    assert 'lang="nl-BE"' in etree.tostring(title._element).decode()
    document.language = "de-DE"
    master = document.package.tree(document.layouts[0].master_part)
    assert find(master, "p:txStyles/p:otherStyle/a:defPPr/a:defRPr").get("lang") == "de-DE"
    assert document.language == "de-DE"
    document.undo()
    assert document.language == "nl-BE"


def _app(data: bytes) -> dict:
    root = _root(data, "docProps/app.xml")
    ns = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
    vt = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
    pairs = [node.text for node in root.iter(vt + "lpstr", vt + "i4")
             if node.getparent().getparent().tag == ns + "HeadingPairs"
             or node.getparent().tag == vt + "variant"]
    return {"slides": root.findtext(ns + "Slides"), "notes": root.findtext(ns + "Notes"),
            "words": root.findtext(ns + "Words"), "paragraphs": root.findtext(ns + "Paragraphs"),
            "pairs": pairs,
            "titles": [n.text for n in root.find(ns + "TitlesOfParts")[0]]}


def test_app_xml_follows_the_slides():
    document = Document.new()
    assert _app(document.to_bytes())["slides"] == "0"
    title = document.add_slide("Title Slide")
    title.shapes[0].set_text("Opening")
    title.shapes[1].set_text("two words")
    document.add_slide("Blank")
    app = _app(document.to_bytes())
    assert app["slides"] == "2" and app["words"] == "3" and app["paragraphs"] == "2"
    assert app["titles"] == ["Aptos", "Aptos Display", "Arial", "Office Theme", "Opening",
                             "PowerPoint Presentation"]
    assert app["pairs"] == ["Fonts Used", "3", "Theme", "1", "Slide Titles", "2"]
    saved = document.to_bytes()
    assert Document.open(saved).to_bytes() == saved  # written once, then left alone


def test_app_xml_follows_a_deleted_slide_in_a_powerpoint_deck():
    from conftest import FIXTURE_DIR

    original = (FIXTURE_DIR / "powerpoint-smartart.pptx").read_bytes()
    document = Document.open(original)
    before = _app(original)
    document.delete_slide(document.slides[-1])
    after = _app(document.to_bytes())
    assert int(after["slides"]) == int(before["slides"]) - 1
    assert len(after["titles"]) == len(before["titles"]) - 1
    document.undo()
    assert document.to_bytes() == Document.open(original).to_bytes()


# ------------------------------------------------------------------------------------------
# Changing the slide size: PowerPoint's "Ensure Fit"
# ------------------------------------------------------------------------------------------


def test_resizing_a_deck_scales_it_as_powerpoint_does():
    """Measured on a deck switched from 16:9 to 4:3 in PowerPoint: shapes on slides scaled
    by 0.75 and centred (y + 857250), group child space kept, table grid scaled, explicit
    sizes scaled, inherited ones written at 13.5 pt, a 24 pt run landing on the 18 pt
    default written without a size; line widths kept."""
    document = Document.new()
    slide = document.add_slide("Blank")
    box = slide.add_shape("rect", 914400, 914400, 1828800, 914400, text="Box",
                          line={"width": 25400})
    box.paragraphs[0].runs[0].size = 24
    oval = slide.add_shape("ellipse", 6096000, 2743200, 1371600, 1371600, text="Oval")
    first = slide.add_shape("rect", 8000000, 500000, 900000, 500000)
    second = slide.add_shape("rect", 9200000, 1200000, 900000, 500000)
    group = slide.group([first, second])
    table = slide.add_table(2, 2, 914400, 5200000, 4000000, 740000)
    picture = slide.add_picture(_PNG, 6000000, 5000000, width=914400)
    before = document.to_bytes()
    document.slide_size = "4:3"
    assert document.slide_size == (9144000, 6858000)
    assert (box.left, box.top, box.width, box.height) == (685800, 1543050, 1371600, 685800)
    assert (oval.left, oval.top, oval.width) == (4572000, 2914650, 1028700)
    assert (group.left, group.top, group.width, group.height) == (6000000, 1232250, 1575000,
                                                                  900000)
    assert group.child_offset == (8000000, 500000)
    assert (table.left, table.top, table.width, table.height) == (685800, 4757250, 3000000,
                                                                  555000)
    assert table.table.column_widths == [1500000, 1500000]
    assert (picture.left, picture.top, picture.width) == (4500000, 4607250, 685800)
    box_xml = etree.tostring(box._element).decode()
    assert '<a:ln w="25400"/>' in box_xml
    assert re.search(r'<a:rPr lang="en-US"/><a:t>Box', box_xml)
    assert 'sz="1350"' in etree.tostring(oval._element).decode()
    assert_new_deck_valid(document.to_bytes())
    document.undo()
    assert document.to_bytes() == before


def test_resizing_without_scaling_changes_only_the_size():
    document = Document.new()
    master = document.package.read(document.layouts[0].master_part)
    document.set_slide_size((9144000, 5143500), scale=False)
    assert document.slide_size == (9144000, 5143500)
    assert document.package.read(document.layouts[0].master_part) == master


# ------------------------------------------------------------------------------------------
# From a template
# ------------------------------------------------------------------------------------------


def test_a_new_deck_from_each_fixture_keeps_its_design_and_drops_its_slides(pptx_path):
    original = pptx_path.read_bytes()
    template = Document.open(original)
    document = Document.new(template=pptx_path, author="Ada", created=CREATED)
    assert document.slides == [] and not document.history.can_undo()
    assert [l.name for l in document.layouts] == [l.name for l in template.layouts]
    assert document.slide_size == template.slide_size
    assert (document.title, document.author, document.created) == ("", "Ada", CREATED)
    data = document.to_bytes()
    check_well_formed(data)
    check_no_dangling_relationships(data)
    check_content_types(data, original)
    assert orphaned_parts(data) <= orphaned_parts(original)
    check_slide_list(data)
    names = set(_parts(data))
    assert not any(name.startswith(("ppt/slides/", "ppt/notesSlides/", "ppt/charts/"))
                   for name in names)
    assert "docProps/thumbnail.jpeg" not in names
    presentation = _root(data, "ppt/presentation.xml")
    assert find(presentation, "p:custShowLst") is None
    assert not [e for e in findall(presentation, "p:extLst/p:ext")
                if e.get("uri") == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"]
    if document.layouts:
        slide = document.add_slide(document.layouts[0])
        assert_valid(document.to_bytes(), data)
        assert Document.open(document.to_bytes()).slides[0].slide_id == slide.slide_id


def test_a_potx_becomes_a_presentation_and_a_deck_saves_as_a_potx(product_page, tmp_path):
    potx = Document.open(product_page.read_bytes()).to_bytes(template=True)
    types = _root(potx, "[Content_Types].xml")
    main = [n.get("ContentType") for n in types if n.get("PartName") == "/ppt/presentation.xml"]
    assert main == [blank.CT_TEMPLATE]
    document = Document.new(template=potx)
    data = document.to_bytes()
    main = [n.get("ContentType") for n in _root(data, "[Content_Types].xml")
            if n.get("PartName") == "/ppt/presentation.xml"]
    assert main == [blank.CT_PRESENTATION]
    document.save_as_template(tmp_path / "design.potx")
    document.save(tmp_path / "again.potx")
    document.save(tmp_path / "deck.pptx")
    for name, wanted in (("design.potx", blank.CT_TEMPLATE), ("again.potx", blank.CT_TEMPLATE),
                         ("deck.pptx", blank.CT_PRESENTATION)):
        saved = (tmp_path / name).read_bytes()
        assert [n.get("ContentType") for n in _root(saved, "[Content_Types].xml")
                if n.get("PartName") == "/ppt/presentation.xml"] == [wanted]


def test_a_template_takes_a_new_size_and_language(product_page):
    template = Document.open(product_page.read_bytes())
    document = Document.new(template=product_page, size="4:3", language="fr-FR")
    assert document.slide_size == (9144000, 6858000)
    assert document.language == "fr-FR"
    assert not document.history.can_undo()
    assert template.slide_size != document.slide_size


def test_a_macro_enabled_template_is_refused(product_page):
    with zipfile.ZipFile(io.BytesIO(Document.open(product_page.read_bytes()).to_bytes())) as src:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as out:
            for info in src.infolist():
                raw = src.read(info.filename)
                if info.filename == "[Content_Types].xml":
                    raw = raw.replace(
                        b"application/vnd.openxmlformats-officedocument.presentationml."
                        b"presentation.main+xml",
                        b"application/vnd.ms-powerpoint.template.macroEnabled.main+xml")
                out.writestr(info, raw)
    with pytest.raises(ValueError):
        Document.new(template=buffer.getvalue())


# ------------------------------------------------------------------------------------------
# PowerPoint (pytest -m oracle)
# ------------------------------------------------------------------------------------------

HOME = Path(os.path.expanduser("~"))


def _new_deck(kind: str, scratch: list[Path]) -> Document:
    from conftest import FIXTURE_DIR

    if kind == "16:9":
        return Document.new(created=CREATED)
    if kind == "4:3":
        return Document.new(size="4:3", created=CREATED)
    if kind == "template-pptx":
        return Document.new(template=FIXTURE_DIR / "sample.pptx", created=CREATED)
    # A .potx PowerPoint wrote itself, from the deck it wrote with a SmartArt diagram.
    source = HOME / "pptx-agent-e5-template-source.pptx"
    potx = HOME / "pptx-agent-e5-template.potx"
    scratch += [source, potx]
    source.write_bytes((FIXTURE_DIR / "powerpoint-smartart.pptx").read_bytes())
    saved = oracle_helper.save_as_potx(source, potx)
    assert saved.ok, f"PowerPoint {saved.outcome}: {saved.detail}"
    return Document.new(template=potx.read_bytes(), created=CREATED)


def _unused_notes_master(data: bytes) -> set[str]:
    """What PowerPoint drops from a deck with no notes (measured): the notes master, what
    only it uses (its theme), and Windows printer settings."""
    from pptx_agent.edit.slides import reachable_from
    from pptx_agent.oxml.package import PresentationPackage, rels_path_for

    package = PresentationPackage.open(data)
    names = set(package.part_names)
    unused = {name for name in names if name.startswith("ppt/printerSettings/")}
    if any(name.startswith("ppt/notesSlides/") for name in names):
        return unused
    for master in (name for name in names if name.startswith("ppt/notesMasters/")
                   and name.endswith(".xml")):
        for part in reachable_from(package, master):
            if part.startswith(("ppt/notesMasters/", "ppt/theme/")):
                unused |= {part, rels_path_for(part)}
    return unused & names


@pytest.mark.oracle
@requires_powerpoint
@pytest.mark.parametrize("kind", ["16:9", "4:3", "template-pptx", "template-potx"])
def test_powerpoint_opens_new_decks(kind):
    """E5's new-deck acceptance.  A deck from nothing at 16:9 and 4:3, and from a template
    (a fixture .pptx, and a .potx PowerPoint saved), built up with slides on several layouts,
    shapes, connectors, a table, a picture and the E5 acceptance slide: PowerPoint exports it
    unprompted with one page per slide and the text in place, and -- saving it again --
    adds no part, keeps every slide, layout, text and connection, and drops nothing but a
    notes master no slide uses (with its theme) and Windows printer settings."""
    scratch: list[Path] = []
    try:
        document = _new_deck(kind, scratch)
        expected = build(document)
        data = document.to_bytes()
        assert_valid(data, data)
        stem = "pptx-agent-e5-new-" + kind.replace(":", "x")
        deck, pdf, again = (HOME / f"{stem}.pptx", HOME / f"{stem}.pdf",
                            HOME / f"{stem}-resaved.pptx")
        scratch += [deck, pdf, again]
        deck.write_bytes(data)
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert oracle_helper.pdf_page_count(pdf) == len(document.slides)
        texts = oracle_helper.pdf_texts(pdf)
        if texts is not None:
            page = expected["acceptance"]["page"]
            for sentinel in ("E5BOX", "E5CELL", "E5KID", "E5FROM"):
                assert sentinel in texts[page], (sentinel, texts[page])
            assert "4,285" in " ".join(texts)

        saved = oracle_helper.save_as_pptx(deck, again)
        assert saved.ok, f"PowerPoint {saved.outcome}: {saved.detail}"
        ours, theirs = _parts(data), _parts(again.read_bytes())
        assert set(theirs) <= set(ours), sorted(set(theirs) - set(ours))
        dropped = set(ours) - set(theirs)
        assert dropped <= _unused_notes_master(data), sorted(dropped)
        resaved = Document.open(again.read_bytes())
        assert [s.slide_id for s in resaved.slides] == [s.slide_id for s in document.slides]
        assert [l.name for l in resaved.layouts] == [l.name for l in document.layouts]
        for identifier in expected["acceptance"]["connectors"] + [expected["connector"]]:
            mine, kept = document.shape(identifier), resaved.shape(identifier)
            assert kept.begin_connection[0].id == mine.begin_connection[0].id
            assert kept.end_connection[0].id == mine.end_connection[0].id
        for identifier, text in expected["texts"].items():
            assert resaved.shape(identifier).text == text
    finally:
        oracle_helper.cleanup(*scratch)


@pytest.mark.oracle
@requires_powerpoint
@pytest.mark.parametrize("size", ["16:9", "4:3"])
def test_a_new_deck_matches_the_one_powerpoint_makes(size):
    """PowerPoint makes a new deck (File > New, switched to 4:3 for the second case) with
    a slide per layout and text in every placeholder; this library makes the same.  Every
    placeholder is within a few EMU of PowerPoint's, the master's text styles and the
    theme's fonts and colours are the same, and the exported pages are the same pixels."""
    stem = "pptx-agent-e5-" + size.replace(":", "x")
    theirs, mine = HOME / f"{stem}-powerpoint.pptx", HOME / f"{stem}-library.pptx"
    pdfs = [theirs.with_suffix(".pdf"), mine.with_suffix(".pdf")]
    try:
        made = oracle_helper.new_powerpoint_deck(theirs, size)
        assert made.ok, made.detail
        document = Document.new(size=size)
        for number, layout in enumerate(document.layouts, start=1):
            slide = document.add_slide(layout)
            for index, shape in enumerate(slide.shapes, start=1):
                if shape.placeholder[0] != "pic":  # PowerPoint cannot type into one
                    shape.set_text(f"S{number}P{index}")
        document.save(mine)
        reference = Document.open(theirs.read_bytes())
        ours = Document.open(mine.read_bytes())
        assert ours.slide_size == reference.slide_size
        # PowerPoint names its layouts in its UI language; their types say which is which.
        assert [ours.package.tree(l.part_path).get("type") for l in ours.layouts] == \
            [reference.package.tree(l.part_path).get("type") for l in reference.layouts]
        for theirs_slide, our_slide in zip(reference.slides, ours.slides):
            assert len(theirs_slide.shapes) == len(our_slide.shapes)
            for a, b in zip(theirs_slide.shapes, our_slide.shapes):
                assert a.placeholder == b.placeholder and a.text == b.text
                assert _close((a.left, a.top, a.width, a.height),
                              (b.left, b.top, b.width, b.height)), (a, b)
        theirs_master = reference.package.tree(reference.layouts[0].master_part)
        our_master = ours.package.tree(ours.layouts[0].master_part)
        assert _styles(theirs_master) == _styles(our_master)
        for part in ("a:themeElements/a:clrScheme", "a:themeElements/a:fontScheme/a:majorFont",
                     "a:themeElements/a:fontScheme/a:minorFont"):
            def faces(document):
                theme = document.package.tree("ppt/theme/theme1.xml")
                return [(etree.QName(n).localname, dict(n.attrib)) for n in
                        find(theme, part).iter() if n.attrib]
            assert faces(reference) == faces(ours), part
        for deck, pdf in zip((theirs, mine), pdfs):
            result = oracle_helper.export_pdf(deck, pdf)
            assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        pages = [oracle_helper.pdf_pages(pdf, dpi=144) for pdf in pdfs]
        if pages[0] is not None:
            from PIL import ImageChops

            assert len(pages[0]) == len(pages[1]) == 11
            for number, (a, b) in enumerate(zip(*pages), start=1):
                difference = ImageChops.difference(a, b).convert("L").point(
                    lambda v: 255 if v > 64 else 0)
                changed = sum(difference.histogram()[255:])
                assert changed <= a.size[0] * a.size[1] // 1000, (number, changed)
    finally:
        oracle_helper.cleanup(theirs, mine, *pdfs)
