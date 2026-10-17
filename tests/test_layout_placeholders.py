"""A layout's own placeholders (production feedback round 7, L1).

What PowerPoint does with them was measured in PowerPoint 16 (Mac) on made-up decks
(agent-work/feedback-round7/oracle/results.md, B and A6): a placeholder written onto a
layout opens without repair; ``type="obj"`` is saved without a type; a duplicate idx is
accepted silently and leaves slides two placeholders to inherit from; a slide made on the
layout afterwards gets every layout placeholder empty, without a frame, a picture's without
a text body; slides already on the layout get it only when the layout is applied again.
The idx rule (largest plus one, at least 13) is inferred from Office's bundled layouts."""

from __future__ import annotations

import datetime as dt
import io
import struct
import zipfile
import zlib

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from pptx_agent.validate import check

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CLOCK = lambda: dt.datetime(2027, 1, 4, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
EMU = 12700


def png(width: int, height: int) -> bytes:
    raw = b"".join(b"\x00" + b"\x20\x60\xa0" * width for _ in range(height))

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def run_deck() -> Document:
    """Two slides on a copy of Title Only, one on Title and Content."""
    deck = Document.new()
    deck.duplicate_layout("Title Only", name="Run")
    for number in (1, 2):
        deck.add_slide("Run").title = f"Item {number}"
    deck.add_slide("Title and Content").title = "Other"
    return deck


def ph_of(shape) -> etree._Element:
    return next(shape._element.iter(f"{P}ph"))


def reopened(deck: Document) -> Document:
    data = deck.to_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        assert archive.testzip() is None
    return Document.open(data)


# -- the layout ---------------------------------------------------------------------------------


def test_a_placeholder_on_a_layout_gets_a_unique_idx_from_13_and_its_frame():
    deck = run_deck()
    layout = deck.layout("Run")
    before = {shape.placeholder.idx for shape in layout.shapes if shape.placeholder}
    assert max(i or 0 for i in before) == 12                 # date 10, footer 11, number 12
    note = layout.add_placeholder("body", 914400, 1600200, 3048000, 914400, name="Note")
    photo = layout.add_placeholder("pic", 6096000, 1600200, 3048000, 2286000)
    more = layout.add_placeholder("content", 914400, 3000000, 3048000, 914400)
    assert [s.placeholder for s in (note, photo, more)] == [
        ("body", 13), ("pic", 14), (None, 15)]
    assert ph_of(more).get("type") is None                   # content: no type="obj"
    assert note.bounds == (914400, 1600200, 3048000, 914400)
    assert note.name == "Note" and photo.name == f"Picture Placeholder {int(photo.id.split('.')[1]) - 1}"
    ids = [shape.id for shape in layout.shapes]
    assert len(ids) == len(set(ids))
    assert note.text == "Note"                               # the text Slide Master view shows
    assert next(photo._element.iter(f"{A}t")).text == photo.name
    assert check(reopened(deck).package) == []


def test_the_idx_follows_the_largest_on_the_layout():
    deck = Document.new()
    layout = deck.layout("Title and Content")
    layout.add_placeholder("body", 0, 0, 914400, 914400)
    tree = deck.package.tree(layout.part_path)
    for ph in tree.iter(f"{P}ph"):
        if ph.get("idx") == "13":
            ph.set("idx", "21")                              # as a template from elsewhere
    deck._reset_caches()
    added = deck.layout("Title and Content").add_placeholder("body", 0, 0, 914400, 914400)
    assert added.placeholder.idx == 22


def test_what_is_refused():
    deck = run_deck()
    layout = deck.layout("Run")
    with pytest.raises(ValueError, match="already has a title"):
        layout.add_placeholder("title", 0, 0, 914400, 914400)
    with pytest.raises(ValueError, match="tbl"):
        layout.add_placeholder("tbl", 0, 0, 914400, 914400)
    with pytest.raises(ValueError, match="width and a height"):
        layout.add_placeholder("body", 0, 0, 0, 914400)
    blank = deck.layout("Blank")
    title = blank.add_placeholder("title", 0, 0, 914400, 914400)
    assert title.placeholder == ("title", None)


def test_adding_is_one_undo_step():
    deck = run_deck()
    layout = deck.layout("Run")
    count = len(layout.shapes)
    layout.add_placeholder("body", 0, 0, 914400, 914400)
    assert len(deck.layout("Run").shapes) == count + 1
    deck.undo()
    assert len(deck.layout("Run").shapes) == count


# -- slides -------------------------------------------------------------------------------------


def test_a_new_slide_gets_it_empty_without_a_frame_and_a_picture_without_text():
    deck = run_deck()
    layout = deck.layout("Run")
    layout.add_placeholder("body", 914400, 1600200, 3048000, 914400, name="Note")
    layout.add_placeholder("pic", 6096000, 1600200, 3048000, 2286000, name="Photo")
    slide = deck.add_slide("Run")
    by_idx = {s.placeholder.idx: s for s in slide.shapes if s.placeholder}
    note, photo = by_idx[13], by_idx[14]
    assert note.name == "Note" and ph_of(note).get("type") == "body"
    assert note._element.find(f"{P}spPr").find(f"{A}xfrm") is None
    assert note.bounds == (914400, 1600200, 3048000, 914400)     # inherited
    assert photo._element.find(f"{P}txBody") is None
    assert photo.bounds == (6096000, 1600200, 3048000, 2286000)
    note.set_text("Ready")
    picture = photo.insert_picture(png(40, 20))
    assert picture.kind == "picture" and picture.id == photo.id and picture.name == "Photo"
    assert picture.placeholder == ("pic", 14)
    crop = next(picture._element.iter(f"{A}srcRect"))
    assert crop.get("l") == crop.get("r") == "16667"            # 2:1 cropped into 4:3
    assert picture.bounds == (6096000, 1600200, 3048000, 2286000)
    again = reopened(deck)
    assert check(again.package) == []
    assert again.slides[-1].shape(note.id).text == "Ready"


def test_a_content_placeholder_typed_obj_reaches_a_new_slide_without_the_type():
    deck = run_deck()
    layout = deck.layout("Run")
    shape = layout.add_placeholder("content", 0, 0, 914400, 914400)
    ph_of(shape).set("type", "obj")
    slide = deck.add_slide("Run")
    copy = next(s for s in slide.shapes if s.placeholder and s.placeholder.idx == 13)
    assert ph_of(copy).get("type") is None and ph_of(copy).get("idx") == "13"


def test_slides_already_on_the_layout_lack_it_until_given_it():
    deck = run_deck()
    layout = deck.layout("Run")
    layout.add_placeholder("body", 914400, 1600200, 3048000, 914400, name="Note")
    first = deck.slides[0]
    assert all(s.placeholder is None or s.placeholder.idx != 13 for s in first.shapes)
    first.add_textbox(0, 0, 914400, 914400, text="mine")
    added = first.add_layout_placeholder(13)
    assert added.placeholder == ("body", 13) and added.name == "Note"
    ph = ph_of(added)
    assert ph.get("sz") is None
    assert added._element.find(f"{P}spPr").getchildren() == []
    body = added._element.find(f"{P}txBody")
    assert [etree.QName(c).localname for c in body] == ["bodyPr", "lstStyle", "p"]
    assert [etree.QName(c).localname for c in body.find(f"{A}p")] == ["endParaRPr"]
    tree = [etree.QName(c).localname for c in first._sp_tree()][2:]
    placeholders = [s for s in first.shapes if s.placeholder]
    assert placeholders[-1].id == added.id
    assert first.shapes[-1].text == "mine" and tree[-1] == "sp"   # the text box stays frontmost
    ids = [s.id for s in first.shapes]
    assert len(ids) == len(set(ids))
    with pytest.raises(ValueError, match="already has"):
        first.add_layout_placeholder(13)
    with pytest.raises(KeyError, match="no placeholder 99"):
        first.add_layout_placeholder(99)


def test_a_copy_takes_its_place_in_the_layouts_order():
    deck = run_deck()
    layout = deck.layout("Run")
    layout.add_placeholder("body", 0, 0, 914400, 914400)
    layout.add_placeholder("pic", 0, 914400, 914400, 914400)
    slide = deck.slides[1]
    slide.add_layout_placeholder(14)
    slide.add_layout_placeholder(13)
    assert [s.placeholder.idx for s in slide.shapes if s.placeholder] == [None, 13, 14]
    bare = deck.add_slide("Run")
    for shape in [s for s in bare.shapes if s.placeholder]:
        shape.delete()
    bare.add_textbox(0, 0, 914400, 914400, text="mine")
    bare.add_layout_placeholder(14)
    bare.add_layout_placeholder("title")
    assert [s.placeholder.idx if s.placeholder else "box" for s in bare.shapes] == [
        None, 14, "box"]


def test_insert_picture_only_into_an_empty_picture_placeholder():
    deck = run_deck()
    layout = deck.layout("Run")
    pic = layout.add_placeholder("pic", 0, 0, 914400, 914400)
    with pytest.raises(ValueError, match="layout"):
        pic.insert_picture(png(4, 4))
    with pytest.raises(ValueError, match="not an empty picture placeholder"):
        deck.slides[0].shapes[0].insert_picture(png(4, 4))


# -- the tool -------------------------------------------------------------------------------------


@pytest.fixture
def tools():
    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=CLOCK)
    session.open(run_deck().to_bytes(), name="run.pptx")
    yield toolbox, session
    toolbox.close()


def call(tools, name, arguments):
    toolbox, session = tools
    return toolbox.dispatch(session, name, {"doc": "d1", **arguments})


def test_add_placeholder_and_fill_it_on_a_new_slide(tools):
    toolbox, session = tools
    layout = session.entry("d1").document.layout("Run").id
    added = call(tools, "ppt_edit_master", {
        "action": "add_placeholder", "target": "Run", "type": "body", "text": "Note",
        "box": {"x": 72, "y": 126, "w": 240, "h": 72}})
    assert added.ok, added.error
    data = added.data
    assert data["idx"] == 13 and data["type"] == "body" and data["name"] == "Note"
    assert data["box"] == {"x": 72, "y": 126, "w": 240, "h": 72}
    assert data["slides_without"] == ["s:256", "s:257"]
    assert f"target {data['placeholder']}" in data["note"]
    assert added.created == [data["placeholder"]] and data["placeholder"].startswith(layout)
    photo = call(tools, "ppt_edit_master", {
        "action": "add_placeholder", "target": layout, "type": "pic",
        "box": {"x": 480, "y": 126, "w": 240, "h": 180}})
    assert photo.data["idx"] == 14 and "ppt_add_picture" in photo.data["note"]
    slide = call(tools, "ppt_add_slide", {"layout": "Run", "title": "Item 3"})
    holders = {entry.get("idx"): entry for entry in slide.data["placeholders"]}
    assert holders[13]["name"] == "Note" and holders[13]["box"] == data["box"]
    assert holders[14]["type"] == "pic"
    assert None in holders                                  # the title: no idx reported
    filled = call(tools, "ppt_set_text", {"items": [{"target": holders[13]["id"],
                                                     "text": "Ready"}]})
    assert filled.ok, filled.error
    blob = session.add_blob(png(30, 30), "photo.png")
    picture = call(tools, "ppt_add_picture", {"image": blob, "target": holders[14]["id"]})
    assert picture.ok, picture.error
    assert picture.data["address"] == holders[14]["id"]
    assert picture.data["box"] == {"x": 480, "y": 126, "w": 240, "h": 180}
    assert "Filled" in picture.summary
    saved = call(tools, "save_document", {"name": "out.pptx", "format": "pptx"})
    assert saved.ok, saved.error


def test_add_placeholder_to_slides_already_on_the_layout(tools):
    added = call(tools, "ppt_edit_master", {
        "action": "add_placeholder", "target": "Run", "type": "content",
        "box": {"x": 72, "y": 126, "w": 240, "h": 72}, "slides": [1]})
    assert added.ok, added.error
    assert added.data["added_to"][0].startswith("256.")
    assert added.data["slides_without"] == ["s:257"]
    later = call(tools, "ppt_edit_master", {"action": "add_placeholder",
                                            "target": added.data["placeholder"],
                                            "slides": ["s:257"]})
    assert later.ok, later.error
    assert later.data["added_to"][0].startswith("257.") and "slides_without" not in later.data
    assert later.created == later.data["added_to"]
    again = call(tools, "ppt_edit_master", {"action": "add_placeholder",
                                            "target": added.data["placeholder"],
                                            "slides": ["s:257"]})
    assert again.error.code == "invalid_arguments" and "already has" in again.error.message


def test_add_placeholder_refusals(tools):
    base = {"action": "add_placeholder", "target": "Run",
            "box": {"x": 0, "y": 0, "w": 10, "h": 10}}
    missing = call(tools, "ppt_edit_master", {k: v for k, v in base.items()})
    assert missing.error.field == "type"
    assert missing.error.valid_options == ["body", "content", "pic", "title"]
    title = call(tools, "ppt_edit_master", {**base, "type": "title"})
    assert title.error.field == "type" and "title placeholder already" in title.error.message
    no_box = call(tools, "ppt_edit_master", {"action": "add_placeholder", "target": "Run",
                                             "type": "body"})
    assert no_box.error.field == "box"
    flat = call(tools, "ppt_edit_master", {**base, "type": "body",
                                           "box": {"x": 0, "y": 0, "w": 0, "h": 10}})
    assert flat.error.field == "box"
    stray = call(tools, "ppt_edit_master", {**base, "type": "body", "slides": [3]})
    assert stray.error.field == "slides" and "s:258" in stray.error.message
    assert stray.error.valid_options == ["s:256", "s:257"]
    nothing = call(tools, "ppt_edit_master", {"action": "add_placeholder",
                                              "target": "L1.2"})
    assert nothing.error.field in ("slides", "target")
