"""Slide masters and layouts as editable parts (production feedback round 3, G5).

What PowerPoint writes for each edit was measured in PowerPoint 16 (Mac): the probes are in
the round's notes; the assertions below are what they showed."""

from __future__ import annotations

import datetime as dt
import io
import re
import zipfile

import pytest
from lxml import etree
from ooxml_edit.tools import Toolbox

from pptx_agent import Document, Layout, Master, Shape
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CLOCK = lambda: dt.datetime(2027, 1, 4, 9, 0, tzinfo=dt.timezone.utc)  # noqa: E731
LEGAL = ('<p:sp><p:nvSpPr><p:cNvPr id="50" name="Legal footer"/><p:cNvSpPr txBox="1"/>'
         '<p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="6400800"/>'
         '<a:ext cx="7000000" cy="300000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
         '</a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r>'
         '<a:rPr lang="en-US" sz="1000"/><a:t>© 2026 Example Corp. Confidential.</a:t>'
         '</a:r></a:p></p:txBody></p:sp>')


def _patched(content: bytes, edits: dict) -> bytes:
    source, out = zipfile.ZipFile(io.BytesIO(content)), io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            data = source.read(name)
            if name in edits:
                data = edits[name](data.decode()).encode()
            target.writestr(name, data)
    return out.getvalue()


def template(*, ids: bool = False) -> bytes:
    """A title slide and two content slides; a legal line on the master; with ``ids``, the
    master and layout ids as a template from elsewhere has them (not 2147483648 + n)."""
    deck = Document.new()
    deck.add_slide("Title Slide").shapes[0].set_text("Plan 2027")
    for number in (2, 3):
        deck.add_slide("Title and Content").shapes[0].set_text(f"Slide {number}")
    edits = {"ppt/slideMasters/slideMaster1.xml":
             lambda x: x.replace("</p:spTree>", LEGAL + "</p:spTree>", 1)}
    if ids:
        edits["ppt/presentation.xml"] = lambda x: x.replace('id="2147483648"',
                                                            'id="2147483700"')
        master = edits["ppt/slideMasters/slideMaster1.xml"]
        edits["ppt/slideMasters/slideMaster1.xml"] = lambda x: re.sub(
            r'sldLayoutId id="21474836(\d\d)"',
            lambda m: f'sldLayoutId id="{2147483800 + 3 * int(m.group(1))}"', master(x))
    return _patched(deck.to_bytes(), edits)


def part(deck: Document, name: str) -> etree._Element:
    return etree.fromstring(deck.to_bytes() and zipfile.ZipFile(io.BytesIO(deck.to_bytes()))
                            .read(name))


# -- objects and addresses ---------------------------------------------------------------------


def test_masters_and_layouts_have_ids_from_the_file_and_editable_shapes():
    deck = Document.open(template())
    master = deck.masters[0]
    assert isinstance(master, Master) and master.id == "M0" and master.name == "Office Theme"
    assert master.part_path == "ppt/slideMasters/slideMaster1.xml"
    assert [layout.id for layout in master.layouts] == [f"L{n}" for n in range(1, 12)]
    legal = next(shape for shape in master.shapes if shape.name == "Legal footer")
    assert isinstance(legal, Shape) and legal.id == "M0.50"
    assert deck.shape("M0.50") .text.startswith("© 2026")
    assert deck.resolve("M0.50/p0/r0").text.startswith("© 2026")
    layout = deck.layout("L2")
    assert isinstance(layout, Layout) and layout.name == "Title and Content"
    assert layout.master == master and layout.type == "obj"
    assert [slide.slide_id for slide in layout.slides] == [257, 258]
    assert [slide.slide_id for slide in master.slides] == [256, 257, 258]
    assert [s.slide_id for s in deck.affected_slides("L2.3")] == [257, 258]


def test_ids_are_the_sld_layout_ids_so_they_survive_a_save():
    """PowerPoint keeps sldMasterId and sldLayoutId when it saves (measured): the key is the
    id less 2147483648, not a position."""
    deck = Document.open(template(ids=True))
    assert deck.masters[0].id == "M52"
    assert [layout.id for layout in deck.layouts][:3] == ["L299", "L302", "L305"]
    assert deck.shape("M52.50").name == "Legal footer"
    again = Document.open(deck.to_bytes())
    assert [layout.id for layout in again.layouts] == [layout.id for layout in deck.layouts]


def test_a_master_text_box_is_edited_like_a_slide_shape_and_undone():
    deck = Document.open(template())
    legal = deck.shape("M0.50")
    legal.set_text(legal.text.replace("2026", "2027"))
    legal.paragraphs[0].runs[0].bold = True
    legal.set_solid_fill("accent1")
    legal.top = 6350000
    assert deck.shape("M0.50").text == "© 2027 Example Corp. Confidential."
    assert deck.shape("M0.50").paragraphs[0].runs[0].size == 10
    assert deck.shape("M0.50").top == 6350000
    for _ in range(4):
        deck.undo()
    assert deck.shape("M0.50").text.startswith("© 2026")


def test_adding_shapes_and_pictures_to_a_layout_and_a_master(tmp_path):
    from images import png

    deck = Document.open(template())
    box = deck.layout("Title Slide").add_textbox(457200, 457200, 3000000, 400000)
    box.set_text("Vertrouwelijk")
    assert box.id.startswith("L1.")
    logo = deck.masters[0].add_picture(png(64, 64, lambda x, y: (0, 0, 255)), 10900000, 228600, width=914400)
    assert logo.id.startswith("M0.") and logo.kind == "picture"
    rect = deck.masters[0].add_shape("rect", 0, 0, 12192000, 100000)
    assert rect.id.startswith("M0.")
    assert deck.validate() == []
    again = Document.open(deck.to_bytes())
    assert any(s.text == "Vertrouwelijk" for s in again.layout("Title Slide").shapes)
    assert again.shape(logo.id).image_part is not None


def test_text_fit_is_measured_on_the_slides_not_the_layout():
    deck = Document.open(template())
    with pytest.raises(ValueError, match="affected_slides"):
        deck.shape("M0.50").text_fit()


# -- prompts -----------------------------------------------------------------------------------


def test_set_prompt_writes_the_first_level_and_marks_it_custom():
    """Measured: typing over the first line of a layout's body placeholder in Slide Master
    view sets hasCustomPrompt="1" and keeps the other levels."""
    deck = Document.open(template())
    body = deck.layout("Title and Content").placeholder("body")
    assert body.prompt is None
    size = body.paragraphs[0].runs[0].effective_size
    body.set_prompt("Typ hier uw bevinding")
    body = deck.layout("Title and Content").placeholder("body")
    assert body.prompt.split("\n")[:2] == ["Typ hier uw bevinding", "Second level"]
    assert len(body.paragraphs) == 5
    assert body.paragraphs[0].runs[0].effective_size == size
    ph = body._element.find(f".//{P}ph")
    assert ph.get("hasCustomPrompt") == "1" and ph.get("idx") == "1"
    deck.undo()
    assert deck.layout("Title and Content").placeholder("body").prompt is None


def test_typing_into_a_layout_placeholder_is_a_prompt_and_formatting_is_not():
    deck = Document.open(template())
    title = deck.layout("Title Slide").placeholder("title")
    title.paragraphs[0].runs[0].bold = True
    assert deck.layout("Title Slide").placeholder("title").prompt is None
    deck.layout("Title Slide").placeholder("subtitle").set_text("Ondertitel")
    assert deck.layout("Title Slide").placeholder("subtitle").prompt == "Ondertitel"
    # the footer placeholders hold content, never a prompt
    deck.layout("Title Slide").placeholder("footer").set_text("Example BV")
    assert deck.layout("Title Slide").placeholder("footer").prompt is None
    with pytest.raises(ValueError, match="layout"):
        deck.masters[0].placeholder("body").set_prompt("x")
    with pytest.raises(ValueError, match="set_footer"):
        deck.layout("L2").placeholder("footer").set_prompt("x")


# -- find and replace --------------------------------------------------------------------------


def test_find_text_scopes():
    deck = Document.open(template())
    assert deck.find_text("2026") == []
    found = deck.find_text("2026", scope="all")
    # the master's legal line; the date placeholders hold only a field: not text
    assert [(block.address, block.part, block.slide) for block in found] == [
        ("M0.50", "M0", 0)]
    assert deck.find_text("Slide 2", scope="layouts") == []
    assert [b.address for b in deck.find_text("Slide 2", scope="all")] == ["257.2"]
    layout_titles = deck.find_text("Click to edit Master title", scope="layouts")
    assert [b.part for b in layout_titles][:2] == ["L1", "L2"]
    with pytest.raises(ValueError, match="scope"):
        deck.find_text("x", scope="everything")


# -- deleting what others inherit from ---------------------------------------------------------


def test_deleting_a_placeholder_slides_inherit_from_is_refused():
    deck = Document.open(template())
    body = deck.layout("Title and Content").placeholder("body")
    assert body.inherited_by() == ["257.3", "258.3"]
    with pytest.raises(ValueError, match="257.3, 258.3"):
        body.delete()
    master_date = deck.masters[0].placeholder("date")
    assert "L2.4" in master_date.inherited_by()
    with pytest.raises(ValueError, match="inherit"):
        master_date.delete()
    # a layout placeholder no slide uses, and a master's own text box, go
    deck.layout("Two Content").placeholder("title").delete()
    deck.shape("M0.50").delete()
    assert deck.find_text("2026", scope="all") == []


# -- the layout list ---------------------------------------------------------------------------


def _master_ids(deck):
    root = deck.package.tree(deck.masters[0].part_path)
    return [node.get("id") for node in root.iter(f"{P}sldLayoutId")]


def test_duplicate_layout_as_powerpoint_does():
    deck = Document.open(template())
    copy = deck.duplicate_layout("Title and Content")
    assert copy.name == "1_Title and Content" and copy.id == "L12"
    assert [layout.name for layout in deck.layouts][:3] == [
        "Title Slide", "Title and Content", "1_Title and Content"]
    assert _master_ids(deck)[2] == "2147483660"
    assert copy.type == "obj" and len(copy.shapes) == len(deck.layout("L2").shapes)
    assert deck.duplicate_layout("Title and Content").name == "2_Title and Content"
    named = deck.duplicate_layout("L2", name="Inhoud NL")
    assert named.name == "Inhoud NL"
    data = deck.to_bytes()
    types = zipfile.ZipFile(io.BytesIO(data)).read("[Content_Types].xml").decode()
    assert f'PartName="/{copy.part_path}"' in types
    assert deck.validate() == []
    slide = deck.add_slide("Inhoud NL")
    assert slide.layout.id == named.id


def test_add_layout_as_insert_layout_makes_one():
    deck = Document.open(template())
    legal = deck.add_layout("Juridisch", after="Title and Content")
    assert [layout.name for layout in deck.layouts][2] == "Juridisch"
    root = deck.package.tree(legal.part_path)
    assert root.get("preserve") == "1" and root.get("userDrawn") == "1"
    assert root.get("type") is None and legal.type is None
    assert [(s.placeholder.type, s.placeholder.idx) for s in legal.shapes] == [
        ("title", None), ("dt", 10), ("ftr", 11), ("sldNum", 12)]
    assert legal.placeholder("title").text == "Click to edit Master title style"
    assert all(s._element.find(f"{P}spPr") is not None and len(s._element.find(f"{P}spPr")) == 0
               for s in legal.shapes)
    with pytest.raises(ValueError, match="already has a layout"):
        deck.add_layout("Juridisch")
    assert deck.validate() == []


def test_rename_and_delete_layouts():
    deck = Document.open(template())
    renamed = deck.layout("Title and Content").rename("Inhoud")
    assert renamed.id == "L2" and deck.layout("Inhoud").slides[0].slide_id == 257
    with pytest.raises(ValueError, match="already has a layout"):
        deck.rename_layout("Inhoud", "Title Slide")
    with pytest.raises(ValueError, match="in use by slide.*s:257, s:258"):
        deck.delete_layout("Inhoud")
    unused = deck.layout("Two Content")
    deck.delete_layout(unused)
    assert "Two Content" not in [layout.name for layout in deck.layouts]
    assert unused.part_path not in deck.package.part_names
    assert "2147483652" not in _master_ids(deck)
    moved = deck.delete_layout("Inhoud", reassign_to="Title Only")
    assert moved == [257, 258]
    assert [slide.layout.name for slide in deck.slides] == ["Title Slide", "Title Only",
                                                             "Title Only"]
    assert deck.slides[1].shapes[0].text == "Slide 2"
    assert deck.validate() == []
    deck.undo()
    assert [slide.layout.name for slide in deck.slides][1] == "Inhoud"


# -- Header & Footer ---------------------------------------------------------------------------


def _footers(deck, part_path):
    tree = deck.package.tree(part_path)
    out = []
    for sp in tree.iter(f"{P}sp"):
        ph = sp.find(f".//{P}ph")
        if ph is not None and ph.get("type") in ("dt", "ftr", "sldNum"):
            out.append((ph.get("type"), "".join(t.text or "" for t in sp.iter(f"{A}t"))))
    return out


def test_set_footer_apply_to_all_as_powerpoint_writes_it():
    deck = Document.open(template())
    changed = deck.set_footer("© 2027 Example Corp", slide_number=True, date="auto",
                              today=dt.date(2027, 1, 5))
    assert changed == [257, 258]                        # not the title slide
    presentation = deck.package.tree("ppt/presentation.xml")
    assert presentation.get("showSpecialPlsOnTitleSld") == "0"
    master = deck.package.tree("ppt/slideMasters/slideMaster1.xml")
    hf = master.find(f"{P}hf")
    assert dict(hf.attrib) == {"hdr": "0"}
    assert [etree.QName(c).localname for c in master][:5] == [
        "cSld", "clrMap", "sldLayoutIdLst", "hf", "txStyles"]
    for path in ("ppt/slideMasters/slideMaster1.xml", deck.layout("L2").part_path,
                 deck.layout("L7").part_path):
        found = dict(_footers(deck, path))
        # en-US datetime1 is M/d/yy (measured: pptx2svg's DATE_FORMATS)
        assert found["ftr"] == "© 2027 Example Corp" and found["dt"] == "1/5/27"
    slide = deck.slides[1]
    assert _footers(deck, slide.part_path) == [
        ("dt", "1/5/27"), ("ftr", "© 2027 Example Corp"), ("sldNum", "2")]
    dt_shape = next(s for s in slide.shapes if s.placeholder and s.placeholder.type == "dt")
    field = dt_shape._element.find(f".//{A}fld")
    assert field.get("type") == "datetime1"
    assert len(dt_shape._element.find(f"{P}spPr")) == 0          # it keeps the layout's place
    number = next(s for s in slide.shapes if s.placeholder and s.placeholder.type == "sldNum")
    layout_number = deck.layout("L2").placeholder("slide_number")
    assert number._element.find(f".//{A}fld").get("id") == \
        layout_number._element.find(f".//{A}fld").get("id")
    assert _footers(deck, deck.slides[0].part_path) == []
    assert deck.validate() == []
    # again: updated in place, nothing doubled
    deck.set_footer("© 2028 Example Corp", date="Januari 2028")
    assert _footers(deck, slide.part_path) == [
        ("dt", "Januari 2028"), ("ftr", "© 2028 Example Corp"), ("sldNum", "2")]


def test_a_new_slide_gets_the_footers_once_header_and_footer_was_applied_to_all():
    """Measured: after Apply to All, PowerPoint's New Slide copies the layout's date, footer
    and slide-number placeholders with their content; before it, none."""
    deck = Document.open(template())
    assert [s.placeholder.type for s in deck.add_slide("Title and Content").shapes] == [
        "title", None]
    deck.set_footer("Example BV", date="Januari 2027")
    slide = deck.add_slide("Title and Content")
    assert _footers(deck, slide.part_path) == [("dt", "Januari 2027"), ("ftr", "Example BV"),
                                               ("sldNum", "5")]
    assert _footers(deck, deck.add_slide("Title Slide").part_path) == []
    deck.set_footer("Example BV", slide_number=False)
    assert [k for k, _ in _footers(deck, deck.add_slide("Title and Content").part_path)] == [
        "ftr"]
    assert deck.validate() == []


def test_set_footer_off_removes_the_slides_placeholders_and_says_so_on_the_master():
    deck = Document.open(template())
    deck.set_footer("Example", date="auto")
    deck.set_footer(None, slide_number=False, date=None)
    assert all(_footers(deck, slide.part_path) == [] for slide in deck.slides)
    hf = deck.package.tree("ppt/slideMasters/slideMaster1.xml").find(f"{P}hf")
    assert dict(hf.attrib) == {"sldNum": "0", "hdr": "0", "ftr": "0", "dt": "0"}
    assert dict(_footers(deck, "ppt/slideMasters/slideMaster1.xml"))["ftr"] == "Example"


def test_set_footer_on_some_slides_leaves_the_master_alone():
    deck = Document.open(template())
    before = etree.tostring(deck.package.tree("ppt/slideMasters/slideMaster1.xml"))
    deck.set_footer("Only here", slides=[2], skip_title=False)
    assert etree.tostring(deck.package.tree("ppt/slideMasters/slideMaster1.xml")) == before
    assert _footers(deck, deck.slides[1].part_path) == [("ftr", "Only here"), ("sldNum", "2")]
    assert _footers(deck, deck.slides[2].part_path) == []
    deck.set_footer("All", skip_title=False)
    assert "showSpecialPlsOnTitleSld" not in deck.package.tree("ppt/presentation.xml").attrib
    assert [kind for kind, _ in _footers(deck, deck.slides[0].part_path)] == ["ftr", "sldNum"]


# -- the tools ---------------------------------------------------------------------------------


@pytest.fixture
def tools():
    toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)
    session = toolbox.session(clock=CLOCK)
    session.open(template(), name="template.pptx")
    yield toolbox, session
    toolbox.close()


def call(tools, name, arguments):
    toolbox, session = tools
    return toolbox.dispatch(session, name, {"doc": "d1", **arguments})


def test_describe_lists_masters_layout_ids_prompts_and_their_text(tools):
    call(tools, "ppt_edit_master", {"action": "set_prompt", "target": "L2.3",
                                    "text": "Typ hier"})
    data = call(tools, "describe", {}).data
    assert data["masters"][0]["id"] == "M0"
    assert data["masters"][0]["text"] == [
        {"id": "M0.50", "text": "© 2026 Example Corp. Confidential."}]
    layout = data["layouts"][1]
    assert layout["id"] == "L2" and layout["name"] == "Title and Content"
    assert layout["placeholders"][1]["id"] == "L2.3"
    assert layout["placeholders"][1]["prompt"].startswith("Typ hier\n")


def test_find_and_replace_reach_layouts_and_masters_with_stories(tools):
    assert call(tools, "find_text", {"text": "2026"}).data == []
    found = call(tools, "find_text", {"text": "2026", "stories": "all"}).data
    assert found == [{"address": "M0.50", "kind": "text", "on": "M0",
                      "context": "© 2026 Example Corp. Confidential."}]
    missed = call(tools, "replace_text", {"find": "2026", "replace": "2027", "expect": "all"})
    assert missed.data == {"count": 0} and "stories=all" in missed.warnings[0]
    one = call(tools, "replace_text", {"find": "2026", "replace": "2027", "expect": "one"})
    assert one.error.code == "not_found" and "M0.50" in one.error.message
    done = call(tools, "replace_text", {"find": "2026", "replace": "2027", "expect": "one",
                                        "stories": "masters"})
    assert done.ok and done.checks["affects"] == {"M0": ["s:256", "s:257", "s:258"]}
    assert done.checks["slides"] == ["s:256", "s:257", "s:258"]
    bad = call(tools, "find_text", {"text": "x", "stories": "layouts", "slides": [1]})
    assert bad.error.code == "invalid_arguments"


def test_set_text_and_set_shape_take_layout_and_master_addresses(tools):
    result = call(tools, "ppt_set_text", {"items": [{"target": "M0.50",
                                                     "text": "© 2027 Voorbeeld BV"}]})
    assert result.ok and result.checks["affects"] == {"M0": ["s:256", "s:257", "s:258"]}
    moved = call(tools, "ppt_set_shape", {"items": [{"target": "L2.3", "y": 150}]})
    assert moved.ok and moved.checks["affects"] == {"L2": ["s:257", "s:258"]}
    refused = call(tools, "ppt_arrange", {"targets": ["L2.3"], "action": "delete"})
    assert refused.error.code == "invalid_arguments" and "257.3" in refused.error.message
    added = call(tools, "ppt_add_shape", {"slide": "L2", "items": [
        {"preset": "textbox", "box": {"x": 36, "y": 500, "w": 400, "h": 20},
         "text": "Juridische tekst"}]})
    assert added.ok and added.created[0].startswith("L2.")
    missing = call(tools, "ppt_set_text", {"items": [{"target": "L99.3", "text": "x"}]})
    assert missing.error.code == "not_found"


def test_edit_master_tool_actions(tools):
    footer = call(tools, "ppt_edit_master", {"action": "set_footer", "text": "Voorbeeld BV",
                                             "date": "auto"})
    assert footer.ok and footer.data == {"slides": ["s:257", "s:258"]}
    toolbox, session = tools
    deck = session.entry("d1").document
    assert dict(_footers(deck, deck.slides[1].part_path))["dt"] == "1/4/27"   # en-US: M/d/yy
    dup = call(tools, "ppt_edit_master", {"action": "duplicate_layout", "target": "L2",
                                          "text": "Inhoud NL"})
    assert dup.data["layout"] == "L12" and dup.created == ["L12"]
    added = call(tools, "ppt_edit_master", {"action": "add_layout", "text": "Juridisch",
                                            "target": "Inhoud NL"})
    assert added.data["layout"] == "L13"
    renamed = call(tools, "ppt_edit_master", {"action": "rename_layout", "target": "L13",
                                              "text": "Juridisch NL"})
    assert {"id": "L13", "name": "Juridisch NL", "slides": 0} in renamed.data["layouts"]
    refused = call(tools, "ppt_edit_master", {"action": "delete_layout", "target": "L2"})
    assert refused.error.code == "invalid_arguments" and "s:257" in refused.error.message
    moved = call(tools, "ppt_edit_master", {"action": "delete_layout", "target": "L2",
                                            "reassign_to": "L12"})
    assert moved.data["moved"] == ["s:257", "s:258"]
    gone = call(tools, "ppt_edit_master", {"action": "delete_layout", "target": "L13"})
    assert gone.removed == ["L13"]
    saved = call(tools, "save_document", {"name": "out.pptx", "format": "pptx"})
    assert saved.ok, saved.error


def test_undo_by_slide_treats_a_layout_edit_as_every_slides(tools):
    call(tools, "ppt_set_text", {"items": [{"target": "257.2", "text": "Mine"}]})
    call(tools, "ppt_edit_master", {"action": "set_prompt", "target": "L2.3",
                                    "text": "Typ hier"})
    refused = call(tools, "undo", {"scope": "257"})
    assert refused.error.code == "entangled"
    assert "slideLayout2.xml" in refused.error.message
    # a slide on another layout is not held up by it
    call(tools, "ppt_set_text", {"items": [{"target": "256.2", "text": "Titel"}]})
    assert call(tools, "undo", {"scope": "256"}).ok
    # plain undo takes the layout edit back
    assert call(tools, "undo", {}).ok
    toolbox, session = tools
    deck = session.entry("d1").document
    assert deck.layout("L2").placeholder("body").prompt is None
    assert deck.slide(257).shapes[0].text == "Mine"


def test_a_layout_edit_on_slides_not_in_scope_is_not_found_by_their_undo(tools):
    call(tools, "ppt_set_text", {"items": [{"target": "256.2", "text": "Titel"}]})
    call(tools, "ppt_edit_master", {"action": "set_prompt", "target": "L2.3",
                                    "text": "Typ hier"})
    undone = call(tools, "undo", {"scope": "256"})
    assert undone.ok
    toolbox, session = tools
    deck = session.entry("d1").document
    assert deck.slide(256).shapes[0].text == "Plan 2027"
    assert deck.layout("L2").placeholder("body").prompt.startswith("Typ hier")
