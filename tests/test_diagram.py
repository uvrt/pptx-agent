"""SmartArt (ROADMAP E4): node text, and adding and removing nodes.

``powerpoint-smartart.pptx`` was written by PowerPoint itself: a hand-made data model (three
nodes in "Basic Block List"; two parents with bullet children in "Vertical Bullet List")
opened in PowerPoint for Mac 16 and saved, so its presentation points and cached drawing are
PowerPoint's own.  In the bullet list a child's text is a paragraph of its parent's
``childText`` shape, which is what the drawing update has to get exactly right.
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest
from lxml import etree

from conftest import FIXTURE_DIR
from pptx_agent import Document
from test_validity import assert_valid

DECK = FIXTURE_DIR / "powerpoint-smartart.pptx"
BLOCKS, BULLETS = "256.3", "257.3"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
DSP = "{http://schemas.microsoft.com/office/drawing/2008/diagram}"


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()
                if not info.is_dir()}


def _drawing_texts(document: Document, shape: str) -> list[list[str]]:
    """Each cached drawing shape's paragraphs, in drawing order."""
    part = document.shape(shape).diagram.drawing_part
    root = etree.fromstring(document.package.read(part))
    return [["".join(t.text or "" for t in p.iter(A + "t")) for p in body.findall(A + "p")]
            for body in root.iter(DSP + "txBody")]


def _gates(document: Document, original: bytes) -> bytes:
    """Save, check validity, undo to the original bytes and redo to the edited ones."""
    edited = document.to_bytes()
    assert_valid(edited, original)
    while document.undo():
        pass
    assert _parts(document.to_bytes()) == _parts(original)
    while document.redo():
        pass
    assert document.to_bytes() == edited
    return edited


def test_reading():
    document = Document.open(str(DECK))
    blocks = document.shape(BLOCKS).diagram
    assert blocks.texts == ["Plan", "Build", "Ship"]
    assert blocks.layout.endswith("/layout/default")
    assert blocks.drawing_part == "ppt/diagrams/drawing1.xml"
    bullets = document.shape(BULLETS).diagram
    assert [(n.level, n.text) for n in bullets.nodes] == [
        (0, "Goals"), (1, "Faster edits"), (1, "Fewer prompts"), (0, "Risks"),
        (1, "Stale caches")]
    assert bullets.nodes[1].parent.text == "Goals"
    assert [c.text for c in bullets.nodes[0].children] == ["Faster edits", "Fewer prompts"]
    assert bullets.model["nodes"][0] == {"id": bullets.nodes[0].id, "lvl": 0, "t": "Goals"}


def test_a_node_and_its_drawing_shape_change_together():
    original = DECK.read_bytes()
    document = Document.open(original)
    diagram = document.shape(BLOCKS).diagram
    diagram.set_text(1, "Build it well")
    assert _drawing_texts(document, BLOCKS) == [["Plan"], ["Build it well"], ["Ship"]]
    edited = _gates(document, original)
    reopened = Document.open(edited).shape(BLOCKS).diagram
    assert reopened.texts == ["Plan", "Build it well", "Ship"]
    assert reopened.drawing_part is not None
    # The run kept PowerPoint's formatting in both places.
    assert b'sz="6500"' in Document.open(edited).package.read(reopened.drawing_part)


def test_a_bullet_childs_paragraph_is_found_in_its_parents_shape():
    original = DECK.read_bytes()
    document = Document.open(original)
    diagram = document.shape(BULLETS).diagram
    diagram.node(2).text = "Far fewer prompts"
    diagram.node(0).text = "Goals\nfor E4"  # two paragraphs where there was one
    assert _drawing_texts(document, BULLETS) == [
        ["Goals", "for E4"], ["Faster edits", "Far fewer prompts"], ["Risks"],
        ["Stale caches"]]
    edited = _gates(document, original)
    assert Document.open(edited).shape(BULLETS).diagram.texts[:3] == [
        "Goals\nfor E4", "Faster edits", "Far fewer prompts"]


def test_a_drawing_that_already_disagrees_is_dropped():
    """If the cache does not say what the data model says, it cannot be patched exactly --
    it goes, and PowerPoint lays the diagram out again (measured)."""
    document = Document.open(DECK.read_bytes())
    diagram = document.shape(BLOCKS).diagram
    part = diagram.drawing_part
    root = document.package.tree(part)
    next(t for t in root.iter(A + "t") if t.text == "Ship").text = "Stale"
    document.package.mark_dirty(part)
    original = document.to_bytes()
    document = Document.open(original)
    diagram = document.shape(BLOCKS).diagram
    diagram.set_text("{5A000000-0000-4000-8000-00000000100A}", "Deliver")
    assert diagram.drawing_part is None
    edited = _gates(document, original)
    assert part not in _parts(edited)
    assert b"dataModelExt" not in _parts(edited)["ppt/diagrams/data1.xml"]
    assert b"diagramDrawing" not in _parts(edited)["ppt/slides/_rels/slide1.xml.rels"]


def test_adding_a_node_drops_the_drawing():
    original = DECK.read_bytes()
    document = Document.open(original)
    diagram = document.shape(BLOCKS).diagram
    added = diagram.add_node("Measure", index=1)
    assert diagram.texts == ["Plan", "Measure", "Build", "Ship"]
    assert added.level == 0 and diagram.drawing_part is None
    child = diagram.node(0).add_child("Scope")
    assert [(n.level, n.text) for n in diagram.nodes][:2] == [(0, "Plan"), (1, "Scope")]
    assert child.parent.text == "Plan"
    edited = _gates(document, original)
    data = _parts(edited)["ppt/diagrams/data1.xml"].decode()
    orders = [int(v) for v in re.findall(
        r'<dgm:cxn modelId="[^"]+" srcId="\{5A000000-0000-4000-8000-000000001001\}" '
        r'destId="[^"]+" srcOrd="(\d+)"', data)]  # parent-of connections: no type
    assert sorted(orders) == [0, 1, 2, 3]
    assert "ppt/diagrams/drawing1.xml" not in _parts(edited)
    assert "ppt/diagrams/drawing2.xml" in _parts(edited)  # the other diagram's is untouched


def test_removing_a_node_takes_its_children_and_presentation_points():
    original = DECK.read_bytes()
    document = Document.open(original)
    diagram = document.shape(BULLETS).diagram
    goals = diagram.node(0)
    pres = [p for p in re.findall(r'presAssocID="([^"]+)"',
                                  document.package.read(diagram.part).decode())]
    assert goals.id in pres
    diagram.remove_node(goals)
    assert diagram.texts == ["Risks", "Stale caches"]
    data = document.package.read(diagram.part).decode()
    assert goals.id not in data
    for gone in ("Faster edits", "Fewer prompts"):
        assert gone not in data
    edited = _gates(document, original)
    assert Document.open(edited).shape(BULLETS).diagram.texts == ["Risks", "Stale caches"]


def test_the_last_node_stays():
    diagram = Document.open(DECK.read_bytes()).shape(BULLETS).diagram
    diagram.remove_node(0)
    with pytest.raises(ValueError):
        diagram.remove_node(0)


def test_unchanged_text_changes_nothing():
    document = Document.open(DECK.read_bytes())
    document.shape(BLOCKS).diagram.set_text(0, "Plan")
    assert not document.history.can_undo() and not document.package.dirty_parts


def test_pptx2svg_draws_the_edited_text():
    pytest.importorskip("pptx2svg")
    document = Document.open(DECK.read_bytes())
    document.shape(BULLETS).diagram.set_text(1, "Much faster edits")
    svg = document.slides[1].render_svg()
    assert "Much faster edits" in svg and "Faster edits" not in svg
