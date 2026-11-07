"""What pptx2svg is handed (``pptx_agent.edit.resolved``), the resolved slides kept across
calls, the per-part validation kept by bytes, and the layout lookups kept while facts read
a deck: each gives what the work done afresh gives.

Made-up decks (``Document.new()``) and the repository's fixtures.
"""

from __future__ import annotations

import io
import json
import re
import zipfile

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document
from pptx_agent import validate as validation
from pptx_agent.edit import resolved
from pptx_agent.edit.document import reading
from pptx_agent.oxml.xml import qn
from pptx_agent.tools.common import facts

pptx2svg = pytest.importorskip("pptx2svg")
if not hasattr(pptx2svg, "ModelCache"):          # pragma: no cover - pptx2svg before 0.3.8
    pytest.skip("pptx2svg keeps no resolved slides", allow_module_level=True)

EMU = 12700


def _deck() -> Document:
    deck = Document.new()
    for number in range(1, 4):
        slide = deck.add_slide("Title and Content" if number % 2 else "Title Only")
        slide.title = f"Section {number}: a heading long enough to wrap on a narrow box"
        table = slide.add_table(3, 3, 40 * EMU, 120 * EMU, 300 * EMU, 90 * EMU)
        for row in range(3):
            for col in range(3):
                table.table.cell(row, col).text_frame.text = "" if row == 2 else f"{row}-{col}"
        slide.add_textbox(320 * EMU, 110 * EMU, 200 * EMU, 40 * EMU,
                          "A box over the table, with more text than it can hold")
    return deck


@pytest.fixture
def afresh(monkeypatch):
    """Nothing kept: the zip round trip and no cache, as before."""
    def run(work):
        with monkeypatch.context() as patch:
            patch.setattr(resolved, "model_cache", lambda: None)
            patch.setattr(resolved, "source_of", lambda document: document.to_bytes())
            patch.setattr(validation, "_PART_PROBLEMS", type(validation._PART_PROBLEMS)())
            patch.setattr(validation, "_PART_LIMIT", 0)
            return work()
    return run


def _facts(deck: Document) -> str:
    return json.dumps(facts(deck, [slide.slide_id for slide in deck.slides]), sort_keys=True,
                      default=str)


def test_the_parts_handed_over_are_the_ones_a_save_writes():
    deck = _deck()
    deck.slides[1].title = "Renamed"
    with zipfile.ZipFile(io.BytesIO(deck.to_bytes())) as archive:
        saved = {info.filename: archive.read(info) for info in archive.infolist()}
    parts = resolved._Parts(deck)
    assert set(parts) == set(saved)
    for name, data in saved.items():
        if name == "docProps/app.xml":
            # Its counts and titles are brought up to date when saved; pptx2svg reads
            # only AppVersion, which that leaves.
            version = re.compile(rb"<AppVersion>[^<]*</AppVersion>")
            assert version.findall(parts[name]) == version.findall(data)
        else:
            assert parts[name] == data, name


def test_facts_and_views_are_those_worked_out_afresh(afresh):
    deck = _deck()
    cache = resolved.model_cache()
    assert _facts(deck) == afresh(lambda: _facts(deck))
    views = [slide.render_svg() for slide in deck.slides]
    assert views == afresh(lambda: [slide.render_svg() for slide in deck.slides])
    # An edit of one slide: only it is resolved again, and still as afresh.
    misses = cache.misses
    deck.slides[2].add_textbox(40 * EMU, 300 * EMU, 300 * EMU, 30 * EMU, "One more line")
    assert _facts(deck) == afresh(lambda: _facts(deck))
    assert cache.misses == misses + 1
    assert deck.slides[2].render_svg() == afresh(lambda: deck.slides[2].render_svg())
    assert deck.render_png([3]) == afresh(lambda: deck.render_png([3]))


@pytest.mark.parametrize("path", sorted(FIXTURE_DIR.glob("*.pptx")), ids=lambda path: path.stem)
def test_the_fixtures_measure_as_afresh(path, afresh):
    deck = Document.open(str(path))
    assert _facts(deck) == afresh(lambda: _facts(deck))
    assert [str(p) for p in deck.validate()] == afresh(lambda: [str(p) for p in deck.validate()])


def test_validation_kept_by_bytes_sees_every_change(afresh):
    deck = _deck()
    assert deck.validate() == []
    slide = deck.slides[0]
    root = deck.package.tree(slide.part_path)
    root.find(f".//{qn('p:cNvPr')}").set(qn("r:id"), "rId999")
    deck.package.mark_dirty(slide.part_path)
    damaged = deck.validate()
    assert any(p.code == "unknown-relationship-id" for p in damaged)
    assert damaged == afresh(deck.validate)
    root.find(f".//{qn('p:cNvPr')}").attrib.pop(qn("r:id"))
    deck.package.mark_dirty(slide.part_path)
    assert deck.validate() == []
    deck.package.replace_part(slide.part_path, b"<p:sld")
    assert [p.code for p in deck.validate()].count("not-well-formed") == 1


def test_a_layout_change_after_reading_is_seen():
    """The layout placeholder found while reading is not kept past the block: a new one
    put in its place afterwards is what the slide's placeholder is drawn by."""
    from copy import deepcopy

    from lxml import etree

    deck = _deck()
    slide = deck.slides[0]
    title = next(shape for shape in slide.shapes if shape.placeholder is not None)
    with reading():
        before = title.drawn_bounds
        assert title.drawn_bounds == before
    old = slide._layout_placeholder(title.placeholder)
    new = deepcopy(old)
    properties = new.find(qn("p:spPr"))
    for xfrm in properties.findall(qn("a:xfrm")):
        properties.remove(xfrm)
    xfrm = etree.SubElement(properties, qn("a:xfrm"))
    properties.insert(0, xfrm)
    x, y, width, height = before
    etree.SubElement(xfrm, qn("a:off"), x=str(x), y=str(y + 100 * EMU))
    etree.SubElement(xfrm, qn("a:ext"), cx=str(width), cy=str(height))
    old.getparent().replace(old, new)
    with reading():
        assert title.drawn_bounds == (x, y + 100 * EMU, width, height)
