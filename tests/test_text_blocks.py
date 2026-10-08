"""Raw text to edit from: ``outline_blocks``, ``find_text``, ``Shape.text``, and the warning
``set_text`` gives for text copied out of the Markdown outline.

The end-to-end pilot (ROADMAP.md, "Usability") found that ``to_outline`` is the natural
first look at a deck, and that its escaped Markdown (``11\\.9%``) pasted into ``set_text``
would have put backslashes on the slide without a word.
"""

from __future__ import annotations

import difflib
import re
import warnings

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document, MarkdownEscapeWarning, TextBlock
from pptx_agent.edit.text import outline_escapes

REPORT = FIXTURE_DIR / "real-financial-report.pptx"
SMARTART = FIXTURE_DIR / "powerpoint-smartart.pptx"


def test_every_block_resolves_to_its_raw_text_and_writing_it_back_changes_nothing(pptx_path):
    """(Writing stamps the shape's id, by design, so the bytes may differ; the text may not.)"""
    document = Document.open(str(pptx_path))
    before = document.to_outline()
    blocks = document.outline_blocks()
    assert blocks or not any(s.shapes for s in document.slides)
    with warnings.catch_warnings():
        warnings.simplefilter("error", MarkdownEscapeWarning)
        for block in blocks:
            assert isinstance(block, TextBlock)
            target = document.resolve(block.address)
            raw = target.text_frame.text if hasattr(target, "text_frame") else target.text
            assert raw == block.text
            target.text = block.text
    assert document.outline_blocks() == blocks
    assert document.to_outline() == before


def test_blocks_follow_the_outline_and_are_not_escaped():
    document = Document.open(str(REPORT))
    blocks = document.outline_blocks(slides=[1])
    by_address = {b.address: b for b in blocks}
    assert by_address["256.29"] == TextBlock("256.29", "text", "11.9%", 1)
    assert by_address["256.30"].text == "+0.6pt YoY"
    assert by_address["256.4"].text == "2025年度\n第3四半期\n決算サマリー"
    # Same order as the outline's id comments.
    outline_ids = re.findall(r"<!-- (\d+\.\d+(?:#\d+)?)", document.to_outline(slides=[1]))
    assert [b.address for b in blocks] == [i for i in outline_ids if i in by_address]
    cells = [b for b in document.outline_blocks(slides=[2]) if b.kind == "cell"]
    assert ("257.3#5/cell1,1", "4,285億円") in [(c.address, c.text) for c in cells]


def test_smartart_nodes_are_blocks_and_resolve():
    document = Document.open(str(SMARTART))
    nodes = [b for b in document.outline_blocks() if b.kind == "smartart"]
    assert nodes
    first = nodes[0]
    document.resolve(first.address).text = "Renamed"
    assert document.resolve(first.address).text == "Renamed"
    with pytest.raises(IndexError):
        document.resolve(first.address.rsplit("/", 1)[0] + "/node999")


def test_find_text_takes_a_substring_or_a_pattern():
    document = Document.open(str(REPORT))
    found = document.find_text("4,285")
    assert [b.address for b in found] == ["256.13", "257.3#5/cell1,1", "258.4#3/cell5,1"]
    percent = document.find_text(re.compile(r"^\d+\.\d%$"), slides=[1])
    assert "256.29" in [b.address for b in percent]
    assert document.find_text("no such text") == []
    with pytest.raises(TypeError):
        document.find_text("")
    # Replacing through the blocks keeps the formatting (the figure stays bold).
    for block in found:
        document.resolve(block.address).text = block.text.replace("4,285", "4,310")
    assert document.find_text("4,285") == []
    assert document.resolve("256.13/p0/r0").bold


def test_a_blocks_text_keeps_a_line_break_as_set_text_writes_it():
    document = Document.new()
    slide = document.add_slide("Title Only")
    box = slide.add_textbox(914400, 914400, 4572000, 914400, "one\vtwo\nthree")
    (block,) = [b for b in document.outline_blocks() if b.address == box.id]
    assert block.text == "one\vtwo\nthree" == box.text_frame.text
    assert box.text == "one\ntwo\nthree"  # E0's reading, kept
    document.resolve(block.address).text = block.text
    assert len(box.paragraphs) == 2


# -- the warning ----------------------------------------------------------------------------


@pytest.mark.parametrize("text", [
    "11\\.9%", "\\+0.6pt YoY", "\\- item", "\\# not a heading", "a\\*b", "R\\&D",
    "**12.1%**", "Revenue **4,310** 億円", "line one\n2\\) second",
])
def test_text_copied_from_the_outline_warns_and_is_written_as_given(text):
    document = Document.open(str(REPORT))
    shape = document.shape("256.29")
    with pytest.warns(MarkdownEscapeWarning, match="256.29"):
        shape.set_text(text)
    assert shape.text == text
    cell = document.shape("257.3#5").table.cell(1, 1)
    with pytest.warns(MarkdownEscapeWarning, match="cell1,1"):
        cell.text = text
    if "\n" not in text:
        with pytest.warns(MarkdownEscapeWarning):
            document.resolve("256.13/p0").text = text


@pytest.mark.parametrize("text", [
    "11.9%", "+0.6pt YoY", "C:\\Users\\ada\\deck.pptx", "\\\\server\\share", "3*4**2 = 48",
    "Revenue*", "Note** see below", "a_b_c", "x \\ y", "regex \\d+\\.\\d", "50% \\ 2",
    "**", "Q3 ** Q4", "price: $5 (approx.)", "2025.10.1", "1. not escaped",
])
def test_real_text_does_not_warn(text):
    assert outline_escapes(text) == []
    document = Document.open(str(REPORT))
    with warnings.catch_warnings():
        warnings.simplefilter("error", MarkdownEscapeWarning)
        document.shape("256.29").set_text(text)


def test_an_escape_the_text_already_has_does_not_warn():
    document = Document.open(str(REPORT))
    shape = document.shape("256.29")
    with pytest.warns(MarkdownEscapeWarning):
        shape.set_text("a\\*b")
    with warnings.catch_warnings():
        warnings.simplefilter("error", MarkdownEscapeWarning)
        shape.set_text("a\\*b and c")


def test_the_warning_names_the_callers_line():
    document = Document.open(str(REPORT))
    with pytest.warns(MarkdownEscapeWarning) as record:
        document.shape("256.29").set_text("**x**")
    assert record[0].filename == __file__


# -- the outline's separators ----------------------------------------------------------------


def test_separators_without_ids_depend_on_structure_only():
    """``ids=False`` puts ``---`` before every block but a slide's first, whatever the text.

    The pilot saw ``---`` "appear" between 256.29 and 256.30 after editing 256.29: that was
    ``difflib``'s autojunk, which treats lines as frequent as ``---`` and blank lines as junk
    and so shows them removed and added around a changed line.  Line by line, the
    separators of the two outlines are the same.
    """
    document = Document.open(str(REPORT))
    before = document.to_outline(ids=False).splitlines()
    document.shape("256.29").set_text("12.1%")
    document.shape("256.30").set_text("+0.8pt YoY")
    after = document.to_outline(ids=False).splitlines()
    assert len(before) == len(after)
    assert [k for k, line in enumerate(before) if line == "---"] == \
        [k for k, line in enumerate(after) if line == "---"]
    changed = [(a, b) for a, b in zip(before, after) if a != b]
    assert changed == [("**11.9%**", "**12.1%**"), ("+0.6pt YoY", "+0.8pt YoY")]
    # What the pilot saw, and how to diff without it.
    noisy = list(difflib.unified_diff(before, after, lineterm="", n=0))
    assert "----" in noisy  # the separator, shown as removed
    quiet = difflib.SequenceMatcher(None, before, after, autojunk=False)
    assert [op for op in quiet.get_opcodes() if op[0] != "equal"] == [
        ("replace", k, k + 1, k, k + 1)
        for k, (a, b) in enumerate(zip(before, after)) if a != b]
