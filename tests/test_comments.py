"""Modern comments: threads on slides and shapes, replies, resolving, and the parts they live in
(what PowerPoint for Mac kept and wrote back when it re-saved them; see edit/comments.py)."""

from __future__ import annotations

import datetime as dt
import io
import zipfile

import pytest

from pptx_agent import Document
from pptx_agent.edit.comments import (COMMENT_REL_EXT_URI, CT_AUTHORS, CT_COMMENTS, P188,
                                      REL_AUTHORS, REL_COMMENTS)

T = dt.datetime(2026, 10, 7, 9, 0, tzinfo=dt.timezone.utc)


@pytest.fixture
def deck():
    deck = Document.new()
    first = deck.add_slide("Title Only")
    first.title = "Plan"
    first.add_shape("rect", 914400, 1828800, 2743200, 914400).text = "Budget"
    deck.add_slide("Title Only").title = "Detail"
    return deck


def _parts(deck) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(deck.to_bytes())) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_a_thread_on_a_shape_with_a_reply_resolved(deck):
    thread = deck.add_comment("256.3", "Is this the Q3 figure?", author="Dana Reviewer", date=T)
    assert thread.target == "256.3" and thread.address.startswith("c:")
    assert (thread.author, thread.initials, thread.created) == ("Dana Reviewer", "DR",
                                                                "2026-10-07T09:00:00Z")
    reply = deck.reply_to_comment(thread.address, "Yes, from the September close.",
                                  author="Sam Author", date=T + dt.timedelta(minutes=5))
    done = deck.resolve_comment(thread.address)
    assert done.done and [r.text for r in done.replies] == ["Yes, from the September close."]
    assert deck.comment(reply.address).id == thread.id          # a reply names its thread
    assert not deck.reopen_comment(thread.address).done
    assert deck.validate() == []


def test_the_parts_are_the_ones_powerpoint_writes(deck):
    deck.add_comment("s:256", "Title should say Q3.\nAnd the date.", author="Dana Reviewer",
                     date=T)
    deck.add_comment("256.3", "Source?", author="Dana Reviewer", date=T)
    parts = _parts(deck)
    types = parts["[Content_Types].xml"].decode()
    assert 'PartName="/ppt/comments/modernComment_100_0.xml" ContentType="' + CT_COMMENTS in types
    assert 'PartName="/ppt/authors.xml" ContentType="' + CT_AUTHORS in types
    assert REL_AUTHORS.encode() in parts["ppt/_rels/presentation.xml.rels"]
    assert REL_COMMENTS.encode() in parts["ppt/slides/_rels/slide1.xml.rels"]
    slide = parts["ppt/slides/slide1.xml"].decode()
    assert COMMENT_REL_EXT_URI in slide and "p188:commentRel" in slide
    comments = parts["ppt/comments/modernComment_100_0.xml"].decode()
    assert f'xmlns:p188="{P188}"' in comments
    assert "<pc:sldMkLst" in comments and \
        '<ac:spMk id="3" creationId="{00000000-0000-0000-0000-000000000000}"/>' in comments
    assert comments.count("<a:p>") == 3                         # two paragraphs, then one
    authors = parts["ppt/authors.xml"].decode()
    assert 'name="Dana Reviewer" initials="DR" userId="Dana Reviewer" providerId="None"' \
        in authors and authors.count("<p188:author ") == 1


def test_reading_back_and_listing_by_slide(deck):
    deck.add_comment("s:257", "Add a source.", author="Dana", date=T)
    deck.add_comment("s:256", "Shorter title.", author="Dana", date=T)
    again = Document.open(deck.to_bytes())
    assert [c.text for c in again.comments()] == ["Shorter title.", "Add a source."]
    assert [c.text for c in again.comments("s:257")] == ["Add a source."]


def test_edit_delete_and_undo(deck):
    thread = deck.add_comment("s:256", "Typo", author="Dana", date=T)
    reply = deck.reply_to_comment(thread.address, "Fixed", author="Sam", date=T)
    deck.edit_comment(reply.address, "Fixed, thanks")
    assert deck.comment(thread.address).replies[0].text == "Fixed, thanks"
    deck.delete_comment(reply.address)
    assert deck.comment(thread.address).replies == []
    deck.delete_comment(thread.address)
    assert deck.comments() == []
    assert not any("modernComment" in name for name in _parts(deck))
    assert COMMENT_REL_EXT_URI not in _parts(deck)["ppt/slides/slide1.xml"].decode()
    deck.undo()
    assert [c.text for c in deck.comments()] == ["Typo"]


def test_same_calls_same_bytes(deck):
    data = deck.to_bytes()
    outputs = []
    for _ in range(2):
        copy = Document.open(data)
        copy.add_comment("256.3", "Check", author="Dana", date=T)
        outputs.append(copy.to_bytes())
    assert outputs[0] == outputs[1]


def test_a_duplicated_slide_does_not_take_the_review_and_a_deleted_one_takes_it_along(deck):
    deck.add_comment("s:256", "Typo", author="Dana", date=T)
    copy = deck.duplicate_slide("s:256")
    assert deck.comments(copy) == [] and len(deck.comments()) == 1
    deck.delete_slide("s:256")
    assert deck.comments() == []
    assert not any("modernComment" in name for name in _parts(deck))
    assert deck.validate() == []


def test_errors(deck):
    with pytest.raises(ValueError):
        deck.add_comment("s:256", " ", author="Dana")
    with pytest.raises(ValueError):
        deck.add_comment("s:256", "x", author="")
    with pytest.raises(KeyError):
        deck.resolve_comment("c:00000000")


# -- the anchor's moniker follows the kind of drawing element ([MS-ODRAWXML] 2.29.3) --------------

FIXTURE = __import__("pathlib").Path(__file__).parent / "fixtures" / "authoring-integration.pptx"


@pytest.fixture
def kinds():
    """sp 256.2, pic 256.3, cxnSp 256.4, table 256.5, chart 256.6; a group 256.9 of 256.7, 256.8."""
    deck = Document.open(str(FIXTURE))
    slide = deck.slides[0]
    first = slide.add_shape("rect", 6000000, 1400000, 1200000, 600000)
    second = slide.add_shape("ellipse", 6000000, 2200000, 1200000, 600000)
    group = slide.group([first, second], name="Pair")
    assert (group.id, first.id) == ("256.9", "256.7")
    return deck


def _anchor(deck, comment) -> str:
    import re

    xml = _parts(deck)["ppt/comments/modernComment_100_0.xml"].decode()
    thread = xml[xml.index(comment.id):]
    return re.search(r"<ac:deMkLst [^>]*>(.*?)</ac:deMkLst>", thread).group(1)


def _monikers_only(anchor: str) -> list[str]:
    import re

    return re.findall(r"<(ac:\w+Mk) id=\"(\d+)\"", anchor)


@pytest.mark.parametrize("target, moniker", [
    ("256.2", ("ac:spMk", "2")),                 # a shape
    ("256.3", ("ac:picMk", "3")),                # a picture
    ("256.4", ("ac:cxnSpMk", "4")),              # a connector
    ("256.5", ("ac:graphicFrameMk", "5")),       # a table
    ("256.6", ("ac:graphicFrameMk", "6")),       # a chart
    ("256.9", ("ac:grpSpMk", "9")),              # a group
    ("256.7", ("ac:spMk", "7")),                 # a shape in a group: its own moniker only
], ids=["sp", "pic", "cxnSp", "table", "chart", "grpSp", "sp-in-group"])
def test_each_kind_of_shape_is_anchored_by_its_own_moniker(kinds, target, moniker):
    thread = kinds.add_comment(target, "Here?", author="Dana", date=T)
    anchor = _anchor(kinds, thread)
    assert anchor.startswith("<pc:docMk ") and 'sldId="256"' in anchor
    assert _monikers_only(anchor) == [moniker]
    # Measured: PowerPoint writes the nil GUID for a shape without a16:creationId.
    assert 'creationId="{00000000-0000-0000-0000-000000000000}"/>' in anchor
    again = Document.open(kinds.to_bytes())
    assert [c.target for c in again.comments()] == [target]
    assert kinds.validate() == []


def test_every_moniker_kind_is_read_including_a_group_chain(kinds):
    from lxml import etree

    from pptx_agent.oxml.xml import parse_xml

    data = kinds.to_bytes()
    part = "ppt/comments/modernComment_100_0.xml"
    for written, target in [('<ac:picMk id="3"/>', "256.3"), ('<ac:cxnSpMk id="4"/>', "256.4"),
                            ('<ac:graphicFrameMk id="6"/>', "256.6"),
                            ('<ac:grpSpMk id="9"/>', "256.9"),
                            ('<ac:grpSpMk id="9"/><ac:spMk id="8"/>', "256.8"),
                            ('<ac:spMk id="6"/>', "256.6")]:     # what earlier versions wrote
        copy = Document.open(data)
        copy.add_comment("256.2", "x", author="Dana", date=T)
        root = copy.package.tree(part)
        xml = etree.tostring(root).decode().replace(
            '<ac:spMk id="2" creationId="{00000000-0000-0000-0000-000000000000}"/>', written)
        root[:] = list(parse_xml(xml.encode()))
        copy.package.mark_dirty(part)
        assert Document.open(copy.to_bytes()).comments()[0].target == target, written


def test_the_moniker_carries_the_shapes_creation_id(kinds):
    from pptx_agent.edit.ids import cnv_pr
    from pptx_agent.oxml.xml import parse_xml

    chart = kinds.shape("256.6")
    properties = cnv_pr(chart._element)
    properties.append(parse_xml(
        b'<a:extLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        b'<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId '
        b'xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" '
        b'id="{0D7E1C55-3A9B-4C2E-8F00-2B6A1D9E4C11}"/></a:ext></a:extLst>'))
    chart._slide._touch()
    thread = kinds.add_comment("256.6", "Q3?", author="Dana", date=T)
    assert ('<ac:graphicFrameMk id="6" creationId="{0D7E1C55-3A9B-4C2E-8F00-2B6A1D9E4C11}"/>'
            in _anchor(kinds, thread))
