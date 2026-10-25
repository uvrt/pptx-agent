"""Round 10, O1 and O3: one call's slides drawn over the workers at once, and images kept
by their content, so a slide an edit left as it was is not drawn again.

The images must be what the serial renderer draws, byte for byte (``render_slides``, one
worker, the deck opened as a :class:`~pptx_agent.Document`), whichever worker drew them and
whether they came from the store."""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from ooxml_edit.tools import InProcess, Limits, Toolbox

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from pptx_agent.tools.render import raster_key, render_slides, render_views

pytest.importorskip("pptx2svg")
pytest.importorskip("resvg_py")

FIXTURES = Path(__file__).parent / "fixtures"


def _deck(count: int = 6) -> bytes:
    deck = Document.new()
    for number in range(1, count + 1):
        slide = deck.add_slide("Title and Content" if number % 2 else "Title Only")
        slide.title = f"Slide {number}: routes and margins"
    return deck.to_bytes()


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=[],
                   limits=options.pop("limits", Limits(image_budget=None)), **options)


def _render(box, session, slides, width=320, **more):
    result = box.dispatch(session, "render", {"doc": "d1", "slides": slides, "width": width,
                                              **more})
    assert result.ok, result.error
    return result


def test_a_call_drawn_over_the_workers_is_the_serial_renderers_byte_for_byte():
    data = _deck()
    expected = render_slides(data, [1, 2, 3, "L2"], 320, [])
    with _box(workers=2) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        result = _render(box, session, [1, 2, 3, "L2"])
        assert box.parallelism == 2
    assert [image.data for image in result.images] == expected
    assert result.data["cached"] == []


def test_a_repeat_is_answered_from_the_store_and_an_edit_redraws_only_its_slide():
    data = _deck()
    with _box(workers=2) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        first = _render(box, session, [1, 2, 3, 4])
        again = _render(box, session, [1, 2, 3, 4])
        assert again.data["cached"] == [1, 2, 3, 4]
        assert [i.data for i in again.images] == [i.data for i in first.images]
        edit = box.dispatch(session, "replace_text", {"doc": "d1", "find": "Slide 3:",
                                                      "replace": "Slide three:",
                                                      "expect": "one"})
        assert edit.ok, edit.error
        after = _render(box, session, [1, 2, 3, 4])
        deck = session.entry("d1").document
        expected = render_slides(deck.to_bytes(), [1, 2, 3, 4], 320, [])
    assert after.data["cached"] == [1, 2, 4]
    assert [i.data for i in after.images] == expected
    assert after.images[2].data != first.images[2].data


def test_the_store_is_the_toolboxs_and_bounded_in_bytes():
    data = _deck(4)
    with _box(workers=0, limits=Limits(image_budget=None, image_cache_bytes=1)) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        _render(box, session, [1, 2])
        assert len(box.images) == 0          # nothing fits: nothing kept, nothing served
        assert _render(box, session, [1, 2]).data["cached"] == []
    with _box(workers=0) as box:
        one, two = box.session(), box.session()
        one.open(data, "a.pptx")
        two.open(data, "b.pptx")
        _render(box, one, [1, 2])
        # The same deck in another session: the same content, the same images.
        assert _render(box, two, [1, 2]).data["cached"] == [1, 2]
        assert box.images.size <= box.limits.image_cache_bytes


def test_a_sheet_reuses_the_tiles_a_render_drew():
    data = _deck(4)
    with _box(workers=2) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        _render(box, session, [1, 2], width=200)
        result = _render(box, session, [1, 2, 3], sheet={"columns": 3, "width": 200})
    assert result.data["cached"] is False
    assert result.data["cached_tiles"] == [1, 2]


def test_a_render_over_the_budget_is_refused_before_anything_is_drawn():
    data = _deck(4)
    with _box(limits=Limits(image_budget=1)) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        result = box.dispatch(session, "render", {"doc": "d1", "slides": [1, 2]})
        assert box._pool is None             # no worker was asked for anything
        sheet = box.dispatch(session, "render", {"doc": "d1", "slides": [1, 2],
                                                 "sheet": {}})
        assert sheet.ok                      # one image: within the budget
    assert not result.ok and result.error.code == "limit"


def test_render_views_draws_only_what_the_caller_does_not_hold():
    data = _deck(2)
    first = render_views(data, [1, 2], 320, [])
    assert [drawn for *_, drawn in first["views"]] == [True, True]
    keys = [key for key, *_ in first["views"]]
    again = render_views(data, [2, 1], 320, [], known=[keys[0]])
    assert again["views"][1] == (keys[0], None, [], False)
    assert again["views"][0][:2] == (keys[1], first["views"][1][1])
    assert first["fonts"] == {}


def test_a_key_depends_on_the_svg_and_the_fingerprint():
    assert raster_key("<svg/>", "a") != raster_key("<svg />", "a")
    assert raster_key("<svg/>", "a") != raster_key("<svg/>", "b")
    assert raster_key("<svg/>", "a") == raster_key("<svg/>", "a")


def _embedding_deck() -> Path:
    for path in sorted(FIXTURES.glob("*.pptx")):
        with zipfile.ZipFile(path) as archive:
            if b"embeddedFontLst" in archive.read("ppt/presentation.xml"):
                return path
    pytest.skip("no fixture deck embeds fonts")


def test_embedded_faces_are_decoded_once_and_handed_to_the_other_workers(monkeypatch):
    path = _embedding_deck()
    data = path.read_bytes()
    count = len(Document.open(data).slides)
    slides = list(range(1, min(count, 4) + 1))
    expected = render_slides(data, slides, 320, [])
    asked: list = []
    import pptx_agent.tools.shared_tools as shared_tools

    real = shared_tools.render_views

    def watched(*args):
        asked.append(len(args[-1]))           # the decoded faces each task was handed
        return real(*args)

    monkeypatch.setattr(shared_tools, "render_views", watched)
    with _box(runner=InProcess(2)) as box:
        session = box.session()
        session.open(data, path.name)
        # In-process: the decoded faces of this process would serve every task; forget them
        # before each, as a worker process would not have them.
        from ooxml_common.fonts import embedded

        embedded.forget_decoded()
        result = _render(box, session, slides)
        decoded = len(box.decoded_fonts)
    assert [image.data for image in result.images] == expected
    if len(slides) > 1 and decoded:
        assert asked[0] == 0 and all(handed == decoded for handed in asked[1:])


def test_bytes_for_a_version_are_made_once():
    data = _deck(2)
    with _box(workers=0) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        entry = session.entry("d1")
        assert entry.snapshot() is entry.snapshot()
        before = entry.snapshot()
        box.dispatch(session, "replace_text", {"doc": "d1", "find": "Slide 1:",
                                               "replace": "First:", "expect": "one"})
        assert entry.snapshot() is not before
        assert Document.open(BytesIO(entry.snapshot()).getvalue()).slides[0].title.startswith(
            "First")
