"""ppt_read_slides(detail="svg"): pptx2svg's compact agent view with this library's ids (LR3).

Skipped where pptx2svg has no agent view yet (convert_pptx_to_agent_svg).
"""

from __future__ import annotations

import datetime as dt
import re

import pytest

pptx2svg = pytest.importorskip("pptx2svg")
if not hasattr(pptx2svg, "convert_pptx_to_agent_svg"):
    pytest.skip("this pptx2svg has no agent view", allow_module_level=True)

from ooxml_edit.tools import Limits, Toolbox  # noqa: E402

from pptx_agent import Document  # noqa: E402
from pptx_agent.tools import FORMAT, GROUPS, TOOLS  # noqa: E402

CLOCK = lambda: dt.datetime(2026, 10, 7, 12, 0, tzinfo=dt.timezone.utc)  # noqa: E731


def _deck() -> bytes:
    deck = Document.new()
    for title in ("One", "Two", "Three"):
        slide = deck.add_slide("Title Only")
        slide.title = title
    slide = deck.slides[1]
    slide.add_shape("roundRect", 100 * 12700, 200 * 12700, 120 * 12700, 40 * 12700,
                    text="Discover").fill = "accent2"
    return deck.to_bytes()


def test_the_svg_view_carries_addresses_the_tools_accept():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=CLOCK)
        session.open(_deck(), "deck.pptx")
        result = toolbox.dispatch(session, "ppt_read_slides",
                                  {"doc": "d1", "detail": "svg", "slides": [2]})
        assert result.ok, result.to_json()
        page = result.data[0]
        svg = page["svg"]
        assert page["slide"] == "s:257" and 'viewBox="0 0 960 540"' in svg
        assert "<image" not in svg and "base64" not in svg and "@font-face" not in svg
        ids = re.findall(r'data-id="(257\.\d+)"', svg)
        assert ids and "data-pptx-id" not in svg
        box = next(i for i in ids if 'data-fill="accent2"' in svg.split(f'data-id="{i}"')[1][:300])
        moved = toolbox.dispatch(session, "ppt_set_shape",
                                 {"doc": "d1", "items": [{"target": box, "x": 300}]})
        assert moved.ok
        again = toolbox.dispatch(session, "ppt_read_slides",
                                 {"doc": "d1", "detail": "svg", "slides": [2]})
        assert 'x="300"' in again.data[0]["svg"]


def test_the_svg_view_pages_by_slide():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        whole = toolbox.session(clock=CLOCK)
        whole.open(_deck(), "deck.pptx")
        sizes = [p["chars"] for p in toolbox.dispatch(whole, "ppt_read_slides",
                                                      {"doc": "d1", "detail": "svg"}).data]
        limit = 2000 + max(sizes) + 1           # room for one slide a page
        session = toolbox.session(clock=CLOCK, limits=Limits(max_result_chars=limit))
        session.open(_deck(), "deck.pptx")
        first = toolbox.dispatch(session, "ppt_read_slides", {"doc": "d1", "detail": "svg"})
        assert first.ok and first.next_cursor and first.total == 3
        seen = [p["n"] for p in first.data]
        cursor = first.next_cursor
        while cursor:
            page = toolbox.dispatch(session, "ppt_read_slides",
                                    {"doc": "d1", "detail": "svg", "cursor": cursor})
            seen += [p["n"] for p in page.data]
            cursor = page.next_cursor
        assert seen == [1, 2, 3]
