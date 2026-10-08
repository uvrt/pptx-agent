"""Where shapes are drawn: rotation, flips, groups and connector routes (the trial's N2).

p4 run 2 drew an elbow connector from one box's bottom to another's top.  PowerPoint
writes such a connector with ``rot="5400000"`` and its width and height swapped, so its
frame ran from y=586,000 to y=8,314,000 on a 6,858,000-high slide, and ``overflows()``
called it off the slide while the line ran between the boxes.
"""

from __future__ import annotations

import math

import pytest

from connector_cases import MEASURED
from pptx_agent import Document
from pptx_agent.edit import connectors
from pptx_agent.edit.connectors import End, axis
from pptx_agent.edit.presets import connection_sites

EMU = 914400


def _slide():
    deck = Document.new()
    return deck, deck.add_slide("Blank")


def _near(a, b, tolerance=2.0):
    return all(abs(x - y) <= tolerance for x, y in zip(a, b))


def test_bounds_is_the_frame_as_one_tuple():
    deck, slide = _slide()
    box = slide.add_shape("rect", EMU, 2 * EMU, 3 * EMU, EMU)
    assert box.bounds == (box.left, box.top, box.width, box.height) == (EMU, 2 * EMU, 3 * EMU, EMU)
    assert box.route is None


def test_a_rotated_shape_is_drawn_turned_about_its_centre():
    deck, slide = _slide()
    box = slide.add_shape("rect", EMU, EMU, 4 * EMU, EMU)
    box.rotation = 90
    assert box.bounds == (EMU, EMU, 4 * EMU, EMU)               # the frame does not turn
    assert _near(box.drawn_bounds, (int(2.5 * EMU), -EMU // 2, EMU, 4 * EMU))
    box.rotation = 45
    left, top, width, height = box.drawn_bounds
    assert abs(width - 5 * EMU / math.sqrt(2)) <= 2 and width == height


def test_a_shape_in_a_rotated_group_turns_with_the_group():
    deck, slide = _slide()
    a = slide.add_shape("rect", 0, 0, 2 * EMU, EMU)
    b = slide.add_shape("rect", 2 * EMU, 0, 2 * EMU, EMU)
    group = slide.group([a, b])
    group._element.find(".//{*}grpSpPr/{*}xfrm").set("rot", str(90 * 60000))
    # The group's frame is (0, 0, 4, 1) inches; turned 90 degrees about (2, 0.5) the left
    # box lands above the centre, from y = -1.5 to 0.5 in.
    assert _near(slide.shape(a.id).drawn_bounds, (int(1.5 * EMU), int(-1.5 * EMU), EMU, 2 * EMU))


def test_the_trial_elbow_is_not_off_the_slide():
    """p4 run 2: pick-and-pack's bottom to ship-order's top, the two facing away."""
    deck, slide = _slide()
    pick = slide.add_shape("roundRect", 8_228_000, 2_350_000, 3_464_000, 900_000, text="Pick")
    ship = slide.add_shape("roundRect", 500_000, 5_650_000, 3_464_000, 900_000, text="Ship")
    elbow = slide.add_connector("elbow", (pick, "bottom"), (ship, "top"))
    assert elbow.rotation in (90.0, 270.0)
    frame_bottom = elbow.top + elbow.height
    assert frame_bottom > deck.slide_size[1] or elbow.top < 0   # the frame is far out
    route = elbow.route
    assert _near(route[0], pick.connection_site("bottom")[0].connection_sites[2])
    assert _near(route[-1], ship.connection_sites[0])
    left, top, width, height = elbow.drawn_bounds
    assert top >= pick.top + pick.height - 2 and top + height <= ship.top + 2
    assert not [o for o in deck.overflows() if o.kind == "off_slide"]


@pytest.mark.parametrize("case", MEASURED[::7])
def test_a_route_is_the_routers_polyline(case):
    """The route read back from the written frame is the polyline the router chose."""
    start, end, start_site, end_site, expected = case
    deck, slide = _slide()
    shapes = []
    for preset, x, y, cx, cy in (start, end):
        shapes.append(slide.add_shape(preset, x, y, cx, cy))
    kind = "straight" if expected[0] == "straightConnector1" else "elbow"
    connector = slide.add_connector(kind, (shapes[0], start_site), (shapes[1], end_site))

    def end_of(shape, site):
        preset, x, y, cx, cy = shape
        angle, sx, sy = connection_sites(preset, cx, cy)[site]
        return End((x + sx, y + sy), axis(angle), (x, y, x + cx, y + cy))

    if kind == "elbow":
        wanted = connectors.route(end_of(start, start_site), end_of(end, end_site))
    else:
        wanted = [end_of(start, start_site).point, end_of(end, end_site).point]
    got = connector.route
    assert len(got) == len(wanted)
    # Within the point the minimum extent can put the far end off its site.
    assert all(_near(g, w, 12701) for g, w in zip(got, wanted)), (got, wanted)


def test_flips_put_the_begin_first():
    deck, slide = _slide()
    line = slide.add_connector("straight", (5 * EMU, 3 * EMU), (EMU, EMU))
    assert line.flip_h and line.flip_v
    assert _near(line.route[0], (5 * EMU, 3 * EMU)) and _near(line.route[-1], (EMU, EMU))


def test_a_curved_connector_is_sampled_along_its_curve():
    deck, slide = _slide()
    a = slide.add_shape("rect", EMU, EMU, EMU, EMU)
    b = slide.add_shape("rect", 5 * EMU, 3 * EMU, EMU, EMU)
    curve = slide.add_connector("curved", (a, "right"), (b, "left"))
    points = curve.route
    assert len(points) > 10
    assert _near(points[0], (2 * EMU, 1.5 * EMU)) and _near(points[-1], (5 * EMU, 3.5 * EMU))
    left, top, width, height = curve.drawn_bounds
    assert left >= 2 * EMU - 2 and left + width <= 5 * EMU + 2


def test_a_table_is_drawn_from_its_grid():
    deck, slide = _slide()
    frame = slide.add_table(2, 2, EMU, EMU, 4 * EMU, EMU)
    frame._element.find(".//{*}xfrm/{*}ext").set("cy", str(20 * EMU))   # a stale frame
    assert frame.drawn_bounds == (EMU, EMU, 4 * EMU, EMU)
    assert not [o for o in deck.overflows() if o.kind == "off_slide"]
