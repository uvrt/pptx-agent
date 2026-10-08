"""Connector geometry against what PowerPoint wrote, and connectors kept on their shapes.

The router and the frame spelling are held to :mod:`connector_cases`, routes PowerPoint
made itself.  The rest checks the property every route must have whatever PowerPoint would
have chosen: both ends on their sites, before and after the shapes move.
"""

from __future__ import annotations

import math

import pytest

from connector_cases import DIVERGENT, MEASURED
from pptx_agent import Document
from pptx_agent.edit import connectors
from pptx_agent.edit.connectors import End, axis, bent_frame, endpoints, frame_for, route
from pptx_agent.edit.creating import container_to_slide, current_ends, shape_sites
from pptx_agent.edit.presets import adjustment_defaults, connection_sites
from pptx_agent.oxml.xml import PRESET_GEOMETRIES, qn

EMU = 914400


def _end(shape: tuple, site: int) -> End:
    preset, x, y, cx, cy = shape
    angle, sx, sy = connection_sites(preset, cx, cy)[site]
    return End((x + sx, y + sy), axis(angle), (x, y, x + cx, y + cy))


def _frame_tuple(frame) -> tuple:
    return (frame.preset, frame.x, frame.y, frame.cx, frame.cy, frame.rot, int(frame.flip_h),
            int(frame.flip_v), tuple(value for _, value in frame.adjustments))


def _close(mine: tuple, theirs: tuple, tolerance: int = 2) -> bool:
    return (mine[0] == theirs[0] and mine[5:8] == theirs[5:8]
            and all(abs(a - b) <= tolerance for a, b in zip(mine[1:5], theirs[1:5]))
            and len(mine[8]) == len(theirs[8])
            and all(abs(a - b) <= tolerance for a, b in zip(mine[8], theirs[8])))


# -- what PowerPoint wrote -------------------------------------------------------------------


@pytest.mark.parametrize("case", MEASURED, ids=lambda case: f"{case[0][0]}{case[2]}-{case[1][0]}{case[3]}-{case[4][0]}")
def test_the_router_writes_what_powerpoint_wrote(case):
    start, end, start_site, end_site, expected = case
    kind = "straight" if expected[0] == "straightConnector1" else "elbow"
    frame = frame_for(kind, _end(start, start_site), _end(end, end_site))
    assert _close(_frame_tuple(frame), expected), (_frame_tuple(frame), expected)


def test_most_measured_routes_are_reproduced():
    """The record: of the distinct routes measured, these many are written identically."""
    assert len(MEASURED) >= 169 and len(DIVERGENT) <= 8


@pytest.mark.parametrize("case", MEASURED + DIVERGENT)
def test_every_route_starts_and_ends_on_its_sites(case):
    """Whatever PowerPoint chose, a route must leave one site and arrive at the other."""
    start, end, start_site, end_site, _ = case
    a, b = _end(start, start_site), _end(end, end_site)
    points = route(a, b)
    assert points[0] == pytest.approx(a.point) and points[-1] == pytest.approx(b.point)
    for p, q in zip(points, points[1:]):  # every run is horizontal or vertical
        assert abs(p[0] - q[0]) < 1e-6 or abs(p[1] - q[1]) < 1e-6
    frame = bent_frame(points)
    drawn = endpoints(frame.x, frame.y, frame.cx, frame.cy, frame.rot, frame.flip_h,
                      frame.flip_v)
    # Within the one point PowerPoint's own minimum extent can put an end off its site.
    assert math.dist(drawn[0], a.point) <= connectors.MIN_EXTENT + 2
    assert math.dist(drawn[1], b.point) <= connectors.MIN_EXTENT + 2


def test_a_curved_connector_takes_the_elbow_frame():
    a = _end(("rect", 0, 0, 1524000, 1016000), 3)
    b = _end(("rect", 3048000, 2032000, 1270000, 762000), 2)
    elbow, curved = frame_for("elbow", a, b), frame_for("curved", a, b)
    assert curved.preset == elbow.preset.replace("bent", "curved")
    assert _frame_tuple(curved)[1:] == _frame_tuple(elbow)[1:]


def test_frames_draw_their_routes():
    """Reading a frame back (preset, flips, rotation, adjustments) gives the route's corners."""
    for case in MEASURED:
        start, end, start_site, end_site, expected = case
        if expected[0] == "straightConnector1":
            continue
        points = route(_end(start, start_site), _end(end, end_site))
        drawn = _local_path(bent_frame(points))
        assert len(drawn) == len(points)
        for got, wanted in zip(drawn, points):
            assert math.dist(got, wanted) <= connectors.MIN_EXTENT + 2, (case, drawn, points)


def _local_path(frame) -> list[tuple[float, float]]:
    """The corners a bentConnector frame draws, on the slide."""
    w, h = frame.cx, frame.cy
    values = [v / 100000 for _, v in frame.adjustments]
    count = int(frame.preset[-1])
    if count == 2:
        local = [(0, 0), (w, 0), (w, h)]
    elif count == 3:
        local = [(0, 0), (w * values[0], 0), (w * values[0], h), (w, h)]
    elif count == 4:
        x1, y2 = w * values[0], h * values[1]
        local = [(0, 0), (x1, 0), (x1, y2), (w, y2), (w, h)]
    else:
        x1, y2, x3 = w * values[0], h * values[1], w * values[2]
        local = [(0, 0), (x1, 0), (x1, y2), (x3, y2), (x3, h), (w, h)]
    angle = math.radians(frame.rot / 60000)
    cos, sin = math.cos(angle), math.sin(angle)
    out = []
    for x, y in local:
        x = w - x if frame.flip_h else x
        y = h - y if frame.flip_v else y
        dx, dy = x - w / 2, y - h / 2
        out.append((frame.x + w / 2 + dx * cos - dy * sin, frame.y + h / 2 + dx * sin + dy * cos))
    return out


def test_sites_match_the_specification():
    assert len(connection_sites("rect", 100, 50)) == 4
    assert [(a, x, y) for a, x, y in connection_sites("rect", 100, 50)] == [
        (16200000, 50, 0), (10800000, 0, 25), (5400000, 50, 50), (0, 100, 25)]
    ellipse = connection_sites("ellipse", 1270000, 1270000)
    assert len(ellipse) == 8
    # PowerPoint put the ellipse's upper-left site 185,987 EMU in from each side.
    assert round(ellipse[1][1]) == 185987 and round(ellipse[1][2]) == 185987
    # Sites move with adjustments: a down arrow's side sites sit at its head's shoulder.
    default = connection_sites("downArrow", 1000, 2000)[1][2]
    moved = connection_sites("downArrow", 1000, 2000, {"adj2": 20000})[1][2]
    assert default != moved
    for preset in PRESET_GEOMETRIES:
        for angle, x, y in connection_sites(preset, 1524000, 1016000):
            assert all(math.isfinite(v) for v in (angle, x, y)), preset
    assert adjustment_defaults("roundRect") == (("adj", 16667),)
    assert adjustment_defaults("bentConnector3") == (("adj1", 50000),)


# -- connectors on slides -----------------------------------------------------------------------


@pytest.fixture
def two_shapes(product_page):
    document = Document.open(product_page.read_bytes())
    slide = document.slides[0]
    a = slide.add_shape("rect", 1 * EMU, 1 * EMU, 2 * EMU, EMU)
    b = slide.add_shape("ellipse", 6 * EMU, 3 * EMU, EMU, EMU)
    return document, slide, a, b


def _ends_on_slide(connector) -> list[tuple[float, float]]:
    ends = current_ends(connector._element)
    parent = connector._element.getparent()
    return [container_to_slide(parent, *point) for point in ends]


def assert_attached(document, connector_id, tolerance=connectors.MIN_EXTENT + 2):
    connector = document.shape(connector_id)
    begin, end = connector.begin_connection, connector.end_connection
    drawn = _ends_on_slide(connector)
    for (shape, site), point in zip((begin, end), drawn):
        wanted = shape_sites(shape)[site].point
        assert math.dist(point, wanted) <= tolerance, (connector.id, point, wanted)


@pytest.mark.parametrize("kind", ["straight", "elbow", "curved"])
def test_every_site_pair_meets(kind, two_shapes):
    document, slide, a, b = two_shapes
    made = []
    for start in range(4):
        for end in range(0, 8, 3):
            made.append(slide.add_connector(kind, (a, start), (b, end)).id)
    for identifier in made:
        assert_attached(document, identifier)


@pytest.mark.parametrize("change", ["move", "resize", "rotate", "flip", "adjust", "preset"])
def test_connectors_follow_their_shapes(change, two_shapes):
    document, slide, a, b = two_shapes
    made = [slide.add_connector(kind, (a, 3), (b, 2)).id for kind in ("straight", "elbow", "curved")]
    made.append(slide.add_connector("elbow", (b, 0), (a, 2)).id)
    target = document.shape(b.id)
    if change == "move":
        target.move_by(-4 * EMU, 2 * EMU)
    elif change == "resize":
        target.width = 3 * EMU
        document.shape(a.id).height = 2 * EMU
    elif change == "rotate":
        target.rotation = 90
        document.shape(a.id).rotation = 45
    elif change == "flip":
        document.shape(a.id).flip_h = True
    elif change == "adjust":
        document.shape(a.id).preset = "downArrow"
        document.shape(a.id).adjustments["adj2"] = 20000
    else:
        document.shape(a.id).preset = "flowChartDecision"
    for identifier in made:
        assert_attached(document, identifier)


def test_a_move_is_one_undo_step_with_its_connectors(two_shapes):
    document, slide, a, b = two_shapes
    connector = slide.add_connector("elbow", (a, 3), (b, 2))
    before = document.to_bytes()
    document.shape(b.id).left = 2 * EMU
    moved = document.to_bytes()
    assert moved != before
    assert document.undo()
    assert document.to_bytes() == before
    assert document.redo()
    assert document.to_bytes() == moved
    assert_attached(document, connector.id)


def test_connectors_follow_a_moved_group_and_a_moved_child(two_shapes):
    document, slide, a, b = two_shapes
    c = slide.add_shape("rect", 6 * EMU, 5 * EMU, EMU, EMU // 2)
    group = slide.group([b, c])
    connector = slide.add_connector("elbow", (a, 3), (document.shape(b.id), 1))
    document.shape(group.id).move_by(-EMU, EMU // 2)
    assert_attached(document, connector.id)
    document.shape(b.id).move_by(EMU // 2, 0)
    assert_attached(document, connector.id)


def test_a_connector_inside_a_scaled_group(two_shapes):
    document, slide, a, b = two_shapes
    c = slide.add_shape("rect", 6 * EMU, 5 * EMU, EMU, EMU // 2)
    group = slide.group([b, c])
    group = document.shape(group.id)
    group.width = group.width * 2  # child space now stretched two to one
    connector = group.add_connector("curved", (document.shape(b.id), 4),
                                    (document.shape(c.id), 0))
    assert connector.parent_group.id == group.id
    assert_attached(document, connector.id, tolerance=2 * connectors.MIN_EXTENT + 2)
    document.shape(c.id).move_by(EMU // 2, EMU // 2)
    assert_attached(document, connector.id, tolerance=2 * connectors.MIN_EXTENT + 2)


def test_free_ends_and_connecting_later(two_shapes):
    document, slide, a, b = two_shapes
    line = slide.add_connector("straight", (EMU, 6 * EMU), (4 * EMU, 7 * EMU))
    assert line.begin_connection is None and line.end_connection is None
    begin, end = _ends_on_slide(line)
    assert begin == pytest.approx((EMU, 6 * EMU)) and end == pytest.approx((4 * EMU, 7 * EMU))
    assert line.flip_h is False and line.flip_v is False
    back = slide.add_connector("straight", (4 * EMU, 7 * EMU), (EMU, 6 * EMU))
    assert back.flip_h and back.flip_v
    elbow = slide.add_connector("elbow", (EMU, 6 * EMU), (4 * EMU, 7 * EMU))
    assert document.shape(elbow.id).preset == "bentConnector3"
    connector = document.shape(line.id)
    connector.connect(end=(document.shape(b.id), 4))
    assert connector.end_connection[0].id == b.id and connector.begin_connection is None
    assert _ends_on_slide(document.shape(line.id))[0] == pytest.approx((EMU, 6 * EMU), abs=1)
    connector = document.shape(line.id)
    connector.connect(begin=(document.shape(a.id), 2))
    assert_attached(document, line.id)
    connector = document.shape(line.id)
    connector.disconnect("begin")
    assert document.shape(line.id).begin_connection is None
    assert document.shape(line.id).end_connection is not None
    element = document.shape(line.id)._element
    assert element.find(".//" + qn("a:stCxn")) is None


def test_deleting_a_shape_detaches_its_connectors(two_shapes):
    document, slide, a, b = two_shapes
    connector = slide.add_connector("elbow", (a, 3), (b, 2))
    where = _ends_on_slide(document.shape(connector.id))
    document.shape(b.id).delete()
    kept = document.shape(connector.id)
    assert kept.begin_connection[0].id == a.id and kept.end_connection is None
    assert _ends_on_slide(kept) == where  # the geometry stays where it was
    reopened = Document.open(document.to_bytes())
    assert reopened.shape(connector.id).end_connection is None


def test_connector_ends_are_checked(two_shapes):
    document, slide, a, b = two_shapes
    with pytest.raises(IndexError):
        slide.add_connector("elbow", (a, 4), (b, 0))
    with pytest.raises(TypeError):
        slide.add_connector("elbow", (a, "3"), (b, 0))
    with pytest.raises(TypeError):
        slide.add_connector("elbow", "nowhere", (b, 0))
    with pytest.raises(ValueError):
        slide.add_connector("zigzag", (a, 0), (b, 0))
    table = slide.add_table(2, 2, 0, 0, EMU, EMU)
    with pytest.raises(ValueError):
        slide.add_connector("elbow", (table, 0), (b, 0))
    other = document.slides[-1] if len(document.slides) > 1 else document.add_slide()
    stranger = other.add_shape("rect", 0, 0, EMU, EMU)
    with pytest.raises(ValueError):
        slide.add_connector("straight", (stranger, 0), (b, 0))


def test_a_shape_with_a_duplicated_raw_id_cannot_be_named_by_a_connector(financial_report):
    document = Document.open(financial_report.read_bytes())
    slide = document.slides[1]
    raws = {}
    for shape in slide.shapes:
        raws.setdefault(shape._element.find(".//" + qn("p:cNvPr")).get("id"), []).append(shape)
    duplicated = next(shapes for shapes in raws.values() if len(shapes) > 1)
    sp = next((s for s in duplicated if s.kind == "shape"), None)
    if sp is None:
        pytest.skip("no autoshape among the duplicated ids")
    new = slide.add_shape("rect", 0, 0, EMU, EMU)
    with pytest.raises(ValueError, match="not unique"):
        slide.add_connector("straight", (sp, 0), (new, 0))


def test_connection_sites_follow_flips_and_rotation(two_shapes):
    document, slide, a, b = two_shapes
    shape = document.shape(a.id)  # 2 x 1 inch at (1, 1)
    assert shape.connection_sites == [(2 * EMU, EMU), (EMU, 3 * EMU // 2),
                                      (2 * EMU, 2 * EMU), (3 * EMU, 3 * EMU // 2)]
    shape.rotation = 90
    sites = document.shape(a.id).connection_sites
    assert sites[0] == (5 * EMU // 2, 3 * EMU // 2)  # the top site now points right
    assert shape_sites(document.shape(a.id))[0].direction == (1, 0)
    document.shape(a.id).rotation = 0
    document.shape(a.id).flip_h = True
    assert shape_sites(document.shape(a.id))[1].direction == (1, 0)  # "left" faces right
    assert document.shape(b.id).connection_sites[2] == (6 * EMU, 7 * EMU // 2)
