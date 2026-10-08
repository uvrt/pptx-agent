"""Scales, align and distribute, stack/column/grid and label placement (LP1, LP21, LP22)."""

from __future__ import annotations

import math

import pytest

from pptx_agent import Document
from pptx_agent.edit.arrange import align, distribute, drawn_box, place
from pptx_agent.edit.layout import LayoutError, column, grid, place_labels, stack
from pptx_agent.edit.scales import (BandScale, DateScale, LinearScale, ScaleError,
                                    nice_step, scale_from)

PT = 12700


def pts(box):
    return tuple(round(v / PT, 2) for v in box)


@pytest.fixture
def slide():
    deck = Document.new()
    page = deck.add_slide("Title Only")
    page.title = "Layout"
    return page


def boxes(slide, specs):
    return [slide.add_shape("rect", round(x * PT), round(y * PT), round(w * PT), round(h * PT))
            for x, y, w, h in specs]


# -- scales ----------------------------------------------------------------------------------------


def test_a_linear_scale_maps_and_reverses():
    score = LinearScale(0, 10, 500, 100)          # a y axis: 0 at the bottom
    assert score.position(0) == 500 and score.position(10) == 100
    assert score.position("7") == pytest.approx(220)
    assert score.position(12) == pytest.approx(20)          # extrapolates
    assert [t.value for t in score.ticks(step=2.5)] == ["0", "2.5", "5", "7.5", "10"]
    assert [t.value for t in LinearScale(0, 100, 0, 1).ticks(count=5)] == \
        ["0", "20", "40", "60", "80", "100"]
    assert nice_step(37, 5) == 10
    with pytest.raises(ScaleError):
        LinearScale(1, 1, 0, 10)
    with pytest.raises(ScaleError):
        score.position("seven")


def test_a_date_scale_has_whole_days_inclusive():
    time = DateScale("2026-11-01", "2027-03-31", 0, 151)    # one point a day
    assert time.days == 151
    assert time.position("2026-11-01") == 0 and time.position("2027-03-31", "end") == 151
    assert time.position("2026-11-02", "center") == pytest.approx(1.5)
    months = time.ticks(every="month")
    assert [t.label for t in months] == ["Nov 2026", "Dec 2026", "Jan 2027", "Feb 2027",
                                         "Mar 2027"]
    assert [(round(t.start, 6), round(t.end, 6)) for t in months][:2] == [(0, 30), (30, 61)]
    weeks = time.ticks(every="week")
    assert weeks[0].value == "2026-10-26" and weeks[0].start == 0     # clipped to the scale
    assert weeks[1].label == "2 Nov"
    assert [t.label for t in time.ticks(every="quarter")] == ["Q4 2026", "Q1 2027"]
    assert len(time.ticks(every="month", step=2)) == 3


def test_excluded_dates_close_up():
    time = DateScale("2026-12-14", "2027-01-08", 0, 14,
                     exclude=[("2026-12-21", "2027-01-01")])
    assert time.days == 14
    assert time.position("2026-12-18", "end") == pytest.approx(5)
    assert time.position("2026-12-21") == time.position("2027-01-01", "end") == \
        time.position("2027-01-02") == pytest.approx(7)
    assert time.position("2026-12-25", "center") == pytest.approx(7)   # inside: the join
    again = scale_from(time.to_json())
    assert again.position("2027-01-05", "center") == time.position("2027-01-05", "center")
    with pytest.raises(ScaleError):
        DateScale("2026-12-01", "2026-11-01", 0, 1)


def test_a_band_scale_has_starts_centres_and_ends():
    rows = BandScale(["Spend", "Category", "Change"], 100, 400, gap=10, padding=5)
    assert rows.width == pytest.approx((300 - 10 - 20) / 3)
    assert rows.position("Spend") == 105 and rows.position("Change", "end") == 395
    assert rows.position("category", "center") == pytest.approx(250)
    up = BandScale(["low", "high"], 400, 100)
    assert up.position("low") == 400 and up.position("low", "end") == 250
    with pytest.raises(ScaleError, match="not a band"):
        rows.position("Nothing")
    with pytest.raises(ScaleError):
        BandScale(["a", "a"], 0, 1)


# -- align and distribute --------------------------------------------------------------------------


def test_align_to_the_selection_the_slide_and_the_first(slide):
    a, b, c = boxes(slide, [(100, 200, 50, 20), (180, 240, 80, 30), (300, 180, 40, 60)])
    align([a, b, c], "left")
    assert {pts(s.drawn_bounds)[0] for s in (a, b, c)} == {100.0}
    align([a, b, c], "middle")
    assert {round(pts(s.drawn_bounds)[1] + pts(s.drawn_bounds)[3] / 2, 2) for s in (a, b, c)} \
        == {225.0}
    align([b, a, c], "center", to="first")
    assert pts(b.drawn_bounds)[0] == 100 and pts(a.drawn_bounds)[0] == 115
    align([c], "right", to="slide")
    assert pts(c.drawn_bounds)[0] == 920
    with pytest.raises(ValueError):
        align([a], "left")


def test_align_uses_what_is_drawn_and_moves_a_group_whole(slide):
    turned, plain = boxes(slide, [(200, 200, 100, 20), (400, 100, 60, 60)])
    turned.rotation = 90                       # drawn 20 wide, 100 tall
    member, other = boxes(slide, [(500, 300, 40, 40), (560, 300, 40, 40)])
    group = slide.group([member, other])
    align([turned, plain, group], "top")
    assert pts(turned.drawn_bounds)[1] == pts(plain.drawn_bounds)[1] == 100
    assert pts(group.drawn_bounds)[1] == 100
    assert pts(other.drawn_bounds)[0] - pts(member.drawn_bounds)[0] == 60   # moved together
    align([turned, plain], "left")
    assert pts(turned.drawn_bounds)[0] == 240                 # its drawn left edge


def test_align_is_one_undo_step(slide):
    deck = slide.document
    a, b = boxes(slide, [(100, 100, 50, 50), (200, 150, 50, 50)])
    before = deck.to_bytes()
    align([a, b], "bottom")
    deck.undo()
    assert deck.to_bytes() == before


def test_distribute_evenly_by_position_or_with_a_gap(slide):
    a, b, c, d = boxes(slide, [(100, 100, 40, 20), (400, 100, 60, 20), (170, 100, 20, 20),
                               (600, 100, 40, 20)])
    distribute([a, b, c, d], "horizontal")
    order = sorted((a, b, c, d), key=lambda s: s.drawn_bounds[0])
    assert [s.id for s in order] == [a.id, c.id, b.id, d.id]
    edges = [pts(s.drawn_bounds) for s in order]
    gaps = [round(edges[i + 1][0] - edges[i][0] - edges[i][2], 2) for i in range(3)]
    assert edges[0][0] == 100 and edges[-1][0] + edges[-1][2] == 640
    assert max(gaps) - min(gaps) <= 0.02
    distribute([a, b, c, d], "horizontal", gap=10 * PT)
    assert pts(c.drawn_bounds)[0] == 150 and pts(b.drawn_bounds)[0] == 180
    distribute([a, b], "vertical", to="slide")
    assert pts(a.drawn_bounds)[1] == 0 and pts(b.drawn_bounds)[1] == 520
    with pytest.raises(ValueError, match="cannot be spread"):
        distribute([a, b, c, d], "horizontal", to="box", box=(0, 0, 100 * PT, 10 * PT))


# -- stack, column, grid ------------------------------------------------------------------------------


def test_stack_packs_aligns_and_spreads(slide):
    shapes = boxes(slide, [(0, 0, 100, 40), (0, 0, 60, 80), (0, 0, 80, 20)])
    box = (50 * PT, 150 * PT, 800 * PT, 100 * PT)
    stack(shapes, box=box, gap=10 * PT, align="middle")
    assert [pts(s.drawn_bounds) for s in shapes] == [(50, 180, 100, 40), (160, 160, 60, 80),
                                                    (230, 190, 80, 20)]
    stack(shapes, box=box, gap=10 * PT, justify="end", align="bottom")
    assert pts(shapes[-1].drawn_bounds)[0] + 80 == 850 and pts(shapes[0].drawn_bounds)[1] == 210
    stack(shapes, box=box, justify="spread")
    assert pts(shapes[0].drawn_bounds)[0] == 50 and pts(shapes[2].drawn_bounds)[0] == 770
    stack(shapes, box=box, gap=12 * PT, equal=True)
    widths = {pts(s.drawn_bounds)[2] for s in shapes}
    assert widths == {round((800 - 24) / 3, 2)}
    with pytest.raises(LayoutError) as caught:
        stack(shapes, box=(0, 0, 100 * PT, 50 * PT), gap=10 * PT)
    assert caught.value.needed > caught.value.available


def test_column_fits_text_heights_then_stacks(slide):
    first = slide.add_textbox(0, 0, 200 * PT, 10 * PT, "A heading", autofit="none")
    second = slide.add_shape("rect", 0, 0, 200 * PT, 10 * PT,
                             text="Body text that is long enough to wrap onto more than one "
                                  "line in a box two hundred points wide")
    column([first, second], box=(100 * PT, 150 * PT, 300 * PT, 300 * PT), gap=6 * PT,
           align="center", fit_text=True)
    a, b = pts(first.drawn_bounds), pts(second.drawn_bounds)
    assert a[0] == b[0] == 150 and a[1] == 150
    assert b[1] == round(a[1] + a[3] + 6, 2)
    assert not first.text_fit().overflows and not second.text_fit().overflows
    assert b[3] > 30


def test_grid_places_in_cells_rows_or_columns_first(slide):
    shapes = boxes(slide, [(0, 0, 40, 20)] * 5)
    box = (100 * PT, 100 * PT, 300 * PT, 120 * PT)
    grid(shapes, box=box, columns=3, gutter=(30 * PT, 20 * PT))
    cell_w, cell_h = (300 - 60) / 3, (120 - 20) / 2
    first, fourth = pts(shapes[0].drawn_bounds), pts(shapes[3].drawn_bounds)
    assert first == (round(100 + (cell_w - 40) / 2, 2), round(100 + (cell_h - 20) / 2, 2), 40, 20)
    assert fourth[1] == round(100 + cell_h + 20 + (cell_h - 20) / 2, 2)
    grid(shapes, box=box, rows=2, order="columns", fit="cell")
    assert pts(shapes[1].drawn_bounds) == (100, 160, 100, 60)
    with pytest.raises(LayoutError):
        grid(shapes, box=box, rows=2, columns=2)


def test_place_moves_and_resizes_by_the_drawn_box(slide):
    shape, = boxes(slide, [(10, 10, 50, 20)])
    shape.rotation = 90
    place(shape, 100 * PT, 120 * PT, 30 * PT, 80 * PT)
    assert pts(shape.drawn_bounds) == (100, 120, 30, 80)
    assert pts(drawn_box(shape)) == (100, 120, 30, 80)


# -- labels ---------------------------------------------------------------------------------------------


def _scatter(slide):
    score = LinearScale(0, 10, 100 * PT, 700 * PT)
    up = LinearScale(0, 10, 480 * PT, 160 * PT)
    points = [(7.5, 8.5), (3.5, 9.0), (8.0, 6.5), (2.5, 7.0), (8.5, 3.5), (6.0, 2.0),
              (9.0, 5.5), (6.5, 7.5), (2.0, 3.0), (4.0, 1.5)]
    bubbles = []
    for x, y in points:
        cx, cy = score.position(x), up.position(y)
        bubbles.append(slide.add_shape("ellipse", round(cx - 9 * PT), round(cy - 9 * PT),
                                       18 * PT, 18 * PT))
    return bubbles


def test_labels_go_beside_their_anchors_without_collisions(slide):
    bubbles = _scatter(slide)
    names = ["Cement & binders", "Steel reinforcement", "Packaging", "Energy contracts",
             "MRO spares", "IT & telecoms", "Professional services", "Inbound freight",
             "Warehousing", "Fleet leasing"]
    labels = [slide.add_textbox(0, 0, 90 * PT, 16 * PT, name, autofit="none")
              for name in names]
    for label in labels:
        label.text_frame.paragraphs[0].runs[0].size = 10
    report = place_labels(slide, [{"shape": l, "anchor": b} for l, b in zip(labels, bubbles)],
                          sides=("right", "left", "above", "below"), distance=4 * PT,
                          max_center=72 * PT)
    assert not report.unplaced and not report.collisions
    for placed in report.placed:
        assert placed.center_distance <= 72 * PT + 1
    assert slide.collisions(boxes=True) == []


def test_a_blocked_label_is_reported_with_what_blocks_it(slide):
    anchor, wall = boxes(slide, [(400, 300, 20, 20), (300, 200, 400, 300)])
    wall.send_to_back()
    wall.set_text("A text panel")       # text: labels may not lie on it
    label = slide.add_textbox(0, 0, 60 * PT, 16 * PT, "Blocked", autofit="none")
    report = place_labels(slide, [{"shape": label, "anchor": anchor}], max_edge=20 * PT)
    assert not report.placed and report.unplaced[0].label == label.id
    assert wall.id in report.unplaced[0].blockers


def test_labels_at_points_with_leaders(slide):
    a, b = (slide.add_textbox(0, 0, 50 * PT, 14 * PT, text, autofit="none")
            for text in ("Kick-off", "SteerCo"))
    report = place_labels(slide, [{"shape": a, "point": (200 * PT, 300 * PT)},
                                  {"shape": b, "point": (220 * PT, 300 * PT)}],
                          sides=("above",), distance=6 * PT, leader="line")
    assert len(report.placed) == 2 and all(p.leader for p in report.placed)
    boxes_ = [p.box for p in report.placed]
    assert not (boxes_[0][0] + boxes_[0][2] > boxes_[1][0] and boxes_[0][1] == boxes_[1][1])
    leader = slide.document.shape(report.placed[0].leader)
    assert leader.is_connector and leader.end_connection[0].id == a.id
