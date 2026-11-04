"""Copies following a revised sample slide (round 31, R8): ``Document.find_copies``,
``Document.follow_sample`` and ``ppt_follow_sample``.  Every deck here is made up: a sample
with a title, a labelled band, a marker and a table, copies duplicated from it and filled
with their own text, then the sample revised."""

from __future__ import annotations

import json

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.errors import ArgumentError
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

E = 12700
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _named(slide, name):
    return next(shape for shape in slide.shapes if shape.name == name)


def _body(slide):
    return next(shape for shape in slide.shapes
                if shape.placeholder is not None and shape.placeholder.idx == 1)


def _sample(deck: Document):
    slide = deck.add_slide("Title and Content")
    slide.title = "Sample heading"
    _body(slide).set_text("First point\nSecond point")
    band = slide.add_shape("rect", 40 * E, 380 * E, 300 * E, 40 * E, text="Band label")
    band.name = "Band"
    band.fill = "accent1"
    marker = slide.add_shape("ellipse", 500 * E, 380 * E, 40 * E, 40 * E)
    marker.name = "Marker"
    table = slide.add_table(3, 2, 380 * E, 440 * E, 300 * E, 60 * E, name="Figures")
    for row in range(3):
        for column in range(2):
            table.table.cell(row, column).set_text(f"r{row}c{column}")
    return slide


def _run(count: int = 3):
    """A deck: the sample, ``count`` copies with their own text, and an unrelated slide on
    the same layout."""
    deck = Document.new()
    sample = _sample(deck)
    copies = []
    for number in range(count):
        copy = deck.duplicate_slide(sample, index=len(deck.slides))
        copy.title = f"Copy {number} heading"
        _named(copy, "Band").set_text(f"Band of copy {number}")
        _body(copy).set_text("\n".join(f"Own point {n}" for n in range(number + 1)))
        figures = _named(copy, "Figures").table
        figures.cell(1, 1).set_text(f"{100 + number}")
        copies.append(copy.slide_id)
    other = deck.add_slide("Title and Content")
    other.title = "Unrelated"
    return deck, sample.slide_id, copies, other.slide_id


def _revise(deck, sample_id):
    sample = deck.slide(sample_id)
    band = _named(sample, "Band")
    band.move_by(dy=-100 * E)
    band.fill = "#C00000"
    band.text_frame.paragraphs[0].runs[0].bold = True
    _named(sample, "Marker").delete()
    mark = sample.add_shape("rect", 600 * E, 40 * E, 40 * E, 40 * E)
    mark.name = "New mark"
    figures = _named(sample, "Figures").table
    figures.set_column_width(0, 200 * E)
    figures.set_column_width(1, 100 * E)
    figures.cell(0, 0).fill = "#FFC000"
    return sample


# -- telling copies -----------------------------------------------------------------------


def test_copies_are_the_slides_on_its_layout_that_kept_its_shapes_by_id_and_name():
    deck, sample, copies, other = _run()
    found = deck.find_copies(f"s:{sample}")
    assert sorted(found.slides) == sorted(f"s:{sid}" for sid in copies)
    assert all(entry["shapes"] == "3 of 3" for entry in found.copies)
    assert f"s:{other}" not in json.dumps(found.to_json())


def test_a_revised_sample_still_finds_its_copies_and_a_slide_on_another_layout_is_near():
    deck, sample, copies, _ = _run()
    _revise(deck, sample)
    moved = deck.duplicate_slide(deck.slide(sample), index=len(deck.slides))
    deck.set_layout(moved, "Title Only") if hasattr(deck, "set_layout") else moved.set_layout(
        "Title Only")
    found = deck.find_copies(sample)
    assert sorted(found.slides) == sorted(f"s:{sid}" for sid in copies)
    assert found.near == [{"slide": f"s:{moved.slide_id}", "shapes": "3 of 3",
                           "why": "on another layout"}]


def test_a_sample_of_placeholders_only_asks_for_its_copies():
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    slide.title = "Only a title"
    deck.duplicate_slide(slide)
    with pytest.raises(ArgumentError, match="name the copies"):
        deck.find_copies(slide)


# -- following ----------------------------------------------------------------------------


def test_copies_follow_the_design_and_keep_their_own_text():
    deck, sample, copies, other = _run()
    _revise(deck, sample)
    result = deck.follow_sample(sample)
    assert [c.slide for c in result.copies] == [f"s:{sid}" for sid in copies]
    expected = _named(deck.slide(sample), "Band")
    for number, sid in enumerate(copies):
        copy = deck.slide(sid)
        band = _named(copy, "Band")
        assert band.bounds == expected.bounds
        assert band.text == f"Band of copy {number}"
        assert band.text_frame.paragraphs[0].runs[0].bold is True
        assert band._element.find(".//{%s}srgbClr" % A).get("val") == "C00000"
        assert copy.title == f"Copy {number} heading"
        assert _body(copy).text == "\n".join(f"Own point {n}" for n in range(number + 1))
        names = [shape.name for shape in copy.shapes]
        assert "Marker" not in names and "New mark" in names
        figures = _named(copy, "Figures").table
        assert figures.column_widths == [200 * E, 100 * E]
        assert figures.cell(1, 1).text == f"{100 + number}"
        assert figures.cell(0, 0).fill is not None
        report = result.copies[number].to_json()
        band_entry = next(e for e in report["shapes"] if e["name"] == "Band")
        assert {"box", "fill", "text format"} <= set(band_entry["what"])
        assert report["removed"][0]["name"] == "Marker"
        assert report["added"][0]["name"] == "New mark"
    assert deck.slide(other).title == "Unrelated"
    assert deck.validate() == []


def test_following_again_changes_nothing_and_says_so():
    deck, sample, copies, _ = _run(2)
    _revise(deck, sample)
    deck.follow_sample(sample)
    version = deck.history.version
    again = deck.follow_sample(sample).to_json()
    assert again["followed"] == 0 and again["unchanged"] == 2
    assert all(entry == {"slide": entry["slide"], "changed": False}
               for entry in again["copies"])
    assert deck.history.version == version


def test_one_undo_step_and_a_preview_that_changes_nothing():
    deck, sample, copies, _ = _run(2)
    _revise(deck, sample)
    before = deck.to_bytes()
    preview = deck.follow_sample(sample, preview=True)
    assert preview.preview and all(c.is_changed for c in preview.copies)
    assert deck.to_bytes() == before
    done = deck.follow_sample(sample)
    def without_new_ids(data):          # a preview's new shape ids are not handed out again
        for entry in data["copies"]:
            for added in entry.get("added", ()):
                added.pop("shape")
        return data["copies"]

    assert without_new_ids(done.to_json()) == without_new_ids(preview.to_json())
    assert deck.to_bytes() != before
    deck.undo()
    assert deck.to_bytes() == before


def test_a_shape_only_on_one_copy_is_its_own_and_stays_unless_remove_all():
    deck, sample, copies, _ = _run(3)
    note = deck.slide(copies[1]).add_textbox(40 * E, 300 * E, 200 * E, 30 * E, "Own note")
    note.name = "Own note"
    result = deck.follow_sample(sample).to_json()
    kept = result["copies"][1]["not_followed"]
    assert kept == [{"shape": note.id, "name": "Own note",
                     "why": "the sample has none: the copy's own, kept (remove: \"all\" "
                            "deletes it)"}]
    assert "Own note" in [s.name for s in deck.slide(copies[1]).shapes]
    deck.follow_sample(sample, remove="all")
    assert "Own note" not in [s.name for s in deck.slide(copies[1]).shapes]


def test_a_shape_taken_off_one_copy_is_not_put_back():
    deck, sample, copies, _ = _run(3)
    _named(deck.slide(copies[0]), "Marker").delete()
    result = deck.follow_sample(sample).to_json()
    entry = result["copies"][0]
    assert "Marker" not in [s.name for s in deck.slide(copies[0]).shapes]
    assert entry["not_followed"][0]["name"] == "Marker"
    assert "taken off here" in entry["not_followed"][0]["why"]


def test_text_with_other_paragraphs_follows_by_level_and_keeps_a_bold_word():
    deck, sample, copies, _ = _run(1)
    copy = deck.slide(copies[0])
    body = _body(copy)
    body.set_text("Alpha\nBeta\nGamma")
    body.text_frame.paragraphs[1].runs[0].bold = True
    for paragraph in _body(deck.slide(sample)).text_frame.paragraphs:
        for run in paragraph.runs:
            run.size = 14
            run.italic = True
    result = deck.follow_sample(sample).to_json()
    runs = [p.runs[0] for p in _body(deck.slide(copies[0])).text_frame.paragraphs]
    assert [r.size for r in runs] == [14, 14, 14]
    assert all(r.italic for r in runs)
    assert [r.bold for r in runs] == [None, True, None]
    entry = next(e for e in result["copies"][0]["shapes"] if e["name"] ==
                 "Content Placeholder 2")
    assert "formatting followed by level" in entry["note"]


def test_a_table_with_more_rows_takes_the_header_and_body_rows_and_its_own_height():
    deck, sample, copies, _ = _run(1)
    figures = _named(deck.slide(copies[0]), "Figures")
    figures.table.insert_row()
    figures.table.insert_row()
    sample_table = _named(deck.slide(sample), "Figures").table
    sample_table.set_row_height(0, 30 * E)
    sample_table.set_row_height(1, 18 * E)
    sample_table.set_row_height(2, 18 * E)
    result = deck.follow_sample(sample).to_json()
    heights = _named(deck.slide(copies[0]), "Figures").table.row_heights
    assert heights == [30 * E] + [18 * E] * 4
    entry = next(e for e in result["copies"][0]["shapes"] if e["name"] == "Figures")
    assert "row heights" in entry["what"] and "rows past the sample's" in entry["note"]


def test_a_table_with_other_columns_is_reported_and_keeps_its_widths():
    deck, sample, copies, _ = _run(1)
    figures = _named(deck.slide(copies[0]), "Figures").table
    figures.insert_column()
    widths = figures.column_widths
    _revise(deck, sample)
    result = deck.follow_sample(sample).to_json()
    assert _named(deck.slide(copies[0]), "Figures").table.column_widths == widths
    why = next(e["why"] for e in result["copies"][0]["not_followed"] if e["name"] == "Figures")
    assert why.startswith("the table has 3 column(s), the sample's 2")


def test_a_sample_in_another_deck_is_followed_by_the_imported_copies():
    samples = Document.new()
    sample = _sample(samples)
    report = Document.new()
    copies = [report.import_slide(sample).slide for _ in range(2)]
    for number, copy in enumerate(copies):
        _named(copy, "Band").set_text(f"Imported {number}")
    report.add_slide("Title and Content").title = "Other"
    _revise(samples, sample.slide_id)
    assert sorted(report.find_copies(sample).slides) == sorted(
        f"s:{c.slide_id}" for c in copies)
    result = report.follow_sample(sample)
    assert result.sample.endswith("of the sample deck")
    for number, copy in enumerate(copies):
        band = _named(report.slide(copy.slide_id), "Band")
        assert band.text == f"Imported {number}" and band.top == 280 * E
    assert report.validate() == []


def test_new_overflows_are_reported_per_copy():
    pytest.importorskip("pptx2svg")
    deck, sample, copies, _ = _run(1)
    band = _named(deck.slide(sample), "Band")
    band.height = 12 * E
    result = deck.follow_sample(sample).to_json()
    overflow = result["copies"][0]["overflows"]
    assert any(e["kind"] == "text" and e["why"] == "its text no longer fits after following"
               for e in overflow)


def test_unknown_remove_and_no_copies_are_refused_with_the_way_out():
    deck, sample, copies, _ = _run(1)
    with pytest.raises(ArgumentError, match="remove is one of auto, all, none"):
        deck.follow_sample(sample, remove="some")
    with pytest.raises(ArgumentError, match="name slides other than the sample"):
        deck.follow_sample(sample, [sample])


# -- the tool -----------------------------------------------------------------------------


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as box:
        yield box


def _session(toolbox, count=3):
    deck, sample, copies, other = _run(count)
    _revise(deck, sample)
    session = toolbox.session()
    doc = session.open(deck.to_bytes(), "run.pptx")
    return session, doc, sample, copies, other


def _call(toolbox, session, **arguments):
    return toolbox.dispatch(session, "ppt_follow_sample", arguments)


def test_the_tool_previews_then_follows_and_accepts_the_forms_other_tools_take(toolbox):
    session, doc, sample, copies, _ = _session(toolbox)
    result = _call(toolbox, session, doc=doc, sample=f"{sample}.3", copies=[], preview=True,
                   remove="")
    assert result.ok, result.to_json()
    data = result.to_json()
    assert data["data"]["preview"] is True and data["data"]["followed"] == 3
    assert any("is a shape" in w for w in data["warnings"])
    assert any("copies: [] read as every slide" in w for w in data["warnings"])
    assert data["changed"] == []
    result = _call(toolbox, session, doc=doc, sample=1, copies=[2, f"s:{copies[1]}",
                                                                 f"s:{sample}"])
    assert result.ok, result.to_json()
    assert [c["slide"] for c in result.data["copies"]] == [f"s:{copies[0]}",
                                                           f"s:{copies[1]}"]
    assert any("is the sample; skipped" in w for w in result.warnings)
    assert sorted(result.changed) == sorted([f"s:{copies[0]}", f"s:{copies[1]}"])


def test_the_tool_in_a_batch_reports_an_already_following_copy_as_unchanged(toolbox):
    session, doc, sample, copies, _ = _session(toolbox, 2)
    _call(toolbox, session, doc=doc, sample=f"s:{sample}")
    result = toolbox.dispatch(session, "batch", {"ops": [
        {"tool": "ppt_follow_sample", "arguments": {"doc": doc, "sample": f"s:{sample}"}},
        {"tool": "ppt_set_text", "arguments": {"doc": doc, "target": f"{copies[0]}.2",
                                               "text": "Still editable"}}]})
    assert result.ok, result.to_json()
    first = result.data["ops"][0]["data"]
    assert first["unchanged"] == 2 and first["followed"] == 0


def test_the_tool_refusal_names_the_copies_it_could_follow(toolbox):
    session, doc, sample, copies, other = _session(toolbox, 2)
    result = _call(toolbox, session, doc=doc, sample=f"s:{sample}", copies=[f"s:{sample}"])
    assert not result.ok
    assert result.error.field == "copies"
    assert set(result.error.valid_options) >= {f"s:{sid}" for sid in copies}


def test_the_tool_keeps_to_its_budget_and_lists_what_it_left_out(toolbox, monkeypatch):
    from pptx_agent.tools import follow

    monkeypatch.setattr(follow, "BUDGET", 700)
    session, doc, sample, copies, _ = _session(toolbox, 6)
    result = _call(toolbox, session, doc=doc, sample=f"s:{sample}", preview=True)
    assert result.ok
    data = result.data
    assert all(isinstance(c.get("shapes", 0), int) for c in data["copies"])
    shown = [c["slide"] for c in data["copies"]]
    assert sorted(shown + data.get("left_out", [])) == sorted(f"s:{sid}" for sid in copies)
    assert data.get("left_out")
    assert any("left_out" in w for w in result.warnings)


def test_a_picture_follows_the_box_and_keeps_the_image_its_copy_filled():
    from images import png

    deck = Document.new()
    sample = deck.add_slide("Title Only")
    sample.title = "With a picture"
    photo = sample.add_picture(png(4, 4, lambda x, y: (200, 0, 0)), 40 * E, 120 * E,
                               200 * E, 200 * E)
    photo.name = "Photo"
    copies = [deck.duplicate_slide(sample) for _ in range(2)]
    _named(copies[0], "Photo").replace_image(png(4, 4, lambda x, y: (0, 0, 200)))
    own = _named(deck.slide(copies[0].slide_id), "Photo").image_part
    _named(deck.slide(sample.slide_id), "Photo").move_by(dx=50 * E)
    deck.follow_sample(sample)
    first = _named(deck.slide(copies[0].slide_id), "Photo")
    assert first.left == 90 * E and first.image_part == own
    assert _named(deck.slide(copies[1].slide_id), "Photo").left == 90 * E


def test_default_names_tell_a_copy_only_in_the_samples_place():
    deck = Document.new()
    sample = deck.add_slide("Title Only")
    box = sample.add_shape("rect", 40 * E, 120 * E, 200 * E, 60 * E, text="Box")
    assert box.name == "Rectangle 2"                 # a default name: any slide gets it
    copy = deck.duplicate_slide(sample)
    built = deck.add_slide("Title Only")             # built the same way, elsewhere
    built.add_shape("rect", 300 * E, 300 * E, 100 * E, 60 * E, text="Other")
    found = deck.find_copies(sample)
    assert found.slides == [f"s:{copy.slide_id}"]
    assert found.near[0]["slide"] == f"s:{built.slide_id}"
    assert "none in its places" in found.near[0]["why"]
