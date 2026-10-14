"""Source addresses kept in the deck: a table's or a chart's ``from`` is written into its
graphic frame's tags part (``p:custDataLst/p:tags`` -> ``ppt/tags/tagN.xml``, a
``UVRT_PROVENANCE`` tag) and read back after a save, through ``Session.provenance`` and
``Document.provenance()``; table edits keep it true or drop it; copies get their own part
and deletes leave none behind.

The data is made up: the toy dataset and grid of :mod:`test_tables_from_data`.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile

import pytest

from ooxml_edit import provenance as records
from ooxml_edit.tools import Toolbox, strict_problems

from pptx_agent import Document
from pptx_agent.edit import provenance as kept
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

from test_tables_from_data import GRID, _session, add, call, ds_put, failed, ok

TAGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags"
CT_TAGS = "application/vnd.openxmlformats-officedocument.presentationml.tags+xml"
CHART_BOX = {"x": 40, "y": 300, "w": 300, "h": 200}


@pytest.fixture(scope="module")
def toolbox():
    with Toolbox([*TOOLS, ds_put], formats=[FORMAT, GRID], groups=GROUPS, workers=0) as box:
        yield box


@pytest.fixture
def session(toolbox):
    session = _session(toolbox)
    ok(call(toolbox, session, "ds_put", id="scores"))
    return session


def deck_of(session):
    return session.entry("d1").document


def table_from(toolbox, session, **arguments):
    return ok(add(toolbox, session, **{"from": "d:scores",
                                       "from_columns": ["check", "rating", "due"],
                                       **arguments})).created[0]


def chart_from(toolbox, session):
    return ok(call(toolbox, session, "edit_chart", doc="d1", action="add", target="s:256",
                   box=CHART_BOX, chart_type="bar",
                   **{"from": "t:checks", "columns": ["check", "rating"]})).created[0]


def edit(toolbox, session, target, **arguments):
    return ok(call(toolbox, session, "ppt_edit_table", doc="d1", target=target, **arguments))


def tag_parts(package) -> list[str]:
    return sorted(name for name in package.part_names if name.startswith("ppt/tags/"))


def reopened(session) -> Document:
    return Document.open(deck_of(session).to_bytes())


# -- what is written -----------------------------------------------------------------------------


def test_the_definitions_are_unchanged_and_strict_clean():
    assert strict_problems(TOOLS) == []


def test_a_table_from_a_source_keeps_its_record_in_a_tags_part(toolbox, session):
    result = ok(add(toolbox, session, **{"from": "d:scores",
                                         "from_columns": ["check", "rating"]}))
    target = result.created[0]
    data = deck_of(session).to_bytes()
    package = zipfile.ZipFile(io.BytesIO(data))
    slide = package.read("ppt/slides/slide1.xml").decode()
    rels = package.read("ppt/slides/_rels/slide1.xml.rels").decode()
    types = package.read("[Content_Types].xml").decode()
    # p:nvGraphicFramePr/p:nvPr/p:custDataLst/p:tags r:id -> a slide relationship of type tags
    rel_id = re.search(r'<p:nvPr><p:custDataLst><p:tags r:id="(rId\d+)"/></p:custDataLst>'
                       r'</p:nvPr>', slide).group(1)
    assert f'Id="{rel_id}" Type="{TAGS}" Target="../tags/tag1.xml"' in rels
    assert f'<Override PartName="/ppt/tags/tag1.xml" ContentType="{CT_TAGS}"/>' in types
    tag = package.read("ppt/tags/tag1.xml").decode()
    assert tag.startswith("<?xml")
    assert re.search(r'<p:tagLst [^>]*><p:tag name="UVRT_PROVENANCE" val="[^"]+"/></p:tagLst>$',
                     tag)
    found = Document.open(data).provenance()
    assert list(found) == [target]
    expected = records.record(result.data["from"], document=None)
    assert found[target] == expected
    assert found[target]["first_row"] == 1 and found[target]["header_row"] == 0
    assert found[target]["columns"][1]["sources"] == ["G!B2", "G!B3", "G!B4"]
    assert session.provenance("d1") == found
    assert Document.open(data).validate() == []


def test_a_record_reads_back_through_a_new_session(toolbox, session):
    target = table_from(toolbox, session)
    data = deck_of(session).to_bytes()
    other = toolbox.session()
    other.open(data, "again.pptx")
    assert list(other.provenance("d1")) == [target]


def test_a_chart_from_a_source_keeps_the_same_record_on_its_frame(toolbox, session):
    target = chart_from(toolbox, session)
    deck = reopened(session)
    found = deck.provenance()[target]
    assert found["categories"]["column"] == "check"
    assert [s["name"] for s in found["series"]] == ["rating"]
    # The chart part's own copy says the same.
    from ooxml_edit.charts import provenance as chart_record

    assert list(chart_record(deck.package).values()) == [found]
    assert tag_parts(deck.package) == ["ppt/tags/tag1.xml"]
    assert deck.validate() == []


def test_the_deck_still_renders(toolbox, session):
    pptx2svg = pytest.importorskip("pptx2svg")
    table_from(toolbox, session)
    chart_from(toolbox, session)
    data = deck_of(session).to_bytes()
    assert len(pptx2svg.convert_pptx_to_svg(data)) == 1


def test_undo_takes_the_record_with_the_table(toolbox, session):
    before = deck_of(session).to_bytes()
    target = table_from(toolbox, session)
    edit(toolbox, session, target, action="set_cells", cells=[{"row": 1, "col": 0,
                                                               "text": "x"}])
    ok(call(toolbox, session, "undo", doc="d1"))
    assert [c["col"] for c in session.provenance("d1")[target]["columns"]] == [0, 1, 2]
    ok(call(toolbox, session, "undo", doc="d1"))
    assert session.provenance("d1") == {}
    assert tag_parts(deck_of(session).package) == []
    assert deck_of(session).to_bytes() == before


def test_set_cells_from_a_source_replaces_the_record(toolbox, session):
    target = table_from(toolbox, session)
    edit(toolbox, session, target, action="set_cells", header=[],
         **{"from": "t:checks", "from_columns": ["note"]})
    found = session.provenance("d1")[target]
    assert found["source"] == "t:checks" and found["document"] == "book.grid"
    assert "header_row" not in found and found["first_row"] == 0
    assert [c["column"] for c in found["columns"]] == ["note"]
    assert tag_parts(deck_of(session).package) == ["ppt/tags/tag1.xml"]


# -- the drop rules ------------------------------------------------------------------------------


def columns_of(session, target):
    found = session.provenance("d1").get(target)
    return None if found is None else [(c["col"], c["column"]) for c in found["columns"]]


def test_a_cell_written_without_a_source_drops_its_column(toolbox, session):
    target = table_from(toolbox, session)                    # 4 rows: header + 3
    edit(toolbox, session, target, action="set_cells",
         cells=[{"row": 2, "col": 1, "text": "9"}])
    assert columns_of(session, target) == [(0, "check"), (2, "due")]
    edit(toolbox, session, target, action="set_cells",
         cells=[{"row": 0, "col": 0, "text": "Name"}])       # the header is the block's
    assert columns_of(session, target) == [(2, "due")]
    edit(toolbox, session, target, action="set_cells",
         cells=[{"row": 3, "col": 2, "text": "soon"}])
    assert columns_of(session, target) is None
    deck = deck_of(session)
    assert tag_parts(deck.package) == []
    assert not any(r.type == TAGS for r in
                   deck.package.relationships(deck.slides[0].part_path).values())
    assert b"custDataLst" not in deck.package.read(deck.slides[0].part_path)
    assert deck.validate() == []


def test_a_cell_outside_the_block_keeps_the_record(toolbox, session):
    target = table_from(toolbox, session, rows=6, columns=4)
    edit(toolbox, session, target, action="set_cells",
         cells=[{"row": 5, "col": 0, "text": "total"}, {"row": 1, "col": 3, "text": "x"}])
    assert columns_of(session, target) == [(0, "check"), (1, "rating"), (2, "due")]


def test_set_text_and_replace_text_on_a_cell_drop_its_column(toolbox, session):
    target = table_from(toolbox, session)
    ok(call(toolbox, session, "ppt_set_text", doc="d1",
            items=[{"target": f"{target}/cell1,0", "text": "Delta"}]))
    assert columns_of(session, target) == [(1, "rating"), (2, "due")]
    ok(call(toolbox, session, "replace_text", doc="d1", find="2026-02-10",
            replace="2026-02-11", expect="one"))
    assert columns_of(session, target) == [(1, "rating")]


def test_a_merge_over_the_block_drops_its_columns(toolbox, session):
    target = table_from(toolbox, session)
    edit(toolbox, session, target, action="merge", span={"row": 1, "col": 1, "to_col": 2})
    assert columns_of(session, target) == [(0, "check")]


@pytest.mark.parametrize("action,index,expected", [
    ("insert_row", 0, (1, 2)),          # above the header: the block moves down
    ("insert_row", None, (0, 1)),       # appended below: nothing changes
    ("insert_row", 2, None),            # inside: dropped
    ("delete_row", 3, None),            # the last source row: dropped
    ("delete_row", 0, None),            # the header: dropped
])
def test_row_inserts_and_deletes_move_or_drop_the_record(toolbox, session, action, index,
                                                         expected):
    target = table_from(toolbox, session)
    arguments = {} if index is None else {"index": index}
    edit(toolbox, session, target, action=action, **arguments)
    found = session.provenance("d1").get(target)
    if expected is None:
        assert found is None and tag_parts(deck_of(session).package) == []
    else:
        assert (found["header_row"], found["first_row"]) == expected


def test_a_row_deleted_above_the_block_moves_it_up(toolbox, session):
    target = table_from(toolbox, session, rows=5)
    edit(toolbox, session, target, action="insert_row", index=0)
    edit(toolbox, session, target, action="set_cells",
         cells=[{"row": 0, "col": 0, "text": "above"}])      # above the block: kept
    edit(toolbox, session, target, action="delete_row", index=0)
    found = session.provenance("d1")[target]
    assert (found["header_row"], found["first_row"]) == (0, 1)
    edit(toolbox, session, target, action="delete_row", index=4)    # below it: kept
    assert session.provenance("d1")[target]["first_row"] == 1


def test_column_inserts_and_deletes_move_the_entries(toolbox, session):
    target = table_from(toolbox, session)
    edit(toolbox, session, target, action="insert_column", index=1)
    assert columns_of(session, target) == [(0, "check"), (2, "rating"), (3, "due")]
    edit(toolbox, session, target, action="insert_column")          # appended: no change
    assert columns_of(session, target) == [(0, "check"), (2, "rating"), (3, "due")]
    edit(toolbox, session, target, action="delete_column", index=2)
    assert columns_of(session, target) == [(0, "check"), (2, "due")]
    edit(toolbox, session, target, action="delete_column", index=1)
    assert columns_of(session, target) == [(0, "check"), (1, "due")]
    edit(toolbox, session, target, action="delete_column", index=0)
    edit(toolbox, session, target, action="delete_column", index=0)
    assert columns_of(session, target) is None


def test_formatting_keeps_the_record(toolbox, session):
    target = table_from(toolbox, session)
    before = session.provenance("d1")
    edit(toolbox, session, target, action="format_cells", fill="accent1")
    edit(toolbox, session, target, action="set_flags", flags={"banded_rows": False})
    assert session.provenance("d1") == before


def test_a_chart_edit_that_drops_the_record_drops_the_tag(toolbox, session):
    target = chart_from(toolbox, session)
    ok(call(toolbox, session, "edit_chart", doc="d1", target=target, action="add_category",
            text="Delta", values=[2]))
    assert session.provenance("d1") == {}
    assert tag_parts(deck_of(session).package) == []


# -- copies and deletes --------------------------------------------------------------------------


def test_a_duplicated_shape_gets_its_own_tags_part(toolbox, session):
    table = table_from(toolbox, session)
    chart = chart_from(toolbox, session)
    copies = ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=[table, chart],
                     action="duplicate", dy=10)).created
    deck = deck_of(session)
    package = deck.package
    assert len(tag_parts(package)) == 4
    found = session.provenance("d1")
    assert found[copies[0]] == found[table] and found[copies[1]] == found[chart]
    slide = deck.slides[0].part_path
    targets = [r.target_part for r in package.relationships(slide).values() if r.type == TAGS]
    assert len(set(targets)) == 4
    # Writing to the copy's record leaves the original's alone.
    edit(toolbox, session, copies[0], action="delete_column", index=0)
    assert [c["column"] for c in session.provenance("d1")[table]["columns"]] == \
        ["check", "rating", "due"]
    assert reopened(session).validate() == []


def test_a_duplicated_slide_gets_its_own_tags_parts(toolbox, session):
    table = table_from(toolbox, session)
    copy = ok(call(toolbox, session, "ppt_manage_slides", doc="d1", action="duplicate",
                   slide="s:256")).created[0]
    deck = deck_of(session)
    assert len(tag_parts(deck.package)) == 2
    found = session.provenance("d1")
    copied = [address for address in found if address.startswith(copy[2:] + ".")]
    assert len(copied) == 1 and found[copied[0]] == found[table]
    edit(toolbox, session, copied[0], action="delete_column", index=0)
    assert len(session.provenance("d1")[table]["columns"]) == 3


def test_copied_shapes_get_their_own_tags_part():
    deck = Document.new()
    slide = deck.add_slide("Blank")
    frame = slide.add_table(2, 2, 0, 0, 3000000, 1000000)
    kept.store(frame, {"v": 1, "source": "d:x", "first_row": 0,
                       "columns": [{"col": 0, "column": "a", "rows": "1-2"}]})
    other = Document.new()
    target = other.add_slide("Blank")
    slide.copy_shapes([frame.id], slide, dx=10)
    slide.copy_shapes([frame.id], target)
    assert len(tag_parts(deck.package)) == 2 and len(tag_parts(other.package)) == 1
    assert len(deck.provenance()) == 2 and len(other.provenance()) == 1
    assert deck.validate() == [] and other.validate() == []


def test_deleting_the_shape_or_the_slide_leaves_no_tags_part(toolbox, session):
    table = table_from(toolbox, session)
    chart_from(toolbox, session)
    ok(call(toolbox, session, "ppt_arrange", doc="d1", targets=[table], action="delete"))
    assert tag_parts(deck_of(session).package) == ["ppt/tags/tag2.xml"]
    ok(call(toolbox, session, "ppt_add_slide", doc="d1", layout="Blank"))
    ok(call(toolbox, session, "ppt_manage_slides", doc="d1", action="delete", slide="s:256"))
    deck = deck_of(session)
    assert tag_parts(deck.package) == []
    assert deck.package.unreachable_parts() == set()
    assert reopened(session).validate() == []


# -- the frame's tags part -----------------------------------------------------------------------


def _frame_with(xml_tags: bytes | None = None, ext: bool = False):
    """A table frame; with ``xml_tags``, an add-in's tags part already on it, and with
    ``ext`` an ``p:nvPr/p:extLst`` as PowerPoint writes after an edit."""
    deck = Document.new()
    slide = deck.add_slide("Blank")
    frame = slide.add_table(2, 2, 0, 0, 3000000, 1000000)
    package = deck.package
    nv_pr = frame._element.find(".//{*}nvGraphicFramePr/{*}nvPr")
    if ext:
        from lxml import etree

        nv_pr.append(etree.fromstring(
            '<p:extLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14='
            '"http://schemas.microsoft.com/office/powerpoint/2010/main" val="1"/></p:ext>'
            '</p:extLst>'))
    if xml_tags is not None:
        from lxml import etree

        package.add_part("ppt/tags/tag1.xml", xml_tags, CT_TAGS, override=True)
        rel = package.add_relationship(slide.part_path, TAGS, "ppt/tags/tag1.xml")
        holder = etree.SubElement(nv_pr, "{http://schemas.openxmlformats.org/presentationml/"
                                  "2006/main}custDataLst")
        nv_pr.insert(0, holder)
        etree.SubElement(holder, "{http://schemas.openxmlformats.org/presentationml/2006/"
                         "main}tags").set(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rel)
    slide._touch()
    deck._reset_caches()
    return deck, deck.shape(frame.id)


RECORD = {"v": 1, "source": "d:x", "first_row": 0,
          "columns": [{"col": 0, "column": "a", "rows": "1-2"}]}
ADDIN = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:tagLst xmlns:p='
         b'"http://schemas.openxmlformats.org/presentationml/2006/main"><p:tag name="ADDIN" '
         b'val="keep me"/></p:tagLst>')


def test_custdatalst_goes_before_extlst():
    deck, frame = _frame_with(ext=True)
    kept.store(frame, RECORD)
    nv_pr = frame._element.find(".//{*}nvGraphicFramePr/{*}nvPr")
    assert [child.tag.rsplit("}", 1)[1] for child in nv_pr] == ["custDataLst", "extLst"]
    assert deck.validate() == []


def test_other_tags_stay_and_the_part_stays_for_them():
    deck, frame = _frame_with(ADDIN)
    assert kept.store(frame, RECORD)
    root = deck.package.tree("ppt/tags/tag1.xml")
    assert [tag.get("name") for tag in root] == ["ADDIN", "UVRT_PROVENANCE"]
    assert kept.read(frame) == RECORD
    assert not kept.store(frame, RECORD)                     # the same record: no change
    assert kept.store(frame, None)
    assert [tag.get("name") for tag in deck.package.tree("ppt/tags/tag1.xml")] == ["ADDIN"]
    assert frame._element.find(".//{*}custDataLst/{*}tags") is not None
    assert not kept.store(frame, None)
    assert deck.validate() == []


def test_a_shared_tags_part_is_copied_before_it_is_written():
    deck, frame = _frame_with(ADDIN)
    twin = frame.duplicate()
    # Make the twin share the original's tags part, as a careless tool might.
    rel = frame._element.find(".//{*}custDataLst/{*}tags").get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    twin._element.find(".//{*}custDataLst/{*}tags").set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", rel)
    twin._slide._touch()
    kept.store(twin, RECORD)
    assert kept.read(twin) == RECORD and kept.read(frame) is None
    assert [t.get("name") for t in deck.package.tree("ppt/tags/tag1.xml")] == ["ADDIN"]
    # And a drop from a shared part takes the tag from that frame alone.
    kept.store(frame, RECORD)
    shared = frame._element.find(".//{*}custDataLst/{*}tags").get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    twin._element.find(".//{*}custDataLst/{*}tags").set(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", shared)
    twin._slide._touch()
    assert kept.store(twin, None)
    assert kept.read(frame) == RECORD and kept.read(twin) is None
    assert deck.validate() == []


def test_an_unreadable_record_is_skipped_with_a_warning():
    deck, frame = _frame_with(ADDIN.replace(b'name="ADDIN" val="keep me"',
                                            b'name="UVRT_PROVENANCE" val="[1,2]"'))
    with pytest.warns(kept.ProvenanceWarning, match="no source record"):
        assert deck.provenance() == {}
    deck, frame = _frame_with(ADDIN.replace(b'name="ADDIN" val="keep me"',
                                            b'name="uvrt_provenance" val="{&quot;v&quot;:1}"'))
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        assert deck.provenance() == {frame.id: {"v": 1}}     # a name in any case is ours


def test_a_record_is_kept_on_a_graphic_frame_only():
    deck = Document.new()
    shape = deck.add_slide("Blank").add_shape("rect", 0, 0, 100, 100)
    with pytest.raises(ValueError, match="graphic frame"):
        kept.store(shape, RECORD)


def test_a_refused_edit_leaves_the_record(toolbox, session):
    target = table_from(toolbox, session)
    failed(call(toolbox, session, "ppt_edit_table", doc="d1", target=target,
                action="delete_row", index=40), "not_found")
    assert len(session.provenance("d1")[target]["columns"]) == 3
