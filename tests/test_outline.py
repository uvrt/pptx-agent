"""The Markdown outline (E6): ``to_outline`` reads a deck, ``insert_outline`` drafts slides.

Reading changes no byte and every id it writes resolves.  Drafting is held to the gates
every edit is held to -- the validity checks, save and reopen with nothing lost, one undo
step back to the bytes before and redo, every slide's full-state SVG applied to a fresh
copy changing no byte, and a pptx2svg render with the text in the placeholder it was put
in -- for a hand-written outline into a deck from nothing, and for every fixture's own
outline into a new deck from that fixture.  The round trip (``to_outline``,
``insert_outline``, ``to_outline``) is stable on every fixture up to what
:mod:`pptx_agent.outline` says it does not carry.
"""

from __future__ import annotations

import io
import re
import warnings
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import pytest
from lxml import etree

import oracle as oracle_helper
from conftest import FIXTURE_DIR, fixture_paths
from images import MARKER, ORANGE
from pptx_agent import Document, OutlineWarning
from pptx_agent.outline.markdown import inline_markdown, normalise_inline, parse
from pptx_agent.oxml.xml import find, qn
from test_newdeck import assert_new_deck_valid
from test_validity import assert_valid, requires_powerpoint

CREATED = datetime(2026, 10, 3, 12, 0, 0, tzinfo=timezone.utc)
HOME = Path.home()

HAND = """\
# Quarterly review

Prepared for the board, October 2026

# Where we are

- Revenue grew **12%** to *4,285*
- Margins held
  - Gross margin 41%
  - Operating margin `11.9%`
- See [the report](https://example.org/q3)

Notes:

Open with the headline number.

- Pause after the margins

# Plans

Pros:

- Faster edits
- Fewer prompts

Cons:

- Stale caches

# Results

| Metric | Q2 | Q3 |
| --- | --- | --- |
| Revenue | 3,890 | **4,285** |
| Margin | 11.3% | 11.9% |

# Part two

# Steps

1. Plan
2. Build
3. ~~Ship~~ Launch

---

A second column

# The picture

![A red marker](marker.png)
"""

#: Per slide of HAND: the layout the content rules choose, and the text each shape takes.
HAND_LAYOUTS = ["Title Slide", "Title and Content", "Two Content", "Title Only",
                "Section Header", "Two Content", "Title Only"]


def _parts(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _hand_deck() -> Document:
    document = Document.new(created=CREATED)
    document.insert_outline(HAND, images={"marker.png": MARKER})
    return document


def _template_deck(path: Path) -> tuple[Document, Document, str]:
    """``(source, drafted, outline)``: the fixture's outline inserted into a new deck from
    the fixture."""
    source = Document.open(path)
    outline = source.to_outline()
    drafted = Document.new(template=path, created=CREATED)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", OutlineWarning)
        drafted.insert_outline(outline, images=source)
    return source, drafted, outline


def _ids(outline: str) -> list[str]:
    found = []
    for text in re.findall(r"<!--(.*?)-->", outline):
        words = text.split()
        if words and (re.fullmatch(r"s:\d+", words[0]) or re.fullmatch(r"\d+\.\S+", words[0])):
            found.append(words[0])
    return found


def strip(outline: str) -> str:
    """The outline up to what the round trip does not carry (see pptx_agent.outline): ids,
    media part names, and chart and SmartArt blocks."""
    blocks = re.split(r"\n(?=<!-- )", outline)
    kept = [b for b in blocks if not re.match(r"<!-- \S+ (chart|smartart)\b", b)]
    text = "\n".join(kept)
    text = re.sub(r"<!-- s:\d+ ?", "<!-- ", text)
    text = re.sub(r"<!-- \d+\.\S+ ?", "<!-- ", text)
    text = re.sub(r" <!-- -->", "", text)
    return re.sub(r"\]\(ppt/media/[^)]*\)", "](media)", text)


# ------------------------------------------------------------------------------------------
# Reading
# ------------------------------------------------------------------------------------------


def test_reading_changes_no_byte(pptx_path):
    document = Document.open(pptx_path)
    before = document.to_bytes()
    for ids in (True, False):
        for notes in (True, False):
            document.to_outline(ids=ids, notes=notes)
    assert document.package.dirty_parts == frozenset()
    assert document.to_bytes() == before
    assert not document.history.can_undo()


def test_every_slide_and_id_is_in_the_outline_and_resolves(pptx_path):
    document = Document.open(pptx_path)
    outline = document.to_outline()
    ids = _ids(outline)
    assert [i for i in ids if i.startswith("s:")] == [f"s:{s.slide_id}" for s in document.slides]
    for identifier in ids:
        if identifier.startswith("s:"):
            document.slide(int(identifier[2:]))
        else:
            document.shape(identifier)
    # Every shape with text is named, groups' children included.
    def walk(shapes):
        for shape in shapes:
            yield shape
            if shape.kind == "group":
                yield from walk(shape.children)
    for slide in document.slides:
        for shape in walk(slide.shapes):
            if shape.kind == "shape" and shape.text.strip():
                assert shape.id in ids, shape
    # The slides are the outline's level-one headings.
    headings = [b for b in parse(outline) if b[0] == "h" and b[1] == 1]
    assert len(headings) == len(document.slides)


def test_the_outline_of_a_deck(sample_deck=FIXTURE_DIR / "sample.pptx"):
    outline = Document.open(sample_deck).to_outline()
    assert outline.startswith(
        "<!-- s:256 layout: Title Slide -->\n# md-pptx サンプル <!-- 256.2 -->\n\n"
        "<!-- 256.3 ph: subTitle 1 -->\nVS Code拡張の動作確認用\n\n")
    assert (
        "<!-- s:257 layout: Title and Content -->\n# 概要 <!-- 257.2 -->\n\n"
        "<!-- 257.3 ph: obj 1 -->\n- MarkdownからPPTXを生成するツール\n"
        "- **テンプレートPPTX**のスライドマスターを活用\n- *編集可能な*PPTXを出力\n\n"
        "<!-- 257/notes -->\nNotes:\n\nここがプレゼンターノートになります\n\n") in outline
    assert "- 通常テキスト、**太字テキスト**、*斜体テキスト*、`コード`、取り消し線\n" \
           "- [リンクのサンプル](https://example.com)\n" in outline
    assert "- 箇条書き項目2\n  - ネストされた項目\n    - さらにネスト\n- 箇条書き項目3\n" in outline
    assert "お問い合わせ: [info@example.com](mailto:info@example.com)\n" in outline


def test_tables_pictures_charts_and_smartart():
    basic = Document.open(FIXTURE_DIR / "real-basic-theme.pptx")
    outline = basic.to_outline()
    assert ("<!-- 257.93 ph: body 1 -->\nほんぶん\n\n- かじょうがき１\n- かじょうがき２\n"
            "- かじょうがき３\n") in outline
    assert ("<!-- 257.97 table -->\n| てーぶる | てーぶる |\n| --- | --- |\n"
            "| てーぶる | てーぶる |\n| てーぶる | てーぶる |\n") in outline
    assert basic.resolve("257.97/cell2,1").text == "てーぶる"  # the cells a table's id implies
    assert "<!-- 257.98 picture -->\n![](ppt/media/image1.png)\n" in outline

    contract = Document.open(FIXTURE_DIR / "authoring-integration.pptx").to_outline()
    assert ("<!-- 256.6 chart -->\n[chart: column; title: Chart contract; series: Coverage; "
            "categories: Reader, Writer, Renderer]\n") in contract
    smartart = Document.open(FIXTURE_DIR / "powerpoint-smartart.pptx").to_outline()
    assert ("<!-- 257.3 smartart -->\n- Goals\n  - Faster edits\n  - Fewer prompts\n- Risks\n"
            "  - Stale caches\n") in smartart


def test_slides_can_be_chosen_and_ids_and_notes_left_out():
    document = Document.open(FIXTURE_DIR / "sample.pptx")
    second = document.slides[1]
    assert document.to_outline([2]) == document.to_outline(["s:257"]) == \
        document.to_outline(second)
    assert document.to_outline([2]).startswith("<!-- s:257 ")
    plain = document.to_outline(ids=False, notes=False)
    assert "<!--" not in plain and "Notes:" not in plain
    assert "# 概要\n\n- MarkdownからPPTXを生成するツール" in plain
    with pytest.raises(IndexError):
        document.to_outline([7])
    with pytest.raises(KeyError):
        document.to_outline(["s:999"])


def test_text_is_escaped_so_it_reads_back_as_text():
    document = Document.new(created=CREATED)
    slide = document.add_slide("Title Only")
    tricky = ["*not emphasis* [a] `b` <c> 1. ~x~ # | & &amp; \\", "Notes:", "- not a list",
              "1) not numbered", "bold and plain"]
    box = slide.add_textbox(914400, 914400, 4572000, 914400, "\n".join(tricky))
    document.resolve(f"{box.id}/p4").segment(["bold", " and plain"])
    document.resolve(f"{box.id}/p4/r0").bold = True
    outline = document.to_outline(ids=False)
    assert "Notes\\:" in outline
    blocks = [b for b in parse(outline) if b[0] == "p"]
    texts = ["".join(i[1] for i in b[1] if i[0] == "t") for b in blocks]
    assert texts == tricky
    assert blocks[-1][1][0] == ("t", "bold", frozenset({"strong"}), None)
    # Drafted again, the text is the same.
    again = Document.new(created=CREATED)
    again.insert_outline(outline)
    assert again.to_outline(ids=False) == outline


#: Text at the start of a line: what CommonMark would read as a block is escaped, and only that.
LINE_STARTS = {
    "11.9%": "11.9%", "+0.6pt YoY": "+0.6pt YoY", "-5%": "-5%", "#1 in Japan": "#1 in Japan",
    "2025.10.1": "2025.10.1", "=A1": "=A1", "1)a": "1)a",
    "1. one": "1\\. one", "12) twelve": "12\\) twelve", "+ plus": "\\+ plus",
    "- minus": "\\- minus", "-": "\\-", "---": "\\---", "- - -": "\\- - -", "===": "\\===",
    "# heading": "\\# heading", "###": "\\###", "> quote": "\\> quote", ">x": "\\>x",
    "3.": "3\\.",
}


@pytest.mark.parametrize("text", sorted(LINE_STARTS))
def test_only_what_would_start_a_block_is_escaped(text):
    written = inline_markdown([("t", text, frozenset(), None)])
    assert written == LINE_STARTS[text]
    blocks = parse(written)
    assert [b[0] for b in blocks] == ["p"]
    assert blocks[0][1] == (("t", text, frozenset(), None),)
    # Bold, as the pilot's KPI figures were: no escape where none is needed.
    bold = inline_markdown([("t", text, frozenset({"strong"}), None)])
    assert [b[1] for b in parse(bold) if b[0] == "p"] == \
        [(("t", text, frozenset({"strong"}), None),)]


@pytest.mark.parametrize("items", [
    [("t", "a", frozenset({"strong"}), None), ("t", "b", frozenset({"em"}), None)],
    [("t", "x*", frozenset({"strong"}), None), ("t", "y", frozenset(), None)],
    [("t", "under_score", frozenset({"em", "strong"}), None), ("br",), ("t", "z", frozenset(), "https://e.org/(1)")],
    [("t", " lead", frozenset({"strike"}), None), ("t", "``code``", frozenset({"code"}), None)],
])
def test_inline_markdown_reads_back_or_drops_marks(items):
    """What CommonMark can say is read back exactly; marks it cannot place are dropped from
    the block, and the text never changes."""
    text = inline_markdown(items)
    read = [b for b in parse(text) if b[0] == "p"][0][1]
    expected = normalise_inline(items)

    def characters(inline):
        return [(i[1], i[3]) if i[0] == "t" else i for i in normalise_inline(
            [("t", i[1], frozenset(), i[3]) if i[0] == "t" else i for i in inline])]

    assert characters(read) == characters(expected)
    marks = {m for i in read if i[0] == "t" for m in i[2]}
    assert read == expected or marks <= {"code"} | ({"strong", "em"} if "strike" not in marks
                                                     else set())


# ------------------------------------------------------------------------------------------
# Drafting
# ------------------------------------------------------------------------------------------


def test_the_content_chooses_the_layout():
    document = _hand_deck()
    assert [s.layout.name for s in document.slides] == HAND_LAYOUTS
    titles = [next(s for s in slide.shapes if s.placeholder and s.placeholder[0] in
                   ("title", "ctrTitle")).text for slide in document.slides]
    assert titles == ["Quarterly review", "Where we are", "Plans", "Results", "Part two",
                      "Steps", "The picture"]
    cover = document.slides[0]
    assert [(s.placeholder, s.text) for s in cover.shapes] == [
        (("ctrTitle", None), "Quarterly review"),
        (("subTitle", 1), "Prepared for the board, October 2026")]
    plans = document.slides[2]
    assert [(s.placeholder, s.text) for s in plans.shapes][1:] == [
        ((None, 1), "Pros:\nFaster edits\nFewer prompts"), ((None, 2), "Cons:\nStale caches")]
    # Placeholders nothing went into are gone: the section header's text, the title only
    # slide's nothing.
    assert [s.placeholder for s in document.slides[4].shapes] == [("title", None)]
    steps = document.slides[5]
    assert [s.text for s in steps.shapes][1:] == ["Plan\nBuild\nShip Launch", "A second column"]
    # An empty heading and nothing else is a blank slide; heading only on the first slide a
    # title slide.
    blank = Document.new(created=CREATED)
    blank.insert_outline("#\n\n# Cover\n")
    assert [s.layout.name for s in blank.slides] == ["Blank", "Section Header"]
    first = Document.new(created=CREATED)
    first.insert_outline("# Cover\n")
    assert first.slides[0].layout.name == "Title Slide"


def test_bullets_marks_and_links_become_runs():
    document = _hand_deck()
    body = document.slides[1].shapes[1]
    assert body.placeholder == (None, 1)
    assert [p.level for p in body.paragraphs] == [0, 0, 1, 1, 0]
    assert all(p.bullet is None for p in body.paragraphs)  # inherited from the master
    revenue = body.paragraphs[0]
    assert [(r.text, r.bold, r.italic) for r in revenue.runs] == [
        ("Revenue grew ", None, None), ("12%", True, None), (" to ", None, None),
        ("4,285", None, True)]
    assert body.paragraphs[3].runs[1].typeface == "Consolas"
    assert body.paragraphs[4].runs[1].hyperlink.address == "https://example.org/q3"
    # Paragraphs in a bulleted placeholder do without; a numbered list is numbered; a
    # struck word is struck.
    plans = document.slides[2].shapes[1]
    assert plans.paragraphs[0].bullet is not None and plans.paragraphs[0].bullet.kind == "none"
    steps = document.slides[5].shapes[1]
    assert [p.bullet.kind for p in steps.paragraphs] == ["number"] * 3
    assert steps.paragraphs[2].runs[0].strike is True
    for slide in document.slides:
        for shape in slide.shapes:
            if shape.kind == "shape":
                for paragraph in shape.paragraphs:
                    for run in paragraph.runs:
                        assert run._rpr().get("lang") == "en-US"


def test_a_table_and_a_picture():
    document = _hand_deck()
    table = document.slides[3].shapes[1]
    assert table.has_table
    cells = table.table
    assert [[cells.cell(r, c).text for c in range(3)] for r in range(3)] == [
        ["Metric", "Q2", "Q3"], ["Revenue", "3,890", "4,285"], ["Margin", "11.3%", "11.9%"]]
    assert document.resolve(f"{table.id}/cell1,2/p0/r0").bold is True
    picture = document.slides[6].shapes[1]
    assert picture.kind == "picture" and picture.image_part.startswith("ppt/media/")
    assert find(picture._element, "p:nvPicPr/p:cNvPr").get("descr") == "A red marker"
    assert document.package.read(picture.image_part) == MARKER
    # Fitted into the body area, keeping its 2:1 shape.
    assert abs(picture.width / picture.height - 2) < 0.01


def test_a_picture_fills_a_picture_placeholder():
    document = Document.new(created=CREATED)
    document.insert_outline("<!-- layout: Picture with Caption -->\n# Square\n\n"
                            "![Orange](o.png)\n\nThe caption\n", images={"o.png": ORANGE})
    slide = document.slides[0]
    assert slide.layout.name == "Picture with Caption"
    picture = next(s for s in slide.shapes if s.kind == "picture")
    assert picture.placeholder == ("pic", 1)
    assert find(picture._element, "p:blipFill/a:srcRect") is not None  # cropped to fill
    caption = next(s for s in slide.shapes if s.placeholder == ("body", 2))
    assert caption.text == "The caption"
    assert "<!-- 256.3 picture ph: pic 1 -->" in document.to_outline() or \
        "picture ph: pic 1 -->" in document.to_outline()


def test_speaker_notes_get_a_notes_page_and_master():
    document = _hand_deck()
    data = document.to_bytes()
    parts = _parts(data)
    assert "ppt/notesMasters/notesMaster1.xml" in parts and "ppt/notesSlides/notesSlide1.xml" in parts
    assert sum(name.startswith("ppt/notesSlides/") and name.endswith(".xml") for name in parts) == 1
    presentation = etree.fromstring(parts["ppt/presentation.xml"])
    assert find(presentation, "p:notesMasterIdLst/p:notesMasterId") is not None
    # The notes master's theme is the slide master's, as PowerPoint copies it.
    assert parts["ppt/theme/theme2.xml"] == parts["ppt/theme/theme1.xml"]
    outline = document.to_outline()
    assert "Notes:\n\nOpen with the headline number.\n\n- Pause after the margins\n" in outline
    app = parts["docProps/app.xml"].decode()
    assert "<Notes>1</Notes>" in app
    # A template with a notes master gets no second one.
    template = Document.new(template=FIXTURE_DIR / "sample.pptx", created=CREATED)
    template.insert_outline("# One\n\nNotes:\n\nSay it\n")
    names = _parts(template.to_bytes())
    assert sum(n.startswith("ppt/notesMasters/") and n.endswith(".xml") for n in names) == 1
    assert "Notes:\n\nSay it\n" in template.to_outline()


def test_layouts_are_found_by_type_in_a_localised_template():
    dutch = {"Title Slide": "Titeldia", "Title and Content": "Titel en object",
             "Section Header": "Sectiekop", "Two Content": "Twee objecten",
             "Title Only": "Alleen titel", "Blank": "Leeg"}
    document = Document.new(created=CREATED)
    for layout in document.layouts:
        if layout.name in dutch:
            root = document.package.tree(layout.part_path)
            find(root, "p:cSld").set("name", dutch[layout.name])
            document.package.mark_dirty(layout.part_path)
    template = document.to_bytes()
    drafted = Document.new(template=template, created=CREATED)
    drafted.insert_outline(HAND, images={"marker.png": MARKER})
    assert [s.layout.name for s in drafted.slides] == [dutch[n] for n in HAND_LAYOUTS]


def test_a_hand_written_outline_drafts_into_any_template(pptx_path):
    """Google Slides' layout types (``tx``, ``twoColTx``) stand in for PowerPoint's; a
    template without placeholders gets text boxes, the title in the title's place."""
    document = Document.new(template=pptx_path, created=CREATED)
    empty = document.to_bytes()
    document.insert_outline(HAND, images={"marker.png": MARKER})
    assert_valid(document.to_bytes(), empty)
    names = [s.layout.name for s in document.slides]
    if pptx_path.stem == "real-basic-theme":
        assert names == ["TITLE", "TITLE_AND_BODY", "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY",
                         "SECTION_HEADER", "TITLE_AND_TWO_COLUMNS", "TITLE_ONLY"]
    elif pptx_path.stem == "sample":
        assert names == HAND_LAYOUTS
    texts = [[s.text for s in slide.shapes if s.kind == "shape"] for slide in document.slides]
    assert [t[0] for t in texts] == ["Quarterly review", "Where we are", "Plans", "Results",
                                     "Part two", "Steps", "The picture"]


def test_a_named_layout_and_the_layout_map_override_the_rules():
    document = Document.new(created=CREATED)
    document.insert_outline("<!-- layout: Comparison -->\n# A\n\n- x\n\n# B\n\n- y\n\n# C\n",
                            layout_map={"obj": "Content with Caption", "Section Header":
                                        "Title Only"})
    assert [s.layout.name for s in document.slides] == ["Comparison", "Content with Caption",
                                                        "Title Only"]
    with pytest.raises(KeyError):
        document.insert_outline("# D\n\n- z\n", layout_map={"obj": "No such layout"})


def test_at_inserts_where_asked_and_returns_the_slides():
    document = Document.new(created=CREATED)
    document.insert_outline("# One\n\n- a\n\n# Three\n\n- c\n")
    made = document.insert_outline("# Two\n\n- b\n", at=1)
    assert [s.slide_id for s in made] == [document.slides[1].slide_id]
    titles = [s.shapes[0].text for s in document.slides]
    assert titles == ["One", "Two", "Three"]


def test_insertion_is_one_undo_step():
    document = Document.new(created=CREATED)
    document.insert_outline("# Before\n")
    before = document.to_bytes()
    document.insert_outline(HAND, images={"marker.png": MARKER})
    after = document.to_bytes()
    assert document.undo()
    assert document.to_bytes() == before
    assert document.redo()
    assert document.to_bytes() == after


def test_bad_outlines_are_refused_and_change_nothing():
    document = Document.new(created=CREATED)
    before = document.to_bytes()
    with pytest.raises(ValueError, match="starts with"):
        document.insert_outline("Text first\n\n# Then a slide\n")
    with pytest.raises(ValueError, match="image"):
        document.insert_outline("# Picture\n\n![x](missing.png)\n")
    with pytest.raises(IndexError):
        document.insert_outline("# A\n", at=5)
    assert document.to_bytes() == before and not document.history.can_undo()


def test_charts_and_smartart_are_not_drafted():
    document = Document.new(created=CREATED)
    with pytest.warns(OutlineWarning):
        document.insert_outline("# Sales\n\n[chart: column; series: A; categories: Q1]\n\n- kept\n")
    assert [s.text for s in document.slides[0].shapes] == ["Sales", "kept"]
    outline = Document.open(FIXTURE_DIR / "powerpoint-smartart.pptx").to_outline()
    fresh = Document.new(created=CREATED)
    with pytest.warns(OutlineWarning):
        fresh.insert_outline(outline)


def test_hidden_slides_and_empty_named_placeholders_are_kept():
    document = Document.new(created=CREATED)
    document.insert_outline("<!-- hidden layout: Title and Content -->\n# Hidden\n\n"
                            "<!-- ph: obj 1 empty -->\n")
    slide = document.slides[0]
    assert document.package.tree(slide.part_path).get("show") == "0"
    assert [s.placeholder for s in slide.shapes] == [("title", None), (None, 1)]
    assert "<!-- s:256 hidden layout: Title and Content -->" in document.to_outline()
    assert "ph: obj 1 empty -->" in document.to_outline()


def test_lists_outside_a_bulleted_placeholder_get_powerpoints_bullets():
    document = Document.new(created=CREATED)
    document.insert_outline("# Free\n\n<!-- 256.9 -->\n- one\n  - two\n\n1. first\n\n"
                            "<!-- 256.3 ph: subTitle 1 -->\n")
    box = next(s for s in document.slides[0].shapes if s.placeholder is None)
    bullets = [p.bullet for p in box.paragraphs]
    assert [(b.kind, b.char or b.scheme, b.font) for b in bullets] == [
        ("char", "•", "Arial"), ("char", "•", "Arial"), ("number", "arabicPeriod", "+mj-lt")]
    assert [(p.level, p.margin_left, p.indent) for p in box.paragraphs] == [
        (0, 285750, -285750), (1, 742950, -285750), (0, 514350, -514350)]
    assert "- one\n  - two\n\n1. first\n" in document.to_outline()


def test_a_level_that_skips_one_is_written_one_deeper():
    document = Document.new(created=CREATED)
    slide = document.add_slide("Title and Content")
    body = slide.shapes[1]
    body.set_text("top\ndeep\nback")
    body.paragraphs[1].level = 2
    assert "- top\n  - deep\n- back\n" in document.to_outline()


def test_a_comment_before_the_next_heading_still_names_its_placeholder():
    document = Document.new(created=CREATED)
    document.insert_outline("<!-- layout: Title and Content -->\n# One\n\n<!-- ph: obj 1 empty -->\n"
                            "# Two\n\n- b\n")
    assert [s.placeholder for s in document.slides[0].shapes] == [("title", None), (None, 1)]


def test_images_are_found_in_a_directory_by_a_callable_or_a_path(tmp_path):
    (tmp_path / "m.png").write_bytes(MARKER)
    for images, source in ((tmp_path, "m.png"), (lambda src: ORANGE, "any"),
                           (None, str(tmp_path / "m.png"))):
        document = Document.new(created=CREATED)
        document.insert_outline(f"# P\n\n![p]({source})\n", images=images)
        picture = next(s for s in document.slides[0].shapes if s.kind == "picture")
        assert document.package.read(picture.image_part) in (MARKER, ORANGE)


def test_the_notes_page_and_master_are_written_as_powerpoint_writes_them():
    document = _hand_deck()
    parts = _parts(document.to_bytes())
    master = etree.fromstring(parts["ppt/notesMasters/notesMaster1.xml"])
    frames = []
    for shape in master.iter(qn("p:sp")):
        ph = shape.find(f".//{qn('p:ph')}")
        off, ext = shape.find(f".//{qn('a:off')}"), shape.find(f".//{qn('a:ext')}")
        frames.append((ph.get("type"), ph.get("idx"), int(off.get("x")), int(off.get("y")),
                       int(ext.get("cx")), int(ext.get("cy"))))
    assert frames == [("hdr", None, 0, 0, 2971800, 458788),
                      ("dt", "1", 3884613, 0, 2971800, 458788),
                      ("sldImg", "2", 685800, 1143000, 5486400, 3086100),
                      ("body", "3", 685800, 4400550, 5486400, 3600450),
                      ("ftr", "4", 0, 8685213, 2971800, 458787),
                      ("sldNum", "5", 3884613, 8685213, 2971800, 458787)]
    styles = find(master, "p:notesStyle")
    assert [l.find(qn("a:defRPr")).get("sz") for l in styles] == ["1200"] * 9
    page = etree.fromstring(parts["ppt/notesSlides/notesSlide1.xml"])
    assert [(ph.get("type"), ph.get("idx")) for ph in page.iter(qn("p:ph"))] == [
        ("sldImg", None), ("body", "1"), ("sldNum", "5")]
    # At 4:3 the slide image is the slide's shape, fitted into the same box.
    narrow = Document.new(size="4:3", created=CREATED)
    narrow.insert_outline("# A\n\nNotes:\n\nB\n")
    image = next(s for s in etree.fromstring(_parts(narrow.to_bytes())[
        "ppt/notesMasters/notesMaster1.xml"]).iter(qn("p:sp"))
        if s.find(f".//{qn('p:ph')}").get("type") == "sldImg")
    assert (int(image.find(f".//{qn('a:ext')}").get("cx")),
            int(image.find(f".//{qn('a:ext')}").get("cy"))) == (4114800, 3086100)


# ------------------------------------------------------------------------------------------
# The gates, on drafted decks
# ------------------------------------------------------------------------------------------


def _drafted_cases():
    return ["hand"] + [path.stem for path in fixture_paths()]


def _drafted(case: str) -> tuple[Document, bytes]:
    """The drafted deck, and the deck it was drafted into, as bytes."""
    if case == "hand":
        base = Document.new(created=CREATED)
        empty = base.to_bytes()
        base.insert_outline(HAND, images={"marker.png": MARKER})
        return base, empty
    path = FIXTURE_DIR / f"{case}.pptx"
    empty = Document.new(template=path, created=CREATED).to_bytes()
    _, drafted, _ = _template_deck(path)
    return drafted, empty


@pytest.mark.parametrize("case", _drafted_cases())
def test_drafted_decks_pass_the_gates(case):
    document, empty = _drafted(case)
    data = document.to_bytes()
    if case == "hand":
        assert_new_deck_valid(data)
    else:
        assert_valid(data, empty)
    # Saved and reopened, nothing is lost.
    reopened = Document.open(data)
    assert reopened.to_bytes() == data
    assert reopened.to_outline() == document.to_outline()
    # One undo step back to the deck it was drafted into; redo forward again.
    assert document.undo()
    assert document.to_bytes() == empty
    assert document.redo()
    assert document.to_bytes() == data


@pytest.mark.parametrize("case", _drafted_cases())
def test_drafted_decks_survive_their_full_state_svgs(case):
    pytest.importorskip("pptx2svg")
    document, _ = _drafted(case)
    data = document.to_bytes()
    for slide in Document.open(data).slides:
        copy = Document.open(data)
        report = copy.apply_svg(slide.render_svg(full_state=True))
        assert not report, slide
        assert copy.to_bytes() == data


@pytest.mark.parametrize("case", _drafted_cases())
def test_pptx2svg_draws_the_text_in_its_placeholders(case):
    pytest.importorskip("pptx2svg")
    document, _ = _drafted(case)
    svgs = document.render_svg()
    assert len(svgs) == len(document.slides)
    for slide in document.slides:
        svg = etree.fromstring(slide.render_svg().encode())
        groups = {g.get("data-pptx-id"): g for g in svg.iter("{http://www.w3.org/2000/svg}g")
                  if g.get("data-pptx-id")}
        for shape in slide.shapes:
            if shape.kind != "shape" or not shape.text.strip():
                continue
            drawn = re.sub(r"\s+", "", "".join(t.text or "" for t in groups[shape.id].iter(
                "{http://www.w3.org/2000/svg}tspan")))
            for paragraph in shape.paragraphs:
                for run in paragraph.runs:  # pptx2svg draws a line's runs by face, and
                    # a wrapped line without the space it wrapped at
                    assert re.sub(r"\s+", "", run.text) in drawn, (shape.id, run.text, drawn)


# ------------------------------------------------------------------------------------------
# The round trip
# ------------------------------------------------------------------------------------------


def test_the_round_trip_is_stable(pptx_path):
    _, drafted, outline = _template_deck(pptx_path)
    again = drafted.to_outline()
    assert strip(again) == strip(outline)
    # And once more: drafting the drafted deck's outline changes nothing further.
    third = Document.new(template=pptx_path, created=CREATED)
    third.insert_outline(again, images=drafted)
    assert strip(third.to_outline()) == strip(again)


def test_the_round_trip_without_ids_keeps_the_text(pptx_path):
    source = Document.open(pptx_path)
    outline = source.to_outline(ids=False)
    drafted = Document.new(template=pptx_path, created=CREATED)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", OutlineWarning)
        drafted.insert_outline(outline, images=source)

    def without_charts(text: str) -> str:
        text = re.sub(r"\]\(ppt/media/[^)]*\)", "](media)", text)
        lines = [line for line in text.split("\n") if not line.startswith("[chart: ")]
        kept: list[str] = []
        for line in lines:  # a separator with nothing after it on its slide goes too
            if line == "" and kept and kept[-1] == "":
                continue
            if line.startswith("#") and len(kept) >= 2 and kept[-2] == "---":
                del kept[-2:]
            if line == "---" and len(kept) >= 2 and kept[-2] == "---":
                kept.pop()
                continue
            kept.append(line)
        while kept and kept[-1] in ("", "---"):
            kept.pop()
        return "\n".join(kept)

    assert without_charts(drafted.to_outline(ids=False)) == without_charts(outline)


def test_the_round_trip_of_a_hand_written_outline():
    first = _hand_deck().to_outline()
    again = Document.new(created=CREATED)
    again.insert_outline(first, images=_hand_deck())
    assert strip(again.to_outline()) == strip(first)


# ------------------------------------------------------------------------------------------
# PowerPoint (pytest -m oracle)
# ------------------------------------------------------------------------------------------


@pytest.mark.oracle
@requires_powerpoint
@pytest.mark.parametrize("case", _drafted_cases())
def test_powerpoint_opens_drafted_decks(case):
    """E6's acceptance.  A deck drafted from a hand-written outline into a deck from
    nothing, and each fixture's own outline drafted into a new deck from that fixture:
    PowerPoint exports it unprompted, with one page per slide and every slide's text on its
    page."""
    document, _ = _drafted(case)
    stem = f"pptx-agent-e6-{case}"
    deck, pdf = HOME / f"{stem}.pptx", HOME / f"{stem}.pdf"
    try:
        deck.write_bytes(document.to_bytes())
        result = oracle_helper.export_pdf(deck, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert oracle_helper.pdf_page_count(pdf) == len(document.slides)
        texts = oracle_helper.pdf_texts(pdf)
        if texts is not None:
            for page, slide in zip(texts, document.slides):
                flat = re.sub(r"\s+", "", page)
                for shape in slide.shapes:
                    if shape.kind == "shape":
                        for word in shape.text.split():
                            assert re.sub(r"\s+", "", word) in flat, (slide.slide_id, word)
                    elif shape.has_table:
                        cell = shape.table.cell(0, 0).text
                        assert re.sub(r"\s+", "", cell) in flat, (slide.slide_id, cell)
    finally:
        oracle_helper.cleanup(deck, pdf)


@pytest.mark.oracle
@requires_powerpoint
@pytest.mark.parametrize("case", _drafted_cases())
def test_powerpoint_keeps_a_drafted_deck_as_drafted(case):
    """Saved again by PowerPoint, a drafted deck reads as the same outline -- layouts,
    placeholders, text, lists and their levels, marks, links, tables, pictures and notes.
    (The parts PowerPoint adds on saving -- a thumbnail, and to some templates presProps or
    app.xml -- are the template's business, not the outline's.)"""
    document, _ = _drafted(case)
    stem = f"pptx-agent-e6-{case}"
    deck, again = HOME / f"{stem}.pptx", HOME / f"{stem}-resaved.pptx"
    try:
        data = document.to_bytes()
        deck.write_bytes(data)
        saved = oracle_helper.save_as_pptx(deck, again)
        assert saved.ok, f"PowerPoint {saved.outcome}: {saved.detail}"
        resaved = Document.open(again.read_bytes())
        assert strip(resaved.to_outline()) == strip(document.to_outline())
    finally:
        oracle_helper.cleanup(deck, again)
