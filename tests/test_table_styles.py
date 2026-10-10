"""Table styles and Table Design flags (production feedback round 3, A4).

``deck.table_styles`` lists the deck's own styles (``ppt/tableStyles.xml``) and PowerPoint's
built-in ones; ``table.style`` reads and sets one by name or id; ``table.flags`` and
``set_flags`` are the Header Row, Total Row, First and Last Column and Banded Rows and
Columns check boxes.  The tools: ``ppt_add_table``'s ``style``, ``ppt_edit_table``'s
``set_style`` and ``set_flags``, and ``describe``'s ``table_styles``.
"""

from __future__ import annotations

import pytest
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.tablestyles import BUILTIN_TABLE_STYLES, DEFAULT_TABLE_STYLE, styles_part
from pptx_agent.oxml.xml import qn
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

EMU = 914400
BRAND_ID = "{8F44A2D9-6A0B-4E0C-9C51-2C4A0F2A7B11}"
#: A brand template's own style: an accent-1 header in white bold, a rule under it, a hairline
#: between rows, a light band -- what a corporate template carries.
BRAND_STYLE = f"""<a:tblStyle styleId="{BRAND_ID}" styleName="Acme Table"><a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"/><a:schemeClr val="tx1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:insideH><a:ln w="6350"><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln></a:insideH></a:tcBdr><a:fill><a:noFill/></a:fill></a:tcStyle></a:wholeTbl><a:band1H><a:tcStyle><a:tcBdr/><a:fill><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="20000"/><a:lumOff val="80000"/></a:schemeClr></a:solidFill></a:fill></a:tcStyle></a:band1H><a:firstRow><a:tcTxStyle b="on"><a:fontRef idx="minor"/><a:schemeClr val="bg1"/></a:tcTxStyle><a:tcStyle><a:tcBdr><a:bottom><a:ln w="19050"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></a:bottom></a:tcBdr><a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill></a:tcStyle></a:firstRow></a:tblStyle>"""


def branded(*, default: bool = False) -> Document:
    """A new deck whose tableStyles.xml carries the brand style (and names it ``def``)."""
    deck = Document.new(size="16:9")
    part = styles_part(deck)
    xml = deck.package.read(part).decode("utf-8")
    xml = xml.replace("/>", f">{BRAND_STYLE}</a:tblStyleLst>", 1) if xml.rstrip().endswith(
        "/>") else xml.replace("</a:tblStyleLst>", f"{BRAND_STYLE}</a:tblStyleLst>")
    if default:
        xml = xml.replace(DEFAULT_TABLE_STYLE, BRAND_ID, 1)
    deck.package.replace_part(part, xml.encode("utf-8"))
    return Document.open(deck.to_bytes())


def test_the_deck_lists_its_own_styles_then_powerpoints():
    deck = branded()
    styles = deck.table_styles
    assert styles[0] == {"id": BRAND_ID, "name": "Acme Table", "builtin": False}
    assert len(styles) == 1 + 74 == 1 + len(BUILTIN_TABLE_STYLES)
    default = [s for s in styles if s.get("default")]
    assert default == [{"id": DEFAULT_TABLE_STYLE, "name": "Medium Style 2 - Accent 1",
                        "builtin": True, "default": True}]
    assert [s for s in branded(default=True).table_styles if s.get("default")][0]["id"] \
        == BRAND_ID


def test_a_style_is_set_by_name_or_id_and_read_back():
    deck = branded()
    frame = deck.add_slide("Title Only").add_table(3, 3, EMU, EMU, 6 * EMU, 2 * EMU,
                                                   style="acme table")
    table = frame.table
    assert table.style == {"id": BRAND_ID, "name": "Acme Table", "builtin": False}
    table.style = "Light Style 1 – Accent 1"             # an en dash, any case: one name
    assert table.style["id"] == "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}"
    table.style = "{073a0daa-6af3-43ab-8588-cec1d06c72b9}"
    assert table.style == {"id": "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}",
                           "name": "Medium Style 2", "builtin": True}
    with pytest.raises(ValueError, match="the deck's own: Acme Table; built in: No Style"):
        table.style = "Brand Table"
    with pytest.raises(ValueError, match="no table style"):
        table.style = "{00000000-0000-0000-0000-000000000000}"
    # A built-in style is named by id and not copied into the deck, as PowerPoint writes it.
    assert "Medium Style 2" not in deck.package.read(styles_part(deck)).decode()
    assert deck.validate() == []
    assert deck.undo() and deck.shape(frame.id).table.style["name"] == "Light Style 1 - Accent 1"


def test_flags_are_read_and_set_as_powerpoint_writes_them():
    deck = Document.new()
    table = deck.add_slide("Title Only").add_table(4, 3, EMU, EMU, 6 * EMU, 2 * EMU).table
    assert table.flags == {"first_row": True, "last_row": False, "first_column": False,
                           "last_column": False, "banded_rows": True, "banded_columns": False}
    table.set_flags(banded_rows=False, first_column=True, last_row=True, banded_columns=True)
    properties = table._table().find(qn("a:tblPr"))
    assert list(properties.attrib.items()) == [("firstRow", "1"), ("firstCol", "1"),
                                               ("lastRow", "1"), ("bandCol", "1")]
    assert table.flags["first_column"] and not table.flags["banded_rows"]
    assert table.set_flags() is table
    assert deck.validate() == []
    assert deck.undo() and deck.slides[0].shapes[-1].table.flags["banded_rows"]


def test_a_table_without_a_style_id_says_none():
    deck = Document.new()
    table = deck.add_slide("Title Only").add_table(2, 2, EMU, EMU, 4 * EMU, EMU).table
    properties = table._table().find(qn("a:tblPr"))
    properties.remove(properties.find(qn("a:tableStyleId")))
    assert table.style is None
    table.style = "Medium Style 2"
    assert table.style["name"] == "Medium Style 2"


def test_the_tools_take_a_style_and_flags_and_describe_lists_the_styles():
    deck = branded()
    deck.add_slide("Title Only").title = "Spend"
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, workers=0) as toolbox:
        session = toolbox.session()
        doc = session.open(deck.to_bytes(), "brand.pptx")
        described = toolbox.dispatch(session, "describe", {"doc": doc}).data
        styles = described["theme"]["table_styles"]
        assert styles["default"] == "Medium Style 2 - Accent 1"
        assert styles["own"] == ["Acme Table"] and "Light Style 1|2|3" in styles["builtin"]
        added = toolbox.dispatch(session, "ppt_add_table", {
            "doc": doc, "slide": "s:256", "box": {"x": 36, "y": 100, "w": 600, "h": 120},
            "rows": 3, "columns": 2, "data": [["Item", "EUR m"], ["Rent", "1.2"]],
            "style": "Acme Table"})
        assert added.ok, added.to_json()
        assert added.data["style"] == "Acme Table"
        assert added.data["flags"] == ["first_row", "banded_rows"]
        target = added.created[0]
        flagged = toolbox.dispatch(session, "ppt_edit_table", {
            "doc": doc, "target": target, "action": "set_flags",
            "flags": {"banded_rows": False}})
        assert flagged.ok and flagged.data["flags"] == ["first_row"], flagged.to_json()
        styled = toolbox.dispatch(session, "ppt_edit_table", {
            "doc": doc, "target": target, "action": "set_style",
            "style": "Light Style 2 - Accent 1"})
        assert styled.ok and styled.data["style"] == "Light Style 2 - Accent 1"
        wrong = toolbox.dispatch(session, "ppt_edit_table", {
            "doc": doc, "target": target, "action": "set_style", "style": "Brand"})
        assert wrong.error.code == "not_found" and wrong.error.valid_options == ["Acme Table"]
        empty = toolbox.dispatch(session, "ppt_edit_table", {
            "doc": doc, "target": target, "action": "set_flags", "flags": {}})
        assert empty.error.code == "invalid_arguments"
        assert toolbox.dispatch(session, "ppt_add_table", {
            "doc": doc, "slide": "s:256", "box": {"x": 36, "y": 300, "w": 600, "h": 60},
            "rows": 1, "columns": 1, "style": "Nope"}).error.code == "not_found"


def test_the_renderer_draws_the_brand_style_and_the_flags():
    """pptx2svg draws a custom style from tableStyles.xml and follows the flags: the
    header in accent 1 and the band in its tint, then neither once both flags are off."""
    import re

    def fills(svg: str) -> set[str]:
        return set(re.findall(r'fill="(#[0-9a-f]{6})"', svg.lower()))

    deck = branded()
    slide = deck.add_slide("Title Only")
    table = slide.add_table(4, 2, EMU, EMU, 6 * EMU, 2 * EMU, style="Acme Table").table
    accent = deck.theme.colors["accent1"].lower()
    styled = fills(slide.render_svg())
    band = styled - {accent, "#ffffff", "#000000"}
    assert accent in styled and len(band) == 1          # accent 1 at 20%, lightened 80%
    table.set_flags(banded_rows=False, first_row=False)
    plain = fills(slide.render_svg())
    assert accent not in plain and not band & plain
