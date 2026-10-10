"""Table styles: the deck's own (``ppt/tableStyles.xml``) and PowerPoint's built-in ones.

A table names its style by GUID in ``a:tblPr/a:tableStyleId``; its Table Design flags --
header row, total row, first and last column, banded rows and columns -- are ``a:tblPr``'s
``firstRow``, ``lastRow``, ``firstCol``, ``lastCol``, ``bandRow`` and ``bandCol``, and say
which of the style's regions apply.  ``ppt/tableStyles.xml`` holds the deck's custom styles
(a brand template's own) and ``def``, the style PowerPoint gives a new table.  The built-in
styles live in PowerPoint: a deck names one by GUID and need not define it (see
:data:`BUILTIN_TABLE_STYLES`).
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING

from ..oxml.xml import Element, append_in_order, find, findall, make, qn, remove

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

REL_TABLE_STYLES = ("http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                    "tableStyles")

#: PowerPoint's built-in table styles, GUID -> name, as its Table Design gallery names them:
#: the 74 styles pptx2svg's catalogue measured from PowerPoint 16 (each GUID confirmed by
#: drawing it in PowerPoint: an id PowerPoint does not know draws as no style).
BUILTIN_TABLE_STYLES: dict[str, str] = {
    "{2D5ABB26-0587-4C30-8999-92F81FD0307C}": "No Style, No Grid",
    "{5940675A-B579-460E-94D1-54222C63F5DA}": "No Style, Table Grid",
    "{3C2FFA5D-87B4-456A-9821-1D502468CF0F}": "Themed Style 1 - Accent 1",
    "{284E427A-3D55-4303-BF80-6455036E1DE7}": "Themed Style 1 - Accent 2",
    "{69C7853C-536D-4A76-A0AE-DD22124D55A5}": "Themed Style 1 - Accent 3",
    "{775DCB02-9BB8-47FD-8907-85C794F793BA}": "Themed Style 1 - Accent 4",
    "{35758FB7-9AC5-4552-8A53-C91805E547FA}": "Themed Style 1 - Accent 5",
    "{08FB837D-C827-4EFA-A057-4D05807E0F7C}": "Themed Style 1 - Accent 6",
    "{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}": "Themed Style 2 - Accent 1",
    "{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}": "Themed Style 2 - Accent 2",
    "{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}": "Themed Style 2 - Accent 3",
    "{E269D01E-BC32-4049-B463-5C60D7B0CCD2}": "Themed Style 2 - Accent 4",
    "{327F97BB-C833-4FB7-BDE5-3F7075034690}": "Themed Style 2 - Accent 5",
    "{638B1855-1B75-4FBE-930C-398BA8C253C6}": "Themed Style 2 - Accent 6",
    "{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}": "Light Style 1",
    "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}": "Light Style 1 - Accent 1",
    "{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}": "Light Style 1 - Accent 2",
    "{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}": "Light Style 1 - Accent 3",
    "{D27102A9-8310-4765-A935-A1911B00CA55}": "Light Style 1 - Accent 4",
    "{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}": "Light Style 1 - Accent 5",
    "{68D230F3-CF80-4859-8CE7-A43EE81993B5}": "Light Style 1 - Accent 6",
    "{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}": "Light Style 2",
    "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}": "Light Style 2 - Accent 1",
    "{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}": "Light Style 2 - Accent 2",
    "{F2DE63D5-997A-4646-A377-4702673A728D}": "Light Style 2 - Accent 3",
    "{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}": "Light Style 2 - Accent 4",
    "{5A111915-BE36-4E01-A7E5-04B1672EAD32}": "Light Style 2 - Accent 5",
    "{912C8C85-51F0-491E-9774-3900AFEF0FD7}": "Light Style 2 - Accent 6",
    "{616DA210-FB5B-4158-B5E0-FEB733F419BA}": "Light Style 3",
    "{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}": "Light Style 3 - Accent 1",
    "{5DA37D80-6434-44D0-A028-1B22A696006F}": "Light Style 3 - Accent 2",
    "{8799B23B-EC83-4686-B30A-512413B5E67A}": "Light Style 3 - Accent 3",
    "{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}": "Light Style 3 - Accent 4",
    "{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}": "Light Style 3 - Accent 5",
    "{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}": "Light Style 3 - Accent 6",
    "{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}": "Medium Style 1",
    "{B301B821-A1FF-4177-AEE7-76D212191A09}": "Medium Style 1 - Accent 1",
    "{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}": "Medium Style 1 - Accent 2",
    "{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}": "Medium Style 1 - Accent 3",
    "{1E171933-4619-4E11-9A3F-F7608DF75F80}": "Medium Style 1 - Accent 4",
    "{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}": "Medium Style 1 - Accent 5",
    "{10A1B5D5-9B99-4C35-A422-299274C87663}": "Medium Style 1 - Accent 6",
    "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": "Medium Style 2",
    "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": "Medium Style 2 - Accent 1",
    "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}": "Medium Style 2 - Accent 2",
    "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}": "Medium Style 2 - Accent 3",
    "{00A15C55-8517-42AA-B614-E9B94910E393}": "Medium Style 2 - Accent 4",
    "{7DF18680-E054-41AD-8BC1-D1AEF772440D}": "Medium Style 2 - Accent 5",
    "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}": "Medium Style 2 - Accent 6",
    "{8EC20E35-A176-4012-BC5E-935CFFF8708E}": "Medium Style 3",
    "{6E25E649-3F16-4E02-A733-19D2CDBF48F0}": "Medium Style 3 - Accent 1",
    "{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}": "Medium Style 3 - Accent 2",
    "{EB344D84-9AFB-497E-A393-DC336BA19D2E}": "Medium Style 3 - Accent 3",
    "{EB9631B5-78F2-41C9-869B-9F39066F8104}": "Medium Style 3 - Accent 4",
    "{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}": "Medium Style 3 - Accent 5",
    "{2A488322-F2BA-4B5B-9748-0D474271808F}": "Medium Style 3 - Accent 6",
    "{D7AC3CCA-C797-4891-BE02-D94E43425B78}": "Medium Style 4",
    "{69CF1AB2-1976-4502-BF36-3FF5EA218861}": "Medium Style 4 - Accent 1",
    "{8A107856-5554-42FB-B03E-39F5DBC370BA}": "Medium Style 4 - Accent 2",
    "{0505E3EF-67EA-436B-97B2-0124C06EBD24}": "Medium Style 4 - Accent 3",
    "{C4B1156A-380E-4F78-BDF5-A606A8083BF9}": "Medium Style 4 - Accent 4",
    "{22838BEF-8BB2-4498-84A7-C5851F593DF1}": "Medium Style 4 - Accent 5",
    "{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}": "Medium Style 4 - Accent 6",
    "{E8034E78-7F5D-4C2E-B375-FC64B27BC917}": "Dark Style 1",
    "{125E5076-3810-47DD-B79F-674D7AD40C01}": "Dark Style 1 - Accent 1",
    "{37CE84F3-28C3-443E-9E96-99CF82512B78}": "Dark Style 1 - Accent 2",
    "{D03447BB-5D67-496B-8E87-E561075AD55C}": "Dark Style 1 - Accent 3",
    "{E929F9F4-4A8F-4326-A1B4-22849713DDAB}": "Dark Style 1 - Accent 4",
    "{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}": "Dark Style 1 - Accent 5",
    "{AF606853-7671-496A-8E4F-DF71F8EC918B}": "Dark Style 1 - Accent 6",
    "{5202B0CA-FC54-4496-8BCA-5EF66A818D29}": "Dark Style 2",
    "{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}": "Dark Style 2 - Accent 1/Accent 2",
    "{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}": "Dark Style 2 - Accent 3/Accent 4",
    "{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}": "Dark Style 2 - Accent 5/Accent 6",
}
#: How the built-in names go, for a reader who wants one without the list of 74.
BUILTIN_NAMES = ("No Style, No Grid; No Style, Table Grid; Themed Style 1|2 - Accent 1-6; "
                 "Light Style 1|2|3, Medium Style 1|2|3|4 and Dark Style 1, each plain or "
                 "- Accent 1-6; Dark Style 2, plain or - Accent 1/Accent 2, 3/4, 5/6")
#: PowerPoint's own default, Medium Style 2 - Accent 1: a new deck's ``def``.
DEFAULT_TABLE_STYLE = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

GUID = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
                  r"[0-9A-Fa-f]{12}\}")
_BUILTIN_BY_KEY = {guid.upper(): guid for guid in BUILTIN_TABLE_STYLES}


def _name_key(name: str) -> str:
    """A style name as compared: case, runs of spaces and the dash's kind ignored."""
    text = re.sub(r"\s*[-–—]\s*", " - ", name.strip())
    return re.sub(r"\s+", " ", text).casefold()


def styles_part(document: "Document") -> str | None:
    """The deck's ``tableStyles.xml``, or ``None`` when it has none."""
    package = document.package
    parts = package.related_parts_of_type(package.presentation_part(), REL_TABLE_STYLES)
    return parts[0] if parts else None


def _styles_root(document: "Document") -> Element | None:
    part = styles_part(document)
    return document.package.tree(part) if part is not None else None


def deck_default(document: "Document") -> str | None:
    """``tableStyles.xml``'s ``def``: the style the deck gives a new table, or ``None``."""
    root = _styles_root(document)
    value = root.get("def") if root is not None else None
    return value or None


def table_styles(document: "Document") -> list[dict]:
    """See :attr:`pptx_agent.edit.document.Document.table_styles`."""
    out: list[dict] = []
    seen: set[str] = set()
    root = _styles_root(document)
    default = deck_default(document)
    for node in findall(root, "a:tblStyle") if root is not None else []:
        style_id = node.get("styleId")
        if not style_id or style_id.upper() in seen:
            continue
        seen.add(style_id.upper())
        builtin = style_id.upper() in _BUILTIN_BY_KEY
        name = node.get("styleName") or (BUILTIN_TABLE_STYLES[_BUILTIN_BY_KEY[style_id.upper()]]
                                         if builtin else "")
        out.append(_entry(style_id, name, builtin, default))
    for style_id, name in BUILTIN_TABLE_STYLES.items():
        if style_id.upper() not in seen:
            out.append(_entry(style_id, name, True, default))
    return out


def _entry(style_id: str, name: str, builtin: bool, default: str | None) -> dict:
    entry = {"id": style_id, "name": name, "builtin": builtin}
    if default is not None and style_id.upper() == default.upper():
        entry["default"] = True
    return entry


def describe(document: "Document", style_id: str | None) -> dict | None:
    """``{"id", "name", "builtin"}`` for a style id (the deck's definition first), or
    ``None`` for no id.  An id neither the deck nor PowerPoint has keeps ``name`` ``None``:
    PowerPoint draws such a table with no style."""
    if not style_id:
        return None
    key = style_id.upper()
    for entry in table_styles(document):
        if entry["id"].upper() == key:
            return {"id": entry["id"], "name": entry["name"], "builtin": entry["builtin"]}
    return {"id": style_id, "name": None, "builtin": False}


def resolve(document: "Document", style: str) -> str:
    """The GUID of ``style``: a style id or a name the deck or PowerPoint has (case, spacing
    and the dash's kind do not matter; the deck's own style wins over a built-in of the
    same name).  ``ValueError`` names what there is."""
    if not isinstance(style, str) or not style.strip():
        raise ValueError(f"a table style is a name or a {{GUID}}, not {style!r}")
    styles = table_styles(document)
    text = style.strip()
    if GUID.fullmatch(text):
        for entry in styles:
            if entry["id"].upper() == text.upper():
                return entry["id"]
        raise ValueError(f"no table style {text} in this deck or built into PowerPoint "
                         "(PowerPoint would draw the table with no style); "
                         "deck.table_styles lists them")
    key = _name_key(text)
    for entry in styles:
        if entry["name"] and _name_key(entry["name"]) == key:
            return entry["id"]
    own = [entry["name"] for entry in styles if not entry["builtin"] and entry["name"]]
    raise ValueError(f"no table style named {text!r}; the deck's own: "
                     f"{', '.join(own) if own else 'none'}; built in: {BUILTIN_NAMES}")


#: Table Design flag -> ``a:tblPr`` attribute.
FLAGS = {"first_row": "firstRow", "last_row": "lastRow", "first_column": "firstCol",
         "last_column": "lastCol", "banded_rows": "bandRow", "banded_columns": "bandCol"}


def read_flags(properties: Element | None) -> dict[str, bool]:
    """The Table Design flags ``a:tblPr`` sets (each off when absent)."""
    return {flag: properties is not None and properties.get(attribute) in ("1", "true")
            for flag, attribute in FLAGS.items()}


def write_flags(properties: Element, changes: dict[str, bool]) -> None:
    """Set flags on ``a:tblPr``: a flag that is on is ``"1"`` and one that is off goes, as
    PowerPoint writes them, the attributes in the schema's order."""
    flags = read_flags(properties)
    flags.update(changes)
    on = {attribute for flag, attribute in FLAGS.items() if flags[flag]}
    others = {name: value for name, value in properties.attrib.items()
              if name not in FLAGS.values()}
    properties.attrib.clear()
    for name in _SCHEMA_ORDER:
        if name in on:
            properties.set(name, "1")
        elif name in others:
            properties.set(name, others.pop(name))
    for name, value in others.items():
        properties.set(name, value)


#: ``CT_TableProperties``' attributes in the schema's order.
_SCHEMA_ORDER = ("rtl", "firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol")


def style_id_of(properties: Element | None) -> str | None:
    """``a:tblPr/a:tableStyleId``'s text, or ``None``."""
    node = find(properties, "a:tableStyleId") if properties is not None else None
    text = (node.text or "").strip() if node is not None else ""
    return text or None


def write_style_id(properties: Element, style_id: str) -> None:
    """Set ``a:tableStyleId``, where the schema puts it; an inline ``a:tableStyle``, its
    alternative, goes."""
    for node in list(properties):
        if node.tag in (qn("a:tableStyle"), qn("a:tableStyleId")):
            remove(node)
    node = make("a:tableStyleId")
    node.text = style_id
    append_in_order(properties, node)


__all__ = ["BUILTIN_TABLE_STYLES", "DEFAULT_TABLE_STYLE", "FLAGS", "resolve", "table_styles"]
