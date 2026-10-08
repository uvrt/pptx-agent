"""Speaker notes: a slide's notes page, and the notes master it needs.

A notes page (``p:notes``, ``ppt/notesSlides/notesSlide<n>.xml``) belongs to one slide and is
drawn from the deck's notes master.  A deck from nothing -- PowerPoint's own new deck, and
:meth:`Document.new` -- has no notes master until a slide gets notes, so writing the first
notes may add one.

Measured on PowerPoint for Mac 16 through its AppleScript: a new presentation, a slide, text
set on the notes page's notes placeholder, saved, and the XML read back.  What PowerPoint
writes, and so what is written here:

* **The notes page**: three placeholders -- the slide image (``sldImg``, no ``idx``, locked
  against grouping, rotation and aspect changes), the notes (``body`` ``idx="1"``, whatever
  ``idx`` the master's notes placeholder has; it matches by type) and the slide number
  (``sldNum`` ``sz="quarter"``, the master's ``idx``, a ``slidenum`` field) -- none with a
  frame of its own; relationships to the notes master and back to the slide, and from the
  slide to it.
* **The notes master**, when the deck has none: a background of ``bg1`` (``bgRef idx=1001``);
  header, date, slide image, notes, footer and slide number placeholders on the 6,858,000 x
  9,144,000 notes page (header and date 2,971,800 x 458,788 along the top, the slide image
  at (685,800, 1,143,000) 5,486,400 x 3,086,100 with a 1 pt black outline, the notes at
  (685,800, 4,400,550) 5,486,400 x 3,600,450, footer and number 2,971,800 x 458,787 along
  the bottom; 12 pt, the right-hand ones right-aligned, the bottom ones anchored at the
  bottom); the notes placeholder prompting five levels; the slide master's colour map; a
  ``notesStyle`` of nine 12 pt levels, ``marL`` 457,200 a level; listed in
  ``p:notesMasterIdLst``; and **a theme of its own**, a copy of the slide master's (byte for
  byte the same part in PowerPoint's deck).  The slide image is the slide's shape fitted
  into that 5,486,400 x 3,086,100 box (the 16:9 box exactly).

The prompts are PowerPoint's English ones, as for the slide master (:mod:`.blank`).
"""

from __future__ import annotations

from ..oxml.package import REL_NOTES_MASTER, REL_NOTES_SLIDE, REL_SLIDE, REL_SLIDE_MASTER
from ..oxml.xml import Element, find, local_name, qn, serialize
from . import blank

CT_NOTES_SLIDE = blank.CT + "presentationml.notesSlide+xml"
CT_NOTES_MASTER = blank.CT + "presentationml.notesMaster+xml"
REL_THEME = blank.REL_THEME

#: The notes master's placeholders on the notes page (measured): name, type, idx, sz,
#: frame, alignment, anchor.
_MASTER_PLACEHOLDERS = (
    ("Header Placeholder", "hdr", None, "quarter", (0, 0, 2971800, 458788), "l", None),
    ("Date Placeholder", "dt", 1, None, (3884613, 0, 2971800, 458788), "r", None),
    ("Slide Image Placeholder", "sldImg", 2, None, (685800, 1143000, 5486400, 3086100), None,
     "ctr"),
    ("Notes Placeholder", "body", 3, "quarter", (685800, 4400550, 5486400, 3600450), None,
     None),
    ("Footer Placeholder", "ftr", 4, "quarter", (0, 8685213, 2971800, 458787), "l", "b"),
    ("Slide Number Placeholder", "sldNum", 5, "quarter", (3884613, 8685213, 2971800, 458787),
     "r", "b"),
)
_NOTES_SIZE = 1200
_INSETS = {"lIns": "91440", "tIns": "45720", "rIns": "91440", "bIns": "45720"}
_COLOR_MAP = (("bg1", "lt1"), ("tx1", "dk1"), ("bg2", "lt2"), ("tx2", "dk2"),
              ("accent1", "accent1"), ("accent2", "accent2"), ("accent3", "accent3"),
              ("accent4", "accent4"), ("accent5", "accent5"), ("accent6", "accent6"),
              ("hlink", "hlink"), ("folHlink", "folHlink"))


# -- reading ---------------------------------------------------------------------------------


def notes_part(package, slide_part: str) -> str | None:
    """The slide's notes page, or ``None``."""
    related = package.related_parts_of_type(slide_part, REL_NOTES_SLIDE)
    return related[0] if related and package.has_part(related[0]) else None


def notes_body(package, notes: str) -> Element | None:
    """The notes page's notes placeholder (``p:sp`` with ``ph type="body"``)."""
    root = package.tree(notes)
    tree = find(root, "p:cSld/p:spTree") if root is not None else None
    for shape in [] if tree is None else tree:
        if local_name(shape) != "sp":
            continue
        ph = find(shape, "p:nvSpPr/p:nvPr/p:ph")
        if ph is not None and ph.get("type") == "body":
            return shape
    return None


def notes_master(package) -> str | None:
    presentation = package.presentation_part()
    related = package.related_parts_of_type(presentation, REL_NOTES_MASTER)
    return related[0] if related and package.has_part(related[0]) else None


def master_placeholder(package, master: str, kind: str) -> Element | None:
    root = package.tree(master)
    tree = find(root, "p:cSld/p:spTree") if root is not None else None
    for shape in [] if tree is None else tree:
        ph = shape.find(f".//{qn('p:ph')}")
        if ph is not None and (ph.get("type") or "obj") == kind:
            return shape
    return None


# -- writing ---------------------------------------------------------------------------------


def _sub(parent: Element, tag: str, **attributes: str) -> Element:
    return blank._sub(parent, tag, **attributes)


def _ensure_master(package, language: str, slide_size: tuple[int, int]) -> str:
    """The deck's notes master, added (with a theme of its own) when it has none."""
    existing = notes_master(package)
    if existing is not None:
        return existing
    presentation = package.presentation_part()
    masters = package.related_parts_of_type(presentation, REL_SLIDE_MASTER)
    if not masters:
        raise ValueError("the presentation has no slide master to take a theme from")
    master_themes = package.related_parts_of_type(masters[0], REL_THEME)
    theme_bytes = package.read(master_themes[0]) if master_themes else blank.theme_xml()
    theme = package.unused_part_name("ppt/theme/theme{n}.xml")
    package.add_part(theme, theme_bytes, blank.CT_THEME, override=True)

    color_map = find(package.tree(masters[0]), "p:clrMap")
    part = package.unused_part_name("ppt/notesMasters/notesMaster{n}.xml")
    package.add_part(part, _master_xml(language, slide_size, color_map), CT_NOTES_MASTER,
                     override=True)
    package.add_relationship(part, REL_THEME, theme)

    rel_id = package.add_relationship(presentation, REL_NOTES_MASTER, part)
    root = package.tree(presentation)
    from ..oxml.xml import append_in_order, make

    entries = find(root, "p:notesMasterIdLst")
    if entries is None:
        entries = make("p:notesMasterIdLst")
        append_in_order(root, entries)
    entries.append(make("p:notesMasterId", r__id=rel_id))
    package.mark_dirty(presentation)
    return part


def _slide_image(slide_size: tuple[int, int]) -> tuple[int, int, int, int]:
    """The slide image's frame: the slide's shape fitted into PowerPoint's 16:9 box."""
    x, y, width, height = _MASTER_PLACEHOLDERS[2][4]
    ratio = slide_size[0] / slide_size[1]
    if ratio >= width / height:
        fitted = round(width / ratio)
        return x, y + (height - fitted) // 2, width, fitted
    fitted = round(height * ratio)
    return x + (width - fitted) // 2, y, fitted, height


def _master_xml(language: str, slide_size: tuple[int, int], color_map: Element | None) -> bytes:
    root = blank._root("p:notesMaster", ("a", "r", "p"))
    common = _sub(root, "p:cSld")
    background = _sub(_sub(common, "p:bg"), "p:bgRef", idx="1001")
    _sub(background, "a:schemeClr", val="bg1")
    tree = blank._shape_tree(common)
    for number, (name, kind, idx, size, frame, align, anchor) in enumerate(_MASTER_PLACEHOLDERS):
        identifier = number + 2
        if kind == "sldImg":
            frame = _slide_image(slide_size)
        shape = _sub(tree, "p:sp")
        nv = _sub(shape, "p:nvSpPr")
        _sub(nv, "p:cNvPr", id=str(identifier), name=f"{name} {identifier - 1}")
        locks = _sub(_sub(nv, "p:cNvSpPr"), "a:spLocks", noGrp="1")
        if kind == "sldImg":
            locks.set("noRot", "1")
            locks.set("noChangeAspect", "1")
        ph = _sub(_sub(nv, "p:nvPr"), "p:ph", type=kind)
        if size is not None:
            ph.set("sz", size)
        if idx is not None:
            ph.set("idx", str(idx))
        properties = _sub(shape, "p:spPr")
        xfrm = _sub(properties, "a:xfrm")
        _sub(xfrm, "a:off", x=str(frame[0]), y=str(frame[1]))
        _sub(xfrm, "a:ext", cx=str(frame[2]), cy=str(frame[3]))
        _sub(_sub(properties, "a:prstGeom", prst="rect"), "a:avLst")
        if kind == "sldImg":
            _sub(properties, "a:noFill")
            outline = _sub(properties, "a:ln", w="12700")
            _sub(_sub(outline, "a:solidFill"), "a:prstClr", val="black")
        body = _sub(shape, "p:txBody")
        body_properties = _sub(body, "a:bodyPr", vert="horz", **_INSETS, rtlCol="0")
        if anchor is not None:
            body_properties.set("anchor", anchor)
        styles = _sub(body, "a:lstStyle")
        if align is not None:
            level = _sub(styles, "a:lvl1pPr", algn=align)
            _sub(level, "a:defRPr", sz=str(_NOTES_SIZE))
        if kind == "body":
            for level, prompt in enumerate(blank.TEXT_PROMPTS):
                paragraph = _sub(body, "a:p")
                _sub(paragraph, "a:pPr", lvl=str(level))
                run = _sub(paragraph, "a:r")
                _sub(run, "a:rPr", lang=language)
                _sub(run, "a:t").text = prompt
            _sub(paragraph, "a:endParaRPr", lang=language)
            continue
        paragraph = _sub(body, "a:p")
        if kind in ("dt", "sldNum"):
            field = _sub(paragraph, "a:fld",
                         id=blank.DATE_FIELD_ID if kind == "dt" else blank.NUMBER_FIELD_ID,
                         type="datetimeFigureOut" if kind == "dt" else "slidenum")
            _sub(field, "a:rPr", lang=language)
            _sub(field, "a:t").text = "‹#›" if kind == "sldNum" else ""
        _sub(paragraph, "a:endParaRPr", lang=language)
    mapping = _sub(root, "p:clrMap")
    for name, value in _COLOR_MAP:
        mapping.set(name, color_map.get(name, value) if color_map is not None else value)
    styles = _sub(root, "p:notesStyle")
    scale = blank._Scale(*blank.BASE_SIZE)
    for level in range(1, 10):
        blank._level(styles, level, marL=(level - 1) * blank.LEVEL_STEP, indent=None,
                     size=_NOTES_SIZE, tab=blank.TAB_SIZE, scale=scale)
    return serialize(root)


def _notes_xml(language: str, number_idx: str | None, slide_number: int,
               body_idx: str = "1") -> Element:
    root = blank._root("p:notes", ("a", "r", "p"))
    tree = blank._shape_tree(_sub(root, "p:cSld"))
    specs = [("Slide Image Placeholder", "sldImg", None, None)]
    specs.append(("Notes Placeholder", "body", body_idx, None))
    if number_idx is not None:
        specs.append(("Slide Number Placeholder", "sldNum", number_idx, "quarter"))
    for number, (name, kind, idx, size) in enumerate(specs):
        identifier = number + 2
        shape = _sub(tree, "p:sp")
        nv = _sub(shape, "p:nvSpPr")
        _sub(nv, "p:cNvPr", id=str(identifier), name=f"{name} {identifier - 1}")
        locks = _sub(_sub(nv, "p:cNvSpPr"), "a:spLocks", noGrp="1")
        if kind == "sldImg":
            locks.set("noRot", "1")
            locks.set("noChangeAspect", "1")
        ph = _sub(_sub(nv, "p:nvPr"), "p:ph", type=kind)
        if size is not None:
            ph.set("sz", size)
        if idx is not None:
            ph.set("idx", idx)
        _sub(shape, "p:spPr")
        if kind == "sldImg":
            continue
        body = _sub(shape, "p:txBody")
        _sub(body, "a:bodyPr")
        _sub(body, "a:lstStyle")
        paragraph = _sub(body, "a:p")
        if kind == "sldNum":
            field = _sub(paragraph, "a:fld", id=blank.NUMBER_FIELD_ID, type="slidenum")
            _sub(field, "a:rPr", lang=language)
            _sub(field, "a:t").text = str(slide_number)
        _sub(paragraph, "a:endParaRPr", lang=language)
    _sub(_sub(root, "p:clrMapOvr"), "a:masterClrMapping")
    return root


def add_notes(document, slide_part: str, slide_number: int) -> Element:
    """Give the slide a notes page (and the deck a notes master if it has none); returns
    the page's notes placeholder, holding one empty paragraph.  Inside the caller's batch."""
    package = document.package
    existing = notes_part(package, slide_part)
    if existing is not None:
        body = notes_body(package, existing)
        if body is None:
            raise ValueError(f"{existing} has no notes placeholder")
        return body
    language = _language(document)
    master = _ensure_master(package, language, document.slide_size)
    number = master_placeholder(package, master, "sldNum")
    number_ph = number.find(f".//{qn('p:ph')}") if number is not None else None
    number_idx = None if number_ph is None else (number_ph.get("idx") or "0")
    root = _notes_xml(language, number_idx, slide_number)
    part = package.unused_part_name("ppt/notesSlides/notesSlide{n}.xml")
    package.add_part(part, serialize(root), CT_NOTES_SLIDE, override=True)
    package.add_relationship(part, REL_NOTES_MASTER, master)
    package.add_relationship(part, REL_SLIDE, slide_part)
    package.add_relationship(slide_part, REL_NOTES_SLIDE, part)
    body = notes_body(package, part)
    assert body is not None
    return body


def _language(document) -> str:
    from .authoring import default_language

    return default_language(document.package)


__all__ = ["CT_NOTES_MASTER", "CT_NOTES_SLIDE", "add_notes", "master_placeholder",
           "notes_body", "notes_master", "notes_part"]
