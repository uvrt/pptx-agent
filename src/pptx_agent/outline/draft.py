"""``insert_outline``: slides drafted from Markdown (see :mod:`pptx_agent.outline`)."""

from __future__ import annotations

import os
import re
import warnings
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Callable, Mapping

from ..oxml.package import REL_HYPERLINK, image_size
from ..oxml.xml import Element, find, get_int, local_name, make, qn, remove
from .markdown import comment_text, normalise_inline, parse, plain
from .read import CODE_FONT, NOTES_MARKER, TITLE_TYPES, _bullet, notes_chain, style_chain

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, Slide
    from ..edit.slides import Layout


class OutlineWarning(UserWarning):
    """Something in an outline the drafting half leaves out (a chart, a SmartArt diagram)."""


#: The layout roles the content rules choose, as ``ST_SlideLayoutType`` values, and the
#: English names PowerPoint gives them -- either may key ``layout_map``.
ROLES = {
    "title": "Title Slide",
    "secHead": "Section Header",
    "obj": "Title and Content",
    "twoObj": "Two Content",
    "titleOnly": "Title Only",
    "blank": "Blank",
}
#: Layout types that stand in for a role a template does not have (Google Slides' and
#: older PowerPoint's), before falling back to the placeholders a layout has.
_SIMILAR = {
    "title": ("title",), "secHead": ("secHead",), "obj": ("obj", "tx"),
    "twoObj": ("twoObj", "twoColTx"), "titleOnly": ("titleOnly",), "blank": ("blank",),
}
_TEXT_TYPES = (None, "obj", "body", "subTitle")
_ID = re.compile(r"^\d+\.\S+$")
_CHART = re.compile(r"^\[chart: .*\]$", re.DOTALL)
_SCHEME = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")
#: PowerPoint's default row height for a new table's 18 pt text.
ROW_HEIGHT = 370840
#: The smallest height a free text box is given in the content area (one line of 18 pt).
_LINE = 369332
_BULLET_MARGIN, _BULLET_INDENT = 285750, -285750
_NUMBER_MARGIN, _NUMBER_INDENT = 514350, -514350
_LEVEL_STEP = 457200


# -- the parsed outline ----------------------------------------------------------------------


@dataclass
class Hint:
    """What a shape comment says: ``ph: <type> <idx>``, a kind, ``empty``.  A comment that
    says only an id (a free text shape's) is a hint with no placeholder."""

    kind: str = "text"  # text, table, picture, chart, smartart
    placeholder: tuple | None = None  # (type, idx)
    empty: bool = False


@dataclass
class Item:
    kind: str  # text, table, picture
    blocks: list = field(default_factory=list)
    hint: Hint | None = None
    image: tuple | None = None  # (alt, src, title)
    table: tuple | None = None
    data: bytes | None = None


@dataclass
class SlideDraft:
    heading: tuple = ()
    title_named: bool = False
    layout: str | None = None
    hidden: bool = False
    items: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    skipped: list = field(default_factory=list)


def _slide_meta(text: str) -> dict | None:
    """``s:256 hidden layout: Name`` (any part optional), or ``None`` for another comment."""
    words = text.split()
    if words and re.fullmatch(r"s:\d+", words[0]):
        words = words[1:]
    meta: dict = {"hidden": False, "layout": None}
    while words and words[0] == "hidden":
        meta["hidden"] = True
        words = words[1:]
    if words and words[0] == "layout:":
        meta["layout"] = " ".join(words[1:]) or None
        words = []
    if words:
        return None
    return meta


def _hint(text: str) -> Hint | None:
    """A shape comment's hint, or ``None`` for a comment that is not one."""
    words = text.split()
    if words and _ID.match(words[0]):
        words = words[1:]
    hint = Hint()
    k = 0
    while k < len(words):
        word = words[k]
        if word in ("table", "picture", "chart", "smartart"):
            hint.kind = word
        elif word == "empty":
            hint.empty = True
        elif word == "ph:" and k + 1 < len(words):
            kind = words[k + 1]
            idx = None
            if k + 2 < len(words) and words[k + 2].isdigit():
                idx = int(words[k + 2])
                k += 1
            hint.placeholder = (None if kind == "obj" else kind, idx)
            k += 1
        else:
            return None
        k += 1
    return hint


def _strip_title_id(inline: tuple) -> tuple[tuple, bool]:
    named = False
    kept = []
    for item in inline:
        if item[0] == "html":
            if re.fullmatch(r"<!--\s*\d+\.\S+\s*-->", item[1].strip()):
                named = True
            continue
        kept.append(item)
    return normalise_inline(kept), named


def read_outline(text: str) -> list[SlideDraft]:
    """The outline's slides: ``#`` starts each one."""
    blocks = list(parse(text))
    slides: list[SlideDraft] = []
    pending_meta: dict | None = None
    current: SlideDraft | None = None
    unit: list | None = None
    in_notes = False
    units: list[tuple[Hint | None, list]] = []

    def finish() -> None:
        if current is not None:
            current.items = _items(units, current)

    k = 0
    while k < len(blocks):
        block = blocks[k]
        kind = block[0]
        nxt = blocks[k + 1] if k + 1 < len(blocks) else None
        if kind == "comment" and nxt is not None and nxt[0] == "h" and nxt[1] == 1:
            meta = _slide_meta(block[1])
            if meta is not None:
                pending_meta = meta
                k += 1
                continue
        if kind == "h" and block[1] == 1:
            finish()
            current = SlideDraft()
            current.heading, current.title_named = _strip_title_id(block[2])
            if pending_meta:
                current.layout = pending_meta["layout"]
                current.hidden = pending_meta["hidden"]
            pending_meta = None
            slides.append(current)
            units = []
            unit = None
            in_notes = False
            k += 1
            continue
        pending_meta = None
        if current is None:
            if kind in ("comment", "html"):
                k += 1
                continue
            raise ValueError("an outline starts with a '#' heading: each slide is one")
        if in_notes:
            if kind not in ("comment", "html", "hr"):
                current.notes.append(block)
            k += 1
            continue
        if kind == "p" and block[2].casefold() == NOTES_MARKER.casefold():
            in_notes = True
            rest = _after_marker(block[1])
            if rest:
                current.notes.append(("p", rest, ""))
            k += 1
            continue
        if kind == "comment":
            hint = _hint(block[1])
            if hint is not None:
                unit = []
                units.append((hint, unit))
            k += 1
            continue
        if kind == "hr":
            unit = None
            k += 1
            continue
        if kind == "html":
            k += 1
            continue
        if unit is None:
            unit = []
            units.append((None, unit))
        unit.append(block)
        k += 1
    finish()
    return slides


def _after_marker(inline: tuple) -> tuple:
    """The paragraph after its ``Notes:`` line."""
    items = list(inline)
    if items and items[0][0] == "t":
        text = items[0][1]
        head = text[: len(NOTES_MARKER)]
        if head.casefold() == NOTES_MARKER.casefold():
            items[0] = ("t", text[len(NOTES_MARKER):]) + items[0][2:]
    return normalise_inline(items)


def _items(units: list, slide: SlideDraft) -> list[Item]:
    items: list[Item] = []
    for hint, blocks in units:
        if hint is not None and hint.kind in ("chart", "smartart"):
            slide.skipped.append(hint.kind)
            continue
        found: list[Item] = []
        text: list = []
        for block in blocks:
            if block[0] == "table":
                found.append(Item("table", table=block))
            elif block[0] == "p":
                images = [i for i in block[1] if i[0] == "img"]
                rest = normalise_inline([i for i in block[1] if i[0] not in ("img", "html")])
                if not images and _CHART.match(plain(rest)):
                    slide.skipped.append("chart")
                    continue
                if rest:
                    text.append(("p", rest, block[2]))
                for image in images:
                    found.append(Item("picture", image=image[1:]))
            else:
                text.append(block)
        if text:
            found.insert(0, Item("text", blocks=text))
        if hint is not None:
            if found:
                wanted = {"table": "table", "picture": "picture"}.get(hint.kind, "text")
                target = next((i for i in found if i.kind == wanted), found[0])
                target.hint = hint
            elif hint.placeholder is not None or hint.empty:
                found.append(Item("text", hint=hint))
        items += found
    return items


# -- layouts ---------------------------------------------------------------------------------


def _layout_type(document: "Document", layout: "Layout") -> str | None:
    root = document.package.tree(layout.part_path)
    return root.get("type") if root is not None else None


def _layout_placeholders(document: "Document", layout: "Layout") -> list[tuple]:
    root = document.package.tree(layout.part_path)
    tree = find(root, "p:cSld/p:spTree") if root is not None else None
    found = []
    for shape in [] if tree is None else tree:
        ph = shape.find(f".//{qn('p:ph')}")
        if ph is not None and ph.get("type") not in ("dt", "ftr", "sldNum", "hdr"):
            found.append((ph.get("type"), get_int(ph, "idx")))
    return found


def _shape_of(document: "Document", layout: "Layout") -> tuple[bool, int, bool]:
    """(has a title, how many text placeholders, has other placeholders)."""
    placeholders = _layout_placeholders(document, layout)
    title = any(t in TITLE_TYPES for t, _ in placeholders)
    text = sum(1 for t, _ in placeholders if t in _TEXT_TYPES)
    other = any(t not in _TEXT_TYPES and t not in TITLE_TYPES for t, _ in placeholders)
    return title, text, other


def find_layout(document: "Document", role: str) -> "Layout":
    """The deck's layout for a role: by ``type`` (so a localised template works), then by
    PowerPoint's English name, then by the placeholders it has."""
    layouts = document.layouts
    if not layouts:
        raise ValueError("the presentation has no slide layouts")
    first_master = layouts[0].master_part
    candidates = [l for l in layouts if l.master_part == first_master] + \
        [l for l in layouts if l.master_part != first_master]
    for kind in _SIMILAR.get(role, (role,)):
        for layout in candidates:
            if _layout_type(document, layout) == kind:
                return layout
    english = ROLES.get(role)
    for layout in candidates:
        if english is not None and layout.name.casefold() == english.casefold():
            return layout
    wanted = {"title": (True, 1), "secHead": (True, 1), "obj": (True, 1), "twoObj": (True, 2),
              "titleOnly": (True, 0), "blank": (False, 0)}.get(role)
    if wanted is not None:
        for layout in candidates:
            title, text, other = _shape_of(document, layout)
            if (title, text) == wanted and not other:
                return layout
    fallback = {"title": "secHead", "secHead": "titleOnly", "twoObj": "obj",
                "titleOnly": "obj", "blank": "titleOnly"}.get(role)
    if fallback is not None:
        return find_layout(document, fallback)
    return candidates[0]


def _layout_by_name(document: "Document", name: str) -> "Layout | None":
    for layout in document.layouts:
        if comment_text(layout.name) == name or layout.name == name:
            return layout
    return None


def choose_role(slide: SlideDraft, first: bool) -> str:
    """The content rules (see :mod:`pptx_agent.outline`)."""
    texts = [i for i in slide.items if i.kind == "text" and i.blocks]
    objects = [i for i in slide.items if i.kind in ("table", "picture")]
    has_title = bool(plain(slide.heading).strip())
    if not texts and not objects:
        if not has_title:
            return "blank"
        return "title" if first else "secHead"
    if len(texts) == 1 and not objects:
        lists = any(b[0] == "list" for b in texts[0].blocks)
        return "title" if first and not lists else "obj"
    if len(texts) == 2 and not objects:
        return "twoObj"
    if not texts:
        return "titleOnly"
    if len(texts) == 1 and len(objects) == 1:
        return "twoObj"
    return "obj"


def _split_two_lists(slide: SlideDraft) -> None:
    """One unhinted text unit holding exactly two lists is two columns: the first list (and
    what precedes it) and the rest."""
    texts = [i for i in slide.items if i.kind == "text" and i.blocks]
    if len(texts) != 1 or texts[0].hint is not None or any(i.hint for i in slide.items):
        return
    blocks = texts[0].blocks
    lists = [k for k, b in enumerate(blocks) if b[0] == "list"]
    if len(lists) != 2:
        return
    cut = lists[0] + 1
    if cut >= len(blocks):
        return
    position = slide.items.index(texts[0])
    slide.items[position:position + 1] = [Item("text", blocks=blocks[:cut]),
                                          Item("text", blocks=blocks[cut:])]


def _resolve_layout(document: "Document", slide: SlideDraft, first: bool,
                    layout_map: Mapping | None) -> "Layout":
    explicit = _layout_by_name(document, slide.layout) if slide.layout else None
    if explicit is not None:
        chosen, role = explicit, None
    else:
        _split_two_lists(slide)
        role = choose_role(slide, first)
        chosen = None
    if layout_map:
        keys = ([role, ROLES.get(role)] if role else []) + \
            ([explicit.name] if explicit is not None else [])
        for key in keys:
            if key is not None and key in layout_map:
                return _layout_value(document, layout_map[key])
        if role is not None:
            for key, value in layout_map.items():
                if isinstance(key, str) and ROLES.get(role, "").casefold() == key.casefold():
                    return _layout_value(document, value)
    if chosen is not None:
        return chosen
    return find_layout(document, role)


def _layout_value(document: "Document", value: Any) -> "Layout":
    if hasattr(value, "part_path"):
        return value
    if isinstance(value, str):
        for layout in document.layouts:
            if value in (layout.name, layout.part_path):
                return layout
        if value in ROLES or value in {t for kinds in _SIMILAR.values() for t in kinds}:
            return find_layout(document, value)
    raise KeyError(f"layout_map names no layout of this deck: {value!r}")


# -- images ----------------------------------------------------------------------------------


def _image_resolver(images) -> Callable[[str], bytes | None]:
    if images is None:
        return lambda src: _read_file(src, None)
    if callable(images) and not hasattr(images, "package"):
        return images
    if hasattr(images, "package"):
        package = images.package

        def from_deck(src: str) -> bytes | None:
            name = src.lstrip("/")
            return package.read(name) if package.has_part(name) else None

        return from_deck
    if isinstance(images, Mapping):
        def from_mapping(src: str) -> bytes | None:
            value = images.get(src)
            if value is None:
                return None
            return value if isinstance(value, bytes) else _read_file(os.fspath(value), None)

        return from_mapping
    base = os.fspath(images)
    return lambda src: _read_file(src, base)


def _read_file(src: str, base: str | None) -> bytes | None:
    if not src or _SCHEME.match(src) and not re.match(r"^[A-Za-z]:[\\/]", src):
        return None
    path = src if base is None or os.path.isabs(src) else os.path.join(base, src)
    try:
        with open(path, "rb") as handle:
            return handle.read()
    except OSError:
        return None


# -- writing text ----------------------------------------------------------------------------


class _Writer:
    """Paragraphs from outline blocks, into one text body of one part."""

    def __init__(self, document: "Document", part: str, chain: list, language: str) -> None:
        self.document = document
        self.part = part
        self.chain = chain
        self.language = language

    def paragraphs(self, blocks: list) -> list[Element]:
        out: list[Element] = []
        for block in blocks:
            out += self._block(block, 0)
        return out

    def _block(self, block: tuple, level: int) -> list[Element]:
        kind = block[0]
        if kind == "p":
            return [self.paragraph("p", 0, 1, block[1])]
        if kind == "h":
            return [self.paragraph("p", 0, 1, block[2])]
        if kind == "quote":
            return [p for inner in block[1] for p in self._block(inner, level)]
        if kind == "code":
            return [self.paragraph("p", 0, 1, (("t", line, frozenset({"code"}), None),))
                    for line in block[1].rstrip("\n").split("\n")]
        if kind == "list":
            return self._list(block, level)
        return []

    def _list(self, block: tuple, level: int) -> list[Element]:
        _, ordered, start, items = block
        out = []
        level = min(level, 8)
        for number, item in enumerate(items):
            first = True
            for inner in item:
                if inner[0] == "list":
                    out += self._list(inner, level + 1)
                elif inner[0] in ("p", "h") and first:
                    inline = inner[1] if inner[0] == "p" else inner[2]
                    out.append(self.paragraph("number" if ordered else "bullet", level,
                                              start if number == 0 else 1, inline,
                                              restart=number == 0 and start != 1))
                    first = False
                else:
                    out += self._block(inner, level)
            if first:
                out.append(self.paragraph("number" if ordered else "bullet", level, 1, ()))
        return out

    def paragraph(self, kind: str, level: int, start: int, inline: tuple, *,
                  restart: bool = False) -> Element:
        paragraph = make("a:p")
        inherited, _ = _bullet(None, level, self.chain)
        properties = None
        if kind == "p":
            if inherited != "p":
                properties = make("a:pPr", marL="0", indent="0")
                properties.append(make("a:buNone"))
        elif kind == inherited and not restart:
            if level:
                properties = make("a:pPr", lvl=str(level))
        else:
            margin, indent = (_BULLET_MARGIN, _BULLET_INDENT) if kind == "bullet" else \
                (_NUMBER_MARGIN, _NUMBER_INDENT)
            properties = make("a:pPr")
            if kind != inherited:
                properties.set("marL", str(margin + level * _LEVEL_STEP))
                properties.set("indent", str(indent))
            if level:
                properties.set("lvl", str(level))
            if kind == "bullet":
                properties.append(make("a:buFont", typeface="Arial", panose="020B0604020202020204",
                                       pitchFamily="34", charset="0"))
                properties.append(make("a:buChar", char="•"))
            else:
                properties.append(make("a:buFont", typeface="+mj-lt"))
                numbering = make("a:buAutoNum", type="arabicPeriod")
                if start != 1:
                    numbering.set("startAt", str(start))
                properties.append(numbering)
        if properties is not None:
            paragraph.append(properties)
        runs = self.runs(inline)
        for run in runs:
            paragraph.append(run)
        if not runs:
            paragraph.append(make("a:endParaRPr", lang=self.language))
        return paragraph

    def runs(self, inline: tuple) -> list[Element]:
        out = []
        for item in normalise_inline([i for i in inline if i[0] in ("t", "br")]):
            if item[0] == "br":
                br = make("a:br")
                br.append(make("a:rPr", lang=self.language))
                out.append(br)
                continue
            _, text, marks, href = item
            run = make("a:r")
            properties = make("a:rPr", lang=self.language)
            if "strong" in marks:
                properties.set("b", "1")
            if "em" in marks:
                properties.set("i", "1")
            if "strike" in marks:
                properties.set("strike", "sngStrike")
            if "code" in marks:
                properties.append(make("a:latin", typeface=CODE_FONT))
            if href and _SCHEME.match(href):
                rel_id = self.document.package.add_external_relationship(
                    self.part, REL_HYPERLINK, href)
                properties.append(make("a:hlinkClick", r__id=rel_id))
            run.append(properties)
            node = make("a:t")
            node.text = text
            run.append(node)
            out.append(run)
        return out


def _fill_body(body: Element, paragraphs: list[Element], language: str) -> None:
    for old in body.findall(qn("a:p")):
        remove(old)
    if not paragraphs:
        paragraph = make("a:p")
        paragraph.append(make("a:endParaRPr", lang=language))
        paragraphs = [paragraph]
    for paragraph in paragraphs:
        body.append(paragraph)


# -- geometry --------------------------------------------------------------------------------


def _frame(xfrm: Element | None) -> tuple[int, int, int, int] | None:
    if xfrm is None:
        return None
    off, ext = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return (get_int(off, "x", 0), get_int(off, "y", 0), get_int(ext, "cx", 0),
            get_int(ext, "cy", 0))


def _placeholder_frame(slide: "Slide", placeholder: tuple) -> tuple[int, int, int, int] | None:
    return _frame(slide._inherited_xfrm(placeholder))


def _master_frame(document: "Document", slide: "Slide", kind: str) -> tuple | None:
    package = document.package
    layout = slide.layout
    if layout is None:
        return None
    master = package.tree(layout.master_part)
    tree = find(master, "p:cSld/p:spTree") if master is not None else None
    for shape in [] if tree is None else tree:
        ph = shape.find(f".//{qn('p:ph')}")
        if ph is not None and (ph.get("type") or "obj") == kind:
            return _frame(find(shape, "p:spPr/a:xfrm"))
    return None


def _content_area(document: "Document", slide: "Slide") -> tuple[int, int, int, int]:
    found = _master_frame(document, slide, "body")
    if found is not None and found[2] > 0 and found[3] > 0:
        return found
    width, height = document.slide_size
    return (round(width * 0.07), round(height * 0.27), round(width * 0.86), round(height * 0.63))


def _title_area(document: "Document", slide: "Slide") -> tuple[int, int, int, int]:
    found = _master_frame(document, slide, "title")
    if found is not None and found[2] > 0 and found[3] > 0:
        return found
    width, height = document.slide_size
    return (round(width * 0.07), round(height * 0.05), round(width * 0.86), round(height * 0.19))


def _below(document: "Document", slide: "Slide", area: tuple[int, int, int, int],
           heading: Element | None) -> tuple[int, int, int, int]:
    """``area`` starting below the title's actual bounds, never over it.

    The master's body is where free content goes, but a layout may put its title lower
    than the master does -- the trial's template's ``TITLE_ONLY`` title sits inside the
    master's body area, and a drafted table covered it.  The content then starts the
    master's own title-to-body gap (at least 0.1 in) below the title, and keeps its bottom
    edge -- or, when that leaves less than two lines, takes two lines, to 0.1 in above the
    slide's bottom at most."""
    if heading is None:
        return area
    shape = slide._wrap(heading)
    if shape.top is None or shape.height is None:
        return area
    title_bottom = shape.top + shape.height
    x, y, width, height = area
    if title_bottom <= y:
        return area
    master_title = _master_frame(document, slide, "title")
    gap = 91440
    if master_title is not None and y - (master_title[1] + master_title[3]) > gap:
        gap = y - (master_title[1] + master_title[3])
    top = title_bottom + gap
    bottom = y + height
    if bottom - top < 2 * _LINE:
        bottom = min(document.slide_size[1] - 91440, top + 2 * _LINE)
    return (x, top, width, max(1, bottom - top))


def _grid(area: tuple[int, int, int, int], count: int) -> list[tuple[int, int, int, int]]:
    """``count`` cells filling ``area``: as many rows as a line of text allows, then
    columns."""
    x, y, width, height = area
    if count <= 0:
        return []
    rows_max = max(1, height // _LINE)
    columns = -(-count // rows_max)
    rows = -(-count // columns)
    cell_w, cell_h = width // columns, height // rows
    return [(x + (k // rows) * cell_w, y + (k % rows) * cell_h, cell_w, cell_h)
            for k in range(count)]


def _fit(frame: tuple[int, int, int, int], pixels: tuple[int, int] | None) -> tuple:
    x, y, width, height = frame
    if not pixels or not all(pixels):
        return frame
    ratio = pixels[0] / pixels[1]
    if width / max(height, 1) > ratio:
        fitted = round(height * ratio)
        return (x + (width - fitted) // 2, y, fitted, height)
    fitted = round(width / ratio)
    return (x, y + (height - fitted) // 2, width, fitted)


def _crop(frame: tuple[int, int, int, int], pixels: tuple[int, int] | None) -> dict[str, str]:
    """``a:srcRect`` cropping an image to fill ``frame`` (a picture placeholder's way)."""
    _, _, width, height = frame
    if not pixels or not all(pixels) or not width or not height:
        return {}
    image, box = pixels[0] / pixels[1], width / height
    if abs(image - box) < 1e-6:
        return {}
    if image > box:
        cut = round((1 - box / image) * 100000 / 2)
        return {"l": str(cut), "r": str(cut)}
    cut = round((1 - image / box) * 100000 / 2)
    return {"t": str(cut), "b": str(cut)}


# -- drafting --------------------------------------------------------------------------------


def insert_outline(document: "Document", text: str, *, at: int | None = None,
                   layout_map: Mapping | None = None, images=None) -> list["Slide"]:
    """See :meth:`pptx_agent.edit.document.Document.insert_outline`."""
    drafts = read_outline(text)
    count = len(document.slides)
    position = count if at is None else at
    if not 0 <= position <= count:
        raise IndexError(f"slide index {at} out of range 0..{count}")
    resolve = _image_resolver(images)
    for draft in drafts:
        for item in draft.items:
            if item.kind == "picture":
                alt, src, _ = item.image
                data = resolve(src)
                if not data:
                    raise ValueError(f"cannot read the image {src!r}; pass images= a directory, "
                                     "a mapping, a callable or the deck it came from")
                item.data = data
    layouts = [_resolve_layout(document, draft, position + k == 0, layout_map)
               for k, draft in enumerate(drafts)]
    for draft in drafts:
        for kind in draft.skipped:
            warnings.warn(OutlineWarning(f"a {kind} in the outline is not drafted"), stacklevel=3)

    from ..edit.authoring import default_language

    language = default_language(document.package)
    made = []
    with document.batch():
        document.history.checkpoint()
        for k, (draft, layout) in enumerate(zip(drafts, layouts)):
            slide = document.add_slide(layout, index=position + k)
            _build(document, slide, draft, language, position + k)
            made.append(slide.slide_id)
        document._reset_caches()
    return [document.slide(slide_id) for slide_id in made]


def _placeholders(slide: "Slide") -> list[Element]:
    out = []
    for shape in slide._sp_tree():
        if local_name(shape) != "sp":
            continue
        if find(shape, "p:nvSpPr/p:nvPr/p:ph") is not None:
            out.append(shape)
    return out


def _ph(shape: Element) -> tuple:
    node = shape.find(f".//{qn('p:ph')}")
    return node.get("type"), get_int(node, "idx")


def _build(document: "Document", slide: "Slide", draft: SlideDraft, language: str,
           index: int) -> None:
    package = document.package
    tree = slide._sp_tree()
    available = _placeholders(slide)
    used: dict[int, Element] = {}  # id(element) -> element, kept
    order: list[Element] = []

    # The title.
    title = next((s for s in available if _ph(s)[0] in TITLE_TYPES), None)
    heading = draft.heading
    if title is not None:
        available.remove(title)
        if plain(heading).strip() or draft.title_named:
            _write_text(document, slide, title, [("p", heading, "")] if heading else [], language)
            used[id(title)] = title
    title_box = None
    if title is None and plain(heading).strip():
        x, y, w, h = _title_area(document, slide)
        title_box = slide.add_textbox(x, y, w, h)._element
        _write_text(document, slide, title_box, [("p", heading, "")], language)

    # Items naming a placeholder take it; then the rest fill free ones in layout order.
    placed: list[tuple[Item, Element | None]] = []
    for item in draft.items:
        target = None
        hint = item.hint
        if hint is not None and hint.placeholder is not None:
            target = _take(available, hint.placeholder)
            if target is None and item.kind == "text":
                target = _new_placeholder(slide, hint.placeholder, language)
        placed.append((item, target))
    # A shape comment without "ph:" says the text was not in a placeholder: it stays free.
    for k, (item, target) in enumerate(placed):
        if target is None and item.hint is None:
            placed[k] = (item, _free_placeholder(available, item.kind))

    free = [item for item, target in placed if target is None and (item.blocks or item.kind != "text")]
    heading_box = title if title is not None and id(title) in used else title_box
    cells = iter(_grid(_below(document, slide, _content_area(document, slide), heading_box),
                       len(free)))

    for item, target in placed:
        if item.kind == "text":
            if target is not None:
                _write_text(document, slide, target, item.blocks, language)
                used[id(target)] = target
                order.append(target)
            elif item.blocks:
                x, y, w, h = next(cells)
                box = slide.add_textbox(x, y, w, h)._element
                _write_text(document, slide, box, item.blocks, language)
                order.append(box)
        elif item.kind == "table":
            frame = _placeholder_frame(slide, _ph(target)) if target is not None else next(cells)
            element = _add_table(document, slide, item, frame, language)
            if target is not None:
                _adopt_placeholder(element, "p:nvGraphicFramePr", _ph(target))
                remove(target)
            order.append(element)
        elif item.kind == "picture":
            frame = _placeholder_frame(slide, _ph(target)) if target is not None else next(cells)
            element = _add_picture(document, slide, item, frame, target)
            order.append(element)

    # Placeholders nothing went into, nor named, go: an empty one draws nothing.
    for shape in available:
        if id(shape) not in used and shape.getparent() is not None:
            remove(shape)
    if title is not None and id(title) not in used and title.getparent() is not None:
        remove(title)

    # The shape tree in the outline's order, the title first.
    front = [title] if title is not None and id(title) in used else []
    front += [title_box] if title_box is not None else []
    for element in front + order:
        if element.getparent() is tree:
            tree.remove(element)
    anchor = tree.find(qn("p:extLst"))
    for element in front + order:
        if anchor is not None:
            anchor.addprevious(element)
        else:
            tree.append(element)

    if draft.hidden:
        root = package.tree(slide.part_path)
        root.set("show", "0")
    package.mark_dirty(slide.part_path)
    slide._invalidate()

    if draft.notes:
        from ..edit import notes as _notes

        body_shape = _notes.add_notes(document, slide.part_path, index + 1)
        part = _notes_part_of(document, slide)
        writer = _Writer(document, part, notes_chain(document, body_shape), language)
        body = body_shape.find(qn("p:txBody"))
        _fill_body(body, writer.paragraphs(draft.notes), language)
        package.mark_dirty(part)


def _notes_part_of(document: "Document", slide: "Slide") -> str:
    from ..edit import notes as _notes

    part = _notes.notes_part(document.package, slide.part_path)
    assert part is not None
    return part


def _take(available: list[Element], placeholder: tuple) -> Element | None:
    kind, idx = placeholder
    for shape in available:
        if _ph(shape)[1] == idx and (idx is not None or _ph(shape)[0] == kind):
            available.remove(shape)
            return shape
    for shape in available:
        if idx is None and _ph(shape)[0] == kind:
            available.remove(shape)
            return shape
    return None


def _free_placeholder(available: list[Element], kind: str) -> Element | None:
    if kind == "text":
        wanted = lambda t: t in _TEXT_TYPES  # noqa: E731
    elif kind == "table":
        wanted = lambda t: t in (None, "obj", "tbl")  # noqa: E731
    else:
        wanted = lambda t: t in (None, "obj", "pic")  # noqa: E731
    preferred = {"table": "tbl", "picture": "pic"}.get(kind)
    for shape in available:
        if preferred is not None and _ph(shape)[0] == preferred:
            available.remove(shape)
            return shape
    for shape in available:
        if wanted(_ph(shape)[0]):
            available.remove(shape)
            return shape
    return None


def _new_placeholder(slide: "Slide", placeholder: tuple, language: str) -> Element:
    """A placeholder the layout does not have (the outline names it): inheriting from the
    master by type."""
    kind, idx = placeholder
    identifier = slide._next_shape_id()
    shape = make("p:sp")
    nv = make("p:nvSpPr")
    nv.append(make("p:cNvPr", id=str(identifier), name=f"Placeholder {identifier - 1}"))
    locks = make("p:cNvSpPr")
    locks.append(make("a:spLocks", noGrp="1"))
    nv.append(locks)
    nv_pr = make("p:nvPr")
    ph = make("p:ph")
    if kind is not None:
        ph.set("type", kind)
    if idx is not None:
        ph.set("idx", str(idx))
    nv_pr.append(ph)
    nv.append(nv_pr)
    shape.append(nv)
    shape.append(make("p:spPr"))
    body = make("p:txBody")
    body.append(make("a:bodyPr"))
    body.append(make("a:lstStyle"))
    paragraph = make("a:p")
    paragraph.append(make("a:endParaRPr", lang=language))
    body.append(paragraph)
    shape.append(body)
    tree = slide._sp_tree()
    anchor = tree.find(qn("p:extLst"))
    if anchor is not None:
        anchor.addprevious(shape)
    else:
        tree.append(shape)
    slide._invalidate()
    return shape


def _write_text(document: "Document", slide: "Slide", element: Element, blocks: list,
                language: str) -> None:
    slide._invalidate()
    shape = slide._wrap(element)
    writer = _Writer(document, slide.part_path, style_chain(document, slide, shape), language)
    body = element.find(qn("p:txBody"))
    if body is None:
        body = shape._text_body(create=True)
    _fill_body(body, writer.paragraphs(blocks), language)
    document.package.mark_dirty(slide.part_path)


def _adopt_placeholder(element: Element, nv_tag: str, placeholder: tuple, *,
                       picture: bool = False) -> None:
    """Make a new table or picture the placeholder's content, as PowerPoint does when one
    is inserted into a placeholder: the frame carries the placeholder's ``p:ph``."""
    nv = element.find(qn(nv_tag))
    nv_pr = nv.find(qn("p:nvPr"))
    kind, idx = placeholder
    ph = make("p:ph")
    if kind is not None:
        ph.set("type", kind)
    if idx is not None:
        ph.set("idx", str(idx))
    nv_pr.insert(0, ph)
    if picture:
        locks = nv.find(f"{qn('p:cNvPicPr')}/{qn('a:picLocks')}")
        if locks is not None:
            locks.attrib.clear()
            locks.set("noGrp", "1")
            locks.set("noChangeAspect", "1")


def _add_table(document: "Document", slide: "Slide", item: Item, frame: tuple,
               language: str) -> Element:
    _, header, rows = item.table
    grid = [list(header)] + [list(r) for r in rows]
    columns = max(len(r) for r in grid) if grid else 1
    x, y, width, height = frame
    height = max(1, min(height, ROW_HEIGHT * len(grid))) if height else ROW_HEIGHT * len(grid)
    shape = slide.add_table(len(grid), columns, x, y, width, height)
    element = shape._element
    writer = _Writer(document, slide.part_path, [], language)
    tbl = element.find(f".//{qn('a:tbl')}")
    for r, row in enumerate(tbl.findall(qn("a:tr"))):
        for c, cell in enumerate(row.findall(qn("a:tc"))):
            inline = grid[r][c] if c < len(grid[r]) else ()
            body = cell.find(qn("a:txBody"))
            if inline:
                _fill_body(body, [writer.paragraph("p", 0, 1, inline)], language)
    document.package.mark_dirty(slide.part_path)
    return element


def _add_picture(document: "Document", slide: "Slide", item: Item, frame: tuple,
                 target: Element | None) -> Element:
    alt, _, _ = item.image
    pixels = image_size(item.data)
    placeholder = _ph(target) if target is not None else None
    if placeholder is not None and placeholder[0] == "pic":
        x, y, w, h = frame
        crop = _crop(frame, pixels)
    else:
        x, y, w, h = _fit(frame, pixels)
        crop = {}
    shape = slide.add_picture(item.data, x, y, w, h, description=alt or None)
    element = shape._element
    if crop:
        blip_fill = element.find(qn("p:blipFill"))
        stretch = blip_fill.find(qn("a:stretch"))
        stretch.addprevious(make("a:srcRect", **crop))
    if placeholder is not None:
        _adopt_placeholder(element, "p:nvPicPr", placeholder, picture=True)
        remove(target)
    document.package.mark_dirty(slide.part_path)
    return element


__all__ = ["OutlineWarning", "ROLES", "choose_role", "find_layout", "insert_outline",
           "read_outline"]
