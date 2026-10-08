"""The deck's metadata: ``docProps/core.xml`` and ``docProps/app.xml``.

**Core properties** (title, author, created, modified) are ordinary, undoable edits of
``core.xml``; a deck without one gets one, related from the package root.

**Extended properties** (``app.xml``) restate the deck's structure for indexers: how many
slides, notes and hidden slides it has, how many words and paragraphs, and -- under
``TitlesOfParts`` -- every slide's title, after the fonts and themes.  PowerPoint rewrites
the whole part on every save and does not check it on open, so nothing breaks when it is
stale, but a deck should not claim three slides when it has five.  :func:`refreshed_app`
brings those values up to date as the deck is *written* (:meth:`Document.save`), and only
those an edit changed: it compares the slides with the deck as it was opened, writes the
slide count, moves the other counts by the difference, and replaces the slide titles when
they changed.  The edits themselves never touch ``app.xml``, so undo and redo stay exact; a
deck saved without changes -- or with changes that alter no count and no title -- keeps the
part's bytes.

PowerPoint's own rules, measured on a deck it wrote with a slide per layout: a slide's title
is its title placeholder's text, and a slide without one is listed as "PowerPoint
Presentation" (in the UI language: "PowerPoint-presentatie" in Dutch); ``Paragraphs`` counts
paragraphs with text and ``Words`` their words.  Which ``HeadingPairs`` group holds the slide
titles is found by its English name ("Slide Titles") or, in a localised file, as the group
whose entries are the deck's slide titles as they were read.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import TYPE_CHECKING

from lxml import etree

from ..oxml.xml import Element, find, local_name, qn, serialize

if TYPE_CHECKING:  # pragma: no cover
    from ..oxml.package import PresentationPackage

CORE_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "dcmitype": "http://purl.org/dc/dcmitype/",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
}
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

CORE_PART = "docProps/core.xml"
APP_PART = "docProps/app.xml"
REL_CORE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
REL_APP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
CT_CORE = "application/vnd.openxmlformats-package.core-properties+xml"

APPLICATION = "pptx-agent"
UNTITLED = "PowerPoint Presentation"
SLIDE_TITLES = "Slide Titles"
FONTS_USED = "Fonts Used"
THEME = "Theme"

_TITLE_TYPES = {"title", "ctrTitle"}


def _c(tag: str) -> str:
    prefix, _, name = tag.partition(":")
    return "{%s}%s" % (CORE_NS[prefix], name)


def w3cdtf(moment: datetime) -> str:
    """``2026-10-03T09:40:41Z``: UTC, whole seconds, as Office writes it."""
    if moment.tzinfo is None:
        moment = moment.replace(tzinfo=timezone.utc)
    return moment.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_w3cdtf(text: str | None) -> datetime | None:
    if not text:
        return None
    value = text.strip()
    if value.endswith("Z"):
        value = value[:-1] + "+00:00"
    try:
        moment = datetime.fromisoformat(value)
    except ValueError:
        return None
    return moment if moment.tzinfo else moment.replace(tzinfo=timezone.utc)


def now() -> datetime:
    return datetime.now(timezone.utc).replace(microsecond=0)


def core_xml(*, title: str | None, author: str | None, created: datetime,
             modified: datetime | None = None) -> bytes:
    """A new ``core.xml``, as PowerPoint writes one for a new deck."""
    root = etree.Element(_c("cp:coreProperties"), nsmap=CORE_NS)
    etree.SubElement(root, _c("dc:title")).text = title or ""
    if author:
        etree.SubElement(root, _c("dc:creator")).text = author
        etree.SubElement(root, _c("cp:lastModifiedBy")).text = author
    etree.SubElement(root, _c("cp:revision")).text = "1"
    for tag, moment in (("dcterms:created", created), ("dcterms:modified", modified or created)):
        node = etree.SubElement(root, _c(tag))
        node.set(_c("xsi:type"), "dcterms:W3CDTF")
        node.text = w3cdtf(moment)
    return serialize(root)


def app_xml(*, format_name: str, fonts: tuple[str, ...], themes: tuple[str, ...],
            titles: tuple[str, ...], slides: int = 0, notes: int = 0, hidden: int = 0,
            words: int = 0, paragraphs: int = 0) -> bytes:
    """A new ``app.xml`` in PowerPoint's element order, with English group names."""
    root = etree.Element("{%s}Properties" % APP_NS, nsmap={None: APP_NS, "vt": VT_NS})

    def add(tag: str, text: str) -> None:
        etree.SubElement(root, "{%s}%s" % (APP_NS, tag)).text = text

    add("TotalTime", "0")
    add("Words", str(words))
    add("Application", APPLICATION)
    add("PresentationFormat", format_name)
    add("Paragraphs", str(paragraphs))
    add("Slides", str(slides))
    add("Notes", str(notes))
    add("HiddenSlides", str(hidden))
    add("MMClips", "0")
    add("ScaleCrop", "false")
    groups = [(FONTS_USED, fonts), (THEME, themes), (SLIDE_TITLES, titles)]
    pairs = etree.SubElement(root, "{%s}HeadingPairs" % APP_NS)
    vector = etree.SubElement(pairs, "{%s}vector" % VT_NS, size=str(2 * len(groups)),
                              baseType="variant")
    for name, members in groups:
        etree.SubElement(etree.SubElement(vector, "{%s}variant" % VT_NS),
                         "{%s}lpstr" % VT_NS).text = name
        etree.SubElement(etree.SubElement(vector, "{%s}variant" % VT_NS),
                         "{%s}i4" % VT_NS).text = str(len(members))
    parts = etree.SubElement(root, "{%s}TitlesOfParts" % APP_NS)
    entries = [member for _, members in groups for member in members]
    vector = etree.SubElement(parts, "{%s}vector" % VT_NS, size=str(len(entries)),
                              baseType="lpstr")
    for entry in entries:
        etree.SubElement(vector, "{%s}lpstr" % VT_NS).text = entry
    add("Company", "")
    add("LinksUpToDate", "false")
    add("SharedDoc", "false")
    add("HyperlinksChanged", "false")
    return serialize(root)


# -- core properties -------------------------------------------------------------------------


class CoreProperties:
    """Read and write ``docProps/core.xml`` through the package (no history of its own)."""

    def __init__(self, package: "PresentationPackage") -> None:
        self._package = package

    def _part(self) -> str | None:
        for rel in self._package.relationships("").values():
            if rel.type == REL_CORE and rel.target_part and self._package.has_part(rel.target_part):
                return rel.target_part
        return None

    def _root(self, create: bool) -> Element | None:
        part = self._part()
        if part is None:
            if not create:
                return None
            part = self._package.unused_part_name(CORE_PART) \
                if self._package.has_part(CORE_PART) else CORE_PART
            root = etree.Element(_c("cp:coreProperties"), nsmap=CORE_NS)
            self._package.add_part(part, serialize(root), CT_CORE, override=True)
            self._package.add_relationship("", REL_CORE, part)
        return self._package.tree(part)

    def get(self, tag: str) -> str | None:
        root = self._root(False)
        node = None if root is None else root.find(_c(tag))
        return None if node is None else (node.text or "")

    def set(self, tag: str, value: str | None, *, dated: bool = False) -> None:
        root = self._root(True)
        assert root is not None
        node = root.find(_c(tag))
        if value is None:
            if node is not None:
                root.remove(node)
        else:
            if node is None:
                node = etree.SubElement(root, _c(tag))
            node.text = value
            if dated:
                node.set(_c("xsi:type"), "dcterms:W3CDTF")
        self._package.mark_dirty(self._part())


# -- app.xml ---------------------------------------------------------------------------------


def slide_title(root: Element | None) -> str | None:
    """A slide's title placeholder's text (paragraphs and breaks joined by spaces)."""
    if root is None:
        return None
    for placeholder in root.iter(qn("p:ph")):
        if placeholder.get("type") not in _TITLE_TYPES:
            continue
        shape = placeholder.getparent().getparent().getparent()
        body = shape.find(qn("p:txBody"))
        if body is None:
            return None
        lines = []
        for paragraph in body.findall(qn("a:p")):
            pieces = []
            for node in paragraph:
                name = local_name(node)
                if name in ("r", "fld"):
                    text = node.find(qn("a:t"))
                    pieces.append("" if text is None else (text.text or ""))
                elif name == "br":
                    pieces.append(" ")
            lines.append("".join(pieces))
        text = " ".join(line for line in lines if line).strip()
        return text or None
    return None


def _text_counts(root: Element | None) -> tuple[int, int]:
    """``(words, paragraphs)``: paragraphs with text, in shapes and table cells."""
    if root is None:
        return 0, 0
    words = paragraphs = 0
    tree = find(root, "p:cSld/p:spTree")
    if tree is None:
        return 0, 0
    for paragraph in tree.iter(qn("a:p")):
        text = "".join(node.text or "" for node in paragraph.iter(qn("a:t")))
        if text.strip():
            paragraphs += 1
            words += len(text.split())
    return words, paragraphs


def deck_summary(package: "PresentationPackage") -> dict:
    """What ``app.xml`` should say about the slides."""
    from ..oxml.package import REL_NOTES_SLIDE

    titles, notes, hidden, words, paragraphs = [], 0, 0, 0, 0
    for _, part in package.slide_parts():
        root = package.tree(part)
        titles.append(slide_title(root) or UNTITLED)
        if package.related_parts_of_type(part, REL_NOTES_SLIDE):
            notes += 1
        if root is not None and root.get("show") in {"0", "false"}:
            hidden += 1
        counted = _text_counts(root)
        words += counted[0]
        paragraphs += counted[1]
    return {"titles": titles, "notes": notes, "hidden": hidden, "words": words,
            "paragraphs": paragraphs}


def _app_part(package: "PresentationPackage") -> str | None:
    for rel in package.relationships("").values():
        if rel.type == REL_APP and rel.target_part and package.has_part(rel.target_part):
            return rel.target_part
    return None


def _groups(root: Element) -> tuple[Element | None, list[tuple[str, Element]], Element | None]:
    """``(heading vector, [(name, count node)], titles vector)``."""
    pairs = root.find("{%s}HeadingPairs/{%s}vector" % (APP_NS, VT_NS))
    titles = root.find("{%s}TitlesOfParts/{%s}vector" % (APP_NS, VT_NS))
    groups = []
    if pairs is not None:
        variants = pairs.findall("{%s}variant" % VT_NS)
        for name_variant, count_variant in zip(variants[::2], variants[1::2]):
            name = name_variant.find("{%s}lpstr" % VT_NS)
            count = count_variant.find("{%s}i4" % VT_NS)
            if name is not None and count is not None:
                groups.append((name.text or "", count))
    return pairs, groups, titles


def refreshed_app(package: "PresentationPackage", before: dict | None) -> bytes | None:
    """``app.xml`` brought up to date with the slides, or ``None`` when there is no
    ``app.xml`` or nothing it describes changed.

    ``before`` is :func:`deck_summary` of the deck as it was opened.  Only what differs from
    it is written, so an edit that changes no count and no title leaves the part alone, and
    a count is moved by the difference -- whatever the application that wrote it counted
    (notes, say), its basis is kept.  Slide titles are replaced as a group when they
    changed; the group is found by its English name or, in a localised file, as the group
    listing the titles the deck was opened with.
    """
    part = _app_part(package)
    if part is None:
        return None
    data = package.read(part)
    try:
        root = etree.fromstring(data, etree.XMLParser(resolve_entities=False))
    except etree.XMLSyntaxError:
        return None
    now_ = deck_summary(package)
    before = before or {"titles": [], "notes": 0, "hidden": 0, "words": 0, "paragraphs": 0}
    changed = False

    def shift(tag: str, current: int, previous: int, absolute: bool = False) -> None:
        nonlocal changed
        node = root.find("{%s}%s" % (APP_NS, tag))
        if node is None or current == previous:
            return
        if absolute:
            value = current
        else:
            try:
                value = max(0, int(node.text or "0") + current - previous)
            except ValueError:
                value = current
        if node.text != str(value):
            node.text = str(value)
            changed = True

    shift("Slides", len(now_["titles"]), len(before["titles"]), absolute=True)
    shift("Notes", now_["notes"], before["notes"])
    shift("HiddenSlides", now_["hidden"], before["hidden"])
    shift("Words", now_["words"], before["words"])
    shift("Paragraphs", now_["paragraphs"], before["paragraphs"])

    pairs, groups, titles = _groups(root)
    if now_["titles"] != before["titles"] and pairs is not None and titles is not None:
        entries = titles.findall("{%s}lpstr" % VT_NS)
        spans, start = [], 0
        for name, count in groups:
            try:
                size = int(count.text or "0")
            except ValueError:
                spans = []
                break
            spans.append((name, count, start, start + size))
            start += size
        target = None
        if spans and start == len(entries):
            target = next((span for span in spans if span[0] == SLIDE_TITLES), None)
            if target is None:
                for span in spans:
                    listed = [e.text or "" for e in entries[span[2]:span[3]]]
                    if len(listed) == len(before["titles"]) and all(
                            a == b or b == UNTITLED for a, b in zip(listed, before["titles"])):
                        target = span
        if target is not None:
            _, count, first, last = target
            anchor = entries[first - 1] if first > 0 else None
            for entry in entries[first:last]:
                titles.remove(entry)
            new_entries = []
            for text in now_["titles"]:
                node = etree.Element("{%s}lpstr" % VT_NS)
                node.text = text
                new_entries.append(node)
            if anchor is not None:
                for node in reversed(new_entries):
                    anchor.addnext(node)
            else:
                for node in reversed(new_entries):
                    titles.insert(0, node)
            count.text = str(len(new_entries))
            titles.set("size", str(len(titles.findall("{%s}lpstr" % VT_NS))))
            changed = True
    if not changed:
        return None
    return serialize(root)


def _canonical(data: bytes) -> bytes:
    try:
        return etree.tostring(etree.fromstring(data), method="c14n")
    except etree.XMLSyntaxError:
        return data


__all__ = ["CoreProperties", "app_xml", "core_xml", "deck_summary", "refreshed_app",
           "slide_title"]
