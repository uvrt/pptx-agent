"""What a change did to a deck, in the deck's own addresses.

The tool layer reports, for every call that changed a document, what it changed
(``Result.touched``, for the application -- a live view that re-renders only what changed;
the model never reads it).  ooxml-edit finds which parts' bytes differ
(:class:`ooxml_edit.history.Change`, from the snapshots its history takes anyway); this
module says what that means in a deck, so it holds for every tool, undo and redo
included, and for a tool written tomorrow, without the tool saying anything:

* a slide part that changed is compared shape by shape -- each shape's own XML (a group's
  without its members), its group and its place among its siblings -- by the shape ids the
  tools use (:class:`~pptx_agent.edit.ids.SlideShapeIndex`): ``{"slide": "s:258",
  "shapes": ["258.7", "258.9"]}``, with ``created`` and ``removed`` listing the shapes in
  ``shapes`` that are new or gone.  A change outside every shape -- the background, the
  layout the slide is on, a transition, hidden -- adds ``"whole": true``: redraw the slide.
* a part a shape draws from -- a chart, its workbook, a diagram's parts, a picture's
  image -- counts as a change to the shape whose XML refers to it, found through the
  relationships; a change to a slide's relationships, as one to the shapes that use them.
* speaker notes and review comments are ``"notes": true`` and ``"comments": true`` on the
  slide's entry; neither is drawn on the slide.
* a new or deleted slide is ``{"slide": "s:300", "created": ["s:300"]}`` (``removed``).
* a layout or master is ``{"layout": "L2", ...}`` (``"master": "M0"``) as a slide is,
  with ``affects``: the slides built on it, which show the change -- those slides are not
  listed again.  A theme is ``{"theme": "M0", "affects": [...]}``: the master whose theme
  it is.
* the deck itself is one ``{"deck": [...], "affects": [...]}`` entry: ``order`` (a slide
  moved, or slides before it came or went: ``affects`` the slides whose position
  changed, which matters to their slide numbers), ``size`` (every slide affected),
  ``text_styles`` (the presentation's default text styles: every slide), ``sections``,
  ``table_styles`` (the slides with tables), ``document_properties``, and the name of any
  other deck-level part's relationship (``commentAuthors``, ``presProps``, ...).
* anything else is ``{"part": name}``: never left out.

Addresses only read -- a ``like`` shape copied from, a source slide -- never appear: only
bytes that changed count.  A shape whose only change is the id stamp the tools write the
first time they edit one (:mod:`~pptx_agent.edit.ids`) has not changed.
"""

from __future__ import annotations

import bisect
import functools
import posixpath
import re
from typing import Any, Callable, Iterable

from lxml import etree
from ooxml_edit.opc import normalize_part_path, part_for_rels, rels_path_for, resolve_target
from ooxml_edit.xml import parse_xml

from ..oxml.xml import PA_ID_EXT_URI, find, qn
from .comments import COMMENT_REL_EXT_URI
from .ids import SlideShapeIndex

_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_PR = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_SHAPES = frozenset(qn(tag) for tag in ("p:sp", "p:pic", "p:cxnSp", "p:grpSp",
                                        "p:graphicFrame"))
_GROUP = qn("p:grpSp")
_XMLNS = re.compile(rb'\s+xmlns(?::[\w.-]+)?="[^"]*"')

#: Presentation-part children and the deck change they make (``None``: what the slide
#: and template entries already say).
_PRESENTATION = {"sldIdLst": None, "sldMasterIdLst": None, "notesMasterIdLst": None,
                 "handoutMasterIdLst": None, "sldSz": "size", "notesSz": "notes_size",
                 "defaultTextStyle": "text_styles"}
#: Deck changes every slide shows.
_EVERY_SLIDE = ("size", "text_styles")


def touched(deck, change) -> list[dict[str, Any]]:
    """What ``change`` (an :class:`ooxml_edit.history.Change` of ``deck``'s package) did,
    as ``Result.touched`` entries (see the module); ``[]`` when no part changed.

    For example::

        with deck.history.watch() as change:
            deck.shape("256.3").set_text("Q3")
        touched(deck, change)          # [{'slide': 's:256', 'shapes': ['256.3']}]
    """
    if not change.parts:
        return []
    return _Touched(deck, change).entries()


class _Host:
    """What changed on one slide, layout or master."""

    def __init__(self) -> None:
        self.shapes: dict[str, None] = {}
        self.created: dict[str, None] = {}
        self.removed: dict[str, None] = {}
        self.whole = False
        self.notes = False
        self.comments = False
        self.made: str | None = None     # "created" / "removed": the host itself

    def json(self, kind: str, address: str, affects: list[str] | None) -> dict[str, Any]:
        out: dict[str, Any] = {kind: address}
        if self.made is not None:
            out[self.made] = [address]
            return out
        if self.shapes:
            out["shapes"] = list(self.shapes)
        if self.created:
            out["created"] = list(self.created)
        if self.removed:
            out["removed"] = list(self.removed)
        for flag in ("whole", "notes", "comments"):
            if getattr(self, flag):
                out[flag] = True
        if affects is not None:
            out["affects"] = affects
        return out


class _Touched:
    def __init__(self, deck, change) -> None:
        self.deck = deck
        self.change = change
        self.package = deck.package
        self.parts = change.parts
        presentation = self.package.presentation_part()
        self.presentation = presentation
        self.after_slides = _slide_map(change.after, presentation)
        if presentation in self.parts or rels_path_for(presentation) in self.parts:
            self.before_slides = _slide_map(change.before, presentation)
        else:
            self.before_slides = self.after_slides
        self._templates: dict[str, Any] | None = None
        self.slides: dict[int, _Host] = {}
        self.sheets: dict[str, _Host] = {}       # template part -> its changes
        self.themes: dict[str, None] = {}        # master keys
        self.deck_changes: dict[str, None] = {}
        self.deck_affects: dict[str, None] = {}
        self.other: dict[str, None] = {}

    @property
    def templates(self) -> dict[str, Any]:
        """``{part: layout or master}``, read when a part is no slide's."""
        if self._templates is None:
            self._templates = {t.part_path: t for t in [*self.deck.masters,
                                                        *self.deck.layouts]}
        return self._templates

    @property
    def masters(self) -> set[str]:
        return {part for part, template in self.templates.items() if template.id[0] == "M"}

    # -- the walk ----------------------------------------------------------------------

    def entries(self) -> list[dict[str, Any]]:
        before_ids = {sid: part for part, sid in self.before_slides.items()}
        after_ids = {sid: part for part, sid in self.after_slides.items()}
        for sid in after_ids:
            if sid not in before_ids:
                self._slide(sid).made = "created"
        for sid in before_ids:
            if sid not in after_ids:
                self._slide(sid).made = "removed"
        for sid, part in after_ids.items():
            if sid in before_ids and before_ids[sid] != part:
                self._slide(sid).whole = True        # another part now: redraw it
        self._order(list(before_ids), list(after_ids))

        for part in sorted(self.parts):
            self._part(part)
        return self._json()

    def _part(self, part: str) -> None:
        if part == "[Content_Types].xml" or part == "_rels/.rels":
            return
        if part.endswith(".rels"):
            owner = part_for_rels(part)
            if owner == self.presentation:
                return                              # the slide list says it
            host = self._host_of(owner)
            if host is not None:
                self._relationships(owner, host)
            elif owner not in self.parts:
                self._drawn_from(owner, set())      # a chart's: as a change to the chart
            return
        if part == self.presentation:
            self._presentation()
            return
        host = self._host_of(part)
        if host is not None:
            self._shapes(part, host)
            return
        self._drawn_from(part, set())

    def _host_of(self, part: str) -> tuple[str, Any] | None:
        """``("slide", sid)`` or ``("sheet", part)`` for a slide's or template's part."""
        sid = self.after_slides.get(part, self.before_slides.get(part))
        if sid is not None:
            return ("slide", sid)
        if part in self.templates:
            return ("sheet", part)
        return None

    def _entry(self, host: tuple[str, Any]) -> _Host:
        return self._slide(host[1]) if host[0] == "slide" else self._sheet(host[1])

    def _slide(self, sid: int) -> _Host:
        return self.slides.setdefault(sid, _Host())

    def _sheet(self, part: str) -> _Host:
        return self.sheets.setdefault(part, _Host())

    # -- a slide's or template's own part --------------------------------------------------

    def _shapes(self, part: str, host: tuple[str, Any]) -> None:
        entry = self._entry(host)
        if entry.made is not None:
            return
        before, after = self.change.before(part), self.change.after(part)
        if before is None or after is None:
            # A template added or deleted (a slide's are known by the slide list).
            if host[0] == "sheet":
                entry.made = "created" if before is None else "removed"
            return
        key = host[1] if host[0] == "slide" else self.templates[part].id
        old, new = _Shapes(key, parse_xml(before)), _Shapes(key, parse_xml(after))
        for local in new.own:
            if local not in old.own:
                entry.created[local] = None
            elif (new.own[local] != old.own[local]
                  or new.parent[local] != old.parent[local]):
                entry.shapes[local] = None
        for local in old.own:
            if local not in new.own:
                entry.removed[local] = None
        for parent, order in new.order.items():
            was = [local for local in old.order.get(parent, ()) if local in new.own
                   and new.parent[local] == parent]
            now = [local for local in order if local in old.own and old.parent[local] == parent]
            if was != now:
                for local in _moved(was, now):
                    entry.shapes[local] = None
        if new.rest != old.rest:
            entry.whole = True
        prefix = f"{key}."
        entry.shapes.update(entry.created)
        entry.shapes.update(entry.removed)
        entry.shapes = {(local if local.startswith(prefix) else prefix + local): None
                        for local in entry.shapes}
        entry.created = {prefix + local: None for local in entry.created}
        entry.removed = {prefix + local: None for local in entry.removed}

    def _relationships(self, owner: str, host: tuple[str, Any]) -> None:
        """A slide's (template's) relationships changed: the shapes using each one that
        came, went or points elsewhere now; its layout (master) one, the whole slide."""
        entry = self._entry(host)
        if entry.made is not None:
            return
        rels = rels_path_for(owner)
        before = _relationships(self.change.before(rels), owner)
        after = _relationships(self.change.after(rels), owner)
        for rid in sorted(set(before) | set(after)):
            if before.get(rid) == after.get(rid):
                continue
            kind = (after.get(rid) or before.get(rid))[0]
            self._used(owner, host, rid, kind)

    def _used(self, owner: str, host: tuple[str, Any], rid: str, kind: str) -> None:
        """What on ``owner`` a change through relationship ``rid`` (of type ``kind``)
        changes."""
        entry = self._entry(host)
        name = kind.rsplit("/", 1)[-1]
        if name in ("slideLayout", "slideMaster"):
            # What a slide (layout) is built on; a master's layouts are their own entries.
            if owner not in self.masters:
                entry.whole = True
            return
        if name == "notesSlide":
            entry.notes = True
            return
        if name in ("comments", "commentAuthors"):
            entry.comments = True
            return
        key = host[1] if host[0] == "slide" else self.templates[owner].id
        found = False
        for state in (self.change.after, self.change.before):
            data = state(owner)
            if data is None:
                continue
            shapes = _Shapes(key, parse_xml(data))
            for local in shapes.using(rid):
                found = True
                if local is None:
                    entry.whole = True
                else:
                    entry.shapes[f"{key}.{local}"] = None
            if found:
                return

    # -- parts drawn from ----------------------------------------------------------------

    def _drawn_from(self, part: str, seen: set[str]) -> None:
        """A part that is no slide's or template's own: whatever refers to it changed."""
        if part in seen:
            return
        seen.add(part)
        sources = self.package.relationship_sources(part)
        if not sources:
            if self.change.after(part) is not None:
                self.other[part] = None     # an orphan, or the package's own: say it
            return                          # gone with what used it, which says so
        for owner, rel in sources:
            name = rel.type.rsplit("/", 1)[-1]
            if owner == "":
                self.deck_changes["document_properties"] = None
                continue
            if owner == self.presentation:
                if name == "theme" and any(
                        o in self.templates for o, r in sources if r.type == rel.type):
                    continue                        # its master's, said there
                if name == "tableStyles":
                    self.deck_changes["table_styles"] = None
                    self.deck_affects.update(dict.fromkeys(self._slides_with_tables()))
                elif name != "slideMaster":
                    self.deck_changes[name] = None
                continue
            if name == "theme" and owner in self.templates:
                template = self.templates[owner]
                self.themes[template.id] = None
                continue
            host = self._host_of(owner)
            if host is not None:
                self._used(owner, host, rel.id, rel.type)
                continue
            notes_of = self._notes_owner(owner)
            if notes_of is not None:
                self._slide(notes_of).notes = True
                continue
            if name == "theme":
                self.deck_changes[_kind_of(owner)] = None     # the notes master's, say
                continue
            self._drawn_from(owner, seen)

    def _notes_owner(self, part: str) -> int | None:
        """The slide whose notes ``part`` is (its relationship back to the slide)."""
        for owner, rel in self.package.relationship_sources(part):
            if rel.type.endswith("/notesSlide"):
                return self.after_slides.get(owner)
        return None

    def _slides_with_tables(self) -> list[str]:
        out = []
        for sid, part in ((sid, part) for part, sid in self.after_slides.items()):
            data = self.package.read(part) or b""
            if b"<a:tbl>" in data or b"<a:tbl " in data:
                out.append(f"s:{sid}")
        return out

    # -- the presentation part -----------------------------------------------------------

    def _presentation(self) -> None:
        before = self.change.before(self.presentation)
        after = self.change.after(self.presentation)
        if before is None or after is None:
            self.deck_changes["presentation"] = None
            return
        old, new = _children(parse_xml(before)), _children(parse_xml(after))
        for name in sorted(set(old) | set(new)):
            if old.get(name) == new.get(name):
                continue
            if name in _PRESENTATION:
                what = _PRESENTATION[name]
            elif name == "extLst":
                what = ("sections" if b"sectionLst" in (old.get(name, b"") + new.get(name, b""))
                        else "presentation")
            else:
                what = "presentation"
            if what is not None:
                self.deck_changes[what] = None
        if any(what in self.deck_changes for what in _EVERY_SLIDE):
            self.deck_affects.update(dict.fromkeys(f"s:{sid}" for sid in
                                                   self.after_slides.values()))

    def _order(self, before: list[int], after: list[int]) -> None:
        """Slides whose position changed (and stayed in the deck): their numbers did."""
        where = {sid: n for n, sid in enumerate(before)}
        moved = [f"s:{sid}" for n, sid in enumerate(after) if sid in where and where[sid] != n]
        if moved:
            self.deck_changes["order"] = None
            self.deck_affects.update(dict.fromkeys(moved))

    # -- out -----------------------------------------------------------------------------

    def _json(self) -> list[dict[str, Any]]:
        out: list[dict[str, Any]] = []
        order = [*self.after_slides.values(),
                 *(sid for sid in self.before_slides.values()
                   if sid not in set(self.after_slides.values()))]
        for sid in order:
            entry = self.slides.get(sid)
            if entry is not None and _says(entry):
                out.append(entry.json("slide", f"s:{sid}", None))
        layouts = [t for t in self.deck.layouts if t.part_path in self.sheets] \
            if self.sheets else []
        masters = [t for t in self.deck.masters if t.part_path in self.sheets] \
            if self.sheets else []
        for kind, templates in (("layout", layouts), ("master", masters)):
            for template in templates:
                entry = self.sheets[template.part_path]
                if _says(entry):
                    out.append(entry.json(kind, template.id, None if entry.made else
                                          self._on(template.part_path)))
        for part, entry in self.sheets.items():
            if part not in self.templates and _says(entry):
                out.append({"part": part, **({entry.made: [part]} if entry.made else {})})
        for key in self.themes:
            master = next((t for t in self.deck.masters if t.id == key), None)
            out.append({"theme": key,
                        "affects": self._on(master.part_path) if master is not None else []})
        if self.deck_changes:
            deck: dict[str, Any] = {"deck": list(self.deck_changes)}
            if self.deck_affects:
                deck["affects"] = list(self.deck_affects)
            out.append(deck)
        out += [{"part": part} for part in self.other]
        return out

    def _on(self, part: str) -> list[str]:
        return [f"s:{slide.slide_id}" for slide in self.deck._slides_on(part)]


def _moved(was: list[str], now: list[str]) -> list[str]:
    """The fewest shapes of ``now`` whose moving explains the new order (the rest, the
    longest run in the order they had, stays): ``[2, 3, 4] -> [2, 4, 3, ...]`` moved one."""
    where = {local: n for n, local in enumerate(was)}
    places = [where[local] for local in now]
    tails: list[int] = []            # index into now of the smallest tail of each length
    back: list[int] = [-1] * len(now)
    for n, place in enumerate(places):
        low = bisect.bisect_left([places[t] for t in tails], place)
        back[n] = tails[low - 1] if low else -1
        if low == len(tails):
            tails.append(n)
        else:
            tails[low] = n
    kept: set[int] = set()
    n = tails[-1] if tails else -1
    while n != -1:
        kept.add(n)
        n = back[n]
    return [local for n, local in enumerate(now) if n not in kept]


def _says(entry: _Host) -> bool:
    return bool(entry.made or entry.shapes or entry.whole or entry.notes or entry.comments)


class _Shapes:
    """One slide's (template's) shapes as compared: each one's own XML, its group, the
    order of each group's members, and everything outside the shapes."""

    def __init__(self, key: "int | str", root) -> None:
        tree = find(root, "p:cSld/p:spTree")
        # The ids first, stamps and all; then the stamps go from what is compared.
        self.index = SlideShapeIndex(key, tree) if tree is not None else None
        _unstamp(root)
        self.own: dict[str, bytes] = {}
        self.parent: dict[str, str | None] = {}
        self.order: dict[str | None, list[str]] = {}
        self.root = root
        if self.index is not None:
            for shape_id, element in self.index:
                local = shape_id.local
                self.own[local] = _own(element)
                up = element.getparent()
                parent = None
                if up is not None and up.tag == _GROUP:
                    found = self.index.id_of(up)
                    parent = found.local if found is not None else None
                self.parent[local] = parent
                self.order.setdefault(parent, []).append(local)
        self.rest = _rest(root, tree)

    def using(self, rid: str) -> list[str | None]:
        """The shapes whose XML refers to relationship ``rid`` (innermost); ``None`` for a
        reference outside every shape (the background's picture)."""
        found: dict[str | None, None] = {}
        for element in self.root.iter():
            if not isinstance(element.tag, str):
                continue
            for name, value in element.attrib.items():
                if value == rid and (name.startswith(_R) or name.endswith("relId")):
                    found[self._shape_at(element)] = None
        return list(found)

    def _shape_at(self, element) -> str | None:
        while element is not None:
            if element.tag in _SHAPES and self.index is not None:
                found = self.index.id_of(element)
                if found is not None:
                    return found.local
            element = element.getparent()
        return None


def _own(element) -> bytes:
    """A shape's own XML: a group's without its members."""
    if element.tag == _GROUP:
        parts = [repr(sorted(element.attrib.items())).encode()]
        parts += [_bytes(child) for child in element if child.tag not in _SHAPES]
        return b"\0".join(parts)
    return _bytes(element)


def _bytes(element) -> bytes:
    """``element``'s XML without namespace declarations, which say where a prefix was
    declared (the part's root, say, once another part of it needs one), not what is drawn."""
    return _XMLNS.sub(b"", etree.tostring(element, with_tail=False))


def _unstamp(root) -> None:
    """Take the tools' id stamps (``a:ext`` of :data:`~pptx_agent.oxml.xml.PA_ID_EXT_URI`)
    out of a parsed part, and an ``a:extLst`` they leave empty: written the first time a
    tool edits a shape, they change no drawing (the shape's id stays what it was)."""
    for ext in list(root.iter(qn("a:ext"))):
        if ext.get("uri") == PA_ID_EXT_URI:
            holder = ext.getparent()
            holder.remove(ext)
            if len(holder) == 0 and holder.getparent() is not None:
                holder.getparent().remove(holder)


def _rest(root, tree) -> bytes:
    """Everything of a slide part outside its shapes -- but the link to its review
    comments (:data:`COMMENT_REL_EXT_URI`) and a master's list of its layouts, which are
    no drawing: taken out of ``root``."""
    for listed in root.findall(qn("p:sldLayoutIdLst")):
        root.remove(listed)
    for ext in list(root.iterfind(f"{qn('p:extLst')}/{qn('p:ext')}")):
        if ext.get("uri") == COMMENT_REL_EXT_URI:
            holder = ext.getparent()
            holder.remove(ext)
            if len(holder) == 0:
                root.remove(holder)
    parts = [repr(sorted(root.attrib.items())).encode()]
    for child in root:
        if not isinstance(child.tag, str):
            continue
        if child.tag != qn("p:cSld"):
            parts.append(_bytes(child))
            continue
        parts.append(repr(sorted(child.attrib.items())).encode())
        for grand in child:
            if not isinstance(grand.tag, str):
                continue
            if grand is not tree:
                parts.append(_bytes(grand))
                continue
            parts += [_bytes(node) for node in grand
                      if isinstance(node.tag, str) and node.tag not in _SHAPES]
    return b"\0".join(parts)


def _children(root) -> dict[str, bytes]:
    out: dict[str, bytes] = {"": repr(sorted(root.attrib.items())).encode()}
    for child in root:
        if isinstance(child.tag, str):
            name = etree.QName(child).localname
            out[name] = out.get(name, b"") + _bytes(child)
    return out


def _relationships(data: bytes | None, owner: str) -> dict[str, tuple[str, str]]:
    """A relationships part's ``{Id: (Type, target)}``."""
    if data is None:
        return {}
    base = posixpath.dirname(owner)
    out = {}
    for rel in parse_xml(data).iter(f"{_PR}Relationship"):
        target = rel.get("Target") or ""
        if rel.get("TargetMode") != "External":
            target = resolve_target(base, target)
        out[rel.get("Id") or ""] = (rel.get("Type") or "", target)
    return out


def _slide_map(read: Callable[[str], bytes | None], presentation: str) -> dict[str, int]:
    """``{slide part: sldId}`` in presentation order, from one state's presentation part."""
    data = read(presentation)
    if data is None:
        return {}
    return dict(_slides_of(data, read(rels_path_for(presentation)), presentation))


@functools.lru_cache(maxsize=8)
def _slides_of(data: bytes, rels_data: bytes | None, presentation: str
               ) -> tuple[tuple[str, int], ...]:
    """:func:`_slide_map` of a presentation part's bytes and its relationships' -- the same
    bytes call after call while the slide list stays as it is."""
    rels = _relationships(rels_data, presentation)
    root = parse_xml(data)
    out: dict[str, int] = {}
    for node in root.iter(qn("p:sldId")):
        target = rels.get(node.get(f"{_R}id") or "")
        raw = node.get("id")
        if target is None or raw is None or not raw.isdigit():
            continue
        out[normalize_part_path(target[1])] = int(raw)
    return tuple(out.items())


def _kind_of(part: str) -> str:
    """``notesMaster`` for ``ppt/notesMasters/notesMaster1.xml``: a part's kind by name."""
    return re.sub(r"\d*\.xml$", "", posixpath.basename(part)) or part


__all__: Iterable[str] = ["touched"]
