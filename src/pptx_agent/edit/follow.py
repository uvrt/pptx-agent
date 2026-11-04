"""Copies of a sample slide made to follow it after the sample was revised (round 31, R8).

A run of slides is often built from one sample: the sample is duplicated (or imported from a
deck of samples) and each copy gets its own text, values and pictures.  When the sample's
design changes afterwards, :func:`follow_sample` gives every copy the new design in one step
and keeps what each copy filled in.  :func:`find_copies` finds the copies when the caller
does not name them.

**Which slides are copies** (:func:`find_copies`).  What a copy keeps of its sample, measured
(PowerPoint 16 for Mac, a slide copied and pasted through AppleScript -- its Duplicate is not
scriptable -- 2026-11-04, made-up deck): every shape's ``cNvPr`` id and name, and the
pptx-agent id stamp; PowerPoint gave every shape of the pasted slide a new
``a16:creationId``.  :meth:`Document.duplicate_slide <pptx_agent.Document.duplicate_slide>`
and :meth:`~pptx_agent.Document.import_slide` keep ids, names, stamps and creation ids.  So
a shape *answers* a sample's shape when both carry the same GUID-based id (a creation id, or a
stamp made from one) or, outside placeholders, the same kind, ``cNvPr`` id and name -- with
the same place, or the same preset and size, too when the name is a default one
(``Rectangle 3`` for id 4, :func:`default_name`), which any slide built in the same order
gets.  A placeholder's id and
name are what every slide of its layout gets, so they say nothing.  A
slide is a copy when it is on the sample's layout (the same one, or one of the same name when
the sample is in another deck) and at least half of the sample's telling shapes -- its
shapes other than placeholders, and placeholders with a GUID -- answer on it, at least one.
Slides with some answering shapes that miss the bar, and slides whose default-named shapes
match by id and name but none in its place or at its size, are listed (``near``) with why, so a caller can
name them.  A sample with no telling shape (placeholders only) cannot be
told from its layout's other slides: name the copies.

**Matching** is :func:`pptx_agent.edit.compare.match`'s, the sample with one copy at a time:
placeholders by type and idx, then names unique on both, then kind and nearest box.

**What follows** for each matched pair, from the sample: the shape properties (box,
geometry, fill, outline, effects -- a placeholder without a box of its own on the sample
loses the copy's, so it follows its layout again), the shape style, the name, the text
frame (``a:bodyPr``, its autofit's scale the copy's own) and list styles; paragraph and run
formatting, the text staying the copy's own.  When the copy's paragraphs and runs line up
with the sample's (as many paragraphs, each with the same runs, breaks and fields) each
takes the formatting of the one at its position; when they do not, each paragraph takes
that of the sample's first paragraph at its level (a level the sample does not have is left
as it is and reported), and a run keeps what it states differently from the run
formatting most of the copy's text at that level has (a bold word, a coloured one).  A run's language and its
hyperlinks are always the copy's.  Tables: the table style and its options, column widths
and row heights, cell properties (fill, borders, insets, anchor) and cell text formatting.
A copy with more or fewer rows takes the header row from the sample's first row (Header
Row on), a last row from its last (Total Row on) and the others from the sample's body rows
in turn; a copy with other columns keeps its widths and cells (reported).  Pictures: the
box and outline; the image and its crop stay the copy's unless it shows the sample's own
image.  Charts, SmartArt and other frames: the box.  Groups: the group's box, and members
in turn when the members line up, else reported.  The stacking order of the matched shapes
follows the sample's.

**Shapes added or removed.**  A sample's shape no copy has is new: each copy gets it, with
the sample's text.  One that some other copy has but this one does not was taken off this
copy: it is not put back (reported).  A copy's shape the sample has none of was taken off
the sample when at least two of the copies followed have it, and more than half: it goes
(``remove="auto"``); else it is the copy's own and stays (reported).  ``remove="all"``
deletes every shape the sample has none of; ``"none"`` keeps them all.

**What could not follow** is reported per copy with why -- a structure that differs, a
level the sample lacks, a shape the copy filled with something else (a picture where the
sample's placeholder holds text: the box follows) -- and, when pptx2svg is installed, the
text that no longer fits and the collisions that following made (each checked before and
after, only the new ones reported).

The whole follow is one undo step; ``preview=True`` reports the same and changes nothing
(the ids it gives shapes it would add are not handed out again, so a follow after it gives
others).  A copy that already follows is reported with ``changed`` false.
"""

from __future__ import annotations

import copy as _copy
from collections import Counter
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any, Iterable

from lxml import etree

from ..oxml.xml import NAMESPACES, append_in_order, insert_in_order, local_name, qn, remove
from .compare import REFERENCE, kind_of, match
from .errors import ArgumentError
from .ids import cnv_pr

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

EMU_PER_POINT = 12700
_R = "{%s}" % NAMESPACES["r"]
_TEXT_ITEMS = ("r", "br", "fld")
#: A run's attributes that belong to its text, not to the design: always the copy's.
_OWN_ATTRIBUTES = ("lang", "altLang", "dirty", "err", "smtClean", "smtId", "noProof")
#: A run's children that belong to its text: always the copy's.
_OWN_CHILDREN = ("hlinkClick", "hlinkMouseOver")
_FILLS = frozenset(("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"))
#: What a shape-properties child is, as a change is reported.
_SP_PR_PART = {"xfrm": "box", "prstGeom": "geometry", "custGeom": "geometry", "ln": "line",
               "effectLst": "effects", "effectDag": "effects", "scene3d": "effects",
               "sp3d": "effects", **{name: "fill" for name in _FILLS}}
REMOVE = ("auto", "all", "none")


def _points(emu: float) -> float:
    return round(float(emu) / EMU_PER_POINT, 2)


# -- comparing XML across slides ------------------------------------------------------------


def _canon(element, slide: "Slide | None") -> Any:
    """``element`` as a comparable value: tags, attributes and text, a relationship id read
    as what it points at (so the same picture on two slides is the same)."""
    if element is None:
        return None
    rels = slide.document.package.relationships(slide.part_path) if slide is not None else {}

    def walk(node):
        if not isinstance(node.tag, str):
            return None
        attributes = []
        for name, value in sorted(node.attrib.items()):
            if name.startswith(_R):
                rel = rels.get(value)
                value = ("rel", rel.target if rel is not None and rel.is_external
                         else rel.target_part if rel is not None else value)
            attributes.append((name, value))
        children = tuple(item for item in (walk(child) for child in node) if item is not None)
        return (node.tag, tuple(attributes), (node.text or "").strip(), children)

    return walk(element)


def _same(a, a_slide, b, b_slide) -> bool:
    return _canon(a, a_slide) == _canon(b, b_slide)


# -- identity: which shapes a copy kept of its sample ---------------------------------------


def _guid(shape: "Shape") -> str | None:
    """The shape's GUID-based local id (``c<creation id>``, kept as a stamp too), or None."""
    local = shape._id.local
    return local if local.startswith("c") and "-" in local else None


def default_name(name: str | None, raw_id: str | None) -> bool:
    """Whether ``name`` is the kind of name PowerPoint (and pptx-agent) gives a new shape:
    a word or words and the shape's id less one (``Rectangle 3`` for id 4, in any
    interface language) -- a name any slide built the same way gets."""
    if not name or not raw_id or not raw_id.isdigit():
        return False
    head, _, number = name.rpartition(" ")
    return bool(head) and number.isdigit() and int(number) == int(raw_id) - 1


def _preset(shape: "Shape") -> str | None:
    geometry = next(shape._element.iter(qn("a:prstGeom")), None)
    return geometry.get("prst") if geometry is not None else None


def _telling(shape: "Shape", loose: bool = False) -> list[tuple]:
    """What identifies ``shape`` as the copy of another: its GUID, and outside
    placeholders its kind, ``cNvPr`` id and name -- with its place, or its preset and size,
    too when the name is a default one (:func:`default_name`); ``loose``, without the box (which shape of a copy
    came from which of the sample, whatever the copy did to it since)."""
    keys = []
    guid = _guid(shape)
    if guid:
        keys.append(("guid", guid.split("#")[0]))
    if shape.placeholder is None:
        properties = cnv_pr(shape._element)
        raw, name = (properties.get("id"), properties.get("name")) if properties is not None \
            else (None, None)
        if raw:
            if default_name(name, raw) and not loose:
                bounds = shape.bounds
                keys.append(("placed", kind_of(shape), raw, name, bounds))
                if bounds is not None:
                    keys.append(("sized", kind_of(shape), raw, name, _preset(shape),
                                 bounds[2], bounds[3]))
            else:
                keys.append(("id", kind_of(shape), raw, name))
    return keys


def _top(slide: "Slide") -> list["Shape"]:
    return [shape for shape in slide.shapes if shape.parent_group is None]


def _identities(slide: "Slide", loose: bool = False) -> set[tuple]:
    return {key for shape in _top(slide) for key in _telling(shape, loose)}


def _same_layout(sample: "Slide", other: "Slide") -> bool:
    a, b = sample.layout, other.layout
    if a is None or b is None:
        return a is b
    if sample.document is other.document:
        return a.part_path == b.part_path
    return a.name == b.name


@dataclass
class Found:
    """The copies :func:`find_copies` found, and the slides that came near."""

    sample: str
    copies: list[dict] = field(default_factory=list)   #: {"slide", "shapes": "3 of 4"}
    near: list[dict] = field(default_factory=list)     #: {"slide", "shapes", "why"}

    @property
    def slides(self) -> list[str]:
        return [entry["slide"] for entry in self.copies]

    def to_json(self) -> dict[str, Any]:
        data: dict[str, Any] = {"sample": self.sample, "copies": self.copies}
        if self.near:
            data["near"] = self.near
        return data


def _address(slide: "Slide") -> str:
    return f"s:{slide.slide_id}"


def find_copies(sample: "Slide", deck=None) -> Found:
    """The slides of ``deck`` (default the sample's own) made from ``sample`` (the module's
    rule).  Raises ``ArgumentError`` when the sample has no shape that tells its copies
    from other slides of its layout."""
    deck = sample.document if deck is None else deck
    wanted = [_telling(shape) for shape in _top(sample)]
    wanted = [keys for keys in wanted if keys]
    if not wanted:
        raise ArgumentError(
            f"{_address(sample)} holds only placeholders, which every slide of its layout "
            "has too: nothing tells its copies from the others; name the copies")
    loosely = [_telling(shape, True) for shape in _top(sample)]
    loosely = [keys for keys in loosely if keys]
    found = Found(sample=_address(sample))
    for slide in deck.slides:
        if slide.document is sample.document and slide.slide_id == sample.slide_id:
            continue
        have = _identities(slide)
        answered = sum(1 for keys in wanted if any(key in have for key in keys))
        if not answered:
            loose = _identities(slide, True)
            named = sum(1 for keys in loosely if any(key in loose for key in keys))
            if named:
                found.near.append({"slide": _address(slide), "shapes": f"0 of {len(wanted)}",
                                   "why": f"{named} shape(s) with the sample's default names "
                                          "and ids, none in its places or sizes"})
            continue
        count = f"{answered} of {len(wanted)}"
        if not _same_layout(sample, slide):
            found.near.append({"slide": _address(slide), "shapes": count,
                               "why": "on another layout"})
        elif 2 * answered >= len(wanted):
            found.copies.append({"slide": _address(slide), "shapes": count})
        else:
            found.near.append({"slide": _address(slide), "shapes": count,
                               "why": "fewer than half of the sample's shapes"})
    return found


# -- following one copy ---------------------------------------------------------------------


@dataclass
class CopyFollow:
    """What following did, or would do, to one copy."""

    slide: str
    changed: list[dict] = field(default_factory=list)       #: {"shape", "name", "what": [...]}
    added: list[dict] = field(default_factory=list)         #: {"shape", "name", "from"}
    removed: list[dict] = field(default_factory=list)       #: {"shape", "name"}
    not_followed: list[dict] = field(default_factory=list)  #: {"shape", "name", "why"}
    overflows: list[dict] = field(default_factory=list)     #: new problems after following

    @property
    def is_changed(self) -> bool:
        return bool(self.changed or self.added or self.removed)

    def to_json(self, *, brief: bool = False) -> dict[str, Any]:
        data: dict[str, Any] = {"slide": self.slide, "changed": self.is_changed}
        for name in ("changed", "added", "removed"):
            items = getattr(self, name)
            if items:
                data[name if name != "changed" else "shapes"] = len(items) if brief else items
        if self.not_followed:
            data["not_followed"] = self.not_followed
        if self.overflows:
            data["overflows"] = self.overflows
        return data


@dataclass
class Following:
    """What :func:`follow_sample` did (or, with ``preview``, would do)."""

    sample: str
    copies: list[CopyFollow] = field(default_factory=list)
    preview: bool = False
    found: Found | None = None
    checked: bool = True                        #: whether the fit check ran (pptx2svg)
    notes: list[str] = field(default_factory=list)

    def to_json(self, *, brief: bool = False) -> dict[str, Any]:
        data: dict[str, Any] = {"sample": self.sample, "preview": self.preview,
                                "followed": sum(1 for c in self.copies if c.is_changed),
                                "unchanged": sum(1 for c in self.copies if not c.is_changed),
                                "copies": [c.to_json(brief=brief) for c in self.copies]}
        if self.found is not None:
            data["found_by"] = "same layout, at least half of the sample's shapes by id " \
                               "and name"
            if self.found.near:
                data["near"] = self.found.near
        if not self.checked:
            data["fit_check"] = "not run: pptx2svg is not installed"
        if self.notes:
            data["notes"] = self.notes
        return data


class _Pair:
    """One sample shape and its copy's, followed: the changes found on the way."""

    def __init__(self, work: "_Copy", sample: "Shape", target: "Shape") -> None:
        self.work, self.sample, self.target = work, sample, target
        self.what: list[str] = []
        self.notes: list[str] = []

    def changed(self, what: str) -> None:
        if what not in self.what:
            self.what.append(what)

    def not_followed(self, why: str) -> None:
        self.work.report.not_followed.append(
            {"shape": self.target.id, "name": self.target.name, "why": why})


class _Copy:
    def __init__(self, sample: "Slide", slide: "Slide") -> None:
        self.sample, self.slide = sample, slide
        self.report = CopyFollow(_address(slide))

    # pieces of the sample, brought over with what they relate to
    def bring(self, element):
        """A deep copy of the sample's ``element`` for the copy's slide, every relationship
        it uses re-pointed there (a picture fill's image, ...)."""
        clone = _copy.deepcopy(element)
        if any(name.startswith(_R) for node in clone.iter() if isinstance(node.tag, str)
               for name in node.attrib):
            from .copying import _carry_relationships

            holder = etree.Element("holder")
            holder.append(clone)
            _carry_relationships(holder, self.sample.document.package, self.sample.part_path,
                                 self.slide.document.package, self.slide.part_path)
            holder.remove(clone)
        return clone

    def same(self, sample_element, copy_element) -> bool:
        return _same(sample_element, self.sample, copy_element, self.slide)


def _child(element, tag: str):
    return element.find(qn(tag))


def _replace(work: _Copy, parent, tag: str, sample_child) -> bool:
    """Make ``parent``'s ``tag`` child the sample's (none when the sample has none);
    whether it changed."""
    mine = _child(parent, tag)
    if work.same(sample_child, mine):
        return False
    if mine is not None:
        remove(mine)
    if sample_child is not None:
        insert_in_order(parent, work.bring(sample_child))
    return True


def _shape_properties(pair: _Pair, tag: str = "p:spPr") -> None:
    """The sample's shape properties (box, geometry, fill, outline, effects), what changed
    named part by part."""
    work = pair.work
    theirs = _child(pair.sample._element, tag)
    mine = _child(pair.target._element, tag)
    if work.same(theirs, mine):
        return
    def children(side, name):
        return [c for c in (side if side is not None else ()) if isinstance(c.tag, str)
                and local_name(c) == name]

    names = {local_name(c) for side in (theirs, mine) if side is not None for c in side
             if isinstance(c.tag, str)}
    before = len(pair.what)
    for name in sorted(names):
        a, b = children(theirs, name), children(mine, name)
        if len(a) != len(b) or any(not work.same(x, y) for x, y in zip(a, b)):
            pair.changed(_SP_PR_PART.get(name, "shape"))
    if len(pair.what) == before:
        pair.changed("shape")
    new = work.bring(theirs) if theirs is not None else etree.Element(qn(tag))
    if mine is not None:
        mine.addprevious(new)
        remove(mine)
    else:
        insert_in_order(pair.target._element, new)


def _name(pair: _Pair) -> None:
    theirs, mine = cnv_pr(pair.sample._element), cnv_pr(pair.target._element)
    if theirs is None or mine is None or theirs.get("name") == mine.get("name"):
        return
    mine.set("name", theirs.get("name") or "")
    pair.changed("name")


# -- text --------------------------------------------------------------------------------


def _paragraphs(body) -> list:
    return [] if body is None else body.findall(qn("a:p"))


def _items(paragraph) -> list:
    return [child for child in paragraph if isinstance(child.tag, str)
            and local_name(child) in _TEXT_ITEMS]


def _shape_of(paragraphs) -> list[tuple[str, ...]]:
    return [tuple(local_name(item) for item in _items(p)) for p in paragraphs]


def _level(paragraph) -> int:
    properties = paragraph.find(qn("a:pPr"))
    try:
        return int(properties.get("lvl", "0")) if properties is not None else 0
    except ValueError:
        return 0


def _run_properties(item):
    return item.find(qn("a:rPr"))


def _with_own(work: _Copy, design, own, base=None):
    """``design`` (a sample's run properties, or None) for a copy's run whose properties are
    ``own``: the run's language and hyperlinks stay its own, and -- given the copy
    paragraph's ``base``, its most used run properties -- whatever it states differently
    from that base (a bold word) too."""
    new = work.bring(design) if design is not None else etree.Element(qn("a:rPr"))
    new.tag = qn("a:rPr")
    for child in list(new):
        if isinstance(child.tag, str) and local_name(child) in _OWN_CHILDREN:
            new.remove(child)
    for name in _OWN_ATTRIBUTES:
        new.attrib.pop(name, None)
        if own is not None and own.get(name) is not None:
            new.set(name, own.get(name))
    if own is not None:
        for child in own:
            if isinstance(child.tag, str) and local_name(child) in _OWN_CHILDREN:
                insert_in_order(new, _copy.deepcopy(child))
    if base is not None or own is not None:
        mine = own if own is not None else etree.Element(qn("a:rPr"))
        usual = base if base is not None else etree.Element(qn("a:rPr"))
        if not _same(mine, None, usual, None):
            for name in set(mine.attrib) | set(usual.attrib):
                if name in _OWN_ATTRIBUTES or mine.get(name) == usual.get(name):
                    continue
                if mine.get(name) is None:
                    new.attrib.pop(name, None)
                else:
                    new.set(name, mine.get(name))
            names = {local_name(c) for side in (mine, usual) for c in side
                     if isinstance(c.tag, str)} - set(_OWN_CHILDREN)
            for name in names:
                a = [c for c in mine if isinstance(c.tag, str) and local_name(c) == name]
                b = [c for c in usual if isinstance(c.tag, str) and local_name(c) == name]
                if len(a) == len(b) and all(_same(x, None, y, None) for x, y in zip(a, b)):
                    continue
                for old in [c for c in new if isinstance(c.tag, str) and local_name(c) == name]:
                    new.remove(old)
                for item in a:
                    insert_in_order(new, _copy.deepcopy(item))
    return new


def _set_run_properties(item, properties) -> None:
    old = item.find(qn("a:rPr"))
    empty = not len(properties) and not properties.attrib
    if old is not None:
        remove(old)
    if not empty:
        item.insert(0, properties)


def _dominant(paragraph):
    """The run properties most of the paragraph's text is written with."""
    weights: Counter = Counter()
    seen: dict[Any, Any] = {}
    for item in _items(paragraph):
        if local_name(item) == "br":
            continue
        text = "".join(item.itertext())
        properties = _run_properties(item)
        key = _canon(properties, None)
        seen.setdefault(key, properties)
        weights[key] += len(text.strip()) or 0.5
    if not weights:
        return None
    return seen[weights.most_common(1)[0][0]]


def _paragraph_properties(work: _Copy, paragraph, design, level: int | None = None) -> None:
    old = paragraph.find(qn("a:pPr"))
    new = work.bring(design) if design is not None else None
    if new is not None and level is not None:
        if level:
            new.set("lvl", str(level))
        else:
            new.attrib.pop("lvl", None)
    if old is not None:
        remove(old)
    if new is not None:
        paragraph.insert(0, new)


def _end_properties(work: _Copy, paragraph, design) -> None:
    old = paragraph.find(qn("a:endParaRPr"))
    own = old
    if design is None and old is None:
        return
    new = _with_own(work, design, own) if design is not None else None
    if old is not None:
        remove(old)
    if new is not None:
        new.tag = qn("a:endParaRPr")
        append_in_order(paragraph, new)


def _follow_body(pair: _Pair, theirs, mine, *, where: str = "") -> None:
    """The sample's text body formatting on the copy's, the copy's text kept."""
    work = pair.work
    if theirs is None or mine is None:
        return
    before = _canon(mine, work.slide)
    # the frame: the sample's, the autofit's scale the copy's own (it is its text's)
    frame_before = _canon(_child(mine, "a:bodyPr"), work.slide)
    scale = None
    old_frame = _child(mine, "a:bodyPr")
    if old_frame is not None:
        fit = _child(old_frame, "a:normAutofit")
        scale = dict(fit.attrib) if fit is not None else None
    if _replace(work, mine, "a:bodyPr", _child(theirs, "a:bodyPr")):
        frame = _child(mine, "a:bodyPr")
        fit = _child(frame, "a:normAutofit") if frame is not None else None
        if fit is not None:
            fit.attrib.clear()
            fit.attrib.update(scale or {})
    if _canon(_child(mine, "a:bodyPr"), work.slide) != frame_before:
        pair.changed(f"{where}text frame")
    _replace(work, mine, "a:lstStyle", _child(theirs, "a:lstStyle"))
    sample_paragraphs, copy_paragraphs = _paragraphs(theirs), _paragraphs(mine)
    if _shape_of(sample_paragraphs) == _shape_of(copy_paragraphs):
        for s_par, c_par in zip(sample_paragraphs, copy_paragraphs):
            _paragraph_properties(work, c_par, s_par.find(qn("a:pPr")))
            for s_item, c_item in zip(_items(s_par), _items(c_par)):
                _set_run_properties(c_item, _with_own(work, _run_properties(s_item),
                                                      _run_properties(c_item)))
            _end_properties(work, c_par, s_par.find(qn("a:endParaRPr")))
    else:
        by_level: dict[int, Any] = {}
        for s_par in sample_paragraphs:
            if _items(s_par):
                by_level.setdefault(_level(s_par), s_par)
        for s_par in sample_paragraphs:
            by_level.setdefault(_level(s_par), s_par)
        missing = set()
        usual: dict[int, Any] = {}          # the copy's most used run formatting, per level
        for level in {_level(p) for p in copy_paragraphs}:
            holder = etree.Element(qn("a:p"))
            for c_par in copy_paragraphs:
                if _level(c_par) == level:
                    holder.extend(_copy.deepcopy(item) for item in _items(c_par))
            usual[level] = _dominant(holder)
        for c_par in copy_paragraphs:
            level = _level(c_par)
            s_par = by_level.get(level)
            if s_par is None:
                if _items(c_par):
                    missing.add(level)
                continue
            design = _dominant(s_par)
            if design is None:
                design = s_par.find(qn("a:endParaRPr"))
            base = usual.get(level)
            _paragraph_properties(work, c_par, s_par.find(qn("a:pPr")), level)
            for c_item in _items(c_par):
                _set_run_properties(c_item, _with_own(work, design, _run_properties(c_item),
                                                      base))
            _end_properties(work, c_par, s_par.find(qn("a:endParaRPr")) if
                            s_par.find(qn("a:endParaRPr")) is not None else design)
        if _rest_changed(before, _canon(mine, work.slide)):
            pair.notes.append(f"{where}text laid out unlike the sample's "
                              f"({len(copy_paragraphs)} paragraph(s), the sample "
                              f"{len(sample_paragraphs)}): formatting followed by level")
        for level in sorted(missing):
            pair.not_followed(f"{where}paragraphs at level {level + 1}: the sample has "
                              "none at that level; their formatting is the copy's")
    if _rest_changed(before, _canon(mine, work.slide)):
        pair.changed(f"{where}text format")


def _rest_changed(before, after) -> bool:
    """Whether two canonical text bodies differ apart from their frame."""
    def rest(value):
        return tuple(child for child in value[3] if not child[0].endswith("}bodyPr"))
    return rest(before) != rest(after)


def _text(pair: _Pair) -> None:
    theirs = _child(pair.sample._element, "p:txBody")
    mine = _child(pair.target._element, "p:txBody")
    if theirs is None:
        return
    if mine is None:
        if "".join(theirs.itertext()).strip():
            insert_in_order(pair.target._element, pair.work.bring(theirs))
            pair.changed("text")
        return
    _follow_body(pair, theirs, mine)


# -- tables ------------------------------------------------------------------------------


def _table(frame):
    return next(frame.iter(qn("a:tbl")), None)


def _row_map(sample_rows: int, copy_rows: int, first: bool, last: bool) -> list[int]:
    """For each of the copy's rows, the sample's row whose design it takes."""
    if sample_rows == copy_rows:
        return list(range(copy_rows))
    body = list(range(1 if first and sample_rows > 1 else 0,
                      sample_rows - 1 if last and sample_rows > 1 else sample_rows))
    if not body:
        body = list(range(sample_rows))
    out = []
    for index in range(copy_rows):
        if first and index == 0:
            out.append(0)
        elif last and index == copy_rows - 1:
            out.append(sample_rows - 1)
        else:
            out.append(body[(index - (1 if first else 0)) % len(body)])
    return out


def _merges(row) -> list[tuple]:
    return [tuple(cell.get(name) for name in ("gridSpan", "rowSpan", "hMerge", "vMerge"))
            for cell in row.findall(qn("a:tc"))]


def _follow_table(pair: _Pair) -> None:
    work = pair.work
    theirs, mine = _table(pair.sample._element), _table(pair.target._element)
    if theirs is None or mine is None:
        return
    if _replace(work, mine, "a:tblPr", _child(theirs, "a:tblPr")):
        pair.changed("table style")
    s_cols = theirs.findall("a:tblGrid/a:gridCol", NAMESPACES)
    c_cols = mine.findall("a:tblGrid/a:gridCol", NAMESPACES)
    s_rows, c_rows = theirs.findall(qn("a:tr")), mine.findall(qn("a:tr"))
    properties = _child(theirs, "a:tblPr")
    first = properties is not None and properties.get("firstRow") in ("1", "true")
    last = properties is not None and properties.get("lastRow") in ("1", "true")
    if len(s_cols) != len(c_cols):
        pair.not_followed(f"the table has {len(c_cols)} column(s), the sample's "
                          f"{len(s_cols)}: its column widths, row heights and cells are "
                          "the copy's")
        return
    widths = False
    for s_col, c_col in zip(s_cols, c_cols):
        if s_col.get("w") != c_col.get("w"):
            c_col.set("w", s_col.get("w"))
            widths = True
    if widths:
        pair.changed("column widths")
    if len(s_rows) != len(c_rows):
        pair.notes.append(f"the table has {len(c_rows)} rows, the sample's {len(s_rows)}: "
                          "rows past the sample's take its body rows' design in turn")
    mapping = _row_map(len(s_rows), len(c_rows), first, last)
    heights = cells = False
    unmerged = []
    for number, (c_row, source) in enumerate(zip(c_rows, mapping)):
        s_row = s_rows[source]
        if s_row.get("h") != c_row.get("h") and s_row.get("h") is not None:
            c_row.set("h", s_row.get("h"))
            heights = True
        if _merges(s_row) != _merges(c_row):
            unmerged.append(number)
            continue
        for s_cell, c_cell in zip(s_row.findall(qn("a:tc")), c_row.findall(qn("a:tc"))):
            if _replace(work, c_cell, "a:tcPr", _child(s_cell, "a:tcPr")):
                cells = True
            before = len(pair.what)
            _follow_body(pair, _child(s_cell, "a:txBody"), _child(c_cell, "a:txBody"),
                         where="cell ")
            if len(pair.what) != before:
                cells = True
    pair.what = [what for what in pair.what if not what.startswith("cell ")]
    if heights:
        pair.changed("row heights")
    if cells:
        pair.changed("cells")
    if unmerged:
        rows = ", ".join(str(n) for n in unmerged[:10])
        pair.not_followed(f"row(s) {rows} (from 0) merge cells unlike the sample's: their "
                          "cells are the copy's")


def _frame_box(pair: _Pair) -> None:
    """A graphic frame's box from the sample; a table's height its rows' when they differ
    in number from the sample's."""
    work = pair.work
    theirs = _child(pair.sample._element, "p:xfrm")
    mine = _child(pair.target._element, "p:xfrm")
    if theirs is None or work.same(theirs, mine):
        return
    new = work.bring(theirs)
    s_tbl, c_tbl = _table(pair.sample._element), _table(pair.target._element)
    if s_tbl is not None and c_tbl is not None and \
            len(s_tbl.findall(qn("a:tr"))) != len(c_tbl.findall(qn("a:tr"))):
        extent = new.find(qn("a:ext"))
        total = sum(int(row.get("h") or 0) for row in c_tbl.findall(qn("a:tr")))
        if extent is not None and total:
            extent.set("cy", str(total))
    if work.same(new, mine):
        return
    if mine is not None:
        mine.addprevious(new)
        remove(mine)
    else:
        insert_in_order(pair.target._element, new)
    pair.changed("box")


# -- pictures, groups, the pair ------------------------------------------------------------


def _image(element, slide: "Slide"):
    """What the shape's picture shows: the image's bytes (a linked one: its address)."""
    blip = next(element.iter(qn("a:blip")), None)
    if blip is None:
        return None
    rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
    rel = slide.document.package.relationships(slide.part_path).get(rid)
    if rel is None:
        return None
    if rel.is_external:
        return ("link", rel.target)
    return ("data", slide.document.package.read(rel.target_part))


def _picture(pair: _Pair) -> None:
    work = pair.work
    own = _image(pair.target._element, work.slide)
    if own is not None and own == _image(pair.sample._element, work.sample):
        if _replace(work, pair.target._element, "p:blipFill",
                    _child(pair.sample._element, "p:blipFill")):
            pair.changed("picture crop")


def _box_only(pair: _Pair, tag: str) -> None:
    """Only the box: the sample's ``a:xfrm`` on the copy's shape properties."""
    work = pair.work
    s_pr = _child(pair.sample._element, "p:spPr")
    theirs = _child(s_pr, "a:xfrm") if s_pr is not None else None
    if pair.target.kind == "graphic_frame":
        mine_parent = pair.target._element
        mine = _child(mine_parent, "p:xfrm")
        if theirs is None or work.same(theirs, mine):
            return
        new = work.bring(theirs)
        new.tag = qn("p:xfrm")
    else:
        mine_parent = _child(pair.target._element, tag)
        if mine_parent is None:
            mine_parent = etree.Element(qn(tag))
            insert_in_order(pair.target._element, mine_parent)
        mine = _child(mine_parent, "a:xfrm")
        if work.same(theirs, mine):
            return
        new = work.bring(theirs) if theirs is not None else None
    if mine is not None:
        remove(mine)
    if new is not None:
        insert_in_order(mine_parent, new)
    pair.changed("box")


def _members(element) -> list:
    return [child for child in element if isinstance(child.tag, str)
            and local_name(child) in ("sp", "pic", "cxnSp", "grpSp", "graphicFrame")]


_KIND_WORD = {"sp": "shape", "pic": "picture", "cxnSp": "connector", "grpSp": "group",
              "graphicFrame": "frame"}


def _follow_pair(pair: _Pair) -> None:
    sample, target = pair.sample, pair.target
    s_tag, c_tag = local_name(sample._element), local_name(target._element)
    if s_tag != c_tag:
        if sample.placeholder is not None and target.placeholder is not None:
            _box_only(pair, "p:spPr")
            pair.notes.append(f"the copy filled the placeholder with a "
                              f"{_KIND_WORD.get(c_tag, c_tag)} where the sample's holds a "
                              f"{_KIND_WORD.get(s_tag, s_tag)}: only the box followed")
        else:
            pair.not_followed(f"the copy has a {_KIND_WORD.get(c_tag, c_tag)} where the "
                              f"sample has a {_KIND_WORD.get(s_tag, s_tag)}")
        return
    if s_tag == "graphicFrame" and kind_of(sample) != kind_of(target):
        _frame_box(pair)
        pair.not_followed(f"the copy has a {kind_of(target)} where the sample has a "
                          f"{kind_of(sample)}: only the box followed")
        return
    _name(pair)
    if s_tag in ("sp", "cxnSp", "pic"):
        _shape_properties(pair)
        if _replace(pair.work, target._element, "p:style", _child(sample._element, "p:style")):
            pair.changed("style")
        if s_tag == "sp":
            _text(pair)
        elif s_tag == "pic":
            _picture(pair)
    elif s_tag == "graphicFrame":
        _frame_box(pair)
        if kind_of(sample) == "table":
            _follow_table(pair)
    elif s_tag == "grpSp":
        _shape_properties(pair, "p:grpSpPr")
        s_members, c_members = _members(sample._element), _members(target._element)
        if [local_name(m) for m in s_members] != [local_name(m) for m in c_members]:
            pair.not_followed(f"the group holds {len(c_members)} shape(s) unlike the "
                              f"sample's {len(s_members)}: only the group's box followed")
            return
        for s_member, c_member in zip(s_members, c_members):
            inner = _Pair(pair.work, sample._slide._wrap(s_member),
                          target._slide._wrap(c_member))
            _follow_pair(inner)
            for what in inner.what:
                pair.changed(f"members' {what}" if not what.startswith("members") else what)
            pair.notes += inner.notes


# -- the run of copies ------------------------------------------------------------------


def _problems(slides: "list[Slide]") -> dict[int, dict]:
    """Each slide's problems by key, the slides measured in one reading of the deck."""
    from .fit import _Measured, slide_problems

    measured = _Measured.many(slides)
    return {slide.slide_id: {_problem_key(p): p for p in
                             slide_problems(slide, measured=measured[slide.slide_id])}
            for slide in slides}


def _problem_key(problem) -> tuple:
    return (problem.kind, problem.shape, problem.other, problem.detail)


def _problem_json(problem) -> dict:
    entry: dict[str, Any] = {"shape": problem.shape, "kind": problem.kind}
    if problem.kind == "text":
        entry["by"] = _points(problem.amount)
        entry["why"] = "its text no longer fits after following"
    elif problem.kind == "overlap":
        entry["with"] = problem.other
        entry["why"] = "it overlaps another shape after following"
    else:
        entry["by"] = _points(problem.amount)
        entry["why"] = "it runs past the slide's edge after following"
    return entry


def _renderer() -> bool:
    try:
        import pptx2svg  # noqa: F401
    except ImportError:  # pragma: no cover - the extras-less install
        return False
    return True


class _Preview(Exception):
    """Raised inside the savepoint to take a preview's changes back."""


def follow_sample(sample: "Slide", copies: "Iterable[Slide] | None" = None, *, deck=None,
                  preview: bool = False, remove: str = "auto", check: bool = True
                  ) -> Following:
    """Make ``copies`` (default: :func:`find_copies` in ``deck``, the sample's own deck
    unless given) follow ``sample``'s design, keeping each copy's text, values and
    pictures (the module's rules).  ``remove``: which shapes the sample has none of go --
    ``"auto"`` those most of the copies have, ``"all"``, ``"none"``.  One undo step;
    ``preview`` reports and changes nothing; ``check`` runs the fit check (pptx2svg) and
    reports what following made overflow or collide.  Raises ``ArgumentError`` for an
    unknown ``remove``, no copies, or a sample without telling shapes when none are named."""
    if remove not in REMOVE:
        raise ArgumentError(f"remove is one of {', '.join(REMOVE)}, not {remove!r}")
    found = None
    if copies is None:
        found = find_copies(sample, deck)
        target_deck = sample.document if deck is None else deck
        chosen = [target_deck.slide(address) for address in found.slides]
        if not chosen:
            near = ", ".join(f"{e['slide']} ({e['shapes']}, {e['why']})" for e in found.near[:20])
            raise ArgumentError(f"no slide is a copy of {_address(sample)}: none is on its "
                                "layout with at least half of its shapes by id and name"
                                + (f"; near: {near} -- name them as copies" if near else
                                   "; name the copies"))
    else:
        chosen = []
        for slide in copies:
            if slide.document is sample.document and slide.slide_id == sample.slide_id:
                continue
            if all(slide is not other for other in chosen):
                chosen.append(slide)
        if not chosen:
            raise ArgumentError(f"no copies to follow {_address(sample)}: name slides other "
                                "than the sample")
    documents = {id(slide.document): slide.document for slide in chosen}
    if len(documents) != 1:
        raise ArgumentError("the copies are in one deck; follow each deck's in its own call")
    document = next(iter(documents.values()))
    result = Following(sample=_address(sample), preview=preview, found=found)
    if sample.document is not document:
        result.sample = f"{_address(sample)} of the sample deck"
    check = check and _renderer()
    result.checked = check
    identities = {slide.slide_id: _identities(slide, True) for slide in chosen}
    with document.batch():
        document.history.checkpoint()
        try:
            with document.history.savepoint():
                before = _problems(chosen) if check else {}
                for slide in chosen:
                    result.copies.append(_follow_one(sample, slide, identities, remove))
                if check:
                    slides = [document.slide(slide.slide_id) for slide in chosen]
                    after = _problems(slides)
                    for report, slide in zip(result.copies, slides):
                        old = before.get(slide.slide_id, {})
                        report.overflows = [_problem_json(p) for key, p in
                                            after[slide.slide_id].items() if key not in old]
                if preview:
                    raise _Preview()
        except _Preview:
            document._reset_caches()
    return result


def _follow_one(sample: "Slide", slide: "Slide", identities: dict, remove_rule: str
                ) -> CopyFollow:
    work = _Copy(sample, slide)
    report = work.report
    slide._watch()
    groups = match([sample, slide],
                   lambda s: REFERENCE if s is sample else s.slide_id)
    others = [ids for sid, ids in identities.items() if sid != slide.slide_id]
    followed_count = len(identities)
    order: list[tuple] = []          # (sample z, copy element) of matched shapes
    sample_z = {id(shape._element): z for z, shape in enumerate(_top(sample))}
    adds: list["Shape"] = []
    removes: list["Shape"] = []
    for group in groups:
        theirs = group.shapes.get(REFERENCE)
        mine = group.shapes.get(slide.slide_id)
        if theirs is not None and mine is not None:
            pair = _Pair(work, theirs, mine)
            address, name = mine.id, mine.name
            _follow_pair(pair)
            if pair.what:
                entry: dict[str, Any] = {"shape": address, "name": name, "what": pair.what}
                if pair.notes:
                    entry["note"] = "; ".join(pair.notes)
                report.changed.append(entry)
            elif pair.notes:
                report.not_followed.append({"shape": address, "name": name,
                                            "why": "; ".join(pair.notes)})
            order.append((sample_z.get(id(theirs._element), 0), mine._element))
        elif theirs is not None:
            keys = _telling(theirs, True)
            if keys and any(key in ids for ids in others for key in keys):
                report.not_followed.append({
                    "shape": theirs.id, "name": theirs.name,
                    "why": "the sample's shape is on other copies but not this one: taken "
                           "off here, so not put back"})
            else:
                adds.append(theirs)
        elif mine is not None:
            keys = _telling(mine, True)
            holders = 1 + sum(1 for ids in others if any(key in ids for key in keys))
            if remove_rule == "all" or (remove_rule == "auto" and keys and holders >= 2
                                        and 2 * holders > followed_count):
                removes.append(mine)
            else:
                why = ("the sample has none: the copy's own, kept"
                       + (" (remove: \"all\" deletes it)" if remove_rule == "auto" else ""))
                report.not_followed.append({"shape": mine.id, "name": mine.name, "why": why})
    _restack(slide, order)
    for shape in removes:
        report.removed.append({"shape": shape.id, "name": shape.name})
    for shape in removes:
        _delete(slide, shape)
    for shape in adds:
        report.added.append(_add(work, shape))
    slide._invalidate()
    slide._settle()
    if report.changed or report.added or report.removed:
        slide._touch()
    return report


def _restack(slide: "Slide", order: list[tuple]) -> None:
    """The matched shapes in the sample's stacking order, in the places they hold now."""
    if len(order) < 2:
        return
    elements = [element for _, element in order]
    tree = slide._sp_tree()
    slots = [child for child in tree if any(child is element for element in elements)]
    wanted = [element for _, element in sorted(order, key=lambda item: item[0])]
    if all(a is b for a, b in zip(slots, wanted)):
        return
    markers = []
    for slot in slots:
        marker = etree.Element("marker")
        slot.addprevious(marker)
        markers.append(marker)
    for element in wanted:
        tree.remove(element)
    for marker, element in zip(markers, wanted):
        marker.addprevious(element)
        tree.remove(marker)
    slide._invalidate()


def _delete(slide: "Slide", shape: "Shape") -> None:
    from . import creating as _creating

    gone = _creating.raw_ids_within(shape._element)
    remove(shape._element)
    slide._detach(gone)
    slide._invalidate()


def _add(work: _Copy, shape: "Shape") -> dict:
    """The sample's ``shape`` copied onto the copy, at the stacking place after the copy's
    shape that answers the one below it on the sample."""
    from .copying import _drop_placeholder, _write_frame
    from .document import _strip_identity, renumber_copy

    slide = work.slide
    clone = work.bring(shape._element)
    _strip_identity(clone)
    renumber_copy(clone, slide._next_shape_ids)
    for glue in [*clone.iter(qn("a:stCxn")), *clone.iter(qn("a:endCxn"))]:
        remove(glue)
    if shape.placeholder is not None and not _same_layout(work.sample, slide):
        box = shape.bounds
        _drop_placeholder(clone)
        if box is not None:
            _write_frame(clone, box)
    tree = slide._sp_tree()
    below = None
    names = {}
    for child in tree:
        properties = cnv_pr(child) if isinstance(child.tag, str) else None
        if properties is not None:
            names[properties.get("name")] = child
    for previous in _top(work.sample):
        if previous._element is shape._element:
            break
        if previous.name in names:
            below = names[previous.name]
    if below is not None:
        below.addnext(clone)
    else:
        anchor = tree.find(qn("p:grpSpPr"))
        if anchor is not None:
            anchor.addnext(clone)
        else:
            append_in_order(tree, clone)
    slide._invalidate()
    made = slide._wrap(clone)
    return {"shape": made.id, "name": made.name, "from": shape.id}


__all__ = ["CopyFollow", "Following", "Found", "REMOVE", "default_name", "find_copies",
           "follow_sample"]
