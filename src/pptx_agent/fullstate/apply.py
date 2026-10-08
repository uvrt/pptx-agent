"""Reading a full-state SVG back into the open document.

The rules, in the order they are applied (all of it one undo step):

1. **Refuse what is not ours.**  The root must be an ``<svg>`` carrying this library's
   vocabulary marker, and the slide it names must exist, with the same part, layout, master,
   theme and size.  Parsing follows :mod:`.safe`: no DTDs or entities, bounded sizes, strict
   base64, validated JSON.  Nothing is changed until the whole SVG has been read and checked.
2. **Deletions** (only with ``delete_missing=True``): a shape on the slide whose id appears
   nowhere in the SVG is deleted, with E2's relationship reaping.  The default is to leave
   such shapes alone, so an SVG cut down to the shapes of interest is safe to apply.
3. **The raw floor.**  A shape whose ``data-ooxml-xml`` differs from its XML in the document
   is replaced by it, relationships re-pointed (see below).  A group's XML is its shell --
   its children are their own groups -- so replacing it keeps the children.
4. **Additions** (``add_new=True``, the default): a ``data-ooxml-xml`` whose ``data-pptx-id``
   the slide does not know becomes a new shape, at the place the SVG puts it: inside the
   group it is nested in, after the shape it follows.  Its ``cNvPr@id`` is renumbered if
   taken, and charts, diagrams and other per-slide parts it refers to are copied, as when a
   slide is duplicated; media is shared.
5. **Typed edits.**  A typed attribute that differs from the document is applied through the
   E1/E2 API -- ``data-ooxml-fill-scheme`` through :meth:`Shape.set_solid_fill`, a run's text
   through :meth:`TextFrame.set_text`, a table's rows through :meth:`Table.insert_row` --
   so the result is exactly what that edit would have written.  An absent attribute means
   "no opinion"; ``inherit`` is how to say "remove the explicit value".  Typed edits are
   compared with the document *before* step 3, so changing only a shape's raw XML is not
   undone by the (now stale) typed attributes travelling with it.

**Relationships.**  A fragment's ``r:id``s are resolved through the shape's ``data-ooxml-rels``
description: an id that still means the same target in the slide is kept; otherwise the
slide's existing relationship to that target is used, or a new one is made.  Internal targets
must already be parts of the package -- the SVG refers to media and charts, it does not
carry them -- and relationships to package structure (layouts, masters, notes, themes) are
refused outright.
"""

from __future__ import annotations

import copy
import difflib
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Any

from ..oxml.package import (
    REL_NOTES_MASTER,
    REL_NOTES_SLIDE,
    REL_SLIDE_LAYOUT,
    REL_SLIDE_MASTER,
    SHARED_ON_DUPLICATE,
    Relationship,
    normalize_part_path,
)
from ..oxml.xml import (
    CREATION_ID_EXT_URI,
    NAMESPACES,
    Element,
    append_in_order,
    local_name,
    qn,
    remove,
)
from ..edit import creating as _creating
from ..edit.fill import Arrowhead
from ..edit.ids import cnv_pr, read_stamp
from ..edit.units import Pt
from . import model as m
from .emit import SVG_NS, fragment_bytes, resolver_for, shape_ids, slide_attributes
from .safe import (
    DEFAULT_LIMITS,
    FullStateError,
    Limits,
    NotFullStateSvg,
    decode_base64,
    load_json,
    parse_untrusted,
)

if TYPE_CHECKING:  # pragma: no cover
    from ..edit.document import Document, Shape, Slide

_SHAPE_TAGS = frozenset(qn(f"p:{name}") for name in m.SHAPE_TAG_NAMES)
#: Relationship types a shape never has: they are the package's skeleton.
_STRUCTURAL = frozenset({
    REL_SLIDE_LAYOUT, REL_SLIDE_MASTER, REL_NOTES_SLIDE, REL_NOTES_MASTER,
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
})
#: Slide-level attributes that must match: a mismatch means another deck, or another state.
_SLIDE_CHECKED = ("slide-part", "layout", "master", "theme", "slide-cx", "slide-cy")
#: Read-only shape attributes: they describe, and changing them is refused rather than lost.
_READ_ONLY = ("kind", "cnvpr-id", "ph-type", "ph-idx", "ch-x", "ch-y", "ch-cx", "ch-cy",
              "graphic", "chart", "diagram")


@dataclass
class ApplyReport:
    """What :meth:`Document.apply_svg` did.  Falsy when the SVG matched the document.

    For example::

        report = deck.apply_svg(svg)
        if report: print(report.edited, report.added, report.deleted)
    """

    slide_id: int
    #: Shapes whose raw XML was replaced from ``data-ooxml-xml``.
    replaced: list[str] = field(default_factory=list)
    #: Shapes a typed attribute was applied to, with the features applied.
    edited: dict[str, list[str]] = field(default_factory=dict)
    #: SVG id (or ``#<n>`` for a group without one) -> the new shape's id.
    added: dict[str, str] = field(default_factory=dict)
    deleted: list[str] = field(default_factory=list)
    #: Groups that named no shape on the slide and carried no XML to create one from.
    ignored: list[str] = field(default_factory=list)

    def __bool__(self) -> bool:
        return bool(self.replaced or self.edited or self.added or self.deleted)


@dataclass
class _Record:
    """One shape group read from the SVG."""

    key: str
    svg_id: str | None
    attributes: dict[str, str]
    parent: "_Record | None"
    order: int
    fragment: Element | None = None
    fragment_bytes: bytes | None = None
    rels: dict[str, dict] = field(default_factory=dict)
    #: The document element this record names (before the edit) or became.
    element: Element | None = None
    is_new: bool = False


# ------------------------------------------------------------------------------------------
# Reading the SVG
# ------------------------------------------------------------------------------------------


def read_svg(svg: str | bytes, limits: Limits) -> tuple[dict[str, str], list[_Record]]:
    root = parse_untrusted(svg, limit=limits.svg_bytes, what="the SVG")
    if root.tag != "{%s}svg" % SVG_NS:
        raise NotFullStateSvg("not an SVG document")
    vocabulary = root.get(m.PREFIX + "vocabulary")
    if vocabulary is None:
        raise NotFullStateSvg("this SVG carries no full-state vocabulary; only SVG written by "
                              "Slide.render_svg(full_state=True) can be applied")
    if vocabulary != m.VOCABULARY:
        raise NotFullStateSvg(f"unsupported vocabulary {vocabulary!r}; expected {m.VOCABULARY!r}")
    slide = {key[len(m.PREFIX):]: value for key, value in root.attrib.items()
             if key.startswith(m.PREFIX)}

    records: list[_Record] = []
    # Keyed by the element itself, not id(): lxml proxies that are not kept alive are
    # recreated, and their ids reused.
    by_element: dict[Element, _Record] = {}
    seen: set[str] = set()
    for node in root.iter():
        if not isinstance(node.tag, str) or node is root:
            continue
        svg_id = node.get("data-pptx-id")
        raw = node.get(m.PREFIX + "xml")
        if svg_id is None and raw is None:
            continue
        if svg_id is not None and svg_id.startswith(("lay:", "mst:")):
            continue  # layout and master decoration: drawn, not editable on the slide
        if svg_id is not None and "/" in svg_id and raw is None:
            continue  # a drawn piece of a shape (SmartArt's cached shapes), not a shape
        if len(records) >= limits.shapes:
            raise FullStateError(f"the SVG has more than {limits.shapes} shapes")
        if svg_id is not None:
            if svg_id in seen:
                raise FullStateError(f"data-pptx-id {svg_id!r} appears twice; give a new "
                                     f"shape no id (or a new one)")
            seen.add(svg_id)
        parent = None
        for ancestor in node.iterancestors():
            parent = by_element.get(ancestor)
            if parent is not None:
                break
        attributes = {key[len(m.PREFIX):]: value for key, value in node.attrib.items()
                      if key.startswith(m.PREFIX)}
        key = svg_id if svg_id is not None else f"#{len(records)}"
        record = _Record(key, svg_id, attributes, parent, len(records))
        if raw is not None:
            data = decode_base64(raw, limit=limits.xml_attribute, what=f"{key} data-ooxml-xml")
            element = parse_untrusted(data, limit=limits.xml_attribute,
                                      what=f"{key} data-ooxml-xml")
            if element.tag not in _SHAPE_TAGS:
                raise FullStateError(f"{key}: data-ooxml-xml holds a {element.tag}, not a shape")
            if local_name(element) == "grpSp" and any(c.tag in _SHAPE_TAGS for c in element):
                raise FullStateError(f"{key}: a group's data-ooxml-xml is its shell; its "
                                     f"children are SVG groups of their own")
            record.fragment, record.fragment_bytes = element, data
        if "rels" in attributes:
            record.rels = _read_rels(attributes["rels"], key, limits)
        records.append(record)
        by_element[node] = record
    return slide, records


def _read_rels(value: str, key: str, limits: Limits) -> dict[str, dict]:
    entries = load_json(value, limit=limits.json_attribute, what=f"{key} data-ooxml-rels",
                        kind=list)
    result = {}
    for entry in entries:
        if not isinstance(entry, dict) or set(entry) - {"id", "type", "target", "external"} \
                or not all(isinstance(entry.get(k), str) for k in ("id", "type", "target")):
            raise FullStateError(f"{key}: a relationship is {{id, type, target, external?}}")
        if not isinstance(entry.get("external", False), bool):
            raise FullStateError(f"{key}: a relationship's 'external' is true or false")
        if entry["type"] in _STRUCTURAL:
            raise FullStateError(f"{key}: a shape cannot relate to {entry['type']}")
        result[entry["id"]] = entry
    return result


def _fragment_resolver(record: _Record, slide: "Slide") -> m.Resolver:
    """Resolve a fragment's ids the way its SVG describes them, else through the slide."""
    relationships = dict(slide.document.package.relationships(slide.part_path))
    for rel_id, entry in record.rels.items():
        external = entry.get("external", False)
        relationships[rel_id] = Relationship(
            rel_id, entry["type"], entry["target"],
            None if external else normalize_part_path(entry["target"]), external)
    return m.Resolver(relationships, {s.part_path: s.slide_id for s in slide.document.slides},
                      slide.document.package, shape_ids(slide))


# ------------------------------------------------------------------------------------------
# Comparing
# ------------------------------------------------------------------------------------------


def _tolerant(canonical, value: Any, what: str) -> Any:
    """A document's value in canonical form -- or as it is, if the file is unusual."""
    try:
        return canonical(value, what)
    except FullStateError:
        return value


def _feature(attributes: dict[str, str], feature: str, what: str, *,
             trusted: bool, limits: Limits = DEFAULT_LIMITS) -> Any:
    """One editable feature's canonical value, or ``None`` when the attributes lack it."""
    if feature in {"fill", "line"}:
        model = m.decode_attributes(feature, attributes)
        if model is None:
            return None
        canonical = m.canonical_fill if feature == "fill" else m.canonical_line
        return _tolerant(canonical, model, what) if trusted else canonical(model, what)
    if feature not in attributes:
        return None
    value = attributes[feature]
    if feature in {"text", "table", "chart-data", "diagram-nodes"}:
        loaded = load_json(value, limit=limits.json_attribute,
                           what=f"{what} data-ooxml-{feature}", kind=dict)
        canonical = {
            "text": m.canonical_text,
            "table": m.canonical_table,
            "chart-data": lambda v, w: m.canonical_chart(v, w, limits.chart_points),
            "diagram-nodes": lambda v, w: m.canonical_diagram(v, w, limits.diagram_nodes),
        }[feature]
        return _tolerant(canonical, loaded, what) if trusted else canonical(loaded, what)
    if feature == "image":
        return normalize_part_path(value)
    if feature == "adj":
        if trusted:
            return _tolerant(m.canonical_adjustments, value, what)
        return m.canonical_adjustments(value, what, limits.adjustments)
    if feature in m.CONNECTIONS:
        if trusted:
            return _tolerant(m.canonical_connection, value, f"{what} {feature}")
        return m.canonical_connection(value, f"{what} {feature}")
    if trusted:
        return _tolerant(lambda v, w: m.canonical_geometry(feature, v, w), value, what)
    return m.canonical_geometry(feature, value, what)


def _typed_changes(record: _Record, base: dict[str, str], limits: Limits) -> dict[str, Any]:
    """Editable features whose SVG value differs from ``base`` (the document's)."""
    changes = {}
    for feature in m.EDITABLE:
        wanted = _feature(record.attributes, feature, record.key, trusted=False, limits=limits)
        if wanted is None:
            continue
        current = _feature(base, feature, record.key, trusted=True)
        if feature == "adj":
            if "geom" in changes and record.attributes.get("adj") == base.get("adj"):
                continue  # the old preset's values, left as they were: no opinion
            preset = changes["geom"][0] if "geom" in changes else base.get("geom")
            wanted_values = m.effective_adjustments(preset, wanted, record.key)
            try:
                current_values = m.effective_adjustments(base.get("geom"), current or (),
                                                         record.key)
            except FullStateError:
                current_values = None
            if "geom" not in changes and wanted_values == current_values:
                continue
            changes[feature] = (wanted, current)
            continue
        if feature in m.CONNECTIONS and base.get("kind") != "connector":
            raise FullStateError(f"{record.key}: data-ooxml-{feature} is a connector's; a "
                                 f"{base.get('kind')} has no connections")
        if wanted != current or (record.is_new and feature in m.CONNECTIONS):
            # A new connector's ends are always attached afresh: the shapes they name may
            # have been renumbered on the way in.
            changes[feature] = (wanted, current)
    return changes


def _check_read_only(record: _Record, *bases: dict[str, str]) -> None:
    for key in _READ_ONLY:
        if key not in record.attributes:
            continue
        if not any(base.get(key) == record.attributes[key] for base in bases):
            raise FullStateError(f"{record.key}: data-ooxml-{key} is read-only (it was "
                                 f"{bases[0].get(key)!r}); edit data-ooxml-xml instead")


# ------------------------------------------------------------------------------------------
# Applying
# ------------------------------------------------------------------------------------------


def apply_full_state(document: "Document", svg: str | bytes, *, add_new: bool = True,
                     delete_missing: bool = False, limits: Limits = DEFAULT_LIMITS
                     ) -> ApplyReport:
    slide_attrs, records = read_svg(svg, limits)
    try:
        slide_id = int(slide_attrs.get("slide-id", ""))
    except ValueError:
        raise FullStateError("the SVG names no slide (data-ooxml-slide-id)") from None
    try:
        slide = document.slide(slide_id)
    except KeyError:
        raise FullStateError(f"the document has no slide {slide_id}") from None
    expected = slide_attributes(slide)
    for key in _SLIDE_CHECKED:
        if key in slide_attrs and slide_attrs[key] != expected.get(key):
            raise FullStateError(f"data-ooxml-{key} is {slide_attrs[key]!r} but the document's "
                                 f"is {expected.get(key)!r}: the SVG describes another deck or "
                                 f"another state of it")

    report = ApplyReport(slide_id)
    resolver = resolver_for(slide)
    index = slide._ensure_index()
    shapes = {str(shape_id): element for shape_id, element in index}
    for record in records:
        if record.svg_id is not None and record.svg_id in shapes:
            record.element = shapes[record.svg_id]
        elif record.fragment is not None:
            record.is_new = True
        else:
            report.ignored.append(record.key)

    # -- plan, reading only ------------------------------------------------------------------
    plan: list[tuple[_Record, bool, dict[str, Any]]] = []
    for record in records:
        if record.element is not None:
            base = m.shape_attributes(record.element, resolver)
            raw_changed = (record.fragment_bytes is not None
                           and record.fragment_bytes != fragment_bytes(record.element))
            raw_base = (m.shape_attributes(record.fragment, _fragment_resolver(record, slide))
                        if raw_changed else base)
            _check_read_only(record, base, raw_base)
            if raw_changed and m.KINDS[local_name(record.fragment)] != base["kind"]:
                raise FullStateError(f"{record.key}: data-ooxml-xml turns a {base['kind']} "
                                     f"into something else; delete it and add a new shape")
            changes = _typed_changes(record, base, limits)
        elif record.is_new:
            if not add_new:
                raise FullStateError(f"{record.key} is not a shape on slide {slide_id}; pass "
                                     f"add_new=True to add it")
            if record.parent is not None and record.parent.fragment is not None \
                    and local_name(record.parent.fragment) != "grpSp":
                raise FullStateError(f"{record.key} is nested in {record.parent.key}, "
                                     f"which is not a group")
            base = m.shape_attributes(record.fragment, _fragment_resolver(record, slide))
            _check_read_only(record, base)
            raw_changed, changes = True, _typed_changes(record, base, limits)
        else:
            continue
        if raw_changed or changes:
            plan.append((record, raw_changed, changes))

    known = {record.svg_id for record in records
             if record.svg_id is not None and (record.element is not None or record.is_new)}
    for record, _, changes in plan:
        for feature in m.CONNECTIONS:
            if feature in changes and changes[feature][0] != "none":
                target = changes[feature][0].partition(" ")[0]
                if target not in known and target not in shapes:
                    raise FullStateError(f"{record.key} data-ooxml-{feature}: no shape "
                                         f"{target!r} on slide {slide_id} or in the SVG")
                if target == record.svg_id:
                    raise FullStateError(f"{record.key} data-ooxml-{feature}: a connector "
                                         f"cannot attach to itself")

    present = {record.svg_id for record in records if record.svg_id is not None}
    doomed = []
    if delete_missing:
        for shape_id, element in index:
            if str(shape_id) in present:
                continue
            ancestors = list(element.iterancestors(qn("p:grpSp")))
            if any(index.id_of(a) is not None and str(index.id_of(a)) not in present
                   for a in ancestors):
                continue  # goes with the group being deleted
            kept = [e for e in element.iter(*_SHAPE_TAGS)
                    if e is not element and str(index.id_of(e)) in present]
            if kept:
                raise FullStateError(f"{shape_id} is missing from the SVG but its child "
                                     f"{index.id_of(kept[0])} is not; ungroup it instead")
            doomed.append(str(shape_id))

    if not plan and not doomed:
        return report

    # -- apply, as one undo step -------------------------------------------------------------
    with document.batch():
        document.history.checkpoint()
        for identifier in doomed:
            slide.shape(identifier).delete()  # reaps what only it used
            report.deleted.append(identifier)
        # Relationship ids the raw XML stops using are released after it is in place.
        slide._watch()
        moved: set[str] = set()          # cNvPr ids of shapes whose raw geometry changed
        rerouted: list[Element] = []     # connectors to route once everything is in place
        for record, raw_changed, _ in plan:
            if raw_changed and not record.is_new:
                before = _geometry_key(record.element)
                _replace(record, slide)
                report.replaced.append(record.key)
                if _geometry_key(record.element) != before:
                    if local_name(record.element) == "cxnSp":
                        rerouted.append(record.element)
                    else:
                        moved |= _creating.raw_ids_within(record.element)
        for record, _, _ in plan:
            if record.is_new:
                _add(record, records, slide)
                if local_name(record.element) == "cxnSp":
                    rerouted.append(record.element)
        slide._invalidate()
        slide._settle()
        connections = []
        for record, _, changes in plan:
            if record.is_new:
                report.added[record.key] = slide._wrap(record.element).id
            if not changes:
                continue
            shape = slide._wrap(record.element)
            for feature, (wanted, current) in changes.items():
                if feature in m.CONNECTIONS:
                    continue
                _apply_feature(shape, feature, wanted, current)
            ends = {feature: changes[feature][0] for feature in m.CONNECTIONS
                    if feature in changes}
            if ends:
                connections.append((record, ends))
            report.edited[shape.id] = list(changes)
        # Connections, once every shape is where the SVG puts it; then the routes.
        by_svg_id = {record.svg_id: record.element for record in records
                     if record.svg_id is not None and record.element is not None}
        for record, ends in connections:
            _apply_connections(slide, record, ends, by_svg_id)
        if moved:
            rerouted.extend(_creating.attached_connectors(slide, moved))
        for element in dict.fromkeys(rerouted):
            _reroute(slide, element)
    return report


def _geometry_key(element: Element | None) -> bytes:
    """What a connector attached to the shape depends on: frame, geometry, attachments."""
    from lxml import etree

    if element is None:
        return b""
    parts = []
    for path in ("p:spPr/a:xfrm", "p:spPr/a:prstGeom", "p:spPr/a:custGeom",
                 "p:grpSpPr/a:xfrm", "p:xfrm", "p:nvCxnSpPr/p:cNvCxnSpPr"):
        node = element.find("/".join(qn(step) for step in path.split("/")))
        parts.append(b"" if node is None else etree.tostring(node, method="c14n"))
    return b"|".join(parts)


def _apply_connections(slide: "Slide", record: _Record, ends: dict[str, str],
                       by_svg_id: dict[str, Element]) -> None:
    shape = slide._wrap(record.element)
    what = f"{shape.id} data-ooxml-cxn"
    begin = end = None
    free = []
    for feature, value in ends.items():
        which = "begin" if feature == "cxn-begin" else "end"
        if value == "none":
            free.append(which)
            continue
        target_id, _, site = value.partition(" ")
        element = by_svg_id.get(target_id)
        target = slide._wrap(element) if element is not None else slide.shape(target_id)
        if which == "begin":
            begin = (target, int(site))
        else:
            end = (target, int(site))
    try:
        if free:
            shape.disconnect("both" if len(free) == 2 else free[0])
        if begin is not None or end is not None:
            shape.connect(begin=begin, end=end)
    except (ValueError, IndexError, TypeError) as error:
        raise FullStateError(f"{what}: {error}") from None


def _reroute(slide: "Slide", element: Element) -> None:
    """Route a connector from its attachments, if it has any and is a kind that routes."""
    if element.getparent() is None or _creating.connector_kind(element) is None:
        return
    if all(_creating.connection(element, which) is None for which in ("begin", "end")):
        return
    slide._wrap(element).reroute()


# -- the raw floor -------------------------------------------------------------------------


def _replace(record: _Record, slide: "Slide") -> None:
    new = _prepared(record, slide, new_shape=False)
    old = record.element
    assert old is not None
    if local_name(old) == "grpSp":
        _replace_shell(old, new)
    else:
        new.tail = old.tail
        old.getparent().replace(old, new)
        record.element = new
    slide._touch()


def _replace_shell(group: Element, shell: Element) -> None:
    """Swap a group's own XML for ``shell``'s, keeping its child shapes where they are."""
    group.attrib.clear()
    group.attrib.update(shell.attrib)
    current = [c for c in group if c.tag not in _SHAPE_TAGS]
    incoming = list(shell)
    if [c.tag for c in current] == [c.tag for c in incoming]:
        for old, new in zip(current, incoming):
            new.tail = old.tail
            group.replace(old, new)
        return
    for old in current:
        remove(old)
    for new in incoming:
        new.tail = None
        append_in_order(group, new)


def _add(record: _Record, records: list[_Record], slide: "Slide") -> None:
    new = _prepared(record, slide, new_shape=True)
    container = slide._sp_tree() if record.parent is None else record.parent.element
    if container is None or local_name(container) not in {"spTree", "grpSp"}:
        raise FullStateError(f"{record.key}: its parent {record.parent and record.parent.key} "
                             f"is not a group on this slide")
    _fresh_identity(new, slide)
    siblings = [r for r in records[:record.order] if r.parent is record.parent
                and r.element is not None and r.element.getparent() is container]
    if siblings:
        previous = siblings[-1].element
        new.tail = previous.tail
        previous.addnext(new)
    else:
        first = next((c for c in container if c.tag in _SHAPE_TAGS), None)
        if first is not None:
            new.tail = first.getprevious().tail if first.getprevious() is not None else None
            first.addprevious(new)
        else:
            append_in_order(container, new)
    record.element = new
    slide._invalidate()


def _fresh_identity(element: Element, slide: "Slide") -> None:
    """Renumber a new shape whose ``cNvPr@id`` is taken; drop a copied durable id."""
    properties = cnv_pr(element)
    if properties is None:
        return
    taken = {node.get("id") for node in slide._sp_tree().iter(qn("p:cNvPr"))}
    if properties.get("id") in taken or not (properties.get("id") or "").isdigit():
        properties.set("id", str(slide._next_shape_id()))
    identities = set()
    for node in slide._sp_tree().iter(qn("p:cNvPr")):
        identities.add(read_stamp(node))
        identities.add(_creation(node))
    if read_stamp(properties) in identities - {None} or _creation(properties) in identities - {None}:
        from ..edit.document import _strip_identity

        _strip_identity(element)


def _creation(properties: Element) -> str | None:
    ext_list = properties.find(qn("a:extLst"))
    for extension in [] if ext_list is None else ext_list.findall(qn("a:ext")):
        if extension.get("uri") == CREATION_ID_EXT_URI:
            node = extension.find(qn("a16:creationId"))
            return None if node is None else node.get("id")
    return None


def _prepared(record: _Record, slide: "Slide", *, new_shape: bool) -> Element:
    """A copy of the record's fragment with its relationship ids valid in ``slide``."""
    assert record.fragment is not None
    element = copy.deepcopy(record.fragment)
    namespace = "{%s}" % NAMESPACES["r"]
    mapping: dict[str, str] = {}
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        for name, value in list(node.attrib.items()):
            if not name.startswith(namespace) or not value:
                continue
            if value not in mapping:
                mapping[value] = _relationship_for(record, value, slide, new_shape)
            node.set(name, mapping[value])
    return element


def _relationship_for(record: _Record, rel_id: str, slide: "Slide", new_shape: bool) -> str:
    package = slide.document.package
    part = slide.part_path
    current = package.relationships(part).get(rel_id)
    entry = record.rels.get(rel_id)
    if entry is None:
        if current is None:
            raise FullStateError(f"{record.key}: data-ooxml-xml uses {rel_id}, which neither "
                                 f"data-ooxml-rels nor the slide defines")
        entry = {"id": rel_id, "type": current.type, "external": current.is_external,
                 "target": current.target if current.is_external else current.target_part}
    rel_type, external = entry["type"], entry.get("external", False)
    if rel_type in _STRUCTURAL:
        raise FullStateError(f"{record.key}: a shape cannot relate to {rel_type}")
    target = entry["target"] if external else normalize_part_path(entry["target"])
    if not external:
        if not package.has_part(target) or target.endswith(".rels") \
                or target == "[Content_Types].xml":
            raise FullStateError(f"{record.key}: {rel_id} points at {target!r}, which is not "
                                 f"a part of this package")
        if new_shape and rel_type not in SHARED_ON_DUPLICATE:
            # A chart or diagram is per-shape state, as when a slide is duplicated.
            copied = package.copy_part(target, share=lambda rel: rel.type in SHARED_ON_DUPLICATE)
            return package.add_relationship(part, rel_type, copied)
    if current is not None and current.type == rel_type and current.is_external == external \
            and (current.target if external else current.target_part) == target:
        return rel_id
    if external:
        return package.add_external_relationship(part, rel_type, target)
    return package.add_relationship(part, rel_type, target)


# -- typed edits ---------------------------------------------------------------------------


def _apply_feature(shape: "Shape", feature: str, wanted: Any, current: Any) -> None:
    what = f"{shape.id} data-ooxml-{feature}"
    if feature == "name":
        shape.name = wanted
    elif feature in {"x", "y", "cx", "cy"}:
        if wanted == "inherit":
            raise FullStateError(f"{what}: an explicit position cannot be made to inherit again")
        attribute = {"x": "left", "y": "top", "cx": "width", "cy": "height"}[feature]
        setattr(shape, attribute, int(wanted))
    elif feature == "rot":
        shape.rotation = m.parse_degrees(wanted, what) / m.ROTATION_UNIT
    elif feature == "flip-h":
        shape.flip_h = wanted == "1"
    elif feature == "flip-v":
        shape.flip_v = wanted == "1"
    elif feature == "geom":
        if wanted == "custom":
            raise FullStateError(f"{what}: a custom geometry is edited in data-ooxml-xml")
        shape.preset = wanted
    elif feature == "adj":
        adjustments = shape.adjustments
        given = dict(wanted)
        adjustments._write({name: given.get(name) for name in adjustments.names})
    elif feature == "fill":
        _apply_fill(shape, wanted, what)
    elif feature == "line":
        _apply_line(shape.line, wanted, current or {"kind": "inherit"}, what)
    elif feature == "image":
        if shape.kind != "picture":
            raise FullStateError(f"{what}: only a picture has an image")
        shape.replace_image(_part_bytes(shape, wanted, what))
    elif feature == "text":
        if shape.kind != "shape":
            raise FullStateError(f"{what}: a {shape.kind} cannot hold text")
        _apply_text(shape.text_frame, wanted, shape._slide, what)
    elif feature == "table":
        if not shape.has_table:
            raise FullStateError(f"{what}: the shape holds no table")
        _apply_table(shape, wanted, what)
    elif feature == "chart-data":
        if not shape.has_chart:
            raise FullStateError(f"{what}: the shape holds no chart")
        _apply_chart(shape, wanted, what)
    elif feature == "diagram-nodes":
        if not shape.has_diagram:
            raise FullStateError(f"{what}: the shape holds no SmartArt diagram")
        _apply_diagram(shape, wanted, what)


def _part_bytes(shape: "Shape", part: str, what: str) -> bytes:
    data = shape._slide.document.package.read(part)
    if data is None or not part.startswith("ppt/media/"):
        raise FullStateError(f"{what}: {part!r} is not an image in this package")
    return data


def _apply_fill(target, wanted: dict[str, str], what: str) -> None:
    kind = wanted["kind"]
    if kind == "inherit":
        target.fill = None
    elif kind == "none":
        target.fill = "none"
    elif kind == "solid":
        target.fill = m.color_from_fields(wanted, what)
    elif kind == "gradient":
        stops = []
        for stop in wanted["stops"].split(";"):
            position, _, color = stop.partition(" ")
            stops.append((int(position) / 100000, m.color_from_text(color, what)))
        angle = m.parse_degrees(wanted["angle"], what) / m.ROTATION_UNIT \
            if "angle" in wanted else None
        target.set_gradient_fill(stops, angle=angle, path=wanted.get("path"))
    elif kind == "image" and hasattr(target, "set_image_fill"):
        target.set_image_fill(_part_bytes(target, wanted["image"], what))
    else:
        raise FullStateError(f"{what}: a {kind} fill is not editable through typed attributes; "
                             f"edit data-ooxml-xml instead")


def _apply_line(line, wanted: dict[str, str], current: dict[str, str], what: str) -> None:
    if wanted["kind"] == "inherit":
        line.clear()
        return
    if current["kind"] == "inherit":
        with line._change():
            pass  # an explicit, empty outline: the fields below fill it in
        current = {"kind": "set"}
    if wanted.get("w") != current.get("w"):
        line.width = None if "w" not in wanted else int(wanted["w"])
    color_keys = ("fill", "scheme", "rgb", "sys", "prst", "color", "mods")
    if any(wanted.get(k) != current.get(k) for k in color_keys):
        fill = wanted.get("fill")
        if fill is None:
            line.color = None
        elif fill == "none":
            line.visible = False
        elif fill == "solid":
            line.color = m.color_from_fields(wanted, what)
        else:
            raise FullStateError(f"{what}: a {fill} outline is edited in data-ooxml-xml")
    if wanted.get("dash") != current.get("dash"):
        if wanted.get("dash") == "custom":
            raise FullStateError(f"{what}: a custom dash is edited in data-ooxml-xml")
        line.dash = wanted.get("dash")
    if wanted.get("cap") != current.get("cap"):
        line.cap = wanted.get("cap")
    for end in ("head", "tail"):
        if wanted.get(end) != current.get(end):
            parts = wanted[end].split() if end in wanted else None
            setattr(line, end, None if parts is None else Arrowhead(*parts))


def _live_text(frame, slide: "Slide") -> dict[str, Any] | None:
    body = frame._resolve()._text_body(False)
    return None if body is None else m.text_model(body, resolver_for(slide))


def _apply_text(frame, wanted: dict[str, Any], slide: "Slide", what: str) -> None:
    current = _live_text(frame, slide)
    if current is not None and _tolerant(m.canonical_text, current, what) == wanted:
        return
    text = m.plain_text(wanted)
    if current is None or frame.text != text:
        frame.set_text(text)
    for index, wanted_paragraph in enumerate(wanted["p"]):
        paragraph = frame.paragraph(index)
        where = f"{what} p{index}"
        live = _live_paragraph(paragraph, slide, where)
        _apply_paragraph(paragraph, wanted_paragraph, live, where)
        pieces = ["\v" if "br" in item else item["t"] for item in wanted_paragraph["c"]]
        if pieces != ["\v" if "br" in item else item["t"] for item in live["c"]]:
            if "" in pieces:
                raise FullStateError(f"{where}: an empty run can only stay where it is")
            paragraph.segment(pieces)
        live = _live_paragraph(paragraph, slide, where)
        wanted_runs = [item for item in wanted_paragraph["c"] if "br" not in item]
        live_runs = [item for item in live["c"] if "br" not in item]
        for number, (run_wanted, run_live) in enumerate(zip(wanted_runs, live_runs)):
            _apply_run(paragraph.run(number), run_wanted, run_live, f"{where} r{number}")


def _live_paragraph(paragraph, slide: "Slide", what: str) -> dict[str, Any]:
    raw = m.paragraph_model(paragraph._element(), resolver_for(slide))
    return _tolerant(lambda value, w: m.canonical_text({"p": [value]}, w)["p"][0], raw, what)


def _apply_paragraph(paragraph, wanted: dict[str, Any], live: dict[str, Any], what: str) -> None:
    if wanted.get("algn") != live.get("algn"):
        paragraph.alignment = wanted.get("algn")
    if wanted.get("lvl", 0) != live.get("lvl", 0):
        paragraph.level = wanted.get("lvl", 0)
    if wanted.get("marL") != live.get("marL"):
        paragraph.margin_left = wanted.get("marL")
    if wanted.get("indent") != live.get("indent"):
        paragraph.indent = wanted.get("indent")
    for key, attribute in (("spcBef", "space_before"), ("spcAft", "space_after")):
        if wanted.get(key) != live.get(key):
            value = wanted.get(key)
            if value is not None and not value.startswith("pts:"):
                raise FullStateError(f"{what}: percentage {key} is edited in data-ooxml-xml")
            setattr(paragraph, attribute, None if value is None else Pt(int(value[4:]) / 100))
    if wanted.get("lnSpc") != live.get("lnSpc"):
        value = wanted.get("lnSpc")
        if value is None:
            paragraph.line_spacing = None
        elif value.startswith("pct:"):
            paragraph.line_spacing = int(value[4:]) / 100000
        else:
            paragraph.line_spacing_points = int(value[4:]) / 100
    if wanted.get("bu") != live.get("bu"):
        bullet = wanted.get("bu")
        if bullet is None:
            paragraph.clear_bullet()
        elif bullet == "none":
            paragraph.set_no_bullet()
        elif bullet == "picture":
            raise FullStateError(f"{what}: a picture bullet is edited in data-ooxml-xml")
        else:
            paragraph.clear_bullet()
            color = m.color_from_text(bullet["color"], what) if "color" in bullet else None
            size = bullet["size"] / 100000 if "size" in bullet else None
            if "char" in bullet:
                paragraph.set_bullet(bullet["char"], font=bullet.get("font"), color=color,
                                     size=size)
            else:
                paragraph.set_numbering(bullet["num"], start_at=bullet.get("start"),
                                        font=bullet.get("font"), color=color, size=size)


_STRIKE = {None: None, "sngStrike": True, "noStrike": False}


def _apply_run(run, wanted: dict[str, Any], live: dict[str, Any], what: str) -> None:
    if wanted.get("fld") != live.get("fld"):
        raise FullStateError(f"{what}: a field cannot be made or unmade through typed "
                             f"attributes; edit data-ooxml-xml instead")
    if wanted["t"] != live["t"]:
        run.text = wanted["t"]
    for key, attribute in (("b", "bold"), ("i", "italic"), ("latin", "typeface")):
        if wanted.get(key) != live.get(key):
            setattr(run, attribute, wanted.get(key))
    if wanted.get("u") != live.get("u"):
        run.underline = wanted.get("u")
    if wanted.get("strike") != live.get("strike"):
        if wanted.get("strike") not in _STRIKE:
            raise FullStateError(f"{what}: strike {wanted['strike']!r} is edited in "
                                 f"data-ooxml-xml")
        run.strike = _STRIKE[wanted.get("strike")]
    if wanted.get("sz") != live.get("sz"):
        run.size = None if "sz" not in wanted else wanted["sz"] / 100
    if (wanted.get("color"), wanted.get("fill")) != (live.get("color"), live.get("fill")):
        if "color" in wanted:
            run.color = m.color_from_text(wanted["color"], what)
        elif "fill" in wanted:
            raise FullStateError(f"{what}: a {wanted['fill']} text fill is edited in "
                                 f"data-ooxml-xml")
        else:
            run.color = None
    if wanted.get("link") != live.get("link"):
        link = wanted.get("link")
        if link is None:
            run.remove_hyperlink()
        elif "url" in link:
            run.set_hyperlink(link["url"], tooltip=link.get("tip"))
        elif "slide" in link:
            run.set_hyperlink(link["slide"], tooltip=link.get("tip"))
        else:
            raise FullStateError(f"{what}: a click action is edited in data-ooxml-xml")


# -- tables --------------------------------------------------------------------------------


def _live_table(shape: "Shape") -> dict[str, Any]:
    from ..edit.table import _table_element

    return m.table_model(_table_element(shape._element), resolver_for(shape._slide))


def _cell_texts(model: dict[str, Any]) -> list[list[str]]:
    return [[m.plain_text(cell["text"]) if "text" in cell else "" for cell in row]
            for row in model["cells"]]


def _realign(count_live: int, count_wanted: int, live_keys: list, wanted_keys: list,
             delete, insert) -> None:
    """Insert and delete rows (or columns) so the live sequence lines up with the wanted one."""
    matcher = difflib.SequenceMatcher(None, live_keys, wanted_keys, autojunk=False)
    for tag, i1, i2, j1, j2 in reversed(matcher.get_opcodes()):
        if tag == "equal":
            continue
        common = min(i2 - i1, j2 - j1)
        for index in reversed(range(i1 + common, i2)):
            delete(index)
        for offset in range(j2 - j1 - common):
            insert(i1 + common + offset)


def _apply_table(shape: "Shape", wanted: dict[str, Any], what: str) -> None:
    table = shape.table
    live = _live_table(shape)
    if _tolerant(m.canonical_table, live, what) == wanted:
        return
    rows_differ = len(live["rows"]) != len(wanted["rows"])
    columns_differ = len(live["cols"]) != len(wanted["cols"])
    if rows_differ and columns_differ:
        raise FullStateError(f"{what}: change the number of rows and of columns in two steps")
    if rows_differ:
        texts = [tuple(row) for row in _cell_texts(live)]
        wanted_texts = [tuple(row) for row in _cell_texts(wanted)]
        _realign(len(texts), len(wanted_texts), texts, wanted_texts,
                 table.delete_row, lambda index: table.insert_row(index))
    elif columns_differ:
        texts = list(zip(*_cell_texts(live)))
        wanted_texts = list(zip(*_cell_texts(wanted)))
        _realign(len(texts), len(wanted_texts), texts, wanted_texts,
                 table.delete_column, lambda index: table.insert_column(index))
    live = _live_table(shape)
    if len(live["rows"]) != len(wanted["rows"]) or len(live["cols"]) != len(wanted["cols"]):
        raise FullStateError(f"{what}: could not line the table's rows and columns up")

    current = {tuple(region) for region in live["merges"]}
    target = {tuple(region) for region in wanted["merges"]}
    for top, left, _, _ in sorted(current - target):
        table.split(top, left)
    for region in sorted(target - current):
        table.merge(*region)

    for column, width in enumerate(wanted["cols"]):
        if table.column_widths[column] != width:
            table.set_column_width(column, width)
    for row, height in enumerate(wanted["rows"]):
        if table.row_heights[row] != height:
            table.set_row_height(row, height)

    live = _tolerant(m.canonical_table, _live_table(shape), what)
    from ..edit.table import BORDER_TAGS

    for r, row in enumerate(wanted["cells"]):
        for c, cell_wanted in enumerate(row):
            cell_live = live["cells"][r][c]
            if cell_wanted == cell_live:
                continue
            cell = table.cell(r, c)
            where = f"{what} cell{r},{c}"
            if "text" in cell_wanted and cell_wanted["text"] != cell_live.get("text"):
                _apply_text(cell.text_frame, cell_wanted["text"], shape._slide, where)
            fill_wanted = cell_wanted.get("fill", {"kind": "inherit"})
            if fill_wanted != cell_live.get("fill", {"kind": "inherit"}):
                _apply_fill(cell, fill_wanted, where)
            for side in BORDER_TAGS:
                line_wanted = cell_wanted.get(side, {"kind": "inherit"})
                line_live = cell_live.get(side, {"kind": "inherit"})
                if line_wanted != line_live:
                    _apply_line(cell.border(side), line_wanted, line_live, f"{where} {side}")


# -- charts --------------------------------------------------------------------------------


def _apply_chart(shape: "Shape", wanted: dict[str, Any], what: str) -> None:
    """Bring a chart's data, titles and legend to ``wanted`` through the E4 chart edits:
    series and categories are lined up by name and label (like table rows), then renamed,
    relabelled and revalued where they still differ (:func:`ooxml_edit.charts.model.
    apply_chart_model`)."""
    from ooxml_edit.charts.model import ChartModelError, apply_chart_model

    chart = shape.chart
    try:
        apply_chart_model(chart, wanted, what,
                          read_only_hint="edit data-ooxml-xml or the chart part instead")
    except ChartModelError as error:
        raise FullStateError(str(error)) from None


# -- diagrams ------------------------------------------------------------------------------


def _apply_diagram(shape: "Shape", wanted: dict[str, Any], what: str) -> None:
    """Bring a diagram's nodes to ``wanted``: nodes missing from the SVG are removed, nodes
    without an ``id`` are added where the list puts them, and texts are set.  Existing nodes
    keep their order and level -- moving one is refused."""
    from ooxml_edit.charts.model import ChartModelError, apply_diagram_model

    diagram = shape.diagram
    try:
        apply_diagram_model(diagram, wanted, what, through="the SVG")
    except ChartModelError as error:
        raise FullStateError(str(error)) from None


__all__ = ["ApplyReport", "apply_full_state", "read_svg"]
