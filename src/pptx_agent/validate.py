"""Structural validity checks: what PowerPoint repairs, or refuses, when it is wrong.

The only authoritative answer to "will it open?" is PowerPoint itself (the opt-in oracle in
``tests/test_validity.py``).  These checks cover the failure modes that actually produce a
repair prompt, run anywhere, and are what :meth:`Document.validate` returns.  Each is a
:class:`Problem` with a stable ``code`` and a ``detail`` that does not depend on positions,
so the problems of a deck before and after an edit can be compared: an edit may leave a
problem the deck arrived with (one fixture's generator declares a master it does not have),
but no edit adds one.

* ``not-well-formed`` -- an XML or relationships part that does not parse;
* ``no-content-type``, ``override-without-part`` -- ``[Content_Types].xml`` and the parts
  disagree;
* ``dangling-relationship`` -- a relationship whose target part is missing;
* ``orphaned-part`` -- a part no chain of relationships reaches;
* ``unknown-relationship-id`` -- an ``r:id``/``r:embed``/... a part's XML spells that is
  not one of that part's relationships;
* ``slide-list`` -- a ``p:sldId`` that does not resolve to a slide part, a repeated or
  out-of-range id, two entries for one slide part, or sections that no longer partition
  the slide list in order;
* ``child-order`` -- children out of their schema sequence, for every element whose order
  the library registers (in slides, layouts and masters);
* ``table-grid`` -- a table row whose cell count is not the grid's, or a span that runs
  past the table;
* ``ignorable-prefix-undeclared`` -- an ``mc:Ignorable`` prefix the part does not declare;
* ``package-type`` -- with a target file name only (:func:`package_type_problems`): the main
  part declares a presentation or a template, plain or macro-enabled, and the extension says
  otherwise.  PowerPoint refuses such a file outright, with no offer to repair it.
"""

from __future__ import annotations

from dataclasses import dataclass

from lxml import etree
from ooxml_edit.xml import _ranks, prefixed_name

from .oxml.package import CONTENT_TYPES_PART, PresentationPackage
from .oxml.xml import Element, qn

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_MC_IGNORABLE = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
_ORDERED_PARTS = ("ppt/slides/", "ppt/slideLayouts/", "ppt/slideMasters/")


@dataclass(frozen=True, order=True)
class Problem:
    """One structural problem: ``code`` (stable), the ``part`` it is in, and a ``detail``::

        for problem in deck.validate():
            print(problem)        # [dangling-relationship] ppt/slides/slide1.xml: image -> ...
    """

    code: str
    part: str
    detail: str = ""

    def __str__(self) -> str:
        return f"[{self.code}] {self.part}: {self.detail}"


def check(package: PresentationPackage) -> list[Problem]:
    """Every problem found in the package, sorted."""
    problems: list[Problem] = []
    roots: dict[str, Element] = {}
    for name in package.part_names:
        if not name.endswith((".xml", ".rels")):
            continue
        try:
            etree.fromstring(package.read(name), etree.XMLParser(resolve_entities=False))
        except etree.XMLSyntaxError as error:
            problems.append(Problem("not-well-formed", name, str(error)))
            continue
        if name.endswith(".xml") and name != CONTENT_TYPES_PART:
            roots[name] = package.tree(name)
    problems += _package_problems(package)
    for name, root in roots.items():
        problems += _relationship_id_problems(package, name, root)
        problems += _ignorable_problems(name, root)
        if name.startswith(_ORDERED_PARTS):
            problems += _order_problems(name, root)
        if name.startswith("ppt/slides/"):
            problems += _table_problems(name, root)
    problems += _slide_list_problems(package)
    return sorted(set(problems))


def _package_problems(package: PresentationPackage) -> list[Problem]:
    out: list[Problem] = []
    parts = set(package.part_names)
    for name in parts:
        if name != CONTENT_TYPES_PART and not name.endswith("/") \
                and package.content_type(name) is None:
            out.append(Problem("no-content-type", name))
    types = package.tree(CONTENT_TYPES_PART)
    if types is not None:
        for override in types:
            if not isinstance(override.tag, str) or not override.tag.endswith("}Override"):
                continue
            target = (override.get("PartName") or "").lstrip("/")
            if target and target not in parts:
                out.append(Problem("override-without-part", CONTENT_TYPES_PART, target))
    for owner in [""] + sorted(parts):
        if owner.endswith(".rels"):
            continue
        for relationship in package.relationships(owner).values():
            if not relationship.is_external and relationship.target_part not in parts:
                out.append(Problem("dangling-relationship", owner or "/",
                                   f"{relationship.type.rpartition('/')[2]} -> "
                                   f"{relationship.target_part}"))
    for orphan in sorted(package.unreachable_parts()):
        out.append(Problem("orphaned-part", orphan))
    return out


def _relationship_id_problems(package: PresentationPackage, name: str,
                              root: Element) -> list[Problem]:
    known = set(package.relationships(name))
    out: list[Problem] = []
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        for attribute, value in element.attrib.items():
            if attribute.startswith("{" + _R_NS + "}") and value and value not in known:
                out.append(Problem("unknown-relationship-id", name,
                                   f"{prefixed_name(element)}/@r:{attribute.rpartition('}')[2]}"
                                   f"={value}"))
    return out


def _ignorable_problems(name: str, root: Element) -> list[Problem]:
    listed = (root.get(_MC_IGNORABLE) or "").split()
    return [Problem("ignorable-prefix-undeclared", name, prefix)
            for prefix in listed if prefix not in root.nsmap]


def _order_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    for element in root.iter():
        if not isinstance(element.tag, str):
            continue
        parent = prefixed_name(element)
        ranks = _ranks(parent)
        if not ranks:
            continue
        unknown = len(ranks) + 1
        last = -1
        for child in element:
            if not isinstance(child.tag, str):
                continue
            rank = ranks.get(prefixed_name(child), unknown)
            if rank < last:
                out.append(Problem("child-order", name, f"{prefixed_name(child)} in {parent}"))
                break
            last = max(last, rank)
    return out


def _table_problems(name: str, root: Element) -> list[Problem]:
    out: list[Problem] = []
    for table in root.iter(qn("a:tbl")):
        grid = table.find(qn("a:tblGrid"))
        columns = len(grid.findall(qn("a:gridCol"))) if grid is not None else 0
        rows = table.findall(qn("a:tr"))
        for r, row in enumerate(rows):
            cells = row.findall(qn("a:tc"))
            if len(cells) != columns:
                out.append(Problem("table-grid", name,
                                   f"a row has {len(cells)} cells for {columns} grid columns"))
                continue
            for c, cell in enumerate(cells):
                span_c = int(cell.get("gridSpan") or 1)
                span_r = int(cell.get("rowSpan") or 1)
                if c + span_c > columns or r + span_r > len(rows):
                    out.append(Problem("table-grid", name, "a merge runs past the table"))
    return out


def _slide_list_problems(package: PresentationPackage) -> list[Problem]:
    from .edit.slides import sections

    presentation = package.presentation_part()
    root = package.tree(presentation) if presentation else None
    if root is None:
        return []
    out: list[Problem] = []
    listed = root.findall(f"{qn('p:sldIdLst')}/{qn('p:sldId')}")
    relationships = package.relationships(presentation)
    ids: list[int] = []
    parts: list[str] = []
    for entry in listed:
        rel = relationships.get(entry.get(qn("r:id")) or "")
        try:
            slide_id = int(entry.get("id") or "")
        except ValueError:
            out.append(Problem("slide-list", presentation, f"sldId {entry.get('id')!r}"))
            continue
        ids.append(slide_id)
        if rel is None or rel.target_part is None or not package.has_part(rel.target_part):
            out.append(Problem("slide-list", presentation, f"sldId {slide_id} has no slide"))
            continue
        parts.append(rel.target_part)
        if not 256 <= slide_id < 2147483648:
            out.append(Problem("slide-list", presentation, f"sldId {slide_id} out of range"))
    for slide_id in sorted({i for i in ids if ids.count(i) > 1}):
        out.append(Problem("slide-list", presentation, f"sldId {slide_id} repeated"))
    for part in sorted({p for p in parts if parts.count(p) > 1}):
        out.append(Problem("slide-list", presentation, f"two entries for {part}"))
    in_sections = [slide_id for _, members in sections(root) for slide_id in members]
    if in_sections and in_sections != ids:
        out.append(Problem("slide-list", presentation,
                           "sections no longer partition the slide list in order"))
    return out


def package_type_problems(package: PresentationPackage, target) -> list[Problem]:
    """A ``package-type`` problem when the main part's content type does not match what
    ``target``'s extension promises (``.pptx``/``.pptm`` a presentation, ``.potx``/``.potm``
    a template; the ``m`` forms macro-enabled)."""
    import os

    from ooxml_common.kinds import KINDS, kind_for, kind_mismatch

    expected = kind_for(target, "powerpoint")
    if expected is None:
        return []
    main = package.presentation_part()
    current = package.content_type(main) or ""
    mismatch = kind_mismatch(current, target)
    if mismatch is None:
        return []
    template = KINDS[expected].template
    declared_template = mismatch.actual is not None and KINDS[mismatch.actual].template
    extension = os.path.splitext(os.fspath(target))[1].lower()
    out = []
    if mismatch.template_differs:
        out.append(Problem(
            "package-type", main,
            f"declared a {'template' if declared_template else 'presentation'} "
            f"({current.split('.', 2)[-1]}) but the target is {extension}; save() "
            "writes the type the extension needs, bytes written otherwise would not open"))
    if mismatch.macros_differ and not mismatch.macros_forbidden:
        out.append(Problem("package-type", main,
                           f"declares no macros but the target is {extension}; save() "
                           "writes the type the extension needs"))
    if mismatch.macros_forbidden:
        out.append(Problem("package-type", main,
                           f"macro-enabled, but the target is {extension}, which cannot carry "
                           f"macros; save as "
                           f"{'.potm' if template else '.pptm'}"))
    return out


__all__ = ["Problem", "check", "package_type_problems"]
