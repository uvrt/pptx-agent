"""Slide masters and layouts as PowerPoint's Slide Master view edits them, and Insert >
Header & Footer.

A layout or master shape is edited through the same :class:`~pptx_agent.Shape` a slide's
is (:attr:`Layout.shapes <pptx_agent.Layout.shapes>`, :attr:`Master.shapes
<pptx_agent.Master.shapes>`); this module changes the *list* of layouts and fills the
footer placeholders.  What each writes was measured in PowerPoint 16 (Mac), and it writes
the same:

* **Duplicate layout** (Slide Master view, Insert > Duplicate Layout): a copy of the part,
  right after the original, named ``"1_<name>"``; a new ``sldLayoutId``, one more than the
  largest master or layout id; the master's relationship and the content type.
* **Insert layout**: ``preserve="1" userDrawn="1"`` and no ``type`` (a custom layout), a
  title placeholder and the master's date, footer and slide-number placeholders with their
  text; placed after the layout named, else last.
* **Rename**: ``cSld@name`` alone.
* **Delete**: PowerPoint disables it for a layout a slide uses; so does this, unless the
  slides are moved to another layout first (``reassign_to``).  The part, its relationship
  and its ``sldLayoutId`` go; other ids stay.
* **Header & Footer, Apply to All**: the footer text and the date (an updating
  ``datetime1`` field, or fixed text) on the master's and *every* layout's placeholders;
  ``<p:hf>`` on the master saying which are off (``hdr="0"`` always); on each slide the
  layout's ``dt``, ``ftr`` and ``sldNum`` placeholders copied in (empty ``spPr``, so they
  inherit position) or, turned off, removed; a title slide (a ``"title"`` layout's) left
  out with ``showSpecialPlsOnTitleSld="0"``.  **Apply** (some slides) changes only those
  slides.

Layout and master edits show on every slide built on them, so the tools report the slides
each one ``affects`` and check them.
"""

from __future__ import annotations

import copy
import uuid
from datetime import date as _date
from typing import TYPE_CHECKING, Iterable

from ..oxml.package import REL_SLIDE_LAYOUT, REL_SLIDE_MASTER
from ..oxml.xml import (
    NAMESPACES,
    Element,
    append_in_order,
    find,
    findall,
    local_name,
    make,
    qn,
    register_child_order,
    remove,
    serialize,
    subelement,
)
from .ids import TEMPLATE_ID_BASE

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, Slide
    from .slides import Layout, Master

CT_SLIDE_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"

register_child_order({
    "p:sldMaster": ("p:cSld", "p:clrMap", "p:sldLayoutIdLst", "p:transition", "p:timing",
                    "p:hf", "p:txStyles", "p:extLst"),
    "p:sldLayoutIdLst": ("p:sldLayoutId", "p:extLst"),
    "p:sldLayout": ("p:cSld", "p:clrMapOvr", "p:transition", "p:timing", "p:hf", "p:extLst"),
})

#: The Header & Footer placeholders, in the order PowerPoint appends them to a slide.
FOOTER_TYPES = ("dt", "ftr", "sldNum")
_GUID_NAMESPACE = uuid.UUID("6f1d3c52-8a7e-4a43-9a51-2d1c6c0e9b7d")


# -- the layout list ---------------------------------------------------------------------------


def _next_template_id(package) -> int:
    """One more than the largest ``sldMasterId`` or ``sldLayoutId`` in the deck: they share
    one id space (PowerPoint gives a duplicated layout max + 1)."""
    used = [TEMPLATE_ID_BASE - 1]
    presentation = package.tree(package.presentation_part())
    for node in findall(presentation, "p:sldMasterIdLst/p:sldMasterId"):
        if (node.get("id") or "").isdigit():
            used.append(int(node.get("id")))
        master = package.related_part(package.presentation_part(), node.get(qn("r:id")))
        root = package.tree(master) if master else None
        for entry in findall(root, "p:sldLayoutIdLst/p:sldLayoutId") if root is not None else ():
            if (entry.get("id") or "").isdigit():
                used.append(int(entry.get("id")))
    return max(used) + 1


def _entry_for(package, layout: "Layout") -> Element:
    root = package.tree(layout.master_part)
    for node in findall(root, "p:sldLayoutIdLst/p:sldLayoutId"):
        if package.related_part(layout.master_part, node.get(qn("r:id"))) == layout.part_path:
            return node
    raise KeyError(f"{layout.part_path} is not in its master's layout list")


def _list_layout(document: "Document", master_part: str, part: str,
                 after: "Layout | None") -> str:
    """Relate ``part`` to its master and list it after ``after`` (else last); returns the
    new layout's key."""
    package = document.package
    rel_id = package.add_relationship(master_part, REL_SLIDE_LAYOUT, part)
    if not package.related_parts_of_type(part, REL_SLIDE_MASTER):
        package.add_relationship(part, REL_SLIDE_MASTER, master_part)
    identifier = _next_template_id(package)
    node = make("p:sldLayoutId", id=str(identifier), r__id=rel_id)
    root = package.tree(master_part)
    id_list = subelement(root, "p:sldLayoutIdLst")
    if after is not None:
        _entry_for(package, after).addnext(node)
    else:
        entries = findall(id_list, "p:sldLayoutId")
        if entries:
            entries[-1].addnext(node)
        else:
            id_list.insert(0, node)
    package.mark_dirty(master_part)
    return f"L{identifier - TEMPLATE_ID_BASE}"


def _unique_name(document: "Document", master_part: str, wanted: str, *,
                 numbered: bool) -> str:
    names = {layout.name for layout in document.layouts if layout.master_part == master_part}
    if not numbered:
        return wanted
    k = 1
    while f"{k}_{wanted}" in names:
        k += 1
    return f"{k}_{wanted}"


def _check_name(document: "Document", master_part: str, name: str,
                skip: str | None = None) -> str:
    name = str(name).strip()
    if not name:
        raise ValueError("a layout needs a name")
    for layout in document.layouts:
        if layout.master_part == master_part and layout.name == name \
                and layout.part_path != skip:
            raise ValueError(f"the master already has a layout named {name!r} ({layout.id}); "
                             "layouts are found by name, so names stay unique")
    return name


def duplicate_layout(document: "Document", layout: "Layout", name: str | None = None
                     ) -> "Layout":
    package = document.package
    target_name = _check_name(document, layout.master_part, name) if name is not None \
        else _unique_name(document, layout.master_part, layout.name, numbered=True)
    with document.batch():
        document.history.checkpoint()
        # Everything a layout relates to is shared: its master, its pictures.
        part = package.copy_part(layout.part_path, share=lambda rel: True)
        root = package.tree(part)
        find(root, "p:cSld").set("name", target_name)
        package.mark_dirty(part)
        key = _list_layout(document, layout.master_part, part, layout)
        document._reset_caches()
    return document.layout(key)


def add_layout(document: "Document", name: str, master: "Master | None",
               after: "Layout | None") -> "Layout":
    package = document.package
    if master is None:
        master = after.master if after is not None else document.masters[0]
    if after is not None and after.master_part != master.part_path:
        raise ValueError(f"{after.id} belongs to {after.master.id}, not {master.id}")
    name = _check_name(document, master.part_path, name)
    with document.batch():
        document.history.checkpoint()
        part = package.unused_part_name("ppt/slideLayouts/slideLayout{n}.xml")
        package.add_part(part, _new_layout_xml(document, master, name), CT_SLIDE_LAYOUT,
                         override=True)
        key = _list_layout(document, master.part_path, part, after)
        document._reset_caches()
    return document.layout(key)


def _new_layout_xml(document: "Document", master: "Master", name: str) -> bytes:
    """A custom layout as Insert Layout makes one: a title, and the master's date, footer
    and slide-number placeholders with their text, all inheriting their place."""
    from . import inherit
    from .slides import make_root

    package = document.package
    root = make_root("p:sldLayout", {p: NAMESPACES[p] for p in ("a", "r", "p")})
    root.set("preserve", "1")
    root.set("userDrawn", "1")
    common = subelement(root, "p:cSld")
    common.set("name", name)
    tree = subelement(common, "p:spTree")
    group = subelement(tree, "p:nvGrpSpPr")
    group.append(make("p:cNvPr", id="1", name=""))
    group.append(make("p:cNvGrpSpPr"))
    group.append(make("p:nvPr"))
    xfrm = subelement(subelement(tree, "p:grpSpPr"), "a:xfrm")
    for tag, attributes in (("a:off", {"x": "0", "y": "0"}), ("a:ext", {"cx": "0", "cy": "0"}),
                            ("a:chOff", {"x": "0", "y": "0"}),
                            ("a:chExt", {"cx": "0", "cy": "0"})):
        xfrm.append(make(tag, **attributes))
    # idx and sz as the master's other layouts give them (10, 11, 12 in Office's own).
    known: dict[str, Element] = {}
    for layout in master.layouts:
        for shape in document._template_sheet(layout.part_path)._sp_tree():
            ph = inherit.placeholder_of(shape)
            if ph is not None and ph.get("type") in FOOTER_TYPES and ph.get("type") not in known:
                known[ph.get("type")] = ph
    next_id = 2
    for kind in ("title", *FOOTER_TYPES):
        source = inherit.master_placeholder(package, master.part_path, inherit.master_kind(kind))
        if source is None:
            continue
        sp = make("p:sp")
        nv = make("p:nvSpPr")
        source_pr = find(source, "p:nvSpPr/p:cNvPr")
        label = (source_pr.get("name") if source_pr is not None else None) or "Placeholder"
        label = " ".join(label.split()[:-1]) if label.split()[-1].isdigit() else label
        if kind == "title":
            label = "Title"
        nv.append(make("p:cNvPr", id=str(next_id), name=f"{label} {next_id - 1}"))
        locks = make("p:cNvSpPr")
        locks.append(make("a:spLocks", noGrp="1"))
        nv.append(locks)
        nv_pr = make("p:nvPr")
        ph = make("p:ph", type=kind)
        template = known.get(kind)
        for attribute in ("sz", "idx"):
            value = template.get(attribute) if template is not None else None
            if value is None and kind != "title":
                value = {"sz": {"dt": "half"}.get(kind, "quarter"),
                         "idx": str(10 + FOOTER_TYPES.index(kind))}[attribute]
            if value is not None:
                ph.set(attribute, value)
        nv_pr.append(ph)
        nv.append(nv_pr)
        sp.append(nv)
        sp.append(make("p:spPr"))
        body = make("p:txBody")
        body.append(make("a:bodyPr"))
        body.append(make("a:lstStyle"))
        paragraphs = findall(source, "p:txBody/a:p")[:1]
        for paragraph in paragraphs:
            body.append(copy.deepcopy(paragraph))
        if not paragraphs:
            body.append(make("a:p"))
        sp.append(body)
        tree.append(sp)
        next_id += 1
    color_map = subelement(root, "p:clrMapOvr")
    color_map.append(make("a:masterClrMapping"))
    return serialize(root)


def rename_layout(document: "Document", layout: "Layout", name: str) -> "Layout":
    name = _check_name(document, layout.master_part, name, skip=layout.part_path)
    package = document.package
    with document.batch():
        document.history.checkpoint()
        find(package.tree(layout.part_path), "p:cSld").set("name", name)
        package.mark_dirty(layout.part_path)
        document._reset_caches()
    return document.layout(layout.part_path)


def delete_layout(document: "Document", layout: "Layout",
                  reassign_to: "Layout | None") -> list[int]:
    """Delete ``layout``; returns the ``sldId`` of each slide moved to ``reassign_to``."""
    package = document.package
    using = layout.slides
    if using and reassign_to is None:
        raise ValueError(
            f"{layout.id} {layout.name!r} is in use by slide(s) "
            f"{', '.join(f's:{s.slide_id}' for s in using)}; PowerPoint does not delete a "
            "layout in use either. Move them to another layout (reassign_to) or delete them "
            "first")
    if reassign_to is not None and reassign_to.part_path == layout.part_path:
        raise ValueError("reassign_to is the layout being deleted")
    if len([lay for lay in document.layouts if lay.master_part == layout.master_part]) == 1:
        raise ValueError(f"{layout.id} is its master's only layout")
    moved = []
    with document.batch():
        document.history.checkpoint()
        for slide in using:
            for rel_id, rel in list(package.relationships(slide.part_path).items()):
                if rel.type == REL_SLIDE_LAYOUT:
                    package.remove_relationship(slide.part_path, rel_id)
            package.add_relationship(slide.part_path, REL_SLIDE_LAYOUT, reassign_to.part_path)
            moved.append(slide.slide_id)
        entry = _entry_for(package, layout)
        rel_id = entry.get(qn("r:id"))
        remove(entry)
        package.mark_dirty(layout.master_part)
        package.release(layout.master_part, [rel_id])
        document._reset_caches()
    return moved


# -- Header & Footer ---------------------------------------------------------------------------


def _date_text(language: str, today: _date) -> str:
    """``datetime1``'s text: the short date as the deck's language writes it."""
    if language.lower().startswith("en-us"):
        return f"{today.month}/{today.day}/{today.year}"
    return f"{today.day}/{today.month}/{today.year}"


def _guid(*parts: object) -> str:
    return "{" + str(uuid.uuid5(_GUID_NAMESPACE, ":".join(map(str, parts)))).upper() + "}"


def _paragraph(sp: Element) -> Element:
    body = subelement(sp, "p:txBody")
    subelement(body, "a:bodyPr")
    subelement(body, "a:lstStyle")
    paragraphs = findall(body, "a:p")
    for extra in paragraphs[1:]:
        remove(extra)
    return paragraphs[0] if paragraphs else subelement(body, "a:p")


def _clear(paragraph: Element) -> Element | None:
    """Empty a paragraph but for its ``a:pPr``; returns the first run's (or field's)
    ``a:rPr``, to keep its formatting."""
    keep = None
    for child in list(paragraph):
        if local_name(child) in ("r", "fld") and keep is None:
            keep = child.find(qn("a:rPr"))
        if local_name(child) not in ("pPr", "endParaRPr"):
            remove(child)
    return keep


def _insert_content(paragraph: Element, node: Element, language: str) -> None:
    end = paragraph.find(qn("a:endParaRPr"))
    if end is None:
        end = make("a:endParaRPr", lang=language)
        paragraph.append(end)
    end.addprevious(node)


def _rpr(kept: Element | None, language: str, smart: bool) -> Element:
    rpr = copy.deepcopy(kept) if kept is not None else make("a:rPr")
    rpr.set("lang", rpr.get("lang") or language)
    if smart:
        rpr.set("smtClean", "0")
    return rpr


def _write_footer_text(sp: Element, text: str, language: str) -> None:
    paragraph = _paragraph(sp)
    kept = _clear(paragraph)
    if text:
        run = make("a:r")
        run.append(_rpr(kept, language, smart=False))
        subelement(run, "a:t").text = text
        _insert_content(paragraph, run, language)


def _write_date(sp: Element, date: str, language: str, today: _date, part: str) -> None:
    paragraph = _paragraph(sp)
    old = paragraph.find(qn("a:fld"))
    kind = old.get("type") if old is not None else None
    kept = _clear(paragraph)
    if date == "auto":
        if not (kind or "").startswith("datetime") or kind == "datetimeFigureOut":
            kind = "datetime1"
        node = make("a:fld", id=_guid(part, "dt"), type=kind)
        node.append(_rpr(kept, language, smart=True))
        subelement(node, "a:t").text = _date_text(language, today)
    else:
        node = make("a:r")
        node.append(_rpr(kept, language, smart=False))
        subelement(node, "a:t").text = date
    _insert_content(paragraph, node, language)


def _placeholders(tree: Element, kind: str) -> list[Element]:
    from . import inherit

    return [shape for shape in tree if local_name(shape) == "sp"
            and (inherit.placeholder_of(shape) is not None)
            and inherit.placeholder_of(shape).get("type") == kind]


def _copy_to_slide(slide: "Slide", source: Element, kind: str, number: int, *, text: str,
                   date: str | None, language: str, today: _date) -> Element:
    """A slide's ``dt``/``ftr``/``sldNum`` placeholder copied from its layout's, as Header &
    Footer adds one: its name and ``p:ph``, an empty ``spPr`` (it inherits its place)."""
    from . import inherit

    sp = make("p:sp")
    nv = make("p:nvSpPr")
    source_pr = find(source, "p:nvSpPr/p:cNvPr")
    nv.append(make("p:cNvPr", id=str(slide._next_shape_id()),
                   name=(source_pr.get("name") if source_pr is not None else None)
                   or f"{kind} placeholder"))
    locks = make("p:cNvSpPr")
    locks.append(make("a:spLocks", noGrp="1"))
    nv.append(locks)
    nv_pr = make("p:nvPr")
    ph = make("p:ph")
    for attribute in ("type", "orient", "sz", "idx"):
        value = inherit.placeholder_of(source).get(attribute)
        if value is not None:
            ph.set(attribute, value)
    nv_pr.append(ph)
    nv.append(nv_pr)
    sp.append(nv)
    sp.append(make("p:spPr"))
    body = make("p:txBody")
    body.append(make("a:bodyPr"))
    body.append(make("a:lstStyle"))
    paragraph = make("a:p")
    body.append(paragraph)
    sp.append(body)
    if kind == "sldNum":
        field = next(iter(source.iter(qn("a:fld"))), None)
        node = make("a:fld", id=field.get("id") if field is not None else _guid(
            slide.part_path, "sldNum"), type="slidenum")
        node.append(make("a:rPr", lang=language, smtClean="0"))
        subelement(node, "a:t").text = str(number)
        paragraph.append(node)
        paragraph.append(make("a:endParaRPr", lang=language))
    elif kind == "ftr":
        _write_footer_text(sp, text, language)
    else:
        _write_date(sp, date, language, today, slide.part_path)
    append_in_order(slide._sp_tree(), sp)
    return sp


def footers_for_new_slide(document: "Document", slide: "Slide") -> bool:
    """After Header & Footer's Apply to All a slide made with New Slide gets its layout's
    date, footer and slide-number placeholders, content and all, for each the master's
    ``<p:hf>`` has on -- and none on a title slide when ``showSpecialPlsOnTitleSld`` is
    ``0`` (measured: PowerPoint's New Slide).  A master without ``<p:hf>`` (Header & Footer
    never applied) gives none.  Returns whether any was added."""
    from . import inherit

    package = document.package
    layout_part = inherit.layout_of(package, slide.part_path)
    master_part = inherit.master_of(package, layout_part)
    master = package.tree(master_part) if master_part else None
    hf = master.find(qn("p:hf")) if master is not None else None
    if hf is None:
        return False
    layout = package.tree(layout_part)
    presentation = package.tree(package.presentation_part())
    if layout.get("type") == "title" and presentation.get("showSpecialPlsOnTitleSld") in (
            "0", "false"):
        return False
    tree = find(layout, "p:cSld/p:spTree")
    added = False
    for kind in FOOTER_TYPES:
        if hf.get(kind) in ("0", "false") or _placeholders(slide._sp_tree(), kind):
            continue
        source = _placeholders(tree, kind)
        if not source:
            continue
        sp = _copy_to_slide(slide, source[0], kind, slide.index + 1, text="", date="auto",
                            language=_language(document), today=_date.today())
        paragraph = find(sp, "p:txBody/a:p")
        for child in list(paragraph):
            remove(child)
        for child in find(source[0], "p:txBody/a:p") if find(source[0], "p:txBody/a:p")\
                is not None else ():
            paragraph.append(copy.deepcopy(child))
        if kind == "sldNum":
            for text in paragraph.iter(qn("a:t")):
                text.text = str(slide.index + 1)
        added = True
    if added:
        slide._touch()
    return added


def _language(document: "Document") -> str:
    from .authoring import default_language

    return default_language(document.package)


def set_footer(document: "Document", *, text: str | None, slide_number: bool,
               date: str | None, slides: "Iterable[Slide] | None", skip_title: bool,
               today: _date | None) -> list[int]:
    """See :meth:`Document.set_footer`; returns the ``sldId`` of every slide changed."""
    from . import inherit
    from .authoring import default_language

    package = document.package
    language = default_language(package)
    today = today or _date.today()
    text = text or ""
    if date is not None and not str(date).strip():
        date = None
    wanted = {"dt": date is not None, "ftr": bool(text), "sldNum": bool(slide_number)}
    every = slides is None
    targets = list(document.slides) if every else list(slides)
    changed: list[int] = []
    with document.batch():
        document.history.checkpoint()
        if every:
            presentation = package.tree(package.presentation_part())
            if skip_title:
                presentation.set("showSpecialPlsOnTitleSld", "0")
            elif presentation.get("showSpecialPlsOnTitleSld") is not None:
                del presentation.attrib["showSpecialPlsOnTitleSld"]
            package.mark_dirty(package.presentation_part())
            for master in document.masters:
                parts = [master.part_path, *[lay.part_path for lay in master.layouts]]
                for part in parts:
                    tree = find(package.tree(part), "p:cSld/p:spTree")
                    for sp in _placeholders(tree, "ftr") if text else ():
                        _write_footer_text(sp, text, language)
                    for sp in _placeholders(tree, "dt") if date is not None else ():
                        _write_date(sp, date, language, today, part)
                    package.mark_dirty(part)
                hf = subelement(package.tree(master.part_path), "p:hf")
                for attribute in list(hf.attrib):
                    del hf.attrib[attribute]
                for kind in ("sldNum", "hdr", "ftr", "dt"):
                    if not wanted.get(kind, False):
                        hf.set(kind, "0")
        for slide in targets:
            layout_part = inherit.layout_of(package, slide.part_path)
            layout_tree = find(package.tree(layout_part), "p:cSld/p:spTree") \
                if layout_part else None
            title_slide = layout_part is not None and \
                package.tree(layout_part).get("type") == "title"
            tree = slide._sp_tree()
            touched = False
            for kind in FOOTER_TYPES:
                existing = _placeholders(tree, kind)
                show = wanted[kind] and not (skip_title and title_slide)
                if not show:
                    for sp in existing:
                        remove(sp)
                        touched = True
                    continue
                if existing:
                    for sp in existing:
                        if kind == "ftr":
                            _write_footer_text(sp, text, language)
                        elif kind == "dt":
                            _write_date(sp, date, language, today, slide.part_path)
                    touched = True
                    continue
                source = _placeholders(layout_tree, kind) if layout_tree is not None else []
                if source:
                    _copy_to_slide(slide, source[0], kind, slide.index + 1, text=text,
                                   date=date, language=language, today=today)
                    touched = True
            if touched:
                slide._touch()
                changed.append(slide.slide_id)
        document._reset_caches()
    return changed


__all__ = ["CT_SLIDE_LAYOUT", "FOOTER_TYPES", "add_layout", "delete_layout", "duplicate_layout",
           "footers_for_new_slide", "rename_layout", "set_footer"]
