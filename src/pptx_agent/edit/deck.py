"""Deck-level operations: a new deck (from nothing or from a template), saving one as a
template, and the deck's metadata, language and slide size.

**From nothing** the parts are :mod:`.blank`'s.  **From a template** -- a ``.pptx`` or a
``.potx``, as a path or bytes -- the template's masters, layouts, themes, notes and handout
masters and properties are kept and its slides go: each is deleted as
:meth:`Document.delete_slide` deletes one, so its notes, charts, media and relationships go
with it and only what nothing else uses is reaped.  Then what described those slides goes
too: sections and custom shows (which only list slides), and the thumbnail (a picture of
the first one).  A ``.potx``'s main part is a *template* (``presentationml.template.main``);
the new deck's is a presentation.  The core properties start again (no title, the given
author, created and modified now, revision 1); others the template set, like its company or
category, stay.

Either way the result is opened afresh, so creating the deck is the base state undo returns
to.
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import TYPE_CHECKING, BinaryIO

from ooxml_common.kinds import KINDS, kind_for

from ..oxml.xml import find, findall, make, qn, remove, subelement
from . import blank
from .properties import CoreProperties, now, refreshed_app, w3cdtf

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

# The package kinds -- which extension names which kind, and the main part's content type
# for it -- are ooxml-common's, shared with docx-agent.
_POWERPOINT_KINDS = {name: kind for name, kind in KINDS.items() if kind.application == "powerpoint"}
CT_MACRO_TEMPLATE = KINDS["potm"].content_type
CT_MACRO_PRESENTATION = KINDS["pptm"].content_type
REL_THUMBNAIL = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"


def new(cls, *, size=None, template=None, title: str | None = None,
        author: str | None = None, language: str | None = None,
        created: datetime | None = None) -> "Document":
    created = (created or now()).replace(microsecond=0)
    if template is None:
        width, height = blank.slide_size("16:9" if size is None else size)
        data = blank.build(width, height, language=language or "en-US", created=created,
                           title=title, author=author)
        return _open(cls, data)
    document = _open(cls, _read(template))
    _strip_template(document)
    properties = CoreProperties(document.package)
    for tag in ("dc:title", "dc:creator", "cp:lastModifiedBy"):
        properties.set(tag, None)
    properties.set("dc:title", title or "")
    if author:
        properties.set("dc:creator", author)
        properties.set("cp:lastModifiedBy", author)
    properties.set("cp:revision", "1")
    properties.set("dcterms:created", w3cdtf(created), dated=True)
    properties.set("dcterms:modified", w3cdtf(created), dated=True)
    if language is not None:
        set_language(document, language)
    if size is not None:
        document.set_slide_size(size)
    return _open(cls, document.to_bytes())


def _open(cls, data: bytes) -> "Document":
    """Open without :meth:`Document.open`'s template warning: here a template is the point."""
    from ..oxml.package import OoxmlPackage

    return cls(OoxmlPackage.open(data))


def _read(source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        with open(os.fspath(source), "rb") as handle:
            return handle.read()
    if hasattr(source, "read"):
        return source.read()
    raise TypeError(f"a template is a path or bytes, not {type(source).__name__}")


def _strip_template(document: "Document") -> None:
    package = document.package
    main = package.presentation_part()
    content_type = package.content_type(main)
    if content_type in (CT_MACRO_TEMPLATE, CT_MACRO_PRESENTATION):
        raise ValueError("a macro-enabled template cannot be used: macros are out of scope")
    if content_type == blank.CT_TEMPLATE:
        set_main_content_type(package, blank.CT_PRESENTATION)
    for slide in list(document.slides):
        document.delete_slide(slide)
    root = package.tree(main)
    changed = False
    for extension in findall(root, "p:extLst/p:ext"):
        if extension.get("uri") == SECTION_EXT_URI:
            remove(extension)
            changed = True
    extensions = find(root, "p:extLst")
    if extensions is not None and len(extensions) == 0:
        remove(extensions)
    shows = find(root, "p:custShowLst")
    if shows is not None:
        remove(shows)
        changed = True
    if changed:
        package.mark_dirty(main)
    for rel in list(package.relationships("").values()):
        if rel.type == REL_THUMBNAIL:
            package.remove_relationship("", rel.id)
            if rel.target_part is not None:
                package.reap([rel.target_part])


def set_main_content_type(package, content_type: str) -> None:
    """Point the main part's ``Override`` at ``content_type`` (presentation <-> template)."""
    main = package.presentation_part()
    root = package.tree("[Content_Types].xml")
    for node in root:
        if node.get("PartName", "").lstrip("/") == main:
            node.set("ContentType", content_type)
    package.mark_dirty("[Content_Types].xml")


#: The main part's content type for each kind of package: (presentation, template), plain
#: and macro-enabled.  PowerPoint refuses a file whose extension and main content type
#: disagree -- a .pptx declaring a template, a .potx declaring a presentation.
_MAIN_TYPES = {(kind.macro_enabled, kind.template): kind.content_type
               for kind in _POWERPOINT_KINDS.values()}
TEMPLATE_TYPES = frozenset(kind.content_type for kind in _POWERPOINT_KINDS.values()
                           if kind.template)
MACRO_TYPES = frozenset(kind.content_type for kind in _POWERPOINT_KINDS.values()
                        if kind.macro_enabled)


class TemplateOpenedWarning(UserWarning):
    """:meth:`Document.open` was given a template (``.potx``): it opens the template
    itself, its sample slides and all.  To make a deck *from* a template, use
    ``Document.new(template=...)``; a template opened this way and saved as ``.pptx``
    is declared a presentation by :meth:`Document.save`.

    For example::

        warnings.simplefilter("error", TemplateOpenedWarning)   # make it an error
    """


def target_kind(target) -> tuple[bool, bool] | None:
    """``(macro-enabled, template)`` as a target's extension says, or ``None`` when it is
    not a path or names no PowerPoint extension."""
    name = kind_for(target, "powerpoint")
    if name is None:
        return None
    kind = KINDS[name]
    return kind.macro_enabled, kind.template


def main_content_type(package) -> str | None:
    return package.content_type(package.presentation_part())


def wanted_content_type(package, *, template: bool | None,
                        macro: bool | None = None) -> str | None:
    """The main content type a save should write: ``template`` and ``macro`` as given
    (``None``: as the package is now).  ``None`` when that is what it already has.

    A macro-enabled deck stays macro-enabled: saving one as ``.pptx`` or ``.potx`` would
    keep its VBA project in a file that may not carry one, so that is refused.
    """
    current = main_content_type(package)
    is_macro = current in MACRO_TYPES
    is_template = current in TEMPLATE_TYPES
    if current not in _MAIN_TYPES.values():
        return None  # not a main part this library knows how to retype
    if macro is not None and is_macro and not macro:
        raise ValueError("a macro-enabled deck is saved as .pptm or .potm, not .pptx or .potx")
    want_macro = is_macro if macro is None else macro
    want_template = is_template if template is None else template
    wanted = _MAIN_TYPES[(want_macro, want_template)]
    return None if wanted == current else wanted


def write_replacements(document: "Document", *, template: bool | None = None,
                       macro: bool | None = None) -> dict[str, bytes]:
    """What a save writes in place of a part's own bytes: ``app.xml`` brought up to date
    when the slides changed, and the main part's content type -- a presentation or a
    template (``template``), plain or macro-enabled (``macro``); ``None`` keeps what the
    package has."""
    from .properties import _app_part

    package = document.package
    replacements: dict[str, bytes] = {}
    if _slides_changed(package):
        refreshed = refreshed_app(package, _opened_summary(package))
        if refreshed is not None:
            replacements[_app_part(package)] = refreshed
    main = package.presentation_part()
    wanted = wanted_content_type(package, template=template, macro=macro)
    if wanted is not None:
        # Serialized as Office writes the part (ooxml-edit), the open package unchanged.
        replacements["[Content_Types].xml"] = package.content_types_with(main, wanted)
    return replacements


def _slides_changed(package) -> bool:
    changed = package.changed_parts()
    if not changed:
        return False
    try:
        main = package.presentation_part()
    except ValueError:
        return False
    folders = {"ppt/slides", "ppt/notesSlides"}
    folders |= {part.rpartition("/")[0] for _, part in package.slide_parts()}
    return any(path == main or path.rpartition("/")[0] in folders
               or path.rpartition("/")[0].removesuffix("/_rels") in folders
               for path in changed)


def _opened_summary(package) -> dict | None:
    from .properties import deck_summary

    try:
        return deck_summary(package.opened())
    except (ValueError, KeyError):
        return None


# -- language ---------------------------------------------------------------------------------


def set_language(document: "Document", language: str) -> None:
    """The deck's default text language: ``defaultTextStyle`` and every master's
    ``otherStyle`` -- what new text, and this library's new shapes, take."""
    from ..oxml.package import REL_SLIDE_MASTER

    package = document.package
    main = package.presentation_part()
    targets = [(main, "p:defaultTextStyle")]
    for master in package.related_parts_of_type(main, REL_SLIDE_MASTER):
        targets.append((master, "p:txStyles/p:otherStyle"))
    for part, path in targets:
        root = package.tree(part)
        holder = find(root, path) if root is not None else None
        if holder is None:
            if part != main:
                continue
            holder = subelement(root, "p:defaultTextStyle")
        properties = holder.find(qn("a:defPPr"))
        if properties is None:
            properties = make("a:defPPr")
            holder.insert(0, properties)
        run = properties.find(qn("a:defRPr"))
        if run is None:
            run = make("a:defRPr")
            properties.append(run)
        if run.get("lang") != language:
            run.set("lang", language)
            package.mark_dirty(part)




# -- slide size -------------------------------------------------------------------------------

#: Paragraph attributes that are text metrics, and so scale with the text.
_PARAGRAPH_METRICS = ("marL", "marR", "indent", "defTabSz")
_RUN_TAGS = ("a:rPr", "a:defRPr", "a:endParaRPr")
_PARAGRAPH_TAGS = ("a:pPr", "a:defPPr") + tuple(f"a:lvl{n}pPr" for n in range(1, 10))
#: A run's size when nothing on the way says otherwise.
_DEFAULT_SIZE = 1800


def resize(document: "Document", width: int, height: int, *, scale: bool = True) -> None:
    """Set ``p:sldSz``, and with ``scale`` lay the deck out for it as PowerPoint does.

    Measured by switching decks between sizes in PowerPoint ("Ensure Fit"): masters and
    layouts are stretched -- frames axis by axis -- while slide content is scaled uniformly
    by the smaller ratio and centred, so nothing on a slide is distorted.  Text metrics
    scale by the smaller ratio everywhere; line widths and insets stay.
    """
    from ..oxml.package import REL_SLIDE_LAYOUT, REL_SLIDE_MASTER

    package = document.package
    main = package.presentation_part()
    root = package.tree(main)
    node = subelement(root, "p:sldSz")
    old_width = int(node.get("cx") or 9144000)
    old_height = int(node.get("cy") or 6858000)
    node.set("cx", str(width))
    node.set("cy", str(height))
    known = blank._SIZE_TYPES.get((width, height))
    if known:
        node.set("type", known)
    else:
        node.attrib.pop("type", None)
    package.mark_dirty(main)
    if not scale:
        return
    sx, sy = width / old_width, height / old_height
    defaults = _default_sizes(root)
    designs: list[str] = []
    for master in package.related_parts_of_type(main, REL_SLIDE_MASTER):
        designs.append(master)
        designs.extend(package.related_parts_of_type(master, REL_SLIDE_LAYOUT))
    for part in dict.fromkeys(designs):
        tree = package.tree(part)
        if tree is not None and scale_design(tree, sx, sy):
            package.mark_dirty(part)
    for _, part in package.slide_parts():
        tree = package.tree(part)
        if tree is not None and scale_slide(tree, old_width, old_height, width, height,
                                            defaults):
            package.mark_dirty(part)


def scale_design(root, sx: float, sy: float) -> bool:
    """A master or layout: top-level frames (and table grids) axis by axis, every explicit
    text metric by ``min(sx, sy)``.  Returns whether anything changed."""
    changed = _scale_frames(root, lambda x: x * sx, lambda y: y * sy, sx, sy, sx, sy)
    return _scale_text(root, min(sx, sy)) or changed


def scale_slide(root, old_width: int, old_height: int, width: int, height: int,
                defaults: dict[int, int]) -> bool:
    """A slide: top-level frames scaled uniformly by the smaller ratio and centred; text
    metrics by the same ratio.  Text outside placeholders that took its size from the
    presentation's default text style -- which does not scale -- gets the scaled size
    written, as PowerPoint writes it."""
    factor = min(width / old_width, height / old_height)
    dx = (width - old_width * factor) / 2
    dy = (height - old_height * factor) / 2
    changed = _scale_frames(root, lambda x: x * factor + dx, lambda y: y * factor + dy,
                            factor, factor, factor, factor)
    written = _materialise_sizes(root, factor, defaults)
    return _scale_text(root, factor, written) or changed or bool(written)


def _scale_frames(root, map_x, map_y, sx, sy, grid_x, grid_y) -> bool:
    changed = False
    tree = find(root, "p:cSld/p:spTree")
    for shape in [] if tree is None else list(tree):
        xfrm = _frame_of(shape)
        if xfrm is None:
            continue
        offset, extent = xfrm.find(qn("a:off")), xfrm.find(qn("a:ext"))
        if offset is not None:
            offset.set("x", str(round(map_x(int(offset.get("x", "0"))))))
            offset.set("y", str(round(map_y(int(offset.get("y", "0"))))))
            changed = True
        if extent is not None:
            extent.set("cx", str(round(int(extent.get("cx", "0")) * sx)))
            extent.set("cy", str(round(int(extent.get("cy", "0")) * sy)))
            changed = True
        for grid in shape.iter(qn("a:gridCol")):
            grid.set("w", str(round(int(grid.get("w", "0")) * grid_x)))
        for row in shape.iter(qn("a:tr")):
            row.set("h", str(round(int(row.get("h", "0")) * grid_y)))
    return changed


def _scale_text(root, factor: float, done: "set | None" = None) -> bool:
    changed = False
    for tag in _RUN_TAGS:
        for node in root.iter(qn(tag)):
            if node.get("sz") is not None and (done is None or node not in done):
                node.set("sz", str(max(100, round(int(node.get("sz")) * factor))))
                changed = True
    for tag in _PARAGRAPH_TAGS:
        for node in root.iter(qn(tag)):
            for attribute in _PARAGRAPH_METRICS:
                if node.get(attribute) is not None:
                    node.set(attribute, str(round(int(node.get(attribute)) * factor)))
                    changed = True
    for node in root.iter(qn("a:spcPts")):
        node.set("val", str(round(int(node.get("val", "0")) * factor)))
        changed = True
    return changed


def _default_sizes(presentation) -> dict[int, int]:
    """Level -> size in the presentation's ``defaultTextStyle``."""
    sizes = {}
    style = find(presentation, "p:defaultTextStyle")
    for level in range(1, 10):
        node = find(style, f"a:lvl{level}pPr/a:defRPr") if style is not None else None
        if node is not None and node.get("sz"):
            sizes[level - 1] = int(node.get("sz"))
    return sizes


def _materialise_sizes(root, factor: float, defaults: dict[int, int]) -> set:
    """Size the runs outside placeholders against what they inherit, which does not scale:
    an inherited size is written scaled, an explicit one is scaled -- and dropped when it
    lands on the inherited size, as PowerPoint writes it.  Returns the properties done."""
    from ..oxml.xml import local_name

    written = set()
    for body in list(root.iter(qn("p:txBody"))) + list(root.iter(qn("a:txBody"))):
        shape = body.getparent()
        if local_name(shape) == "sp" and find(shape, "p:nvSpPr/p:nvPr/p:ph") is not None:
            continue  # a placeholder inherits from its layout and master, which scale
        styles = body.find(qn("a:lstStyle"))
        for paragraph in body.findall(qn("a:p")):
            properties = paragraph.find(qn("a:pPr"))
            level = int(properties.get("lvl", "0")) if properties is not None else 0
            inherited = None
            if styles is not None:
                node = find(styles, f"a:lvl{level + 1}pPr/a:defRPr")
                if node is not None and node.get("sz"):
                    inherited = int(node.get("sz"))
            if inherited is None:
                inherited = defaults.get(level, _DEFAULT_SIZE)
            target = round(inherited * factor)
            for item in paragraph:
                name = local_name(item)
                if name in ("r", "fld", "br"):
                    run = item.find(qn("a:rPr"))
                    if run is None:
                        run = make("a:rPr")
                        item.insert(0, run)
                elif name == "endParaRPr":
                    run = item
                else:
                    continue
                if run.get("sz") is None:
                    if target != inherited:
                        run.set("sz", str(target))
                elif round(int(run.get("sz")) * factor) == inherited:
                    del run.attrib["sz"]  # scaled onto the inherited size: PowerPoint drops it
                else:
                    run.set("sz", str(max(100, round(int(run.get("sz")) * factor))))
                written.add(run)
    return written


def _frame_of(shape):
    from ..oxml.xml import local_name

    name = local_name(shape)
    if name == "graphicFrame":
        return shape.find(qn("p:xfrm"))
    holder = shape.find(qn("p:grpSpPr" if name == "grpSp" else "p:spPr"))
    return None if holder is None else holder.find(qn("a:xfrm"))


__all__ = ["MACRO_TYPES", "TEMPLATE_TYPES", "TemplateOpenedWarning", "main_content_type", "new",
           "resize", "scale_design", "scale_slide", "set_language", "set_main_content_type",
           "target_kind", "wanted_content_type", "write_replacements"]
