"""SVG pictures, as PowerPoint 2016 and later store an inserted SVG.

A picture keeps two images: the PNG every consumer can draw, in ``a:blip@r:embed``, and
the SVG it came from, behind an extension of that blip::

    <a:blip r:embed="rId2">
      <a:extLst>
        <a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
          <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main"
                        r:embed="rId3"/>
        </a:ext>
      </a:extLst>
    </a:blip>

Measured in PowerPoint for Mac 16 (2026-10-11, an SVG inserted and the deck saved): the
SVG is a ``ppt/media/imageN.svg`` part (``Default Extension="svg"
ContentType="image/svg+xml"``) behind an ordinary ``.../relationships/image``
relationship -- not the ``2017/06/relationships/svgImage`` type, which is read but not
written; ``xmlns:asvg`` is declared on the ``svgBlip``; the picture is named "Graphic N"
with an empty ``a:picLocks``; the PNG is the SVG drawn at twice its own size (192 dpi).

The SVG is untrusted data: it is parsed before it is stored, with no document type,
no entities and no network (:func:`pptx_agent.fullstate.safe.parse_untrusted`, as
:meth:`Document.apply_svg <pptx_agent.Document.apply_svg>` does), and the bytes stored are
the ones given.
"""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Sequence

if TYPE_CHECKING:  # pragma: no cover
    from ..oxml.package import PresentationPackage

#: ``a:ext`` URI of an SVG alternative to a bitmap blip.
SVG_BLIP_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
#: The ``asvg:`` namespace of ``svgBlip``.
ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_CONTENT_TYPE = "image/svg+xml"
SVG_NS = "http://www.w3.org/2000/svg"
#: The largest SVG stored as a picture, in bytes.
SVG_LIMIT = 16 * 1024 * 1024
#: The longest side of a fallback PNG made here, in pixels.
FALLBACK_MAX_PX = 4096

_EMU_PER_INCH = 914400
#: CSS units per inch: an SVG's lengths, its user units being pixels at 96 dpi.
_PER_INCH = {"px": 96.0, "": 96.0, "pt": 72.0, "pc": 6.0, "in": 1.0, "cm": 2.54,
             "mm": 25.4}
_LENGTH = re.compile(r"^\s*([+]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*([a-z]*)\s*$")
_SVG_TAG = re.compile(rb"<svg[\s>/]")


def looks_like_svg(data: bytes, suffix: str | None = None) -> bool:
    """Whether ``data`` is meant as an SVG: a ``.svg`` path, or markup with an ``<svg`` tag
    in its first 4 KiB (a declaration, a comment or a document type may come first).
    :func:`parse` decides whether it is one, and refuses what is unsafe."""
    if suffix and suffix.lower().lstrip(".") == "svg":
        return True
    head = data[:4096].lstrip(b"\xef\xbb\xbf \t\r\n")
    return head.startswith(b"<") and _SVG_TAG.search(head) is not None


def parse(data: bytes):
    """The SVG's root element, parsed as untrusted data; a ``ValueError`` (``FullStateError``
    or ``UnsafeInput``) for anything else -- a DTD, an entity, more than :data:`SVG_LIMIT`
    bytes, a root that is not ``svg`` in the SVG namespace."""
    from ..fullstate.safe import FullStateError, parse_untrusted

    root = parse_untrusted(data, limit=SVG_LIMIT, what="the SVG")
    if root.tag != f"{{{SVG_NS}}}svg":
        raise FullStateError(f"not an SVG document: the root is {root.tag!r}, not svg in "
                             f"the namespace {SVG_NS}")
    return root


def _length(value: str | None) -> float | None:
    if not value:
        return None
    match = _LENGTH.match(value)
    if match is None or match.group(2) not in _PER_INCH:
        return None              # a percentage, em, or something unreadable
    number = float(match.group(1))
    return number * 96.0 / _PER_INCH[match.group(2)] if number > 0 else None


def native_size(root) -> tuple[float, float] | None:
    """``(width, height)`` in pixels at 96 dpi: the root's ``width`` and ``height`` (in px,
    pt, pc, in, cm or mm), one of them with the ``viewBox``'s proportions, or the
    ``viewBox`` alone -- what PowerPoint inserts an SVG at.  ``None`` when it states none."""
    width, height = _length(root.get("width")), _length(root.get("height"))
    box = None
    numbers = re.split(r"[\s,]+", (root.get("viewBox") or "").strip())
    if len(numbers) == 4:
        try:
            box = tuple(float(n) for n in numbers)
        except ValueError:
            box = None
        if box is not None and (box[2] <= 0 or box[3] <= 0):
            box = None
    if width and height:
        return width, height
    if box is not None:
        if width:
            return width, width * box[3] / box[2]
        if height:
            return height * box[2] / box[3], height
        return box[2], box[3]
    return None


def size_emu(root, width: int | None, height: int | None) -> tuple[int, int]:
    """The frame for an SVG: the size given, the other side in the SVG's proportions, or
    its :func:`native_size` at 96 dpi.  A ``ValueError`` when the SVG states no size and
    the call gives too little."""
    if width is not None and height is not None:
        return int(width), int(height)
    native = native_size(root)
    if native is None:
        raise ValueError("the SVG states no size (width, height or viewBox): pass width "
                         "and height")
    natural_w = round(native[0] * _EMU_PER_INCH / 96)
    natural_h = round(native[1] * _EMU_PER_INCH / 96)
    if width is None and height is None:
        return natural_w, natural_h
    if width is None:
        return round(height * natural_w / natural_h), int(height)
    return int(width), round(width * natural_h / natural_w)


def rasterise(data: bytes, width_emu: int, height_emu: int, dpi: float = 300,
              font_dirs: Sequence[str] | None = None) -> bytes:
    """A PNG of the SVG at its placed size and ``dpi`` (its longer side at most
    :data:`FALLBACK_MAX_PX`), drawn by pptx2svg's rasteriser; an ``ImportError`` naming
    ``pptx-agent[png]`` without it."""
    if dpi <= 0:
        raise ValueError("fallback_dpi must be above 0")
    width = width_emu * dpi / _EMU_PER_INCH
    height = height_emu * dpi / _EMU_PER_INCH
    scale = min(1.0, FALLBACK_MAX_PX / max(width, height, 1))
    width, height = max(1, round(width * scale)), max(1, round(height * scale))
    message = ("an SVG picture needs a PNG fallback: install pptx-agent[png] "
               "(pip install 'pptx-agent[png]'), or pass fallback= a PNG")
    try:
        from pptx2svg.png import RasterizerNotAvailable, svg_to_png
    except ImportError as error:
        raise ImportError(message) from error
    # The rasteriser reads a size in pixels only (resvg-py refuses width="72pt"): the root's
    # is restated in pixels as PowerPoint reads it, a viewBox added where there is none, and
    # a non-UTF-8 SVG re-encoded -- nothing else changes (ooxml_common.svgimage, measured on
    # PowerPoint 16).
    from ooxml_common.svgimage import with_pixel_size

    data = with_pixel_size(data) or data
    if not _is_utf8(data):
        raise ValueError("the SVG is not UTF-8 text")
    text = data.decode("utf-8-sig")
    try:
        return svg_to_png(text, width=width, height=height,
                          check_glyphs=False,
                          **({"font_dirs": list(font_dirs)} if font_dirs is not None else {}))
    except RasterizerNotAvailable as error:
        raise ImportError(message) from error


def _is_utf8(data: bytes) -> bool:
    try:
        data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return False
    return True


def store(package: "PresentationPackage", data: bytes) -> str:
    """Store an SVG under ``ppt/media/``, reusing an identical one already there; its part
    name."""
    existing = package.find_part_with_bytes(data, "ppt/media")
    if existing is not None and package.content_type(existing) == SVG_CONTENT_TYPE:
        return existing
    return package.add_part(package.unused_part_name("ppt/media/image{n}.svg"), data,
                            SVG_CONTENT_TYPE)


def svg_blip(blip):
    """The ``asvg:svgBlip`` under a blip, else ``None``."""
    from ..oxml.xml import qn

    extensions = blip.find(qn("a:extLst"))
    if extensions is None:
        return None
    for extension in extensions:
        if extension.get("uri") == SVG_BLIP_EXT_URI:
            found = extension.find(f"{{{ASVG_NS}}}svgBlip")
            if found is not None:
                return found
    return None


def drop(blip) -> None:
    """Remove a blip's SVG extension (and an ``a:extLst`` it leaves empty)."""
    from ..oxml.xml import qn, remove

    extensions = blip.find(qn("a:extLst"))
    if extensions is None:
        return
    for extension in list(extensions):
        if extension.get("uri") == SVG_BLIP_EXT_URI:
            remove(extension)
    if len(extensions) == 0:
        remove(extensions)


def attach(blip, rel_id: str) -> None:
    """Point a blip's SVG extension at ``rel_id``, writing the extension as PowerPoint does
    (``xmlns:asvg`` on the ``svgBlip``) when there is none."""
    from lxml import etree

    from ..oxml.xml import make, qn

    found = svg_blip(blip)
    if found is not None:
        found.set(qn("r:embed"), rel_id)
        return
    extensions = blip.find(qn("a:extLst"))
    if extensions is None:
        extensions = make("a:extLst")
        blip.append(extensions)
    extension = make("a:ext", uri=SVG_BLIP_EXT_URI)
    extensions.append(extension)
    element = etree.SubElement(extension, f"{{{ASVG_NS}}}svgBlip", nsmap={"asvg": ASVG_NS})
    element.set(qn("r:embed"), rel_id)


# -- recolouring (PowerPoint's Graphics Fill) ---------------------------------------------------

#: The class PowerPoint names a theme colour by in a recoloured SVG, which is how the theme
#: link survives a save and a reopen.  Measured for accent 2 only (PowerPoint for Mac,
#: 2026-10-15); the other slots' names are not written until they are measured -- a
#: colour from any other slot is written as its value alone.
THEME_CLASSES = {"accent2": "MsftOfcThm_Accent2_Fill_v2"}

#: The elements a Graphics Fill fills: SVG's basic shapes, paths and text.
_FILLED = frozenset({"path", "rect", "circle", "ellipse", "polygon", "polyline", "line",
                     "text", "tspan", "textPath"})
#: Containers whose shapes are not drawn as themselves: a clip path's outline, a mask's
#: luminance.  Their shapes keep their colours.
_NOT_DRAWN = frozenset({"clipPath", "mask"})
_THEME_CLASS = re.compile(r"^MsftOfcThm_\w+_Fill_v2$")
_THEME_RULES = re.compile(r"^\s*(?:\.MsftOfcThm_\w+_Fill_v2\s*\{[^}]*\}\s*)+$")
_FILL_DECLARATION = re.compile(r"^\s*(fill|fill-opacity)\s*:", re.IGNORECASE)


class RecolorWarning(UserWarning):
    """A recoloured SVG kept some of its shapes' colours (shapes with no fill)."""


def _declared_fill(element) -> str | None:
    """The fill an element states itself, in its ``style`` or its ``fill`` attribute."""
    for declaration in (element.get("style") or "").split(";"):
        name, _, value = declaration.partition(":")
        if name.strip().lower() == "fill":
            return value.strip().removesuffix("!important").strip()
    value = element.get("fill")
    return value.strip() if value is not None else None


def _fill_is_none(element) -> bool:
    """Whether the element's fill, its own or inherited, is ``none`` (CSS rules aside)."""
    node = element
    while node is not None and isinstance(node.tag, str):
        declared = _declared_fill(node)
        if declared is not None and declared.lower() != "inherit":
            return declared.lower() == "none"
        node = node.getparent()
    return False


def recolor(data: bytes, color: str, theme_class: str | None = None) -> tuple[bytes, int, int]:
    """An SVG with its shapes filled with ``color`` (``"#RRGGBB"``), as PowerPoint's Graphics
    Fill rewrites one; ``(svg, shapes filled, shapes left)``.

    Measured in PowerPoint for Mac 16 (2026-10-15), on a picture of four paths -- two
    filled black, one with no ``fill``, one ``currentColor``: every path gets
    ``fill="#RRGGBB" fill-opacity="1"``; a theme colour also gets a ``<style>`` rule for its
    class (``.MsftOfcThm_Accent2_Fill_v2 { fill:#E97132; }``) and that ``class`` on every
    path.  What is written here, and what is not:

    * every path, basic shape and text element gets ``fill`` and ``fill-opacity="1"`` --
      ``currentColor``, a gradient and no fill at all included -- and, with
      ``theme_class``, the class (added to a class it has) and the rule, first in the root;
    * a shape whose fill is ``none`` (its own or inherited from a group) keeps it, and
      its stroke stays as it is: how PowerPoint fills an outline-only shape is not
      measured.  Shapes in a clip path or a mask keep their colours too;
    * a ``fill`` in a shape's ``style`` attribute goes (it would win over the attribute);
      an SVG with ``<style>`` rules of its own gets the fill in the shapes' ``style`` as
      well, since a class rule wins over an attribute;
    * a class and rule of an earlier recolour (``MsftOfcThm_..._Fill_v2``) go first;
    * strokes are not touched -- PowerPoint writes ``stroke="none"`` and its defaults on
      paths that had none, which draws the same -- and PowerPoint's path normalisation,
      ``overflow`` and ``xmlns:xlink`` are not repeated: they change nothing drawn.

    ``ValueError`` when nothing can be filled (an SVG of outlines only)."""
    from lxml import etree

    root = parse(data)
    svg = f"{{{SVG_NS}}}"
    for style in list(root.iter(f"{svg}style")):
        if style.text and _THEME_RULES.match(style.text):
            _drop(style)
    own_rules = any(True for _ in root.iter(f"{svg}style"))
    filled = left = 0
    for element in root.iter():
        if not isinstance(element.tag, str) or not element.tag.startswith(svg):
            continue
        if element.tag[len(svg):] not in _FILLED:
            continue
        if any(isinstance(a.tag, str) and a.tag[len(svg):] in _NOT_DRAWN
               for a in element.iterancestors()):
            continue
        if _fill_is_none(element):
            left += 1
            continue
        classes = [c for c in (element.get("class") or "").split()
                   if not _THEME_CLASS.match(c)]
        if theme_class:
            classes.append(theme_class)
        if classes:
            element.set("class", " ".join(classes))
        else:
            element.attrib.pop("class", None)
        declarations = [d for d in (element.get("style") or "").split(";")
                        if d.strip() and not _FILL_DECLARATION.match(d)]
        if own_rules:
            declarations += [f"fill:{color}", "fill-opacity:1"]
        if declarations:
            element.set("style", ";".join(d.strip() for d in declarations))
        else:
            element.attrib.pop("style", None)
        element.set("fill", color)
        element.set("fill-opacity", "1")
        filled += 1
    if not filled:
        raise ValueError("the SVG has no shape to fill: its shapes are outlines only (fill "
                         "none), and how PowerPoint recolours those is not measured yet")
    if theme_class:
        style = etree.SubElement(root, f"{svg}style")    # in the root's namespaces
        root.insert(0, style)
        style.tail = "\n"
        style.text = f"\n.{theme_class} {{\n fill:{color}; \n}}\n"
        root.insert(0, style)
    return etree.tostring(root, encoding="utf-8"), filled, left


def _drop(element) -> None:
    parent = element.getparent()
    if element.tail and element.tail.strip():
        previous = element.getprevious()
        if previous is not None:
            previous.tail = (previous.tail or "") + element.tail
        else:
            parent.text = (parent.text or "") + element.tail
    parent.remove(element)


__all__ = ["ASVG_NS", "SVG_BLIP_EXT_URI", "SVG_CONTENT_TYPE", "SVG_LIMIT", "attach", "drop",
           "looks_like_svg", "native_size", "parse", "rasterise", "recolor", "RecolorWarning",
           "size_emu", "store", "svg_blip", "THEME_CLASSES"]
