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
    text = data.decode("utf-8-sig", errors="strict") if _is_utf8(data) else None
    if text is None:
        raise ValueError("the SVG is not UTF-8 text")
    # The rasteriser reads a size in pixels only (resvg-py refuses width="72pt"): the root
    # gets its native size in pixels, or the placed one when it states none.
    root = parse(data)
    native = native_size(root) or (width, height)
    if (root.get("width"), root.get("height")) != tuple(f"{v:g}" for v in native):
        from lxml import etree

        root.set("width", f"{native[0]:g}")
        root.set("height", f"{native[1]:g}")
        text = etree.tostring(root, encoding="unicode")
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


__all__ = ["ASVG_NS", "SVG_BLIP_EXT_URI", "SVG_CONTENT_TYPE", "SVG_LIMIT", "attach", "drop",
           "looks_like_svg", "native_size", "parse", "rasterise", "size_emu", "store",
           "svg_blip"]
