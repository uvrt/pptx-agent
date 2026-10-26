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

What is the same in every format -- the blip's extension, the native size, the recolouring
as Graphics Fill and Outline rewrite an SVG, measured here (PowerPoint for Mac 16,
2026-10-15, oracle2/results.md) -- is ooxml-edit's :mod:`ooxml_edit.pictures` since round
17 (P6: docx-agent tints an icon the same way, Word drawing it as PowerPoint does); its
names are this module's still.  What is a deck's stays here: :func:`parse`'s errors (a
``FullStateError``), the frame's size in EMU, the PNG fallback drawn by pptx2svg, the media
part under ``ppt/media``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Sequence

from ooxml_edit.pictures import svg as _shared
from ooxml_edit.pictures.svg import (ASVG_NS, SVG_BLIP_EXT_URI, SVG_CONTENT_TYPE, SVG_LIMIT, SVG_NS,
                                     THEME_SLOTS, attach, drop, looks_like_svg, native_size,
                                     stroke_units, svg_blip, theme_class)

from .errors import ArgumentError

if TYPE_CHECKING:  # pragma: no cover
    from ..oxml.package import PresentationPackage

#: The longest side of a fallback PNG made here, in pixels.
FALLBACK_MAX_PX = 4096

_EMU_PER_INCH = 914400


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


def recolor(data: bytes, fill: str | None = None, fill_class: str | None = None,
            opacity: float = 1.0, **options) -> bytes:
    """:func:`ooxml_edit.pictures.recolor`, the SVG read by :func:`parse` first (so what is
    refused is refused as a deck refuses it)."""
    parse(data)
    return _shared.recolor(data, fill, fill_class, opacity, **options)


def size_emu(root, width: int | None, height: int | None) -> tuple[int, int]:
    """The frame for an SVG: the size given, the other side in the SVG's proportions, or
    its :func:`native_size` at 96 dpi.  A ``ValueError`` when the SVG states no size and
    the call gives too little."""
    if width is not None and height is not None:
        return int(width), int(height)
    native = native_size(root)
    if native is None:
        raise ArgumentError("the SVG states no size (width, height or viewBox): pass width "
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
        raise ArgumentError("fallback_dpi must be above 0")
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
        raise ArgumentError("the SVG is not UTF-8 text")
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



__all__ = ["ASVG_NS", "SVG_BLIP_EXT_URI", "SVG_CONTENT_TYPE", "SVG_LIMIT", "attach", "drop",
           "looks_like_svg", "native_size", "parse", "rasterise", "recolor", "size_emu",
           "store", "stroke_units", "svg_blip", "theme_class", "THEME_SLOTS"]
