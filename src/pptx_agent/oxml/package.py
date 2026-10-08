"""The ``.pptx`` package: the format-neutral OPC container plus PresentationML entry points.

Everything about reading, editing and losslessly writing parts lives in
:mod:`ooxml_edit.opc`; this module only adds how to find the presentation and its slides.
"""

from __future__ import annotations

from . import xml as _xml  # noqa: F401  registers the p:/a: namespaces before qn() is used
from ooxml_edit.opc import (  # noqa: F401  (re-exported)
    CONTENT_TYPES_PART,
    REL_IMAGE,
    REL_OFFICE_DOCUMENT,
    RELS_NS,
    OpcPackage,
    Relationship,
    _part_for_rels,
    normalize_part_path,
    rels_path_for,
    resolve_target,
)
from .xml import findall, qn

_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
REL_SLIDE = _REL + "slide"
REL_SLIDE_LAYOUT = _REL + "slideLayout"
REL_SLIDE_MASTER = _REL + "slideMaster"
REL_NOTES_SLIDE = _REL + "notesSlide"
REL_NOTES_MASTER = _REL + "notesMaster"
REL_HYPERLINK = _REL + "hyperlink"
REL_CHART = _REL + "chart"

CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"

#: Relationships a duplicated slide may share with its original, because their targets are
#: never edited through one slide: the layout and masters it inherits from, media (stored
#: once and never rewritten in place -- an edit adds a new part), other slides it links to,
#: fonts.  Everything else -- notes, charts and their embedded workbooks, diagrams, OLE
#: objects, comments, tags -- is per-slide state and is copied, which is also the safe answer
#: for a relationship type this list does not know.
SHARED_ON_DUPLICATE: frozenset[str] = frozenset({
    REL_SLIDE_LAYOUT, REL_SLIDE_MASTER, REL_NOTES_MASTER, REL_SLIDE, REL_HYPERLINK,
    REL_IMAGE, _REL + "theme", _REL + "font", _REL + "audio", _REL + "video",
    "http://schemas.microsoft.com/office/2007/relationships/media",
    "http://schemas.microsoft.com/office/2007/relationships/hdphoto",
    "http://schemas.microsoft.com/office/2017/06/relationships/svgImage",
})


#: Image content types by extension, for the formats PowerPoint embeds natively.
IMAGE_CONTENT_TYPES: dict[str, str] = {
    "png": "image/png",
    "jpeg": "image/jpeg",
    "jpg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
}


def sniff_image_extension(data: bytes) -> str | None:
    """The file extension an image's magic number says it has."""
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if data.startswith(b"BM"):
        return "bmp"
    if data[:4] in (b"II*\x00", b"MM\x00*"):
        return "tiff"
    return None


def image_size(data: bytes) -> tuple[int, int] | None:
    """``(width, height)`` in pixels from an image header, or ``None`` if it cannot be read."""
    import struct

    try:
        if data.startswith(b"\x89PNG\r\n\x1a\n"):
            return struct.unpack(">II", data[16:24])
        if data[:6] in (b"GIF87a", b"GIF89a"):
            return struct.unpack("<HH", data[6:10])
        if data.startswith(b"BM"):
            width, height = struct.unpack("<ii", data[18:26])
            return width, abs(height)
        if data.startswith(b"\xff\xd8"):
            position = 2
            while position + 9 < len(data):
                if data[position] != 0xFF:
                    position += 1
                    continue
                marker = data[position + 1]
                if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7 or marker == 0xFF:
                    position += 1 if marker == 0xFF else 2
                    continue
                length = struct.unpack(">H", data[position + 2:position + 4])[0]
                # Start-of-frame markers, excluding DHT (C4), JPG (C8) and DAC (CC).
                if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
                    height, width = struct.unpack(">HH", data[position + 5:position + 9])
                    return width, height
                position += 2 + length
    except struct.error:
        return None
    return None


def image_dpi(data: bytes) -> tuple[float, float] | None:
    """The resolution an image states, ``(x, y)`` dots per inch: a PNG's ``pHYs`` chunk in
    pixels per metre, a JPEG's JFIF density.  ``None`` when it states none."""
    import struct

    try:
        if data.startswith(b"\x89PNG\r\n\x1a\n"):
            position = 8
            while position + 8 <= len(data):
                length, kind = struct.unpack(">I4s", data[position:position + 8])
                if kind == b"pHYs" and length >= 9:
                    x, y, unit = struct.unpack(">IIB", data[position + 8:position + 17])
                    return (round(x * 0.0254, 2), round(y * 0.0254, 2)) if unit == 1 and x \
                        and y else None
                if kind in (b"IDAT", b"IEND"):
                    return None
                position += 12 + length
            return None
        if data.startswith(b"\xff\xd8") and data[6:11] == b"JFIF\x00":
            unit, x, y = struct.unpack(">BHH", data[13:18])
            if not x or not y or unit not in (1, 2):
                return None
            return (float(x), float(y)) if unit == 1 else (round(x * 2.54, 2), round(y * 2.54, 2))
    except struct.error:
        return None
    return None


class PresentationPackage(OpcPackage):
    """An open .pptx: the OPC package, plus where its presentation and slides are."""

    def add_image(self, data: bytes, extension: str | None = None) -> str:
        """Store an image under ``ppt/media/``; returns its part name.

        An identical image already in the deck is reused instead of stored twice, which is
        also what PowerPoint does.
        """
        extension = (extension or sniff_image_extension(data) or "").lower().lstrip(".")
        if extension not in IMAGE_CONTENT_TYPES:
            raise ValueError(
                "unsupported image format; expected one of "
                + ", ".join(sorted(set(IMAGE_CONTENT_TYPES)))
            )
        existing = self.find_part_with_bytes(data, "ppt/media")
        if existing is not None and self.content_type(existing) == IMAGE_CONTENT_TYPES[extension]:
            return existing
        path = self.unused_part_name(f"ppt/media/image{{n}}.{extension}")
        return self.add_part(path, data, IMAGE_CONTENT_TYPES[extension])

    def presentation_part(self) -> str:
        main = self.main_document_part()
        if main is not None:
            return main
        if self.has_part("ppt/presentation.xml"):
            return "ppt/presentation.xml"
        raise ValueError("not a PowerPoint package: no presentation part")

    def slide_parts(self) -> list[tuple[int, str]]:
        """``(sldId, part path)`` for every slide, in presentation order.

        ``p:sldId/@id`` is a deck-unique integer that survives reordering, unlike the slide's
        position -- which is why the id scheme keys off it rather than an index.
        """
        presentation = self.presentation_part()
        root = self.tree(presentation)
        if root is None:
            return []
        slides: list[tuple[int, str]] = []
        for node in findall(root, "p:sldIdLst/p:sldId"):
            raw_id = node.get("id")
            part = self.related_part(presentation, node.get(qn("r:id")))
            if raw_id is None or part is None:
                continue
            try:
                slides.append((int(raw_id), part))
            except ValueError:
                continue
        return slides


#: The E0 name, kept so existing imports work.
OoxmlPackage = PresentationPackage
