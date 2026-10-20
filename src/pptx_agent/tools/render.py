"""Rendering slides in a worker process: the functions the tool layer's pool runs.

Module-level, so a spawned worker can import them; they get the deck's bytes and return
PNGs.  Registered (``worker_task``), so a subprocess worker -- what a daemonic process, a
Celery prefork child, renders with -- may run them.

:func:`render_views` is what the ``render`` tool runs, a chunk of a call's slides in each
worker at once: it draws each view's SVG, keys the image by its content
(:func:`raster_key`: the SVG, and everything else the pixels depend on), and rasterises
only the views whose key the caller does not hold already.  It works from the bytes it is given
(``pptx2svg`` opens them itself); nothing is parsed into a :class:`~pptx_agent.Document`
or written back to bytes.  Its PNGs are :func:`render_slides`' byte for byte: the same
pptx2svg passes, one view at a time.
"""

from __future__ import annotations

import hashlib
import os
import struct
import tempfile
import types
import warnings
from typing import Any

from ooxml_edit.tools import content_key, worker_task
from ooxml_edit.tools.rendercache import code_stamp


@worker_task
def render_slides(data: bytes, numbers: "list[int | str]", width: int,
                  font_dirs: "list[str] | None" = None, locale: Any = None) -> list[bytes]:
    """PNG bytes of the 1-based slides ``numbers`` of the deck ``data``, ``width`` px wide.
    An item may be a layout's or master's key (``"L2"``, ``"M0"``): its template view."""
    return render_slides_reporting(data, numbers, width, font_dirs, locale)[0]


@worker_task
def render_slides_reporting(data: bytes, numbers: "list[int | str]", width: int,
                            font_dirs: "list[str] | None" = None, locale: Any = None
                            ) -> tuple[list[bytes], list[list[dict]]]:
    """:func:`render_slides`, and for each slide the text its image leaves out: pptx2svg's
    ``glyphs-missing`` warnings -- no font the renderer has can draw it -- as
    ``{"face", "script", "sample"}``, reported on the first slide that has them.

    ``font_dirs`` is the application's font folders (``Document.font_dirs``), resolved by
    the caller -- the session's, or ``OOXML_FONT_DIRS`` as the caller read it -- since the
    worker sees neither; ``None`` leaves the worker's own default.  ``locale`` is the
    session's (``call.locale``, an :class:`ooxml_common.locales.Locale`: picklable), handed
    over the same way; ``None`` is en-US."""
    from ..edit.document import Document

    deck = Document.open(data)
    if font_dirs is not None:
        deck.font_dirs = tuple(font_dirs)
    deck.locale = locale
    found: list = []
    with warnings.catch_warnings():
        # The same facts come back in ``found``; the worker's stderr is nobody's to read.
        warnings.filterwarnings("ignore", message=".*no font this render loads")
        images = list(deck.render_png(list(numbers), width=int(width), warnings=found))
    # A layout's or master's view is reported by its part (its slide number is 0).
    keys = {template.part_path: template.id for template in [*deck.layouts, *deck.masters]}
    by_slide: dict = {}
    for warning in found:
        detail = getattr(warning, "detail", None)
        if getattr(warning, "code", None) != "glyphs-missing" or detail is None:
            continue
        where = keys.get(getattr(warning, "part_path", None)) or warning.slide_number
        by_slide.setdefault(where, []).append(
            {"face": detail.face, "script": detail.script, "sample": detail.sample})
    return images, [by_slide.get(number, []) for number in numbers]


@worker_task
def rasterise_tiles(svg: str, font_dirs: "list[str] | None" = None) -> bytes:
    """The PNG of a contact image's SVG (``ooxml_edit.tools.tiles.tiled_svg``: the slides'
    PNGs placed at their own size, their labels): drawn once, at the SVG's own size.  The
    labels are in the bundle's sans-serif, the machine's fonts left out so the same sheet
    draws the same everywhere; a label a font cannot draw is not a deck's text, so it is
    not reported."""
    from pptx2svg.png import svg_to_png

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        return svg_to_png(svg, font_dirs=font_dirs, host_fonts=False, check_glyphs=False)


def png_size(png: bytes) -> tuple[int, int]:
    """A PNG's width and height, read from its header."""
    if png[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("not a PNG")
    return struct.unpack(">II", png[16:24])


# -- the render tool's task -----------------------------------------------------------------


@worker_task
def warm_up() -> None:
    """A new worker's first task, before any call's: what the first render would pay for
    -- the renderer's imports, the index of the faces PowerPoint draws with here, the
    rasteriser's own load of the system's fonts (once per process) -- by drawing one
    line of text, and the stamp of the renderers' code (:func:`_code_stamp`)."""
    from pptx2svg.png import svg_to_png

    _code_stamp()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        svg_to_png('<svg xmlns="http://www.w3.org/2000/svg" width="64" height="16">'
                   '<text x="0" y="12" font-family="Calibri, sans-serif" font-size="10">'
                   'Aa</text></svg>')


@worker_task
def render_views(data: bytes, items: "list[int | str]", width: int,
                 font_dirs: "list[str]", locale: Any = None, known: "list[str]" = (),
                 fonts: "dict[str, bytes] | None" = None) -> dict:
    """Each of ``items`` -- a 1-based slide number, or a layout's or master's part path
    (its template view) -- drawn ``width`` px wide from the deck ``data``.

    Returns ``{"views": [(key, png, missing, drawn), ...], "fonts": {...}}``, the views in
    ``items``' order: ``key`` is the image's content key (:func:`raster_key`); ``png`` is
    ``None`` where the key is in ``known`` (the caller holds that image) and ``drawn`` is
    then false; ``missing`` is pptx2svg's ``glyphs-missing`` facts of the view,
    ``{"face", "script", "sample"}`` each.  ``"fonts"`` is the faces this task decoded
    from the deck's embedded fonts, by the hex digest of their payload, for the caller to
    hand the next task as ``fonts`` -- which this worker takes as decoded
    (:func:`ooxml_common.fonts.embedded.remember_decoded`), so no worker decodes a face
    twice.  ``font_dirs`` and ``locale`` as :func:`render_slides_reporting` takes them."""
    import pptx2svg
    from ooxml_common.fonts import embedded
    from ooxml_common.locales import EN_US

    locale = EN_US if locale is None else locale
    embedded.remember_decoded({bytes.fromhex(digest): font
                               for digest, font in (fonts or {}).items()})
    before = set(embedded.decoded_faces())
    numbers = sorted({item for item in items if isinstance(item, int)})
    parts = [item for item in items if isinstance(item, str)]
    extra = {"templates": parts} if parts else {}
    options = pptx2svg.ConvertOptions(slide_numbers=numbers, width=int(width), locale=locale,
                                      font_dirs=list(font_dirs), **extra)
    with warnings.catch_warnings():
        # The same facts come back in the result; the worker's stderr is nobody's to read.
        warnings.filterwarnings("ignore", message=".*no font this render loads")
        documents, resolved = pptx2svg._render(data, options)
        host_fonts = pptx2svg._host_fonts(options)
        dirs = pptx2svg.user_font_dirs(options)
        fingerprint = raster_fingerprint(dirs, host_fonts, resolved.embedded_fonts,
                                         options.check_glyphs)
        views = {}
        for document, slide in zip(documents, resolved.slides):
            ident = slide.template_part if slide.template_part else slide.slide_number
            views[ident] = (document, (slide.slide_number, slide.template_part))
        known_keys = set(known)
        out: dict = {}
        todo = []
        for item in dict.fromkeys(items):
            document, number = views[item]
            key = raster_key(document, fingerprint)
            if key in known_keys:
                out[item] = (key, None, [], False)
            else:
                todo.append((item, key, document, number))
        if todo:
            with _embedded_files(resolved.embedded_fonts) as files:
                for item, key, document, number in todo:
                    report = types.SimpleNamespace(check_glyphs=options.check_glyphs,
                                                   warnings=[])
                    png = pptx2svg._rasterise([document], "auto", dirs, files, None, True,
                                              host_fonts, report, [number])[0]
                    missing = [{"face": w.detail.face, "script": w.detail.script,
                                "sample": w.detail.sample} for w in report.warnings
                               if getattr(w, "code", None) == "glyphs-missing"
                               and getattr(w, "detail", None) is not None]
                    out[item] = (key, png, missing, True)
    decoded = {digest.hex(): font for digest, font in embedded.decoded_faces().items()
               if digest not in before}
    return {"views": [out[item] for item in items], "fonts": decoded}


class _embedded_files:
    """The deck's embedded faces written to a temporary folder for the rasteriser, as
    ``pptx2svg.convert_pptx_to_png`` writes them (``None`` when there are none)."""

    def __init__(self, faces: Any) -> None:
        self._faces = faces
        self._folder: "tempfile.TemporaryDirectory | None" = None

    def __enter__(self) -> "list[str] | None":
        if not self._faces:
            return None
        self._folder = tempfile.TemporaryDirectory(prefix="pptx2svg-fonts-")
        return self._faces.write(self._folder.name)

    def __exit__(self, *exc: Any) -> None:
        if self._folder is not None:
            self._folder.cleanup()


def _code_stamp() -> list:
    """The code that draws the pixels -- pptx2svg, ooxml-common, resvg-py
    (:func:`ooxml_edit.tools.rendercache.code_stamp`): a new install is a new key."""
    return code_stamp("pptx2svg", "ooxml_common", "resvg_py")


def raster_fingerprint(font_dirs: "list[str]", host_fonts: bool, embedded: Any,
                       check_glyphs: bool) -> str:
    """What a slide's pixels depend on besides its SVG, as one key: every font file the
    rasteriser may be handed -- the application's folders and the bundle, the faces
    PowerPoint draws with here (``host_fonts``), the system's where they take part --
    each with its time and size; the deck's embedded faces, by digest; whether missing
    glyphs are reported; and the renderers' code (:func:`_code_stamp`).  All of them, not
    only those one SVG names: a font changed anywhere is a new key, never a stale image."""
    from ooxml_common.fonts.rasterdb import file_stamps, resvg_system_dirs
    from pptx2svg.fonts import font_dirs as bundled_font_dirs

    bundled = [str(path) for path in bundled_font_dirs()]
    system: Any = None
    if not bundled:                      # svg_to_png skips the system's fonts with a bundle
        found = resvg_system_dirs()
        system = file_stamps((), found) if found is not None else "fontconfig"
    host: Any = None
    if host_fonts:
        from pptx2svg.fonts import office

        host = [file_stamps((), [str(path)]) for _, path in office.search_dirs()]
        emoji = office.emoji_face()
        host.append(getattr(emoji, "path", None) and str(emoji.path))
    faces = [(face.family, face.bold, face.italic, hashlib.sha256(face.source).hexdigest())
             for face in getattr(embedded, "faces", ())]
    return content_key("pptx-agent raster", _code_stamp(), bool(check_glyphs),
                       file_stamps((), [*map(str, font_dirs), *bundled]), system, host, faces,
                       os.environ.get("PPTX2SVG_OFFICE_FONTS"))


def raster_key(svg: str, fingerprint: str) -> str:
    """A slide image's content key: its SVG and :func:`raster_fingerprint`."""
    return content_key("pptx-agent png", svg, fingerprint)
