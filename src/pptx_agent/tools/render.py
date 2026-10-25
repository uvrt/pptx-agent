"""Rendering slides: pptx2svg behind the tool layer's render pipeline, and the serial renderer.

:data:`RENDERER` is the deck's :class:`ooxml_edit.tools.render.Renderer`: the pipeline
(:func:`ooxml_edit.tools.render.draw`, its worker tasks, the cache keys, the PNG checks) is
ooxml-edit's, the same for every format; this module says what is a deck's -- how its bytes
become each view's SVG (``pptx2svg.render_svgs``, a chunk of a call's slides in each worker
at once: a slide stands alone), what the pixels depend on besides the SVG
(:meth:`SlideRenderer.raster_fingerprint`), and how an SVG becomes a PNG.  It works from the
bytes it is given (``pptx2svg`` opens them itself); nothing is parsed into a
:class:`~pptx_agent.Document` or written back to bytes.  Its PNGs are
:func:`render_slides`' byte for byte: the same pptx2svg passes, one view at a time.

The functions here are module-level worker tasks (``worker_task``), so a spawned worker
can import them and a subprocess worker -- what a daemonic process, a Celery prefork child,
renders with -- may run them.
"""

from __future__ import annotations

import hashlib
import os
import tempfile
import types
import warnings
from typing import Any

from ooxml_edit.tools import render, worker_task
from ooxml_edit.tools.render import Layout, RendererBase, View


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


class SlideRenderer(RendererBase):
    """pptx2svg as the render pipeline calls it (:class:`ooxml_edit.tools.render.Renderer`):
    a slide stands alone, so each worker lays out and draws its own chunk of a call's views.

    A view is a 1-based slide number, or a layout's or master's part path (its template
    view).  The ``state`` a task is handed is the faces earlier tasks decoded from the
    deck's embedded fonts, by the hex digest of their payload -- which this worker takes as
    decoded (:func:`ooxml_common.fonts.embedded.remember_decoded`), so no worker decodes a
    face twice -- and its facts are ``{"fonts": {...}}``, the faces it decoded itself, for
    the caller to hand the next task.  An image's meta is pptx2svg's ``glyphs-missing``
    facts of the view, ``{"face", "script", "sample"}`` each."""

    name = "pptx-agent"
    whole_document = False
    #: The code that draws the pixels: a new install is a new key.
    modules = ("pptx2svg", "ooxml_common", "resvg_py")

    def lay_out(self, data: bytes, items: list, width: int, font_dirs: "list[str]",
                locale: Any, state: Any) -> Layout:
        import pptx2svg
        from ooxml_common.fonts import embedded
        from ooxml_common.locales import EN_US

        locale = EN_US if locale is None else locale
        embedded.remember_decoded({bytes.fromhex(digest): font
                                   for digest, font in (state or {}).items()})
        before = set(embedded.decoded_faces())
        numbers = sorted({item for item in items if isinstance(item, int)})
        parts = [item for item in items if isinstance(item, str)]
        extra = {"templates": parts} if parts else {}
        options = pptx2svg.ConvertOptions(slide_numbers=numbers, width=int(width), locale=locale,
                                          font_dirs=list(font_dirs), **extra)
        with warnings.catch_warnings():
            # The same facts come back with the images; the worker's stderr is nobody's to read.
            warnings.filterwarnings("ignore", message=".*no font this render loads")
            documents, resolved = pptx2svg.render_svgs(data, options)
            host_fonts = pptx2svg.uses_host_fonts(options)
            dirs = pptx2svg.user_font_dirs(options)
            fingerprint = self.raster_fingerprint(dirs, host_fonts, resolved.embedded_fonts,
                                                  options.check_glyphs)
        found = {}
        for document, slide in zip(documents, resolved.slides):
            ident = slide.template_part if slide.template_part else slide.slide_number
            found[ident] = View(ident, document, (slide.slide_number, slide.template_part))
        return Layout([found[item] for item in dict.fromkeys(items) if item in found],
                      fingerprint,
                      raster=(resolved.embedded_fonts, dirs, host_fonts, options.check_glyphs),
                      facts=before)

    def rasterise(self, views: "list[View]", raster: Any) -> "list[tuple[bytes, Any]]":
        import pptx2svg

        faces, dirs, host_fonts, check_glyphs = raster
        out = []
        with warnings.catch_warnings(), _embedded_files(faces) as files:
            warnings.filterwarnings("ignore", message=".*no font this render loads")
            for view in views:
                report = types.SimpleNamespace(check_glyphs=check_glyphs, warnings=[])
                png = pptx2svg.rasterise_svgs([view.svg], report, numbers=[view.info],
                                              font_dirs=dirs, font_files=files,
                                              host_fonts=host_fonts)[0]
                out.append((png, [{"face": w.detail.face, "script": w.detail.script,
                                   "sample": w.detail.sample} for w in report.warnings
                                  if getattr(w, "code", None) == "glyphs-missing"
                                  and getattr(w, "detail", None) is not None]))
        return out

    def facts(self, layout: Layout) -> dict:
        from ooxml_common.fonts import embedded

        return {"fonts": {digest.hex(): font for digest, font in embedded.decoded_faces().items()
                          if digest not in layout.facts}}

    def rasterise_tiles(self, svg: str, font_dirs: "list[str] | None") -> bytes:
        """A contact image's labels are in the bundle's sans-serif, the machine's fonts --
        and the faces PowerPoint draws with here -- left out, so the same sheet draws the
        same everywhere; a label a font cannot draw is not a deck's text, so it is not
        reported."""
        from pptx2svg.png import svg_to_png

        return svg_to_png(svg, font_dirs=font_dirs, host_fonts=False, check_glyphs=False)

    def warm_up(self) -> None:
        """The renderer's imports, the index of the faces PowerPoint draws with here, the
        rasteriser's own load of the system's fonts (once per process): one line of text."""
        from pptx2svg.png import svg_to_png

        svg_to_png('<svg xmlns="http://www.w3.org/2000/svg" width="64" height="16">'
                   '<text x="0" y="12" font-family="Calibri, sans-serif" font-size="10">'
                   'Aa</text></svg>')

    def raster_fingerprint(self, font_dirs: "list[str]", host_fonts: bool, embedded: Any,
                           check_glyphs: bool) -> str:
        """What a slide's pixels depend on besides its SVG
        (:func:`ooxml_edit.tools.render.fingerprint`): every font file the rasteriser may be
        handed -- the application's folders and the bundle, the faces PowerPoint draws with
        here (``host_fonts``), the system's where they take part -- each with its time and
        size; the deck's embedded faces, by digest; whether missing glyphs are reported."""
        from ooxml_common.fonts.rasterdb import file_stamps
        from pptx2svg.fonts import font_dirs as bundled_font_dirs

        bundled = [str(path) for path in bundled_font_dirs()]
        # svg_to_png skips the system's fonts with a bundle.
        system = render.system_fonts() if not bundled else None
        host: Any = None
        if host_fonts:
            from pptx2svg.fonts import office

            host = [file_stamps((), [str(path)]) for _, path in office.search_dirs()]
            emoji = office.emoji_face()
            host.append(getattr(emoji, "path", None) and str(emoji.path))
        faces = [(face.family, face.bold, face.italic, hashlib.sha256(face.source).hexdigest())
                 for face in getattr(embedded, "faces", ())]
        return render.fingerprint(self, bool(check_glyphs),
                                  file_stamps((), [*map(str, font_dirs), *bundled]), system,
                                  host, faces, os.environ.get("PPTX2SVG_OFFICE_FONTS"))


#: The deck's renderer, handed to the pipeline's worker tasks.
RENDERER = SlideRenderer()


@worker_task
def warm_up() -> None:
    """A new worker's first task, before any call's (:func:`ooxml_edit.tools.render.warm_up`):
    what the first render would pay for, and the stamp of the renderers' code."""
    render.warm_up(RENDERER)


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
