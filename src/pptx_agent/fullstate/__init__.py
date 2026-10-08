"""The full-state SVG: a slide's picture that also carries the slide, losslessly.

:func:`emit_full_state` takes pptx2svg's rendering and adds the ``data-ooxml-*`` vocabulary
to every shape's group -- typed attributes for what an agent edits, read from the
*unresolved* document (a theme colour stays ``data-ooxml-fill-scheme="accent1"``), and the
shape's own OOXML as base64 (``data-ooxml-xml``), the floor that keeps the SVG lossless for
everything the typed attributes do not cover.  :func:`apply_full_state` reads one back into
the open document as ordinary, undoable edits.

The vocabulary is PowerPoint's, so it lives here and not in the format-neutral core.  It is
documented attribute by attribute in ``ROADMAP.md`` (Phase E3).

Entry points: ``Slide.render_svg(full_state=True)`` and ``Document.apply_svg(svg)``.
"""

from .apply import ApplyReport, apply_full_state
from .emit import emit_full_state
from .model import PREFIX, VOCABULARY
from .safe import DEFAULT_LIMITS, FullStateError, Limits, NotFullStateSvg, UnsafeInput

__all__ = [
    "ApplyReport",
    "DEFAULT_LIMITS",
    "FullStateError",
    "Limits",
    "NotFullStateSvg",
    "PREFIX",
    "UnsafeInput",
    "VOCABULARY",
    "apply_full_state",
    "emit_full_state",
]
