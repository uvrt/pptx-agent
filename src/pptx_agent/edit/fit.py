"""Does the text fit?  As PowerPoint will show it, not as a renderer might shrink it.

**What is measured.**  The deck is resolved by pptx2svg (inheritance, theme fonts, the
geometry's text rectangle) and the text laid out by ooxml-common's text measurement --
the same advance widths, kerning, line heights and paragraph spacing pptx2svg wraps and
draws with (:mod:`._textlayout` is the one module that reaches into it).  Nothing is
rendered.

**What is not done: shrinking.**  PowerPoint computes a ``normAutofit`` scale when it
*edits* text and stores it as ``fontScale``; opening a file, it draws exactly what is
stored.  So a body with ``normAutofit`` and no ``fontScale`` -- every placeholder this
library or any generator fills -- is drawn full size and overflows when the text is too
long.  :func:`measure` lays the text out at the stored scale (1.0 when none), which is
what PowerPoint shows; a renderer that re-derives a scale to make the text fit shows
something PowerPoint does not.

**Agreement with PowerPoint** is the measurement's: pptx2svg's line breaks match
PowerPoint's to the word on its corpus, and the oracle test checks overflow verdicts
against PowerPoint's own PDF (tests/test_fit.py).  A shape within a line of the edge can
go either way; :attr:`TextFit.overflow` says by how much.

**Near the right edge a line can go either way too** (:data:`WRAP_MARGIN`).  Measured on
290 one-word text boxes in PowerPoint 16 for Mac (``tools/wrap_boundary_probe.py``: Aptos,
Calibri and Arial at 12, 18 and 24 pt, box widths from 0.3 pt under the measured width to
0.85 pt over it): PowerPoint keeps a word on its line when its advances, kerning included,
fit the text width -- the box less its left and right insets, which it subtracts exactly
-- but it draws each glyph at an advance up to about 0.05 pt off the font's.  It kerns
only the pairs of the face's legacy ``kern`` table, not the class pairs only its ``GPOS``
table has, and so does this measurement since ooxml-common 0.5
(``DrawingRules.kerning``): Aptos's ``s s`` pair is one, and "Pass" at 18 pt in a 650,000
EMU box breaks "Pas / s" here as in PowerPoint (trial 2, p4).  What is left is a band
about -0.05 to +0.16 pt wide around the edge: 260 of the 290 verdicts agree, and every
line PowerPoint wrapped that this measurement keeps had less than 0.16 pt to spare (the
most, 2,009 EMU), while PowerPoint kept a word that measured up to 0.05 pt too wide.  So
:attr:`TextFit.margin_to_wrap` reports the smallest room any line has left, and
:attr:`TextFit.near_wrap` whether that is under :data:`WRAP_MARGIN`, 0.16 pt.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING


def _engine():
    """:mod:`._textlayout`, imported when first needed: it needs ooxml-common, which comes
    with pptx2svg (``pptx-agent[render]``), and the package imports without it."""
    from . import _textlayout

    return _textlayout

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, Shape, Slide

#: Below this (EMU, about 0.1 pt) a difference is measurement noise, not an overflow.
_TOLERANCE = 1270
#: How close (EMU) a line may come to its frame's right edge before PowerPoint might break
#: it where this measurement does not: 0.16 pt, the measured band (see the module doc).
#: It was half a point while the measurement charged Aptos's GPOS-only pairs.
WRAP_MARGIN = 2032
#: A line crossing less of a text area than this (EMU, a point) only touches it.
_CROSSING = 12700


@dataclass(frozen=True)
class TextFit:
    """How a shape's text fits it, as PowerPoint draws it (:meth:`Shape.text_fit`).

    ``needed`` is the height, EMU, the text takes -- its lines, spacing and the frame's top
    and bottom insets -- and ``available`` the height of the frame's text area (the shape's,
    or its geometry's text rectangle); ``overflow`` is how much more is needed (0 when
    nothing).  ``overflows`` is not simply ``needed > available``: it allows the bottom of
    the last line box, which draws no ink -- a fifth of the tallest line, and never less
    than 1,270 EMU (0.1 pt) -- so a box sized tightly to one line is not reported, a line
    that does not fit is.  ``sizes`` is each paragraph's effective font sizes in points
    (stored autofit scale applied; a paragraph of one size is a one-element tuple, an
    empty one ``()``), ``lines`` how many lines each wraps to.  ``autofit`` is
    ``"none"``, ``"normal"`` or ``"shape"`` and ``font_scale`` the stored scale the text
    is drawn at.

    ``margin_to_wrap`` is the least room, EMU, any line leaves before the right edge of
    the text area (``None`` when the frame does not wrap or holds no text), and
    ``near_wrap`` says it is under :data:`WRAP_MARGIN` (0.16 pt): PowerPoint may then
    break that line where this measurement does not (see :mod:`.fit`).

    ``slack`` is that allowance, EMU: how far ``needed`` may pass ``available`` and the
    text still be drawn inside the frame.  So "fits while needed > available" is no
    contradiction: ``overflows`` is ``overflow > slack``.  ``needed - available`` up to
    ``slack`` is the empty bottom of the last line box (its descent and line gap), which
    PowerPoint draws no ink in.

    For example::

        fit = deck.shape("258.3").text_fit()
        if fit.overflows:
            print(f"{fit.overflow / 12700:.0f} pt too tall at {fit.smallest_size} pt")
        if fit.near_wrap:
            shape.width += 2 * 12700         # give the tightest line a little room
    """

    needed: int
    available: int
    overflows: bool
    overflow: int
    sizes: tuple[tuple[float, ...], ...]
    lines: tuple[int, ...]
    autofit: str
    font_scale: float
    margin_to_wrap: int | None = None
    slack: int = _TOLERANCE

    @property
    def smallest_size(self) -> float | None:
        """The smallest effective size of any paragraph, points; ``None`` without text.

        For example::

            assert shape.text_fit().smallest_size >= 18
        """
        sizes = [size for paragraph in self.sizes for size in paragraph]
        return min(sizes) if sizes else None

    @property
    def near_wrap(self) -> bool:
        """``True`` when some line is within :data:`WRAP_MARGIN` of wrapping.

        For example::

            assert not label.text_fit().near_wrap
        """
        return self.margin_to_wrap is not None and self.margin_to_wrap < WRAP_MARGIN


@dataclass(frozen=True)
class TextMeasure:
    """Text laid out in a width before any shape exists (:func:`measure_text`).

    ``lines`` is the text of each line as it breaks (a paragraph per ``"\\n"``, a forced
    break per ``"\\v"``), ``breaks`` the offset in the text where each line starts,
    ``height`` the height, EMU, a box needs for it -- lines, spacing and the top and
    bottom insets, exactly what :attr:`TextFit.needed` would say -- ``widest`` the widest
    line, EMU, and ``margin_to_wrap`` the least room any line leaves (see
    :attr:`TextFit.margin_to_wrap`).

    For example::

        m = measure_text("Order received", size=18, width=1371600, deck_or_shape=deck)
        m.lines, m.height                    # (('Order', 'received'), 640080)
    """

    lines: tuple[str, ...]
    breaks: tuple[int, ...]
    height: int
    widest: int
    margin_to_wrap: int | None
    #: What a box needs, EMU: :attr:`height` for a rectangle or text box; more for a
    #: geometry whose text area is a share of it (an ellipse, a chevron), as
    #: :meth:`Shape.fit_height` grows it.  Set when a text spec was measured.
    box_height: int | None = None
    #: How many lines each paragraph wraps to (a text spec's measurement).
    paragraph_lines: tuple[int, ...] = ()
    #: The insets measured with, EMU ``(left, top, right, bottom)``.
    insets: tuple[int, int, int, int] | None = None

    @property
    def line_count(self) -> int:
        """How many lines.

        For example::

            measure_text("Pass", size=18, width=650000, font="Aptos").line_count   # 2
        """
        return len(self.lines)

    @property
    def near_wrap(self) -> bool:
        """As :attr:`TextFit.near_wrap`.

        For example::

            measure_text("Pass", size=18, width=652000, font="Aptos").near_wrap   # True
        """
        return self.margin_to_wrap is not None and self.margin_to_wrap < WRAP_MARGIN


@dataclass(frozen=True)
class Overflow:
    """One problem :meth:`Document.overflows` found on a slide.

    ``kind`` is

    * ``"text"``: the shape's text is taller than its frame; ``amount`` EMU, ``fit`` the
      :class:`TextFit`;
    * ``"overlap"``: two things collide, ``shape`` and ``other``, and ``detail`` says how:
      ``"placeholder"`` -- the shape's bounds cross those of a title or another
      placeholder holding something; ``"text"`` -- two shapes holding text overlap
      (their drawn boxes, or for a text box with neither fill nor outline the text
      itself); ``amount`` is the overlap's area in EMU² -- ``"line"`` -- the line or
      connector ``shape`` crosses the text area of ``other``; ``amount`` is the length
      inside it, EMU -- or, only when asked for (``boxes=True``), ``"box"``: the boxes of
      two shapes holding text overlap although their text does not; ``amount`` is the
      area, EMU²;
    * ``"off_slide"``: the shape is drawn past the slide's edge by ``amount`` EMU.

    ``slide`` is the ``sldId``, ``shape`` the shape id.

    For example::

        for problem in deck.overflows():
            print(problem)          # [text] s:258 258.3: 2,214,600 EMU too tall
    """

    kind: str
    slide: int
    shape: str
    amount: int
    other: str | None = None
    fit: TextFit | None = None
    detail: str | None = None

    def __str__(self) -> str:
        if self.kind == "text":
            return f"[text] s:{self.slide} {self.shape}: {self.amount:,} EMU too tall"
        if self.kind == "overlap":
            if self.detail == "line":
                return (f"[overlap] s:{self.slide} line {self.shape} crosses the text of "
                        f"{self.other} for {self.amount:,} EMU")
            return f"[overlap] s:{self.slide} {self.shape} over {self.other}"
        return f"[off_slide] s:{self.slide} {self.shape}: {self.amount:,} EMU past the edge"


# -- the layout ------------------------------------------------------------------------------


class _Measured:
    """pptx2svg's resolved slide and a measuring context, for one slide's shapes."""

    def __init__(self, slide: "Slide") -> None:
        import pptx2svg

        document = slide.document
        options = pptx2svg.ConvertOptions(slide_numbers=[slide.index + 1],
                                          warn_on_font_substitution=False)
        model = pptx2svg.convert_pptx_to_model(document.to_bytes(), options)
        self.context = _engine().context(model.embedded_fonts.metrics)
        self._by_id: dict[str, list[tuple[object, tuple[float, float]]]] = {}
        self._walk(model.slides[0].elements, (1.0, 1.0))

    def _walk(self, elements, scale: tuple[float, float]) -> None:
        for element in elements:
            if getattr(element, "type", None) == "group":
                outer, inner = element.transform, element.child_transform
                sx = outer.extent_width / inner.extent_width if inner.extent_width else 1.0
                sy = outer.extent_height / inner.extent_height if inner.extent_height else 1.0
                self._by_id.setdefault(element.element_id or "", []).append((element, scale))
                self._walk(element.children, (scale[0] * sx, scale[1] * sy))
            else:
                self._by_id.setdefault(element.element_id or "", []).append((element, scale))

    def element(self, shape: "Shape", occurrence: int | None = None):
        if occurrence is None:
            occurrence = _occurrence(shape)
        found = self._by_id.get(f"{shape._slide.slide_id}.{_raw_id(shape)}", [])
        return found[min(occurrence, len(found) - 1)] if found else (None, (1.0, 1.0))


def _raw_id(shape: "Shape") -> str | None:
    from .ids import cnv_pr

    properties = cnv_pr(shape._element)
    return None if properties is None else properties.get("id")


def _occurrence(shape: "Shape") -> int:
    """How many shapes before this one on the slide, in document order, share its raw id."""
    from .ids import cnv_pr

    raw = _raw_id(shape)
    count = 0
    for element in shape._slide._sp_tree().iter():
        if element is shape._element:
            return count
        if _local(element) in ("sp", "pic", "cxnSp", "grpSp", "graphicFrame"):
            properties = cnv_pr(element)
            if properties is not None and properties.get("id") == raw:
                count += 1
    return count


def _local(element) -> str:
    tag = element.tag
    return tag.rpartition("}")[2] if isinstance(tag, str) else ""


@dataclass(frozen=True)
class _Laid:
    """A shape's text laid out: the :class:`TextFit`, and where the text is drawn in the
    shape's frame -- ``(x0, y0, x1, y1)`` in slide EMU before the frame's rotation, relative
    to its top-left corner -- with the text area it is laid out in, the same way."""

    fit: TextFit
    ink: tuple[float, float, float, float] | None
    area: tuple[float, float, float, float] | None
    lines: tuple[str, ...] = ()
    breaks: tuple[int, ...] = ()
    widest: int = 0
    insets: tuple[int, int, int, int] | None = None


def _lay_out(shape: "Shape", measured: _Measured | None = None) -> _Laid:
    if shape.kind != "shape":
        raise ValueError(f"{shape.id}: a {shape.kind} holds no text to fit")
    measured = measured or _Measured(shape._slide)
    element, (sx, sy) = measured.element(shape)
    frame_mode = shape.text_frame.autofit
    stored_scale = shape.text_frame.font_scale
    if element is None or getattr(element, "text_body", None) is None:
        height = shape.height or 0
        return _Laid(TextFit(0, height, False, 0, (), (), frame_mode, stored_scale), None, None)
    engine = _engine()
    context = measured.context
    properties = element.text_body.body_properties
    laid = engine.measure_shape_text(element, context, scale=(sx, sy), as_drawn=True)
    to_emu = engine.px_to_emu
    has_text = laid.has_text
    needed = round(to_emu(laid.needed)) if has_text else 0
    available = round(to_emu(laid.available))
    overflow = max(0, needed - available)
    margins = [paragraph.width - line.width for paragraph in laid.paragraphs
               for line in paragraph.lines if line.text]
    margin = round(to_emu(min(margins))) if laid.wraps and margins else None
    tallest = laid.tallest_line / engine.PX_PER_PT
    slack = _slack(tallest)
    fit = TextFit(needed, available, overflow > slack, overflow,
                  tuple(paragraph.sizes for paragraph in laid.paragraphs),
                  tuple(len(paragraph.lines) if paragraph.has_text else 0
                        for paragraph in laid.paragraphs),
                  frame_mode, stored_scale, margin, slack)

    # Where the text is: the area within the frame, the lines placed by alignment and the
    # block by the frame's anchor.  Only for horizontal text in one column.
    offset_x, offset_y = to_emu(laid.area_left), to_emu(laid.area_top)
    inset_left, inset_top, inset_right, inset_bottom = laid.insets
    area_rect = (offset_x + to_emu(inset_left), offset_y + to_emu(inset_top),
                 offset_x + to_emu(laid.width - inset_right),
                 offset_y + to_emu(laid.height - inset_bottom))
    ink = None
    text_px = laid.text_height
    if has_text and properties.vert == "horz" and laid.columns == 1:
        spans = []
        for source, paragraph in zip(element.text_body.paragraphs, laid.paragraphs):
            margin_left = engine.emu_to_px(source.properties.margin_left or 0)
            alignment = source.properties.alignment
            for line in paragraph.lines:
                if not line.text:
                    continue
                if alignment == "ctr":
                    start = margin_left + (paragraph.width - line.width) / 2
                elif alignment == "r":
                    start = margin_left + paragraph.width - line.width
                else:
                    start = margin_left
                spans.append((start, start + min(line.width, paragraph.width)
                              if laid.wraps else start + line.width))
        if spans:
            left = inset_left + min(span[0] for span in spans)
            right = inset_left + max(span[1] for span in spans)
            inner = laid.height - inset_top - inset_bottom
            anchor = properties.anchor
            if anchor == "ctr":
                top = inset_top + (inner - text_px) / 2
            elif anchor == "b":
                top = laid.height - inset_bottom - text_px
            else:
                top = inset_top
            # The last line box ends below the ink by about a fifth of a line (see
            # _slack): stacked labels whose line boxes touch do not collide.
            bottom = top + max(0.0, text_px - 0.2 * laid.tallest_line)
            ink = (offset_x + to_emu(left), offset_y + to_emu(top),
                   offset_x + to_emu(right), offset_y + to_emu(bottom))
    lines, breaks, base = [], [], 0
    for paragraph in laid.paragraphs:
        for line in paragraph.lines:
            lines.append(line.text)
            breaks.append(base + line.start)
        base += len(paragraph.text) + 1
    widest = round(to_emu(max((line.width for paragraph in laid.paragraphs
                               for line in paragraph.lines), default=0.0)))
    insets = tuple(round(to_emu(v)) for v in laid.insets)
    return _Laid(fit, ink, area_rect, tuple(lines), tuple(breaks), widest, insets)


def measure(shape: "Shape", measured: _Measured | None = None) -> TextFit:
    """See :meth:`Shape.text_fit`."""
    return _lay_out(shape, measured).fit


def _slack(line_pt: float) -> int:
    """How far text may run past its frame and still draw inside it, EMU: the part of the
    last line box below the ink -- a fifth of the tallest line, line spacing included --
    and never less than measurement noise."""
    return max(_TOLERANCE, round(0.2 * line_pt * 12700))


def fit_height(shape: "Shape") -> int:
    """See :meth:`Shape.fit_height`."""
    fit = measure(shape)
    height = shape.height or 0
    if fit.needed == 0:
        return height
    if fit.available <= 0 or fit.available == height:
        return height + fit.needed - fit.available
    # A geometry whose text rectangle is a share of the shape (an ellipse's, a chevron's)
    # grows with it in proportion.
    return round(height * fit.needed / fit.available)


# -- measuring before building ---------------------------------------------------------------


def measure_text(text, *, font: str | None = None, size: float | None = None,
                 width: int, line_spacing: float | None = None, deck_or_shape=None,
                 bold: bool = False, insets: tuple[int, int, int, int] | None = None,
                 wrap: bool = True, like=None, preset: str | None = None,
                 slide=None, height: int | None = None) -> TextMeasure:
    """Lay ``text`` out in a box ``width`` EMU wide before building anything, exactly as
    :meth:`Shape.text_fit` lays a shape's text out: its lines, where they break and the
    height a box needs (:class:`TextMeasure`).

    **One measuring model.**  ``text`` may be a :class:`~pptx_agent.TextSpec` --
    paragraphs, runs with their own sizes, bold and fonts, bullets, spacing, insets --
    the very value ``add_shape(text=...)`` and :meth:`Shape.set_text` build with.  Name
    what will hold it: ``like=shape`` (an existing shape: its font, size, insets and
    wrapping, a placeholder's inheritance included) or ``preset=`` (a new shape's
    defaults: ``"rect"``, ``"chevron"``..., or ``"textbox"`` for a text box), with the
    deck (or ``slide=``, the slide it will go on) as ``deck_or_shape``.  The spec is then
    built on a copy of the deck, in such a shape, and measured as :meth:`Shape.text_fit`
    measures it -- so what is measured is what is built.  ``box_height`` is the height to
    give the box (:func:`fit_box`).  Plain text with ``like`` or ``preset`` is measured
    the same way, its runs formatted by ``font``, ``size`` and ``bold``.

    Without ``like``/``preset``, plain ``text`` is laid out directly, faster:
    ``size`` is in points and ``font`` a typeface; ``bold=True`` measures it bold (a bold
    run is wider: pass it whenever the text will be bold); ``line_spacing`` a multiple of
    single spacing (``1.2``).  ``deck_or_shape`` supplies what is not given: a
    :class:`~pptx_agent.Document` or :class:`~pptx_agent.Slide` its theme's body font
    (the face a new text box draws in), a :class:`~pptx_agent.Shape` its first
    paragraph's effective font, size and line spacing, its insets and its wrapping -- so
    ``measure_text(new_text, width=shape.width, deck_or_shape=shape)`` answers "will this
    text fit that box?" before it is set.  ``width`` is the box's, insets included:
    ``insets`` are ``(left, top, right, bottom)`` EMU, PowerPoint's ``(91440, 45720,
    91440, 45720)`` by default.  ``"\n"`` starts a paragraph, ``"\v"`` breaks a line.
    Needs pptx2svg's measurement (``pptx-agent[render]``); a face the deck embeds is
    measured with the built-in tables, not its own.

    For example::

        from pptx_agent import measure_text
        m = measure_text("Order received", size=18, width=1371600, deck_or_shape=deck)
        box = slide.add_textbox(left, top, 1371600, m.height, "Order received",
                                autofit="none")
    """
    from .document import Document, Shape, Slide, _require_renderer
    from .textspec import ParagraphSpec, RunSpec, TextSpec, as_spec

    _require_renderer()
    spec = as_spec(text)
    if spec is not None or like is not None or preset is not None:
        if spec is None:
            run = {key: value for key, value in (("font", font), ("size", size),
                                                 ("bold", bold or None)) if value is not None}
            spec = TextSpec([ParagraphSpec([RunSpec(part, **run)] if part else [],
                                           line_spacing=line_spacing)
                             for part in str(text).split("\n")],
                            insets=insets, wrap=None if wrap is True else wrap)
        return _measure_spec(spec, width=width, like=like, preset=preset,
                             where=slide if slide is not None else deck_or_shape,
                             height=height)
    engine = _engine()
    east_asian = None
    if isinstance(deck_or_shape, Shape):
        shape = deck_or_shape
        if shape.kind != "shape":
            raise ValueError(f"{shape.id}: a {shape.kind} holds no text to measure like")
        frame = shape.text_frame
        first = frame.paragraph(0).effective if len(frame) else None
        font = font or (first.font if first else None)
        size = size if size is not None else (first.size if first else None)
        if line_spacing is None and first is not None:
            line_spacing = first.line_spacing
        insets = insets if insets is not None else frame.insets
        wrap = frame.wrap if wrap is True else wrap
        theme = shape._slide.theme
        east_asian = theme.fonts.minor_east_asian
    elif isinstance(deck_or_shape, (Document, Slide)):
        theme = deck_or_shape.theme
        font = font or theme.fonts.minor
        east_asian = theme.fonts.minor_east_asian
    elif deck_or_shape is not None:
        raise TypeError("deck_or_shape is a Document, a Slide or a Shape")
    if font is None:
        raise ValueError("name a font, or pass the deck (its theme's body font) or a shape")
    if size is None:
        raise ValueError("give a size, in points")
    if size <= 0 or int(width) <= 0:
        raise ValueError("size and width must be positive")
    left, top, right, bottom = insets if insets is not None else engine.DEFAULT_INSETS
    model = engine.model
    properties = model.RunProperties(font_size=float(size), font_family=font,
                                     font_family_ea=east_asian, bold=bool(bold))
    spacing = None if line_spacing is None else model.PercentSpacing(round(line_spacing * 100000))
    paragraphs = [model.Paragraph(runs=[model.TextRun(part.replace("\v", "\n"), properties)],
                                  properties=model.ParagraphProperties(line_spacing=spacing),
                                  end_para_run_properties=properties)
                  for part in str(text).split("\n")]
    body = model.BodyProperties(margin_left=left, margin_top=top, margin_right=right,
                                margin_bottom=bottom, wrap="square" if wrap else "none")
    laid = engine.measure_text_body(model.TextBody(paragraphs, body),
                                    model.Transform(extent_width=int(width)), engine.context(),
                                    as_drawn=True)
    # A line starts where the break before it ends: an empty line between two "\v" breaks
    # starts after the first, at the second -- where PowerPoint puts the caret on it.
    lines, breaks, base = [], [], 0
    for paragraph in laid.paragraphs:
        for line in paragraph.lines:
            lines.append(line.text)
            breaks.append(base + line.start)
        base += len(paragraph.text) + 1
    to_emu = engine.px_to_emu
    height = round(to_emu(laid.text_height)) + top + bottom if laid.has_text else 0
    widths = [line.width for line in laid.lines]
    margins = [paragraph.width - line.width for paragraph in laid.paragraphs
               for line in paragraph.lines if line.text]
    return TextMeasure(tuple(lines), tuple(breaks), height, round(to_emu(max(widths, default=0.0))),
                       round(to_emu(min(margins))) if wrap and margins else None)


def fit_box(spec, width: int, *, like=None, preset: str | None = None, deck_or_shape=None,
            slide=None) -> int:
    """The height, EMU, a box ``width`` EMU wide needs for ``spec`` (a
    :class:`~pptx_agent.TextSpec` or plain text), insets included: the
    :attr:`TextMeasure.box_height` of :func:`measure_text` -- as tall as
    :meth:`Shape.fit_height` would make the shape once built.  ``preset`` defaults to
    ``"rect"``.

    For example::

        height = fit_box(spec, Pt(160), preset="roundRect", deck_or_shape=slide)
        slide.add_shape("roundRect", left, top, Pt(160), height, text=spec)
    """
    if like is None and preset is None:
        preset = "rect"
    measured = measure_text(spec, width=width, like=like, preset=preset,
                            deck_or_shape=deck_or_shape, slide=slide)
    return measured.box_height if measured.box_height is not None else measured.height


def _scratch(document: "Document") -> "Document":
    """A copy of ``document`` to build measuring shapes in, kept while the document stays
    at the same version (its edits make a new one)."""
    from .document import Document

    version = getattr(document.history, "version", None)
    cached = document.__dict__.get("_measure_scratch")
    if cached is not None and version is not None and cached[0] == version:
        return cached[1]
    copy = Document.open(document.to_bytes())
    document.__dict__["_measure_scratch"] = (version, copy)
    return copy


def _height_for(fit: TextFit, height: int) -> int:
    """:func:`fit_height`'s answer from a fit already taken at ``height``."""
    if fit.needed == 0:
        return height
    if fit.available <= 0 or fit.available == height:
        return height + fit.needed - fit.available
    return round(height * fit.needed / fit.available)


def _converged_height(shape: "Shape", rounds: int = 4, precision: int = 6350) -> int:
    """The least height at which the shape's text fits (``needed <= available``), and the
    shape left at it.

    A rectangle's is :func:`fit_height` at once.  A geometry whose text area depends on
    its height -- an ellipse's grows with it, a chevron's narrows as its point deepens --
    is grown to :func:`fit_height` again until that stops changing; when it does not
    settle (a chevron can wrap differently at every height), the least fitting height is
    found by halving the interval, to ``precision`` EMU (half a point).
    """
    def fits_at(height: int) -> tuple[bool, TextFit]:
        shape.height = height
        fit = measure(shape)
        return fit.needed <= fit.available, fit

    start = shape.height or 0
    ok, fit = fits_at(start)
    low, high = (None, start) if ok else (start, None)
    box = _height_for(fit, start)
    for _ in range(rounds):
        if box <= 0 or (low is not None and box <= low) or (high is not None and box >= high):
            break
        ok, fit = fits_at(box)
        if ok:
            high = box
            nxt = _height_for(fit, box)
            if box - nxt <= 2:
                shape.height = box
                return box
            box = nxt
        else:
            low = box
            box = _height_for(fit, box)
    if high is None:
        grow = max(low or 0, 12700)
        for _ in range(12):
            grow *= 2
            ok, _ = fits_at(grow)
            if ok:
                high = grow
                break
        else:
            shape.height = grow
            return grow
    low = low if low is not None else 0
    while high - low > precision:
        middle = (low + high) // 2
        ok, _ = fits_at(middle)
        if ok:
            high = middle
        else:
            low = middle
    shape.height = high
    return high


def _measure_spec(spec, *, width: int, like=None, preset: str | None = None, where=None,
                  height: int | None = None) -> TextMeasure:
    from .document import Document, Shape, Slide
    from .textspec import apply_spec
    from ..oxml.xml import PRESET_GEOMETRIES

    if int(width) <= 0:
        raise ValueError("width must be positive")
    if like is not None:
        if not isinstance(like, Shape) or like.kind != "shape":
            raise ValueError("like is an autoshape or text box (kind 'shape') to measure in")
        document = like._slide.document
    elif isinstance(where, Shape):
        document = where._slide.document
    elif isinstance(where, Slide):
        document = where.document
    elif isinstance(where, Document):
        document = where
    else:
        raise ValueError("pass the deck or slide (deck_or_shape= or slide=) the text will "
                         "go into, or like= a shape")
    if like is None and preset not in PRESET_GEOMETRIES and preset != "textbox":
        raise ValueError(f"{preset!r} is not a preset geometry or 'textbox'; see "
                         "pptx_agent.PRESETS")
    copy = _scratch(document)
    tall = int(height) if height is not None else int(width)
    if like is not None:
        source = copy.shape(like.id)
        shape = source.duplicate()
        shape.width = int(width)
        if height is not None:
            shape.height = int(height)
        tall = shape.height
    else:
        if isinstance(where, Slide):
            target = copy.slide(where.slide_id)
        elif isinstance(where, Shape):
            target = copy.slide(where._slide.slide_id)
        elif copy.slides:
            target = copy.slides[0]
        else:
            target = copy.add_slide(copy.layouts[0])
        if preset == "textbox":
            shape = target.add_textbox(0, 0, int(width), tall, autofit="none")
        else:
            shape = target.add_shape(preset, 0, 0, int(width), tall)
    try:
        apply_spec(shape.text_frame, spec)
        laid = _lay_out(shape)
        box = _converged_height(shape)
        if height is None and shape.height != tall:
            laid = _lay_out(shape)        # as it will be: at the height it needs
    finally:
        shape.delete()
    fit = laid.fit
    return TextMeasure(laid.lines, laid.breaks, fit.needed, laid.widest, fit.margin_to_wrap,
                       box_height=box, paragraph_lines=fit.lines, insets=laid.insets)


# -- the deck --------------------------------------------------------------------------------


def _holds_something(shape: "Shape") -> bool:
    if shape.kind == "shape":
        return bool(shape.text.strip())
    return True


def _holds_text(shape: "Shape") -> bool:
    if shape.kind == "shape":
        return bool(shape.text.strip())
    if shape.kind == "graphic_frame" and shape.has_table:
        return any(cell.text.strip() for row in range(shape.table.rows)
                   for cell in (shape.table.cell(row, column)
                                for column in range(shape.table.columns)))
    return False


def _opaque(shape: "Shape", element) -> bool:
    """Is the shape drawn with something that hides what is behind it -- a fill, a
    picture, a table, a chart?"""
    if shape.kind in ("picture", "graphic_frame"):
        return True
    if shape.kind != "shape" or element is None:
        return False
    fill = getattr(element, "fill", None)
    return fill is not None and getattr(fill, "type", "none") != "none"


def _outlined(element) -> bool:
    outline = getattr(element, "outline", None)
    return outline is not None and getattr(outline, "fill", None) is not None


def _slide_rect(shape: "Shape", rect) -> tuple[float, float, float, float] | None:
    """A rectangle given relative to the shape's frame (slide EMU, before rotation),
    mapped onto the slide -- the box around it when the frame is turned."""
    from . import geometry

    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if None in (left, top, width, height) or rect is None:
        return None
    scale_x, scale_y = shape._parent_scale()
    x0, y0, x1, y1 = rect
    corners = [(left + x0 / scale_x, top + y0 / scale_y), (left + x1 / scale_x, top + y0 / scale_y),
               (left + x1 / scale_x, top + y1 / scale_y), (left + x0 / scale_x, top + y1 / scale_y)]
    frame = (float(left), float(top), float(width), float(height))
    points = geometry.to_slide(shape._element,
                               geometry._turn(corners, frame, shape._xfrm(create=False)))
    xs, ys = [p[0] for p in points], [p[1] for p in points]
    return min(xs), min(ys), max(xs), max(ys)


def _box(bounds) -> tuple[float, float, float, float]:
    left, top, width, height = bounds
    return left, top, left + width, top + height


def _intersection(a, b) -> tuple[float, float]:
    return min(a[2], b[2]) - max(a[0], b[0]), min(a[3], b[3]) - max(a[1], b[1])


def _inside(inner, outer) -> bool:
    return (inner[0] >= outer[0] - _TOLERANCE and inner[1] >= outer[1] - _TOLERANCE
            and inner[2] <= outer[2] + _TOLERANCE and inner[3] <= outer[3] + _TOLERANCE)


@dataclass
class _Item:
    shape: "Shape"
    order: int
    box: tuple[float, float, float, float]
    opaque: bool
    text: bool
    region: tuple[float, float, float, float] | None = None   # where its text shows
    area: tuple[float, float, float, float] | None = None     # its text area
    route: list | None = None


def _layered(a: _Item, b: _Item) -> bool:
    """Intended layering: one sits wholly on an opaque shape behind it -- a label on a
    bar, a title on a banner, text on a picture."""
    for upper, lower in ((a, b), (b, a)):
        if lower.opaque and lower.order < upper.order and _inside(upper.box, lower.box):
            return True
    return False


def _items(slide: "Slide", measured_ref: list, fits: dict) -> list[_Item]:
    from . import geometry

    items = []
    for order, shape in enumerate(_flatten(slide.shapes)):
        bounds = shape.drawn_bounds
        if bounds is None:
            continue
        element = None
        if shape.kind == "shape":
            if measured_ref[0] is None:
                measured_ref[0] = _Measured(slide)
            element, _ = measured_ref[0].element(shape)
        item = _Item(shape, order, _box(bounds), _opaque(shape, element), _holds_text(shape))
        if shape.kind == "shape" and item.text:
            laid = _lay_out(shape, measured_ref[0])
            fits[shape.id] = laid.fit
            visible = item.opaque or _outlined(element)
            ink = _slide_rect(shape, laid.ink)
            item.region = item.box if visible or ink is None else ink
            item.area = _slide_rect(shape, laid.area) or item.box
        elif item.text:
            item.region = item.area = item.box
        item.route = geometry.route(shape)
        items.append(item)
    return items


def _collisions(slide: "Slide", items: list[_Item], boxes: bool = False) -> list[Overflow]:
    out: list[Overflow] = []
    reported: set = set()

    def report(kind_detail, shape, other, amount):
        pair = frozenset((shape.id, other.id))
        if pair in reported:
            return
        reported.add(pair)
        out.append(Overflow("overlap", slide.slide_id, shape.id, round(amount),
                            other=other.id, detail=kind_detail))

    # A shape over a title or another placeholder holding something.
    anchors = [item for item in items
               if item.shape.placeholder is not None and _holds_something(item.shape)]
    for anchor in anchors:
        for item in items:
            shape = item.shape
            if item is anchor or shape.kind == "connector" or item.route is not None:
                continue
            if shape.placeholder is not None and not _holds_something(shape):
                continue
            # What a placeholder holds: its text where it is drawn, else its box.
            held = anchor.region if anchor.region is not None else anchor.box
            mine = item.region if item.text and item.region is not None else item.box
            dx, dy = _intersection(held, mine)
            if dx > _TOLERANCE and dy > _TOLERANCE and not _layered(anchor, item):
                report("placeholder", shape, anchor.shape, dx * dy)
    # Two shapes holding text.
    texts = [item for item in items if item.text and item.region is not None]
    for index, first in enumerate(texts):
        for second in texts[index + 1:]:
            dx, dy = _intersection(first.region, second.region)
            if dx > _CROSSING and dy > _CROSSING and not _layered(first, second):
                report("text", second.shape, first.shape, dx * dy)
    # Two text boxes whose boxes overlap though their text does not: labels set side by
    # side whose frames run into each other (the spike's p8 A1).
    if boxes:
        for index, first in enumerate(texts):
            for second in texts[index + 1:]:
                dx, dy = _intersection(first.box, second.box)
                if dx > _CROSSING and dy > _CROSSING and not _layered(first, second):
                    report("box", second.shape, first.shape, dx * dy)
    # Lines through text.
    from .geometry import length_inside

    for line in items:
        if line.route is None:
            continue
        for item in texts:
            if item is line:
                continue
            region = item.region if not item.opaque else item.area
            if region is None:
                continue
            if item.opaque and item.order > line.order:
                continue                      # an opaque shape in front hides the line
            crossed = length_inside(line.route, region)
            if crossed > _CROSSING:
                report("line", line.shape, item.shape, crossed)
    return out


def slide_problems(slide: "Slide", *, collisions_only: bool = False,
                   boxes: bool = False) -> list[Overflow]:
    """Every problem on one slide: text overflows, collisions and shapes off the slide;
    with ``boxes``, also text shapes whose boxes overlap (``detail="box"``)."""
    return slide_report(slide, collisions_only=collisions_only, boxes=boxes)[0]


def slide_report(slide: "Slide", *, collisions_only: bool = False,
                 boxes: bool = False) -> tuple[list[Overflow], dict[str, TextFit]]:
    """:func:`slide_problems`, and the :class:`TextFit` of every shape holding text, by
    shape id -- from one measurement of the slide (its near-wrap lines, its fits within
    the allowance), for a caller that reports both.

    For example::

        problems, fits = slide_report(slide)
        tight = [shape for shape, fit in fits.items() if fit.near_wrap]
    """
    width, height = slide.document.slide_size
    measured_ref: list = [None]
    fits: dict = {}
    items = _items(slide, measured_ref, fits)
    out: list[Overflow] = []
    if not collisions_only:
        for item in items:
            fit = fits.get(item.shape.id)
            if fit is not None and fit.overflows:
                out.append(Overflow("text", slide.slide_id, item.shape.id, fit.overflow, fit=fit))
    out += _collisions(slide, items, boxes)
    if not collisions_only:
        for item in items:
            x0, y0, x1, y1 = item.box
            past = max(-x0, -y0, x1 - width, y1 - height)
            if past > _TOLERANCE:
                out.append(Overflow("off_slide", slide.slide_id, item.shape.id, round(past)))
    return out, fits


def overflows(document: "Document", slides=None, *, boxes: bool = False) -> list[Overflow]:
    """See :meth:`Document.overflows`."""
    from ..outline.read import select

    out: list[Overflow] = []
    for slide in select(document, slides):
        out += slide_problems(slide, boxes=boxes)
    return out


def _flatten(shapes):
    """Every shape that is not a group, once each, back to front.  (``slide.shapes`` lists
    a group's members after it as well, so a member is met twice on the way down.)"""
    out, seen = [], set()

    def walk(items):
        for shape in items:
            if shape._element in seen:
                continue
            seen.add(shape._element)
            if shape.kind == "group":
                walk(shape.children)
            else:
                out.append(shape)

    walk(shapes)
    return out


__all__ = ["Overflow", "TextFit", "TextMeasure", "WRAP_MARGIN", "fit_box", "fit_height",
           "measure", "measure_text", "overflows", "slide_problems", "slide_report"]
