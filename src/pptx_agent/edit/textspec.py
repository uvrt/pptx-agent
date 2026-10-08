r"""A text spec: a shape's whole text -- paragraphs, runs, bullets and frame -- as one value.

One value says what a text body should hold, so the same value can build it
(:meth:`Shape.set_text <pptx_agent.Shape.set_text>`, ``add_shape(text=...)``,
``add_textbox(text=...)``) and measure it first (:func:`~pptx_agent.measure_text`,
:func:`~pptx_agent.fit_box`).  What is measured is what is built: both go through
:func:`apply_spec`, so a box sized from the measurement fits the text it is then given --
insets, sizes, bold and bullets included.

For example::

    from pptx_agent import TextSpec, ParagraphSpec, RunSpec

    spec = TextSpec([
        ParagraphSpec([RunSpec("Discover", bold=True, size=14)], align="center"),
        ParagraphSpec([RunSpec("Interviews, data audit")], bullet="bullet", level=1),
    ], insets=(Pt(7.2), Pt(3.6), Pt(7.2), Pt(3.6)), anchor="middle")
    box = slide.add_shape("rect", left, top, width, fit_box(spec, width, deck_or_shape=slide),
                          text=spec)

Lengths are EMU, as everywhere in the library; font sizes are points.  A property left
``None`` is not written: the text keeps what it has (on an existing shape) or inherits
(on a new one).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping, Sequence

#: The run properties a :class:`RunSpec` may set.
RUN_FIELDS = ("bold", "italic", "underline", "strike", "size", "font", "color", "hyperlink")
#: What a paragraph's ``bullet`` may be.
BULLETS = ("none", "bullet", "number")
#: What a frame's ``autofit`` may be.
AUTOFITS = ("none", "normal", "shape")


@dataclass(frozen=True)
class RunSpec:
    r"""One run: its text (``"\v"`` breaks the line) and the formatting it states.

    ``size`` is points; ``font`` a typeface; ``color`` a theme colour or ``#RRGGBB``;
    ``underline`` ``True`` or a style; ``hyperlink`` an address.  ``None`` states nothing.

    For example::

        RunSpec("Q3", bold=True, color="accent1")
    """

    text: str
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | str | None = None
    strike: bool | None = None
    size: float | None = None
    font: str | None = None
    color: str | None = None
    hyperlink: str | None = None

    def formatting(self) -> dict[str, Any]:
        """The stated run properties, as :meth:`Run.format <pptx_agent.Run.format>` takes
        them.

        For example::

            RunSpec("x", bold=True).formatting()     # {'bold': True}
        """
        return {name: getattr(self, name) for name in RUN_FIELDS
                if getattr(self, name) is not None}


@dataclass(frozen=True)
class ParagraphSpec:
    """One paragraph: its runs, and the paragraph properties it states.

    ``align`` is ``"left"``, ``"center"``, ``"right"`` or ``"justify"``; ``bullet``
    ``"none"``, ``"bullet"`` (a hanging ``•``) or ``"number"`` (1. 2. 3.); ``level`` 0-8;
    ``space_before``/``space_after`` EMU; ``line_spacing`` a multiple of single spacing.

    For example::

        ParagraphSpec([RunSpec("Interviews")], bullet="bullet", level=1)
    """

    runs: Sequence[RunSpec] = ()
    align: str | None = None
    bullet: str | None = None
    level: int | None = None
    space_before: int | None = None
    space_after: int | None = None
    line_spacing: float | None = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "runs", tuple(_run(run) for run in self.runs))
        if self.bullet is not None and self.bullet not in BULLETS:
            raise ValueError(f"bullet is one of {BULLETS}, not {self.bullet!r}")
        if self.level is not None and not 0 <= int(self.level) <= 8:
            raise ValueError("level must be between 0 and 8")

    @property
    def text(self) -> str:
        """The paragraph's text, its runs joined.

        For example::

            ParagraphSpec([RunSpec("a"), RunSpec("b")]).text     # 'ab'
        """
        return "".join(run.text for run in self.runs)


@dataclass(frozen=True)
class TextSpec:
    r"""A whole text body: paragraphs, and the frame properties it states.

    ``insets`` is ``(left, top, right, bottom)`` EMU; ``anchor`` ``"top"``, ``"middle"``
    or ``"bottom"``; ``wrap`` whether lines wrap; ``autofit`` ``"none"``, ``"normal"`` or
    ``"shape"``.  :meth:`from_text` makes one of plain text (``"\n"`` paragraphs, ``"\v"``
    line breaks) with one run formatting.

    For example::

        TextSpec.from_text("Plan\nBuild", size=14, bold=True)
    """

    paragraphs: Sequence[ParagraphSpec] = ()
    insets: tuple[int, int, int, int] | None = None
    anchor: str | None = None
    wrap: bool | None = None
    autofit: str | None = None

    def __post_init__(self) -> None:
        paragraphs = tuple(_paragraph(p) for p in self.paragraphs) or (ParagraphSpec(),)
        object.__setattr__(self, "paragraphs", paragraphs)
        if self.insets is not None:
            insets = tuple(int(v) for v in self.insets)
            if len(insets) != 4 or min(insets) < 0:
                raise ValueError("insets are four lengths, EMU: (left, top, right, bottom)")
            object.__setattr__(self, "insets", insets)
        if self.autofit is not None and self.autofit not in AUTOFITS:
            raise ValueError(f"autofit is one of {AUTOFITS}, not {self.autofit!r}")

    @classmethod
    def from_text(cls, text: str, **formatting: Any) -> "TextSpec":
        r"""A spec of plain text, every run formatted alike (``size=14``, ``bold=True``).

        For example::

            TextSpec.from_text("Kick-off\vNov 2", size=11)
        """
        unknown = set(formatting) - set(RUN_FIELDS)
        if unknown:
            raise TypeError(f"unknown run formatting {sorted(unknown)}; one of {RUN_FIELDS}")
        return cls([ParagraphSpec([RunSpec(part, **formatting)] if part else [])
                    for part in str(text).split("\n")])

    @property
    def text(self) -> str:
        r"""The plain text: paragraphs joined by ``"\n"``.

        For example::

            TextSpec.from_text("a\nb").text          # 'a\nb'
        """
        return "\n".join(paragraph.text for paragraph in self.paragraphs)


def _run(value: Any) -> RunSpec:
    if isinstance(value, RunSpec):
        return value
    if isinstance(value, str):
        return RunSpec(value)
    if isinstance(value, Mapping):
        return RunSpec(**dict(value))
    raise TypeError(f"a run is a RunSpec, a mapping or a string, not {type(value).__name__}")


def _paragraph(value: Any) -> ParagraphSpec:
    if isinstance(value, ParagraphSpec):
        return value
    if isinstance(value, str):
        return ParagraphSpec([RunSpec(value)] if value else [])
    if isinstance(value, Mapping):
        return ParagraphSpec(**dict(value))
    raise TypeError(f"a paragraph is a ParagraphSpec, a mapping or a string, not "
                    f"{type(value).__name__}")


def as_spec(value: Any) -> TextSpec | None:
    """``value`` as a :class:`TextSpec` when it is one (or a mapping with ``paragraphs``);
    ``None`` for plain text."""
    if isinstance(value, TextSpec):
        return value
    if isinstance(value, Mapping) and "paragraphs" in value:
        return TextSpec(**dict(value))
    return None


_ANCHORS = {"top": "top", "middle": "middle", "bottom": "bottom", "t": "top", "ctr": "middle",
            "b": "bottom", "center": "middle"}


def apply_spec(frame, spec: TextSpec) -> None:
    """Write ``spec`` into a :class:`~pptx_agent.TextFrame`: the text (each surviving
    character keeping its formatting, as ``set_text`` keeps it), cut into the spec's runs,
    each run's stated formatting, then each paragraph's, then the frame's.  One undo step
    when the caller holds a batch (the shape and document setters each are one)."""
    frame.set_text(spec.text)
    for index, paragraph_spec in enumerate(spec.paragraphs):
        paragraph = frame.paragraph(index)
        pieces: list[str] = []
        owners: list[RunSpec] = []
        for run in paragraph_spec.runs:
            for position, part in enumerate(run.text.split("\v")):
                if position:
                    pieces.append("\v")
                if part:
                    pieces.append(part)
                    owners.append(run)
        if pieces:
            paragraph.segment(pieces)
        for run, owner in zip(paragraph.runs, owners):
            stated = owner.formatting()
            if stated:
                run.format(**stated)
        if paragraph_spec.align is not None:
            paragraph.alignment = paragraph_spec.align
        if paragraph_spec.level is not None:
            paragraph.level = int(paragraph_spec.level)
        if paragraph_spec.space_before is not None:
            paragraph.space_before = _length(paragraph_spec.space_before)
        if paragraph_spec.space_after is not None:
            paragraph.space_after = _length(paragraph_spec.space_after)
        if paragraph_spec.line_spacing is not None:
            paragraph.line_spacing = float(paragraph_spec.line_spacing)
        if paragraph_spec.bullet == "bullet":
            paragraph.set_bullet()
        elif paragraph_spec.bullet == "number":
            paragraph.set_numbering()
        elif paragraph_spec.bullet == "none":
            paragraph.set_no_bullet()
    if spec.insets is not None:
        frame.insets = spec.insets
    if spec.anchor is not None:
        anchor = _ANCHORS.get(spec.anchor)
        if anchor is None:
            raise ValueError(f"anchor is top, middle or bottom, not {spec.anchor!r}")
        frame.anchor = anchor
    if spec.wrap is not None:
        frame.wrap = bool(spec.wrap)
    if spec.autofit is not None:
        frame.autofit = spec.autofit


def _length(value: Any) -> Any:
    """A spacing as the setter takes it: a :class:`~pptx_agent.Pt` stays one (so a small
    one is not mistaken for EMU), anything else is EMU."""
    from .units import Pt

    return value if isinstance(value, Pt) else int(round(float(value)))


__all__ = ["AUTOFITS", "BULLETS", "ParagraphSpec", "RUN_FIELDS", "RunSpec", "TextSpec",
           "apply_spec", "as_spec"]
