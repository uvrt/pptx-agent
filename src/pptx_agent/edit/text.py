"""Text: ``TextFrame`` -> ``Paragraph`` -> ``Run``, live views over ``a:txBody``.

**Addressing is positional and re-resolved on every call.**  Runs and paragraphs have no
identity in OOXML -- no id attribute, nothing that survives PowerPoint re-saving the file -- so
pretending otherwise would be a lie the caller pays for later.  ``"256.5/p1/r0"`` means "the
first run of the second paragraph of shape 256.5", looked up afresh each time it is used, and
a :class:`Run` object holds exactly that address, never the element.  Inserting a run before
it therefore makes the same object refer to the run that is now in that position; that is the
honest semantics, and it is what lets a facade survive undo.

**What counts as a run.**  ``r<n>`` counts ``a:r`` and ``a:fld`` (a field such as the slide
number is a run whose text PowerPoint recomputes).  ``a:br`` is a line break, not a run; in
text it reads as ``"\\v"`` (vertical tab), which is also how to write one.  ``"\\n"`` separates
paragraphs.

**Replacing text keeps mixed formatting.**  :meth:`TextFrame.set_text` (and so
``Shape.set_text``) diffs the new text against the old -- paragraphs first, then characters
within each changed paragraph -- and gives every new character the formatting of the old
character it replaces or follows.  Changing "Revenue grew **12%**" to "Revenue grew **15%**"
leaves the figure bold; a paragraph whose text did not change is not touched at all.
Retyping a box completely still behaves like E0's ``set_text``: everything takes the first
run's formatting.
"""

from __future__ import annotations

import os
import re
import sys
import warnings
from dataclasses import dataclass
from contextlib import contextmanager
from typing import Any, Callable, ContextManager, Iterator, Protocol

from ..oxml.xml import (
    FILL_TAGS,
    Element,
    get_int,
    local_name,
    make,
    qn,
    remove,
    replace_choice,
    set_attr,
    subelement,
)
from ..oxml.package import REL_HYPERLINK, REL_SLIDE
# The paragraph rewriting below a TextFrame -- runs, breaks and fields, the text diff that
# keeps mixed formatting -- is DrawingML a chart title and a diagram node share; it lives in
# ooxml-edit's charts subpackage and is imported back under the names this module used.
from ooxml_edit.charts.dmltext import (  # noqa: F401  (re-exported)
    LINE_BREAK as _LINE_BREAK,
    RUN_TAGS as _RUN_TAGS,
    inline_items as _inline_items,
    make_break as _make_break,
    make_run as _make_run,
    paragraph_like as _paragraph_like,
    paragraph_text as _paragraph_text,
    rebuild_paragraph as _rebuild_paragraph,
    replace_body_text as _replace_body_text,
    resegment as _resegment,
    runs as _runs,
)
from .color import Color
from .fill import read_fill, solid_fill

if False:  # pragma: no cover - for type checkers only
    from .effective import EffectiveParagraph

#: ``a:hlinkClick@action`` for a jump to another slide of the same deck.
SLIDE_JUMP_ACTION = "ppaction://hlinksldjump"


# ------------------------------------------------------------------------------------------
# Text that was copied out of the Markdown outline
# ------------------------------------------------------------------------------------------


class MarkdownEscapeWarning(UserWarning):
    r"""Text given to ``set_text`` (or a ``text`` setter) looks copied from
    :meth:`Document.to_outline`: it has a Markdown escape such as ``\*`` or ``1\.`` at the
    start of a line, or ``**bold**`` markup, that the text it replaces does not have.  The
    text is written as given -- the warning only says so.  Take text to edit from
    ``shape.text``, :meth:`Document.outline_blocks` or :meth:`Document.find_text`, which are
    raw.  Silence it, when the backslashes are meant, with::

        warnings.simplefilter("ignore", pptx_agent.MarkdownEscapeWarning)
    """


#: What :func:`pptx_agent.outline.markdown.escape` writes anywhere in a line: a backslash
#: before ``* _ ` [ ] < ~ | &``.  Not ``\\``: a UNC path or a regular expression on a
#: slide is real text more often than an escaped backslash is a mistake.
_ESCAPED = re.compile(r"\\[*_`\[\]<~|&]")
#: What it writes at the start of a line only: ``\#``, ``\>``, ``\+``, ``\-``, ``\=``,
#: and ``1\.`` / ``1\)``.
_ESCAPED_LINE_START = re.compile(r"^[ \t]*(\\[#>+=-]|\d{1,9}(\\[.)]))", re.MULTILINE)
#: ``**bold**`` around a word, as the outline marks a bold run.
_BOLD_MARKUP = re.compile(r"(?<![\w*])\*\*[^\s*](?:[^*]*[^\s*])?\*\*(?![\w*])")


def outline_escapes(text: str, old: str = "") -> list[str]:
    """The pieces of ``text`` that look like the outline's Markdown -- escapes and ``**``
    markup -- and that ``old`` (the text being replaced) does not already contain.

    The check is deliberately narrow, so real text does not trigger it: a backslash only
    counts before the characters the outline escapes, and ``\\.`` and ``\\+`` only where
    the outline writes them, at the start of a line after a number or alone.  A Windows
    path (``C:\\Users``) or ``\\server`` passes; so does any escape already in the text.
    """
    found: list[str] = []
    for match in _ESCAPED.finditer(text):
        found.append(match.group(0))
    for match in _ESCAPED_LINE_START.finditer(text):
        found.append(match.group(1).lstrip())
    for match in _BOLD_MARKUP.finditer(text):
        found.append(match.group(0))
    seen: list[str] = []
    for piece in found:
        key = piece if piece.startswith("**") else (piece[-2:])
        if key not in old and piece not in seen:
            seen.append(piece)
    return seen


def _warn_if_escaped(value: str, old: str, address: str) -> None:
    pieces = outline_escapes(value, old)
    if not pieces:
        return
    shown = ", ".join(repr(p) for p in pieces[:3])
    warnings.warn(
        f"{address}: the new text has Markdown from to_outline ({shown}); it is written "
        f"as given, backslashes and asterisks included.  For raw text to edit from use "
        f"shape.text, deck.outline_blocks() or deck.find_text().",
        MarkdownEscapeWarning, stacklevel=_caller_level())


_PACKAGE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _caller_level() -> int:
    """The ``stacklevel`` that names the first frame outside pptx-agent and ooxml-edit."""
    level = 1
    frame = sys._getframe(1)
    while frame is not None:
        filename = os.path.abspath(frame.f_code.co_filename)
        if not (filename.startswith(_PACKAGE_DIR + os.sep) or "ooxml_edit" in filename):
            return level
        frame = frame.f_back
        level += 1
    return level


@dataclass(frozen=True)
class Hyperlink:
    """Where a run's click goes: a web or mail ``address``, or a ``slide_id`` in this deck.

    ``action`` is set for PowerPoint's other click actions (``ppaction://hlinkshowjump?
    jump=nextslide`` and the like), which have no address.

    For example::

        run.hyperlink                                # Hyperlink(address="https://...", ...)
    """

    address: str | None = None
    slide_id: int | None = None
    tooltip: str | None = None
    action: str | None = None

#: Friendly alignment names -> ``ST_TextAlignType``.
ALIGNMENTS: dict[str, str] = {
    "left": "l",
    "center": "ctr",
    "right": "r",
    "justify": "just",
    "distributed": "dist",
    "justify_low": "justLow",
    "thai_distributed": "thaiDist",
}
_ALIGNMENT_NAMES = {value: key for key, value in ALIGNMENTS.items()}

#: ``ST_TextUnderlineType``.
UNDERLINES = frozenset({
    "none", "words", "sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy",
    "dashLong", "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy",
    "wavy", "wavyHeavy", "wavyDbl",
})

#: ``ST_TextAutonumberScheme`` -- the common ones; any valid scheme is accepted.
NUMBERING_SCHEMES = frozenset({
    "alphaLcParenBoth", "alphaUcParenBoth", "alphaLcParenR", "alphaUcParenR", "alphaLcPeriod",
    "alphaUcPeriod", "arabicParenBoth", "arabicParenR", "arabicPeriod", "arabicPlain",
    "romanLcParenBoth", "romanUcParenBoth", "romanLcParenR", "romanUcParenR", "romanLcPeriod",
    "romanUcPeriod", "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain",
    "arabicDbPeriod", "arabicDbPlain", "ea1ChsPeriod", "ea1ChsPlain", "ea1ChtPeriod",
    "ea1ChtPlain", "ea1JpnChsDbPeriod", "ea1JpnKorPlain", "ea1JpnKorPeriod", "arabic1Minus",
    "arabic2Minus", "hebrew2Minus", "thaiAlphaPeriod", "thaiAlphaParenR", "thaiAlphaParenBoth",
    "thaiNumPeriod", "thaiNumParenR", "thaiNumParenBoth", "hindiAlphaPeriod",
    "hindiNumPeriod", "hindiNumParenR", "hindiAlpha1Period",
})

_BULLET_TAGS = ("a:buNone", "a:buAutoNum", "a:buChar", "a:buBlip")

#: Friendly anchor names -> ``ST_TextAnchoringType``.
ANCHORS: dict[str, str] = {"top": "t", "middle": "ctr", "bottom": "b",
                           "justified": "just", "distributed": "dist"}
_ANCHOR_NAMES = {value: key for key, value in ANCHORS.items()}
#: PowerPoint's insets when nothing states them, EMU: left, top, right, bottom.
_DEFAULT_INSETS = (91440, 45720, 91440, 45720)

SPACING_UNIT = 100  # spcPts: hundredths of a point
EMU_PER_SPACING_UNIT = 127  # 12,700 EMU per point / 100
PERCENT_UNIT = 100000  # spcPct, buSzPct: thousandths of a percent
SIZE_UNIT = 100  # sz: hundredths of a point


class TextHost(Protocol):
    """What a text frame needs from whatever holds it -- a shape, or a table cell."""

    address: str

    def _text_body(self, create: bool) -> Element | None: ...

    def _before_change(self) -> None: ...

    def _after_change(self) -> None: ...

    def _batch(self) -> ContextManager[None]: ...


@contextmanager
def _editing(host: TextHost) -> Iterator[None]:
    """One undo step: batch, checkpoint, edit, mark the part dirty."""
    with host._batch():
        host._before_change()
        yield
        host._after_change()


@dataclass(frozen=True)
class Bullet:
    """A paragraph's explicit bullet.

    ``kind`` is ``"none"`` (``a:buNone``), ``"char"``, ``"number"`` or ``"picture"``.

    For example::

        paragraph.bullet                             # Bullet(kind="char", char="•", ...)
    """

    kind: str
    char: str | None = None
    scheme: str | None = None
    start_at: int | None = None
    font: str | None = None
    color: Color | None = None
    #: Bullet size as a fraction of the text size.
    size: float | None = None


#: PowerPoint's hanging indent for a bullet, by font size: ``(up to this size, points;
#: indent, EMU)``.  Measured: a text box at each of 8 to 96 pt, bulleted through
#: PowerPoint 16 for Mac's AppleScript (``bullet type unnumbered``), saved and read back --
#: it wrote ``indent`` = -171,450 (3/16 in) up to 12 pt, -285,750 (5/16 in) from 12.5 to 18,
#: -342,900 from 18.5 to 25, -457,200 from 26 to 35, -571,500 from 36 to 44, -685,800
#: from 45 to 55, -857,250 from 57 to 72 and -1,143,000 at 80 and 96.  The step between two
#: sizes measured either side of it is placed at the lower one.
BULLET_HANGING = ((12.0, 171450), (18.0, 285750), (25.0, 342900), (35.0, 457200),
                  (44.0, 571500), (55.0, 685800), (72.0, 857250), (float("inf"), 1143000))


def bullet_hanging(size: float) -> int:
    """The hanging indent, EMU, PowerPoint gives a bullet at ``size`` points
    (:data:`BULLET_HANGING`).

    For example::

        bullet_hanging(18)                           # 285750, 0.3125 in
    """
    return next(amount for limit, amount in BULLET_HANGING if size <= limit)


def _spacing_value(emu, what: str) -> float:
    """A spacing given in EMU (checked: :func:`.units.spacing`), as the points
    ``a:spcPts`` stores."""
    from .units import spacing

    return spacing(emu, what, stacklevel=4) / (SPACING_UNIT * EMU_PER_SPACING_UNIT)


#: What ``format(**formatting)`` -- on a run, a paragraph or a frame -- and
#: ``add_run``/``add_paragraph``'s ``**formatting`` take: each is the :class:`Run` property
#: of that name (``font`` is ``typeface``'s other name).
FORMAT_KEYS = frozenset({"bold", "italic", "underline", "strike", "size", "typeface", "font",
                         "color", "hyperlink"})
_FORMAT_HELP = ("bold, italic, strike (True, False or None to inherit); underline (True, "
                "a style such as 'dbl', or None); size (points, or Pt); typeface or font "
                "(a face, or '+mn-lt'); color (a colour string or Color); hyperlink (an "
                "address, a Hyperlink, or None)")


def check_format_keys(formatting: dict) -> None:
    """Refuse a formatting key no run property has, naming the ones there are."""
    unknown = set(formatting) - FORMAT_KEYS
    if unknown:
        raise TypeError(f"unknown run formatting {sorted(unknown)}; the keys are "
                        f"{_FORMAT_HELP}.  Paragraph properties (alignment, space_after, "
                        f"bullets...) are set on the paragraph itself")


# ------------------------------------------------------------------------------------------
# TextFrame
# ------------------------------------------------------------------------------------------


class TextFrame:
    """The text of a shape or table cell: an ordered list of paragraphs.

    For example::

        frame = deck.shape("256.5").text_frame
        frame.paragraph(1).run(0).bold = True
    """

    def __init__(self, resolve: Callable[[], TextHost], address: str) -> None:
        self._resolve = resolve
        self.address = address

    # -- structure -------------------------------------------------------------------------

    @property
    def paragraphs(self) -> list["Paragraph"]:
        """Every paragraph, in order.

        For example::

            [paragraph.text for paragraph in frame.paragraphs]
        """
        return [Paragraph(self, index) for index in range(len(self._paragraph_elements()))]

    def paragraph(self, index: int) -> "Paragraph":
        """Paragraph ``index`` (negative counts from the end).

        For example::

            frame.paragraph(-1).alignment = "right"
        """
        count = len(self._paragraph_elements())
        if not -count <= index < count:
            raise IndexError(f"{self.address} has {count} paragraphs; no p{index}")
        return Paragraph(self, index % count)

    def __len__(self) -> int:
        return len(self._paragraph_elements())

    @property
    def text(self) -> str:
        """Paragraphs joined by ``"\\n"``; line breaks inside a paragraph read as ``"\\v"``.

        For example::

            frame.text                                   # 'line one\\nline two'
        """
        return "\n".join(_paragraph_text(p) for p in self._paragraph_elements())

    @text.setter
    def text(self, value: str) -> None:
        self.set_text(value)

    def set_text(self, value: "str | TextSpec") -> "TextFrame":
        r"""Replace the text, keeping each surviving character's formatting (see the
        module doc): ``"\n"`` separates paragraphs, ``"\v"`` is a line break::

            frame.set_text("Revenue grew 15%\nOperating margin 12.1%")

        Or a :class:`~pptx_agent.TextSpec`: paragraphs, runs, bullets and frame
        properties in one value, applied as one undo step -- the value
        :func:`~pptx_agent.measure_text` measures.

        The text is written as given; text that looks copied from
        :meth:`Document.to_outline` (Markdown escapes, ``**bold**``) draws a
        :class:`MarkdownEscapeWarning`.
        """
        from .textspec import apply_spec, as_spec

        spec = as_spec(value)
        if spec is not None:
            with self._resolve()._batch():
                apply_spec(self, spec)
            return self
        value = str(value)
        _warn_if_escaped(value, self.text, self.address)
        host = self._resolve()
        with _editing(host):
            body = host._text_body(True)
            assert body is not None
            _replace_body_text(body, value)
        return self

    def add_paragraph(self, text: str = "", *, index: int | None = None,
                      like: int | None = None, **formatting: Any) -> "Paragraph":
        """Insert a paragraph at ``index`` (default: at the end).

        Paragraph and run formatting are copied from paragraph ``like`` -- by default the one
        before the insertion point, or the first one when inserting at the start -- the way
        pressing Enter in PowerPoint continues the current paragraph's style.  ``formatting``
        is applied to the new runs afterwards (``bold=True``, ``size=18``...).

        For example::

            frame.add_paragraph("速報値", like=0, bold=True)
        """
        if "\n" in text:
            raise ValueError("one paragraph at a time; use '\\v' for a line break")
        check_format_keys(formatting)
        host = self._resolve()
        with _editing(host):
            body = host._text_body(True)
            assert body is not None
            paragraphs = body.findall(qn("a:p"))
            count = len(paragraphs)
            position = count if index is None else _clamp_index(index, count + 1, self.address)
            if like is None:
                like = position - 1 if position > 0 else 0
            template = paragraphs[like] if paragraphs and 0 <= like < count else None
            new = _paragraph_like(template, text)
            if position < count:
                paragraphs[position].addprevious(new)
            elif paragraphs:
                paragraphs[-1].addnext(new)
            else:
                subelement(body, "a:bodyPr")
                subelement(body, "a:lstStyle")
                body.append(new)
            paragraph = Paragraph(self, position)
            for run in paragraph.runs if formatting else ():
                run._apply(formatting)
        return paragraph

    def delete_paragraph(self, index: int) -> "TextFrame":
        """Remove paragraph ``index``; returns the frame.

        A text body must keep at least one paragraph, so deleting the last one empties it
        instead (keeping its paragraph properties and end-of-paragraph formatting).

        For example::

            frame.delete_paragraph(1)
        """
        host = self._resolve()
        paragraphs = self._paragraph_elements(host)
        index = _clamp_index(index, len(paragraphs), self.address)
        with _editing(host):
            target = paragraphs[index]
            if len(paragraphs) == 1:
                for child in list(target):
                    if child.tag not in {qn("a:pPr"), qn("a:endParaRPr")}:
                        remove(child)
            else:
                remove(target)
        return self

    @property
    def autofit(self) -> str:
        """How the text fits its shape, as it applies -- the frame's own setting or the one
        it inherits from its layout and master: ``"none"`` (text may overflow),
        ``"normal"`` (shrink text on overflow) or ``"shape"`` (resize the shape to fit
        the text).  Setting it writes the frame's own; ``None`` inherits again.

        PowerPoint shrinks text only when it edits it, and stores the result as
        :attr:`font_scale`; a file says ``"normal"`` with no scale until then, and is drawn
        at full size.  Setting ``"normal"`` keeps a scale already stored.

        For example::

            frame.autofit = "none"                       # never shrink: check text_fit()
        """
        from . import effective

        return effective.autofit(self._host()).mode

    @autofit.setter
    def autofit(self, value: str | None) -> None:
        self.set_autofit(value)

    def set_autofit(self, mode: str | None, *, font_scale: float | None = None) -> "TextFrame":
        """Set :attr:`autofit` (and, for ``"normal"``, the stored :attr:`font_scale`);
        returns the frame.

        For example::

            frame.set_autofit("normal", font_scale=0.9)
        """
        from . import effective

        host = self._host()
        if mode is not None and mode not in effective.AUTOFIT_NAMES:
            raise ValueError(f"autofit is one of {sorted(effective.AUTOFIT_NAMES)} or None")
        if font_scale is not None and mode != "normal":
            raise ValueError("a font scale belongs to autofit 'normal'")
        with _editing(host):
            body = host._text_body(True)
            assert body is not None
            effective.write_autofit(body, mode, font_scale, keep_scale=font_scale is None)
        return self

    @property
    def font_scale(self) -> float:
        """The stored ``normAutofit`` scale PowerPoint draws the text at (``0.9`` for 90%);
        ``1.0`` when none is stored or the autofit is not ``"normal"``.  Setting it sets
        :attr:`autofit` to ``"normal"`` with that scale; ``1.0`` removes it.

        For example::

            frame.font_scale = 1.0                       # draw at full size again
        """
        from . import effective

        return effective.autofit(self._host()).font_scale

    @font_scale.setter
    def font_scale(self, value: float) -> None:
        self.set_autofit("normal", font_scale=float(value))

    # -- the frame: insets, anchor, wrapping -----------------------------------------------

    @property
    def insets(self) -> tuple[int, int, int, int]:
        """The room between the frame's edges and its text, EMU: ``(left, top, right,
        bottom)``, as it applies -- the frame's own, else its layout's and master's, else
        PowerPoint's defaults ``(91440, 45720, 91440, 45720)`` (0.1 in at the sides, 0.05
        in above and below).  A table cell's are its margins (``a:tcPr@marL``...).
        Setting a 4-tuple writes the frame's own; ``None`` inherits again.  To change one
        side: :meth:`set_insets`.

        For example::

            frame.insets = (0, 0, 0, 0)          # text to the edges: a label sized exactly
        """
        values = []
        for name, default in zip(self._inset_names(), _DEFAULT_INSETS):
            raw = self._frame_attribute(name)
            values.append(default if raw is None else int(raw))
        return tuple(values)  # type: ignore[return-value]

    @insets.setter
    def insets(self, value: "tuple[int, int, int, int] | None") -> None:
        if value is None:
            self._set_frame_attributes({name: None for name in self._inset_names()})
            return
        values = tuple(value)
        if len(values) != 4:
            raise ValueError("insets are (left, top, right, bottom), EMU")
        self.set_insets(*values)

    def set_insets(self, left: int | None = None, top: int | None = None,
                   right: int | None = None, bottom: int | None = None) -> "TextFrame":
        """Set some of the insets, EMU, keeping the others as they apply; returns the frame.

        For example::

            frame.set_insets(left=0, right=0)
        """
        given = {"left": left, "top": top, "right": right, "bottom": bottom}
        current = dict(zip(("left", "top", "right", "bottom"), self.insets))
        changes = {}
        for (side, value), name in zip(given.items(), self._inset_names()):
            if value is None:
                if self._frame_attribute(name, own=True) is None and \
                        self._frame_attribute(name) is not None:
                    changes[name] = str(current[side])     # keep what applies now
                continue
            value = int(value)
            if value < 0:
                raise ValueError("an inset cannot be negative")
            changes[name] = str(value)
        self._set_frame_attributes(changes)
        return self

    @property
    def anchor(self) -> str:
        """Where the text sits between the top and bottom insets: ``"top"``,
        ``"middle"`` or ``"bottom"`` (``"justified"`` and ``"distributed"`` too, when a
        deck says so), as it applies -- the frame's own or inherited, ``"top"`` when
        nothing says.  Settable; ``"t"``, ``"ctr"`` and ``"b"`` are accepted, ``None``
        inherits again.

        For example::

            frame.anchor = "middle"              # centre a label vertically in its box
        """
        raw = self._frame_attribute("anchor")
        return _ANCHOR_NAMES.get(raw or "t", raw or "top")

    @anchor.setter
    def anchor(self, value: str | None) -> None:
        if value is not None:
            if value not in ANCHORS and value not in _ANCHOR_NAMES:
                raise ValueError(f"anchor is one of {sorted(ANCHORS)} (or 't', 'ctr', 'b')")
            value = ANCHORS.get(value, value)
        self._set_frame_attributes({"anchor": value})

    @property
    def wrap(self) -> bool:
        """Whether lines wrap at the frame's width (``a:bodyPr@wrap`` ``"square"``) or run
        on (``"none"``), as it applies; ``True`` when nothing says.  Settable; ``None``
        inherits again.  A table cell always wraps.

        For example::

            frame.wrap = False                   # one line, however long
        """
        if self._is_cell():
            return True
        return self._frame_attribute("wrap") != "none"

    @wrap.setter
    def wrap(self, value: bool | None) -> None:
        if self._is_cell():
            raise ValueError(f"{self.address}: a table cell always wraps")
        self._set_frame_attributes({"wrap": None if value is None
                                    else ("square" if value else "none")})

    def format(self, **formatting: Any) -> "TextFrame":
        """Apply run formatting to every run (``bold=True``, ``color="accent2"``...).

        The keys: ``bold``, ``italic``, ``strike`` (``True``, ``False``, ``None`` to
        inherit); ``underline`` (``True``, a style such as ``"dbl"``, ``None``); ``size``
        (points, or :class:`~pptx_agent.Pt`); ``typeface`` or ``font`` (a face, or
        ``"+mn-lt"``); ``color`` (a colour string or :class:`~pptx_agent.Color`);
        ``hyperlink`` (an address, a :class:`Hyperlink`, ``None``).  Any other raises
        ``TypeError`` naming these (:data:`FORMAT_KEYS`).

        For example::

            frame.format(size=14, color="tx1")
        """
        check_format_keys(formatting)
        with self._resolve()._batch():
            for paragraph in self.paragraphs:
                for run in paragraph.runs:
                    run._apply(formatting)
        return self

    # -- internals -------------------------------------------------------------------------

    def _host(self) -> TextHost:
        return self._resolve()

    def _is_cell(self) -> bool:
        return hasattr(self._host(), "_tc")

    def _inset_names(self) -> tuple[str, str, str, str]:
        return ("marL", "marT", "marR", "marB") if self._is_cell() else \
            ("lIns", "tIns", "rIns", "bIns")

    def _frame_attribute(self, name: str, *, own: bool = False) -> str | None:
        """An ``a:bodyPr`` attribute as it applies -- the frame's own, then its layout's
        and master's placeholders' -- or a table cell's ``a:tcPr`` one."""
        host = self._host()
        if self._is_cell():
            properties = host._tc().find(qn("a:tcPr"))
            return None if properties is None else properties.get(name)
        from . import effective

        chain = effective.body_chain(host)
        for properties in chain[:1] if own else chain:
            if properties is not None and properties.get(name) is not None:
                return properties.get(name)
        return None

    def _set_frame_attributes(self, values: dict) -> None:
        host = self._host()
        with _editing(host):
            if self._is_cell():
                properties = subelement(host._tc(), "a:tcPr")
            else:
                body = host._text_body(True)
                assert body is not None
                properties = subelement(body, "a:bodyPr")
            for name, value in values.items():
                set_attr(properties, name, value)

    def _paragraph_elements(self, host: TextHost | None = None) -> list[Element]:
        body = (host or self._resolve())._text_body(False)
        return [] if body is None else list(body.findall(qn("a:p")))

    def __repr__(self) -> str:
        return f"<TextFrame {self.address} {self.text[:40]!r}>"


# ------------------------------------------------------------------------------------------
# Paragraph
# ------------------------------------------------------------------------------------------


class Paragraph:
    """One ``a:p``, addressed as ``<frame>/p<index>``.

    For example::

        paragraph = deck.resolve("256.5/p1")
        paragraph.run(0).bold = True; paragraph.alignment = "center"
    """

    def __init__(self, frame: TextFrame, index: int) -> None:
        self._frame = frame
        self.index = index

    @property
    def address(self) -> str:
        """``<frame>/p<index>``, as :meth:`Document.resolve` takes it.

        For example::

            paragraph.address                            # '256.5/p1'
        """
        return f"{self._frame.address}/p{self.index}"

    # -- text and runs ---------------------------------------------------------------------

    @property
    def text(self) -> str:
        """The paragraph's text; a line break reads as a vertical tab. Settable.

        For example::

            paragraph.text = "発表日：2026年2月6日"
        """
        return _paragraph_text(self._element())

    @text.setter
    def text(self, value: str) -> None:
        if "\n" in value:
            raise ValueError("a paragraph cannot contain '\\n'; use '\\v' for a line break")
        _warn_if_escaped(value, self.text, self.address)
        with self._change() as element:
            _rebuild_paragraph(element, value)

    @property
    def runs(self) -> list["Run"]:
        """Every run (fields included, line breaks not), in order.

        For example::

            [run.text for run in paragraph.runs]
        """
        return [Run(self, index) for index in range(len(_runs(self._element())))]

    def run(self, index: int) -> "Run":
        """Run ``index`` (negative counts from the end).

        For example::

            paragraph.run(0).italic = True
        """
        count = len(_runs(self._element()))
        if not -count <= index < count:
            raise IndexError(f"{self.address} has {count} runs; no r{index}")
        return Run(self, index % count)

    def add_run(self, text: str, *, index: int | None = None, like: int | None = None,
                **formatting: Any) -> "Run":
        """Insert a run at ``index`` among the runs (default: at the end).

        Its formatting starts as a copy of run ``like`` -- by default the run before the
        insertion point, else the one after, else the paragraph's end-of-paragraph
        properties -- and ``formatting`` overrides from there.

        For example::

            paragraph.add_run("（予定）", size=10, italic=True)
        """
        if "\n" in text or _LINE_BREAK in text:
            raise ValueError("a run cannot contain line breaks; add a paragraph instead")
        check_format_keys(formatting)
        with self._change() as element:
            runs = _runs(element)
            position = len(runs) if index is None else _clamp_index(index, len(runs) + 1, self.address)
            if like is None:
                like = position - 1 if position > 0 else 0
            template = runs[like] if runs and 0 <= like < len(runs) else None
            properties = None
            if template is not None:
                properties = template.find(qn("a:rPr"))
            else:
                end = element.find(qn("a:endParaRPr"))
                if end is not None:
                    properties = end
            new = _make_run(text, properties)
            if position < len(runs):
                runs[position].addprevious(new)
            elif runs:
                runs[-1].addnext(new)
            else:
                end = element.find(qn("a:endParaRPr"))
                if end is not None:
                    end.addprevious(new)
                else:
                    element.append(new)
            run = Run(self, position)
            run._apply(formatting)
        return run

    def segment(self, pieces: "list[str]") -> "Paragraph":
        r"""Re-cut the paragraph's text into runs at new boundaries, keeping the text.

        ``pieces`` must join to :attr:`text`; each is one run's text, or ``"\v"`` for a line
        break.  A new run takes the formatting of the character it starts with, so splitting a
        run gives two runs formatted like it, ready to be formatted apart; merging runs keeps
        the first one's formatting.  A field whose text is one piece, where it was, stays a
        field.  Inline objects that are not text (equations) keep their place.

        For example::

            paragraph.segment(["発表日：", "2026年2月6日"])   # two runs, same text
        """
        pieces = [str(piece) for piece in pieces]
        for piece in pieces:
            if not piece or "\n" in piece or (_LINE_BREAK in piece and piece != _LINE_BREAK):
                raise ValueError(f"{piece!r} is not a run's text or a single line break")
        if "".join(pieces) != self.text:
            raise ValueError("the pieces must join to the paragraph's text; use set_text "
                             "to change the text")
        if self._pieces() == pieces:
            return self
        with self._change() as element:
            _resegment(element, pieces)
        return self

    def _pieces(self) -> list[str]:
        items, _ = _inline_items(self._element())
        return [item.text for item in items]

    def delete_run(self, index: int) -> "Paragraph":
        """Remove run ``index``; returns the paragraph.

        For example::

            paragraph.delete_run(-1)
        """
        with self._change() as element:
            runs = _runs(element)
            remove(runs[_clamp_index(index, len(runs), self.address)])
        return self

    def delete(self) -> None:
        """Remove the paragraph from its frame.

        For example::

            paragraph.delete()
        """
        self._frame.delete_paragraph(self.index)

    def format(self, **formatting: Any) -> "Paragraph":
        """Apply run formatting to every run in the paragraph.

        The keys: ``bold``, ``italic``, ``strike`` (``True``, ``False``, ``None`` to
        inherit); ``underline`` (``True``, a style such as ``"dbl"``, ``None``); ``size``
        (points, or :class:`~pptx_agent.Pt`); ``typeface`` or ``font`` (a face, or
        ``"+mn-lt"``); ``color`` (a colour string or :class:`~pptx_agent.Color`);
        ``hyperlink`` (an address, a :class:`Hyperlink`, ``None``).  Any other raises
        ``TypeError`` naming these (:data:`FORMAT_KEYS`).

        For example::

            paragraph.format(bold=True, color="accent1")
        """
        check_format_keys(formatting)
        with self._frame._host()._batch():
            for run in self.runs:
                run._apply(formatting)
        return self

    # -- paragraph properties --------------------------------------------------------------

    @property
    def alignment(self) -> str | None:
        """``"left"``, ``"center"``, ``"right"``, ``"justify"``, ``"distributed"``... or
        ``None`` when inherited from the layout/master list styles.

        For example::

            paragraph.alignment = "center"
        """
        raw = self._ppr_get("algn")
        return None if raw is None else _ALIGNMENT_NAMES.get(raw, raw)

    @alignment.setter
    def alignment(self, value: str | None) -> None:
        if value is not None:
            value = ALIGNMENTS.get(value, value)
            if value not in _ALIGNMENT_NAMES:
                raise ValueError(f"alignment must be one of {sorted(ALIGNMENTS)}")
        self._ppr_set("algn", value)

    @property
    def level(self) -> int:
        """Outline level, 0-8 (``a:pPr@lvl``; 0 when absent).

        For example::

            paragraph.level = 1                          # indent one level
        """
        properties = self._ppr()
        return get_int(properties, "lvl", 0) or 0

    @level.setter
    def level(self, value: int) -> None:
        level = int(value)
        if not 0 <= level <= 8:
            raise ValueError("level must be between 0 and 8")
        if level == 0 and self._ppr_get("lvl") is None:
            return  # already the default; do not create a:pPr just to say so
        self._ppr_set("lvl", str(level))

    @property
    def margin_left(self) -> int | None:
        """Left margin in EMU (``a:pPr@marL``), ``None`` when inherited.

        For example::

            paragraph.margin_left = 342900
        """
        return get_int(self._ppr(), "marL")

    @margin_left.setter
    def margin_left(self, emu: int | None) -> None:
        self._ppr_set("marL", None if emu is None else str(int(emu)))

    @property
    def indent(self) -> int | None:
        """First-line indent in EMU, relative to ``margin_left`` (negative for a hanging one).

        For example::

            paragraph.indent = -342900                    # a hanging indent
        """
        return get_int(self._ppr(), "indent")

    @indent.setter
    def indent(self, emu: int | None) -> None:
        self._ppr_set("indent", None if emu is None else str(int(emu)))

    @property
    def space_before(self) -> int | None:
        """Space above the paragraph, EMU, like every other length -- when the paragraph
        gives it as a length (``a:spcBef/a:spcPts``); ``None`` when it inherits or gives
        a share of a line.  Write points with :class:`~pptx_agent.Pt`: ``Pt(6)`` is
        76,200.  A value over 1,000 pt, or under a hundredth of one, is refused -- it is
        points written as EMU -- and one under a point that is not a ``Pt`` draws a
        :class:`~pptx_agent.UnitWarning`.  ``None`` inherits again.

        For example::

            paragraph.space_before = Pt(6)           # or 76200
        """
        return self._spacing_emu("a:spcBef")

    @space_before.setter
    def space_before(self, emu: int | None) -> None:
        self._set_spacing("a:spcBef", None if emu is None else _spacing_value(emu, "space_before"),
                          percent=False)

    @property
    def space_after(self) -> int | None:
        """Space below the paragraph, EMU, as :attr:`space_before`.

        For example::

            paragraph.space_after = Pt(6)            # or 76200
        """
        return self._spacing_emu("a:spcAft")

    @space_after.setter
    def space_after(self, emu: int | None) -> None:
        self._set_spacing("a:spcAft", None if emu is None else _spacing_value(emu, "space_after"),
                          percent=False)

    @property
    def line_spacing(self) -> float | None:
        """Line spacing as a multiple of single spacing (``1.5``), when given as a percentage.

        For example::

            paragraph.line_spacing = 1.2
        """
        properties = self._ppr()
        node = None if properties is None else properties.find(qn("a:lnSpc"))
        pct = None if node is None else node.find(qn("a:spcPct"))
        value = get_int(pct, "val")
        return None if value is None else value / PERCENT_UNIT

    @line_spacing.setter
    def line_spacing(self, multiple: float | None) -> None:
        self._set_spacing("a:lnSpc", multiple, percent=True)

    @property
    def line_spacing_points(self) -> float | None:
        """Exact line spacing in points, when given that way.

        For example::

            paragraph.line_spacing_points = 18
        """
        return self._spacing_points("a:lnSpc")

    @line_spacing_points.setter
    def line_spacing_points(self, points: float | None) -> None:
        from .units import points_of

        self._set_spacing("a:lnSpc", None if points is None else points_of(points),
                          percent=False)

    @property
    def effective(self) -> "EffectiveParagraph":
        """The paragraph's formatting as drawn, inheritance resolved and autofit's stored
        scale applied: ``size``, ``font``, ``alignment``, ``bullet``, margins, spacing
        (:class:`~pptx_agent.EffectiveParagraph`).

        For example::

            paragraph.effective.size                     # 24.0, though run.size is None
        """
        from . import effective

        return effective.paragraph(self._frame._host(), self._element())

    # -- bullets ---------------------------------------------------------------------------

    @property
    def bullet(self) -> Bullet | None:
        """The explicit bullet, or ``None`` when the list style decides.

        For example::

            paragraph.bullet = "•"                        # None inherits; "none" suppresses
        """
        properties = self._ppr()
        if properties is None:
            return None
        node = next((c for c in properties if c.tag in {qn(t) for t in _BULLET_TAGS}), None)
        if node is None:
            return None
        font_node = properties.find(qn("a:buFont"))
        color = Color.from_element(properties.find(qn("a:buClr")))
        size_node = properties.find(qn("a:buSzPct"))
        size = None if size_node is None else (get_int(size_node, "val", 0) or 0) / PERCENT_UNIT
        font = None if font_node is None else font_node.get("typeface")
        name = local_name(node)
        if name == "buNone":
            return Bullet("none")
        if name == "buChar":
            return Bullet("char", char=node.get("char"), font=font, color=color, size=size)
        if name == "buAutoNum":
            return Bullet("number", scheme=node.get("type"), start_at=get_int(node, "startAt"),
                          font=font, color=color, size=size)
        return Bullet("picture", font=font, color=color, size=size)

    @bullet.setter
    def bullet(self, value: "str | Bullet | None") -> None:
        """``None`` inherits, ``"none"`` suppresses, any other string is a bullet character."""
        if value is None:
            self.clear_bullet()
        elif isinstance(value, Bullet):
            if value.kind == "none":
                self.set_no_bullet()
            elif value.kind == "char":
                self.set_bullet(value.char or "•", font=value.font, color=value.color,
                                size=value.size)
            elif value.kind == "number":
                self.set_numbering(value.scheme or "arabicPeriod", start_at=value.start_at,
                                   font=value.font, color=value.color, size=value.size)
            else:
                raise ValueError("picture bullets cannot be set through this API")
        elif value == "none":
            self.set_no_bullet()
        else:
            self.set_bullet(value)

    def set_bullet(self, char: str = "•", *, font: str | None = None,
                   color: "str | Color | None" = None, size: float | None = None,
                   hanging: "bool | int" = True) -> "Paragraph":
        """A character bullet.  ``size`` is a fraction of the text size (``0.8``).

        ``hanging=True`` (the default) also hangs the bullet the way PowerPoint does when
        it bullets a paragraph: the text indented by an amount that steps with the font
        size (:func:`bullet_hanging`: 0.3125 in at 18 pt) and the bullet set back by the
        same (``margin_left`` = h, ``indent`` = -h), so it no longer sits on the text --
        unless the paragraph already hangs (its own indent, or its list style's, is
        negative).  An ``int`` is the hanging indent, EMU; ``False`` leaves the indents
        alone.

        For example::

            paragraph.set_bullet("–", color="accent1", size=0.8)
        """
        if not char:
            raise ValueError("a bullet character cannot be empty")
        if isinstance(hanging, bool):
            amount = None
            if hanging and self.effective.indent >= 0:
                amount = bullet_hanging(self.effective.size)
        else:
            amount = int(hanging)
            if amount < 0:
                raise ValueError("a hanging indent is a positive length, EMU")
        with self._change() as element:
            properties = _ppr(element, create=True)
            self._write_bullet_decoration(properties, font, color, size)
            replace_choice(properties, _BULLET_TAGS, make("a:buChar", char=char))
            if amount is not None:
                properties.set("marL", str(amount))
                properties.set("indent", str(-amount))
        return self

    def set_numbering(self, scheme: str = "arabicPeriod", *, start_at: int | None = None,
                      font: str | None = None, color: "str | Color | None" = None,
                      size: float | None = None) -> "Paragraph":
        """An automatically numbered bullet, e.g. ``"arabicPeriod"`` for 1. 2. 3.

        For example::

            paragraph.set_numbering("arabicPeriod", start_at=3)
        """
        if scheme not in NUMBERING_SCHEMES:
            raise ValueError(f"unknown numbering scheme {scheme!r}")
        with self._change() as element:
            properties = _ppr(element, create=True)
            self._write_bullet_decoration(properties, font, color, size)
            node = make("a:buAutoNum", type=scheme)
            if start_at is not None:
                node.set("startAt", str(int(start_at)))
            replace_choice(properties, _BULLET_TAGS, node)
        return self

    def set_no_bullet(self) -> "Paragraph":
        """Suppress the bullet (``a:buNone``), whatever the list style says.

        For example::

            paragraph.set_no_bullet()
        """
        with self._change() as element:
            replace_choice(_ppr(element, create=True), _BULLET_TAGS, make("a:buNone"))
        return self

    def clear_bullet(self) -> "Paragraph":
        """Remove explicit bullet settings, so the list style applies again.

        For example::

            paragraph.clear_bullet()
        """
        with self._change() as element:
            properties = _ppr(element, create=False)
            if properties is not None:
                replace_choice(properties, _BULLET_TAGS + (
                    "a:buClrTx", "a:buClr", "a:buSzTx", "a:buSzPct", "a:buSzPts",
                    "a:buFontTx", "a:buFont"), None)
        return self

    @staticmethod
    def _write_bullet_decoration(properties: Element, font, color, size) -> None:
        if font is not None:
            replace_choice(properties, ("a:buFontTx", "a:buFont"), make("a:buFont", typeface=font))
        if color is not None:
            node = make("a:buClr")
            node.append(Color.parse(color).to_element())
            replace_choice(properties, ("a:buClrTx", "a:buClr"), node)
        if size is not None:
            replace_choice(properties, ("a:buSzTx", "a:buSzPct", "a:buSzPts"),
                           make("a:buSzPct", val=str(round(size * PERCENT_UNIT))))

    # -- internals -------------------------------------------------------------------------

    def _element(self) -> Element:
        paragraphs = self._frame._paragraph_elements()
        if not 0 <= self.index < len(paragraphs):
            raise IndexError(f"{self.address} no longer exists ({len(paragraphs)} paragraphs)")
        return paragraphs[self.index]

    def _change(self):
        return _ElementChange(self._frame._host(), self._element)

    def _ppr(self) -> Element | None:
        return _ppr(self._element(), create=False)

    def _ppr_get(self, attribute: str) -> str | None:
        properties = self._ppr()
        return None if properties is None else properties.get(attribute)

    def _ppr_set(self, attribute: str, value: str | None) -> None:
        with self._change() as element:
            properties = _ppr(element, create=value is not None)
            if properties is not None:
                set_attr(properties, attribute, value)

    def _spacing_emu(self, tag: str) -> int | None:
        properties = self._ppr()
        node = None if properties is None else properties.find(qn(tag))
        pts = None if node is None else node.find(qn("a:spcPts"))
        value = get_int(pts, "val")
        return None if value is None else value * EMU_PER_SPACING_UNIT

    def _spacing_points(self, tag: str) -> float | None:
        properties = self._ppr()
        node = None if properties is None else properties.find(qn(tag))
        pts = None if node is None else node.find(qn("a:spcPts"))
        value = get_int(pts, "val")
        return None if value is None else value / SPACING_UNIT

    def _set_spacing(self, tag: str, value: float | None, *, percent: bool) -> None:
        with self._change() as element:
            properties = _ppr(element, create=value is not None)
            if properties is None:
                return
            if value is None:
                replace_choice(properties, (tag,), None)
                return
            node = make(tag)
            if percent:
                node.append(make("a:spcPct", val=str(round(value * PERCENT_UNIT))))
            else:
                node.append(make("a:spcPts", val=str(round(value * SPACING_UNIT))))
            replace_choice(properties, (tag,), node)

    def __repr__(self) -> str:
        return f"<Paragraph {self.address} {self.text[:40]!r}>"


# ------------------------------------------------------------------------------------------
# Run
# ------------------------------------------------------------------------------------------


class Run:
    """One ``a:r`` (or ``a:fld``), addressed as ``<paragraph>/r<index>``.

    Formatting properties read ``None`` when the run inherits that property -- from the
    paragraph, the shape's list style, the layout or the master.  Setting ``None`` removes the
    explicit value so it inherits again.

    For example::

        run = deck.resolve("256.5/p1/r0")
        run.format(bold=True, color="accent2", size=14)
    """

    def __init__(self, paragraph: Paragraph, index: int) -> None:
        self._paragraph = paragraph
        self.index = index

    @property
    def address(self) -> str:
        """``<paragraph>/r<index>``, as :meth:`Document.resolve` takes it.

        For example::

            run.address                                  # '256.5/p1/r0'
        """
        return f"{self._paragraph.address}/r{self.index}"

    @property
    def is_field(self) -> bool:
        """``True`` for a field (``a:fld``), such as the slide number.

        For example::

            run.is_field
        """
        return local_name(self._element()) == "fld"

    # -- text ------------------------------------------------------------------------------

    @property
    def text(self) -> str:
        """The run's text (no line breaks: those are between runs). Settable.

        For example::

            run.text = "2026年2月6日"
        """
        node = self._element().find(qn("a:t"))
        return "" if node is None or node.text is None else node.text

    @text.setter
    def text(self, value: str) -> None:
        if "\n" in value or _LINE_BREAK in value:
            raise ValueError("a run cannot contain line breaks; use the paragraph API")
        _warn_if_escaped(value, self.text, self.address)
        with self._change() as element:
            subelement(element, "a:t").text = value

    def delete(self) -> None:
        """Remove the run.

        For example::

            run.delete()
        """
        self._paragraph.delete_run(self.index)

    # -- character properties --------------------------------------------------------------

    @property
    def bold(self) -> bool | None:
        """Bold: ``True``, ``False``, or ``None`` when inherited. Settable.

        For example::

            run.bold = True
        """
        return self._flag("b")

    @bold.setter
    def bold(self, value: bool | None) -> None:
        self._set_flag("b", value)

    @property
    def italic(self) -> bool | None:
        """Italic: ``True``, ``False``, or ``None`` when inherited. Settable.

        For example::

            run.italic = True
        """
        return self._flag("i")

    @italic.setter
    def italic(self, value: bool | None) -> None:
        self._set_flag("i", value)

    @property
    def strike(self) -> bool | None:
        """Struck through: ``True``, ``False``, or ``None`` when inherited. Settable.

        For example::

            run.strike = True
        """
        raw = self._rpr_get("strike")
        return None if raw is None else raw != "noStrike"

    @strike.setter
    def strike(self, value: bool | None) -> None:
        self._rpr_set("strike", None if value is None else ("sngStrike" if value else "noStrike"))

    @property
    def underline(self) -> bool | str | None:
        """``True`` for a single underline, ``False`` for none, another
        ``ST_TextUnderlineType`` string (``"dbl"``, ``"wavy"``...) for the rest.

        For example::

            run.underline = True                           # or "dbl", "wavy"...
        """
        raw = self._rpr_get("u")
        if raw is None:
            return None
        if raw == "sng":
            return True
        if raw == "none":
            return False
        return raw

    @underline.setter
    def underline(self, value: bool | str | None) -> None:
        if value is True:
            value = "sng"
        elif value is False:
            value = "none"
        if value is not None and value not in UNDERLINES:
            raise ValueError(f"underline must be a bool or one of {sorted(UNDERLINES)}")
        self._rpr_set("u", value)

    @property
    def size(self) -> float | None:
        """Font size in points.

        For example::

            run.size = 14
        """
        raw = self._rpr_get("sz")
        return None if raw is None else int(raw) / SIZE_UNIT

    @size.setter
    def size(self, points: float | None) -> None:
        from .units import points_of

        points = None if points is None else points_of(points)
        if points is not None and not 1 <= float(points) <= 4000:
            raise ValueError("size must be between 1 and 4000 points")
        self._rpr_set("sz", None if points is None else str(round(float(points) * SIZE_UNIT)))

    @property
    def typeface(self) -> str | None:
        """The Latin typeface (``a:latin@typeface``), e.g. ``"Arial"`` or ``"+mn-lt"``.

        For example::

            run.typeface = "Meiryo"
        """
        properties = self._rpr()
        node = None if properties is None else properties.find(qn("a:latin"))
        return None if node is None else node.get("typeface")

    @typeface.setter
    def typeface(self, value: str | None) -> None:
        with self._change() as element:
            properties = _rpr(element, create=value is not None)
            if properties is None:
                return
            if value is None:
                replace_choice(properties, ("a:latin",), None)
            else:
                subelement(properties, "a:latin").set("typeface", value)

    font = typeface

    @property
    def effective_size(self) -> float:
        """The size in points the run is drawn at: its own, or what it inherits from its
        paragraph, list styles, layout, master and the presentation -- times the stored
        autofit scale (:attr:`TextFrame.font_scale`).

        For example::

            run.size, run.effective_size                 # (None, 24.0)
        """
        from . import effective

        return effective.size(self._paragraph._frame._host(), self._paragraph._element(),
                              self._element())

    @property
    def effective_font(self) -> str | None:
        """The Latin typeface the run is drawn in, inherited and with the theme's
        ``+mn-lt``/``+mj-lt`` resolved: the theme's body font when nothing names one, and
        Arial, PowerPoint's own default, in a shape that is not a placeholder.

        For example::

            run.typeface, run.effective_font             # ('+mj-lt', 'Calibri Light')
        """
        from . import effective

        return effective.typeface(self._paragraph._frame._host(), self._paragraph._element(),
                                  self._element())

    @property
    def color(self) -> Color | None:
        """The run's solid text colour, or ``None`` when inherited (or not a solid fill).

        For example::

            run.color = "accent2 lumMod=75%"
        """
        fill = read_fill(self._rpr())
        return fill.color if fill is not None and fill.kind == "solid" else None

    @color.setter
    def color(self, value: "str | Color | None") -> None:
        with self._change() as element:
            properties = _rpr(element, create=value is not None)
            if properties is not None:
                replace_choice(properties, FILL_TAGS, None if value is None else solid_fill(value))

    # -- hyperlink -------------------------------------------------------------------------

    @property
    def hyperlink(self) -> Hyperlink | None:
        """The run's click hyperlink, or ``None``.

        For example::

            run.hyperlink.address                        # 'https://example.org/ir'
        """
        properties = self._rpr()
        node = None if properties is None else properties.find(qn("a:hlinkClick"))
        if node is None:
            return None
        tooltip, action = node.get("tooltip") or None, node.get("action") or None
        slide = self._paragraph._frame._host()._slide
        package = slide.document.package
        relationship = package.relationships(slide.part_path).get(node.get(qn("r:id")) or "")
        if relationship is None:
            return Hyperlink(tooltip=tooltip, action=action)
        if relationship.is_external:
            return Hyperlink(address=relationship.target, tooltip=tooltip)
        if relationship.type == REL_SLIDE:
            for candidate in slide.document.slides:
                if candidate.part_path == relationship.target_part:
                    return Hyperlink(slide_id=candidate.slide_id, tooltip=tooltip)
        return Hyperlink(tooltip=tooltip, action=action)

    @hyperlink.setter
    def hyperlink(self, value: "str | Hyperlink | None") -> None:
        if value is None:
            self.remove_hyperlink()
        elif isinstance(value, Hyperlink):
            self.set_hyperlink(value.address if value.address is not None else value.slide_id,
                               tooltip=value.tooltip)
        else:
            self.set_hyperlink(value)

    def set_hyperlink(self, target: Any, *, tooltip: str | None = None) -> "Run":
        """Link the run to a web or mail address (``"https://..."``, ``"mailto:..."``), or
        to a slide of this deck (a ``Slide`` or its ``sldId``).

        The relationship is added to the slide; one the old link used is removed once nothing
        references it any more.

        For example::

            run.set_hyperlink("https://example.org/ir", tooltip="IR")
        """
        if target is None:
            raise ValueError("use remove_hyperlink() to remove a link")
        with self._change() as element:
            slide = self._paragraph._frame._host()._slide
            package = slide.document.package
            node = make("a:hlinkClick")
            if isinstance(target, str):
                if not target:
                    raise ValueError("a hyperlink needs an address")
                rel_id = package.add_external_relationship(slide.part_path, REL_HYPERLINK, target)
                node.set(qn("r:id"), rel_id)
            else:
                slide_id = getattr(target, "slide_id", target)
                destination = slide.document.slide(int(slide_id))
                rel_id = package.add_relationship(slide.part_path, REL_SLIDE,
                                                  destination.part_path)
                node.set(qn("r:id"), rel_id)
                node.set("action", SLIDE_JUMP_ACTION)
            if tooltip:
                node.set("tooltip", tooltip)
            properties = _rpr(element, create=True)
            assert properties is not None
            replace_choice(properties, ("a:hlinkClick",), node)
        return self

    def remove_hyperlink(self) -> "Run":
        """Unlink the run; its relationship goes too once nothing else uses it.

        For example::

            run.remove_hyperlink()
        """
        properties = self._rpr()
        if properties is None or properties.find(qn("a:hlinkClick")) is None:
            return self
        with self._change() as element:
            properties = _rpr(element, create=False)
            if properties is not None:
                replace_choice(properties, ("a:hlinkClick",), None)
        return self

    def format(self, **formatting: Any) -> "Run":
        """Set several properties as one undo step: ``run.format(bold=True, size=24)``.

        The keys: ``bold``, ``italic``, ``strike`` (``True``, ``False``, ``None`` to
        inherit); ``underline`` (``True``, a style such as ``"dbl"``, ``None``); ``size``
        (points, or :class:`~pptx_agent.Pt`); ``typeface`` or ``font`` (a face, or
        ``"+mn-lt"``); ``color`` (a colour string or :class:`~pptx_agent.Color`);
        ``hyperlink`` (an address, a :class:`Hyperlink`, ``None``).  Any other raises
        ``TypeError`` naming these (:data:`FORMAT_KEYS`).

        For example::

            run.format(bold=True, size=14, color="accent1")
        """
        with self._paragraph._frame._host()._batch():
            self._apply(formatting)
        return self

    # -- internals -------------------------------------------------------------------------

    _FORMAT_KEYS = FORMAT_KEYS

    def _apply(self, formatting: dict) -> None:
        check_format_keys(formatting)
        for key, value in formatting.items():
            setattr(self, key, value)

    def _element(self) -> Element:
        runs = _runs(self._paragraph._element())
        if not 0 <= self.index < len(runs):
            raise IndexError(f"{self.address} no longer exists ({len(runs)} runs)")
        return runs[self.index]

    def _change(self):
        return _ElementChange(self._paragraph._frame._host(), self._element)

    def _rpr(self) -> Element | None:
        return _rpr(self._element(), create=False)

    def _rpr_get(self, attribute: str) -> str | None:
        properties = self._rpr()
        return None if properties is None else properties.get(attribute)

    def _rpr_set(self, attribute: str, value: str | None) -> None:
        with self._change() as element:
            properties = _rpr(element, create=value is not None)
            if properties is not None:
                set_attr(properties, attribute, value)

    def _flag(self, attribute: str) -> bool | None:
        raw = self._rpr_get(attribute)
        return None if raw is None else raw.strip().lower() in {"1", "true", "on"}

    def _set_flag(self, attribute: str, value: bool | None) -> None:
        self._rpr_set(attribute, None if value is None else ("1" if value else "0"))

    def __repr__(self) -> str:
        return f"<Run {self.address} {self.text[:40]!r}>"


# ------------------------------------------------------------------------------------------
# Change bracketing
# ------------------------------------------------------------------------------------------


class _ElementChange:
    """One undo step around an edit of a freshly resolved element."""

    def __init__(self, host: TextHost, locate: Callable[[], Element]) -> None:
        self._host = host
        self._locate = locate
        self._batch: Any = None

    def __enter__(self) -> Element:
        element = self._locate()  # resolve first, so a stale address fails before a checkpoint
        self._batch = self._host._batch()
        self._batch.__enter__()
        self._host._before_change()
        return element

    def __exit__(self, *exc) -> Any:
        if exc[0] is None:
            self._host._after_change()
        return self._batch.__exit__(*exc)


# ------------------------------------------------------------------------------------------
# XML helpers
# ------------------------------------------------------------------------------------------


def _ppr(paragraph: Element, *, create: bool) -> Element | None:
    if create:
        return subelement(paragraph, "a:pPr")
    return paragraph.find(qn("a:pPr"))


def _rpr(run: Element, *, create: bool) -> Element | None:
    if create:
        properties = run.find(qn("a:rPr"))
        if properties is None:
            properties = subelement(run, "a:rPr")
        return properties
    return run.find(qn("a:rPr"))


def _clamp_index(index: int, count: int, address: str) -> int:
    if not -count <= index < count:
        raise IndexError(f"{address}: index {index} is out of range ({count})")
    return index % count if count else 0
