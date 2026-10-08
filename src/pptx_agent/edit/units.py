"""Lengths are EMU everywhere -- 914,400 per inch, 12,700 per point -- and :class:`Pt` says
"points" where that reads better.

Font sizes are the one exception, and stay in points, as every Office application shows
them: :attr:`Run.size <pptx_agent.Run.size>` is ``18``, not 228,600.  A :class:`Pt` given
to a size is read as its points.
"""

from __future__ import annotations

import warnings

EMU_PER_POINT = 12700
EMU_PER_INCH = 914400
EMU_PER_CM = 360000


class Pt(int):
    """A length in points, as the EMU everything else takes: ``Pt(6) == 76200``.

    It *is* that ``int``, so it goes wherever a length does -- ``left``, ``width``,
    ``insets``, ``space_after`` -- and a font size takes it as points.  ``points`` is the
    value given.

    For example::

        paragraph.space_after = Pt(6)            # the same as 76200
        shape.top += Pt(12)
    """

    points: float

    def __new__(cls, points: float) -> "Pt":
        value = super().__new__(cls, round(float(points) * EMU_PER_POINT))
        value.points = float(points)
        return value

    def __repr__(self) -> str:
        return f"Pt({self.points:g})"


class Inches(int):
    """A length in inches, as EMU: ``Inches(1) == 914400``.  Like :class:`Pt`, it *is* the
    ``int``, and ``inches`` is the value given.

    For example::

        slide.add_shape("rect", Inches(1), Inches(1), Inches(2), Inches(0.75))
    """

    inches: float

    def __new__(cls, inches: float) -> "Inches":
        value = super().__new__(cls, round(float(inches) * EMU_PER_INCH))
        value.inches = float(inches)
        return value

    def __repr__(self) -> str:
        return f"Inches({self.inches:g})"


class Cm(int):
    """A length in centimetres, as EMU: ``Cm(1) == 360000``.

    For example::

        shape.width = Cm(4.5)
    """

    cm: float

    def __new__(cls, cm: float) -> "Cm":
        value = super().__new__(cls, round(float(cm) * EMU_PER_CM))
        value.cm = float(cm)
        return value

    def __repr__(self) -> str:
        return f"Cm({self.cm:g})"


def to_pt(emu: float, digits: int = 2) -> float:
    """A length in EMU as points, rounded to ``digits`` decimals: ``to_pt(914400) == 72.0``.

    For example::

        print(to_pt(shape.width), "pt wide")
    """
    return round(float(emu) / EMU_PER_POINT, digits)


class UnitWarning(UserWarning):
    """A length that looks like it was meant in other units: ``space_after = 600`` is 600
    EMU, under a twentieth of a point -- was ``Pt(600)``, or 6 points, meant?  Lengths are
    EMU; say points with :class:`Pt`.

    For example::

        warnings.simplefilter("error", pptx_agent.UnitWarning)   # make it fail loudly
    """


#: A paragraph spacing above this (EMU, 1,000 pt) is not a spacing: a value in the wrong
#: units, or a mistake.
MAX_SPACING = 1000 * EMU_PER_POINT
#: ``a:spcPts`` counts hundredths of a point: 127 EMU.  Less than that writes nothing.
SPACING_STEP = 127


def points_of(value) -> float:
    """A font size in points: a :class:`Pt` by its points, a number as it is."""
    return value.points if isinstance(value, Pt) else float(value)


def spacing(value, what: str, *, stacklevel: int = 3) -> int:
    """A paragraph spacing, EMU, checked: negative, implausibly large (over 1,000 pt) and
    sub-hundredth-of-a-point values -- what points given as EMU look like -- are refused,
    and a value under a point that is not a :class:`Pt` is warned about."""
    if isinstance(value, bool):
        raise TypeError(f"{what} is a length in EMU, not {value!r}")
    emu = int(value) if isinstance(value, Pt) else round(float(value))
    if emu < 0:
        raise ValueError(f"{what} cannot be negative")
    if emu > MAX_SPACING:
        raise ValueError(f"{what}={emu:,} EMU is {emu / EMU_PER_POINT:,.0f} pt; lengths are "
                         f"EMU (12,700 per point) -- for points write Pt(...)")
    if 0 < emu < SPACING_STEP and not isinstance(value, Pt):
        raise ValueError(f"{what}={value!r} EMU is under a hundredth of a point and would "
                         f"write nothing; lengths are EMU -- for {value!r} points write "
                         f"Pt({value!r})")
    if 0 < emu < EMU_PER_POINT and not isinstance(value, Pt):
        warnings.warn(f"{what}={value!r} EMU is {emu / EMU_PER_POINT:.2f} pt; lengths are EMU "
                      f"-- for points write Pt(...)", UnitWarning, stacklevel=stacklevel)
    return emu


__all__ = ["Cm", "EMU_PER_CM", "EMU_PER_INCH", "EMU_PER_POINT", "Inches", "MAX_SPACING", "Pt",
           "UnitWarning", "points_of", "spacing", "to_pt"]
