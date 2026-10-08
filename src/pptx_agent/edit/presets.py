"""Preset geometry facts an editor needs: adjustment defaults and connection sites.

Both come from ECMA-376's preset definitions, compiled into :mod:`.preset_sites` by
``tools/derive_connection_sites.py``.  A preset's sites are positions written in DrawingML's
guide language (``"il"`` is ``hc - wd2 cos 45deg`` for an ellipse), so placing a connector's
end means evaluating those guides for the shape's size and adjustments -- which is all this
module does.  It draws nothing; pptx2svg does the drawing.

Values are EMU and DrawingML's raw units: adjustments in 1/100,000 (or 1/60,000 of a degree
for an angle), angles in 1/60,000 of a degree.  Guides are evaluated in floating point and
rounded only where a caller needs an integer, which matches the positions PowerPoint wrote
in the probes (an ellipse's ``il``/``it`` site to the EMU).
"""

from __future__ import annotations

import math
from typing import Mapping

from .preset_sites import PRESET_SITES

#: One degree in DrawingML's angle unit.
DEGREE = 60000.0

_ANGLES = {
    "cd8": 2700000.0, "cd4": 5400000.0, "cd2": 10800000.0, "3cd8": 8100000.0,
    "3cd4": 16200000.0, "5cd8": 13500000.0, "7cd8": 18900000.0,
}
_DIVISORS = (2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 24, 32)


def adjustment_defaults(preset: str) -> tuple[tuple[str, int], ...]:
    """``((name, default), ...)`` in the preset's own order; ``()`` for one without any."""
    entry = PRESET_SITES.get(preset)
    return () if entry is None else entry[0]


def has_sites(preset: str | None) -> bool:
    return preset is not None and bool(PRESET_SITES.get(preset, ((), (), ()))[2])


def connection_sites(preset: str, width: float, height: float,
                     adjustments: Mapping[str, float] | None = None
                     ) -> list[tuple[float, float, float]]:
    """``(angle, x, y)`` for each site of ``preset`` in a ``width`` x ``height`` box.

    ``x``/``y`` are relative to the box's top-left corner, before any flip or rotation of
    the shape; ``angle`` (1/60,000 degree, clockwise from +x) is the way a connector leaves
    the site.  The list is in ``idx`` order, the order ``stCxn``/``endCxn`` count in.
    """
    entry = PRESET_SITES.get(preset)
    if entry is None or not entry[2]:
        return []
    defaults, guides, sites = entry
    variables = _builtins(float(width), float(height))
    given = adjustments or {}
    for name, default in defaults:
        variables[name] = float(given.get(name, default))
    for name, formula in guides:
        variables[name] = evaluate(formula, variables)
    return [(_value(angle, variables), _value(x, variables), _value(y, variables))
            for angle, x, y in sites]


def _builtins(width: float, height: float) -> dict[str, float]:
    shortest = min(width, height)
    variables = {
        "w": width, "h": height, "l": 0.0, "t": 0.0, "r": width, "b": height,
        "hc": width / 2, "vc": height / 2, "ss": shortest, "ls": max(width, height),
    }
    for divisor in _DIVISORS:
        variables[f"wd{divisor}"] = width / divisor
        variables[f"hd{divisor}"] = height / divisor
        variables[f"ssd{divisor}"] = shortest / divisor
    variables.update(_ANGLES)
    return variables


def _value(token: str, variables: Mapping[str, float]) -> float:
    try:
        return float(token)
    except ValueError:
        return variables.get(token, 0.0)


def evaluate(formula: str, variables: Mapping[str, float]) -> float:
    """One guide formula (ECMA-376 Part 1, 20.1.9.11): an operator and up to three operands."""
    tokens = formula.split()
    operator = tokens[0] if tokens else "val"

    def arg(index: int) -> float:
        return _value(tokens[index], variables) if index < len(tokens) else 0.0

    def radians(angle: float) -> float:
        return math.radians(angle / DEGREE)

    if operator == "val":
        return arg(1)
    if operator == "*/":
        return arg(1) * arg(2) / (arg(3) or 1.0)
    if operator == "+-":
        return arg(1) + arg(2) - arg(3)
    if operator == "+/":
        return (arg(1) + arg(2)) / (arg(3) or 1.0)
    if operator == "?:":
        return arg(2) if arg(1) > 0 else arg(3)
    if operator == "abs":
        return abs(arg(1))
    if operator == "at2":
        return math.degrees(math.atan2(arg(2), arg(1))) * DEGREE
    if operator == "cat2":
        return arg(1) * math.cos(math.atan2(arg(3), arg(2)))
    if operator == "sat2":
        return arg(1) * math.sin(math.atan2(arg(3), arg(2)))
    if operator == "cos":
        return arg(1) * math.cos(radians(arg(2)))
    if operator == "sin":
        return arg(1) * math.sin(radians(arg(2)))
    if operator == "tan":
        return arg(1) * math.tan(radians(arg(2)))
    if operator == "max":
        return max(arg(1), arg(2))
    if operator == "min":
        return min(arg(1), arg(2))
    if operator == "mod":
        return math.sqrt(arg(1) ** 2 + arg(2) ** 2 + arg(3) ** 2)
    if operator == "pin":
        return max(arg(1), min(arg(2), arg(3)))
    if operator == "sqrt":
        return math.sqrt(max(arg(1), 0.0))
    raise ValueError(f"unknown guide operator {operator!r} in {formula!r}")


def _preset_table() -> dict[str, tuple[str, ...]]:
    from ..oxml.xml import PRESET_GEOMETRIES

    return {name: tuple(adjustment for adjustment, _ in adjustment_defaults(name))
            for name in sorted(PRESET_GEOMETRIES)}


#: Every preset geometry name ``add_shape`` and ``Shape.preset`` take (ECMA-376's
#: ``ST_ShapeType``), each with its adjustment names in the preset's own order -- they
#: differ: a rounded rectangle's is ``adj``, a chevron's ``adj``, a right arrow's ``adj1``
#: and ``adj2``.  Defaults and values are in :attr:`Shape.adjustments`.
#:
#: For example::
#:
#:     from pptx_agent import PRESETS
#:     PRESETS["roundRect"], PRESETS["rightArrow"]     # ('adj',), ('adj1', 'adj2')
PRESETS: dict[str, tuple[str, ...]] = _preset_table()

