"""Colours as the file stores them -- theme colours stay theme colours.

A theme colour is a *reference*: ``<a:schemeClr val="accent1"/>`` means "whatever accent 1 is
in this deck's theme", so changing the theme recolours the shape.  Writing the resolved hex
instead looks identical today and silently detaches the shape from the theme forever.  So a
:class:`Color` keeps the reference, never resolves it, and is written back exactly as given.

Specifications accepted wherever a colour is (``run.color = ...``, ``shape.fill = ...``)::

    "4472C4" / "#4472C4"              an explicit RGB colour   -> <a:srgbClr val="4472C4"/>
    "accent1"                         a theme colour           -> <a:schemeClr val="accent1"/>
    "accent1 lumMod=75% lumOff=25%"   a tint or shade of one   -> the same, with modifiers
    Color.theme("accent1", lum_mod=0.75, lum_off=0.25)          the same, from Python

Modifier values in the string form are either percentages (``75%``) or the raw OOXML integer
in thousandths of a percent (``75000``); the keyword form takes fractions (``0.75``).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..oxml.xml import COLOR_TAGS, Element, local_name, make, qn

#: ``ST_SchemeColorVal``.  On a slide, ``bg1``/``tx1``/``bg2``/``tx2`` are the usual spellings
#: (they go through the master's colour map); ``dk1``/``lt1``/... name theme slots directly.
SCHEME_COLORS = frozenset({
    "bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5",
    "accent6", "hlink", "folHlink", "phClr", "dk1", "lt1", "dk2", "lt2",
})

#: ``EG_ColorTransform`` members that carry a ``val``.
VALUED_TRANSFORMS = frozenset({
    "tint", "shade", "alpha", "alphaOff", "alphaMod", "hue", "hueOff", "hueMod", "sat",
    "satOff", "satMod", "lum", "lumOff", "lumMod", "red", "redOff", "redMod", "green",
    "greenOff", "greenMod", "blue", "blueOff", "blueMod",
})
#: ``EG_ColorTransform`` members with no value.
FLAG_TRANSFORMS = frozenset({"comp", "inv", "gray", "gamma", "invGamma"})

_KIND_BY_TAG = {
    "srgbClr": "rgb",
    "schemeClr": "scheme",
    "sysClr": "system",
    "prstClr": "preset",
    "hslClr": "hsl",
    "scrgbClr": "scrgb",
}
_TAG_BY_KIND = {kind: tag for tag, kind in _KIND_BY_TAG.items()}
_HEX = re.compile(r"#?([0-9A-Fa-f]{6})")


@dataclass(frozen=True, eq=False)
class Color:
    """A DrawingML colour, unresolved.

    ``kind`` is ``"rgb"`` or ``"scheme"`` for anything this API writes; colours read from a
    file may also be ``"system"``, ``"preset"``, ``"hsl"`` or ``"scrgb"``, and those are written
    back verbatim.  ``transforms`` are the raw ``(name, value)`` modifiers in document order;
    ``value`` is ``None`` for the flag modifiers such as ``gray``.

    For example::

        Color.parse("accent1 lumMod=75% lumOff=25%")   # or Color.rgb("1F4E79")
    """

    kind: str
    value: str
    transforms: tuple[tuple[str, int | None], ...] = ()
    #: Attributes beyond ``val`` on colours read from a file (``sysClr@lastClr``, the
    #: ``hslClr`` components), kept so they write back verbatim.
    extra: tuple[tuple[str, str], ...] = ()

    # -- construction ----------------------------------------------------------------------

    @classmethod
    def rgb(cls, hex_value: str) -> "Color":
        """An RGB colour from ``RRGGBB`` (with or without ``#``).

        For example::

            Color.rgb("#1F4E79")
        """
        match = _HEX.fullmatch(hex_value.strip())
        if match is None:
            raise ValueError(f"{hex_value!r} is not a six-digit hex colour")
        return cls("rgb", match.group(1).upper())

    @classmethod
    def theme(
        cls,
        name: str,
        *,
        lum_mod: float | None = None,
        lum_off: float | None = None,
        tint: float | None = None,
        shade: float | None = None,
        alpha: float | None = None,
    ) -> "Color":
        """A theme colour, optionally tinted: ``Color.theme("accent1", lum_mod=0.75)``.

        Fractions, so ``lum_mod=0.75`` is ``<a:lumMod val="75000"/>``.  PowerPoint's own
        "Lighter 40%" is ``lum_mod=0.6, lum_off=0.4``; "Darker 25%" is ``lum_mod=0.75``.

        For example::

            Color.theme("accent1", lum_mod=0.75)
        """
        if name not in SCHEME_COLORS:
            raise ValueError(f"{name!r} is not a theme colour; expected one of {sorted(SCHEME_COLORS)}")
        transforms = []
        # Office writes tint/shade before lumMod/lumOff, and alpha last; follow it.
        for key, value in (("tint", tint), ("shade", shade), ("lumMod", lum_mod),
                           ("lumOff", lum_off), ("alpha", alpha)):
            if value is not None:
                transforms.append((key, round(value * 100000)))
        return cls("scheme", name, tuple(transforms))

    @classmethod
    def parse(cls, spec: "str | Color") -> "Color":
        """Accept a :class:`Color` or a string spec (see the module docstring).

        For example::

            shape.fill = Color.parse("accent2 lumMod=50%")
        """
        if isinstance(spec, Color):
            return spec
        if not isinstance(spec, str):
            raise TypeError(f"expected a colour string or Color, got {type(spec).__name__}")
        base, *modifiers = spec.split()
        if base in SCHEME_COLORS:
            color = cls("scheme", base)
        else:
            try:
                color = cls.rgb(base)
            except ValueError:
                raise ValueError(
                    f"{base!r} is neither a hex colour nor a theme colour "
                    f"({', '.join(sorted(SCHEME_COLORS))})"
                ) from None
        transforms = []
        for modifier in modifiers:
            name, separator, raw = modifier.partition("=")
            if name in FLAG_TRANSFORMS and not separator:
                transforms.append((name, None))
                continue
            if name not in VALUED_TRANSFORMS or not separator:
                raise ValueError(f"unknown colour modifier {modifier!r}")
            transforms.append((name, _modifier_value(raw)))
        return cls(color.kind, color.value, tuple(transforms))

    # -- the XML ---------------------------------------------------------------------------

    @classmethod
    def from_element(cls, element: Element | None) -> "Color | None":
        """Read a colour element (``a:srgbClr`` etc.), or the first one under a container.

        For example::

            Color.from_element(solid_fill_element)
        """
        if element is None:
            return None
        if local_name(element) not in _KIND_BY_TAG:
            wanted = {qn(tag) for tag in COLOR_TAGS}
            element = next((child for child in element if child.tag in wanted), None)
            if element is None:
                return None
        kind = _KIND_BY_TAG[local_name(element)]
        transforms: list[tuple[str, int | None]] = []
        for child in element:
            name = local_name(child)
            if not name:
                continue
            raw = child.get("val")
            transforms.append((name, int(raw) if raw is not None and _is_int(raw) else None))
        extra = tuple((k, v) for k, v in element.attrib.items() if k != "val")
        return cls(kind, element.get("val", ""), tuple(transforms), extra)

    def to_element(self) -> Element:
        """The colour as a new DrawingML element (``a:schemeClr``, ``a:srgbClr``...).

        For example::

            color.to_element()
        """
        element = make(f"a:{_TAG_BY_KIND[self.kind]}")
        if self.kind != "hsl" or self.value:
            element.set("val", self.value)
        for key, value in self.extra:
            element.set(key, value)
        for name, value in self.transforms:
            child = make(f"a:{name}")
            if value is not None:
                child.set("val", str(value))
            element.append(child)
        return element

    # -- reading ---------------------------------------------------------------------------

    @property
    def is_theme(self) -> bool:
        """``True`` for a theme (scheme) colour.

        For example::

            Color.parse("accent1").is_theme       # True
        """
        return self.kind == "scheme"

    def resolve(self, where) -> str:
        """What this colour looks like, as ``"#RRGGBB"``: resolved against ``where`` -- a
        :class:`Document` (its first master's theme), a :class:`Slide` or a shape (the
        theme and colour map that slide uses) -- with every modifier applied as PowerPoint
        composes them.  The colour itself stays unresolved; alpha is not part of the answer.

        For example::

            Color.parse("accent1 lumMod=75%").resolve(deck)     # '#2F5597'
            shape.fill.color.resolve(shape)
        """
        from .theme import theme_for

        return theme_for(where).resolve(self)

    def transform(self, name: str) -> int | None:
        """A modifier's raw value (``lumMod`` in 1000ths of a percent), or ``None``.

        For example::

            Color.parse("accent1 lumMod=75%").transform("lumMod")   # 75000
        """
        for key, value in self.transforms:
            if key == name:
                return value
        return None

    def __str__(self) -> str:
        base = f"#{self.value}" if self.kind == "rgb" else self.value
        parts = [base] + [name if value is None else f"{name}={value}" for name, value in self.transforms]
        return " ".join(parts)

    def __repr__(self) -> str:
        return f"Color({str(self)!r})"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, str):
            try:
                other = Color.parse(other)
            except (ValueError, TypeError):
                return False
        if not isinstance(other, Color):
            return NotImplemented
        return (self.kind, self.value, self.transforms) == (other.kind, other.value, other.transforms)

    def __hash__(self) -> int:
        return hash((self.kind, self.value, self.transforms))


def _modifier_value(raw: str) -> int:
    raw = raw.strip()
    if raw.endswith("%"):
        return round(float(raw[:-1]) * 1000)
    if not _is_int(raw):
        raise ValueError(f"colour modifier value {raw!r} must be an integer or a percentage")
    return int(raw)


def _is_int(raw: str) -> bool:
    try:
        int(raw)
    except ValueError:
        return False
    return True
