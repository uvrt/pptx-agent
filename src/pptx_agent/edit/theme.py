"""The deck's theme, read: its colours as RGB and its fonts -- and a colour resolved.

Colours stay unresolved everywhere else in this library (:mod:`.color`): a theme colour is
written as a reference, so that the theme can change.  Reading what a reference looks like
*today* is still worth having -- to choose a text colour that reads on an accent fill, or
to map a hard-coded RGB back onto the theme slot it equals.  The resolution is
ooxml-common's, the one pptx2svg draws with (scheme lookup through the colour map, then
every transform as PowerPoint composes them, measured swatch by swatch).

A slide's colour map is its master's ``p:clrMap``, overridden by its layout's and its own
``p:clrMapOvr``: ``tx1`` is usually ``dk1``, ``bg1`` ``lt1``.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Mapping, NamedTuple

from ..oxml.package import REL_SLIDE_LAYOUT, REL_SLIDE_MASTER
from ..oxml.xml import Element, find, make, qn, subelement
from . import inherit

if TYPE_CHECKING:  # pragma: no cover
    from .color import Color

REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
#: The theme's twelve slots, in the order PowerPoint lists them.
SLOTS = ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
         "accent6", "hlink", "folHlink")
#: The colour map's own names for the slots a slide usually spells.
MAPPED = ("bg1", "tx1", "bg2", "tx2")


class ThemeFonts(NamedTuple):
    """The theme's fonts: ``major`` for headings (``+mj-lt``), ``minor`` for body text
    (``+mn-lt``), and their East Asian faces (``+mj-ea``, ``+mn-ea``) when the theme names
    them.

    For example::

        deck.theme.fonts.major                      # 'Calibri Light'
    """

    major: str | None
    minor: str | None
    major_east_asian: str | None = None
    minor_east_asian: str | None = None


@dataclass(frozen=True)
class Theme:
    """A slide master's theme, resolved: ``colors`` maps every scheme name -- the twelve
    slots (``dk1``...``folHlink``) and the colour map's ``tx1``, ``bg1``, ``tx2``,
    ``bg2`` -- to ``"#RRGGBB"``; ``fonts`` is a :class:`ThemeFonts`.

    For example::

        theme = deck.theme
        theme.colors["accent1"]                     # '#4472C4'
        theme.fonts.minor                           # 'Calibri'
        theme.resolve("accent1 lumMod=75%")         # '#2F5597'
    """

    name: str
    colors: dict[str, str]
    fonts: ThemeFonts
    _scheme: dict = field(default_factory=dict, repr=False, compare=False)
    _color_map: dict = field(default_factory=dict, repr=False, compare=False)
    _document: object = field(default=None, repr=False, compare=False)
    _master: str | None = field(default=None, repr=False, compare=False)

    def resolve(self, color: "str | Color") -> str:
        """The ``"#RRGGBB"`` a colour (a :class:`Color` or a colour string) is drawn with
        under this theme and its master's colour map; alpha is not part of the answer.

        For example::

            deck.theme.resolve("tx1")               # '#000000'
        """
        return resolve_with(color, self._scheme, self._color_map)

    # -- changing it -------------------------------------------------------------------------

    def set_colors(self, colors: Mapping[str, str]) -> "Theme":
        """Set some of the theme's twelve colours, ``{"accent1": "#1F4E79", ...}``: a slot
        (``dk1``...``folHlink``) or the colour map's name for one (``tx1``, ``bg2``...), to
        an RGB colour.  Everything that refers to the slot -- every ``accent1`` fill, tint
        and text in the deck -- follows, as when a theme's colours are changed in
        PowerPoint.  One undo step.  Returns the theme as it now is (this object keeps
        what it read).

        For example::

            theme = deck.theme.set_colors({"accent1": "#1F4E79", "dk2": "#0B2545"})
        """
        changes = {}
        for name, value in dict(colors).items():
            slot = self._color_map.get(name, name) if name in MAPPED else name
            if slot not in SLOTS:
                raise ValueError(f"{name!r} is not a theme colour; the slots are "
                                 f"{', '.join(SLOTS)} (or {', '.join(MAPPED)})")
            match = re.fullmatch(r"#?([0-9A-Fa-f]{6})", str(value).strip())
            if match is None:
                raise ValueError(f"{name}: a theme colour is RGB, '#RRGGBB', not {value!r}")
            changes[slot] = match.group(1).upper()
        scheme = self._element("a:themeElements/a:clrScheme")
        with self._editing():
            for slot, rgb in changes.items():
                node = scheme.find(qn(f"a:{slot}"))
                if node is None:
                    raise ValueError(f"the theme has no {slot} to set")
                for child in list(node):
                    node.remove(child)
                node.text = None
                node.append(make("a:srgbClr", val=rgb))
        return self._reread()

    def set_fonts(self, major: str | None = None, minor: str | None = None, *,
                  major_east_asian: str | None = None,
                  minor_east_asian: str | None = None) -> "Theme":
        """Set the theme's fonts: ``major`` for headings (``+mj-lt``), ``minor`` for body
        text (``+mn-lt``), and their East Asian faces; ``None`` leaves one as it is.
        Every title and body that names the theme's font follows.  One undo step; returns
        the theme as it now is.

        For example::

            deck.theme.set_fonts(major="Georgia", minor="Arial")
        """
        wanted = {("a:majorFont", "a:latin"): major, ("a:minorFont", "a:latin"): minor,
                  ("a:majorFont", "a:ea"): major_east_asian,
                  ("a:minorFont", "a:ea"): minor_east_asian}
        scheme = self._element("a:themeElements/a:fontScheme")
        with self._editing():
            for (holder, script), face in wanted.items():
                if face is None:
                    continue
                if not str(face).strip():
                    raise ValueError("a typeface is a name, like 'Arial'")
                font = subelement(subelement(scheme, holder), script)
                for attribute in ("panose", "pitchFamily", "charset"):
                    font.attrib.pop(attribute, None)          # they described the old face
                font.set("typeface", str(face).strip())
        return self._reread()

    # -- design guidance, as data ------------------------------------------------------------

    @property
    def roles(self) -> "ThemeRoles":
        """Which theme colours act as the template's primary, secondary, neutral and
        highlight colours (:class:`ThemeRoles`), read from how its master and layouts use
        them -- data to choose colours by, so a new shape looks like the template's own.

        For example::

            roles = deck.theme.roles
            box.fill = f"{roles.neutral} lumMod=20% lumOff=80%"
            bar.fill = roles.primary
        """
        return theme_roles(self)

    def tints(self, slot: str) -> tuple["Tint", ...]:
        """The recommended lighter and darker variants of one theme colour (:class:`Tint`),
        lightest first: the five a colour picker's column offers under it.  A mid-tone
        colour gets lighter 80%, 60% and 40% (``lumMod``/``lumOff``) and darker 25% and
        50% (``lumMod``); one darker than 20% luminance lighter 90%, 75%, 50%, 25% and
        10%; one lighter than 80% darker 10%, 25%, 50%, 75% and 90%.

        For example::

            [tint.color for tint in deck.theme.tints("accent1")]
            # ['accent1 lumMod=20% lumOff=80%', ..., 'accent1 lumMod=50%']
        """
        return theme_tints(self, slot)

    @property
    def ramps(self) -> dict[str, tuple["Tint", ...]]:
        """:meth:`tints` for every colour a design uses: ``dk2``, ``lt2`` and
        ``accent1``...``accent6``.

        For example::

            deck.theme.ramps["accent2"][0].hex          # the lightest accent2 tint
        """
        return {slot: self.tints(slot) for slot in RAMP_SLOTS}

    # -- internals -----------------------------------------------------------------------------

    def _element(self, path: str) -> Element:
        document = self._require_document()
        part = theme_part(document.package, self._master) if self._master else None
        root = document.package.tree(part) if part else None
        node = find(root, path) if root is not None else None
        if node is None:
            raise ValueError(f"the theme has no {path.rpartition('/')[2]} to change")
        return node

    def _editing(self):
        from contextlib import contextmanager

        document = self._require_document()
        part = theme_part(document.package, self._master)

        @contextmanager
        def editing():
            with document.batch():
                document.history.checkpoint()
                yield
                document.package.mark_dirty(part)

        return editing()

    def _reread(self) -> "Theme":
        return read_theme(self._document.package, self._master, document=self._document)

    def _require_document(self):
        if self._document is None or self._master is None:
            raise ValueError("this Theme was not read from a deck; take it from deck.theme")
        return self._document


def theme_part(package, master_part: str) -> str | None:
    related = package.related_parts_of_type(master_part, REL_THEME)
    return related[0] if related and package.has_part(related[0]) else None


def first_master(package) -> str | None:
    related = package.related_parts_of_type(package.presentation_part(), REL_SLIDE_MASTER)
    return related[0] if related else None


def color_map(package, master_part: str | None, overrides: list[Element | None] = ()) -> dict:
    """The effective colour map: the schema's default, the master's ``p:clrMap``, then each
    ``p:clrMapOvr/a:overrideClrMapping`` in ``overrides`` (layout, then slide)."""
    from ooxml_common.drawingml.color import DEFAULT_COLOR_MAP

    mapping = dict(DEFAULT_COLOR_MAP)
    root = package.tree(master_part) if master_part else None
    node = find(root, "p:clrMap") if root is not None else None
    if node is not None:
        mapping.update(dict(node.attrib))
    for override in overrides:
        if override is None:
            continue
        explicit = override.find(qn("a:overrideClrMapping"))
        if explicit is not None:
            mapping.update(dict(explicit.attrib))
    return mapping


def read_theme(package, master_part: str | None, overrides: list[Element | None] = (), *,
               document=None) -> Theme:
    from ooxml_common.drawingml.read import parse_color_scheme

    part = theme_part(package, master_part) if master_part else None
    root = package.tree(part) if part else None
    elements = find(root, "a:themeElements") if root is not None else None
    scheme_node = find(elements, "a:clrScheme") if elements is not None else None
    try:
        scheme = parse_color_scheme(scheme_node)
    except AttributeError:  # a comment among the slots: read what is there
        scheme = {}
    mapping = color_map(package, master_part, overrides)
    colors = {name: resolve_with(name, scheme, mapping) for name in SLOTS + MAPPED}
    fonts = find(elements, "a:fontScheme") if elements is not None else None

    def face(path: str) -> str | None:
        node = find(fonts, path) if fonts is not None else None
        return (node.get("typeface") or None) if node is not None else None

    return Theme(
        name=(root.get("name") if root is not None else None) or "",
        colors=colors,
        fonts=ThemeFonts(face("a:majorFont/a:latin"), face("a:minorFont/a:latin"),
                         face("a:majorFont/a:ea"), face("a:minorFont/a:ea")),
        _scheme=scheme, _color_map=mapping, _document=document, _master=master_part,
    )


class _SchemeTheme:
    __slots__ = ("color_scheme",)

    def __init__(self, scheme: dict) -> None:
        self.color_scheme = scheme


def resolve_with(color: "str | Color", scheme: dict, mapping: dict) -> str:
    """``color`` as ``"#RRGGBB"`` under a theme's scheme and a colour map."""
    from ooxml_common.drawingml.color import ColorContext, resolve_color
    from ooxml_common.drawingml.read import parse_color_node

    from .color import Color

    parsed = Color.parse(color) if isinstance(color, str) else color
    source = parse_color_node(parsed.to_element())
    if source is None:
        raise ValueError(f"cannot resolve the colour {parsed}")
    context = ColorContext(_SchemeTheme(scheme), mapping)
    resolved = resolve_color(context, source)
    if resolved is None:
        raise ValueError(f"cannot resolve the colour {parsed}")
    return "#" + resolved.hex.lstrip("#").upper()


def theme_fill_style(slide, index: int) -> Element | None:
    """The fill a shape style's ``fillRef idx`` names in its master's theme: 1-999 count
    the format scheme's ``fillStyleLst``, 1001 and up its ``bgFillStyleLst``; ``None``
    for 0 (no fill) or past the list."""
    package = slide.document.package
    layout = inherit.layout_of(package, slide.part_path)
    master = inherit.master_of(package, layout) or first_master(package)
    part = theme_part(package, master) if master else None
    root = package.tree(part) if part else None
    if root is None or index <= 0:
        return None
    path = "a:themeElements/a:fmtScheme/" + ("a:bgFillStyleLst" if index > 1000
                                             else "a:fillStyleLst")
    styles = find(root, path)
    entries = [child for child in styles if isinstance(child.tag, str)] if styles is not None else []
    position = (index - 1001) if index > 1000 else index - 1
    return entries[position] if 0 <= position < len(entries) else None


def theme_for(target) -> Theme:
    """The theme a :class:`Document` (its first master), :class:`Slide` or shape draws with,
    through the colour-map overrides of its layout and slide."""
    from .document import Document, Shape, Slide

    if isinstance(target, Document):
        package = target.package
        return read_theme(package, first_master(package), document=target)
    if isinstance(target, Shape):
        target = target._slide
    if isinstance(target, Slide):
        package = target.document.package
        layout = inherit.layout_of(package, target.part_path)
        master = inherit.master_of(package, layout) or first_master(package)
        overrides = []
        for part in (layout, target.part_path):
            root = package.tree(part) if part else None
            overrides.append(find(root, "p:clrMapOvr") if root is not None else None)
        return read_theme(package, master, overrides, document=target.document)
    raise TypeError(f"a colour resolves against a Document, a Slide or a Shape, not "
                    f"{type(target).__name__}")


# -- roles and tints --------------------------------------------------------------------------

#: The slots a design draws its colours from (dk1 and lt1 are text and background).
DESIGN_SLOTS = ("dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
RAMP_SLOTS = ("dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6")
_ACCENTS = DESIGN_SLOTS[1:]


@dataclass(frozen=True)
class ThemeRoles:
    """Which theme colours do which job in a template (:attr:`Theme.roles`).

    ``primary`` is the colour the template carries its identity in -- titles, bars,
    rules; ``secondary`` the one that partners it; ``neutral`` the quiet one boxes and
    panels are tinted from; ``highlight`` the one kept for the one thing that must stand
    out.  Each is a slot name (``"dk2"``, ``"accent1"``...), ``colors`` maps each role to
    its ``"#RRGGBB"``, ``source`` says for each role whether the template showed it
    (``"master"``) or the convention filled it in (``"convention"``), and ``usage``
    counts how often the master and layouts use each design slot.

    **The rule.**  Count every theme colour the master and its layouts use themselves --
    shape fills, outlines and text colours, the master's background and its title style
    -- with ``tx2``/``bg2`` mapped through the colour map, and leave out ``dk1``/``lt1``
    (text and background) and the hyperlink colours.  ``primary`` is the most used of
    ``dk2`` and ``accent1``-``accent6``, ``secondary`` the next; a tie goes to ``dk2``,
    then the accents in order.  ``highlight`` is the most used remaining accent.  Where
    the template uses fewer, the convention fills the gap -- Office's own: ``dk2``
    primary, ``accent1`` secondary, ``accent2`` highlight (the first of these not taken).
    ``neutral`` is ``lt2``.

    For example::

        roles = deck.theme.roles
        roles.primary, roles.colors["primary"]       # ('dk2', '#0E2841')
    """

    primary: str
    secondary: str
    neutral: str
    highlight: str
    colors: dict[str, str]
    source: dict[str, str]
    usage: dict[str, int]


@dataclass(frozen=True)
class Tint:
    """One variant of a theme colour (:meth:`Theme.tints`): ``name`` as a colour picker
    shows it (``"Lighter 80%"``), ``color`` the colour string to write (``"accent1
    lumMod=20% lumOff=80%"``, a theme reference that follows the theme) and ``hex`` what
    it looks like under this theme.

    For example::

        tint = deck.theme.tints("accent1")[0]
        box.fill = tint.color
    """

    name: str
    color: str
    hex: str


def _usage(theme: Theme) -> Counter:
    document = theme._document
    counts: Counter = Counter()
    if document is None or theme._master is None:
        return counts
    package = document.package
    parts = [theme._master] + [part for part in package.related_parts_of_type(
        theme._master, REL_SLIDE_LAYOUT) if package.has_part(part)]
    for part in parts:
        root = package.tree(part)
        if root is None:
            continue
        holders = [find(root, "p:cSld/p:bg"), find(root, "p:cSld/p:spTree"),
                   find(root, "p:txStyles/p:titleStyle/a:lvl1pPr")]
        for holder in holders:
            if holder is None:
                continue
            for node in holder.iter(qn("a:schemeClr")):
                name = node.get("val")
                slot = theme._color_map.get(name, name) if name in MAPPED else name
                if slot in DESIGN_SLOTS:
                    counts[slot] += 1
    return counts


def theme_roles(theme: Theme) -> ThemeRoles:
    counts = _usage(theme)
    order = {slot: index for index, slot in enumerate(DESIGN_SLOTS)}
    used = sorted((slot for slot in counts if counts[slot]),
                  key=lambda slot: (-counts[slot], order[slot]))
    roles: dict[str, str] = {}
    source: dict[str, str] = {}
    for role, fallback in (("primary", ("dk2", "accent1", "accent2")),
                           ("secondary", ("accent1", "dk2", "accent2")),
                           ("highlight", ("accent2",) + _ACCENTS)):
        taken = set(roles.values())
        pool = used if role != "highlight" else [slot for slot in used if slot in _ACCENTS]
        found = next((slot for slot in pool if slot not in taken), None)
        if found is not None:
            roles[role], source[role] = found, "master"
        else:
            roles[role] = next(slot for slot in fallback + _ACCENTS if slot not in taken)
            source[role] = "convention"
    roles["neutral"], source["neutral"] = "lt2", "convention"
    return ThemeRoles(primary=roles["primary"], secondary=roles["secondary"],
                      neutral=roles["neutral"], highlight=roles["highlight"],
                      colors={role: theme.colors[slot] for role, slot in roles.items()},
                      source=source, usage={slot: counts.get(slot, 0) for slot in DESIGN_SLOTS})


def _luminance(hex_color: str) -> float:
    red, green, blue = (int(hex_color.lstrip("#")[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return (max(red, green, blue) + min(red, green, blue)) / 2


#: (name, lumMod %, lumOff %) for a mid-tone colour, and for very dark and very light ones.
_MID_STEPS = (("Lighter 80%", 20, 80), ("Lighter 60%", 40, 60), ("Lighter 40%", 60, 40),
              ("Darker 25%", 75, 0), ("Darker 50%", 50, 0))
_DARK_STEPS = (("Lighter 90%", 10, 90), ("Lighter 75%", 25, 75), ("Lighter 50%", 50, 50),
               ("Lighter 25%", 75, 25), ("Lighter 10%", 90, 10))
_LIGHT_STEPS = (("Darker 10%", 90, 0), ("Darker 25%", 75, 0), ("Darker 50%", 50, 0),
                ("Darker 75%", 25, 0), ("Darker 90%", 10, 0))


def theme_tints(theme: Theme, slot: str) -> tuple[Tint, ...]:
    if slot not in theme.colors:
        raise ValueError(f"{slot!r} is not a theme colour; they are {', '.join(SLOTS)}")
    luminance = _luminance(theme.colors[slot])
    steps = _DARK_STEPS if luminance < 0.2 else _LIGHT_STEPS if luminance > 0.8 else _MID_STEPS
    out = []
    for name, modulation, offset in steps:
        color = f"{slot} lumMod={modulation}%" + (f" lumOff={offset}%" if offset else "")
        out.append(Tint(name, color, theme.resolve(color)))
    return tuple(out)


def resolve_font(typeface: str | None, theme: Theme) -> str | None:
    """A typeface with the theme's ``+mj-lt``/``+mn-lt``/``+mj-ea``/``+mn-ea`` resolved."""
    if typeface is None:
        return None
    return {"+mj-lt": theme.fonts.major, "+mn-lt": theme.fonts.minor,
            "+mj-ea": theme.fonts.major_east_asian,
            "+mn-ea": theme.fonts.minor_east_asian}.get(typeface, typeface)


__all__ = ["Theme", "ThemeFonts", "ThemeRoles", "Tint", "color_map", "read_theme",
           "resolve_font", "resolve_with", "theme_for", "theme_roles", "theme_tints"]
