"""Author a graphic as SVG, apply it to a slide as native, editable PowerPoint objects.

SPIKE (branch ``svg-profile-spike``): a prototype of an *authoring profile* -- a small,
restricted SVG vocabulary an agent writes, which :func:`apply_svg_graphic` turns into real
shapes: presets where a preset is named, custom geometry for free paths, text frames with
wrap, anchor and insets, glued connectors, groups, theme colours.  It is not the full-state
SVG of :meth:`Document.apply_svg`; it never writes ``data-ooxml-*``.

The coordinate system: the root ``viewBox`` maps onto ``box`` (EMU ``left, top, width,
height``; by default :func:`content_area`), honouring ``preserveAspectRatio`` (default
``xMidYMid meet``, as SVG does).  Font sizes, stroke widths, insets and paragraph spacing
are in **points** and are not scaled by the viewBox.

Everything is checked before anything changes; then the whole graphic is created as one
undo step (rolled back if anything fails).
"""

from __future__ import annotations

import html.entities
import math
import re
from dataclasses import dataclass, field
from typing import Any

from lxml import etree

from ..oxml.xml import PRESET_GEOMETRIES, make, replace_choice, subelement

SVG_NS = "http://www.w3.org/2000/svg"
EMU_PER_PT = 12700
EMU_PER_IN = 914400

THEME_TOKENS = {"accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk1", "lt1",
                "dk2", "lt2", "tx1", "tx2", "bg1", "bg2", "hlink", "folHlink"}
SHAPE_TAGS = {"rect", "ellipse", "circle", "polygon", "polyline", "path"}
SUPPORTED = SHAPE_TAGS | {"svg", "g", "line", "text", "tspan", "title", "desc", "defs",
                          "linearGradient", "stop", "metadata"}
UNSUPPORTED_HINT = {
    "filter": "filters (shadows, blur) are not supported",
    "mask": "masks are not supported",
    "clipPath": "clipping is not supported",
    "pattern": "pattern fills are not supported",
    "radialGradient": "only linearGradient is supported",
    "image": "images are not supported; add pictures with slide.add_picture",
    "use": "<use> is not supported; write each element out",
    "symbol": "<symbol> is not supported",
    "style": "CSS stylesheets are not supported; use presentation attributes",
    "foreignObject": "foreignObject is not supported",
    "script": "scripts are not supported",
    "marker": "SVG markers are not supported; use data-arrow-end / data-arrow-start",
    "textPath": "text on a path is not supported",
    "switch": "<switch> is not supported",
    "a": "links are not supported",
}
UNSUPPORTED_ATTRS = {
    "class": "CSS classes are not supported; use presentation attributes",
    "filter": "filters are not supported", "mask": "masks are not supported",
    "clip-path": "clipping is not supported",
    "marker-start": "SVG markers are not supported; use data-arrow-start",
    "marker-end": "SVG markers are not supported; use data-arrow-end",
    "marker-mid": "SVG markers are not supported",
}
STYLE_PROPS = {"fill", "stroke", "stroke-width", "stroke-dasharray", "opacity", "fill-opacity",
               "stroke-opacity", "font-size", "font-weight", "font-style", "font-family",
               "text-anchor", "text-decoration", "stroke-linecap", "stroke-linejoin"}
#: Inherited presentation attributes (as in SVG); ``opacity`` is multiplied down the tree.
INHERITED = {"fill", "stroke", "stroke-width", "stroke-dasharray", "fill-opacity",
             "stroke-opacity", "font-size", "font-weight", "font-style", "font-family",
             "text-anchor", "text-decoration", "stroke-linecap"}
DASH_STYLES = {"solid", "dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot",
               "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"}
ARROWS = {"none", "triangle", "stealth", "diamond", "oval", "arrow"}
SIDES = {"top", "right", "bottom", "left"}


class SvgProfileError(ValueError):
    """The SVG uses something outside the authoring profile, or uses it wrongly.
    Nothing on the slide has changed."""


# -- result ------------------------------------------------------------------------------------


@dataclass
class GraphicResult:
    """What :func:`apply_svg_graphic` made, and what the slide's checks say now."""

    #: Every shape created, in z-order (back to front), groups included.
    shapes: list[str] = field(default_factory=list)
    #: SVG ``id`` (or ``data-name``) -> the shape id it became.
    ids: dict[str, str] = field(default_factory=dict)
    #: Shape id -> its :class:`TextFit`, for every shape that got text.
    text_fits: dict[str, Any] = field(default_factory=dict)
    #: ``deck.overflows()`` for this slide, after the change (text, overlap, off_slide).
    overflows: list = field(default_factory=list)
    #: ``slide.collisions()`` after the change.
    collisions: list = field(default_factory=list)
    #: Notes on approximations (a dash pattern mapped to a preset, a skew dropped...).
    warnings: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        """No overflow, no collision, nothing off the slide."""
        return not self.overflows

    def __str__(self) -> str:
        lines = [f"{len(self.shapes)} shapes; {len(self.overflows)} problems"]
        for problem in self.overflows:
            lines.append(f"  {problem}")
        for warning in self.warnings:
            lines.append(f"  warning: {warning}")
        return "\n".join(lines)


# -- the content area --------------------------------------------------------------------------


def content_area(slide) -> tuple[int, int, int, int]:
    """The slide's free content area, EMU ``(left, top, width, height)``: the library's
    :attr:`Slide.content_area` (LP17), which began here."""
    return slide.content_area


def suggested_viewbox(box: tuple[int, int, int, int]) -> str:
    """A viewBox in points for ``box``: one user unit is one point, so sizes read true."""
    return f"0 0 {box[2] / EMU_PER_PT:g} {box[3] / EMU_PER_PT:g}"


# -- affine transforms -------------------------------------------------------------------------


class Affine:
    __slots__ = ("a", "b", "c", "d", "e", "f")

    def __init__(self, a=1.0, b=0.0, c=0.0, d=1.0, e=0.0, f=0.0):
        self.a, self.b, self.c, self.d, self.e, self.f = a, b, c, d, e, f

    def __matmul__(self, o: "Affine") -> "Affine":
        return Affine(self.a * o.a + self.c * o.b, self.b * o.a + self.d * o.b,
                      self.a * o.c + self.c * o.d, self.b * o.c + self.d * o.d,
                      self.a * o.e + self.c * o.f + self.e, self.b * o.e + self.d * o.f + self.f)

    def apply(self, x: float, y: float) -> tuple[float, float]:
        return self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f

    def scales(self) -> tuple[float, float]:
        return math.hypot(self.a, self.b), math.hypot(self.c, self.d)

    def angle(self) -> float:
        return math.degrees(math.atan2(self.b, self.a))

    def det(self) -> float:
        return self.a * self.d - self.b * self.c

    def skewed(self) -> bool:
        sx, sy = self.scales()
        if sx == 0 or sy == 0:
            return False
        return abs((self.a * self.c + self.b * self.d) / (sx * sy)) > 1e-6


_NUM = r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?"


def _numbers(text: str) -> list[float]:
    return [float(n) for n in re.findall(_NUM, text or "")]


def parse_transform(text: str, where: str) -> Affine:
    result = Affine()
    for name, args in re.findall(r"([a-zA-Z]+)\s*\(([^)]*)\)", text or ""):
        v = _numbers(args)
        if name == "translate":
            m = Affine(e=v[0], f=v[1] if len(v) > 1 else 0.0)
        elif name == "scale":
            m = Affine(a=v[0], d=v[1] if len(v) > 1 else v[0])
        elif name == "rotate":
            r = math.radians(v[0])
            m = Affine(math.cos(r), math.sin(r), -math.sin(r), math.cos(r))
            if len(v) == 3:
                m = Affine(e=v[1], f=v[2]) @ m @ Affine(e=-v[1], f=-v[2])
        elif name == "matrix" and len(v) == 6:
            m = Affine(*v)
        else:
            raise SvgProfileError(f"{where}: transform {name}() is not supported "
                                  "(use translate, scale, rotate)")
        result = result @ m
    return result


# -- colours -----------------------------------------------------------------------------------


def parse_color(value: str, where: str, opacity: float = 1.0) -> str | None:
    """A profile colour -> a pptx-agent colour string; ``None`` for ``none``."""
    v = (value or "").strip()
    if v in ("", "none", "transparent"):
        return None
    base, _, modifier = v.partition("/")
    mods = []
    if base.startswith("#"):
        hexpart = base[1:]
        if len(hexpart) == 3:
            hexpart = "".join(ch * 2 for ch in hexpart)
        if not re.fullmatch(r"[0-9a-fA-F]{6}", hexpart):
            raise SvgProfileError(f"{where}: {value!r} is not a colour")
        color = "#" + hexpart.upper()
    elif base in THEME_TOKENS:
        color = base
    else:
        raise SvgProfileError(
            f"{where}: colour {value!r} is not in the profile: use #RRGGBB or a theme token "
            f"({', '.join(sorted(THEME_TOKENS))}), optionally /tintNN or /shadeNN")
    if modifier:
        m = re.fullmatch(r"(tint|shade)(\d{1,2})", modifier)
        if not m:
            raise SvgProfileError(f"{where}: colour modifier {modifier!r}: use /tint10../tint90 "
                                  "or /shade10../shade90")
        amount = int(m.group(2))
        if m.group(1) == "tint":       # lighter by amount%: lumMod=(100-a)% lumOff=a%
            mods.append(f"lumMod={100 - amount}% lumOff={amount}%")
        else:                          # darker by amount%: lumMod=(100-a)%
            mods.append(f"lumMod={100 - amount}%")
    if opacity < 0.999:
        mods.append(f"alpha={round(max(0.0, opacity) * 100)}%")
    return " ".join([color] + mods)


# -- path data ---------------------------------------------------------------------------------


def parse_path(d: str, where: str) -> list[list[tuple]]:
    """SVG path data -> subpaths of absolute segments: ("M", (x, y)), ("L", (x, y)),
    ("C", c1, c2, p), ("Z",).  Quadratics and arcs become cubics."""
    tokens = re.findall(r"[MmLlHhVvCcSsQqTtAaZz]|" + _NUM, d or "")
    if not tokens or tokens[0] not in "Mm":
        raise SvgProfileError(f"{where}: path data must start with M")
    out: list[list[tuple]] = []
    i, cmd = 0, None
    x = y = sx = sy = 0.0
    last_c2 = last_q = None
    current: list[tuple] = []

    def nums(n):
        nonlocal i
        vals = []
        for _ in range(n):
            if i >= len(tokens) or re.fullmatch(r"[A-Za-z]", tokens[i]):
                raise SvgProfileError(f"{where}: path data {d[:60]!r}: command {cmd} needs {n} numbers")
            vals.append(float(tokens[i]))
            i += 1
        return vals

    while i < len(tokens):
        if re.fullmatch(r"[A-Za-z]", tokens[i]):
            cmd = tokens[i]
            i += 1
        elif cmd is None:
            raise SvgProfileError(f"{where}: bad path data")
        rel = cmd.islower()
        c = cmd.upper()
        if c == "Z":
            current.append(("Z",))
            x, y = sx, sy
            last_c2 = last_q = None
            continue
        if c == "M":
            px, py = nums(2)
            if rel:
                px, py = x + px, y + py
            if current:
                out.append(current)
            current = [("M", (px, py))]
            x, y, sx, sy = px, py, px, py
            cmd = "l" if rel else "L"
            last_c2 = last_q = None
            continue
        if not current:
            current = [("M", (x, y))]
        if c in "LHV":
            if c == "L":
                px, py = nums(2)
                if rel:
                    px, py = x + px, y + py
            elif c == "H":
                (px,) = nums(1)
                px, py = (x + px if rel else px), y
            else:
                (py,) = nums(1)
                px, py = x, (y + py if rel else py)
            current.append(("L", (px, py)))
            x, y = px, py
            last_c2 = last_q = None
        elif c in "CS":
            if c == "C":
                v = nums(6)
                c1 = (v[0], v[1]); c2 = (v[2], v[3]); p = (v[4], v[5])
                if rel:
                    c1 = (x + c1[0], y + c1[1]); c2 = (x + c2[0], y + c2[1]); p = (x + p[0], y + p[1])
            else:
                v = nums(4)
                c2 = (v[0], v[1]); p = (v[2], v[3])
                if rel:
                    c2 = (x + c2[0], y + c2[1]); p = (x + p[0], y + p[1])
                c1 = (2 * x - last_c2[0], 2 * y - last_c2[1]) if last_c2 else (x, y)
            current.append(("C", c1, c2, p))
            x, y = p
            last_c2, last_q = c2, None
        elif c in "QT":
            if c == "Q":
                v = nums(4)
                q = (v[0], v[1]); p = (v[2], v[3])
                if rel:
                    q = (x + q[0], y + q[1]); p = (x + p[0], y + p[1])
            else:
                v = nums(2)
                p = (v[0], v[1])
                if rel:
                    p = (x + p[0], y + p[1])
                q = (2 * x - last_q[0], 2 * y - last_q[1]) if last_q else (x, y)
            c1 = (x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y))
            c2 = (p[0] + 2 / 3 * (q[0] - p[0]), p[1] + 2 / 3 * (q[1] - p[1]))
            current.append(("C", c1, c2, p))
            x, y = p
            last_q, last_c2 = q, None
        elif c == "A":
            rx, ry, phi, large, sweep, px, py = nums(7)
            if rel:
                px, py = x + px, y + py
            for seg in _arc_to_cubics(x, y, rx, ry, phi, bool(large), bool(sweep), px, py):
                current.append(seg)
            x, y = px, py
            last_c2 = last_q = None
        else:
            raise SvgProfileError(f"{where}: path command {cmd} is not supported")
    if current:
        out.append(current)
    return out


def _arc_to_cubics(x1, y1, rx, ry, phi, large, sweep, x2, y2):
    if (x1, y1) == (x2, y2):
        return []
    if rx == 0 or ry == 0:
        return [("L", (x2, y2))]
    rx, ry = abs(rx), abs(ry)
    p = math.radians(phi)
    cp, sp = math.cos(p), math.sin(p)
    dx, dy = (x1 - x2) / 2, (y1 - y2) / 2
    x1p, y1p = cp * dx + sp * dy, -sp * dx + cp * dy
    lam = (x1p / rx) ** 2 + (y1p / ry) ** 2
    if lam > 1:
        rx, ry = rx * math.sqrt(lam), ry * math.sqrt(lam)
    num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p
    den = rx * rx * y1p * y1p + ry * ry * x1p * x1p
    coef = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        coef = -coef
    cxp, cyp = coef * rx * y1p / ry, -coef * ry * x1p / rx
    cx = cp * cxp - sp * cyp + (x1 + x2) / 2
    cy = sp * cxp + cp * cyp + (y1 + y2) / 2

    def ang(ux, uy, vx, vy):
        a = math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
        return a

    t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
    dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
    if not sweep and dt > 0:
        dt -= 2 * math.pi
    elif sweep and dt < 0:
        dt += 2 * math.pi
    n = max(1, math.ceil(abs(dt) / (math.pi / 2)))
    step = dt / n
    k = 4 / 3 * math.tan(step / 4)
    segs = []

    def point(t):
        return (cx + rx * math.cos(t) * cp - ry * math.sin(t) * sp,
                cy + rx * math.cos(t) * sp + ry * math.sin(t) * cp)

    def deriv(t):
        return (-rx * math.sin(t) * cp - ry * math.cos(t) * sp,
                -rx * math.sin(t) * sp + ry * math.cos(t) * cp)

    t = t1
    for _ in range(n):
        p0, p3 = point(t), point(t + step)
        d0, d3 = deriv(t), deriv(t + step)
        c1 = (p0[0] + k * d0[0], p0[1] + k * d0[1])
        c2 = (p3[0] - k * d3[0], p3[1] - k * d3[1])
        segs.append(("C", c1, c2, p3))
        t += step
    return segs


# -- the intermediate model --------------------------------------------------------------------


@dataclass
class Run:
    text: str
    fmt: dict


@dataclass
class Para:
    runs: list[Run]
    align: str | None = None
    bullet: str | None = None
    level: int = 0
    space_before: float | None = None
    space_after: float | None = None
    line_spacing: float | None = None


@dataclass
class TextSpec:
    paragraphs: list[Para]
    anchor: str
    insets: tuple | None
    wrap: bool


@dataclass
class Op:
    kind: str                    # "shape" | "custom" | "line" | "group-start" | "group-end"
    where: str
    key: str | None = None       # svg id / data-name
    name: str | None = None
    preset: str | None = None
    adjustments: dict | None = None
    frame: tuple | None = None   # left, top, width, height in EMU (unrotated)
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False
    fill: Any = "none"           # colour string, None (no fill) or ("gradient", stops, angle)
    line: dict | None = None     # None -> no outline
    text: TextSpec | None = None
    textbox: bool = False
    path: list | None = None     # custom geometry in frame-local EMU
    points: tuple | None = None  # line: ((x1, y1), (x2, y2)) in EMU
    connect: dict | None = None  # from, to, from_side, to_side, kind
    group_id: int | None = None


# -- parsing -----------------------------------------------------------------------------------


def _local(el) -> str:
    tag = el.tag
    if not isinstance(tag, str):
        return ""
    if tag.startswith("{"):
        ns, _, name = tag[1:].partition("}")
        return name if ns == SVG_NS else f"{{{ns}}}{name}"
    return tag


def _where(el) -> str:
    """``<rect id='q1'> (line 4)``, or ``<rect> #3 (line 4)`` -- the third <rect> in the
    document -- for an element without an id."""
    name = _local(el)
    ident = el.get("id") or el.get("data-name")
    if ident:
        return f"<{name} id={ident!r}> (line {el.sourceline})"
    index = 0
    try:
        for index, other in enumerate(el.getroottree().getroot().iter(el.tag), 1):
            if other is el:
                break
    except (TypeError, ValueError):
        index = 0
    number = f" #{index}" if index else ""
    return f"<{name}>{number} (line {el.sourceline})"


def _length(value, where, what) -> float:
    if value is None:
        raise SvgProfileError(f"{where}: {what} is required")
    m = re.fullmatch(r"\s*(" + _NUM + r")\s*(pt|px)?\s*", str(value))
    if not m:
        raise SvgProfileError(f"{where}: {what}={value!r} is not a number (units: none, pt or px; "
                              "no %, em or in)")
    return float(m.group(1))


def _attrs(el, inherited: dict) -> dict:
    """Presentation attributes: inherited ones, then the element's, then its style=""."""
    out = {k: v for k, v in inherited.items() if k in INHERITED}
    out["opacity"] = inherited.get("opacity", 1.0)
    for key in STYLE_PROPS:
        if el.get(key) is not None:
            out[key] = el.get(key).strip()
    style = el.get("style")
    if style:
        for decl in style.split(";"):
            if not decl.strip():
                continue
            k, _, v = decl.partition(":")
            k, v = k.strip(), v.strip()
            if k not in STYLE_PROPS:
                raise SvgProfileError(f"{_where(el)}: style property {k!r} is not in the profile")
            out[k] = v
    own = el.get("opacity")
    if own is None and style:
        m = re.search(r"(?:^|;)\s*opacity\s*:\s*(" + _NUM + ")", style)
        own = m.group(1) if m else None
    if own is not None:
        out["opacity"] = inherited.get("opacity", 1.0) * float(own)
    return out


class _Parser:
    def __init__(self, root, box):
        self.root = root
        self.ops: list[Op] = []
        self.warnings: list[str] = []
        self.gradients: dict[str, Any] = {}
        self.keys: dict[str, str] = {}
        self.group_counter = 0
        vb = root.get("viewBox")
        if vb:
            vx, vy, vw, vh = _numbers(vb)
        elif root.get("width") and root.get("height"):
            vx, vy = 0.0, 0.0
            vw, vh = _length(root.get("width"), "<svg>", "width"), _length(root.get("height"), "<svg>", "height")
        else:
            raise SvgProfileError("<svg>: give a viewBox (or width and height)")
        if vw <= 0 or vh <= 0:
            raise SvgProfileError("<svg>: the viewBox must have a positive width and height")
        L, T, W, H = box
        par = (root.get("preserveAspectRatio") or "xMidYMid meet").split()
        if par[0] == "none":
            sx, sy = W / vw, H / vh
            ox, oy = L, T
        else:
            meet = len(par) < 2 or par[1] == "meet"
            s = min(W / vw, H / vh) if meet else max(W / vw, H / vh)
            sx = sy = s
            align = par[0]
            ax = {"xMin": 0.0, "xMid": 0.5, "xMax": 1.0}[align[:4]]
            ay = {"YMin": 0.0, "YMid": 0.5, "YMax": 1.0}[align[4:]]
            ox = L + (W - vw * s) * ax
            oy = T + (H - vh * s) * ay
            if abs(W / vw - H / vh) / max(W / vw, H / vh) > 0.01:
                self.warnings.append(
                    f"viewBox aspect {vw:g}x{vh:g} differs from the box's "
                    f"{W / EMU_PER_PT:.1f}x{H / EMU_PER_PT:.1f} pt; scaled uniformly and centred")
        self.base = Affine(sx, 0, 0, sy, ox - vx * sx, oy - vy * sy)

    # .. walking ..

    def parse(self):
        self._collect_defs(self.root)
        self._children(self.root, self.base, {"opacity": 1.0}, None)
        return self.ops

    def _check(self, el):
        name = _local(el)
        if name in UNSUPPORTED_HINT:
            raise SvgProfileError(f"{_where(el)}: {UNSUPPORTED_HINT[name]}")
        if name not in SUPPORTED:
            raise SvgProfileError(f"{_where(el)}: <{name}> is not in the authoring profile")
        for attr, why in UNSUPPORTED_ATTRS.items():
            if el.get(attr) is not None:
                raise SvgProfileError(f"{_where(el)}: attribute {attr}: {why}")

    def _collect_defs(self, root):
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            self._check(el)
            if _local(el) == "linearGradient":
                gid = el.get("id")
                if not gid:
                    raise SvgProfileError(f"{_where(el)}: a gradient needs an id")
                x1, y1 = float(el.get("x1", "0").rstrip("%")), float(el.get("y1", "0").rstrip("%"))
                x2, y2 = float(el.get("x2", "1").rstrip("%")), float(el.get("y2", "0").rstrip("%"))
                if el.get("x2", "").endswith("%") or el.get("x1", "").endswith("%"):
                    pass
                angle = math.degrees(math.atan2(y2 - y1, x2 - x1)) % 360
                stops = []
                for st in el:
                    if _local(st) != "stop":
                        continue
                    off = st.get("offset", "0").strip()
                    pos = float(off[:-1]) / 100 if off.endswith("%") else float(off)
                    style = dict(
                        (k.strip(), v.strip()) for k, _, v in
                        (d.partition(":") for d in (st.get("style") or "").split(";") if d.strip()))
                    color = st.get("stop-color") or style.get("stop-color")
                    op = float(st.get("stop-opacity") or style.get("stop-opacity") or 1)
                    c = parse_color(color, _where(st), op)
                    if c is None:
                        raise SvgProfileError(f"{_where(st)}: a gradient stop needs a colour")
                    stops.append((pos, c))
                if len(stops) < 2:
                    raise SvgProfileError(f"{_where(el)}: a gradient needs two stops")
                self.gradients[gid] = ("gradient", stops, angle)

    def _children(self, el, ctm, inherited, group_id):
        for child in el:
            if not isinstance(child.tag, str):
                continue
            self._element(child, ctm, inherited, group_id)

    def _register(self, el, op: Op):
        key = el.get("id") or el.get("data-name")
        if key:
            if key in self.keys:
                raise SvgProfileError(f"{_where(el)}: id {key!r} is used twice")
            self.keys[key] = op.kind
        op.key = key
        op.name = el.get("data-name") or el.get("id")

    def _element(self, el, ctm, inherited, group_id):
        name = _local(el)
        if name in ("title", "desc", "defs", "metadata", "linearGradient", "stop"):
            return
        if el.get("display") == "none" or el.get("visibility") == "hidden":
            return
        where = _where(el)
        if el.get("transform"):
            ctm = ctm @ parse_transform(el.get("transform"), where)
        attrs = _attrs(el, inherited)
        if name == "g":
            if el.get("data-text-box") is not None:
                self._text_box_group(el, ctm, attrs, group_id)
                return
            self.group_counter += 1
            gid = self.group_counter
            start = Op("group-start", where, group_id=group_id)
            start.key = el.get("id") or el.get("data-name")
            start.name = el.get("data-name") or el.get("id")
            if start.key:
                if start.key in self.keys:
                    raise SvgProfileError(f"{where}: id {start.key!r} is used twice")
                self.keys[start.key] = "group"
            start.adjustments = {"gid": gid}
            self.ops.append(start)
            self._children(el, ctm, attrs, gid)
            end = Op("group-end", where, group_id=group_id)
            end.adjustments = {"gid": gid}
            self.ops.append(end)
            return
        if name == "text":
            self._free_text(el, ctm, attrs, group_id)
            return
        if name == "tspan":
            raise SvgProfileError(f"{where}: <tspan> belongs inside <text>")
        if name == "line" or (name in ("path", "polyline") and
                              (el.get("data-from") or el.get("data-to"))):
            self._line(el, ctm, attrs, group_id)
            return
        if name in SHAPE_TAGS:
            texts = [c for c in el if isinstance(c.tag, str) and _local(c) == "text"]
            op = self._shape(el, ctm, attrs, group_id)
            if texts:
                if len(texts) > 1:
                    raise SvgProfileError(f"{where}: one <text> per shape")
                op.text = self._text_spec(texts[0], _attrs(texts[0], attrs), [el], default_anchor="middle")
            self._register(el, op)
            self.ops.append(op)
            return
        raise SvgProfileError(f"{where}: <{name}> is not in the authoring profile")

    # .. paint ..

    def _fill(self, attrs, where):
        value = attrs.get("fill")
        if value is None:
            value = "tx1"               # SVG's default fill is black: the theme's dark text colour
        value = value.strip()
        m = re.fullmatch(r"url\(\s*#([^)\s]+)\s*\)", value)
        if m:
            if m.group(1) not in self.gradients:
                raise SvgProfileError(f"{where}: fill refers to unknown gradient {m.group(1)!r}")
            return self.gradients[m.group(1)]
        op = attrs.get("opacity", 1.0) * float(attrs.get("fill-opacity", 1))
        return parse_color(value, where, op)

    def _line_spec(self, el, attrs, where, scale):
        value = attrs.get("stroke")
        if value is None or value.strip() in ("none", ""):
            return None
        op = attrs.get("opacity", 1.0) * float(attrs.get("stroke-opacity", 1))
        color = parse_color(value, where, op)
        if color is None:
            return None
        width = _length(attrs.get("stroke-width", "1"), where, "stroke-width")
        spec = {"color": color, "width": max(1, round(width * EMU_PER_PT))}
        dash = el.get("data-dash")
        if dash:
            if dash not in DASH_STYLES:
                raise SvgProfileError(f"{where}: data-dash={dash!r}; use one of {sorted(DASH_STYLES)}")
            spec["dash"] = dash
        elif attrs.get("stroke-dasharray") and attrs["stroke-dasharray"].strip() != "none":
            pattern = _numbers(attrs["stroke-dasharray"])
            spec["dash"] = _dash_preset(pattern, width)
            self.warnings.append(f"{where}: stroke-dasharray {attrs['stroke-dasharray']!r} "
                                 f"drawn as PowerPoint's {spec['dash']!r}")
        for attr, key in (("data-arrow-end", "end"), ("data-arrow-start", "start")):
            arrow = el.get(attr)
            if arrow:
                if arrow not in ARROWS:
                    raise SvgProfileError(f"{where}: {attr}={arrow!r}; use one of {sorted(ARROWS)}")
                spec[key] = arrow
        return spec

    # .. geometry ..

    def _frame(self, ctm: Affine, x, y, w, h, where) -> tuple:
        """An axis-aligned local box under ``ctm`` -> EMU frame, rotation, flips."""
        if ctm.skewed():
            raise SvgProfileError(f"{where}: the transform skews this shape; only translate, "
                                  "scale and rotate are supported")
        cx, cy = ctm.apply(x + w / 2, y + h / 2)
        sx, sy = ctm.scales()
        W, H = abs(w) * sx, abs(h) * sy
        flip = ctm.det() < 0
        angle = ctm.angle()
        if flip:
            # a reflection: express it as a horizontal flip and the remaining rotation
            angle = math.degrees(math.atan2(-ctm.b, -ctm.a))
        angle = round(angle % 360, 4)
        if abs(angle - 360) < 1e-3:
            angle = 0.0
        frame = (round(cx - W / 2), round(cy - H / 2), max(1, round(W)), max(1, round(H)))
        return frame, angle, flip

    def _shape(self, el, ctm, attrs, group_id) -> Op:
        name = _local(el)
        where = _where(el)
        preset = el.get("data-preset")
        adjustments = {}
        for key, value in el.attrib.items():
            if isinstance(key, str) and key.startswith("data-adj-"):
                adjustments[key[len("data-adj-"):]] = int(round(float(value)))
        op = Op("shape", where, group_id=group_id)
        op.fill = self._fill(attrs, where)
        op.line = self._line_spec(el, attrs, where, ctm)
        if name == "rect":
            x, y = float(el.get("x", 0)), float(el.get("y", 0))
            w = _length(el.get("width"), where, "width")
            h = _length(el.get("height"), where, "height")
            rx = el.get("rx") or el.get("ry")
            if not preset:
                preset = "rect"
                if rx and float(rx) > 0:
                    preset = "roundRect"
                    adjustments.setdefault("adj", min(50000, round(float(rx) / min(w, h) * 100000)))
            op.frame, op.rotation, op.flip_h = self._frame(ctm, x, y, w, h, where)
        elif name in ("ellipse", "circle"):
            cx, cy = float(el.get("cx", 0)), float(el.get("cy", 0))
            if name == "circle":
                rx = ry = _length(el.get("r"), where, "r")
            else:
                rx, ry = _length(el.get("rx"), where, "rx"), _length(el.get("ry"), where, "ry")
            preset = preset or "ellipse"
            op.frame, op.rotation, op.flip_h = self._frame(ctm, cx - rx, cy - ry, 2 * rx, 2 * ry, where)
        else:   # path, polygon, polyline
            if name == "path":
                subpaths = parse_path(el.get("d"), where)
            else:
                pts = _numbers(el.get("points"))
                if len(pts) < 4 or len(pts) % 2:
                    raise SvgProfileError(f"{where}: points needs pairs of numbers")
                pairs = list(zip(pts[::2], pts[1::2]))
                sub = [("M", pairs[0])] + [("L", p) for p in pairs[1:]]
                if name == "polygon":
                    sub.append(("Z",))
                subpaths = [sub]
            if preset:
                xs, ys = _extent(subpaths)
                op.frame, op.rotation, op.flip_h = self._frame(
                    ctm, min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys), where)
            else:
                if ctm.skewed():
                    pass   # custom geometry takes any affine map: points are transformed
                mapped = [[_map_seg(seg, ctm) for seg in sub] for sub in subpaths]
                xs, ys = _extent(mapped)
                left, top = min(xs), min(ys)
                op.kind = "custom"
                op.frame = (round(left), round(top), max(1, round(max(xs) - left)),
                            max(1, round(max(ys) - top)))
                op.path = [[_offset_seg(seg, left, top) for seg in sub] for sub in mapped]
                op.closed = [any(seg[0] == "Z" for seg in sub) for sub in mapped]
                op.preset = None
                op.adjustments = None
                if el.get("data-preset") is None and name == "polyline" and op.fill not in (None,) \
                        and attrs.get("fill") is None:
                    op.fill = None   # an open polyline with no fill stated: drawn as a line
                return op
        if preset not in PRESET_GEOMETRIES:
            raise SvgProfileError(f"{where}: data-preset={preset!r} is not a PowerPoint preset "
                                  "(e.g. rect, roundRect, chevron, homePlate, ellipse, diamond)")
        if adjustments:
            _check_adjustments(preset, adjustments, where)
        op.preset = preset
        op.adjustments = adjustments or None
        return op

    def _line(self, el, ctm, attrs, group_id):
        where = _where(el)
        op = Op("line", where, group_id=group_id)
        if attrs.get("stroke") is None:
            attrs = dict(attrs, stroke="tx1")      # a line with no stroke would be invisible
        op.line = self._line_spec(el, attrs, where, ctm)
        if op.line is None:
            raise SvgProfileError(f"{where}: a line needs a stroke")
        name = _local(el)
        if name == "line":
            p1 = ctm.apply(float(el.get("x1", 0)), float(el.get("y1", 0)))
            p2 = ctm.apply(float(el.get("x2", 0)), float(el.get("y2", 0)))
        else:
            sub = parse_path(el.get("d"), where) if name == "path" else None
            if sub is None:
                pts = _numbers(el.get("points"))
                sub = [[("M", (pts[0], pts[1])), ("L", (pts[-2], pts[-1]))]]
            first, last = sub[0][0][1], sub[-1][-1][-1] if sub[-1][-1][0] != "Z" else sub[0][0][1]
            p1, p2 = ctm.apply(*first), ctm.apply(*last)
        op.points = ((round(p1[0]), round(p1[1])), (round(p2[0]), round(p2[1])))
        frm, to = el.get("data-from"), el.get("data-to")
        kind = el.get("data-connector")
        if kind and kind not in ("straight", "elbow", "curved"):
            raise SvgProfileError(f"{where}: data-connector={kind!r}: straight, elbow or curved")
        for attr in ("data-from-side", "data-to-side"):
            if el.get(attr) and el.get(attr) not in SIDES:
                raise SvgProfileError(f"{where}: {attr}={el.get(attr)!r}: top, right, bottom or left")
        op.connect = {"from": frm, "to": to, "from_side": el.get("data-from-side"),
                      "to_side": el.get("data-to-side"), "kind": kind}
        self._register(el, op)
        self.ops.append(op)

    # .. text ..

    def _text_box_group(self, g, ctm, attrs, group_id):
        where = _where(g)
        shapes = [c for c in g if isinstance(c.tag, str) and _local(c) in SHAPE_TAGS]
        texts = [c for c in g if isinstance(c.tag, str) and _local(c) == "text"]
        others = [c for c in g if isinstance(c.tag, str)
                  and _local(c) not in SHAPE_TAGS | {"text", "title", "desc"}]
        if len(shapes) != 1 or len(texts) != 1 or others:
            raise SvgProfileError(f"{where}: a data-text-box group holds exactly one shape "
                                  "(rect, ellipse, path...) and one <text>")
        shape_el = shapes[0]
        shape_ctm = ctm @ parse_transform(shape_el.get("transform"), _where(shape_el)) \
            if shape_el.get("transform") else ctm
        op = self._shape(shape_el, shape_ctm, _attrs(shape_el, attrs), group_id)
        if op.kind == "custom":
            raise SvgProfileError(f"{where}: a text box's shape must be a rect, ellipse or a "
                                  "path with data-preset (custom geometry cannot hold text here)")
        text_attrs = _attrs(texts[0], attrs)
        op.text = self._text_spec(texts[0], text_attrs, [g, shape_el], default_anchor="middle")
        self._register(g, op)
        if shape_el.get("id") and not g.get("id"):
            op.key = shape_el.get("id")
            self.keys[op.key] = "shape"
        self.ops.append(op)

    def _free_text(self, el, ctm, attrs, group_id):
        where = _where(el)
        if el.get("data-width") is None or el.get("data-height") is None:
            raise SvgProfileError(
                f"{where}: text must be in a box: give the <text> data-width and data-height "
                "(x, y are then its top-left corner), or put it in a <g data-text-box> with a shape")
        x, y = float(el.get("x", 0)), float(el.get("y", 0))
        w = _length(el.get("data-width"), where, "data-width")
        h = _length(el.get("data-height"), where, "data-height")
        op = Op("shape", where, group_id=group_id)
        op.textbox = True
        op.preset = "rect"
        op.frame, op.rotation, op.flip_h = self._frame(ctm, x, y, w, h, where)
        box_fill = el.get("data-fill")
        op.fill = parse_color(box_fill, where) if box_fill else None
        stroke = el.get("data-stroke")
        op.line = ({"color": parse_color(stroke, where),
                    "width": round(_length(el.get("data-stroke-width", "1"), where, "data-stroke-width")
                                   * EMU_PER_PT)} if stroke and stroke != "none" else None)
        op.text = self._text_spec(el, attrs, [], default_anchor="top")
        self._register(el, op)
        self.ops.append(op)

    def _text_spec(self, text_el, attrs, holders, default_anchor) -> TextSpec:
        """``holders``: the elements that may carry the frame's attributes (data-anchor,
        data-inset, data-wrap) besides the <text> itself -- the shape, and the
        ``<g data-text-box>`` around it.  The most specific one wins: the <text>, then the
        shape, then the group."""
        where = _where(text_el)
        if not isinstance(holders, (list, tuple)):
            holders = [holders]
        chain = [text_el] + [h for h in reversed(holders) if h is not text_el]

        def frame_attr(name):
            for element in chain:
                if element.get(name) is not None:
                    return element.get(name)
            return None

        anchor = frame_attr("data-anchor") or default_anchor
        if anchor not in ("top", "middle", "bottom"):
            raise SvgProfileError(f"{where}: data-anchor={anchor!r}: top, middle or bottom")
        inset_attr = frame_attr("data-inset")
        insets = None
        if inset_attr is not None:
            v = _numbers(inset_attr)
            if len(v) == 1:
                v = v * 4
            elif len(v) == 2:
                v = [v[1], v[0], v[1], v[0]]   # CSS order: vertical, horizontal
            elif len(v) != 4:
                raise SvgProfileError(f"{where}: data-inset takes 1, 2 or 4 numbers (pt): "
                                      "all; vertical horizontal; or left top right bottom")
            insets = tuple(round(n * EMU_PER_PT) for n in v)
        wrap = (frame_attr("data-wrap") or "wrap") != "none"
        base_align = _align(attrs.get("text-anchor"), text_el.get("data-align"), where)
        paragraphs: list[Para] = []
        tspans = [c for c in text_el if isinstance(c.tag, str) and _local(c) == "tspan"]
        if tspans:
            if (text_el.text or "").strip() or any((c.tail or "").strip() for c in tspans):
                raise SvgProfileError(f"{where}: with <tspan> paragraphs, put all the text inside "
                                      "them (loose text between them is ambiguous)")
            for child in text_el:
                if not isinstance(child.tag, str):
                    continue
                if _local(child) != "tspan":
                    raise SvgProfileError(f"{_where(child)}: only <tspan> goes inside <text>")
                pattrs = _attrs(child, attrs)
                runs = self._runs(child, pattrs)
                para = Para(runs=runs, align=_align(pattrs.get("text-anchor") if child.get("text-anchor")
                                                    else None, child.get("data-align"), where) or base_align)
                bullet = child.get("data-bullet")
                if bullet is not None:
                    para.bullet = bullet or "•"
                level = child.get("data-level")
                para.level = int(level) if level else 0
                for attr, key in (("data-space-before", "space_before"),
                                  ("data-space-after", "space_after")):
                    val = child.get(attr) or text_el.get(attr)
                    if val is not None:
                        setattr(para, key, _length(val, where, attr))
                ls = child.get("data-line-spacing") or text_el.get("data-line-spacing")
                para.line_spacing = float(ls) if ls else None
                paragraphs.append(para)
        else:
            runs = self._runs(text_el, attrs)
            para = Para(runs=runs, align=base_align)
            ls = text_el.get("data-line-spacing")
            para.line_spacing = float(ls) if ls else None
            paragraphs.append(para)
        for para in paragraphs:
            _collapse(para)
        if not any(r.text for p in paragraphs for r in p.runs):
            raise SvgProfileError(f"{where}: the text is empty")
        return TextSpec(paragraphs, anchor, insets, wrap)

    def _runs(self, el, attrs) -> list[Run]:
        out: list[Run] = []
        fmt = _run_format(attrs, _where(el))
        if el.text:
            out.append(Run(el.text, fmt))
        for child in el:
            if not isinstance(child.tag, str):
                if child.tail:
                    out.append(Run(child.tail, fmt))
                continue
            if _local(child) != "tspan":
                raise SvgProfileError(f"{_where(child)}: only <tspan> goes inside <text>")
            if child.get("data-bullet") is not None:
                raise SvgProfileError(f"{_where(child)}: data-bullet goes on a paragraph <tspan> "
                                      "(a direct child of <text>)")
            out.extend(self._runs(child, _attrs(child, attrs)))
            if child.tail:
                out.append(Run(child.tail, fmt))
        return out

    def report_keys(self):
        return self.keys


def _check_adjustments(preset: str, adjustments: dict, where: str) -> None:
    """Every ``data-adj-NAME`` must be one of the preset's own adjustment names."""
    from .presets import adjustment_defaults

    known = adjustment_defaults(preset)
    if not known:
        return      # no table for this preset: add_shape decides
    names = [name for name, _ in known]
    wrong = [name for name in adjustments if name not in names]
    if wrong:
        listing = ", ".join(f"data-adj-{name} (default {value})" for name, value in known)
        raise SvgProfileError(f"{where}: {preset} has no adjustment "
                              f"{', '.join(repr(w) for w in wrong)}; it has {listing}")


def _collapse(para: Para):
    """SVG whitespace: newlines and runs of spaces become one space; trimmed at the ends."""
    for r in para.runs:
        r.text = re.sub(r"\s+", " ", r.text.replace(" ", "\x00")).replace("\x00", " ")
    # drop double spaces across run boundaries
    prev_space = True
    for r in para.runs:
        if prev_space and r.text.startswith(" "):
            r.text = r.text[1:]
        if r.text:
            prev_space = r.text.endswith(" ")
    for r in reversed(para.runs):
        if r.text:
            r.text = r.text.rstrip(" ")
            break
    para.runs = [r for r in para.runs if r.text]


def _align(text_anchor, data_align, where):
    if data_align:
        table = {"left": "l", "center": "ctr", "centre": "ctr", "right": "r", "justify": "just"}
        if data_align not in table:
            raise SvgProfileError(f"{where}: data-align={data_align!r}: left, center, right, justify")
        return table[data_align]
    if text_anchor:
        table = {"start": "l", "middle": "ctr", "end": "r"}
        if text_anchor not in table:
            raise SvgProfileError(f"{where}: text-anchor={text_anchor!r}")
        return table[text_anchor]
    return None


def _run_format(attrs, where) -> dict:
    fmt: dict = {}
    fmt["size"] = _length(attrs.get("font-size", "12"), where, "font-size")
    weight = (attrs.get("font-weight") or "normal").strip()
    fmt["bold"] = weight == "bold" or weight == "bolder" or (weight.isdigit() and int(weight) >= 600)
    fmt["italic"] = (attrs.get("font-style") or "normal").strip() in ("italic", "oblique")
    if "underline" in (attrs.get("text-decoration") or ""):
        fmt["underline"] = True
    color = attrs.get("fill")
    fmt["color"] = parse_color(color if color is not None else "tx1", where,
                               float(attrs.get("fill-opacity", 1)))
    family = attrs.get("font-family")
    if family:
        family = family.split(",")[0].strip().strip("'\"")
        fmt["typeface"] = {"major": "+mj-lt", "minor": "+mn-lt"}.get(family, family)
    return fmt


def _dash_preset(pattern, width):
    if not pattern:
        return "solid"
    first = pattern[0]
    if len(pattern) >= 4:
        return "sysDashDot" if first <= 3 * max(width, 1) else "dashDot"
    if first <= 1.5 * max(width, 1):
        return "sysDot"
    if first <= 3.5 * max(width, 1):
        return "sysDash"
    if first >= 7 * max(width, 1):
        return "lgDash"
    return "dash"


def _map_seg(seg, ctm):
    if seg[0] == "Z":
        return seg
    return (seg[0],) + tuple(ctm.apply(*p) for p in seg[1:])


def _offset_seg(seg, dx, dy):
    if seg[0] == "Z":
        return seg
    return (seg[0],) + tuple((p[0] - dx, p[1] - dy) for p in seg[1:])


def _extent(subpaths):
    xs, ys = [], []
    for sub in subpaths:
        for seg in sub:
            for p in seg[1:]:
                xs.append(p[0]); ys.append(p[1])
    return xs, ys


# -- reading the markup ------------------------------------------------------------------------

_XML_REFERENCE = re.compile(r"&(?:(amp|lt|gt|quot|apos)|#[0-9]+|#x[0-9a-fA-F]+);")
_NAMED_REFERENCE = re.compile(r"&([A-Za-z][A-Za-z0-9]*);")
_CDATA = re.compile(r"(<!\[CDATA\[.*?\]\]>|<!--.*?-->)", re.S)


def repair_markup(text: str) -> tuple[str, list[str]]:
    """Make hand-written SVG well-formed where its intent is unambiguous.

    * A bare ``&`` that does not start a valid reference (``&amp;``, ``&lt;``, ``&#8211;``,
      ``&#x2013;``...) is escaped: ``R&D`` and ``Q&A`` mean the characters.
    * An HTML entity name (``&nbsp;``, ``&ndash;``, ``&rarr;``) becomes the character it
      names, as a numeric reference; an unknown name (``&foo;``) is kept as literal text.
    * A ``<`` followed by a space, a digit or ``=`` (``< 5%``, ``<=``) can never start a tag
      and is escaped.

    Comments and CDATA sections are left alone.  Returns the text and one warning per kind
    of repair made."""
    counts = {"amp": 0, "html": 0, "unknown": [], "lt": 0}

    def fix(chunk: str) -> str:
        def named(m):
            name = m.group(1)
            if name in ("amp", "lt", "gt", "quot", "apos"):
                return m.group(0)
            code = html.entities.name2codepoint.get(name)
            if code is not None:
                counts["html"] += 1
                return f"&#{code};"
            counts["unknown"].append(name)
            return "&amp;" + name + ";"

        chunk = _NAMED_REFERENCE.sub(named, chunk)
        out, pos = [], 0
        for m in re.finditer("&", chunk):
            out.append(chunk[pos:m.start()])
            if _XML_REFERENCE.match(chunk, m.start()):
                out.append("&")
            else:
                out.append("&amp;")
                counts["amp"] += 1
            pos = m.end()
        out.append(chunk[pos:])
        chunk = "".join(out)

        def less(m):
            counts["lt"] += 1
            return "&lt;"

        return re.sub(r"<(?=[\s0-9=])", less, chunk)

    parts = _CDATA.split(text)
    repaired = "".join(part if i % 2 else fix(part) for i, part in enumerate(parts))
    warnings = []
    if counts["amp"]:
        warnings.append(f"escaped {counts['amp']} bare '&' as '&amp;' (a literal ampersand)")
    if counts["html"]:
        warnings.append(f"{counts['html']} HTML entity reference(s) written as the characters")
    if counts["unknown"]:
        names = ", ".join(sorted(set(counts["unknown"])))
        warnings.append(f"unknown entity reference(s) kept as literal text: {names}")
    if counts["lt"]:
        warnings.append(f"escaped {counts['lt']} '<' that could not start a tag as '&lt;'")
    return repaired, warnings


def _nearest_element(text: str, line: int, column: int) -> str:
    """Name the element a parse error at ``line``/``column`` (1-based) falls in or after:
    ``<rect id='q1'>`` or ``<text> #4`` (the fourth <text>), with the line it starts on."""
    lines = text.split("\n")
    offset = sum(len(l) + 1 for l in lines[:max(0, line - 1)]) + max(0, column - 1)
    head = text[:offset + 1]
    starts = list(re.finditer(r"<([A-Za-z][\w:.-]*)", head))
    if not starts:
        return "before the first element"
    last = starts[-1]
    tag = last.group(1)
    start_line = text.count("\n", 0, last.start()) + 1
    tag_text = text[last.start():text.find(">", last.start()) + 1 or len(text)]
    ident = re.search(r"\s(?:id|data-name)\s*=\s*[\"']([^\"']*)[\"']", tag_text)
    local = tag.split(":")[-1]
    if ident:
        return f"<{local} id={ident.group(1)!r}> (starts line {start_line})"
    index = sum(1 for m in starts if m.group(1) == tag)
    return f"<{local}> #{index} (starts line {start_line})"


def parse_svg(svg: "str | bytes"):
    """Parse profile SVG: :func:`repair_markup`, then a safe XML parse (no DTD, no entities,
    no network).  Returns ``(root, repair warnings)``; raises :class:`SvgProfileError`
    naming the nearest element when the markup is still not well-formed."""
    text = svg.decode("utf-8") if isinstance(svg, (bytes, bytearray)) else str(svg)
    if "<!ENTITY" in text or "<!DOCTYPE" in text:
        raise SvgProfileError("a DTD or entity is not allowed")
    text, repairs = repair_markup(text)
    parser_opts = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False,
                                  remove_comments=True, remove_pis=True)
    try:
        root = etree.fromstring(text.encode("utf-8"), parser_opts)
    except etree.XMLSyntaxError as error:
        line, column = (error.position if getattr(error, "position", None) else (error.lineno or 1, 1))
        near = _nearest_element(text, line or 1, column or 1)
        source_line = text.split("\n")[max(0, (line or 1) - 1)] if text else ""
        lo = max(0, (column or 1) - 40)
        excerpt = source_line[lo:(column or 1) + 40].strip()
        message = getattr(error, "msg", None) or str(error)
        raise SvgProfileError(f"not well-formed XML in or after {near}: {message} "
                              f"(line {line}, column {column}: ...{excerpt}...)") from None
    if root.getroottree().docinfo.internalDTD is not None:
        raise SvgProfileError("a DTD or entity is not allowed")
    if _local(root) != "svg":
        raise SvgProfileError("the root element must be <svg>")
    return root, repairs


# -- data positions through scales -------------------------------------------------------------

#: Attributes that hold an x (or y) position, and the part of a date or band each means.
_X_ATTRS = {"x": "start", "data-x2": "end", "cx": "center", "data-cx": "center",
            "x1": "center", "x2": "center"}
_Y_ATTRS = {"y": "start", "data-y2": "end", "cy": "center", "data-cy": "center",
            "y1": "center", "y2": "center"}


def resolve_scales(root, box, scales) -> int:
    """Turn data positions into user units, in place; returns how many were turned.

    Inside an element carrying ``data-scale-x="$time"`` (or ``data-scale-y``), a position
    attribute whose value starts with ``@`` is a data value of that scale: ``x="@2026-11-16"``,
    ``cy="@Category strategy"``, ``x2="@7.5"``.  ``|start``, ``|center`` or ``|end`` picks
    the part of a date or band (defaults: ``x``/``y`` the start, ``data-x2``/``data-y2`` the
    end, centres and line ends the centre).  A scale maps onto **slide points**, as in the
    shape tools; the point is then mapped into the SVG's user units through the viewBox.
    ``data-x2``/``data-y2`` (a right or bottom edge) and ``data-cx``/``data-cy`` (a centre)
    give a ``rect``'s, or a boxed ``text``'s, width and height.  ``scales`` maps a name
    (without ``$``) to a :class:`~pptx_agent.edit.scales.Scale`.
    """
    from .scales import ScaleError

    parser = _Parser(root, box)
    base = parser.base
    count = 0

    def scope(el, axis):
        node = el
        while node is not None:
            ref = node.get(f"data-scale-{axis}")
            if ref is not None:
                return ref, node
            node = node.getparent()
        return None, None

    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        for axis, attrs in (("x", _X_ATTRS), ("y", _Y_ATTRS)):
            for name, part in attrs.items():
                value = el.get(name)
                if value is None or not value.strip().startswith("@"):
                    continue
                where = _where(el)
                ref, holder = scope(el, axis)
                if ref is None:
                    raise SvgProfileError(f"{where}: {name}={value!r} is a data value but no "
                                          f"ancestor has data-scale-{axis}")
                key = ref.strip().lstrip("$")
                if scales is None or key not in scales:
                    known = ", ".join(f"${k}" for k in sorted(scales or {})) or "none"
                    raise SvgProfileError(f"{where}: data-scale-{axis}={ref!r} is not a scale "
                                          f"(scales: {known})")
                node = el
                while node is not None and node is not root:
                    if node.get("transform"):
                        raise SvgProfileError(f"{where}: data values cannot sit under a "
                                              f"transform ({_where(node)})")
                    node = node.getparent()
                raw = value.strip()[1:]
                if "|" in raw:
                    raw, part = raw.rsplit("|", 1)
                    if part not in ("start", "center", "end"):
                        raise SvgProfileError(f"{where}: {name}: |{part} must be |start, "
                                              "|center or |end")
                try:
                    points = scales[key].position(raw.strip(), part)
                except ScaleError as exc:
                    raise SvgProfileError(f"{where}: {name}: {exc}") from None
                emu = points * EMU_PER_PT
                user = (emu - base.e) / base.a if axis == "x" else (emu - base.f) / base.d
                el.set(name, f"{user:.4f}")
                count += 1
        for axis, size in (("x", "width"), ("y", "height")):
            boxed = _local(el) == "text"
            size_attr = f"data-{size}" if boxed else size
            end, centre = el.get(f"data-{axis}2"), el.get(f"data-c{axis}")
            if end is not None:
                start = el.get(axis)
                if start is None:
                    raise SvgProfileError(f"{_where(el)}: data-{axis}2 needs {axis}")
                el.set(size_attr, f"{float(end) - float(start):.4f}")
                del el.attrib[f"data-{axis}2"]
            elif centre is not None:
                length = el.get(size_attr)
                if length is None:
                    raise SvgProfileError(f"{_where(el)}: data-c{axis} needs {size_attr}")
                el.set(axis, f"{float(centre) - float(length) / 2:.4f}")
                del el.attrib[f"data-c{axis}"]
    return count


# -- building ----------------------------------------------------------------------------------


def _custom_geometry(op: Op):
    w, h = op.frame[2], op.frame[3]
    geom = make("a:custGeom")
    for tag in ("a:avLst", "a:gdLst", "a:ahLst", "a:cxnLst"):
        geom.append(make(tag))
    geom.append(make("a:rect", l="l", t="t", r="r", b="b"))
    path_list = make("a:pathLst")
    geom.append(path_list)
    path = make("a:path", w=str(w), h=str(h))
    path_list.append(path)

    def pt(p):
        return make("a:pt", x=str(round(p[0])), y=str(round(p[1])))

    for sub in op.path:
        for seg in sub:
            if seg[0] == "M":
                node = make("a:moveTo"); node.append(pt(seg[1]))
            elif seg[0] == "L":
                node = make("a:lnTo"); node.append(pt(seg[1]))
            elif seg[0] == "C":
                node = make("a:cubicBezTo")
                for p in seg[1:]:
                    node.append(pt(p))
            else:
                node = make("a:close")
            path.append(node)
    return geom


def _site_for(shape, other_center, side):
    if side:
        return shape.connection_site(side)
    l, t, w, h = shape.slide_bounds
    cx, cy = l + w / 2, t + h / 2
    dx, dy = other_center[0] - cx, other_center[1] - cy
    if abs(dx) * h >= abs(dy) * w:
        return shape.connection_site("right" if dx > 0 else "left")
    return shape.connection_site("bottom" if dy > 0 else "top")


def auto_sides(source_box, target_box) -> tuple[str, str]:
    """The sides a connector between two shapes uses when none is given.

    A target lying wholly below the source (its top at or under the source's bottom) is
    joined bottom -> top, and one wholly above top -> bottom, however far it is to the
    side: the tree layouts of org charts and hierarchies.  Otherwise a target wholly to
    the right or left is joined side to side, and overlapping boxes fall back to the
    direction between the centres."""
    al, at, aw, ah = source_box
    bl, bt, bw, bh = target_box
    if bt >= at + ah:
        return "bottom", "top"
    if bt + bh <= at:
        return "top", "bottom"
    if bl >= al + aw:
        return "right", "left"
    if bl + bw <= al:
        return "left", "right"
    dx = (bl + bw / 2) - (al + aw / 2)
    dy = (bt + bh / 2) - (at + ah / 2)
    if abs(dx) * ah >= abs(dy) * aw:
        return ("right", "left") if dx > 0 else ("left", "right")
    return ("bottom", "top") if dy > 0 else ("top", "bottom")


def _center(shape):
    l, t, w, h = shape.slide_bounds
    return l + w / 2, t + h / 2


def _apply_text(shape, spec: TextSpec):
    frame = shape.text_frame
    frame.autofit = "none"
    frame.wrap = spec.wrap
    frame.anchor = spec.anchor
    if spec.insets is not None:
        frame.insets = spec.insets
    shape.set_text("\n".join("".join(r.text for r in p.runs) for p in spec.paragraphs))
    for index, para in enumerate(spec.paragraphs):
        paragraph = frame.paragraph(index)
        texts = [r.text for r in para.runs]
        if len(texts) > 1:
            paragraph.segment(texts)
        for r_index, run in enumerate(para.runs):
            paragraph.run(r_index).format(**run.fmt)
        paragraph.alignment = para.align or "l"
        if para.space_before is not None:
            paragraph.space_before = round(para.space_before * EMU_PER_PT)
        if para.space_after is not None:
            paragraph.space_after = round(para.space_after * EMU_PER_PT)
        if para.line_spacing is not None:
            paragraph.line_spacing = para.line_spacing
        if para.bullet is not None:
            size = para.runs[0].fmt.get("size", 12) if para.runs else 12
            paragraph.set_bullet(para.bullet)
            if para.level:
                paragraph.level = para.level
                hang = -(paragraph.indent or 0)
                paragraph.margin_left = (paragraph.margin_left or 0) + para.level * hang
            del size
        elif para.level:
            paragraph.margin_left = para.level * 171450


def _apply_paint(shape, op: Op):
    fill = op.fill
    if isinstance(fill, tuple) and fill and fill[0] == "gradient":
        shape.set_gradient_fill(fill[1], angle=fill[2])
    elif fill is None:
        shape.fill = "none"
    else:
        shape.fill = fill
    if op.line is None:
        shape.line.visible = False
    else:
        for key, value in op.line.items():
            setattr(shape.line, key, value)


def apply_svg_graphic(slide, svg: "str | bytes", *, box: "tuple[int, int, int, int] | None" = None,
                      scales=None) -> GraphicResult:
    """Apply an SVG in the authoring profile (see PROFILE.md) to ``slide`` as native shapes.

    ``box`` is where the SVG's viewBox lands, EMU ``(left, top, width, height)``; by default
    :func:`content_area` of the slide.  Raises :class:`SvgProfileError` naming the element for
    anything outside the profile, before anything changes.  Returns a :class:`GraphicResult`:
    the shapes made, SVG id -> shape id, each text shape's :class:`TextFit`, and the slide's
    ``overflows()`` and ``collisions()`` afterwards.  ``scales`` (name -> scale) serves
    data positions (:func:`resolve_scales`).
    """
    root, repairs = parse_svg(svg)
    if box is None:
        box = content_area(slide)
    box = tuple(int(v) for v in box)
    resolve_scales(root, box, scales)
    parser = _Parser(root, box)
    parser.warnings[:0] = repairs
    ops = parser.parse()
    # connector targets must exist
    for op in ops:
        if op.kind == "line" and op.connect:
            for end in ("from", "to"):
                key = op.connect[end]
                if key and key not in parser.keys:
                    raise SvgProfileError(f"{op.where}: data-{end}={key!r} names no element")
                if key and parser.keys[key] in ("line", "group"):
                    raise SvgProfileError(f"{op.where}: data-{end}={key!r} must name a shape, "
                                          f"not a {parser.keys[key]}")
    result = GraphicResult(warnings=list(parser.warnings))
    deck = slide.document
    made: dict[str, Any] = {}
    created: list[Any] = []
    group_members: dict[int | None, list] = {}
    group_meta: dict[int, Op] = {}
    group_parent: dict[int, int | None] = {}
    pending: list[tuple[Any, Op]] = []
    with deck.batch():
        for op in ops:
            if op.kind == "group-start":
                gid = op.adjustments["gid"]
                group_meta[gid] = op
                group_parent[gid] = op.group_id
                group_members.setdefault(gid, [])
                group_members.setdefault(op.group_id, []).append(("group", gid))
                continue
            if op.kind == "group-end":
                continue
            if op.kind == "line":
                c = op.connect or {}
                kind = c.get("kind") or "straight"
                shape = slide.add_connector(kind, op.points[0], op.points[1],
                                            name=op.name, line=op.line)
                if c.get("from") or c.get("to"):
                    pending.append((shape, op))
            elif op.kind == "custom":
                l, t, w, h = op.frame
                shape = slide.add_shape("rect", l, t, w, h, name=op.name)
                if not op.name:
                    shape.name = "Freeform: Shape " + shape.id.split(".")[-1].split("#")[0]
                shape._before_change()
                properties = subelement(shape._element, "p:spPr")
                replace_choice(properties, ("a:custGeom", "a:prstGeom"), _custom_geometry(op))
                shape._slide._touch()
                _apply_paint(shape, op)
            else:
                l, t, w, h = op.frame
                if op.textbox:
                    shape = slide.add_textbox(l, t, w, h, name=op.name, autofit="none")
                else:
                    shape = slide.add_shape(op.preset, l, t, w, h, name=op.name,
                                            adjustments=op.adjustments)
                if op.rotation:
                    shape.rotation = op.rotation
                if op.flip_h:
                    shape.flip_h = True
                _apply_paint(shape, op)
                if op.text is not None:
                    _apply_text(shape, op.text)
            created.append(shape)
            if op.key:
                made[op.key] = shape
            group_members.setdefault(op.group_id, []).append(("shape", shape))
        # glue connectors, now every shape exists
        for shape, op in pending:
            c = op.connect
            begin = end = None
            a = made.get(c["from"]) if c["from"] else None
            b = made.get(c["to"]) if c["to"] else None
            far_a = _center(b) if b is not None else op.points[1]
            far_b = _center(a) if a is not None else op.points[0]
            from_side, to_side = c["from_side"], c["to_side"]
            if a is not None and b is not None:
                auto_from, auto_to = auto_sides(a.slide_bounds, b.slide_bounds)
                from_side, to_side = from_side or auto_from, to_side or auto_to
            if a is not None:
                begin = _site_for(a, far_a, from_side)
            if b is not None:
                end = _site_for(b, far_b, to_side)
            if not c.get("kind") and a is not None and b is not None:
                # straight when the two sites line up, elbow otherwise
                pa = a.connection_sites[begin[1]] if hasattr(a, "connection_sites") else None
                pb = b.connection_sites[end[1]] if hasattr(b, "connection_sites") else None
                if pa is not None and pb is not None:
                    pa, pb = _site_xy(pa), _site_xy(pb)
                    if pa and pb and abs(pa[0] - pb[0]) > 6350 and abs(pa[1] - pb[1]) > 6350:
                        _set_connector_kind(shape, "elbow")
            shape.connect(begin=begin, end=end)
        # groups, innermost first
        group_shapes: dict[int, Any] = {}

        def depth(gid):
            d = 0
            while group_parent.get(gid) is not None:
                gid = group_parent[gid]
                d += 1
            return d

        for gid in sorted(group_meta, key=depth, reverse=True):
            members = []
            for kind, item in group_members.get(gid, []):
                if kind == "shape":
                    members.append(item)
                elif item in group_shapes:
                    members.append(group_shapes[item])
                else:
                    members.extend(_flatten(item, group_members, group_shapes))
            if len(members) >= 2:
                meta = group_meta[gid]
                group = slide.group(members, name=meta.name)
                group_shapes[gid] = group
                created.append(group)
                if meta.key:
                    made[meta.key] = group
    # report
    order = {s.id: i for i, s in enumerate(slide.shapes)}
    result.shapes = sorted({s.id for s in created}, key=lambda i: order.get(i, 0))
    result.ids = {k: v.id for k, v in made.items()}
    for shape in created:
        if shape.kind == "shape" and shape.text:
            result.text_fits[shape.id] = shape.text_fit()
    prefix = f"{slide.slide_id}."
    result.overflows = [o for o in deck.overflows() if str(getattr(o, "slide", "")) == str(slide.slide_id)
                        or str(getattr(o, "shape", "")).startswith(prefix)]
    result.collisions = slide.collisions()
    return result


def _flatten(gid, members, built):
    out = []
    for kind, item in members.get(gid, []):
        if kind == "shape":
            out.append(item)
        elif item in built:
            out.append(built[item])
        else:
            out.extend(_flatten(item, members, built))
    return out


def _site_xy(site):
    if isinstance(site, (tuple, list)) and len(site) >= 2 and all(
            isinstance(v, (int, float)) for v in site[:2]):
        return site[0], site[1]
    for attrs in (("x", "y"),):
        if all(hasattr(site, a) for a in attrs):
            return getattr(site, "x"), getattr(site, "y")
    return None


def _set_connector_kind(shape, kind):
    geometry = {"straight": "straightConnector1", "elbow": "bentConnector3",
                "curved": "curvedConnector3"}[kind]
    shape.preset = geometry


# -- measuring before applying -------------------------------------------------------------------


@dataclass
class TextBoxMeasure:
    """One text shape of an SVG, laid out exactly as :func:`apply_svg_graphic` would make it
    (:func:`measure_svg_text`).  Lengths are in points.

    ``needed`` is the height the text takes inside the frame (lines, spacing, top and bottom
    insets) and ``available`` the frame's text height for the box as drawn; ``fits`` is
    PowerPoint's verdict (see :class:`TextFit`).  ``height_to_fit`` is the least height to give
    the shape, in SVG user units (the ``height`` attribute; points with
    :func:`suggested_viewbox`), for the text to fit at the same width, rounded up to 0.5.  ``lines`` is how many lines each
    paragraph wraps to; ``near_wrap`` warns that a line ends within 0.16 pt of the edge."""

    key: str
    shape_id: str
    width: float
    height: float
    needed: float
    available: float
    fits: bool
    overflow: float
    height_to_fit: float
    lines: tuple
    sizes: tuple
    near_wrap: bool

    def as_dict(self) -> dict:
        return {k: (round(v, 2) if isinstance(v, float) else v) for k, v in self.__dict__.items()}


def _wrap_fragment(text: str, box) -> str:
    stripped = re.sub(r"^\s*(<\?xml[^>]*\?>)?\s*", "", text)
    if re.match(r"<svg[\s>]", stripped):
        return text
    return (f'<svg xmlns="{SVG_NS}" viewBox="{suggested_viewbox(box)}">' + text + "</svg>")


def measure_svg_text(slide, svg: "str | bytes", *, box: "tuple[int, int, int, int] | None" = None,
                     scales=None) -> dict[str, TextBoxMeasure]:
    """Measure every text shape of ``svg`` exactly as :func:`apply_svg_graphic` will build it
    -- the same insets, sizes, bold, bullets, spacing, wrap and shape geometry -- without
    changing ``slide``.

    ``svg`` is a whole profile SVG, or a fragment: one or more elements such as a
    ``<g data-text-box>`` or a ``<text data-width data-height>``, which is placed in a
    viewBox of :func:`suggested_viewbox` (one unit = one point).  Connectors are left out.
    Returns ``{svg id (or "#n" for the n-th unnamed text shape): TextBoxMeasure}``.

    For example::

        m = measure_svg_text(slide, '<g id="b" data-text-box="" data-inset="4 6">'
                                    '<rect width="180" height="30" fill="accent1"/>'
                                    '<text font-size="12" font-weight="bold">Programme lead</text></g>')
        m["b"].fits, m["b"].height_to_fit
    """
    if box is None:
        box = content_area(slide)
    box = tuple(int(v) for v in box)
    text = svg.decode("utf-8") if isinstance(svg, (bytes, bytearray)) else str(svg)
    root, _ = parse_svg(_wrap_fragment(text, box))
    for el in list(root.iter()):
        if isinstance(el.tag, str) and (el.get("data-from") is not None or el.get("data-to") is not None):
            el.getparent().remove(el)
    # measured on a copy of the deck, so the caller's deck, slide objects and undo history
    # are untouched
    from .document import Document

    twin_deck = Document.open(slide.document.to_bytes())
    twin = twin_deck.slides[slide.index]
    scale_y = _Parser(root, box).base.d          # EMU per SVG user unit, vertically
    result = apply_svg_graphic(twin, etree.tostring(root), box=box, scales=scales)
    by_shape = {shape_id: key for key, shape_id in result.ids.items()}
    found: dict[str, TextBoxMeasure] = {}
    unnamed = 0
    pt = EMU_PER_PT
    for shape_id, fit in result.text_fits.items():
        key = by_shape.get(shape_id)
        if key is None:
            unnamed += 1
            key = f"#{unnamed}"
        _, _, w, h = twin_deck.shape(shape_id).bounds
        # the frame's text height scales with the shape's (a preset's text rectangle is a
        # fixed share of it), so the shape grows by the missing text height times h/available
        ratio = h / fit.available if fit.available > 0 else 1.0
        least = h + (fit.needed - fit.available) * ratio
        found[key] = TextBoxMeasure(
            key=key, shape_id=shape_id, width=w / pt, height=h / pt,
            needed=fit.needed / pt, available=fit.available / pt,
            fits=not fit.overflows, overflow=fit.overflow / pt,
            height_to_fit=math.ceil(least / scale_y * 2) / 2,
            lines=tuple(fit.lines), sizes=tuple(fit.sizes),
            near_wrap=bool(getattr(fit, "near_wrap", False)))
    return found
