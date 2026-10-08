"""The ``data-ooxml-*`` vocabulary: typed values read straight from the unresolved XML.

Nothing here resolves anything.  A theme colour is read as the reference it is
(``scheme:accent1 lumMod=75000``), a position as the EMU the file holds or ``inherit`` when
the shape takes it from its layout, a picture as the package part it points at.  Every
function reads an element and nothing else, so the same code describes a shape in the
document *and* the raw XML that comes back in an SVG -- which is what lets the reader tell
what an SVG changed.

Two encodings carry the values:

* flat **attributes** on the shape's group, for scalar features (geometry, fill, outline):
  ``data-ooxml-fill="solid" data-ooxml-fill-scheme="accent1" data-ooxml-fill-mods="lumMod=75000"``;
* compact **JSON** for the nested ones (text, tables), which are trees: paragraphs of runs,
  rows of cells.  See ``ROADMAP.md`` for why JSON rather than a child-element encoding.

Every ``canonical_*`` function validates a value that came back from outside and rewrites it
in the exact form the emitter writes, so comparing two values compares meaning, not
spelling -- and so untrusted input is checked field by field before anything is applied.
"""

from __future__ import annotations

import re
from typing import Any, Callable

from ..oxml.package import REL_SLIDE, Relationship, normalize_part_path
from ..oxml.xml import (
    FILL_TAGS,
    LINE_FILL_TAGS,
    PRESET_GEOMETRIES,
    Element,
    get_int,
    local_name,
    qn,
)
from ooxml_edit.charts import model as _charts
from ooxml_edit.charts.model import (  # noqa: F401  (re-exported)
    CHART_KEYS,
    CHART_READ_ONLY,
    ChartModelError,
    chart_model,
    diagram_model,
)
from ..edit.color import FLAG_TRANSFORMS, SCHEME_COLORS, VALUED_TRANSFORMS, Color
from ..edit.fill import ARROWHEAD_SIZES, ARROWHEADS, DASH_STYLES, LINE_CAPS
from ..edit.table import BORDER_TAGS
from ..edit.text import NUMBERING_SCHEMES, UNDERLINES, _ALIGNMENT_NAMES
from .safe import FullStateError

#: The attribute prefix, and the version the emitter writes into the root ``<svg>``.
PREFIX = "data-ooxml-"
VOCABULARY = "pptx-agent/1"

#: Shape tag -> the vocabulary's ``kind`` (the same words :attr:`Shape.kind` uses).
KINDS = {"sp": "shape", "pic": "picture", "cxnSp": "connector", "grpSp": "group",
         "graphicFrame": "graphic_frame"}
SHAPE_TAG_NAMES = tuple(KINDS)

ROTATION_UNIT = 60000

#: ``a:graphicData@uri`` -> the vocabulary's ``graphic``.
GRAPHIC_KINDS = {
    "http://schemas.openxmlformats.org/drawingml/2006/table": "table",
    "http://schemas.openxmlformats.org/drawingml/2006/chart": "chart",
    "http://schemas.openxmlformats.org/drawingml/2006/diagram": "diagram",
    "http://schemas.openxmlformats.org/presentationml/2006/ole": "ole",
}

_STRIKES = {"sngStrike": True, "noStrike": False}
_INT = re.compile(r"-?\d{1,15}")
_DEGREES = re.compile(r"-?\d{1,12}(\.\d{1,12})?")


# ------------------------------------------------------------------------------------------
# Context
# ------------------------------------------------------------------------------------------


class Resolver:
    """What reading needs beyond the element: the slide's relationships and slide ids -- and
    the package, for the parts a chart or a diagram keeps its data in."""

    def __init__(self, relationships: dict[str, Relationship],
                 slide_ids: dict[str, int], package=None,
                 shape_ids: dict[str, str] | None = None) -> None:
        self._relationships = relationships
        self._slide_ids = slide_ids
        self._package = package
        #: ``cNvPr@id`` -> the shape's id on the slide, for ids that are unique there.
        self._shape_ids = shape_ids or {}

    def shape_id(self, raw: str | None) -> str | None:
        """The id of the shape a connector's ``stCxn``/``endCxn`` names, if it is unique."""
        return None if raw is None else self._shape_ids.get(raw)

    def part(self, path: str) -> Element | None:
        """A part's parsed root, ``None`` without a package or when it is missing."""
        if self._package is None or not path or not self._package.has_part(path):
            return None
        return self._package.tree(path)

    def rel(self, rel_id: str | None) -> Relationship | None:
        return None if not rel_id else self._relationships.get(rel_id)

    def slide_id(self, part: str | None) -> int | None:
        return None if part is None else self._slide_ids.get(part)

    def target(self, rel_id: str | None) -> str:
        """The part a relationship points at, ``""`` when it does not resolve."""
        rel = self.rel(rel_id)
        if rel is None:
            return ""
        return rel.target if rel.is_external else (rel.target_part or "")


# ------------------------------------------------------------------------------------------
# Small codecs
# ------------------------------------------------------------------------------------------


def degrees_text(raw: int) -> str:
    """60000ths of a degree as the shortest exact decimal: ``420000`` -> ``"7"``."""
    if raw % ROTATION_UNIT == 0:
        return str(raw // ROTATION_UNIT)
    return repr(raw / ROTATION_UNIT)


def parse_degrees(text: str, what: str) -> int:
    if not isinstance(text, str) or not _DEGREES.fullmatch(text.strip()):
        raise FullStateError(f"{what}: {text!r} is not a number of degrees")
    return round(float(text) * ROTATION_UNIT)


def parse_int(text: Any, what: str, low: int = -(2 ** 40), high: int = 2 ** 40) -> int:
    if isinstance(text, bool):
        raise FullStateError(f"{what}: expected an integer, got {text!r}")
    if isinstance(text, int):
        value = text
    elif isinstance(text, str) and _INT.fullmatch(text.strip()):
        value = int(text)
    else:
        raise FullStateError(f"{what}: {text!r} is not an integer")
    if not low <= value <= high:
        raise FullStateError(f"{what}: {value} is out of range {low}..{high}")
    return value


def _text(value: Any, what: str, limit: int = 10_000) -> str:
    if not isinstance(value, str) or len(value) > limit:
        raise FullStateError(f"{what}: expected a string")
    return value


def mods_text(color: Color) -> str:
    return " ".join(name if value is None else f"{name}={value}"
                    for name, value in color.transforms)


def parse_mods(text: str, what: str) -> tuple[tuple[str, int | None], ...]:
    result: list[tuple[str, int | None]] = []
    for token in _text(text, what).split():
        name, separator, raw = token.partition("=")
        if not separator and name in FLAG_TRANSFORMS:
            result.append((name, None))
        elif separator and name in VALUED_TRANSFORMS:
            result.append((name, parse_int(raw, f"{what} {name}")))
        else:
            raise FullStateError(f"{what}: unknown colour modifier {token!r}")
    return tuple(result)


#: Colour kind -> the field (attribute suffix, or codec prefix) that carries it.
_COLOR_FIELDS = {"scheme": "scheme", "rgb": "rgb", "system": "sys", "preset": "prst"}
_COLOR_KINDS = {field: kind for kind, field in _COLOR_FIELDS.items()}
_RGB = re.compile(r"[0-9A-F]{6}")
_NAME = re.compile(r"[A-Za-z][A-Za-z0-9]{0,63}")


def color_fields(color: Color | None) -> dict[str, str]:
    """``{"scheme": "accent1", "mods": "lumMod=75000"}`` -- never a resolved value."""
    if color is None:
        return {}
    field = _COLOR_FIELDS.get(color.kind)
    fields = {field: color.value} if field else {"color": color.kind}
    if color.transforms:
        fields["mods"] = mods_text(color)
    return fields


def _check_color_value(kind: str, value: str, what: str) -> str:
    if kind == "rgb":
        value = value.upper()
        if not _RGB.fullmatch(value):
            raise FullStateError(f"{what}: {value!r} is not a six-digit hex colour")
    elif kind == "scheme":
        if value not in SCHEME_COLORS:
            raise FullStateError(f"{what}: {value!r} is not a theme colour")
    elif not _NAME.fullmatch(value):
        raise FullStateError(f"{what}: {value!r} is not a colour name")
    return value


def canonical_color_fields(fields: dict[str, str], what: str) -> dict[str, str]:
    present = [key for key in ("scheme", "rgb", "sys", "prst", "color") if key in fields]
    if len(present) != 1:
        raise FullStateError(f"{what}: give exactly one of scheme, rgb, sys, prst")
    key = present[0]
    if key == "color":  # a colour kind the vocabulary carries opaquely (hsl, scrgb)
        result = {"color": _text(fields["color"], what, 16)}
    else:
        result = {key: _check_color_value(_COLOR_KINDS[key], _text(fields[key], what, 64), what)}
    if fields.get("mods"):
        result["mods"] = mods_text(Color("scheme", "", parse_mods(fields["mods"], what)))
    return result


def color_from_fields(fields: dict[str, str], what: str) -> Color:
    """The :class:`Color` canonical fields describe; opaque kinds cannot be written."""
    if "color" in fields:
        raise FullStateError(f"{what}: a {fields['color']} colour is not editable through "
                             f"typed attributes; edit data-ooxml-xml instead")
    key = next(k for k in ("scheme", "rgb", "sys", "prst") if k in fields)
    return Color(_COLOR_KINDS[key], fields[key], parse_mods(fields.get("mods", ""), what))


def color_text(color: Color) -> str:
    """The one-string form used inside JSON and gradient stops: ``scheme:accent1 lumMod=75000``."""
    fields = color_fields(color)
    key = next(iter(fields))
    head = f"{key}:{fields[key]}" if key != "color" else f"color:{fields['color']}"
    return f"{head} {fields['mods']}" if "mods" in fields else head


def canonical_color_text(text: str, what: str) -> str:
    head, _, mods = _text(text, what, 512).strip().partition(" ")
    key, separator, value = head.partition(":")
    if not separator:
        raise FullStateError(f"{what}: {text!r} should read like 'scheme:accent1 lumMod=75000'")
    fields = canonical_color_fields({key: value, "mods": mods}, what)
    key = next(iter(fields))
    head = f"{key}:{fields[key]}"
    return f"{head} {fields['mods']}" if "mods" in fields else head


def color_from_text(text: str, what: str) -> Color:
    head, _, mods = text.partition(" ")
    key, _, value = head.partition(":")
    return color_from_fields({key: value, "mods": mods}, what)


# ------------------------------------------------------------------------------------------
# Fills
# ------------------------------------------------------------------------------------------

FILL_KINDS = ("inherit", "none", "solid", "gradient", "image", "pattern", "group")


def fill_model(container: Element | None, resolver: Resolver,
               choices: tuple[str, ...] = FILL_TAGS) -> dict[str, str]:
    """``{"kind": ..., ...}`` for the fill a properties element holds."""
    node = None
    if container is not None:
        wanted = {qn(tag) for tag in choices}
        node = next((child for child in container if child.tag in wanted), None)
    if node is None:
        return {"kind": "inherit"}
    name = local_name(node)
    if name == "noFill":
        return {"kind": "none"}
    if name == "solidFill":
        return {"kind": "solid", **color_fields(Color.from_element(node))}
    if name == "grpFill":
        return {"kind": "group"}
    if name == "pattFill":
        return {"kind": "pattern", "pattern": node.get("prst") or ""}
    if name == "blipFill":
        blip = node.find(qn("a:blip"))
        return {"kind": "image",
                "image": resolver.target(blip.get(qn("r:embed")) if blip is not None else None)}
    # gradFill
    stops = []
    gs_list = node.find(qn("a:gsLst"))
    for stop in [] if gs_list is None else gs_list.findall(qn("a:gs")):
        color = Color.from_element(stop)
        if color is not None:
            stops.append(f"{get_int(stop, 'pos', 0) or 0} {color_text(color)}")
    model = {"kind": "gradient", "stops": ";".join(stops)}
    linear, path = node.find(qn("a:lin")), node.find(qn("a:path"))
    if linear is not None:
        model["angle"] = degrees_text(get_int(linear, "ang", 0) or 0)
    elif path is not None:
        model["path"] = path.get("path") or ""
    return model


def canonical_fill(model: dict[str, Any], what: str) -> dict[str, str]:
    if not isinstance(model, dict):
        raise FullStateError(f"{what}: expected an object")
    kind = model.get("kind")
    if kind not in FILL_KINDS:
        raise FullStateError(f"{what}: fill kind {kind!r} is not one of {FILL_KINDS}")
    allowed = {"kind"}
    result: dict[str, str] = {"kind": kind}
    if kind == "solid":
        allowed |= {"scheme", "rgb", "sys", "prst", "color", "mods"}
        result.update(canonical_color_fields(model, what))
    elif kind == "gradient":
        allowed |= {"stops", "angle", "path"}
        stops = []
        for stop in _text(model.get("stops", ""), what, 100_000).split(";"):
            position, _, color = stop.strip().partition(" ")
            stops.append(f"{parse_int(position, what, 0, 100000)} "
                         f"{canonical_color_text(color, what)}")
        if len(stops) < 2:
            raise FullStateError(f"{what}: a gradient needs at least two stops")
        result["stops"] = ";".join(stops)
        if "angle" in model:
            result["angle"] = degrees_text(parse_degrees(model["angle"], what))
        elif "path" in model:
            if model["path"] not in {"circle", "rect", "shape"}:
                raise FullStateError(f"{what}: gradient path must be circle, rect or shape")
            result["path"] = model["path"]
    elif kind == "image":
        allowed |= {"image"}
        result["image"] = normalize_part_path(_text(model.get("image", ""), what, 1024))
    elif kind == "pattern":
        allowed |= {"pattern"}
        result["pattern"] = _text(model.get("pattern", ""), what, 64)
    unknown = set(model) - allowed
    if unknown:
        raise FullStateError(f"{what}: unexpected fields {sorted(unknown)} for a {kind} fill")
    return result


# ------------------------------------------------------------------------------------------
# Outlines
# ------------------------------------------------------------------------------------------


def line_model(line: Element | None) -> dict[str, str]:
    """An ``a:ln`` (or a cell border): ``{"kind": "inherit"}`` when there is none."""
    if line is None:
        return {"kind": "inherit"}
    model: dict[str, str] = {"kind": "set"}
    if line.get("w") is not None:
        model["w"] = line.get("w") or ""
    fill = fill_model(line, Resolver({}, {}), LINE_FILL_TAGS)
    if fill["kind"] != "inherit":
        model["fill"] = fill["kind"]
        if fill["kind"] == "solid":
            model.update({k: v for k, v in fill.items() if k != "kind"})
    dash = line.find(qn("a:prstDash"))
    if dash is not None:
        model["dash"] = dash.get("val") or ""
    elif line.find(qn("a:custDash")) is not None:
        model["dash"] = "custom"
    if line.get("cap") is not None:
        model["cap"] = line.get("cap") or ""
    for end in ("head", "tail"):
        node = line.find(qn(f"a:{end}End"))
        if node is not None:
            model[end] = " ".join(v for v in (node.get("type", "none"), node.get("w"),
                                              node.get("len")) if v)
    return model


def canonical_line(model: dict[str, Any], what: str) -> dict[str, str]:
    if not isinstance(model, dict):
        raise FullStateError(f"{what}: expected an object")
    kind = model.get("kind")
    if kind == "inherit":
        if set(model) != {"kind"}:
            raise FullStateError(f"{what}: an inherited outline has no other fields")
        return {"kind": "inherit"}
    if kind != "set":
        raise FullStateError(f"{what}: outline kind must be 'inherit' or 'set'")
    allowed = {"kind", "w", "fill", "scheme", "rgb", "sys", "prst", "color", "mods", "dash",
               "cap", "head", "tail"}
    unknown = set(model) - allowed
    if unknown:
        raise FullStateError(f"{what}: unexpected outline fields {sorted(unknown)}")
    result: dict[str, str] = {"kind": "set"}
    if "w" in model:
        result["w"] = str(parse_int(model["w"], f"{what} width", 0, 20116800))
    if "fill" in model:
        fill = model["fill"]
        if fill not in {"none", "solid", "gradient", "pattern"}:
            raise FullStateError(f"{what}: outline fill {fill!r} is not none/solid/gradient/pattern")
        result["fill"] = fill
        if fill == "solid":
            result.update(canonical_color_fields(model, what))
    if "dash" in model:
        if model["dash"] not in DASH_STYLES | {"custom"}:
            raise FullStateError(f"{what}: unknown dash {model['dash']!r}")
        result["dash"] = model["dash"]
    if "cap" in model:
        if model["cap"] not in LINE_CAPS:
            raise FullStateError(f"{what}: unknown cap {model['cap']!r}")
        result["cap"] = model["cap"]
    for end in ("head", "tail"):
        if end in model:
            parts = _text(model[end], what, 64).split()
            if not parts or parts[0] not in ARROWHEADS or len(parts) > 3 \
                    or any(p not in ARROWHEAD_SIZES for p in parts[1:]):
                raise FullStateError(f"{what}: {end} must read 'type [width [length]]'")
            result[end] = " ".join(parts)
    return result


# ------------------------------------------------------------------------------------------
# Text
# ------------------------------------------------------------------------------------------

_RUN_KEYS = ("t", "fld", "b", "i", "u", "strike", "sz", "latin", "color", "fill", "link")
_PARAGRAPH_KEYS = ("algn", "lvl", "marL", "indent", "spcBef", "spcAft", "lnSpc", "bu", "c")


def _spacing(properties: Element, tag: str) -> str | None:
    node = properties.find(qn(tag))
    if node is None:
        return None
    for unit, child in (("pts", "a:spcPts"), ("pct", "a:spcPct")):
        value = get_int(node.find(qn(child)), "val")
        if value is not None:
            return f"{unit}:{value}"
    return None


def _bullet(properties: Element) -> Any:
    """The explicit bullet, as :attr:`Paragraph.bullet` reads it."""
    tags = {qn(t): local_name_ for t, local_name_ in (
        ("a:buNone", "none"), ("a:buChar", "char"), ("a:buAutoNum", "num"), ("a:buBlip", "picture"))}
    node = next((child for child in properties if child.tag in tags), None)
    if node is None:
        return None
    kind = tags[node.tag]
    if kind in {"none", "picture"}:
        return kind
    bullet: dict[str, Any] = {}
    if kind == "char":
        bullet["char"] = node.get("char") or ""
    else:
        bullet["num"] = node.get("type") or ""
        if node.get("startAt") is not None:
            bullet["start"] = get_int(node, "startAt", 1)
    font = properties.find(qn("a:buFont"))
    if font is not None and font.get("typeface") is not None:
        bullet["font"] = font.get("typeface")
    color = Color.from_element(properties.find(qn("a:buClr")))
    if color is not None:
        bullet["color"] = color_text(color)
    size = properties.find(qn("a:buSzPct"))
    if size is not None:
        bullet["size"] = get_int(size, "val", 0) or 0
    return bullet


def _hyperlink(node: Element, resolver: Resolver) -> dict[str, Any]:
    """A run's click, as :attr:`Run.hyperlink` reads it."""
    link: dict[str, Any] = {}
    rel = resolver.rel(node.get(qn("r:id")))
    action = node.get("action") or None
    if rel is not None and rel.is_external:
        link["url"] = rel.target
    elif rel is not None and rel.type == REL_SLIDE and resolver.slide_id(rel.target_part):
        link["slide"] = resolver.slide_id(rel.target_part)
    elif action:
        link["action"] = action
    if node.get("tooltip"):
        link["tip"] = node.get("tooltip")
    return link


def run_model(run: Element, resolver: Resolver) -> dict[str, Any]:
    node = run.find(qn("a:t"))
    model: dict[str, Any] = {"t": "" if node is None or node.text is None else node.text}
    if local_name(run) == "fld":
        model["fld"] = run.get("type") or ""
    properties = run.find(qn("a:rPr"))
    if properties is None:
        return model
    for key in ("b", "i"):
        raw = properties.get(key)
        if raw is not None:
            model[key] = raw.strip().lower() in {"1", "true", "on"}
    for key in ("u", "strike"):
        if properties.get(key) is not None:
            model[key] = properties.get(key)
    size = get_int(properties, "sz")
    if size is not None:
        model["sz"] = size
    latin = properties.find(qn("a:latin"))
    if latin is not None and latin.get("typeface") is not None:
        model["latin"] = latin.get("typeface")
    fill = fill_model(properties, resolver)
    if fill["kind"] == "solid" and "color" not in fill:
        model["color"] = color_text(color_from_fields(fill, "run colour"))
    elif fill["kind"] == "solid":
        model["fill"] = f"solid:{fill['color']}"
    elif fill["kind"] != "inherit":
        model["fill"] = fill["kind"]
    click = properties.find(qn("a:hlinkClick"))
    if click is not None:
        model["link"] = _hyperlink(click, resolver)
    return model


def paragraph_model(paragraph: Element, resolver: Resolver) -> dict[str, Any]:
    model: dict[str, Any] = {}
    properties = paragraph.find(qn("a:pPr"))
    if properties is not None:
        if properties.get("algn") is not None:
            model["algn"] = properties.get("algn")
        level = get_int(properties, "lvl", 0) or 0
        if level:
            model["lvl"] = level
        for key in ("marL", "indent"):
            value = get_int(properties, key)
            if value is not None:
                model[key] = value
        for key, tag in (("spcBef", "a:spcBef"), ("spcAft", "a:spcAft"), ("lnSpc", "a:lnSpc")):
            value = _spacing(properties, tag)
            if value is not None:
                model[key] = value
        bullet = _bullet(properties)
        if bullet is not None:
            model["bu"] = bullet
    items: list[dict[str, Any]] = []
    for child in paragraph:
        name = local_name(child)
        if name in {"r", "fld"}:
            items.append(run_model(child, resolver))
        elif name == "br":
            items.append({"br": True})
    model["c"] = items
    return model


def text_model(body: Element, resolver: Resolver) -> dict[str, Any]:
    """``{"p": [{"algn": "ctr", ..., "c": [{"t": "Hello", "b": true}, {"br": true}]}]}``."""
    return {"p": [paragraph_model(p, resolver) for p in body.findall(qn("a:p"))]}


def plain_text(model: dict[str, Any]) -> str:
    """The text a model spells, in :attr:`TextFrame.text`'s form (``\\n``, ``\\v``)."""
    return "\n".join("".join("\v" if "br" in item else item["t"] for item in p["c"])
                     for p in model["p"])


def _canonical_link(link: Any, what: str) -> dict[str, Any]:
    if not isinstance(link, dict) or set(link) - {"url", "slide", "action", "tip"}:
        raise FullStateError(f"{what}: a link is {{url | slide | action, tip}}")
    result: dict[str, Any] = {}
    targets = [key for key in ("url", "slide", "action") if key in link]
    if len(targets) > 1:
        raise FullStateError(f"{what}: a link has one of url, slide or action")
    if "url" in link:
        result["url"] = _text(link["url"], what, 4096)
    elif "slide" in link:
        result["slide"] = parse_int(link["slide"], what, 256, 2147483647)
    elif "action" in link:
        result["action"] = _text(link["action"], what, 1024)
    if "tip" in link:
        result["tip"] = _text(link["tip"], what, 4096)
    return result


def _canonical_run(item: Any, what: str) -> dict[str, Any]:
    if not isinstance(item, dict):
        raise FullStateError(f"{what}: expected an object")
    if "br" in item:
        if item != {"br": True}:
            raise FullStateError(f"{what}: a line break is exactly {{\"br\": true}}")
        return {"br": True}
    unknown = set(item) - set(_RUN_KEYS)
    if unknown or "t" not in item:
        raise FullStateError(f"{what}: a run needs 't' and knows only {list(_RUN_KEYS)}")
    result: dict[str, Any] = {"t": _text(item["t"], what, 1_000_000)}
    if "\n" in result["t"] or "\v" in result["t"]:
        raise FullStateError(f"{what}: a run's text cannot break lines; use {{\"br\": true}}")
    if "fld" in item:
        result["fld"] = _text(item["fld"], what, 256)
    for key in ("b", "i"):
        if key in item:
            if not isinstance(item[key], bool):
                raise FullStateError(f"{what}: {key} must be true or false")
            result[key] = item[key]
    if "u" in item:
        if item["u"] not in UNDERLINES:
            raise FullStateError(f"{what}: unknown underline {item['u']!r}")
        result["u"] = item["u"]
    if "strike" in item:
        result["strike"] = _text(item["strike"], what, 32)
    if "sz" in item:
        result["sz"] = parse_int(item["sz"], f"{what} sz", 100, 400000)
    if "latin" in item:
        result["latin"] = _text(item["latin"], what, 256)
    if "color" in item:
        result["color"] = canonical_color_text(item["color"], what)
    if "fill" in item:
        result["fill"] = _text(item["fill"], what, 64)
    if "link" in item:
        result["link"] = _canonical_link(item["link"], what)
    return result


def _canonical_bullet(bullet: Any, what: str) -> Any:
    if bullet in ("none", "picture"):
        return bullet
    if not isinstance(bullet, dict) or set(bullet) - {"char", "num", "start", "font", "color",
                                                       "size"}:
        raise FullStateError(f"{what}: a bullet is 'none', 'picture' or {{char | num, ...}}")
    result: dict[str, Any] = {}
    if ("char" in bullet) == ("num" in bullet):
        raise FullStateError(f"{what}: a bullet has exactly one of char and num")
    if "char" in bullet:
        result["char"] = _text(bullet["char"], what, 8)
        if not result["char"]:
            raise FullStateError(f"{what}: a bullet character cannot be empty")
    else:
        if bullet["num"] not in NUMBERING_SCHEMES:
            raise FullStateError(f"{what}: unknown numbering scheme {bullet['num']!r}")
        result["num"] = bullet["num"]
        if "start" in bullet:
            result["start"] = parse_int(bullet["start"], what, 1, 32767)
    if "font" in bullet:
        result["font"] = _text(bullet["font"], what, 256)
    if "color" in bullet:
        result["color"] = canonical_color_text(bullet["color"], what)
    if "size" in bullet:
        result["size"] = parse_int(bullet["size"], what, 25000, 400000)
    return result


def _canonical_spacing(value: Any, what: str) -> str:
    unit, separator, raw = _text(value, what, 32).partition(":")
    if not separator or unit not in {"pts", "pct"}:
        raise FullStateError(f"{what}: spacing reads 'pts:<1/100 pt>' or 'pct:<1/1000 %>'")
    return f"{unit}:{parse_int(raw, what, 0, 158400000)}"


def canonical_text(model: Any, what: str) -> dict[str, Any]:
    if not isinstance(model, dict) or set(model) != {"p"} or not isinstance(model["p"], list):
        raise FullStateError(f"{what}: text is {{\"p\": [paragraphs]}}")
    if not model["p"]:
        raise FullStateError(f"{what}: a text body has at least one paragraph")
    paragraphs = []
    for index, paragraph in enumerate(model["p"]):
        where = f"{what} p{index}"
        if not isinstance(paragraph, dict) or set(paragraph) - set(_PARAGRAPH_KEYS):
            raise FullStateError(f"{where}: a paragraph knows only {list(_PARAGRAPH_KEYS)}")
        result: dict[str, Any] = {}
        if "algn" in paragraph:
            if paragraph["algn"] not in _ALIGNMENT_NAMES:
                raise FullStateError(f"{where}: unknown alignment {paragraph['algn']!r}")
            result["algn"] = paragraph["algn"]
        if paragraph.get("lvl"):
            result["lvl"] = parse_int(paragraph["lvl"], where, 0, 8)
        for key in ("marL", "indent"):
            if key in paragraph:
                result[key] = parse_int(paragraph[key], f"{where} {key}", -51206400, 51206400)
        for key in ("spcBef", "spcAft", "lnSpc"):
            if key in paragraph:
                result[key] = _canonical_spacing(paragraph[key], f"{where} {key}")
        if "bu" in paragraph:
            result["bu"] = _canonical_bullet(paragraph["bu"], f"{where} bu")
        items = paragraph.get("c", [])
        if not isinstance(items, list):
            raise FullStateError(f"{where}: 'c' is a list of runs and breaks")
        result["c"] = [_canonical_run(item, f"{where} c{i}") for i, item in enumerate(items)]
        paragraphs.append(result)
    return {"p": paragraphs}


# ------------------------------------------------------------------------------------------
# Tables
# ------------------------------------------------------------------------------------------


def _is_set(tc: Element, attribute: str) -> bool:
    return (tc.get(attribute) or "").strip().lower() in {"1", "true", "on"}


def table_regions(table: Element) -> list[list[int]]:
    """Merged areas as ``[top, left, bottom, right]``, inclusive, in reading order."""
    regions = []
    for r, row in enumerate(table.findall(qn("a:tr"))):
        for c, tc in enumerate(row.findall(qn("a:tc"))):
            if _is_set(tc, "hMerge") or _is_set(tc, "vMerge"):
                continue
            height = max(get_int(tc, "rowSpan", 1) or 1, 1)
            width = max(get_int(tc, "gridSpan", 1) or 1, 1)
            if height > 1 or width > 1:
                regions.append([r, c, r + height - 1, c + width - 1])
    return regions


def cell_model(tc: Element, resolver: Resolver) -> dict[str, Any]:
    model: dict[str, Any] = {}
    body = tc.find(qn("a:txBody"))
    if body is not None:
        model["text"] = text_model(body, resolver)
    properties = tc.find(qn("a:tcPr"))
    fill = fill_model(properties, resolver)
    if fill["kind"] != "inherit":
        model["fill"] = fill
    if properties is not None:
        for side, tag in BORDER_TAGS.items():
            line = properties.find(qn(tag))
            if line is not None:
                model[side] = line_model(line)
    return model


def table_model(table: Element, resolver: Resolver) -> dict[str, Any]:
    grid = table.find(qn("a:tblGrid"))
    columns = [] if grid is None else grid.findall(qn("a:gridCol"))
    rows = table.findall(qn("a:tr"))
    return {
        "cols": [get_int(col, "w", 0) or 0 for col in columns],
        "rows": [get_int(row, "h", 0) or 0 for row in rows],
        "merges": table_regions(table),
        "cells": [[cell_model(tc, resolver) for tc in row.findall(qn("a:tc"))] for row in rows],
    }


def canonical_table(model: Any, what: str) -> dict[str, Any]:
    if not isinstance(model, dict) or set(model) != {"cols", "rows", "merges", "cells"}:
        raise FullStateError(f"{what}: a table is {{cols, rows, merges, cells}}")
    for key in ("cols", "rows", "merges", "cells"):
        if not isinstance(model[key], list):
            raise FullStateError(f"{what}: {key} must be a list")
    columns = [parse_int(w, f"{what} cols", 0, 51206400) for w in model["cols"]]
    rows = [parse_int(h, f"{what} rows", 0, 51206400) for h in model["rows"]]
    if not columns or not rows:
        raise FullStateError(f"{what}: a table has at least one row and one column")
    if len(model["cells"]) != len(rows):
        raise FullStateError(f"{what}: {len(model['cells'])} rows of cells for {len(rows)} rows")
    merges = []
    for region in model["merges"]:
        if not isinstance(region, list) or len(region) != 4:
            raise FullStateError(f"{what}: a merge is [top, left, bottom, right]")
        top, left, bottom, right = (parse_int(v, f"{what} merge", 0, 10000) for v in region)
        if not (top <= bottom < len(rows) and left <= right < len(columns)):
            raise FullStateError(f"{what}: merge {region} is outside the table")
        merges.append([top, left, bottom, right])
    cells = []
    for r, row in enumerate(model["cells"]):
        if not isinstance(row, list) or len(row) != len(columns):
            raise FullStateError(f"{what}: row {r} needs {len(columns)} cells")
        canonical_row = []
        for c, cell in enumerate(row):
            where = f"{what} cell{r},{c}"
            if not isinstance(cell, dict) or set(cell) - ({"text", "fill"} | set(BORDER_TAGS)):
                raise FullStateError(f"{where}: a cell knows text, fill and the border sides")
            result: dict[str, Any] = {}
            if "text" in cell:
                result["text"] = canonical_text(cell["text"], where)
            if "fill" in cell:
                fill = canonical_fill(cell["fill"], where)
                if fill["kind"] != "inherit":
                    result["fill"] = fill
            for side in BORDER_TAGS:
                if side in cell:
                    result[side] = canonical_line(cell[side], f"{where} {side}")
            canonical_row.append(result)
        cells.append(canonical_row)
    return {"cols": columns, "rows": rows, "merges": merges, "cells": cells}


# ------------------------------------------------------------------------------------------
# Shapes
# ------------------------------------------------------------------------------------------


def nv_container(element: Element) -> Element | None:
    tag = {"sp": "p:nvSpPr", "pic": "p:nvPicPr", "cxnSp": "p:nvCxnSpPr", "grpSp": "p:nvGrpSpPr",
           "graphicFrame": "p:nvGraphicFramePr"}.get(local_name(element))
    return None if tag is None else element.find(qn(tag))


def xfrm_of(element: Element) -> Element | None:
    name = local_name(element)
    if name == "graphicFrame":
        return element.find(qn("p:xfrm"))
    container = element.find(qn("p:grpSpPr" if name == "grpSp" else "p:spPr"))
    return None if container is None else container.find(qn("a:xfrm"))


def properties_of(element: Element) -> Element | None:
    """The element holding a shape's fill and outline, if it can have them."""
    name = local_name(element)
    if name == "grpSp":
        return element.find(qn("p:grpSpPr"))
    if name in {"sp", "pic", "cxnSp"}:
        return element.find(qn("p:spPr"))
    return None


#: Features that are written to the document when they change; everything else the emitter
#: writes is read-only, there to describe the shape.  Connections come last: they are
#: applied once every shape is where the SVG puts it.
EDITABLE = ("name", "x", "y", "cx", "cy", "rot", "flip-h", "flip-v", "geom", "adj", "fill",
            "line", "image", "text", "table", "chart-data", "diagram-nodes", "cxn-begin",
            "cxn-end")
CONNECTIONS = ("cxn-begin", "cxn-end")


def shape_attributes(element: Element, resolver: Resolver) -> dict[str, str]:
    """Every typed attribute of one shape, without the prefix, in emission order."""
    from .safe import dump_json

    name = local_name(element)
    attributes: dict[str, str] = {"kind": KINDS[name]}
    nv = nv_container(element)
    properties = None if nv is None else nv.find(qn("p:cNvPr"))
    attributes["name"] = "" if properties is None else properties.get("name", "")
    if properties is not None and properties.get("id") is not None:
        attributes["cnvpr-id"] = properties.get("id")
    placeholder = None if nv is None else nv.find(f"{qn('p:nvPr')}/{qn('p:ph')}")
    if placeholder is not None:
        attributes["ph-type"] = placeholder.get("type") or "obj"
        if placeholder.get("idx") is not None:
            attributes["ph-idx"] = placeholder.get("idx")

    xfrm = xfrm_of(element)
    offset = None if xfrm is None else xfrm.find(qn("a:off"))
    extent = None if xfrm is None else xfrm.find(qn("a:ext"))
    for key, node, attribute in (("x", offset, "x"), ("y", offset, "y"),
                                 ("cx", extent, "cx"), ("cy", extent, "cy")):
        value = get_int(node, attribute)
        attributes[key] = "inherit" if value is None else str(value)
    attributes["rot"] = degrees_text(get_int(xfrm, "rot", 0) or 0)
    for key, attribute in (("flip-h", "flipH"), ("flip-v", "flipV")):
        attributes[key] = "1" if xfrm is not None and xfrm.get(attribute) in {"1", "true"} else "0"
    if name == "grpSp" and xfrm is not None:
        for key, tag, attribute in (("ch-x", "a:chOff", "x"), ("ch-y", "a:chOff", "y"),
                                    ("ch-cx", "a:chExt", "cx"), ("ch-cy", "a:chExt", "cy")):
            value = get_int(xfrm.find(qn(tag)), attribute)
            if value is not None:
                attributes[key] = str(value)

    shape_properties = properties_of(element)
    if name in {"sp", "pic", "cxnSp"}:
        geometry = None if shape_properties is None else shape_properties.find(qn("a:prstGeom"))
        if geometry is not None:
            attributes["geom"] = geometry.get("prst") or ""
            attributes["adj"] = adjustments_text(geometry)
        elif shape_properties is not None and shape_properties.find(qn("a:custGeom")) is not None:
            attributes["geom"] = "custom"
    if name == "cxnSp":
        for key, tag in (("cxn-begin", "a:stCxn"), ("cxn-end", "a:endCxn")):
            attributes[key] = connection_text(element, tag, resolver)
    if name in {"sp", "pic", "cxnSp", "grpSp"}:
        attributes.update(encode_attributes("fill", fill_model(shape_properties, resolver)))
    if name in {"sp", "pic", "cxnSp"}:
        line = None if shape_properties is None else shape_properties.find(qn("a:ln"))
        attributes.update(encode_attributes("line", line_model(line)))
    if name == "pic":
        blip = element.find(f"{qn('p:blipFill')}/{qn('a:blip')}")
        if blip is not None:
            attributes["image"] = resolver.target(blip.get(qn("r:embed")))
    if name == "graphicFrame":
        data = element.find(f"{qn('a:graphic')}/{qn('a:graphicData')}")
        uri = "" if data is None else data.get("uri", "")
        attributes["graphic"] = GRAPHIC_KINDS.get(uri, "other")
        chart = None if data is None else data.find(qn("c:chart"))
        if chart is not None:
            attributes["chart"] = resolver.target(chart.get(qn("r:id")))
            root = resolver.part(attributes["chart"])
            if root is not None:
                attributes["chart-data"] = dump_json(chart_model(root))
        relids = None if data is None else data.find(qn("dgm:relIds"))
        if relids is not None:
            attributes["diagram"] = resolver.target(relids.get(qn("r:dm")))
            root = resolver.part(attributes["diagram"])
            layout = resolver.part(resolver.target(relids.get(qn("r:lo"))))
            if root is not None:
                attributes["diagram-nodes"] = dump_json(diagram_model(
                    root, None if layout is None else layout.get("uniqueId")))
        table = None if data is None else data.find(qn("a:tbl"))
        if table is not None:
            attributes["table"] = dump_json(table_model(table, resolver))
    if name == "sp":
        body = element.find(qn("p:txBody"))
        if body is not None:
            attributes["text"] = dump_json(text_model(body, resolver))
    return attributes


_ADJUST_NAME = re.compile(r"[A-Za-z][A-Za-z0-9]{0,31}")
_LITERAL = re.compile(r"\s*val\s+(-?\d{1,12})\s*")
_SHAPE_ID = re.compile(r"[^\s]{1,256}")


def adjustments_text(geometry: Element) -> str:
    """A preset's explicit adjust values as written: ``"adj1=30000 adj2=70000"``, ``""``
    when it sets none.  Guides that are not plain values are left to the raw XML."""
    holder = geometry.find(qn("a:avLst"))
    values = []
    for node in [] if holder is None else holder.findall(qn("a:gd")):
        match = _LITERAL.fullmatch(node.get("fmla", ""))
        if match and node.get("name"):
            values.append(f"{node.get('name')}={int(match.group(1))}")
    return " ".join(values)


def canonical_adjustments(value: str, what: str, limit: int = 64) -> tuple[tuple[str, int], ...]:
    """``(name, value)`` pairs, each name once, at most ``limit`` of them."""
    pairs: list[tuple[str, int]] = []
    for token in _text(value, f"{what} adj", 4096).split():
        name, separator, raw = token.partition("=")
        if not separator or not _ADJUST_NAME.fullmatch(name):
            raise FullStateError(f"{what} adj: {token!r} should read like 'adj1=30000'")
        if any(name == seen for seen, _ in pairs):
            raise FullStateError(f"{what} adj: {name} is given twice")
        pairs.append((name, parse_int(raw, f"{what} adj {name}", -(2 ** 31), 2 ** 31 - 1)))
        if len(pairs) > limit:
            raise FullStateError(f"{what} adj: more than {limit} adjust values")
    return tuple(pairs)


def effective_adjustments(preset: str | None, pairs, what: str) -> dict[str, int]:
    """Every adjust value of ``preset``: the given ones, the defaults for the rest."""
    from ..edit.presets import adjustment_defaults

    if preset in (None, "", "custom") or preset not in PRESET_GEOMETRIES:
        if pairs:
            raise FullStateError(f"{what} adj: a custom geometry's guides are edited in "
                                 f"data-ooxml-xml")
        return {}
    values = dict(adjustment_defaults(preset))
    for name, value in pairs:
        if name not in values:
            raise FullStateError(f"{what} adj: {preset} has no adjustment {name!r}; it has "
                                 f"{sorted(values)}")
        values[name] = value
    return values


def connection_text(element: Element, tag: str, resolver: Resolver) -> str:
    """``"257.4 3"`` -- the shape a connector end is attached to and the site -- or
    ``"none"`` when the end is free (or names no shape that is uniquely on the slide)."""
    holder = element.find(f"{qn('p:nvCxnSpPr')}/{qn('p:cNvCxnSpPr')}")
    node = None if holder is None else holder.find(qn(tag))
    if node is None:
        return "none"
    target = resolver.shape_id(node.get("id"))
    site = get_int(node, "idx", 0)
    if target is None or site is None or site < 0:
        return "none"
    return f"{target} {site}"


def canonical_connection(value: str, what: str) -> str:
    text = _text(value, what, 300).strip()
    if text == "none":
        return text
    shape, _, site = text.partition(" ")
    if not _SHAPE_ID.fullmatch(shape):
        raise FullStateError(f"{what}: {text!r} should be 'none' or '<shape id> <site>'")
    return f"{shape} {parse_int(site, what, 0, 10_000)}"


def encode_attributes(feature: str, model: dict[str, str]) -> dict[str, str]:
    """``{"kind": "solid", "scheme": "accent1"}`` -> ``{"fill": "solid", "fill-scheme": ...}``."""
    return {(feature if key == "kind" else f"{feature}-{key}"): value
            for key, value in model.items()}


def decode_attributes(feature: str, attributes: dict[str, str]) -> dict[str, str] | None:
    """The inverse of :func:`encode_attributes`; ``None`` when the feature is absent."""
    if feature not in attributes:
        return None
    model = {"kind": attributes[feature]}
    prefix = f"{feature}-"
    for key, value in attributes.items():
        if key.startswith(prefix):
            model[key[len(prefix):]] = value
    return model


def canonical_geometry(key: str, value: str, what: str) -> str:
    if key in {"x", "y", "cx", "cy"}:
        if value == "inherit":
            return value
        low = 0 if key in {"cx", "cy"} else -27273042316900
        return str(parse_int(value, f"{what} {key}", low, 27273042316900))
    if key == "rot":
        return degrees_text(parse_degrees(value, f"{what} rot"))
    if key in {"flip-h", "flip-v"}:
        if value not in {"0", "1"}:
            raise FullStateError(f"{what} {key}: expected 0 or 1")
        return value
    if key == "geom":
        if value != "custom" and value not in PRESET_GEOMETRIES:
            raise FullStateError(f"{what}: {value!r} is not a preset geometry")
        return value
    if key == "name":
        return _text(value, what, 4096)
    raise KeyError(key)


# ------------------------------------------------------------------------------------------
# Charts and diagrams
# ------------------------------------------------------------------------------------------

# The chart and diagram vocabulary is ooxml-edit's (ooxml_edit.charts.model), shared with any
# document that embeds a chart or a diagram; its refusals become this module's, word for word.


def canonical_chart(model: Any, what: str, max_points: int = 1_000_000) -> dict[str, Any]:
    """Validate a chart's JSON: ``{"types", "title", "axis_titles", "legend", "format",
    "categories", "series": [{"name", "values"}]}``, one value per category per series."""
    try:
        return _charts.canonical_chart(model, what, max_points)
    except ChartModelError as error:
        raise FullStateError(str(error)) from None


def canonical_diagram(model: Any, what: str, max_nodes: int = 10_000) -> dict[str, Any]:
    """Validate a diagram's JSON: ``{"layout"?, "nodes": [{"id"?, "lvl", "t"}]}``.  A node
    without an ``id`` is a new one."""
    try:
        return _charts.canonical_diagram(model, what, max_nodes)
    except ChartModelError as error:
        raise FullStateError(str(error)) from None


Canonicaliser = Callable[[Any, str], Any]
