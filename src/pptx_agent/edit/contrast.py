"""Text drawn in low contrast with what is behind it, measured: :meth:`Slide.contrast
<pptx_agent.Slide.contrast>`.

Measurements only, lowest first, as ``check``'s spacing is: whether a pale caption is a
problem is the application's to say.  For each shape's text (a group's members included):

* **the colour** each run is drawn in -- :attr:`Run.effective_color
  <pptx_agent.Run.effective_color>`, the chain PowerPoint walks
  (:mod:`ooxml_common.drawingml.text_chain`, the one pptx2svg draws with), with where it
  comes from (:func:`.effective.color_origin`: ``"defaultTextStyle lvl1"``...);
* **what is behind it** -- the shape's own fill as drawn; where it has none, the shapes
  beneath it whose drawn box holds the centre of its box, topmost first; then the slide's
  background (the slide's, else its layout's, else its master's).  Fills and backgrounds as
  pptx2svg resolves them.  A fill with alpha is blended over what is under it.  A gradient
  counts as its **worst stop** (the one of least contrast with the text), a pattern as the
  worse of its two colours.  A picture, a picture fill, a table, a chart or a SmartArt
  diagram behind the text, or a background picture, makes it ``"unknown"``: not sampled
  (sampling the render would cost a raster per slide);
* **the ratio** -- WCAG 2's contrast ratio, ``(L1 + 0.05) / (L2 + 0.05)`` of the two
  colours' relative luminances (sRGB, linearised), 1 to 21.

A run is reported when its ratio is below the threshold: ``minimum`` (default 4.5:1, WCAG's
for text), and for **large text** -- 18 pt and up, or 14 pt and up in bold, WCAG's
definition, at the size drawn (autofit applied) -- two thirds of it (3:1 at the default).
One entry per shape, its worst run; text behind which nothing can be told is listed after
the measured ones, without a ratio.  Left out: table cells, chart and SmartArt text, and a
text colour's own transparency (the colour is taken as opaque).

For example::

    for entry in slide.contrast().entries:
        print(entry.to_json())
    # {'shape': '258.7', 'text': 'Hello', 'color': '#FFFFFF', 'behind': '#FFFFFF',
    #  'ratio': 1.0, 'source': 'defaultTextStyle lvl1', 'over': 'background'}
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover
    from .document import Slide

#: WCAG 2's contrast for text, and for large text the share of it its 3:1 is of 4.5:1.
DEFAULT_MINIMUM = 4.5
LARGE_SHARE = 2 / 3
#: WCAG's large text: 18 pt, or 14 pt in bold.
LARGE_SIZE, LARGE_BOLD_SIZE = 18.0, 14.0
#: Words of the shape's text an entry quotes.
WORDS = 8
UNKNOWN = "unknown"


@dataclass(frozen=True)
class Contrast:
    """One shape's text at its least contrast: ``color`` and ``behind`` ``#RRGGBB``
    (``behind`` ``"unknown"`` and ``ratio`` ``None`` where a picture or the like is behind
    it), ``over`` the topmost layer behind the text (``"fill"``, the address of a shape
    beneath, ``"background"``; blended with those under it when translucent), ``source`` where the colour comes from, ``large`` WCAG's large
    text (held to two thirds of the threshold)."""

    shape: str
    text: str
    color: str
    behind: str
    ratio: float | None
    source: str
    over: str
    large: bool = False

    def to_json(self) -> dict[str, Any]:
        data: dict[str, Any] = {"shape": self.shape, "text": self.text, "color": self.color,
                                "behind": self.behind}
        if self.ratio is not None:
            data["ratio"] = self.ratio
        data["source"] = self.source
        data["over"] = self.over
        if self.large:
            data["large"] = True
        return data


@dataclass(frozen=True)
class ContrastReport:
    """:meth:`Slide.contrast <pptx_agent.Slide.contrast>`: the text below the threshold,
    lowest ratio first, then the text whose background is unknown (``entries``)."""

    entries: tuple[Contrast, ...]
    minimum: float


def luminance(hex_color: str) -> float:
    """WCAG 2's relative luminance of ``#RRGGBB``."""
    value = hex_color.lstrip("#")
    channels = []
    for index in (0, 2, 4):
        c = int(value[index:index + 2], 16) / 255
        channels.append(c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4)
    red, green, blue = channels
    return 0.2126 * red + 0.7152 * green + 0.0722 * blue


def ratio(first: str, second: str) -> float:
    """WCAG 2's contrast ratio of two ``#RRGGBB`` colours, 1 to 21."""
    a, b = luminance(first), luminance(second)
    lighter, darker = max(a, b), min(a, b)
    return (lighter + 0.05) / (darker + 0.05)


def _blend(top: str, alpha: float, under: str) -> str:
    t, u = top.lstrip("#"), under.lstrip("#")
    out = []
    for index in (0, 2, 4):
        a, b = int(t[index:index + 2], 16), int(u[index:index + 2], 16)
        out.append(round(a * alpha + b * (1 - alpha)))
    return "#%02X%02X%02X" % tuple(out)


def _candidates(fill) -> list[tuple[str, float]] | None:
    """A resolved fill's colours as ``(hex, alpha)`` -- one, a gradient's stops, a
    pattern's two -- ``[]`` for no fill, ``None`` for one that cannot be told (a picture)."""
    if fill is None:
        return []
    kind = getattr(fill, "type", None)
    if kind == "none":
        return []
    if kind == "solid":
        return [(fill.color.hex.upper(), fill.color.alpha)]
    if kind == "gradient":
        return [(stop.color.hex.upper(), stop.color.alpha) for stop in fill.stops] or []
    if kind == "pattern":
        return [(fill.foreground_color.hex.upper(), fill.foreground_color.alpha),
                (fill.background_color.hex.upper(), fill.background_color.alpha)]
    return None


def _contains(box, x: float, y: float) -> bool:
    return box[0] <= x <= box[0] + box[2] and box[1] <= y <= box[1] + box[3]


def _bold(host, paragraph, run) -> bool:
    from .effective import _run_properties

    for properties in _run_properties(host, paragraph, run):
        value = properties.get("b")
        if value is not None:
            return value in ("1", "true", "on")
    return False


class _Layers:
    """The slide's resolved elements by id, and its background, from pptx2svg."""

    def __init__(self, slide: "Slide") -> None:
        import pptx2svg

        options = pptx2svg.ConvertOptions(slide_numbers=[slide.index + 1],
                                          warn_on_font_substitution=False)
        model = pptx2svg.convert_pptx_to_model(slide.document.to_bytes(), options)
        resolved = model.slides[0]
        self.by_id: dict[str, Any] = {}
        self._walk(resolved.elements)
        background = resolved.background.fill if resolved.background is not None else None
        found = _candidates(background)
        #: The background's colours; white when nothing gives one, as PowerPoint draws it.
        self.background = found if found else ([("#FFFFFF", 1.0)] if found == [] else None)

    def _walk(self, elements) -> None:
        for element in elements:
            self.by_id.setdefault(element.element_id or "", element)
            if getattr(element, "type", None) == "group":
                self._walk(element.children)

    def element(self, shape):
        from .fit import _raw_id

        return self.by_id.get(f"{shape._slide.slide_id}.{_raw_id(shape)}")


def _behind(shape, beneath: list, layers: _Layers):
    """``(colours, over)``: what is behind ``shape``'s text as opaque candidate colours (a
    gradient's stops...), and the topmost layer that gives them; ``(None, over)`` when it
    cannot be told."""
    stack: list[tuple[list[tuple[str, float]], str]] = []
    element = layers.element(shape)
    own = _candidates(getattr(element, "fill", None)) if element is not None else []
    if own is None:
        return None, "fill"
    if own:
        stack.append((own, "fill"))
    if not own or min(alpha for _, alpha in own) < 1:
        box = shape.drawn_bounds
        if box is not None:
            x, y = box[0] + box[2] / 2, box[1] + box[3] / 2
            for other in reversed(beneath):
                bounds = other.drawn_bounds
                if bounds is None or not _contains(bounds, x, y):
                    continue
                if other.kind in ("picture", "graphic_frame"):
                    return None, other.id
                if other.kind != "shape":
                    continue
                found = layers.element(other)
                colours = _candidates(getattr(found, "fill", None)) if found is not None else []
                if colours is None:
                    return None, other.id
                if colours:
                    stack.append((colours, other.id))
                    if min(alpha for _, alpha in colours) >= 1:
                        break
    if not stack or min(alpha for _, alpha in stack[-1][0]) < 1:
        if layers.background is None:
            return None, "background"
        stack.append((layers.background, "background"))
    # Blend from the bottom up; every combination of candidates is a colour that may be drawn.
    colours = [hex_color for hex_color, _ in stack[-1][0]]
    for layer, _ in reversed(stack[:-1]):
        colours = [_blend(top, alpha, under) if alpha < 1 else top
                   for top, alpha in layer for under in colours]
    # What decided it: the topmost layer (blended with those under it when translucent).
    return sorted(set(colours)), stack[0][1]


def contrast(slide: "Slide", *, minimum: float = DEFAULT_MINIMUM) -> ContrastReport:
    from .effective import color_origin
    from .fit import _flatten

    layers = _Layers(slide)
    shapes = _flatten(slide.shapes)
    measured: list[Contrast] = []
    unknown: list[Contrast] = []
    for index, shape in enumerate(shapes):
        if shape.kind != "shape":
            continue
        frame = shape.text_frame
        if frame is None or not frame.text.strip():
            continue
        words = frame.text.split()
        text = " ".join(words[:WORDS]) + ("…" if len(words) > WORDS else "")
        behind, over = _behind(shape, shapes[:index], layers)
        worst = None
        for paragraph in frame.paragraphs:
            host, element = paragraph._frame._host(), paragraph._element()
            for run in paragraph.runs:
                if not run.text.strip():
                    continue
                drawn = run.effective_color
                if drawn is None:
                    continue
                color = f"#{drawn.value.upper()}"
                size = run.effective_size
                large = size >= LARGE_SIZE or (size >= LARGE_BOLD_SIZE
                                               and _bold(host, element, run._element()))
                if behind is None:
                    key = (0.0, 0)
                    found = (color, UNKNOWN, None, large)
                else:
                    value, against = min((ratio(color, under), under) for under in behind)
                    threshold = minimum * LARGE_SHARE if large else minimum
                    if value >= threshold:
                        continue
                    key = (value / threshold, 1)
                    found = (color, against, round(value, 2), large)
                if worst is None or key < worst[0]:
                    worst = (key, found, color_origin(host, element, run._element()))
        if worst is None:
            continue
        (color, against, value, large), source = worst[1], worst[2]
        entry = Contrast(shape.id, text, color, against, value, source, over, large)
        (unknown if value is None else measured).append(entry)
    measured.sort(key=lambda entry: entry.ratio)
    return ContrastReport(tuple(measured + unknown), minimum)


__all__ = ["DEFAULT_MINIMUM", "Contrast", "ContrastReport", "contrast", "luminance", "ratio"]
