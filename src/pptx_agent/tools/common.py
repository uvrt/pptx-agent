"""What every PowerPoint tool shares: units, addresses, the text spec's schema, errors.

Lengths are points at the tool boundary and EMU inside the library: :func:`emu` and
:func:`pt` convert, at 12,700 EMU per point, to two decimals.  Addresses are the library's
own (``s:256``, ``256.5``, ``256.5/p1/r0``, ``256/notes``), passed through unchanged.
"""

from __future__ import annotations

import functools
from typing import Any, Callable, Iterable, Mapping

from ooxml_edit.tools import Param, ToolError, array, boolean, integer, number, obj, string

from ..edit.labels import LabelError
from ..edit.textspec import ParagraphSpec, RunSpec, TextSpec
from ..edit.units import EMU_PER_POINT, Pt

KIND = "pptx"


# -- units -------------------------------------------------------------------------------------


def emu(points: float) -> int:
    """Points as EMU."""
    return int(round(float(points) * EMU_PER_POINT))


def pt(value: float | None) -> float | None:
    """EMU as points, two decimals."""
    return None if value is None else round(float(value) / EMU_PER_POINT, 2)


def box_pt(bounds) -> dict[str, float] | None:
    if bounds is None:
        return None
    left, top, width, height = bounds
    return {"x": pt(left), "y": pt(top), "w": pt(width), "h": pt(height)}


def box_emu(box: Mapping[str, float]) -> tuple[int, int, int, int]:
    return emu(box["x"]), emu(box["y"]), emu(box["w"]), emu(box["h"])


# -- addresses ---------------------------------------------------------------------------------


def slide_id_of(address: str) -> int | None:
    """The ``sldId`` an address belongs to: ``s:256``, ``256.5``, ``256/notes`` -> 256."""
    text = str(address)
    if text.startswith("s:"):
        text = text[2:]
    head = text.split("/")[0].split(".")[0].split("#")[0]
    return int(head) if head.isdigit() else None


def slide(deck, address: str, *, field: str = "slide"):
    """The slide ``address`` names (``s:256``, or the bare ``256``); ``not_found`` lists
    the slides."""
    text = str(address).strip()
    try:
        return deck.slide(text if text.startswith("s:") else f"s:{text}")
    except (KeyError, IndexError, ValueError):
        raise ToolError("not_found", f"no slide {address!r}", field=field,
                        valid_options=[f"s:{s.slide_id}" for s in deck.slides][:50]) from None


def resolve(deck, address: str, *, field: str = "target"):
    """What ``address`` names (a shape, paragraph, run, cell, notes); ``not_found`` lists
    the shapes on its slide."""
    try:
        return deck.resolve(str(address))
    except (KeyError, IndexError, ValueError) as exc:
        raise ToolError("not_found", f"{address!r}: {_message(exc)}", field=field,
                        valid_options=_shapes_near(deck, address)) from None


def shape(deck, address: str, *, field: str = "target"):
    """The shape ``address`` names; ``not_found`` lists the shapes on its slide."""
    try:
        return deck.shape(str(address))
    except (KeyError, IndexError, ValueError) as exc:
        raise ToolError("not_found", f"no shape {address!r}: {_message(exc)}", field=field,
                        valid_options=_shapes_near(deck, address)) from None


def _shapes_near(deck, address: str) -> list[str]:
    sid = slide_id_of(address)
    try:
        target = deck.slide(f"s:{sid}") if sid is not None else None
    except (KeyError, IndexError, ValueError):
        target = None
    if target is None:
        return [f"s:{s.slide_id}" for s in deck.slides][:50]
    return [f"{s.id} ({s.kind}{': ' + s.name if s.name else ''})" for s in target.shapes][:50]


def _message(exc: BaseException) -> str:
    if isinstance(exc, KeyError) and exc.args:
        return str(exc.args[0])
    return str(exc) or type(exc).__name__


def numbers(deck, slides: Iterable[int] | None, *, field: str = "slides") -> list[int]:
    """1-based slide numbers, checked; all slides when ``None``."""
    count = len(deck.slides)
    if slides is None:
        return list(range(1, count + 1))
    out = []
    for number in slides:
        if not 1 <= number <= count:
            raise ToolError("not_found", f"no slide {number}; the deck has {count}",
                            field=field, valid_options=list(range(1, count + 1))[:50])
        if number not in out:
            out.append(number)
    return out


# -- errors ------------------------------------------------------------------------------------


def library_errors(handler: Callable[..., Any]) -> Callable[..., Any]:
    """Run a handler, turning the library's argument errors into tool errors the model can
    act on: a ``ValueError``/``TypeError`` is ``invalid_arguments`` (the library's message
    says what is allowed), a ``LabelError`` ``label_not_found`` with the labels there are."""

    @functools.wraps(handler)
    def run(call, **arguments):
        try:
            return handler(call, **arguments)
        except ToolError:
            raise
        except LabelError as exc:
            raise ToolError("label_not_found", str(exc),
                            valid_options=exc.candidates[:50]) from None
        except (ValueError, TypeError) as exc:
            raise ToolError("invalid_arguments", _message(exc)) from None
        except (KeyError, IndexError) as exc:
            raise ToolError("not_found", _message(exc)) from None

    return run


def refused(message: str, **details: Any) -> ToolError:
    return ToolError("refused", message, details=details or None)


# -- the text spec, in points ------------------------------------------------------------------

ALIGN = ["left", "center", "right", "justify"]
ANCHOR = ["top", "middle", "bottom"]
AUTOFIT = ["none", "normal", "shape"]
BULLET = ["none", "bullet", "number"]


def run_params() -> dict[str, Param]:
    # Nested properties whose name says it all go without a description: lengths are points.
    return {
        "text": string("\\v breaks the line."),
        "bold": boolean(optional=True),
        "italic": boolean(optional=True),
        "underline": boolean(optional=True),
        "size": number(minimum=1, maximum=400, optional=True),
        "font": string(optional=True),
        "color": string(optional=True),
        "hyperlink": string("URL.", optional=True),
    }


def paragraph_params() -> dict[str, Param]:
    return {
        "runs": array(obj(run_params())),
        "align": string(enum=ALIGN, optional=True),
        "bullet": string(enum=BULLET, optional=True),
        "level": integer("0-8.", minimum=0, maximum=8, optional=True),
        "space_before": number(minimum=0, maximum=1000, optional=True),
        "space_after": number(minimum=0, maximum=1000, optional=True),
        "line_spacing": number("Times single.", minimum=0.1, maximum=10, optional=True),
    }


def insets_param(description: str | None = None) -> Param:
    return obj({side: number(minimum=0, maximum=1000)
                for side in ("left", "top", "right", "bottom")}, description, optional=True)


def frame_params() -> dict[str, Param]:
    return {
        "insets": insets_param("Default 7.2, 3.6, 7.2, 3.6."),
        "anchor": string(enum=ANCHOR, optional=True),
        "wrap": boolean("Default true.", optional=True),
        "autofit": string("Default none.", enum=AUTOFIT, optional=True),
    }


def text_params(described: bool = True) -> dict[str, Param]:
    """``text`` or ``paragraphs`` (exactly one), and ``frame``: the text spec."""
    return {
        "text": string("Plain text; \\n paragraphs, \\v line breaks.", optional=True),
        "paragraphs": array(obj(paragraph_params()), "Text spec, instead of text.",
                            optional=True),
        "frame": obj(frame_params(), "Text frame.", optional=True),
    }


def spec_of(item: Mapping[str, Any], *, field: str = "") -> TextSpec | None:
    """The :class:`TextSpec` an item states: ``paragraphs`` (and ``frame``), or ``None``
    for plain ``text``; both at once is an error."""
    where = f"{field}." if field else ""
    if "text" in item and "paragraphs" in item:
        raise ToolError("invalid_arguments", "give text or paragraphs, not both",
                        field=f"{where}paragraphs", valid_options=["text", "paragraphs"])
    if "paragraphs" not in item:
        return None
    paragraphs = []
    for paragraph in item["paragraphs"]:
        runs = [RunSpec(**run) for run in paragraph.get("runs", [])]
        paragraphs.append(ParagraphSpec(
            runs, align=paragraph.get("align"), bullet=paragraph.get("bullet"),
            level=paragraph.get("level"),
            space_before=_points(paragraph.get("space_before")),
            space_after=_points(paragraph.get("space_after")),
            line_spacing=paragraph.get("line_spacing")))
    return TextSpec(paragraphs, **frame_of(item.get("frame") or {}))


def frame_of(frame: Mapping[str, Any]) -> dict[str, Any]:
    """A ``frame`` argument as :class:`TextSpec`'s frame fields (EMU)."""
    out: dict[str, Any] = {}
    if "insets" in frame:
        insets = frame["insets"]
        out["insets"] = tuple(emu(insets[side]) for side in ("left", "top", "right", "bottom"))
    for key in ("anchor", "wrap", "autofit"):
        if key in frame:
            out[key] = frame[key]
    return out


def apply_frame(text_frame, frame: Mapping[str, Any]) -> None:
    fields = frame_of(frame)
    if "insets" in fields:
        text_frame.insets = fields["insets"]
    if "anchor" in fields:
        text_frame.anchor = fields["anchor"]
    if "wrap" in fields:
        text_frame.wrap = fields["wrap"]
    if "autofit" in fields:
        text_frame.autofit = fields["autofit"]


def _points(value: float | None) -> Pt | None:
    return None if value is None else Pt(value)


def write_text(target, item: Mapping[str, Any], *, field: str = "") -> None:
    """Write an item's ``text`` or ``paragraphs`` (and ``frame``) into a shape, cell or
    notes frame: plain text keeps each surviving character's formatting."""
    spec = spec_of(item, field=field)
    text_frame = target if hasattr(target, "paragraphs") and not hasattr(target, "kind") \
        else getattr(target, "text_frame", None)
    if text_frame is None:
        raise ToolError("invalid_arguments", f"{getattr(target, 'id', target)} holds no text",
                        field=f"{field}.target" if field else "target")
    if spec is not None:
        text_frame.set_text(spec)
        return
    if "text" in item:
        text_frame.set_text(item["text"])
    if item.get("frame"):
        apply_frame(text_frame, item["frame"])


# -- facts -------------------------------------------------------------------------------------


def fit_json(fit) -> dict[str, Any]:
    """A :class:`TextFit` in points, with the allowance stated where it matters."""
    data: dict[str, Any] = {"fits": not fit.overflows, "needed": pt(fit.needed),
                            "available": pt(fit.available), "lines": list(fit.lines)}
    if fit.overflows:
        data["overflow"] = pt(fit.overflow)
    elif fit.overflow > 0:
        data["within_allowance"] = pt(fit.overflow)
        data["allowance"] = pt(fit.slack)
    if fit.near_wrap:
        data["near_wrap"] = pt(fit.margin_to_wrap)
    if fit.autofit != "none":
        data["autofit"] = fit.autofit
    return data


ALLOWANCE_NOTE = ("fits while needed > available: up to 'allowance' pt is the empty bottom of "
                  "the last line, which draws no ink")
NEAR_WRAP_NOTE = ("near_wrap: a line has under 0.16 pt to spare; PowerPoint may break it "
                  "where this measurement does not: give it a few points")

ROWS_NOTE = ("rows_past: a table's rows grow to fit their text and PowerPoint cuts off what "
             "passes the slide's bottom; keep rows_fit rows and move the rest to a table on "
             "the next slide, or shorten the text")


def facts(deck, slide_ids: Iterable[int], *, boxes: bool = False,
          include: Iterable[str] = ("fit", "collisions")) -> dict[str, Any]:
    """Fit and collision facts for some slides, in points, from one measurement each."""
    from ..edit.fit import slide_report

    wanted = set(include)
    overflows: list[dict] = []
    collisions: list[dict] = []
    off_slide: list[dict] = []
    near_wrap: list[dict] = []
    allowance: list[dict] = []
    for sid in slide_ids:
        try:
            target = deck.slide(f"s:{sid}")
        except (KeyError, IndexError, ValueError):
            continue
        problems, fits = slide_report(target, boxes=boxes)
        for problem in problems:
            if problem.kind == "text" and "fit" in wanted:
                overflows.append({"shape": problem.shape, **fit_json(problem.fit)})
            elif problem.kind == "overlap" and "collisions" in wanted:
                entry = {"detail": problem.detail, "shape": problem.shape, "other": problem.other}
                if problem.detail == "line":
                    entry["crossing"] = pt(problem.amount)
                else:
                    entry["area"] = round(problem.amount / EMU_PER_POINT ** 2, 1)
                collisions.append(entry)
            elif problem.kind == "off_slide" and "fit" in wanted:
                entry = {"shape": problem.shape, "past": pt(problem.amount)}
                if problem.rows is not None:
                    entry.update(rows_past=list(problem.rows), rows_fit=problem.rows_fit)
                off_slide.append(entry)
        if "fit" in wanted:
            for address, fit in fits.items():
                if fit.near_wrap:
                    near_wrap.append({"shape": address, "margin": pt(fit.margin_to_wrap)})
                if not fit.overflows and fit.overflow > 0:
                    allowance.append({"shape": address, "over": pt(fit.overflow),
                                      "allowance": pt(fit.slack)})
    out: dict[str, Any] = {}
    if "fit" in wanted:
        out["overflows"] = overflows
        out["off_slide"] = off_slide
        if near_wrap:
            out["near_wrap"] = near_wrap
            out["near_wrap_note"] = NEAR_WRAP_NOTE
        if allowance:
            out["within_allowance"] = allowance
            out["allowance_note"] = ALLOWANCE_NOTE
        if any("rows_past" in entry for entry in off_slide):
            out["rows_note"] = ROWS_NOTE
    if "collisions" in wanted:
        out["collisions"] = collisions
    return out


def validate_delta(entry, problems=None) -> dict[str, Any]:
    deck = entry.document
    problems = deck.validate() if problems is None else problems
    baseline = {str(p) for p in entry.baseline_problems}
    now = {str(p) for p in problems}
    return {"new": [str(p) for p in problems if str(p) not in baseline],
            "fixed": [p for p in sorted(baseline) if p not in now],
            "baseline": len(entry.baseline_problems)}


#: At most this many layout facts in one result.
LAYOUT_LIMIT = 5


def checks(entry, touched) -> dict[str, Any]:
    """The facts after a changing call (``DocumentFormat.checks``): fit and collisions on
    the slides it touched, layout facts about the shapes it touched (when there are any),
    and the validation delta against the document as opened."""
    deck = entry.document
    order = {s.slide_id: n for n, s in enumerate(deck.slides)}
    slide_ids = sorted({sid for sid in touched if isinstance(sid, int) and sid in order},
                       key=order.get)
    out: dict[str, Any] = {"slides": [f"s:{sid}" for sid in slide_ids]}
    out.update(facts(deck, slide_ids))
    layout = layout_json(deck, [t for t in touched if isinstance(t, str)])
    if layout:
        out["layout"] = layout
    out["validate"] = validate_delta(entry)
    return out


def layout_json(deck, shapes: Iterable[str]) -> list[dict[str, str]]:
    """Layout facts about ``shapes`` (addresses), slide by slide, at most
    :data:`LAYOUT_LIMIT`: near-alignment, uneven gaps, outlier text sizes, labels far from
    their markers, each with the change that resolves it (:mod:`pptx_agent.edit.feedback`)."""
    from ..edit.feedback import layout_facts

    by_slide: dict[int, list[str]] = {}
    for address in shapes:
        sid = slide_id_of(address)
        if sid is not None:
            by_slide.setdefault(sid, []).append(address.split("/")[0])
    out: list[dict[str, str]] = []
    for sid, addresses in by_slide.items():
        try:
            slide = deck.slide(f"s:{sid}")
        except (KeyError, IndexError, ValueError):
            continue
        for fact in layout_facts(slide, addresses, limit=LAYOUT_LIMIT - len(out)):
            out.append(fact.to_json())
        if len(out) >= LAYOUT_LIMIT:
            break
    return out


def touch_address(call, address: str) -> None:
    """Say a call changed what ``address`` names: its slide (fit and collisions) and, for a
    shape or its text, the shape (layout facts)."""
    sid = slide_id_of(address)
    if sid is not None:
        call.touch(sid)
        if "." in str(address).split("/")[0]:
            call.touch(str(address).split("/")[0])


def touch_shape(call, shape) -> None:
    """Say a call changed ``shape``: its slide and the shape itself."""
    call.touch(shape._slide.slide_id, shape.id)


_PERCENT = {"lumMod", "lumOff", "tint", "shade", "alpha", "alphaMod", "alphaOff", "satMod",
            "satOff", "sat", "lum", "red", "green", "blue", "redMod", "greenMod", "blueMod",
            "redOff", "greenOff", "blueOff"}


def color_text(color) -> str:
    """A colour as the tools take one: ``accent2 lumMod=75%``, ``#1F4E79``."""
    base = f"#{color.value}" if color.kind == "rgb" else color.value
    parts = [base]
    for name, value in color.transforms:
        if value is None:
            parts.append(name)
        elif name in _PERCENT:
            parts.append(f"{name}={value / 1000:g}%")
        else:
            parts.append(f"{name}={value}")
    return " ".join(parts)


def color_json(color, shape=None) -> dict[str, str] | None:
    if color is None:
        return None
    data = {"color": color_text(color)}
    try:
        data["hex"] = color.resolve(shape) if shape is not None else None
    except Exception:  # noqa: BLE001 -- a colour that does not resolve is reported as is
        data["hex"] = None
    return {k: v for k, v in data.items() if v is not None}


__all__ = ["KIND", "box_emu", "box_pt", "checks", "emu", "facts", "fit_json", "layout_json",
           "library_errors", "pt", "resolve", "shape", "slide", "slide_id_of", "spec_of",
           "text_params", "write_text"]
