"""Shape tools: ``ppt_add_shape``, ``ppt_set_shape``, ``ppt_add_picture``,
``ppt_add_connector`` and ``ppt_arrange``.

The plural ones take ``items[]``: a whole graphic -- forty shapes, their connectors -- in
one call and one undo step, each item optionally naming what it makes (``ref``) so later
items and calls can write ``$name`` for its address.
"""

from __future__ import annotations

from typing import Any

from ooxml_edit.tools import (CORE, Result, ToolError, array, boolean, integer, number, obj,
                              string, tool)

from ..edit.fill import ARROWHEADS

#: The dash styles offered (the library takes all of ECMA's eleven).
DASHES = ["solid", "dash", "dot", "dashDot", "lgDash", "sysDash", "sysDot"]
from .common import (ANCHOR, AUTOFIT, KIND, box_emu, box_pt, emu, frame_of, insets_param,
                     library_errors, place, pt, resolve, shape as shape_at, host as slide_at,
                     one_text, spec_of, text_params, write_text)
from . import textfacts
from .text import COMMON_PRESETS, preset_of

from ..edit.arrange import EDGES

#: ppt_arrange's align ``to`` and distribute ``within``, as the library's ``to``.
_ALIGN_TO = {"shapes": "selection", "slide": "slide", "content_area": "box"}
_WITHIN = {"shapes": "selection", "content_area": "box"}

GRAPHICS = "ppt_graphics"
OBJECTS = "ppt_objects"

_REF = string("Ref name (a-z, 0-9, _): later arguments write $name.", optional=True)
_KEY = string("Retry key: a repeat with the same key makes nothing new.", optional=True)


def _box(description: str, *, optional: bool = False):
    return obj({"x": number(minimum=-10000, maximum=10000),
                "y": number(minimum=-10000, maximum=10000),
                "w": number(minimum=0, maximum=10000),
                "h": number(minimum=0, maximum=10000)},
               description, optional=optional)


def line_param(description: str | None = None):
    return obj({"color": string("none hides it.", optional=True),
                "width": number(minimum=0, maximum=1000, optional=True),
                "dash": string(enum=DASHES, optional=True),
                "start": string("Arrowhead.", enum=sorted(ARROWHEADS), optional=True),
                "end": string("Arrowhead.", enum=sorted(ARROWHEADS), optional=True)},
               description, optional=True)


def _adjustments_param():
    return array(obj({"name": string("adj, adj1..."),
                      "value": number("1/100000 of the shorter side.",
                                      minimum=-1000000, maximum=1000000)}),
                 "Preset adjustments.", optional=True)


def apply_line(target, line: dict[str, Any]) -> None:
    outline = target.line
    if outline is None:
        raise ToolError("invalid_arguments", f"{target.id} has no outline", field="line")
    color = line.get("color")
    if color is not None:
        if color.strip().lower() == "none":
            outline.visible = False
        else:
            outline.color = color
    if "width" in line:
        outline.width = emu(line["width"])
    if "dash" in line:
        outline.dash = line["dash"]
    if "start" in line:
        outline.start = line["start"]
    if "end" in line:
        outline.end = line["end"]


def _adjustments(item: dict[str, Any]) -> dict[str, int]:
    return {entry["name"]: int(round(entry["value"])) for entry in item.get("adjustments") or ()}


def _preset_name(name: str, field: str) -> str:
    from ..edit.presets import PRESETS

    if name not in PRESETS:
        raise ToolError("invalid_arguments", f"{name!r} is not a preset geometry", field=field,
                        valid_options=sorted(PRESETS))
    return name


def _define(call, item: dict[str, Any], address: str) -> None:
    if item.get("ref"):
        call.define_ref(item["ref"], address)


def _measuring(arguments) -> bool:
    return bool(arguments.get("measure"))


@tool("ppt_add_shape",
      "Add autoshapes or text boxes, each with text: plain text, or paragraphs of "
      "formatted runs (the text spec) and a frame; give text or paragraphs, not both. "
      "Position by box, in points. measure: true measures text before building and adds nothing: lines, text height and the box_height to use, insets included.",
      {"doc": string("Document id."),
       "slide": string("Slide, s:256 or $ref."),
       "items": array(obj({
           "preset": string("textbox: a plain text box; other: see preset_name.",
                            enum=COMMON_PRESETS),
           "preset_name": string("Any preset, when preset is other.", optional=True),
           "box": _box("Points. Measuring: w is the width; h, if over 0, is checked."),
           **text_params(described=True),
           "fit_height": boolean("Set h to what the text needs. Default false.",
                                 optional=True),
           "fill": string("Colour or none. Default the theme's.", optional=True),
           "line": line_param(),
           "adjustments": _adjustments_param(),
           "like": string("Measuring: measure as this shape (font, insets, placeholder) "
                          "instead of preset.", optional=True),
           "name": string(optional=True),
           "ref": _REF}), "The shapes, back to front."),
       "measure": boolean("Measure the items' text only; add nothing. Default false.",
                          optional=True),
       "key": _KEY},
      kind=KIND, group=GRAPHICS, mutates=True, refs=("slide", "items[].like"),
      facts=("items[].text", "items[].paragraphs[].runs[].text"), reads=_measuring)
@library_errors
def ppt_add_shape(call, doc, slide, items, measure=False, key=None):
    from ..edit.fit import _converged_height

    target = slide_at(call.document, slide)
    if not items:       # a trimmed batch's leftover: nothing to do, said, not refused
        return Result.nothing_given("measured" if measure else "added")
    if measure:
        return _measure(call, target, items)
    made = []
    fits = {}
    for index, item in enumerate(items):
        field = f"items[{index}]"
        item = one_text(call, item, field=field)
        preset = preset_of(item)
        x, y, w, h = box_emu(item["box"])
        if preset == "textbox":
            shape = target.add_textbox(x, y, w, h, name=item.get("name"), autofit="none")
            if "adjustments" in item:
                raise ToolError("invalid_arguments", "a text box has no adjustments",
                                field=f"{field}.adjustments")
        else:
            shape = target.add_shape(preset, x, y, w, h, name=item.get("name"),
                                     adjustments=_adjustments(item) or None)
        if "text" in item or "paragraphs" in item or item.get("frame"):
            write_text(shape, item, field=field)
            textfacts.Spot(shape).written("", textfacts.placed(call, field, item))
        if "fill" in item:
            shape.fill = item["fill"]
        if item.get("line"):
            apply_line(shape, item["line"])
        if item.get("fit_height"):
            _converged_height(shape)
        made.append(shape.id)
        _define(call, item, shape.id)
        if item.get("fit_height"):
            fits[shape.id] = {"h": pt(shape.height)}
    call.touch(target.slide_id, *made)
    return Result(summary=f"Added {len(made)} shape(s) to {place(target)}", created=made,
                  data={"heights": fits} if fits else None)


def _measure(call, where, items) -> Result:
    """``ppt_add_shape`` with ``measure``: each item's text laid out as it would be built
    (one measuring model: the spec is built on a copy of the deck and measured there)."""
    from ..edit.fit import measure_text
    from ..edit.textspec import TextSpec

    deck = call.document
    measured_items = []
    for index, item in enumerate(items):
        field = f"items[{index}]"
        item = one_text(call, item, field=field)
        spec = spec_of(item, field=field)
        if spec is None:
            if "text" not in item:
                raise ToolError("invalid_arguments", "measuring needs text or paragraphs",
                                field=f"{field}.text", valid_options=["text", "paragraphs"])
            spec = TextSpec.from_text(item["text"])
            if item.get("frame"):
                spec = TextSpec(spec.paragraphs, **frame_of(item["frame"]))
        box = item["box"]
        if box["w"] <= 0:
            raise ToolError("invalid_arguments", "measuring needs a width over 0",
                            field=f"{field}.box.w")
        if item.get("like") is not None:
            like = resolve(deck, item["like"], field=f"{field}.like")
            measured = measure_text(spec, width=emu(box["w"]), like=like)
        else:
            measured = measure_text(spec, width=emu(box["w"]), preset=preset_of(item),
                                    deck_or_shape=where)
        data = {"lines": list(measured.lines),
                "paragraph_lines": list(measured.paragraph_lines),
                "text_height": pt(measured.height), "box_height": pt(measured.box_height),
                "widest": pt(measured.widest),
                "insets": [pt(v) for v in measured.insets or ()]}
        if measured.margin_to_wrap is not None:
            data["margin_to_wrap"] = pt(measured.margin_to_wrap)
            if measured.near_wrap:
                data["near_wrap"] = True
        if box["h"] > 0:
            data["fits"] = pt(measured.box_height) <= box["h"]
        measured_items.append(data)
    heights = ", ".join(f"{data['box_height']}" for data in measured_items[:10])
    return Result(summary=f"Measured {len(items)} item(s), nothing added; box heights "
                  f"{heights} pt", data={"measured": measured_items})


@tool("ppt_set_shape",
      "Change shapes' position, size, rotation, flips, fill, gradient, outline, geometry "
      "or name; only the fields given change. Text and frames: ppt_set_text, "
      "ppt_format_text.",
      {"doc": string("Document id."),
       "items": array(obj({
           "target": string("Shape, e.g. 256.5 or $ref."),
           "x": number(minimum=-10000, maximum=10000, optional=True),
           "y": number(minimum=-10000, maximum=10000, optional=True),
           "w": number(minimum=0, maximum=10000, optional=True),
           "h": number(minimum=0, maximum=10000, optional=True),
           "rotation": number("Degrees clockwise.", minimum=-360, maximum=360, optional=True),
           "flip_h": boolean(optional=True),
           "flip_v": boolean(optional=True),
           "fill": string("Colour, none or inherit.", optional=True),
           "gradient": obj({"stops": array(obj({
               "position": number("0-100 %.", minimum=0, maximum=100),
               "color": string()})),
               "angle": number("Degrees. Default 90.", minimum=0, maximum=360,
                               optional=True)}, optional=True),
           "line": line_param(),
           "preset": string("New geometry, e.g. roundRect.", optional=True),
           "adjustments": _adjustments_param(),
           "name": string(optional=True)}),
           "Shapes and their changes.")},
      kind=KIND, group=CORE, mutates=True, refs=("items[].target",))
@library_errors
def ppt_set_shape(call, doc, items):
    if not items:
        return Result.nothing_given("changed")
    deck = call.document
    changed = []
    for index, item in enumerate(items):
        field = f"items[{index}]"
        shape = shape_at(deck, item["target"], field=f"{field}.target")
        if "x" in item or "y" in item:
            bounds = shape.slide_bounds
            if bounds is None:
                raise ToolError("refused", f"{shape.id} has no position to move",
                                field=f"{field}.x")
            dx = emu(item["x"]) - bounds[0] if "x" in item else 0
            dy = emu(item["y"]) - bounds[1] if "y" in item else 0
            if dx or dy:
                shape.move_by(dx, dy, space="slide")
        if "w" in item:
            shape.width = emu(item["w"])
        if "h" in item:
            shape.height = emu(item["h"])
        if "rotation" in item:
            shape.rotation = item["rotation"]
        if "flip_h" in item:
            shape.flip_h = item["flip_h"]
        if "flip_v" in item:
            shape.flip_v = item["flip_v"]
        if "preset" in item:
            shape.preset = _preset_name(item["preset"], f"{field}.preset")
        if "adjustments" in item:
            shape.adjustments.set(**_adjustments(item))
        if "fill" in item and "gradient" in item:
            raise ToolError("invalid_arguments", "give fill or gradient, not both",
                            field=f"{field}.gradient")
        graphic = _graphic(shape, item, field)
        if "fill" in item and "fill" not in graphic:
            shape.fill = None if item["fill"] == "inherit" else item["fill"]
        if "gradient" in item:
            gradient = item["gradient"]
            shape.set_gradient_fill([(stop["position"] / 100, stop["color"])
                                     for stop in gradient["stops"]],
                                    angle=gradient.get("angle", 90.0))
        if graphic:
            _recolor(shape, graphic, field)
        if item.get("line") and not {"line", "line_width"} & graphic.keys():
            apply_line(shape, item["line"])       # an SVG picture's border too, as before
        if "name" in item:
            shape.name = item["name"]
        call.touch(shape._slide.slide_id, shape.id)
        changed.append(shape.id)
    shapes = {address: box_pt(deck.shape(address).slide_bounds) for address in changed[:20]}
    return Result(summary=f"Changed {len(changed)} shape(s)", changed=changed,
                  data={"boxes": shapes})


def _graphic(shape, item: dict[str, Any], field: str) -> dict[str, Any]:
    """What of ``fill`` and ``line`` recolours an SVG picture, as PowerPoint's Graphics
    Fill and Graphics Outline do (:meth:`Shape.recolor <pptx_agent.Shape.recolor>`): a
    fill colour, a line colour and width.  ``none`` and ``inherit``, a bitmap picture's
    fill and line (its background and border, as in PowerPoint: only a graphic has
    Graphics Fill and Outline), and a line of only a dash or arrowheads are left to the
    picture's own fill and line, as before."""
    if shape.kind != "picture" or shape.vector is None:
        return {}
    graphic: dict[str, Any] = {}
    fill = item.get("fill")
    if fill is not None and fill.strip().lower() not in ("none", "inherit"):
        graphic["fill"] = fill
    line = item.get("line") or {}
    color = line.get("color")
    if color is not None and color.strip().lower() != "none":
        graphic["line"] = color
    if "width" in line:
        graphic["line_width"] = emu(line["width"])
    if ("line" in graphic or "line_width" in graphic) and \
            any(key in line for key in ("dash", "start", "end")):
        raise ToolError("invalid_arguments", "an SVG picture's outline takes color and "
                        "width only (its shapes' strokes); give dash or arrowheads in a "
                        "separate change for the picture's border", field=f"{field}.line")
    return graphic


def _recolor(picture, graphic: dict[str, Any], field: str) -> None:
    """Recolour an SVG picture; its errors are the call's, at the field they came from."""
    try:
        picture.recolor(graphic.get("fill"), line=graphic.get("line"),
                        line_width=graphic.get("line_width"))
    except ImportError as error:          # the PNG fallback needs pptx-agent[png]
        raise ToolError("refused", str(error), field=field) from None
    except ValueError as error:
        raise ToolError("invalid_arguments", str(error),
                        field=f"{field}.fill" if "fill" in graphic else f"{field}.line") \
            from None


#: The tool that finds a catalogue's assets, when the application configured one.
FIND_ASSET = "ppt_find_asset"


def picture_tool(catalogue=None):
    """``ppt_add_picture``; with an application's asset catalogue
    (:func:`pptx_agent.tools.deck_format`'s ``assets``), the same tool taking ``asset``, an
    id :data:`FIND_ASSET` found, in place of ``image``."""
    params = {"doc": string("Document id."),
              "image": string("Image blob handle, e.g. b2.",
                              optional=catalogue is not None)}
    if catalogue is not None:
        params["asset"] = string(f"Or an asset id from {FIND_ASSET}.", optional=True)
    params.update({
        "slide": string("Slide to insert on, s:256 or $ref.", optional=True),
        "target": string("Or the picture to replace or placeholder to fill.",
                         optional=True),
        "box": _box("Default: native size at the content area's top left.", optional=True),
        "keep": string("Replacing: what of the frame stays. Default frame.",
                       enum=["frame", "height", "width", "none"], optional=True),
        "anchor": string("Replacing: the point that stays put. Default top_left.",
                         enum=["top_left", "top", "top_right", "left", "center", "right",
                               "bottom_left", "bottom", "bottom_right"], optional=True),
        "ref": _REF, "key": _KEY})
    exactly_one = [("slide", "target")]
    if catalogue is not None:
        exactly_one.append(("image", "asset"))

    @tool("ppt_add_picture",
          "Insert a picture from an image blob"
          + (" or catalogue asset" if catalogue is not None else "")
          + " on a slide, or replace a picture's image (target), keeping its frame, height "
          "or width. An SVG stays vector, with a PNG fallback made for it; ppt_set_shape "
          "fill and line recolour it. Returns its address and native size.",
          params, kind=KIND, group=OBJECTS, mutates=True, exactly_one=exactly_one,
          refs=("slide", "target"))
    @library_errors
    def ppt_add_picture(call, doc, image=None, slide=None, target=None, box=None, keep=None,
                        anchor=None, ref=None, key=None, asset=None):
        deck = call.document
        if asset is not None:
            from ooxml_edit.tools import assets

            data, _ = assets.fetch(catalogue, asset, call.session.limits,
                                   search_tool=FIND_ASSET)
        else:
            blob = call.blob(image)
            if not blob.mime.startswith("image/"):
                raise ToolError("invalid_arguments", f"{image} is {blob.mime}, not an image",
                                field="image")
            data = blob.data
        try:
            if target is not None:
                picture = shape_at(deck, target)
                holder = picture.placeholder
                if picture.kind == "shape" and holder is not None and holder.type == "pic":
                    filled = picture.insert_picture(data)
                    address, verb = filled.id, "Filled placeholder with"
                    picture = filled
                elif picture.kind != "picture":
                    raise ToolError("invalid_arguments",
                                    f"{target} is a {picture.kind}, not a picture",
                                    field="target")
                else:
                    picture.replace_image(data, keep=keep or "frame",
                                          anchor=anchor or "top_left")
                    address, verb = picture.id, "Replaced the image of"
            else:
                host = slide_at(deck, slide)
                if box is not None:
                    x, y, w, h = box_emu(box)
                    picture = host.add_picture(data, x, y, width=w, height=h)
                else:
                    left, top, _, _ = host.content_area
                    picture = host.add_picture(data, left, top)
                address, verb = picture.id, "Inserted"
        except ImportError as error:      # an SVG's PNG fallback needs pptx-agent[png]
            raise ToolError("refused", str(error),
                            field="asset" if asset is not None else "image") from None
        if ref:
            call.define_ref(ref, address)
        call.touch(picture._slide.slide_id, picture.id)
        vector = picture.vector
        if vector is not None:
            from ..edit import svgpicture

            native = svgpicture.native_size(svgpicture.parse(vector))
            size = [round(native[0]), round(native[1])] if native else None
        else:
            stored = picture.image_size
            size = [stored.width, stored.height] if stored else None
        data = {"address": address, "box": box_pt(picture.bounds), "native_px": size}
        if vector is not None:
            data["svg"] = True
        return Result(summary=f"{verb} picture {address}",
                      created=[address] if target is None else [],
                      changed=[address] if target is not None else [], data=data)

    return ppt_add_picture


ppt_add_picture = picture_tool()


def find_asset_tool(catalogue):
    """``ppt_find_asset``: the application's asset catalogue searched by words; offered
    only with one (:func:`pptx_agent.tools.deck_format`'s ``assets``)."""
    from ooxml_edit.tools import assets

    @tool(FIND_ASSET,
          "Search the brand asset catalogue (icons, logos) by words. Returns ids, names "
          "and keywords; place one with ppt_add_picture asset.",
          {"query": string("Words, e.g. arrow up."),
           "limit": integer("Default 10.", minimum=1, maximum=assets.MAX_RESULTS,
                            optional=True)},
          group=OBJECTS, documents=())
    def ppt_find_asset(call, query, limit=None):
        found = assets.find(catalogue, query, limit or 10)
        names = ", ".join(item["id"] for item in found[:10])
        return Result(summary=f"{len(found)} asset(s)" + (f": {names}" if names else
                                                          "; try other words"),
                      data={"assets": found})

    return ppt_find_asset


_SIDES = ["auto", "top", "right", "bottom", "left"]


def _end_param(description: str):
    return obj({"shape": string("Shape to attach to, or $ref.", optional=True),
                "side": string("Default auto: facing the other end.", enum=_SIDES,
                               optional=True),
                "x": number("Or a point: x, y.", minimum=-10000, maximum=10000,
                            optional=True),
                "y": number(minimum=-10000, maximum=10000, optional=True)}, description)


def _auto_sides(first, second) -> tuple[str, str]:
    """Sides that face each other: bottom to top when the second is wholly below, top to
    bottom wholly above, else side to side."""
    a, b = first.drawn_bounds, second.drawn_bounds
    a_bottom, b_bottom = a[1] + a[3], b[1] + b[3]
    if b[1] >= a_bottom:
        return "bottom", "top"
    if b_bottom <= a[1]:
        return "top", "bottom"
    if b[0] >= a[0] + a[2]:
        return "right", "left"
    if b[0] + b[2] <= a[0]:
        return "left", "right"
    return ("right", "left") if b[0] + b[2] / 2 >= a[0] + a[2] / 2 else ("left", "right")


def _end(call, spec: dict[str, Any], field: str):
    deck = call.document
    if "shape" in spec:
        if "x" in spec or "y" in spec:
            raise ToolError("invalid_arguments", "an end is a shape or a point, not both",
                            field=f"{field}.x")
        return shape_at(deck, spec["shape"], field=f"{field}.shape")
    if "x" not in spec or "y" not in spec:
        raise ToolError("invalid_arguments", "an end needs a shape, or both x and y",
                        field=f"{field}.shape")
    return (emu(spec["x"]), emu(spec["y"]))


@tool("ppt_add_connector",
      "Add connectors: straight, elbow or curved lines that stay attached to shapes, or "
      "run between points.",
      {"doc": string("Document id."),
       "slide": string("Slide, s:256 or $ref; needed when an end is a point.",
                       optional=True),
       "items": array(obj({
           "kind": string(enum=["straight", "elbow", "curved"]),
           "from": _end_param("A shape (and side) or a point."),
           "to": _end_param("A shape (and side) or a point."),
           "line": line_param("end=triangle draws an arrow."),
           "name": string(optional=True),
           "ref": _REF}), "The connectors."),
       "key": _KEY},
      kind=KIND, group=GRAPHICS, mutates=True,
      refs=("slide", "items[].from.shape", "items[].to.shape"))
@library_errors
def ppt_add_connector(call, doc, items, slide=None, key=None):
    if not items:
        return Result.nothing_given("added")
    deck = call.document
    made = []
    for index, item in enumerate(items):
        field = f"items[{index}]"
        begin = _end(call, item["from"], f"{field}.from")
        finish = _end(call, item["to"], f"{field}.to")
        shapes = [end for end in (begin, finish) if not isinstance(end, tuple)]
        if shapes:
            host = shapes[0]._slide
            if any(other._slide.slide_id != host.slide_id for other in shapes):
                raise ToolError("invalid_arguments", "both ends must be on one slide",
                                field=f"{field}.to.shape")
        elif slide is not None:
            host = slide_at(deck, slide)
        else:
            raise ToolError("invalid_arguments", "name the slide when both ends are points",
                            field="slide")
        sides = ["auto", "auto"]
        for position, (end, spec) in enumerate(((begin, item["from"]), (finish, item["to"]))):
            sides[position] = spec.get("side", "auto")
        if not isinstance(begin, tuple) and not isinstance(finish, tuple):
            auto = _auto_sides(begin, finish)
            sides = [auto[i] if sides[i] == "auto" else sides[i] for i in range(2)]
        for position, end in enumerate((begin, finish)):
            if not isinstance(end, tuple) and sides[position] == "auto":
                sides[position] = "right" if position == 0 else "left"
        ends = [end if isinstance(end, tuple) else (end, sides[position])
                for position, end in enumerate((begin, finish))]
        connector = host.add_connector(item["kind"], ends[0], ends[1], name=item.get("name"))
        if item.get("line"):
            apply_line(connector, item["line"])
        made.append(connector.id)
        _define(call, item, connector.id)
        call.touch(host.slide_id)
    return Result(summary=f"Added {len(made)} connector(s)", created=made)


@tool("ppt_arrange",
      "Change shapes' z-order; group, ungroup, duplicate, delete, align or distribute "
      "them (equal gaps); grid them in order, row by row, sized to equal cells (pictures "
      "keep aspect), or flow them at their size. Returns new addresses and moved boxes.",
      {"doc": string("Document id."),
       "targets": array(string(), "Shapes, or $refs.", min_items=1),
       "action": string("What to do.", enum=["front", "back", "forward", "backward", "group",
                                             "ungroup", "duplicate", "delete", "align",
                                             "distribute", "grid", "flow"]),
       "dx": number("duplicate: offset. Default 0.", minimum=-10000, maximum=10000,
                    optional=True),
       "dy": number("duplicate: offset. Default 0.", minimum=-10000, maximum=10000,
                    optional=True),
       "ref": string("group, or duplicate of one shape: ref name for the new shape.",
                     optional=True),
       "edge": string("align.", enum=list(EDGES), optional=True),
       "to": string("align: to the targets' bounds (default), slide or content area.",
                    enum=list(_ALIGN_TO), optional=True),
       "axis": string("distribute.", enum=["horizontal", "vertical"], optional=True),
       "within": string("distribute: between the outermost targets (default) or across "
                        "the content area.", enum=list(_WITHIN), optional=True),
       "gap": number("distribute: a fixed gap instead; grid, flow: between cells.",
                     minimum=0, maximum=10000, optional=True),
       "gap_y": number("Between rows, if not gap.", minimum=0, maximum=10000,
                       optional=True),
       "columns": integer("grid; flow: at most.", minimum=1, maximum=200, optional=True),
       "rows": integer("grid.", minimum=1, maximum=200, optional=True),
       "box": _box("grid, flow. Default the content area.", optional=True)},
      kind=KIND, group=GRAPHICS, mutates=True, refs=("targets[]",))
@library_errors
def ppt_arrange(call, doc, targets, action, dx=0, dy=0, ref=None, edge=None, to=None,
                axis=None, within=None, gap=None, gap_y=None, columns=None, rows=None,
                box=None):
    deck = call.document
    shapes = [shape_at(deck, address, field=f"targets[{i}]") for i, address in
              enumerate(targets)]
    for shape in shapes:
        call.touch(shape._slide.slide_id)
        if action not in ("delete", "ungroup"):
            call.touch(shape.id)
    for name, value, owners in (("edge", edge, ("align",)), ("to", to, ("align",)),
                                ("axis", axis, ("distribute",)),
                                ("within", within, ("distribute",)),
                                ("gap", gap, ("distribute", "grid", "flow")),
                                ("gap_y", gap_y, ("grid", "flow")),
                                ("columns", columns, ("grid", "flow")),
                                ("rows", rows, ("grid",)), ("box", box, ("grid", "flow"))):
        if value is not None and action not in owners:
            raise ToolError("invalid_arguments", f"{name} is for {' and '.join(owners)}, "
                            f"not {action}", field=name)
    if action in ("grid", "flow"):
        return _grid(call, shapes, action, gap=gap, gap_y=gap_y, columns=columns, rows=rows,
                     box=box)
    if action in ("align", "distribute"):
        return _line_up(call, shapes, action, edge=edge, to=to, axis=axis, within=within,
                        gap=gap)
    created: list[str] = []
    removed: list[str] = []
    changed: list[str] = []
    if action in ("front", "back", "forward", "backward"):
        method = {"front": "bring_to_front", "back": "send_to_back",
                  "forward": "bring_forward", "backward": "send_backward"}[action]
        ordered = shapes if action in ("front", "forward") else list(reversed(shapes))
        for shape in ordered:
            getattr(shape, method)()
        changed = [shape.id for shape in shapes]
    elif action == "group":
        host = shapes[0]._slide
        group = host.group(shapes)
        created = [group.id]
        if ref:
            call.define_ref(ref, group.id)
    elif action == "ungroup":
        for shape in shapes:
            if shape.kind != "group":
                raise ToolError("invalid_arguments", f"{shape.id} is not a group",
                                field="targets")
            removed.append(shape.id)
            changed += [child.id for child in shape.ungroup()]
    elif action == "duplicate":
        for shape in shapes:
            copy = shape.duplicate(emu(dx or 0), emu(dy or 0))
            created.append(copy.id)
        if ref:
            if len(created) != 1:
                raise ToolError("invalid_arguments", "a ref names one new shape; duplicate "
                                "one target to name its copy", field="ref")
            call.define_ref(ref, created[0])
    elif action == "delete":
        for shape in shapes:
            removed.append(shape.id)
            shape.delete()
    call.touch(*created, *changed)
    return Result(summary=f"{action}: {len(shapes)} shape(s)", created=created,
                  removed=removed, changed=changed)


def _one_slide(shapes):
    host = shapes[0]._slide
    if any(shape._slide.slide_id != host.slide_id for shape in shapes):
        raise ToolError("invalid_arguments", "the targets must be on one slide",
                        field="targets")
    return host


def _grid(call, shapes, action, *, gap, gap_y, columns, rows, box) -> Result:
    """ppt_arrange's grid and flow: :func:`pptx_agent.edit.layout.grid` (``fit="fill"``)
    and :func:`~pptx_agent.edit.layout.flow` on the targets in the order given, in ``box``
    or the slide's content area (as align and distribute take it)."""
    from ..edit import layout
    from ..edit.arrange import outermost

    host = _one_slide(shapes)
    area = box_emu(box) if box is not None else host.content_area
    gutter = (emu(gap or 0), emu(gap_y if gap_y is not None else gap or 0))
    shapes = outermost(shapes)
    try:
        if action == "grid":
            placed = layout.grid(shapes, box=area, rows=rows, columns=columns,
                                 gutter=gutter, fit="fill")
        else:
            placed = layout.flow(shapes, box=area, gutter=gutter, columns=columns)
    except layout.LayoutError as error:
        raise ToolError("refused", str(error), field="box" if box is not None else "gap",
                        details={"area": box_pt(area)}) from None
    changed = [shape.id for shape in shapes]
    boxes = {shape.id: box_pt(bounds) for shape, bounds in zip(shapes, placed)}
    where = "box" if box is not None else "content area"
    return Result(summary=f"{action}: {len(changed)} shape(s) laid out in the {where}",
                  changed=changed, data={"area": box_pt(area), "boxes": boxes})


def _line_up(call, shapes, action, *, edge, to, axis, within, gap) -> Result:
    """ppt_arrange's align and distribute: :mod:`pptx_agent.edit.arrange` on the targets,
    the content area being :attr:`Slide.content_area <pptx_agent.Slide.content_area>` (the
    layout's body placeholders, else the band under the title)."""
    from ..edit import arrange

    host = _one_slide(shapes)
    if action == "align":
        if edge is None:
            raise ToolError("invalid_arguments", "align needs an edge", field="edge",
                            valid_options=list(EDGES))
        where = to or "shapes"
        moved = arrange.align(shapes, edge, to=_ALIGN_TO[where],
                              box=host.content_area if where == "content_area" else None)
        how = f"{edge} of the {where.replace('_', ' ')}"
    else:
        if axis is None:
            raise ToolError("invalid_arguments", "distribute needs an axis", field="axis",
                            valid_options=["horizontal", "vertical"])
        where = within or "shapes"
        moved = arrange.distribute(shapes, axis, to=_WITHIN[where],
                                   box=host.content_area if where == "content_area" else None,
                                   gap=None if gap is None else emu(gap))
        how = f"{axis}ly within the {where.replace('_', ' ')}"
    changed = [shape.id for shape in moved]
    boxes = {shape.id: box_pt(shape.drawn_bounds) for shape in arrange.outermost(shapes)[:20]}
    return Result(summary=f"{action}: {len(changed)} of {len(boxes)} shape(s) moved, {how}",
                  changed=changed, data={"moved": changed, "boxes": boxes})


TOOLS = [ppt_add_shape, ppt_set_shape, ppt_add_picture, ppt_add_connector, ppt_arrange]
