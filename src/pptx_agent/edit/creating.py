"""Creating shapes, and keeping connectors attached to them.

:class:`ShapeFactory` is mixed into :class:`~pptx_agent.edit.document.Slide` and
:class:`~pptx_agent.edit.document.Shape`: on a slide it adds to the shape tree, on a group to
the group, in the group's child space -- and the group is then re-fitted the way PowerPoint
re-fits one, so nothing else moves.  Every creation is one undo step, gets a ``cNvPr@id``
unused on the slide, PowerPoint's name for it, and this library's durable id stamp.  The XML
is :mod:`.authoring`'s, which says what was measured.

Connectors (:meth:`ShapeFactory.add_connector`) attach to a shape's connection sites --
``(shape, site)`` -- or start and end at points.  An attached end is written as
``a:stCxn``/``a:endCxn``, and from then on the connector follows its shapes: moving,
resizing, rotating, flipping or re-shaping one through this API re-routes every connector
attached to it (:func:`reroute`), keeping the sites, as PowerPoint does when a shape is
dragged.  Deleting a shape detaches the connectors that ended on it.

Connector geometry is computed on the slide: a shape inside a group is placed through the
group's offset and scale.  Group rotation and flips are not composed, as for
:attr:`~pptx_agent.edit.document.Shape.slide_bounds`.
"""

from __future__ import annotations

import math
import re
from typing import TYPE_CHECKING, Any, Iterator, Mapping

from ..oxml.xml import (
    PRESET_GEOMETRIES,
    Element,
    append_in_order,
    find,
    get_int,
    local_name,
    qn,
    remove,
    subelement,
)
from . import authoring, connectors
from .color import Color
from .connectors import End, Frame
from .ids import cnv_pr
from .presets import adjustment_defaults, connection_sites, evaluate

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape, Slide

_GUID = re.compile(r"\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}")
_CONNECTOR_NAMES = {"straight": "Straight Arrow Connector", "elbow": "Elbow Connector",
                    "curved": "Curved Connector"}


class ShapeFactory:
    """``add_shape``, ``add_textbox``, ``add_connector`` and ``add_table``.

    On a :class:`Slide` the new shape goes on top of the slide's shapes; on a group, on top
    of the group's children, with ``left``/``top``/``width``/``height`` in the group's child
    coordinate space.  Lengths are EMU.

    For example::

        box = slide.add_shape("roundRect", 914400, 914400, 1828800, 914400, text="Plan")
    """

    def _factory(self) -> tuple["Slide", Element, "Shape | None"]:  # pragma: no cover
        raise NotImplementedError

    # -- autoshapes and text boxes ----------------------------------------------------------

    def add_shape(self, preset: str, left: int, top: int, width: int, height: int, *,
                  text: str | None = None, name: str | None = None,
                  fill: "str | Color | None" = None, line: Any = None,
                  adjustments: Mapping[str, int] | None = None) -> "Shape":
        """A new autoshape of any preset geometry (``"rect"``, ``"roundRect"``,
        ``"rightArrow"``...), styled by the theme as PowerPoint styles a new shape.

        Its text frame, as PowerPoint makes one: autofit ``"none"`` (text may overflow),
        wrapping on, PowerPoint's default insets (0.1 in left and right, 0.05 in top and
        bottom), anchored ``"middle"`` and centred, in the theme's body font at the deck's
        default size (18 pt in PowerPoint's own) and its ``lt1`` colour; filled
        ``accent1`` with an ``accent1`` outline shaded 15%.
        Change any of it on :attr:`Shape.text_frame` (``insets``, ``anchor``, ``wrap``,
        ``autofit``).

        ``text`` fills its (centred) text body; ``fill`` is a colour or ``"none"``;
        ``line`` a colour, ``"none"``, or a mapping of :class:`LineFormat` properties
        (``{"width": 25400, "color": "accent2", "dash": "dash", "tail": "triangle"}``);
        ``adjustments`` the preset's adjust values by name (see :attr:`Shape.adjustments`).

        For example::

            box = slide.add_shape("roundRect", 914400, 914400, 1828800, 914400, text="Plan")
        """
        if preset not in PRESET_GEOMETRIES:
            raise ValueError(f"{preset!r} is not a preset geometry (ST_ShapeType)")
        _check_size(width, height)
        slide, container, group = self._factory()
        language = authoring.default_language(slide.document.package)
        values = _adjustment_values(preset, adjustments or {})

        def build(identifier: int, local: str) -> Element:
            label = name or authoring.default_name(
                authoring.SHAPE_NAMES.get(preset, authoring.GENERIC_NAME), identifier)
            return authoring.autoshape(identifier, label, local, preset, left, top, width,
                                       height, language, values)

        return _create(slide, container, group, build, text=text, fill=fill, line=line)

    def add_textbox(self, left: int, top: int, width: int, height: int, text: str = "", *,
                    name: str | None = None, fill: "str | Color | None" = None,
                    line: Any = None, autofit: str = "shape") -> "Shape":
        """A new text box, as PowerPoint makes one: no fill, no outline, wrapping on,
        PowerPoint's default insets (0.1 in left and right, 0.05 in top and bottom),
        anchored ``"top"``, left-aligned, in the theme's body font at the deck's default
        size (18 pt in PowerPoint's own).

        ``autofit`` is PowerPoint's default, ``"shape"``: the box is resized to its text
        when PowerPoint lays it out (``spAutoFit``), so ``height`` is where it starts, not
        what it keeps.  ``"none"`` keeps the box exactly as given, the text free to
        overflow (check :meth:`~Shape.text_fit`); ``"normal"`` shrinks the text on
        overflow -- once PowerPoint edits it (see :attr:`TextFrame.autofit`).  For a box
        sized to its text from the start: :func:`~pptx_agent.measure_text` or
        :meth:`Shape.fit_height`.

        For example::

            slide.add_textbox(914400, 914400, 3657600, 369332, "Notes", autofit="none")
        """
        if autofit not in authoring.AUTOFIT_ELEMENTS:
            raise ValueError(f"autofit is one of {sorted(authoring.AUTOFIT_ELEMENTS)}, "
                             f"not {autofit!r}")
        _check_size(width, height)
        slide, container, group = self._factory()
        language = authoring.default_language(slide.document.package)

        def build(identifier: int, local: str) -> Element:
            label = name or authoring.default_name(authoring.TEXT_BOX_NAME, identifier)
            return authoring.text_box(identifier, label, local, left, top, width, height,
                                      language, autofit)

        return _create(slide, container, group, build, text=text or None, fill=fill,
                       line=line)

    # -- connectors ---------------------------------------------------------------------------

    def add_connector(self, kind: str, begin: Any, end: Any, *, name: str | None = None,
                      line: Any = None) -> "Shape":
        """A new connector -- ``"straight"``, ``"elbow"`` or ``"curved"`` -- from ``begin``
        to ``end``.

        Each end is a point ``(x, y)`` (in this container's coordinate space) or a
        ``(shape, site)`` pair: a shape (or its id) and the index of one of its
        :attr:`~Shape.connection_sites` -- or the side, ``"top"``, ``"right"``,
        ``"bottom"``, ``"left"`` (:meth:`~Shape.connection_site`).  ``begin`` is where the
        line starts (its ``head``/``start`` arrowhead), ``end`` where it ends (its
        ``tail``/``end``: the point of an arrow from ``begin`` to ``end``).  An attached end is written as ``stCxn``/``endCxn``,
        and the connector is re-routed whenever that shape moves or changes through this
        API.  Elbow and curved connectors are routed the way PowerPoint routes them (see
        :mod:`.connectors`).  Arrowheads are the outline's: ``line={"tail": "triangle"}``,
        or ``connector.line.tail = "triangle"`` afterwards.

        For example::

            slide.add_connector("elbow", (box, 3), (goal, 1), line={"tail": "triangle"})
        """
        if kind not in connectors.KINDS:
            raise ValueError(f"connector kind must be one of {sorted(connectors.KINDS)}")
        slide, container, group = self._factory()
        start, start_ref = _resolve_end(slide, container, begin, "begin")
        finish, end_ref = _resolve_end(slide, container, end, "end")
        frame = _frame_in(slide, container, kind, start, finish)

        def build(identifier: int, local: str) -> Element:
            label = name or authoring.default_name(_CONNECTOR_NAMES[kind], identifier)
            element = authoring.connector(identifier, label, local, frame)
            _write_connections(element, start_ref, end_ref)
            return element

        return _create(slide, container, group, build, line=line)

    # -- tables -------------------------------------------------------------------------------

    def add_table(self, rows: int, cols: int, left: int, top: int, width: int, height: int, *,
                  style: str | None = None, name: str | None = None) -> "Shape":
        """A new ``rows`` x ``cols`` table, its columns and rows sharing the frame evenly.
        Each cell wraps, has PowerPoint's default margins (0.1 in left and right, 0.05 in
        top and bottom) and is anchored ``"top"`` (a cell's :attr:`TableCell.text_frame`
        sets them); a row grows to fit its text when PowerPoint draws it.

        ``style`` is a table style id (``"{5C22544A-...}"``); by default PowerPoint's own
        default, Medium Style 2 - Accent 1, with a header row and banded rows -- as
        PowerPoint inserts one.  Edit it with :attr:`Shape.table`.  PowerPoint does not put
        tables in groups, so neither does this.

        For example::

            table = slide.add_table(3, 4, 914400, 914400, 7315200, 1097280).table
        """
        if int(rows) < 1 or int(cols) < 1:
            raise ValueError("a table needs at least one row and one column")
        _check_size(width, height)
        if style is not None and not _GUID.fullmatch(style):
            raise ValueError(f"a table style id is a GUID in braces, not {style!r}")
        slide, container, group = self._factory()
        if group is not None:
            raise ValueError("PowerPoint does not allow a table in a group")
        language = authoring.default_language(slide.document.package)

        def build(identifier: int, local: str) -> Element:
            label = name or authoring.default_name(authoring.TABLE_NAME, identifier)
            return authoring.table(identifier, label, local, int(rows), int(cols), left, top,
                                   width, height, language,
                                   style or authoring.DEFAULT_TABLE_STYLE)

        return _create(slide, container, group, build)

    def add_chart(self, chart_type: str, categories, series, left: int, top: int, width: int,
                  height: int, *, title: str | None = None, axis_titles: Mapping | None = None,
                  legend: str | None = "bottom", number_format: str | None = None,
                  name: str | None = None) -> "Shape":
        """A new chart from data, in a graphic frame at ``left``/``top`` (EMU), as PowerPoint
        inserts one: the chart part and an embedded workbook that holds the same numbers, so
        Edit Data opens them (ooxml-edit's :func:`~ooxml_edit.charts.add_chart`).

        ``chart_type`` is ``column``, ``stacked_column``, ``bar``, ``stacked_bar``, ``line``,
        ``pie`` or ``scatter``; ``categories`` the category labels (a scatter chart's x
        values); ``series`` ``[{"name": ..., "values": [...]}]``, one value per category
        (``None`` is a blank); data that cannot be charted raises
        :class:`~pptx_agent.ChartDataError` before anything changes.  ``title``, ``axis_titles`` (``{"category", "value"}``),
        ``legend`` (``bottom``, ``right``, ``top``, ``left``, ``top_right`` or ``None``) and
        ``number_format`` (an Excel format code for the values) are optional.  It looks as
        PowerPoint's new chart of that type looks, in the deck's theme (measured: accents in
        order, 18.62 pt title, 11.97 pt labels, legend at the bottom, no fill).  Edit it with
        :attr:`Shape.chart`.  One undo step.

        For example::

            frame = slide.add_chart("column", ["Q1", "Q2"],
                                    [{"name": "North", "values": [12.4, 13.1]}],
                                    457200, 1371600, 8229600, 4572000, title="Revenue")
        """
        from ooxml_edit.charts import POWERPOINT_LOOK, add_chart

        _check_size(width, height)
        slide, container, group = self._factory()
        if group is not None:
            raise ValueError("add a chart to the slide, then group it")
        language = authoring.default_language(slide.document.package)
        made = []

        def build(identifier: int, local: str) -> Element:
            label = name or authoring.default_name(authoring.CHART_NAME, identifier)
            made.append(add_chart(slide.document.package, slide.part_path, chart_type,
                                  categories, series, title=title,
                                  axis_titles=dict(axis_titles or {}), legend=legend,
                                  number_format=number_format, look=POWERPOINT_LOOK,
                                  lang=language))
            return authoring.chart_frame(identifier, label, local, left, top, width, height,
                                         made[-1].graphic())

        return _create(slide, container, group, build)


# -- creation ----------------------------------------------------------------------------------


def _check_size(width: int, height: int) -> None:
    if int(width) < 0 or int(height) < 0:
        raise ValueError("a shape's width and height cannot be negative")


def _create(slide: "Slide", container: Element, group: "Shape | None", build, *,
            text: str | None = None, fill: Any = None, line: Any = None) -> "Shape":
    document = slide.document
    with document.batch():
        document.history.checkpoint()
        slide._watch()
        identifier = slide._next_shape_id()
        element = build(identifier, str(identifier))
        append_in_order(container, element)
        slide._invalidate()
        shape = slide._wrap(element)
        if group is not None:
            group._refit()
        if text is not None:
            shape.set_text(text)
        if fill is not None:
            shape.fill = fill
        if line is not None:
            apply_line(shape, line)
        slide._settle()
        return slide._wrap(element)


def apply_line(shape: "Shape", spec: Any) -> None:
    """A colour, ``"none"``, or a mapping of :class:`LineFormat` properties."""
    outline = shape.line
    if isinstance(spec, Mapping):
        unknown = set(spec) - {"width", "color", "visible", "dash", "cap", "head", "tail",
                               "start", "end"}
        if unknown:
            raise ValueError(f"unknown outline properties {sorted(unknown)}")
        for key in ("visible", "width", "color", "dash", "cap", "head", "tail", "start", "end"):
            if key in spec:
                setattr(outline, key, spec[key])
    elif isinstance(spec, str) and spec.strip().lower() == "none":
        outline.visible = False
    elif isinstance(spec, (str, Color)):
        outline.color = spec
    else:
        raise TypeError(f"cannot use {spec!r} as an outline")


# -- adjustments -------------------------------------------------------------------------------


def _adjustment_values(preset: str, given: Mapping[str, int]) -> list[tuple[str, int]]:
    """The ``a:avLst`` to write: nothing when nothing is set, otherwise *every* adjust value
    of the preset in its order, defaults filled in -- what PowerPoint writes (measured), and
    what it needs: a ``star5`` with ``adj`` but no ``hf``/``vf`` is repaired on open."""
    defaults = adjustment_defaults(preset)
    names = [name for name, _ in defaults]
    unknown = set(given) - set(names)
    if unknown:
        raise KeyError(f"{preset} has no adjustment {sorted(unknown)}; it has {names}")
    if not given:
        return []
    return [(name, int(given.get(name, default))) for name, default in defaults]


class Adjustments:
    """A preset shape's adjust values (``a:avLst``), by name, in raw DrawingML units.

    ``shape.adjustments["adj"] = 30000`` rounds a rounded rectangle's corners to 30% of
    the short side's half; most are 1/100,000 of a length, angles are 1/60,000 of a degree.
    A name the shape does not set reads as the preset's default; ``del`` returns one to its
    default and ``reset`` empties the list.  As PowerPoint does (measured), setting any value
    writes all of them, in the preset's order, defaults filled in -- PowerPoint repairs a
    five-point star that has ``adj`` without ``hf`` and ``vf``.  Connectors attached to the
    shape follow its sites.

    For example::

        shape.adjustments["adj"] = 30000                # corner radius of a roundRect
    """

    def __init__(self, shape: "Shape") -> None:
        self._document = shape._slide.document
        self._identifier = shape.id

    @property
    def _shape(self) -> "Shape":
        return self._document.shape(self._identifier)

    @property
    def preset(self) -> str:
        """The preset geometry the adjust values belong to.

        For example::

            shape.adjustments.preset                     # 'roundRect'
        """
        preset = self._shape.preset
        if preset in (None, "custom"):
            raise ValueError(f"{self._identifier} has no preset geometry to adjust")
        return preset

    @property
    def names(self) -> tuple[str, ...]:
        """The preset's adjust value names, in order.

        For example::

            shape.adjustments.names                      # ('adj',)
        """
        return tuple(name for name, _ in adjustment_defaults(self.preset))

    @property
    def defaults(self) -> dict[str, int]:
        """The preset's default values.

        For example::

            shape.adjustments.defaults                   # {'adj': 16667}
        """
        return dict(adjustment_defaults(self.preset))

    def explicit(self) -> dict[str, int]:
        """Only the values the shape sets itself.

        For example::

            shape.adjustments.explicit
        """
        values = {}
        for node in self._gd_nodes():
            parsed = _literal(node.get("fmla", ""))
            if parsed is not None:
                values[node.get("name")] = parsed
        return values

    def __getitem__(self, key: "str | int") -> int:
        name = self._name(key)
        return self.explicit().get(name, self.defaults[name])

    def __setitem__(self, key: "str | int", value: int) -> None:
        self.set(**{self._name(key): value})

    def __delitem__(self, key: "str | int") -> None:
        self._write({self._name(key): None})

    def __iter__(self) -> Iterator[str]:
        return iter(self.names)

    def __len__(self) -> int:
        return len(self.names)

    def __contains__(self, key: object) -> bool:
        return key in self.names

    def items(self) -> list[tuple[str, int]]:
        """Every value in effect, explicit or default, as ``(name, value)``.

        For example::

            dict(shape.adjustments.items())
        """
        return [(name, self[name]) for name in self.names]

    def set(self, **values: int) -> "Adjustments":
        """Set several at once, as one undo step.

        For example::

            shape.adjustments.set(adj1=20000, adj2=40000)
        """
        for name, value in values.items():
            if name not in self.names:
                raise KeyError(f"{self.preset} has no adjustment {name!r}; it has "
                               f"{list(self.names)}")
            if isinstance(value, bool) or not isinstance(value, int):
                raise TypeError(f"{name}: an adjust value is an integer in DrawingML units")
        self._write(dict(values))
        return self

    def reset(self) -> "Adjustments":
        """Remove every explicit value, back to the preset's defaults.

        For example::

            shape.adjustments.reset()
        """
        self._write({name: None for name in self.names})
        return self

    def _name(self, key: "str | int") -> str:
        names = self.names
        if isinstance(key, int):
            return names[key]
        if key not in names:
            raise KeyError(f"{self.preset} has no adjustment {key!r}; it has {list(names)}")
        return key

    def _gd_nodes(self) -> list[Element]:
        geometry = find(self._shape._element, "p:spPr/a:prstGeom")
        av_list = None if geometry is None else geometry.find(qn("a:avLst"))
        return [] if av_list is None else list(av_list.findall(qn("a:gd")))

    def _write(self, changes: dict[str, int | None]) -> None:
        current = self.explicit()
        wanted = dict(current)
        for name, value in changes.items():
            if value is None:
                wanted.pop(name, None)
            else:
                wanted[name] = int(value)
        values = dict(_adjustment_values(self.preset, wanted))
        written = [(node.get("name"), _literal(node.get("fmla", ""))) for node in self._gd_nodes()]
        if written == list(values.items()):
            return
        shape = self._shape
        with self._document.batch():
            shape._before_change()
            geometry = find(shape._element, "p:spPr/a:prstGeom")
            av_list = subelement(geometry, "a:avLst")
            for node in list(av_list):
                remove(node)
            av_list.text = None
            for name, value in values.items():
                av_list.append(av_list.makeelement(qn("a:gd"), {"name": name,
                                                                "fmla": f"val {value}"}))
            shape._after_change()
            shape._slide._reroute_for(shape)

    def __repr__(self) -> str:
        return f"<Adjustments {self._identifier} {dict(self.items())}>"


def _literal(formula: str) -> int | None:
    match = re.fullmatch(r"\s*val\s+(-?\d+)\s*", formula)
    return int(match.group(1)) if match else None


# -- connection sites --------------------------------------------------------------------------


def local_sites(element: Element, width: float, height: float) -> list[tuple[float, float, float]]:
    """``(angle, x, y)`` of every connection site of a shape element, in its own box."""
    name = local_name(element)
    if name not in {"sp", "pic", "cxnSp"}:
        return []
    properties = element.find(qn("p:spPr"))
    preset_node = None if properties is None else properties.find(qn("a:prstGeom"))
    if preset_node is not None:
        given = {}
        av_list = preset_node.find(qn("a:avLst"))
        for node in [] if av_list is None else av_list.findall(qn("a:gd")):
            value = _literal(node.get("fmla", ""))
            if value is not None:
                given[node.get("name")] = value
        return connection_sites(preset_node.get("prst"), width, height, given)
    custom = None if properties is None else properties.find(qn("a:custGeom"))
    if custom is not None:
        return _custom_sites(custom, width, height)
    if name == "pic":  # a picture without geometry is a rectangle
        return connection_sites("rect", width, height)
    return []


def _custom_sites(custom: Element, width: float, height: float) -> list[tuple[float, float, float]]:
    from .presets import _builtins, _value

    variables = _builtins(width, height)
    for tag in ("a:avLst", "a:gdLst"):
        holder = custom.find(qn(tag))
        for node in [] if holder is None else holder.findall(qn("a:gd")):
            try:
                variables[node.get("name")] = evaluate(node.get("fmla", ""), variables)
            except (ValueError, IndexError):
                variables[node.get("name")] = 0.0
    sites = []
    holder = custom.find(qn("a:cxnLst"))
    for cxn in [] if holder is None else holder.findall(qn("a:cxn")):
        position = cxn.find(qn("a:pos"))
        if position is None:
            continue
        sites.append((_value(cxn.get("ang", "0"), variables),
                      _value(position.get("x", "0"), variables),
                      _value(position.get("y", "0"), variables)))
    return sites


def shape_sites(shape: "Shape") -> list[End]:
    """Every connection site of ``shape`` on the slide: point, direction and the shape's box."""
    left, top, width, height = shape.left, shape.top, shape.width, shape.height
    if None in (left, top, width, height):
        return []
    sites = local_sites(shape._element, width, height)
    if not sites:
        return []
    rotation = shape.rotation
    flip_h, flip_v = shape.flip_h, shape.flip_v
    centre_x, centre_y = left + width / 2, top + height / 2
    angle = math.radians(rotation)
    cos, sin = math.cos(angle), math.sin(angle)

    def place(x: float, y: float) -> tuple[float, float]:
        dx, dy = x - centre_x, y - centre_y
        return to_slide(shape, centre_x + dx * cos - dy * sin, centre_y + dx * sin + dy * cos)

    corners = [place(x, y) for x, y in ((left, top), (left + width, top),
                                         (left, top + height), (left + width, top + height))]
    box = (min(c[0] for c in corners), min(c[1] for c in corners),
           max(c[0] for c in corners), max(c[1] for c in corners))
    ends = []
    for site_angle, x, y in sites:
        if flip_h:
            x, site_angle = width - x, 10800000 - site_angle
        if flip_v:
            y, site_angle = height - y, -site_angle
        point = place(left + x, top + y)
        ends.append(End(point, connectors.axis(site_angle + rotation * 60000), box))
    return ends


# -- coordinate spaces -------------------------------------------------------------------------


def _groups_of(element: Element) -> list[Element]:
    """The groups around ``element``, innermost first."""
    return [ancestor for ancestor in element.iterancestors(qn("p:grpSp"))]


def _group_map(group: Element) -> tuple[float, float, float, float, float, float]:
    """``(ox, oy, cox, coy, sx, sy)``: slide = o + (child - co) * s for one group."""
    xfrm = find(group, "p:grpSpPr/a:xfrm")
    off = None if xfrm is None else xfrm.find(qn("a:off"))
    ext = None if xfrm is None else xfrm.find(qn("a:ext"))
    ch_off = None if xfrm is None else xfrm.find(qn("a:chOff"))
    ch_ext = None if xfrm is None else xfrm.find(qn("a:chExt"))
    ox, oy = get_int(off, "x", 0) or 0, get_int(off, "y", 0) or 0
    ex, ey = get_int(ext, "cx", 0) or 0, get_int(ext, "cy", 0) or 0
    cox = get_int(ch_off, "x", ox) if ch_off is not None else ox
    coy = get_int(ch_off, "y", oy) if ch_off is not None else oy
    cex = get_int(ch_ext, "cx", ex) if ch_ext is not None else ex
    cey = get_int(ch_ext, "cy", ey) if ch_ext is not None else ey
    return ox, oy, cox or 0, coy or 0, (ex / cex if cex else 1.0), (ey / cey if cey else 1.0)


def container_to_slide(container: Element, x: float, y: float) -> tuple[float, float]:
    """A point in ``container``'s coordinate space (the slide, or a group's child space)."""
    chain = ([container] if container.tag == qn("p:grpSp") else []) + _groups_of(container)
    for group in chain:
        ox, oy, cox, coy, sx, sy = _group_map(group)
        x, y = ox + (x - cox) * sx, oy + (y - coy) * sy
    return x, y


def slide_to_container(container: Element, x: float, y: float) -> tuple[float, float]:
    chain = ([container] if container.tag == qn("p:grpSp") else []) + _groups_of(container)
    for group in reversed(chain):
        ox, oy, cox, coy, sx, sy = _group_map(group)
        x, y = cox + (x - ox) / (sx or 1.0), coy + (y - oy) / (sy or 1.0)
    return x, y


def to_slide(shape: "Shape", x: float, y: float) -> tuple[float, float]:
    """A point in ``shape``'s parent space, on the slide."""
    return container_to_slide(shape._element.getparent(), x, y)


# -- connections -------------------------------------------------------------------------------


def _resolve_end(slide: "Slide", container: Element, spec: Any,
                 which: str) -> tuple[End, tuple[str, int] | None]:
    """An end in slide space, and ``(cNvPr id, site)`` when it is attached."""
    from .document import Shape

    if isinstance(spec, (tuple, list)) and len(spec) == 2 and isinstance(spec[0], (Shape, str)):
        target = slide.shape(spec[0]) if isinstance(spec[0], str) else spec[0]
        site = spec[1]
        if isinstance(site, str) and site in ("top", "right", "bottom", "left"):
            site = target.connection_site(site)[1]
        if isinstance(site, bool) or not isinstance(site, int):
            raise TypeError(f"{which}: a connection site is an index or 'top', 'right', "
                            "'bottom', 'left'")
        if target._slide.slide_id != slide.slide_id:
            raise ValueError(f"{which}: {target.id} is on another slide")
        ends = shape_sites(target)
        if not ends:
            raise ValueError(f"{which}: {target.id} ({target.kind}) has no connection sites")
        if not 0 <= site < len(ends):
            raise IndexError(f"{which}: {target.id} has connection sites 0..{len(ends) - 1}, "
                             f"not {site}")
        raw = _unique_raw_id(slide, target)
        return ends[site], (raw, site)
    if isinstance(spec, (tuple, list)) and len(spec) == 2 \
            and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in spec):
        return End(container_to_slide(container, float(spec[0]), float(spec[1]))), None
    raise TypeError(f"{which} must be a point (x, y) or a (shape, site) pair, not {spec!r}")


def _unique_raw_id(slide: "Slide", shape: "Shape") -> str:
    properties = cnv_pr(shape._element)
    raw = None if properties is None else properties.get("id")
    if not raw or sum(1 for node in slide._sp_tree().iter(qn("p:cNvPr"))
                      if node.get("id") == raw) != 1:
        raise ValueError(f"{shape.id}: its cNvPr id {raw!r} is not unique on the slide, so "
                         f"a connector could not name it")
    return raw


def _frame_in(slide: "Slide", container: Element, kind: str, start: End, finish: End) -> Frame:
    """The connector's frame, routed on the slide and written in ``container``'s space."""
    if kind == "straight":
        a = slide_to_container(container, *start.point)
        b = slide_to_container(container, *finish.point)
        return connectors.straight_frame(a, b)
    points = [slide_to_container(container, x, y) for x, y in connectors.route(start, finish)]
    return connectors.bent_frame(points, connectors.KINDS[kind])


def _write_connections(element: Element, start: tuple[str, int] | None,
                       end: tuple[str, int] | None) -> None:
    properties = find(element, "p:nvCxnSpPr/p:cNvCxnSpPr")
    for tag in ("a:stCxn", "a:endCxn"):
        node = properties.find(qn(tag))
        if node is not None:
            remove(node)
    if start is None and end is None:
        return
    subelement(properties, "a:cxnSpLocks")
    for tag, reference in (("a:stCxn", start), ("a:endCxn", end)):
        if reference is not None:
            subelement(properties, tag, id=reference[0], idx=str(reference[1]))


def connection(element: Element, which: str) -> tuple[str, int] | None:
    """``(cNvPr id, site)`` of a connector end (``"begin"``/``"end"``), if attached."""
    tag = "a:stCxn" if which == "begin" else "a:endCxn"
    node = find(element, f"p:nvCxnSpPr/p:cNvCxnSpPr/{tag}")
    if node is None or node.get("id") is None:
        return None
    return node.get("id"), get_int(node, "idx", 0) or 0


def connector_kind(element: Element) -> str | None:
    """``straight``/``elbow``/``curved`` for a connector this module can route."""
    node = find(element, "p:spPr/a:prstGeom")
    preset = None if node is None else node.get("prst")
    if preset in ("straightConnector1", "line"):
        return "straight"
    if preset and re.fullmatch(r"bentConnector[2-5]", preset):
        return "elbow"
    if preset and re.fullmatch(r"curvedConnector[2-5]", preset):
        return "curved"
    return None


def current_ends(element: Element) -> tuple[tuple[float, float], tuple[float, float]] | None:
    """A connector's two ends where its frame draws them, in its parent's space."""
    xfrm = find(element, "p:spPr/a:xfrm")
    off = None if xfrm is None else xfrm.find(qn("a:off"))
    ext = None if xfrm is None else xfrm.find(qn("a:ext"))
    if off is None or ext is None:
        return None
    return connectors.endpoints(
        get_int(off, "x", 0) or 0, get_int(off, "y", 0) or 0,
        get_int(ext, "cx", 0) or 0, get_int(ext, "cy", 0) or 0,
        get_int(xfrm, "rot", 0) or 0, xfrm.get("flipH") in {"1", "true"},
        xfrm.get("flipV") in {"1", "true"})


def write_frame(element: Element, frame: Frame) -> None:
    """Put ``frame`` on a connector: transform, preset and adjustments."""
    properties = subelement(element, "p:spPr")
    xfrm = subelement(properties, "a:xfrm")
    for attribute in ("rot", "flipH", "flipV"):
        xfrm.attrib.pop(attribute, None)
    if frame.rot:
        xfrm.set("rot", str(frame.rot))
    if frame.flip_h:
        xfrm.set("flipH", "1")
    if frame.flip_v:
        xfrm.set("flipV", "1")
    off, ext = subelement(xfrm, "a:off"), subelement(xfrm, "a:ext")
    off.set("x", str(frame.x))
    off.set("y", str(frame.y))
    ext.set("cx", str(frame.cx))
    ext.set("cy", str(frame.cy))
    geometry = properties.find(qn("a:prstGeom"))
    if geometry is None:
        return
    if not (geometry.get("prst") == "line" and frame.preset == "straightConnector1"):
        geometry.set("prst", frame.preset)
    av_list = subelement(geometry, "a:avLst")
    for node in list(av_list):
        remove(node)
    av_list.text = None
    for name, value in frame.adjustments:
        av_list.append(av_list.makeelement(qn("a:gd"), {"name": name, "fmla": f"val {value}"}))


def route_connector(slide: "Slide", element: Element) -> bool:
    """Re-compute one connector's geometry from its attachments.  ``False`` if it cannot be
    routed (an unknown preset, an end on a shape no longer on the slide)."""
    kind = connector_kind(element)
    if kind is None:
        return False
    parent = element.getparent()
    ends = current_ends(element)
    if ends is None:
        return False
    resolved = []
    for which, current in zip(("begin", "end"), ends):
        reference = connection(element, which)
        if reference is None:
            resolved.append(End(container_to_slide(parent, *current)))
            continue
        target = slide._element_by_raw_id(reference[0])
        sites = shape_sites(slide._wrap(target)) if target is not None else []
        if not 0 <= reference[1] < len(sites):
            resolved.append(End(container_to_slide(parent, *current)))
            continue
        resolved.append(sites[reference[1]])
    write_frame(element, _frame_in(slide, parent, kind, resolved[0], resolved[1]))
    return True


def raw_ids_within(element: Element) -> set[str]:
    """``cNvPr@id`` of a shape and, for a group, of everything in it."""
    return {node.get("id") for node in element.iter(qn("p:cNvPr")) if node.get("id")}


def attached_connectors(slide: "Slide", raw_ids: set[str]) -> list[Element]:
    found = []
    for element in slide._sp_tree().iter(qn("p:cxnSp")):
        for which in ("begin", "end"):
            reference = connection(element, which)
            if reference is not None and reference[0] in raw_ids:
                found.append(element)
                break
    return found
