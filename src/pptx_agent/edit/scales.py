"""Data-to-position scales: numbers, dates and categories mapped onto a length.

A chart-like diagram -- a Gantt, a scatter of bubbles, a 2x2 matrix, a timeline -- is
shapes placed at data values.  Working those positions out by hand is where agents went
wrong (the spike's arm C: proportional placement and label distances), so a scale is
declared once and positions are asked of it::

    from pptx_agent.edit.scales import DateScale, LinearScale

    time = DateScale("2026-11-01", "2027-03-31", 160, 920)     # inclusive dates -> points
    time.position("2027-01-22", "center")                      # the middle of that day
    score = LinearScale(0, 10, 540, 140)                       # 0 at the bottom, 10 at top
    score.position(7)

A scale knows no axis: the same one serves x or y, and a reversed range (``540 -> 140``)
runs upwards.  It knows no unit either; the tools give points, ``ppt_draw`` its SVG's user
units.  Three kinds:

* :class:`LinearScale` -- numbers, ``min`` to ``max``;
* :class:`DateScale` -- whole days, ``start`` to ``end`` **inclusive**, optionally
  without some excluded ranges (a holiday break closes up); a date is a day with a start,
  a centre and an end, so a bar from its first day's start to its last day's end covers
  both days;
* :class:`BandScale` -- named categories, each a band of equal width, with a gap between
  bands and padding at both ends; a band too has a start, centre and end.

``where`` -- ``"start"``, ``"center"`` or ``"end"`` -- picks the part of a day or band a
position means (a number is a point, so it ignores it).  :meth:`Scale.ticks` lists tick
or gridline positions: round numbers, the days, weeks, months, quarters or years in the
range, or the bands, each with a label.

:func:`scale_from` builds a scale from a plain mapping (what the tools store) and
:meth:`Scale.to_json` gives that mapping back.
"""

from __future__ import annotations

import datetime as _dt
import math
from dataclasses import dataclass
from typing import Any, Iterable, Mapping, Sequence

#: The parts of a day or a band a position may mean.
WHERE = ("start", "center", "end")
#: Date tick intervals.
DATE_UNITS = ("day", "week", "month", "quarter", "year")


class ScaleError(ValueError):
    """A scale or a value that does not make sense; the message says what would."""


def as_date(value: Any, *, field: str = "date") -> _dt.date:
    """A date from ``YYYY-MM-DD`` (or a date); :class:`ScaleError` names the field."""
    if isinstance(value, _dt.datetime):
        return value.date()
    if isinstance(value, _dt.date):
        return value
    try:
        return _dt.date.fromisoformat(str(value).strip()[:10])
    except ValueError:
        raise ScaleError(f"{field}: {value!r} is not a date YYYY-MM-DD") from None


def as_number(value: Any, *, field: str = "value") -> float:
    """A finite number from a number or its text."""
    try:
        number = float(str(value).strip()) if isinstance(value, str) else float(value)
    except (TypeError, ValueError):
        raise ScaleError(f"{field}: {value!r} is not a number") from None
    if not math.isfinite(number):
        raise ScaleError(f"{field}: {value!r} is not a finite number")
    return number


def _where(where: str | None) -> str:
    where = where or "start"
    if where not in WHERE:
        raise ScaleError(f"where must be one of {', '.join(WHERE)}, not {where!r}")
    return where


@dataclass(frozen=True)
class Tick:
    """One tick: its value (as text), label, and positions.  For a number ``start``,
    ``center`` and ``end`` are the same point; for a day, a period or a band they are its
    edges and middle."""

    value: str
    label: str
    start: float
    center: float
    end: float

    def to_json(self, digits: int = 2) -> dict[str, Any]:
        data: dict[str, Any] = {"value": self.value, "label": self.label}
        if self.start == self.end:
            data["at"] = round(self.center, digits)
        else:
            data.update(start=round(self.start, digits), center=round(self.center, digits),
                        end=round(self.end, digits))
        return data


class Scale:
    """A mapping from data values onto ``start`` .. ``end`` (any length unit)."""

    kind = "scale"

    def __init__(self, start: float, end: float) -> None:
        self.start = as_number(start, field="from")
        self.end = as_number(end, field="to")
        if self.start == self.end:
            raise ScaleError("a scale's range has no length: from and to are equal")

    @property
    def length(self) -> float:
        return self.end - self.start

    def position(self, value: Any, where: str | None = None) -> float:  # pragma: no cover
        raise NotImplementedError

    def ticks(self, **options: Any) -> list[Tick]:  # pragma: no cover
        raise NotImplementedError

    def to_json(self) -> dict[str, Any]:  # pragma: no cover
        raise NotImplementedError

    def describe(self) -> str:  # pragma: no cover
        raise NotImplementedError

    def _at(self, fraction: float) -> float:
        return self.start + fraction * self.length


class LinearScale(Scale):
    """Numbers ``low`` .. ``high`` onto ``start`` .. ``end``; values outside extrapolate."""

    kind = "linear"

    def __init__(self, low: float, high: float, start: float, end: float) -> None:
        super().__init__(start, end)
        self.low = as_number(low, field="min")
        self.high = as_number(high, field="max")
        if self.low == self.high:
            raise ScaleError("min and max are equal: the domain has no width")

    def position(self, value: Any, where: str | None = None) -> float:
        _where(where)
        number = as_number(value)
        return self._at((number - self.low) / (self.high - self.low))

    def ticks(self, *, step: float | None = None, count: int = 6, **_: Any) -> list[Tick]:
        """Ticks at multiples of ``step`` within the domain (ends included when they are
        multiples), or about ``count`` round steps (1, 2 or 5 times a power of ten)."""
        low, high = sorted((self.low, self.high))
        if step is None:
            step = nice_step(high - low, count)
        step = abs(as_number(step, field="step"))
        if step == 0 or (high - low) / step > 1000:
            raise ScaleError(f"step {step:g} gives too many ticks for {low:g}..{high:g}")
        first = math.ceil(low / step - 1e-9)
        out = []
        index = first
        while index * step <= high + step * 1e-9:
            value = round(index * step, 10)
            text = _number_text(value)
            point = self.position(value)
            out.append(Tick(text, text, point, point, point))
            index += 1
        if self.low > self.high:
            out.reverse()
        return out

    def to_json(self) -> dict[str, Any]:
        return {"kind": self.kind, "min": self.low, "max": self.high, "from": self.start,
                "to": self.end}

    def describe(self) -> str:
        return (f"linear {_number_text(self.low)}..{_number_text(self.high)} -> "
                f"{_number_text(round(self.start, 2))}..{_number_text(round(self.end, 2))}")


class DateScale(Scale):
    """Whole days ``first`` .. ``last`` (inclusive) onto ``start`` .. ``end``.

    ``exclude`` lists inclusive ``(from, to)`` date ranges taken out of the scale: they
    close up, so the days either side meet, and a date inside one sits at the join.
    """

    kind = "date"

    def __init__(self, first: Any, last: Any, start: float, end: float, *,
                 exclude: Iterable[tuple[Any, Any]] = ()) -> None:
        super().__init__(start, end)
        self.first = as_date(first, field="start")
        self.last = as_date(last, field="end")
        if self.last < self.first:
            raise ScaleError(f"end {self.last} is before start {self.first}")
        ranges = []
        for low, high in exclude:
            a, b = as_date(low, field="exclude.from"), as_date(high, field="exclude.to")
            if b < a:
                raise ScaleError(f"exclude {a}..{b}: to is before from")
            a, b = max(a, self.first), min(b, self.last)
            if a <= b:
                ranges.append((a, b))
        ranges.sort()
        merged: list[tuple[_dt.date, _dt.date]] = []
        for a, b in ranges:
            if merged and a <= merged[-1][1] + _dt.timedelta(days=1):
                merged[-1] = (merged[-1][0], max(merged[-1][1], b))
            else:
                merged.append((a, b))
        self.exclude = tuple(merged)
        self.days = (self.last - self.first).days + 1 - sum((b - a).days + 1 for a, b in merged)
        if self.days <= 0:
            raise ScaleError("every day of the scale is excluded")

    def _offset(self, day: _dt.date, part: float) -> float:
        """Days from the scale's start to ``part`` (0..1) through ``day``, without the
        excluded days."""
        t = (day - self.first).days + part
        removed = 0
        for a, b in self.exclude:
            lo = (a - self.first).days
            hi = (b - self.first).days + 1
            if t >= hi:
                removed += hi - lo
            elif t > lo:
                return lo - removed
        return t - removed

    def position(self, value: Any, where: str | None = None) -> float:
        day = as_date(value, field="value")
        part = {"start": 0.0, "center": 0.5, "end": 1.0}[_where(where)]
        return self._at(self._offset(day, part) / self.days)

    def ticks(self, *, every: str = "month", step: float | None = None,
              format: str | None = None, **_: Any) -> list[Tick]:
        """One tick per ``every`` (day, week from Monday, month, quarter, year) -- each
        ``step`` of them (default 1) -- that overlaps the scale, clipped to it: its first
        day's start, its middle and its last day's end.  ``format`` is a ``strftime``
        pattern for the label (``{q}`` is the quarter's number, ``{d}`` the day without a
        leading zero)."""
        if every not in DATE_UNITS:
            raise ScaleError(f"every must be one of {', '.join(DATE_UNITS)}, not {every!r}")
        stride = int(step or 1)
        if stride < 1:
            raise ScaleError("step must be 1 or more")
        pattern = format or {"day": "{d} %b", "week": "{d} %b", "month": "%b %Y",
                             "quarter": "Q{q} %Y", "year": "%Y"}[every]
        periods = list(_periods(self.first, self.last, every))[::stride]
        if len(periods) > 1000:
            raise ScaleError(f"{len(periods)} ticks: use a longer interval")
        out = []
        for begin, finish in periods:
            if stride > 1:
                finish = min(_add(begin, every, stride) - _dt.timedelta(days=1), self.last)
            a, b = max(begin, self.first), min(finish, self.last)
            left = self.position(a, "start")
            right = self.position(b, "end")
            label = begin.strftime(pattern.replace("{q}", str((begin.month - 1) // 3 + 1))
                                   .replace("{d}", str(begin.day)))
            out.append(Tick(begin.isoformat(), label, left, (left + right) / 2, right))
        return out

    def to_json(self) -> dict[str, Any]:
        data: dict[str, Any] = {"kind": self.kind, "start": self.first.isoformat(),
                                "end": self.last.isoformat(), "from": self.start,
                                "to": self.end}
        if self.exclude:
            data["exclude"] = [{"from": a.isoformat(), "to": b.isoformat()}
                               for a, b in self.exclude]
        return data

    def describe(self) -> str:
        text = (f"date {self.first}..{self.last} ({self.days} days) -> "
                f"{_number_text(round(self.start, 2))}..{_number_text(round(self.end, 2))}")
        if self.exclude:
            text += " without " + ", ".join(f"{a}..{b}" for a, b in self.exclude)
        return text


class BandScale(Scale):
    """Named bands of equal width across ``start`` .. ``end``, in order, ``gap`` apart,
    with ``padding`` before the first and after the last."""

    kind = "band"

    def __init__(self, bands: Sequence[str], start: float, end: float, *, gap: float = 0,
                 padding: float = 0) -> None:
        super().__init__(start, end)
        self.bands = [str(band) for band in bands]
        if not self.bands:
            raise ScaleError("a band scale needs at least one band")
        if len(set(self.bands)) != len(self.bands):
            raise ScaleError("band names repeat: each band needs its own name")
        self.gap = as_number(gap, field="gap")
        self.padding = as_number(padding, field="padding")
        if self.gap < 0 or self.padding < 0:
            raise ScaleError("gap and padding are lengths, 0 or more")
        count = len(self.bands)
        self.width = (abs(self.length) - 2 * self.padding - (count - 1) * self.gap) / count
        if self.width <= 0:
            raise ScaleError(f"{count} bands with gap {self.gap:g} and padding "
                             f"{self.padding:g} leave no room in {abs(self.length):g}")

    def index(self, value: Any) -> int:
        text = str(value)
        if text in self.bands:
            return self.bands.index(text)
        folded = [band.casefold() for band in self.bands]
        if text.casefold() in folded:
            return folded.index(text.casefold())
        if text.isdigit() and int(text) < len(self.bands):
            return int(text)
        raise ScaleError(f"{value!r} is not a band; the bands are {self.bands}")

    def position(self, value: Any, where: str | None = None) -> float:
        index = self.index(value)
        sign = 1 if self.length > 0 else -1
        begin = self.start + sign * (self.padding + index * (self.width + self.gap))
        offset = {"start": 0.0, "center": 0.5, "end": 1.0}[_where(where)] * self.width
        return begin + sign * offset

    def ticks(self, **_: Any) -> list[Tick]:
        """One per band: its start, centre and end."""
        return [Tick(band, band, self.position(band, "start"), self.position(band, "center"),
                     self.position(band, "end")) for band in self.bands]

    def to_json(self) -> dict[str, Any]:
        return {"kind": self.kind, "bands": list(self.bands), "from": self.start,
                "to": self.end, "gap": self.gap, "padding": self.padding}

    def describe(self) -> str:
        return (f"band {len(self.bands)} x {_number_text(round(self.width, 2))} "
                f"(gap {_number_text(self.gap)}) -> {_number_text(round(self.start, 2))}.."
                f"{_number_text(round(self.end, 2))}")


def scale_from(data: Mapping[str, Any]) -> Scale:
    """A scale from its mapping (:meth:`Scale.to_json`'s form): ``kind`` and ``from``/``to``,
    plus ``min``/``max`` (linear), ``start``/``end``/``exclude`` (date) or
    ``bands``/``gap``/``padding`` (band)."""
    kind = data.get("kind")
    for key in ("from", "to"):
        if data.get(key) is None:
            raise ScaleError(f"a scale needs {key}: the position its domain "
                             f"{'starts' if key == 'from' else 'ends'} at")
    if kind == "linear":
        for key in ("min", "max"):
            if data.get(key) is None:
                raise ScaleError(f"a linear scale needs {key}")
        return LinearScale(data["min"], data["max"], data["from"], data["to"])
    if kind == "date":
        for key in ("start", "end"):
            if data.get(key) is None:
                raise ScaleError(f"a date scale needs {key} (YYYY-MM-DD, inclusive)")
        exclude = [(item.get("from"), item.get("to")) for item in data.get("exclude") or ()]
        return DateScale(data["start"], data["end"], data["from"], data["to"],
                         exclude=exclude)
    if kind == "band":
        if not data.get("bands"):
            raise ScaleError("a band scale needs bands: the names, in order")
        return BandScale(data["bands"], data["from"], data["to"], gap=data.get("gap") or 0,
                         padding=data.get("padding") or 0)
    raise ScaleError(f"kind must be linear, date or band, not {kind!r}")


def nice_step(span: float, count: int = 6) -> float:
    """A round step (1, 2 or 5 times a power of ten) giving about ``count`` intervals."""
    if span <= 0:
        return 1.0
    raw = span / max(1, count)
    power = 10 ** math.floor(math.log10(raw))
    for factor in (1, 2, 5, 10):
        if raw <= factor * power * 1.0000001:
            return factor * power
    return 10 * power


def _number_text(value: float) -> str:
    return f"{value:.10g}" if value != int(value) else str(int(value))


def _add(day: _dt.date, unit: str, count: int = 1) -> _dt.date:
    if unit == "day":
        return day + _dt.timedelta(days=count)
    if unit == "week":
        return day + _dt.timedelta(weeks=count)
    months = {"month": 1, "quarter": 3, "year": 12}[unit] * count
    index = day.year * 12 + day.month - 1 + months
    return _dt.date(index // 12, index % 12 + 1, 1)


def _periods(first: _dt.date, last: _dt.date, unit: str):
    """``(first day, last day)`` of every ``unit`` period overlapping ``first..last``."""
    if unit == "day":
        begin = first
    elif unit == "week":
        begin = first - _dt.timedelta(days=first.weekday())
    elif unit == "month":
        begin = first.replace(day=1)
    elif unit == "quarter":
        begin = _dt.date(first.year, (first.month - 1) // 3 * 3 + 1, 1)
    else:
        begin = _dt.date(first.year, 1, 1)
    while begin <= last:
        following = _add(begin, unit)
        yield begin, following - _dt.timedelta(days=1)
        begin = following


__all__ = ["BandScale", "DATE_UNITS", "DateScale", "LinearScale", "Scale", "ScaleError",
           "Tick", "WHERE", "as_date", "as_number", "nice_step", "scale_from"]
