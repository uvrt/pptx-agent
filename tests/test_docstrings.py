"""Every public class and method an agent reaches for has an example in its docstring.

The end-to-end pilot learned the API from ``help()`` and ``dir()``; an example is what makes
a signature usable.  Internals an agent does not need (part names, the raw data model) are
listed as exceptions.
"""

from __future__ import annotations

import inspect

import pytest

import pptx_agent

CLASSES = ["Document", "Slide", "Shape", "TextFrame", "Paragraph", "Run", "Table", "TableCell",
           "Chart", "Series", "SeriesList", "Diagram", "Fill", "LineFormat", "Color", "Layout",
           "Adjustments", "Hyperlink", "Bullet", "GradientStop", "Arrowhead", "Region",
           "ShapeId", "TextBlock", "ApplyReport", "LabelError", "MarkdownEscapeWarning",
           "TemplateOpenedWarning", "Placeholder", "LayoutPlaceholder", "LayoutShape",
           "Theme", "ThemeFonts", "EffectiveParagraph", "TextFit", "Overflow", "ImageSize"]
#: Plumbing, or inherited from list: no example needed.
EXEMPT = {"Chart.address", "Chart.part", "Chart.point_count", "Diagram.address",
          "Diagram.drawing_part", "Diagram.layout", "Diagram.model", "Diagram.part"} | {
    f"SeriesList.{name}" for name in dir(list) if not name.startswith("_")} | {
    f"{cls}.{name}" for cls in ("Placeholder", "ThemeFonts", "ImageSize") for name in dir(tuple)
    if not name.startswith("_")}


def _members(cls):
    for name in dir(cls):
        if name.startswith("_"):
            continue
        raw = inspect.getattr_static(cls, name)
        if callable(raw) or isinstance(raw, (property, classmethod, staticmethod)):
            if issubclass(cls, BaseException) and hasattr(Exception, name):
                continue
            yield name, raw


@pytest.mark.parametrize("name", CLASSES)
def test_public_api_has_examples(name):
    cls = getattr(pptx_agent, name)
    missing = [] if "::" in (cls.__doc__ or "") else [name]
    for member, raw in _members(cls):
        key = f"{name}.{member}"
        if key not in EXEMPT and "::" not in (inspect.getdoc(raw) or ""):
            missing.append(key)
    assert not missing
