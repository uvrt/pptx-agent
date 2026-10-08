"""Colours: hex and theme specs, modifiers, and verbatim round trips of whatever a deck uses."""

from __future__ import annotations

import pytest

from pptx_agent import Document
from pptx_agent.edit.color import Color
from pptx_agent.oxml.xml import qn


@pytest.mark.parametrize(
    "spec, kind, value, transforms",
    [
        ("4472C4", "rgb", "4472C4", ()),
        ("#4472c4", "rgb", "4472C4", ()),
        ("accent1", "scheme", "accent1", ()),
        ("tx1", "scheme", "tx1", ()),
        ("accent1 lumMod=75%", "scheme", "accent1", (("lumMod", 75000),)),
        ("accent2 lumMod=60000 lumOff=40000", "scheme", "accent2",
         (("lumMod", 60000), ("lumOff", 40000))),
        ("bg1 shade=50% alpha=80%", "scheme", "bg1", (("shade", 50000), ("alpha", 80000))),
    ],
)
def test_colour_specs_parse(spec, kind, value, transforms):
    color = Color.parse(spec)
    assert (color.kind, color.value, color.transforms) == (kind, value, transforms)
    assert Color.parse(str(color)) == color


@pytest.mark.parametrize("spec", ["accent9", "12345", "#GGGGGG", "accent1 glow=3", "accent1 lumMod=x"])
def test_bad_colour_specs_are_refused(spec):
    with pytest.raises(ValueError):
        Color.parse(spec)


def test_theme_keyword_form_matches_string_form():
    assert Color.theme("accent1", lum_mod=0.75, lum_off=0.25) == "accent1 lumMod=75% lumOff=25%"


def test_colours_read_from_a_file_round_trip_verbatim(pptx_path):
    """Whatever colour kinds a deck uses (sysClr, prstClr...), reading and rebuilding is exact."""
    from pptx_agent.oxml.xml import COLOR_TAGS

    def shape_of(element):
        return (element.tag, dict(element.attrib),
                [(child.tag, dict(child.attrib)) for child in element])

    document = Document.open(str(pptx_path))
    seen = 0
    for slide in document.slides:
        root = document.package.tree(slide.part_path)
        for tag in COLOR_TAGS:
            for element in root.iter(qn(tag)):
                assert shape_of(Color.from_element(element).to_element()) == shape_of(element)
                seen += 1
    if not seen:
        pytest.skip("no colours on this fixture's slides")


