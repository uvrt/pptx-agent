"""TEMPORARY (s3): this repository's copies of Office's built-in data are ooxml-common's.

Run once on the pull request that derives them from ``ooxml_common.drawingml.table_styles``
and ``ooxml_common.drawingml.theme`` instead; removed with the copies.
"""

from ooxml_common.drawingml import table_styles, theme

from pptx_agent.edit import blank
from pptx_agent.edit.table_text_colors import BUILTIN_TEXT_COLORS
from pptx_agent.edit.tablestyles import BUILTIN_TABLE_STYLES


def _spec(slot, transforms):
    return " ".join([slot] + [f"{kind}={value}" for kind, value in transforms])


def test_the_names_are_the_shared_catalogues():
    derived = {guid: spec["name"] for guid, spec in table_styles.BUILTIN_TABLE_STYLES.items()}
    assert BUILTIN_TABLE_STYLES == derived
    assert list(BUILTIN_TABLE_STYLES) == list(derived)


def test_the_text_colours_are_the_shared_catalogues():
    derived = {guid: {region: _spec(*color) for region, color
                      in table_styles.builtin_text_colors(guid).items()}
               for guid in table_styles.BUILTIN_TABLE_STYLES}
    assert BUILTIN_TEXT_COLORS == derived
    assert list(BUILTIN_TEXT_COLORS) == list(derived)
    assert all(list(BUILTIN_TEXT_COLORS[guid]) == list(derived[guid]) for guid in derived)


def test_the_theme_is_the_shared_one():
    assert blank.THEME_COLORS == theme.OFFICE_THEME_COLORS
    assert (blank.MAJOR_FONT, blank.MINOR_FONT) == (theme.OFFICE_MAJOR_FONT,
                                                    theme.OFFICE_MINOR_FONT)
    assert blank._APTOS_PANOSE == theme.APTOS_PANOSE
    assert blank._SCRIPT_FONTS == theme.POWERPOINT_SCRIPT_FONTS


def test_the_preset_names_are_the_shared_tables():
    from ooxml_common.drawingml.presets import PRESETS

    from pptx_agent.oxml.xml import PRESET_GEOMETRIES

    assert PRESET_GEOMETRIES == frozenset(PRESETS)


def test_the_format_scheme_is_the_shared_one():
    from lxml import etree

    root = etree.fromstring(blank.theme_xml())
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def stops(gradient):
        return tuple((int(stop.get("pos")), tuple(
            (child.tag.rpartition("}")[2], int(child.get("val"))) for child in stop[0]))
            for stop in gradient.iter(f"{a}gs"))

    fills = root.findall(f".//{a}fillStyleLst/{a}gradFill")
    assert tuple(stops(fill) for fill in fills) == theme.OFFICE_FILL_GRADIENTS
    background = root.find(f".//{a}bgFillStyleLst/{a}gradFill")
    assert stops(background) == theme.OFFICE_BACKGROUND_GRADIENT
    widths = tuple(int(line.get("w")) for line in root.iter(f"{a}ln"))
    assert widths == theme.POWERPOINT_LINE_WIDTHS
