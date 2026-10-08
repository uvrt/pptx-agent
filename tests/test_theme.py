"""The theme, resolved colours and formatting as drawn (the trial's finding 10): five runs
rendered swatches or measured pixels to learn what an accent looks like or what size the
inherited text is."""

from __future__ import annotations

import warnings

import pytest

from conftest import FIXTURE_DIR, fixture_paths
from pptx_agent import Color, Document, EffectiveParagraph, Theme, ThemeFonts
from pptx_agent.oxml.xml import make, qn

PILOT = FIXTURE_DIR / "generated" / "trial" / "pilot-retrospective.pptx"


def _pilot() -> Document:
    return Document.open(PILOT)


def test_the_theme_has_every_scheme_colour_and_its_fonts():
    theme = _pilot().theme
    assert isinstance(theme, Theme) and theme.name == "Office Theme"
    assert theme.colors["accent1"] == "#156082"
    assert theme.colors["tx1"] == theme.colors["dk1"] == "#000000"
    assert theme.colors["bg1"] == theme.colors["lt1"] == "#FFFFFF"
    assert set(theme.colors) == {"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                                 "accent4", "accent5", "accent6", "hlink", "folHlink", "bg1",
                                 "tx1", "bg2", "tx2"}
    assert theme.fonts == ThemeFonts("Aptos Display", "Aptos")
    assert (theme.fonts.major, theme.fonts.minor) == ("Aptos Display", "Aptos")


def test_a_colour_resolves_as_pptx2svg_draws_it():
    """The resolution is the renderer's: a fill set here is the fill pptx2svg draws."""
    pptx2svg = pytest.importorskip("pptx2svg")
    deck = _pilot()
    slide = deck.slides[0]
    specs = ["accent1", "accent2 lumMod=75%", "accent1 lumMod=60% lumOff=40%", "tx1",
             "bg2 shade=50%", "#336699 tint=50%"]
    shapes = [slide.add_shape("rect", 100000 * k, 0, 90000, 90000, fill=spec)
              for k, spec in enumerate(specs)]
    model = pptx2svg.convert_pptx_to_model(
        deck.to_bytes(), pptx2svg.ConvertOptions(slide_numbers=[1]))
    drawn = {e.element_id: e.fill.color.hex.upper() for e in model.slides[0].elements
             if getattr(e, "fill", None) is not None and hasattr(e.fill, "color")}
    for spec, shape in zip(specs, shapes):
        assert Color.parse(spec).resolve(shape) == drawn[shape.id], spec
        assert shape.fill.color.resolve(deck) == drawn[shape.id], spec
    assert Color.parse("#1F4E79").resolve(deck) == "#1F4E79"
    assert deck.theme.resolve("accent1") == "#156082"


def test_a_slide_override_of_the_colour_map_is_followed():
    deck = _pilot()
    slide = deck.slides[0]
    root = deck.package.tree(slide.part_path)
    override = root.find(qn("p:clrMapOvr"))
    for child in list(override):
        override.remove(child)
    override.append(make("a:overrideClrMapping", bg1="dk1", tx1="lt1", bg2="dk2", tx2="lt2",
                         accent1="accent1", accent2="accent2", accent3="accent3",
                         accent4="accent4", accent5="accent5", accent6="accent6",
                         hlink="hlink", folHlink="folHlink"))
    assert slide.theme.colors["bg1"] == "#000000"
    assert Color.parse("tx1").resolve(slide.shapes[0]) == "#FFFFFF"
    assert Color.parse("tx1").resolve(deck) == "#000000"      # the deck's own map


def test_effective_sizes_and_fonts_follow_the_inheritance():
    deck = _pilot()
    body = deck.shape("258.3").text_frame
    topic, point = body.paragraph(0), body.paragraph(1)
    assert topic.runs[0].size is None and topic.runs[0].effective_size == 28.0
    assert point.runs[0].effective_size == 24.0
    assert point.runs[0].effective_font == "Aptos"
    title = deck.shape("258.2").text_frame.paragraph(0).run(0)
    assert (title.effective_size, title.effective_font) == (44.0, "Aptos Display")
    effective = point.effective
    assert isinstance(effective, EffectiveParagraph)
    assert (effective.level, effective.size, effective.alignment) == (1, 24.0, "left")
    assert effective.bullet.char == "•"
    assert effective.margin_left == 685800 and effective.line_spacing == 0.9
    point.runs[0].size = 20
    assert point.runs[0].effective_size == 20.0


def test_a_plain_text_box_without_a_named_face_is_arial():
    deck = Document.open(FIXTURE_DIR / "powerpoint-smartart.pptx")
    run = deck.shape("256.2").text_frame.paragraph(0).run(0)
    assert (run.effective_size, run.effective_font) == (28.0, "Arial")


def test_cells_and_notes_have_effective_sizes_too():
    deck = Document.new()
    slide = deck.add_slide("Title Only")
    table = slide.add_table(2, 2, 914400, 1828800, 3657600, 741680).table
    table.cell(0, 0).text = "Head"
    assert table.cell(0, 0).paragraphs[0].runs[0].effective_size == 18.0
    slide.notes = "Say it"
    assert slide.notes_frame.paragraph(0).run(0).effective_size == 12.0


@pytest.mark.parametrize("path", fixture_paths() + [PILOT], ids=lambda p: p.stem)
def test_effective_formatting_is_what_pptx2svg_resolves(path):
    """Every paragraph's first run, on every fixture: the size and face here are the ones
    the renderer resolves on its own."""
    pptx2svg = pytest.importorskip("pptx2svg")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        deck = Document.open(path)
    model = pptx2svg.convert_pptx_to_model(
        deck.to_bytes(), pptx2svg.ConvertOptions(warn_on_font_substitution=False,
                                                 use_embedded_fonts=False))
    checked = 0
    for slide, resolved in zip(deck.slides, model.slides):
        drawn = {}
        stack = list(resolved.elements)
        while stack:
            element = stack.pop(0)
            if getattr(element, "type", None) == "group":
                stack = list(element.children) + stack
            elif getattr(element, "text_body", None) is not None and element.element_id:
                drawn.setdefault(element.element_id, element)
        for shape in slide.shapes:
            element = drawn.get(shape.id.split("#")[0])
            if shape.kind != "shape" or element is None:
                continue
            scale = shape.text_frame.font_scale
            for paragraph, other in zip(shape.text_frame.paragraphs,
                                        element.text_body.paragraphs):
                runs = [r for r in paragraph.runs if r.text]
                others = [r for r in other.runs if r.text]
                if not runs or not others:
                    continue
                assert runs[0].effective_size == pytest.approx(
                    (others[0].properties.font_size or 18.0) * scale, abs=0.01)
                assert runs[0].effective_font == others[0].properties.font_family
                checked += 1
    assert checked or path.stem in ("powerpoint-smartart",)


def test_autofit_reads_what_applies_and_writes_the_frame_own():
    deck = _pilot()
    frame = deck.shape("258.3").text_frame
    assert (frame.autofit, frame.font_scale) == ("normal", 1.0)     # inherited from the master
    frame.font_scale = 0.8
    assert (frame.autofit, frame.font_scale) == ("normal", 0.8)
    assert frame.paragraph(1).runs[0].effective_size == pytest.approx(19.2)
    frame.autofit = "normal"                                        # keeps the stored scale
    assert frame.font_scale == 0.8
    frame.autofit = "none"
    assert (frame.autofit, frame.font_scale) == ("none", 1.0)
    frame.autofit = "shape"
    assert frame.autofit == "shape"
    frame.autofit = None                                            # inherit again
    assert frame.autofit == "normal"
    for _ in range(5):
        deck.undo()
    assert (deck.shape("258.3").text_frame.autofit,
            deck.shape("258.3").text_frame.font_scale) == ("normal", 1.0)
    with pytest.raises(ValueError):
        frame.autofit = "shrink"
    with pytest.raises(ValueError):
        frame.set_autofit("none", font_scale=0.5)
    assert frame.set_autofit("normal", font_scale=0.9) is frame
    assert deck.validate() == []


# -- changing the theme, and design guidance as data (the trial's N17) -------------------------


def test_set_colors_changes_what_every_reference_draws():
    from pptx_agent import Document

    deck = Document.new()
    box = deck.add_slide("Blank").add_shape("rect", 0, 0, 914400, 914400)
    box.fill = "accent1 lumMod=75%"
    theme = deck.theme.set_colors({"accent1": "#1F4E79", "tx2": "0b2545"})
    assert theme.colors["accent1"] == "#1F4E79" and theme.colors["dk2"] == "#0B2545"
    assert deck.theme.colors["tx2"] == "#0B2545"
    assert str(box.fill.color).startswith("accent1 lumMod")      # still a reference
    assert box.fill.color.resolve(box) == deck.theme.resolve("accent1 lumMod=75%")
    assert deck.validate() == []
    deck.undo()
    assert deck.theme.colors["accent1"] != "#1F4E79"


def test_set_colors_refuses_what_is_not_a_slot_or_rgb():
    from pptx_agent import Document

    deck = Document.new()
    with pytest.raises(ValueError, match="accent1"):
        deck.theme.set_colors({"brand": "#000000"})
    with pytest.raises(ValueError, match="RRGGBB"):
        deck.theme.set_colors({"accent1": "navy"})


def test_set_fonts_changes_the_faces_text_draws_in():
    from pptx_agent import Document

    deck = Document.new()
    slide = deck.add_slide("Title and Content")
    slide.shapes[0].set_text("Plan")
    theme = deck.theme.set_fonts(major="Georgia", minor="Arial")
    assert (theme.fonts.major, theme.fonts.minor) == ("Georgia", "Arial")
    run = slide.shapes[0].text_frame.paragraph(0).run(0)
    assert run.effective_font == "Georgia"
    latin = deck.package.tree("/ppt/theme/theme1.xml").find(".//{*}majorFont/{*}latin")
    assert latin.get("panose") is None
    deck.theme.set_fonts(minor_east_asian="Meiryo")
    assert deck.theme.fonts.minor_east_asian == "Meiryo" and deck.theme.fonts.major == "Georgia"


def test_roles_follow_the_template_and_fall_back_to_the_convention():
    from pptx_agent import Document

    roles = Document.new().theme.roles
    assert (roles.primary, roles.secondary, roles.neutral, roles.highlight) == \
        ("dk2", "accent1", "lt2", "accent2")
    assert set(roles.source.values()) == {"convention"}
    basic = Document.open(FIXTURE_DIR / "real-basic-theme.pptx").theme.roles
    assert (basic.primary, basic.secondary, basic.highlight) == ("accent1", "dk2", "accent3")
    assert basic.source["primary"] == "master" and basic.usage["accent1"] > basic.usage["dk2"]
    assert basic.colors["primary"] == Document.open(
        FIXTURE_DIR / "real-basic-theme.pptx").theme.colors["accent1"]


def test_tints_are_theme_references_lightest_first():
    from pptx_agent import Document

    theme = Document.new().theme
    tints = theme.tints("accent1")
    assert [tint.color for tint in tints] == [
        "accent1 lumMod=20% lumOff=80%", "accent1 lumMod=40% lumOff=60%",
        "accent1 lumMod=60% lumOff=40%", "accent1 lumMod=75%", "accent1 lumMod=50%"]
    assert tints[0].name == "Lighter 80%" and tints[0].hex == theme.resolve(tints[0].color)
    assert theme.tints("dk2")[0].name == "Lighter 90%"            # a dark colour
    assert theme.tints("lt1")[0].name == "Darker 10%"             # a light one
    assert set(theme.ramps) == {"dk2", "lt2", "accent1", "accent2", "accent3", "accent4",
                                "accent5", "accent6"}
    with pytest.raises(ValueError):
        theme.tints("brand")
