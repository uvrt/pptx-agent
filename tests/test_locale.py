"""The regional settings a deck is drawn and edited in: ``Document.locale``, and through the
tool layer ``Toolbox(locale=...)`` -- to every render (in the worker process too), to new
charts, to the language of new text where the deck states none, and to a new date field's
cached text.  en-US unless given: the output without a locale is what it always was."""

from __future__ import annotations

import datetime as dt
import json
import re

import pytest

from ooxml_common.locales import EN_US, Locale
from ooxml_edit.tools import Toolbox

from pptx_agent import Document
from pptx_agent.edit.scales import DateScale, date_label
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

pptx2svg = pytest.importorskip("pptx2svg")

DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-NL", list_separator=";")
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _chart_deck(locale=None) -> Document:
    deck = Document.new(locale=locale)
    slide = deck.add_slide("Blank")
    slide.add_chart("column", ["Q1", "Q2"], [{"name": "North", "values": [0.4, 1.3]}],
                    457200, 457200, 8229600, 4572000)
    return deck


def _texts(svg: str) -> list[str]:
    return re.findall(r">([^<>]+)</t", svg)


def test_without_a_locale_a_deck_is_en_us():
    deck = Document.new()
    assert deck.locale is None and deck.language == "en-US"
    assert Document.new(locale=EN_US).to_bytes() == deck.to_bytes()


def test_a_new_deck_takes_the_locales_language():
    deck = Document.new(locale=DUTCH)
    assert deck.locale is DUTCH and deck.language == "nl-NL"
    assert Document.new(locale=DUTCH, language="de-DE").language == "de-DE"


def test_a_deck_stating_no_language_takes_the_locales():
    deck = Document.new()
    presentation = deck.package.tree(deck.package.presentation_part())
    for node in presentation.iter(f"{A}defRPr"):
        node.attrib.pop("lang", None)
    assert deck.language == "en-US"
    deck.locale = DUTCH
    assert deck.language == "nl-NL"
    shape = deck.add_slide("Blank").add_shape("rect", 0, 0, 914400, 914400, text="Hallo")
    assert {r.get("lang") for r in shape._element.iter(f"{A}rPr")} == {"nl-NL"}


def test_a_new_chart_takes_the_locale():
    deck = _chart_deck(DUTCH)
    (part,) = [p for p in deck.package.part_names if p.startswith("ppt/charts/chart")]
    assert deck.package.tree(part).find(f"{C}lang").get("val") == "nl-NL"


def test_a_render_takes_the_locale():
    english = _texts(_chart_deck().render_svg()[0])
    dutch = _texts(_chart_deck(DUTCH).render_svg()[0])
    assert "0.2" in english and "0,2" in dutch and "0.2" not in dutch
    # the PNG path and the agent view take it too
    deck = _chart_deck(DUTCH)
    assert deck.render_png(width=320) != _chart_deck(DUTCH).render_png(width=320,
                                                                       locale=EN_US)


def _footer_date(deck: Document) -> tuple[str, str, str]:
    deck.set_footer(None, slide_number=False, date="auto", today=dt.date(2027, 1, 5))
    slide = deck.slides[0]
    field = next(deck.package.tree(slide.part_path).iter(f"{A}fld"))
    return field.get("type"), field.find(f"{A}rPr").get("lang"), field.findtext(f"{A}t")


def _deck_with_slide(**options) -> Document:
    deck = Document.new(**options)
    deck.add_slide("Title and Content")
    return deck


def test_a_new_date_field_caches_what_powerpoint_draws():
    # Measured (pptx2svg's DATE_FORMATS): en-US datetime1 is M/d/yy, nl-NL dd-MM-yyyy.
    assert _footer_date(_deck_with_slide()) == ("datetime1", "en-US", "1/5/27")
    assert _footer_date(_deck_with_slide(language="nl-NL")) == ("datetime1", "nl-NL",
                                                               "05-01-2027")
    assert _footer_date(_deck_with_slide(locale=DUTCH)) == ("datetime1", "nl-NL", "05-01-2027")


def test_a_date_field_in_an_unmeasured_language_falls_back_as_pptx2svg_draws_it():
    # ja-JP has no table: the locale's language where it has one, else en-US.
    assert _footer_date(_deck_with_slide(language="ja-JP"))[2] == "1/5/27"
    assert _footer_date(_deck_with_slide(language="ja-JP", locale=DUTCH))[2] == "05-01-2027"


def test_date_scale_labels_name_their_months_themselves():
    time = DateScale("2026-11-01", "2027-03-31", 0, 151)
    assert [t.label for t in time.ticks(every="month")][:2] == ["Nov 2026", "Dec 2026"]
    assert [t.label for t in time.ticks(every="month", language="nl-NL")][2:5] == [
        "jan 2027", "feb 2027", "mrt 2027"]
    assert [t.label for t in time.ticks(every="month", language="ja-JP")][0] == "Nov 2026"
    day = dt.date(2026, 3, 1)
    for pattern in ("%b %Y", "%B %d", "%a %A %p", "%c", "%x", "%X", "100%% %b", "%h"):
        assert date_label(day, pattern) == {
            "%b %Y": "Mar 2026", "%B %d": "March 01", "%a %A %p": "Sun Sunday AM",
            "%c": "Sun Mar  1 00:00:00 2026", "%x": "03/01/26", "%X": "00:00:00",
            "100%% %b": "100% Mar", "%h": "Mar"}[pattern]
    assert date_label(day, "%A %B", "de-DE") == "Sonntag März"


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _render(box: Toolbox, data: bytes, **session_options) -> bytes:
    session = box.session(**session_options)
    session.open(data, "deck.pptx")
    result = box.dispatch(session, "render", {"doc": "d1", "slides": [1], "width": 320})
    assert result.ok, result.error
    return result.images[0].data


def test_the_toolbox_locale_reaches_the_deck_and_the_worker():
    data = _chart_deck().to_bytes()
    english = Document.open(data).render_png([1], width=320)[0]
    dutch_deck = Document.open(data)
    dutch_deck.locale = DUTCH
    dutch = dutch_deck.render_png([1], width=320)[0]
    assert english != dutch
    with _box(workers=1) as plain:
        assert _render(plain, data) == english
    with _box(workers=1, locale=DUTCH) as configured:
        session = configured.session()
        session.open(data, "deck.pptx")
        assert session.entry("d1").document.locale is DUTCH
        assert _render(configured, data) == dutch               # in the worker, in Dutch
        assert _render(configured, data, locale=EN_US) == english
    with _box(workers=0, locale=DUTCH) as in_process:
        assert _render(in_process, data) == dutch


def test_a_subprocess_worker_gets_the_locale_from_its_arguments():
    """What a daemonic process renders with: the locale crosses as an argument."""
    from ooxml_edit.tools import SubprocessPool

    data = _chart_deck().to_bytes()
    dutch_deck = Document.open(data)
    dutch_deck.locale = DUTCH
    with _box(runner=SubprocessPool(1), locale=DUTCH) as box:
        assert _render(box, data) == dutch_deck.render_png([1], width=320)[0]


def test_new_document_takes_the_session_locale():
    with _box(locale=DUTCH) as box:
        session = box.session()
        made = box.dispatch(session, "new_document", {"kind": "pptx"})
        assert made.ok, made.error
        deck = session.entry("d1").document
        assert deck.locale is DUTCH and deck.language == "nl-NL"


def test_the_definitions_and_prompt_do_not_change():
    with _box() as plain, _box(locale=DUTCH) as configured:
        for provider in ("anthropic", "openai-responses"):
            assert json.dumps(plain.definitions(provider)) == \
                json.dumps(configured.definitions(provider))
        assert plain.system_prompt() == configured.system_prompt()
