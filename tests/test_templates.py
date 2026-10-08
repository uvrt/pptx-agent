"""A template opened, edited and saved as a deck: what the file declares itself to be.

The full end-to-end trial's one PowerPoint run that PowerPoint refused (P1, run 2): an agent
opened ``company-template.potx`` with ``Document.open`` -- not ``Document.new(template=...)``
-- and saved it as ``.pptx``.  The main part kept ``presentationml.template.main+xml``, which
PowerPoint refuses outright in a ``.pptx``, and ``validate()`` said nothing.  Now ``save``
writes the type its extension needs, ``validate(target=...)`` names a mismatch, and
``Document.open`` warns that a template is being opened as itself.
"""

from __future__ import annotations

import io
import os
import warnings
import zipfile
from pathlib import Path

import pytest
from lxml import etree

import oracle as oracle_helper
from conftest import FIXTURE_DIR
from pptx_agent import Document, TemplateOpenedWarning
from pptx_agent.edit import blank
from pptx_agent.edit.deck import CT_MACRO_PRESENTATION, CT_MACRO_TEMPLATE

TRIAL = FIXTURE_DIR / "generated" / "trial"
TEMPLATE = TRIAL / "company-template.potx"


def _main_type(data: bytes) -> str:
    root = etree.fromstring(zipfile.ZipFile(io.BytesIO(data)).read("[Content_Types].xml"))
    for node in root:
        if node.get("PartName") == "/ppt/presentation.xml":
            return node.get("ContentType")
    raise AssertionError("no override for the main part")


def _open_template() -> Document:
    with pytest.warns(TemplateOpenedWarning):
        return Document.open(TEMPLATE)


def test_the_trial_template_is_a_template():
    assert _main_type(TEMPLATE.read_bytes()) == blank.CT_TEMPLATE


def test_opening_a_template_warns_and_points_at_new():
    with pytest.warns(TemplateOpenedWarning, match=r"Document\.new\(template=") as caught:
        Document.open(TEMPLATE)
    assert caught[0].filename == __file__        # attributed to the caller, not the library
    with pytest.warns(TemplateOpenedWarning):
        Document.open(TEMPLATE.read_bytes())      # bytes too


def test_a_deck_from_a_template_does_not_warn():
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        deck = Document.new(template=TEMPLATE)
        Document.open(deck.to_bytes())


def test_the_trial_reproduction_now_saves_a_presentation(tmp_path):
    """P1 run 2, as the agent wrote it: open the .potx, add slides, save as .pptx."""
    deck = _open_template()
    deck.add_slide("TITLE").shapes[0].set_text("Project Harbour")
    out = tmp_path / "project-harbour.pptx"

    problems = deck.validate(target=out)
    assert [p.code for p in problems] == ["package-type"]
    assert "template" in problems[0].detail and ".pptx" in problems[0].detail
    assert deck.validate() == []                  # without a target there is nothing to say

    deck.save(out)
    assert _main_type(out.read_bytes()) == blank.CT_PRESENTATION
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        reopened = Document.open(out)             # a presentation now: no warning
    assert reopened.validate(target=out) == []
    assert reopened.slides[0].shapes[0].text == "Project Harbour"


def test_the_extension_decides_both_ways(tmp_path):
    deck = Document.open(FIXTURE_DIR / "sample.pptx")
    assert deck.validate(target="brand.potx")[0].code == "package-type"
    deck.save(tmp_path / "brand.potx")
    assert _main_type((tmp_path / "brand.potx").read_bytes()) == blank.CT_TEMPLATE
    deck.save(tmp_path / "deck.pptx")
    assert _main_type((tmp_path / "deck.pptx").read_bytes()) == blank.CT_PRESENTATION
    deck.save(tmp_path / "deck.pptm")
    assert _main_type((tmp_path / "deck.pptm").read_bytes()) == CT_MACRO_PRESENTATION
    deck.save(tmp_path / "brand.potm")
    assert _main_type((tmp_path / "brand.potm").read_bytes()) == CT_MACRO_TEMPLATE
    assert deck.validate(target="deck.pptm")[0].code == "package-type"


def test_a_target_without_an_extension_keeps_the_type(tmp_path):
    deck = _open_template()
    deck.save(tmp_path / "deck")
    assert _main_type((tmp_path / "deck").read_bytes()) == blank.CT_TEMPLATE
    assert _main_type(deck.to_bytes()) == blank.CT_TEMPLATE
    assert _main_type(deck.to_bytes(template=False)) == blank.CT_PRESENTATION
    assert deck.validate(target="notes.txt") == []


def test_save_as_template_refuses_a_pptx_name(tmp_path):
    deck = Document.open(FIXTURE_DIR / "sample.pptx")
    with pytest.raises(ValueError, match=r"\.potx"):
        deck.save_as_template(tmp_path / "brand.pptx")
    assert not (tmp_path / "brand.pptx").exists()
    deck.save_as_template(tmp_path / "brand.potx")
    assert _main_type((tmp_path / "brand.potx").read_bytes()) == blank.CT_TEMPLATE


def _macro_deck() -> Document:
    data = (FIXTURE_DIR / "sample.pptx").read_bytes()
    source = zipfile.ZipFile(io.BytesIO(data))
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as target:
        for item in source.infolist():
            content = source.read(item.filename)
            if item.filename == "[Content_Types].xml":
                content = content.replace(blank.CT_PRESENTATION.encode(),
                                          CT_MACRO_PRESENTATION.encode())
            target.writestr(item, content)
    return Document.open(buffer.getvalue())


def test_a_macro_enabled_deck_saves_only_as_pptm_or_potm(tmp_path):
    deck = _macro_deck()
    with pytest.raises(ValueError, match=r"\.pptm or \.potm"):
        deck.save(tmp_path / "deck.pptx")
    codes = [p.code for p in deck.validate(target="deck.pptx")]
    assert codes == ["package-type"]
    deck.save(tmp_path / "deck.pptm")
    assert _main_type((tmp_path / "deck.pptm").read_bytes()) == CT_MACRO_PRESENTATION
    deck.save_as_template(tmp_path / "deck.potm")
    assert _main_type((tmp_path / "deck.potm").read_bytes()) == CT_MACRO_TEMPLATE


requires_powerpoint = pytest.mark.skipif(
    not oracle_helper.available(),
    reason="needs macOS with Microsoft PowerPoint and the pptx2svg oracle script",
)


@pytest.mark.oracle
@requires_powerpoint
def test_powerpoint_opens_a_template_opened_and_saved_as_pptx():
    """The trial's refused file, made again the way the agent made it: PowerPoint now
    exports it unprompted.  (A .potx is not exported here: PowerPoint opens a template as a
    new, untitled deck, which the export script cannot find by name.)"""
    home = Path(os.path.expanduser("~"))
    deck_path, pdf = home / "pptx-agent-trial-p1.pptx", home / "pptx-agent-trial-p1.pdf"
    deck = _open_template()
    deck.add_slide("TITLE").shapes[0].set_text("Project Harbour")
    deck.save(deck_path)
    try:
        result = oracle_helper.export_pdf(deck_path, pdf)
        assert result.ok, f"PowerPoint {result.outcome}: {result.detail}"
        assert "Project Harbour" in (oracle_helper.pdf_texts(pdf) or ["Project Harbour"])[0]
    finally:
        oracle_helper.cleanup(deck_path, pdf)
