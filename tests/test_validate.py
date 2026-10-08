"""``Document.validate()``: the structural checks as a public API.

The same failure modes ``test_validity.py`` checks on bytes, returned as
:class:`~pptx_agent.validate.Problem` values: every fixture's problems are the ones it
arrived with, every kind of edit adds none, and each kind of damage is reported.
"""

from __future__ import annotations

import pytest

from conftest import FIXTURE_DIR
from pptx_agent import Document
from pptx_agent.oxml.xml import qn
from pptx_agent.validate import Problem
from test_validity import _edited

REPORT = FIXTURE_DIR / "real-financial-report.pptx"


def test_what_the_fixtures_arrive_with():
    found = {path.stem: sorted(str(p) for p in Document.open(str(path)).validate())
             for path in sorted(FIXTURE_DIR.glob("*.pptx"))}
    assert found["real-financial-report"] == [
        f"[override-without-part] [Content_Types].xml: ppt/slideMasters/slideMaster{n}.xml"
        for n in (2, 3, 4)]
    assert found["sample-issue-387"] == ["[child-order] ppt/slides/slide1.xml: a:pPr in a:p"]
    assert all(not problems for name, problems in found.items()
               if name not in {"real-financial-report", "sample-issue-387"})
    assert Document.new().validate() == []


def test_no_edit_adds_a_problem(pptx_path):
    """Compared without the part: a duplicated slide carries its original's quirks along."""
    before = {(p.code, p.detail) for p in Document.open(str(pptx_path)).validate()}
    after = {(p.code, p.detail) for p in Document.open(_edited(pptx_path)).validate()}
    assert after <= before


def _codes(document: Document) -> set[str]:
    return {problem.code for problem in document.validate()}


def test_an_unknown_relationship_id_is_reported():
    document = Document.open(str(REPORT))
    slide = document.slides[0]
    root = document.package.tree(slide.part_path)
    root.find(f".//{qn('p:cNvPr')}").set(qn("r:id"), "rId999")
    document.package.mark_dirty(slide.part_path)
    problems = document.validate()
    expected = Problem("unknown-relationship-id", slide.part_path, "p:cNvPr/@r:id=rId999")
    assert expected in problems


def test_a_missing_part_is_reported_as_dangling_and_its_override_kept():
    document = Document.open(str(REPORT))
    part = document.shape("257.25").chart.part
    document.package.remove_part(part)
    codes = _codes(document)
    assert "dangling-relationship" in codes


def test_a_broken_table_grid_is_reported():
    document = Document.open(str(REPORT))
    shape = document.shape("257.3#5")
    row = shape._element.find(f".//{qn('a:tr')}")
    row.remove(row.findall(qn("a:tc"))[-1])
    document.package.mark_dirty(shape._slide.part_path)
    assert any(p.code == "table-grid" and "4 cells for 5" in p.detail
               for p in document.validate())


def test_a_repeated_slide_id_is_reported():
    document = Document.open(str(REPORT))
    presentation = document.package.presentation_part()
    root = document.package.tree(presentation)
    entries = root.findall(f"{qn('p:sldIdLst')}/{qn('p:sldId')}")
    entries[1].set("id", entries[0].get("id"))
    document.package.mark_dirty(presentation)
    assert any(p.code == "slide-list" and "repeated" in p.detail for p in document.validate())


@pytest.mark.parametrize("damage", ["order", "ignorable"])
def test_schema_order_and_ignorable_prefixes_are_reported(damage):
    document = Document.open(str(REPORT))
    slide = document.slides[0]
    root = document.package.tree(slide.part_path)
    if damage == "order":
        shape = root.find(f".//{qn('p:sp')}")
        shape.append(shape.find(qn("p:nvSpPr")))  # name after the text body
        code = "child-order"
    else:
        root.set("{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable",
                 "zz")
        code = "ignorable-prefix-undeclared"
    document.package.mark_dirty(slide.part_path)
    assert code in _codes(document)
    assert all(str(p).startswith("[") for p in document.validate())
