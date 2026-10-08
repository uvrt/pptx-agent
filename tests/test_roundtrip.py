"""The losslessness gate.

Everything else in this library rests on one property: reading a deck and writing it back
changes nothing.  If that fails, no amount of careful editing helps -- the file was already
damaged before the first edit.  These tests run over the whole fixture corpus rather than one
convenient deck, because the interesting failures come from decks written by Google Slides and
Keynote, not from the ones PowerPoint made.
"""

from __future__ import annotations

import io
import zipfile

from pptx_agent import Document
from pptx_agent.oxml.package import OoxmlPackage, normalize_part_path


def _entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {
            normalize_part_path(info.filename): archive.read(info)
            for info in archive.infolist()
            if not info.is_dir()
        }


def test_open_and_save_is_byte_identical(pptx_path):
    """No edits means no changes -- in every part, not just the ones we understand."""
    original = _entries(pptx_path.read_bytes())
    written = _entries(OoxmlPackage.open(str(pptx_path)).to_bytes())

    assert set(written) == set(original)
    differing = [name for name, data in original.items() if written[name] != data]
    assert differing == []


def test_entry_order_is_preserved(pptx_path):
    """Order is not required by the format, but changing it makes every saved diff useless."""
    with zipfile.ZipFile(io.BytesIO(pptx_path.read_bytes())) as archive:
        before = [i.filename for i in archive.infolist() if not i.is_dir()]
    with zipfile.ZipFile(io.BytesIO(OoxmlPackage.open(str(pptx_path)).to_bytes())) as archive:
        after = [i.filename for i in archive.infolist() if not i.is_dir()]
    assert before == after


def test_parsing_a_part_does_not_dirty_it(pptx_path):
    """Reading builds lxml trees; only a mutation may cause a re-serialization."""
    package = OoxmlPackage.open(str(pptx_path))
    for name in package.part_names:
        if name.endswith(".xml") or name.endswith(".rels"):
            package.tree(name)
    assert package.dirty_parts == frozenset()
    assert _entries(package.to_bytes()) == _entries(pptx_path.read_bytes())


def test_only_edited_parts_change(product_page):
    """An edit to one slide must not perturb any other part of the package."""
    document = Document.open(str(product_page))
    document.slides[0].shapes[0].move_by(dx=12700)

    original = _entries(product_page.read_bytes())
    written = _entries(document.to_bytes())

    changed = {name for name, data in original.items() if written[name] != data}
    assert changed == {"ppt/slides/slide1.xml"}


def test_an_edited_part_differs_only_where_intended(product_page):
    """Inside an edited part, nothing changes except what the edit asked for.

    The byte comparison used for untouched parts cannot be used here: re-serializing writes
    ``<x/>`` where the original had ``<x></x>``, normalises CRLF to LF, and quotes the XML
    declaration differently.  All three are semantically inert, so the check is made against
    the canonical form of the tree instead -- which still catches what actually matters: a
    dropped namespace declaration, a lost attribute, a reordered child.
    """
    from lxml import etree

    from pptx_agent.oxml.xml import PA_ID_EXT_URI, qn, remove

    document = Document.open(str(product_page))
    shape = document.slides[0].shapes[0]
    shape.rotation = 11

    after = etree.fromstring(_entries(document.to_bytes())["ppt/slides/slide1.xml"])

    # Undo the two intended additions, then the trees must be identical.
    for extension in after.iter(qn("a:ext")):
        if extension.get("uri") == PA_ID_EXT_URI:
            parent = extension.getparent()
            remove(extension)
            if len(parent) == 0:
                remove(parent)
            break
    for xfrm in after.iter(qn("a:xfrm")):
        if xfrm.get("rot") == "660000":
            del xfrm.attrib["rot"]
            break

    before = etree.fromstring(_entries(product_page.read_bytes())["ppt/slides/slide1.xml"])
    assert etree.tostring(after, method="c14n2") == etree.tostring(before, method="c14n2")


def test_an_edit_records_the_change_it_was_asked_for(product_page):
    document = Document.open(str(product_page))
    document.slides[0].shapes[0].rotation = 11
    part = _entries(document.to_bytes())["ppt/slides/slide1.xml"]
    assert b'rot="660000"' in part  # 11 degrees in 1/60,000ths


def test_undo_restores_the_original_bytes(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()

    shape = document.slides[0].shapes[0]
    shape.move_by(dx=914400)
    assert document.to_bytes() != original

    assert document.undo()
    assert _entries(document.to_bytes()) == _entries(original)


def test_a_batch_is_a_single_undo_step(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()

    with document.batch():
        shape = document.slides[0].shapes[0]
        shape.left = 100
        shape.top = 200
        shape.rotation = 5

    assert document.undo()
    assert _entries(document.to_bytes()) == _entries(original)
    assert not document.history.can_undo()


def test_a_failed_batch_rolls_back(product_page):
    document = Document.open(str(product_page))
    original = document.to_bytes()

    class Boom(Exception):
        pass

    try:
        with document.batch():
            document.slides[0].shapes[0].left = 999
            raise Boom
    except Boom:
        pass

    assert _entries(document.to_bytes()) == _entries(original)


def test_redo_reapplies(product_page):
    document = Document.open(str(product_page))
    document.slides[0].shapes[0].left = 123456
    edited = document.to_bytes()

    document.undo()
    assert document.redo()
    assert _entries(document.to_bytes()) == _entries(edited)


# -- the same gates for the semantic API (E1) -------------------------------------------------


def _strip_stamps(root) -> None:
    from pptx_agent.oxml.xml import PA_ID_EXT_URI, qn, remove

    for extension in list(root.iter(qn("a:ext"))):
        if extension.get("uri") == PA_ID_EXT_URI:
            parent = extension.getparent()
            remove(extension)
            if len(parent) == 0:
                remove(parent)


def test_e1_edits_undo_to_the_original_bytes(pptx_path):
    """Text, fills, outlines, image fills, tables across merges, groups: undo is exact."""
    from test_validity import _apply_e1_edits

    document = Document.open(str(pptx_path))
    original = _entries(document.to_bytes())
    with document.batch():
        _apply_e1_edits(document)
    if _entries(document.to_bytes()) == original:
        import pytest

        pytest.skip("nothing in this fixture for the E1 edits to touch")
    assert document.undo()
    assert _entries(document.to_bytes()) == original


def test_e1_edits_undo_step_by_step(pptx_path):
    """Without a batch every edit is its own step, and walking all of them back is exact."""
    from test_validity import _apply_e1_edits

    document = Document.open(str(pptx_path))
    document.history._max_depth = 10_000
    original = _entries(document.to_bytes())
    _apply_e1_edits(document)
    while document.undo():
        pass
    assert _entries(document.to_bytes()) == original


def test_a_run_edit_differs_only_where_intended(product_page):
    from lxml import etree

    from pptx_agent.oxml.xml import qn

    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes if s.kind == "shape" and s.text)
    run = shape.text_frame.paragraph(0).run(0)
    assert run.italic is None
    run.italic = True

    after = etree.fromstring(_entries(document.to_bytes())["ppt/slides/slide1.xml"])
    _strip_stamps(after)
    edited = [rpr for rpr in after.iter(qn("a:rPr")) if rpr.get("i") == "1"]
    assert len(edited) == 1
    del edited[0].attrib["i"]

    before = etree.fromstring(_entries(product_page.read_bytes())["ppt/slides/slide1.xml"])
    assert etree.tostring(after, method="c14n2") == etree.tostring(before, method="c14n2")


def test_a_theme_fill_differs_only_where_intended(product_page):
    """Swapping a hex fill for a theme fill changes that one fill element and nothing else."""
    import copy

    from lxml import etree

    from pptx_agent.oxml.xml import qn

    document = Document.open(str(product_page))
    shape = next(s for s in document.slides[0].shapes
                 if s.kind == "shape" and s.fill is not None and s.fill.kind == "solid")
    name = shape.name
    shape.fill = "accent1"

    def fill_of(root):
        for properties in root.iter(qn("p:cNvPr")):
            if properties.get("name") == name:
                sp = properties.getparent().getparent()
                return sp.find(qn("p:spPr")).find(qn("a:solidFill"))

    after = etree.fromstring(_entries(document.to_bytes())["ppt/slides/slide1.xml"])
    before = etree.fromstring(_entries(product_page.read_bytes())["ppt/slides/slide1.xml"])
    _strip_stamps(after)
    new = fill_of(after)
    assert etree.tostring(new[0]).startswith(b"<a:schemeClr")
    new.addprevious(copy.deepcopy(fill_of(before)))
    new.getparent().remove(new)
    assert etree.tostring(after, method="c14n2") == etree.tostring(before, method="c14n2")
