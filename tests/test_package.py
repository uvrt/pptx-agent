"""Adding parts: media, content types and relationships, all reversible by undo."""

from __future__ import annotations

import io
import zipfile

import pytest

from pptx_agent.core.history import History
from pptx_agent.oxml.package import REL_IMAGE, PresentationPackage

PNG = b"\x89PNG\r\n\x1a\n" + b"\x00" * 32  # enough for the magic-number sniff


def _entries(data: bytes) -> dict[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()}


def test_an_added_image_is_stored_typed_and_related(financial_report):
    package = PresentationPackage.open(str(financial_report))
    slide = package.slide_parts()[0][1]

    media = package.add_image(PNG)
    rel_id = package.add_relationship(slide, REL_IMAGE, media)

    reread = PresentationPackage.open(package.to_bytes())
    assert reread.read(media) == PNG
    assert reread.content_type(media) == "image/png"
    assert reread.related_part(slide, rel_id) == media
    assert reread.relationships(slide)[rel_id].target.startswith("../media/")


def test_identical_images_are_stored_once(product_page):
    package = PresentationPackage.open(str(product_page))
    assert package.add_image(PNG) == package.add_image(PNG)


def test_a_relationship_is_reused_not_duplicated(product_page):
    package = PresentationPackage.open(str(product_page))
    slide = package.slide_parts()[0][1]
    media = package.add_image(PNG)
    assert package.add_relationship(slide, REL_IMAGE, media) == package.add_relationship(
        slide, REL_IMAGE, media)


def test_undo_removes_added_parts_byte_for_byte(pptx_path):
    package = PresentationPackage.open(str(pptx_path))
    original = _entries(package.to_bytes())
    history = History(package)
    slide = package.slide_parts()[0][1]

    history.checkpoint()
    package.add_relationship(slide, REL_IMAGE, package.add_image(PNG))
    assert set(_entries(package.to_bytes())) > set(original)

    assert history.undo()
    assert _entries(package.to_bytes()) == original
    assert history.redo()
    assert any(name.startswith("ppt/media/") and data == PNG
               for name, data in _entries(package.to_bytes()).items())


def test_a_part_without_relationships_gets_a_rels_part(product_page):
    package = PresentationPackage.open(str(product_page))
    owner = package.add_part("ppt/custom/thing.xml", b"<x/>", "application/xml")
    rel_id = package.add_relationship(owner, REL_IMAGE, package.add_image(PNG))
    reread = PresentationPackage.open(package.to_bytes())
    assert reread.related_part(owner, rel_id).startswith("ppt/media/")
    assert reread.content_type("ppt/custom/_rels/thing.xml.rels") is not None


def test_unknown_image_formats_are_refused(product_page):
    package = PresentationPackage.open(str(product_page))
    with pytest.raises(ValueError, match="unsupported image"):
        package.add_image(b"GIF? no")
    with pytest.raises(ValueError, match="already exists"):
        package.add_part(package.slide_parts()[0][1], b"<x/>")


# -- removing, copying and reaping parts (E2) -----------------------------------------------


def _order(data: bytes) -> list[str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        return [info.filename for info in archive.infolist()]


def test_a_removed_part_comes_back_in_place_on_undo(pptx_path):
    package = PresentationPackage.open(str(pptx_path))
    original = package.to_bytes()
    history = History(package)
    slide = package.slide_parts()[0][1]

    history.checkpoint()
    package.remove_part(slide)
    assert not package.has_part(slide)
    assert not package.has_part(slide.replace("slides/", "slides/_rels/") + ".rels")
    assert package.content_type(slide) != "application/vnd.openxmlformats-officedocument." \
                                          "presentationml.slide+xml"
    assert history.undo()
    assert package.to_bytes() == original
    assert _order(package.to_bytes()) == _order(original)


def test_reap_removes_a_cycle_nothing_else_reaches(financial_report):
    """A slide and its notes point at each other; with the presentation's link gone, both go."""
    package = PresentationPackage.open(str(financial_report))
    presentation = package.presentation_part()
    slide = package.slide_parts()[1][1]
    rel_id = next(r.id for r in package.relationships(presentation).values()
                  if r.target_part == slide)
    package.remove_relationship(presentation, rel_id)
    removed = package.reap([slide])
    assert slide in removed
    assert any(name.startswith("ppt/notesSlides/") for name in removed)
    assert any(name.startswith("ppt/embeddings/") for name in removed)
    assert "ppt/notesMasters/notesMaster1.xml" not in removed  # still the presentation's


def test_reap_keeps_a_part_anything_else_relates_to(financial_report):
    package = PresentationPackage.open(str(financial_report))
    slide = package.slide_parts()[1][1]
    assert package.reap([slide]) == []  # the presentation still lists it


def test_release_keeps_a_relationship_still_referenced(product_page):
    package = PresentationPackage.open(str(product_page))
    slide = package.slide_parts()[0][1]
    layout = next(r for r in package.relationships(slide).values() if r.type.endswith("/slideLayout"))
    # Implicit (no XML reference) relationships are never passed by callers; but a referenced
    # one passed by mistake must survive.
    media = package.add_image(PNG)
    rel_id = package.add_relationship(slide, REL_IMAGE, media)
    root = package.tree(slide)
    root.set("probe", rel_id)
    assert package.release(slide, [rel_id]) == []
    assert rel_id in package.relationships(slide)
    del root.attrib["probe"]
    assert package.release(slide, [rel_id]) == [media]
    assert layout.id in package.relationships(slide)


def test_copy_part_copies_what_is_not_shared(financial_report):
    package = PresentationPackage.open(str(financial_report))
    slide = package.slide_parts()[1][1]
    copy = package.copy_part(slide, share=lambda rel: rel.type.endswith("/slideLayout"))
    assert copy == "ppt/slides/slide5.xml"
    old, new = package.relationships(slide), package.relationships(copy)
    assert set(old) == set(new)
    for rel_id in old:
        if old[rel_id].type.endswith("/slideLayout"):
            assert new[rel_id].target_part == old[rel_id].target_part
        else:
            assert new[rel_id].target_part != old[rel_id].target_part
    notes = next(r.target_part for r in new.values() if r.type.endswith("/notesSlide"))
    assert copy in [r.target_part for r in package.relationships(notes).values()]
    assert package.content_type(copy) == package.content_type(slide)


def test_numbered_template():
    from pptx_agent.core.opc import numbered_template

    assert numbered_template("ppt/charts/chart12.xml") == "ppt/charts/chart{n}.xml"
    assert numbered_template("a/b.bin") == "a/b{n}.bin"
    assert numbered_template("noext") == "noext{n}"


def test_an_external_relationship_is_reused(product_page):
    package = PresentationPackage.open(str(product_page))
    slide = package.slide_parts()[0][1]
    kind = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    first = package.add_external_relationship(slide, kind, "https://example.org")
    assert package.add_external_relationship(slide, kind, "https://example.org") == first
    assert package.relationships(slide)[first].is_external


def test_an_embedded_package_round_trips_untouched(financial_report):
    """A package inside the package opens with the same core, and an unedited one is not
    rewritten at all."""
    package = PresentationPackage.open(str(financial_report))
    part = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"
    nested = package.open_embedded(part)
    assert nested.main_document_part() == "xl/workbook.xml"
    nested.tree("xl/worksheets/sheet1.xml")  # parsing alone changes nothing
    assert package.replace_embedded(part, nested) is False
    assert part not in package._raw_changes


def test_an_edited_embedded_package_is_one_undoable_raw_write(financial_report):
    package = PresentationPackage.open(str(financial_report))
    original = _entries(package.to_bytes())
    history = History(package)
    part = "ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx"
    nested = package.open_embedded(part)
    root = nested.tree("xl/sharedStrings.xml")
    root.set("count", "999")
    nested.mark_dirty("xl/sharedStrings.xml")

    history.checkpoint()
    assert package.replace_embedded(part, nested) is True
    reread = PresentationPackage.open(package.to_bytes()).open_embedded(part)
    assert b'count="999"' in reread.read("xl/sharedStrings.xml")
    # Only the edited inner part differs; the rest of the workbook is byte-identical.
    inner_before = _entries(PresentationPackage.open(str(financial_report)).read(part))
    inner_after = _entries(package.read(part))
    assert {n for n in inner_before if not n.endswith("/") and inner_before[n] != inner_after.get(n)} \
        == {"xl/sharedStrings.xml"}

    assert history.undo()
    assert _entries(package.to_bytes()) == original


def test_a_package_is_written_with_replacements_and_keeps_its_own(product_page):
    """Parts derived at save time are written in place of a part's bytes without the
    package changing: the next save without them is the original again."""
    package = PresentationPackage.open(str(product_page))
    original = package.to_bytes()
    written = package.to_bytes({"docProps/app.xml": b"<x/>"})
    assert _entries(written)["docProps/app.xml"] == b"<x/>"
    assert {k: v for k, v in _entries(written).items() if k != "docProps/app.xml"} == \
        {k: v for k, v in _entries(original).items() if k != "docProps/app.xml"}
    assert package.to_bytes() == original
    assert list(_entries(written)) == list(_entries(original))  # entry order kept


def test_changed_parts_and_the_package_as_opened(product_page):
    package = PresentationPackage.open(str(product_page))
    history = History(package)
    assert package.changed_parts() == frozenset()
    slide = package.slide_parts()[0][1]
    history.checkpoint()
    media = package.add_image(PNG)
    package.add_relationship(slide, REL_IMAGE, media)
    assert media in package.changed_parts()
    opened = package.opened()
    assert not opened.has_part(media) and package.has_part(media)
    assert opened.to_bytes() == PresentationPackage.open(str(product_page)).to_bytes()
    history.undo()
    assert package.changed_parts() == frozenset()
