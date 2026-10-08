# Design

Moved from the README.

Three layers:

| Layer | Module | Role |
| --- | --- | --- |
| Document | `pptx_agent.oxml` | the OPC package and typed views over its XML — the source of truth |
| API | `pptx_agent.edit` | slides, shapes, text, stable ids, undo |
| Render | [`pptx2svg`](https://github.com/uvrt/pptx2svg) | SVG and PNG, as a dependency |

**The document is the XML, not a copy of it.** A typed node here is a live view over an lxml
element: `shape.left` reads and writes `sp/spPr/a:xfrm/a:off/@x` on demand rather than holding a
parsed value. Nothing is ever lifted out of the tree, so unknown elements, `extLst`,
`mc:AlternateContent` branches, animations, embedded objects and namespace declarations survive
a read/edit/write cycle because nothing ever removed them. Losslessness is structural rather
than a feature that has to be maintained.

That is checked, not asserted: opening and saving any deck in `tests/fixtures/` reproduces
**every part byte for byte**, and an edited part is compared against the original in canonical
form so a dropped namespace or a reordered child cannot hide inside it.

**Ids are stable.** Shapes are addressed as `"<sldId>.<cNvPr id>"` — never by position, which
changes under reorder. Two wrinkles the obvious implementation gets wrong:

- `cNvPr@id` is **not** unique in real decks. `real-financial-report.pptx` has a `p:sp` and a
  `p:graphicFrame` both at `id="3"`. Collisions are broken apart with the shape-tree path:
  `256.3#1`.
- Editing a shape **stamps** its current id into `cNvPr/a:extLst`, so it keeps that address even
  if PowerPoint later renumbers. Reading never stamps, or the byte-identity guarantee would
  fail.

**The picture is addressable.** `slide.render_svg()` tags every shape group with
`data-pptx-id` carrying the same id the API answers to, so an agent can point at what it sees.

## Why lxml

Editing depends on lxml. That is not negotiable: `xml.etree` rewrites namespace prefixes and
drops declarations it believes are unused, which silently invalidates `mc:Ignorable="a14"` and
produces a file PowerPoint offers to repair. lxml reproduces the input exactly, apart from the
CRLF→LF normalisation the XML specification requires.

## Prior art

The design was informed by [pptx-svg](https://github.com/t-ujiie-g/pptx-svg) and
[moon-pptx](https://github.com/t-ujiie-g/moon-pptx). What was taken, and what was deliberately
done differently, is recorded in [ROADMAP.md](../ROADMAP.md).
