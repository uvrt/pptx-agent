# Roadmap

Working plan for `pptx-agent`. Written to be picked up cold: every phase says what is missing,
where it goes, how to approach it, and how to know it is done.

**Effort key:** S ≈ half a day · M ≈ 1–3 days · L ≈ 1–2 weeks · XL ≈ 3+ weeks.

---

## Prior art, and what was taken from it

Two MIT libraries by the same author solve this problem, and the difference between them is the
most useful thing in the reference material.

### `pptx-svg` (TypeScript + MoonBit/Wasm)

| Module | Size | Role |
| --- | --- | --- |
| `src/xml/xml.mbt` | 24 KB | generic XML DOM |
| `src/ooxml/ooxml_parse.mbt` | 71 KB | OOXML → typed AST |
| `src/ooxml/ooxml.mbt` | 50 KB | the typed AST |
| `src/serializer/serializer.mbt` | 49 KB | typed AST → OOXML |
| `src/main/main_edit.mbt` (+ text, table) | 48 KB | the mutation API |
| `src/renderer/*` | ~480 KB | SVG rendering |
| `src/svg_parser/svg_parser.mbt` | 32 KB | SVG → state |

It parses a slide into a typed in-memory `SlideData`, mutates that, and re-serializes the whole
part on export. Undo snapshots "the OOXML of any in-engine modified slides"; batches nest, and
only the outermost pair takes effect.

**Taken:** the snapshot-per-part undo model — cheap, and exact rather than an attempt to invert
each operation. Implemented in `edit/history.py`.

**Rejected:** addressing by `(slideIdx, shapeIdx)`. Its own documentation concedes that every
reorder invalidates the reference, so each mutator has to return a fresh index. An agent holding
a reference across edits cannot work that way. See `edit/ids.py`.

**The structural weakness.** Because the typed AST is *owned* — parsed out of the XML and written
back from scratch — anything it does not model is lost on serialization. The symptom is visible
in its SVG specification: raw XML stuffed into attributes as escape hatches
(`data-ooxml-transition-xml`, `data-ooxml-timing-xml`, `data-ooxml-mc-choice`, `data-ooxml-ole`,
`data-ooxml-math-xml`) alongside ~200 typed attributes, and documented losses anyway — charts
cannot be copied, merged-cell spans are not adjusted when rows or columns change.

### `moon-pptx` (pure MoonBit)

The same author's later, general-purpose library, and it **fixes exactly that weakness**: "every
model node carries `extension : Array[XmlElement]`, so third-party files round-trip with zero
data loss even through features moon-pptx doesn't model." It carries `roundtrip_test.mbt` at four
levels — xml, opc, slide, integration — plus a corpus test against real-world files.

**Taken:** round-trip fidelity as a first-class, corpus-wide gate, not a nice-to-have.
`tests/test_roundtrip.py` is the first thing that runs and the thing everything else rests on.

**Improved on:** the `extension` field exists because the model is owned. Here, typed nodes are
*live views over lxml elements* — nothing is ever detached from the tree, so there is nothing to
re-attach and no escape hatch to maintain. Losslessness is structural.

---

## Where we are

Phases E0 to E6 are complete. E5 built shapes, text boxes, connectors and tables from
nothing, and new decks: from nothing -- PowerPoint's Office theme, master and eleven
layouts, at any size -- or from a template, with their metadata. E6 reads a deck as a
Markdown outline with ids, and drafts slides from one into the deck's own layouts, speaker
notes included; the outline round trip is stable on every fixture. The full end-to-end
trial's findings are addressed (see "Usability (full trial)"). 1,885 tests over seven
decks -- six real-world, one SmartArt deck PowerPoint itself wrote -- and two of the trial's
inputs, plus 63 opt-in PowerPoint oracle tests.

| Area | State |
| --- | --- |
| OPC read/write, byte-identical round trip | Complete |
| Format-neutral core: [ooxml-edit](https://github.com/uvrt/ooxml-edit), `pptx_agent.core` a shim | Complete — see below |
| Charts, workbooks and SmartArt editing in ooxml-edit's optional `ooxml_edit.charts`, adapted here | Complete — see below |
| Stable ids: derivation, collision handling, stamping | Complete |
| Undo / redo / nested batches, including added parts | Complete |
| Position, size, rotation, placeholder inheritance (through layout and master) | Complete |
| Flips, preset geometry | Complete (E3) |
| Text: frames, paragraphs, runs, positional addresses | Complete (E1) |
| Re-cutting a paragraph into runs without changing its text (`Paragraph.segment`) | Complete (E3) |
| Run formatting: bold, italic, underline, strike, size, typeface, colour | Complete (E1) |
| Paragraphs: alignment, level, indents, bullets/numbering, spacing | Complete (E1) |
| `set_text` that keeps mixed formatting | Complete (E1) |
| Fills: solid, gradient, image, none — theme colours kept as `schemeClr` | Complete (E1) |
| Outlines: width, colour, dash, cap, arrowheads | Complete (E1) |
| Tables: cell text/fill/borders, sizing, merge/split, merge-aware row/column insert/delete | Complete (E1) |
| Groups: addressable children, child-space moves, grouping, PowerPoint-style refit | Complete (E1) |
| Ungrouping, with transforms composed into the parent | Complete (E2) |
| Delete, duplicate, z-order (front, back, one step forward or back) | Complete (steps: E5) |
| Slides: add from a layout, delete, duplicate, reorder; sections and custom shows kept | Complete (E2) |
| Pictures: insert (de-duplicated), replace the image | Complete (E2) |
| Hyperlinks on runs: web/mail addresses and slide jumps, with relationship upkeep | Complete (E2) |
| Relationship reaping: unreferenced relationships, parts and overrides removed | Complete (E2) |
| Render to SVG/PNG with matching ids | Complete |
| Full-state SVG: `data-ooxml-*` vocabulary, raw-XML floor, reader | Complete (E3) |
| Charts: values, labels, series and categories, titles, legend -- cache and embedded workbook in step | Complete (E4) |
| SmartArt: node text (data model and cached drawing), adding and removing nodes | Complete (E4) |
| Charts and SmartArt in the full-state SVG (`chart-data`, `diagram-nodes`) | Complete (E4) |
| New autoshapes (all 187 presets) with adjust values, text boxes, tables -- as PowerPoint writes them | Complete (E5) |
| Connectors (straight, elbow, curved) attached to connection sites, re-routed when a shape moves | Complete (E5) |
| New shapes inside groups, in the group's child space | Complete (E5) |
| A new deck from nothing: Office theme, master, the eleven layouts, at 16:9, 4:3 or any size | Complete (E5) |
| A new deck from a `.pptx` or `.potx` template; saving as `.potx` | Complete (E5) |
| Deck metadata (title, author, created, modified, language), slide size with PowerPoint's scaling, `app.xml` kept true | Complete (E5) |
| Connections and adjust values in the full-state SVG | Complete (E5) |
| Markdown outline: a reading view with slide and shape ids (`to_outline`) | Complete (E6) |
| Slides drafted from Markdown into the deck's own layouts (`insert_outline`) | Complete (E6) |
| Speaker notes: read in the outline, written when drafting (the notes master added as PowerPoint adds it) | Complete (E6); no typed edit API |
| Effects (shadow, glow), text-body properties (autofit, insets, anchor) | **No typed API**; editable as raw XML in the full-state SVG |
| Comments: PowerPoint's modern (threaded) comments on a slide or a shape -- read, add, reply, resolve, reopen, edit, delete (`edit/comments.py`, the parts as PowerPoint for Mac reads and writes them back) | Complete (post-T4) |
| Shape-level click actions, the older (non-threaded) comments; notes text outside drafting | **Not editable** |
| Slide-level state in the SVG (background, transitions, notes), z-order through the SVG | **Not in the vocabulary** |
| Chart formatting (colours, number formats, axis scaling, data labels), chart type changes | **No typed API**; raw XML in the chart part |

### Known approximations

- Reaping only considers relationship ids an edit *stopped* referencing (`Slide._settle`).
  Relationships a deck arrives with that nothing references are left alone on purpose, so an
  edit never changes more of the package than it touched; an explicit "clean up" pass would
  be easy to add on top of `OpcPackage.reap` but is not there.
- A removed part's `[Content_Types].xml` `Override` goes with it, but an extension `Default`
  (`png`, say) stays even when no part of that extension is left. Harmless, and what Office
  itself does.
- `docProps/app.xml` is brought up to date as the deck is written, not as it is edited:
  the slide count, the slide titles, and the notes, hidden-slide, word and paragraph counts
  moved by what the edits changed (so an application's own way of counting is kept). A
  slide without a title is listed as "PowerPoint Presentation", in English; in a localised
  file the titles group is found by matching the titles the deck was opened with, and left
  alone when that fails. Fonts and themes listed there are not recomputed.
- Ungrouping a group scaled *non-uniformly* whose child is rotated by other than a multiple of
  180° cannot be exact: the composed transform is a skew, which a shape's `xfrm` cannot hold.
  The child's frame is scaled axis by axis (pixel-exact in every other case, see
  `test_ungroup_renders_identically`). Which way PowerPoint itself resolves this is unprobed.
  Group effects (shadow, glow on `grpSpPr`) are dropped on ungroup.
- A duplicated slide's hyperlink to *itself* is re-pointed at the copy; links to other slides
  are kept. Comments and tags are copied with the slide (the safe default for any per-slide
  relationship type this library does not know).
- `Slide.add_picture` sizes an image at 96 dpi from its pixel size; a PNG's `pHYs` or a JPEG's
  density is not read. TIFF sizes are not read at all, so a TIFF needs an explicit size.
- Replacing a picture's image keeps its frame and drops its crop (`a:srcRect`) and any SVG
  twin; PowerPoint's "Change Picture" also re-fits the frame to the new aspect ratio.
- New slides copy the layout's placeholders but not its date, footer or slide-number
  placeholders, as PowerPoint does; an empty placeholder shows only in edit mode, so a new
  slide with no text set is blank in a PDF export.
- The id stamp declares its namespace inline (`xmlns:ns0=...`) because the prefix is not on the
  document root. Valid and preserved, but noisy; hoisting it is cosmetic.
- A placeholder's inherited frame is found through its layout and then the master, where
  the master placeholder is matched by kind: a title or centred title to the title, the
  footers to theirs, anything else (body, subtitle, content, picture, chart, table) to the
  body. That is how PowerPoint's own masters are built; a master laid out otherwise (with
  no body, say) is unprobed.
- `Shape.slide_bounds` composes group offsets and scales but not group rotation or flips, and a
  rotated or flipped group is not re-fitted when a child moves (re-fitting moves its rotation
  centre, and so would move the children).
- `Shape.text` still reads a line break (`a:br`) as `"\n"`, as in E0; `TextFrame.text` is the
  exact form, with `"\v"` for a break and `"\n"` between paragraphs.
- A table cell border is per cell, and adjacent cells each carry a copy of their shared edge.
  `TableCell.set_border(..., shared=True)` styles both copies; `cell.border(side)` alone does not.
- No fixture has a merged table or a group, so those paths are exercised on merges and groups
  the tests create. A deck with PowerPoint-authored merges and groups would strengthen it.

**From E3, the full-state SVG:**

- Typed attributes are compared with the *document*, feature by feature, and the SVG's value
  wins where they differ. Applying a stale SVG therefore undoes later changes to the
  features it disagrees on (and only those); it is a merge, not a conflict check. Slide
  identity, part, layout, master, theme and size *are* checked, and a mismatch is refused.
- Some typed values have no E1/E2 edit behind them and are refused with a pointer to
  `data-ooxml-xml` rather than written by hand: pattern and group fills, gradient and pattern
  outlines, a custom dash, a custom geometry, percentage `spcBef`/`spcAft`, picture bullets,
  `dblStrike`, click actions other than an address or a slide, making or unmaking a field,
  and turning an explicit position back into an inherited one.
- Z-order is not read from the SVG for shapes that already exist (only new shapes are placed
  by it), and a shape cannot move between the slide and a group through the SVG: a group
  missing from an SVG that keeps its child is refused rather than read as an ungroup.
- Table rows and columns added or removed in the SVG are lined up with the document by their
  cells' text, so two identical rows are ambiguous (the first match wins). Changing the
  number of rows *and* of columns in one apply is refused; do it in two.
- Moving a group and one of its children in the same apply: the child's move re-fits the
  group (as in PowerPoint), so the group's frame need not end where the SVG put it.
- Media, charts and diagrams travel by reference. An SVG can be applied to its own deck or a
  copy of it; a relationship to a part the package does not have is refused, not embedded.
- An empty run (`{"t": ""}`) can be kept but not created or moved by re-cutting a paragraph.
- Replacing a shape's raw XML with a different `cNvPr@id` changes the shape's id, unless it
  was already stamped.
- `data-ooxml-text` and `data-ooxml-table` describe text with the E1 properties only; East
  Asian and complex-script typefaces, highlight, spacing between characters, text-body
  properties and cell margins ride on the raw floor, not on typed attributes.

**From E4, charts and SmartArt:**

- Adding or removing a category or series needs the workbook laid out the way PowerPoint
  lays it out: each data source one row or column, all of them over the same span, on one
  sheet. Value and label edits only need each formula to be a single range. Anything else
  (several areas, a defined name, multi-level categories, ranges that do not line up) is
  refused with `ChartDataError` before anything changes, rather than guessed at; a cell in
  the way of the data growing is refused the same way.
- Inserting a point moves the cells *of the chart's lines* only, not whole sheet rows:
  cells beside the data in other columns stay where they are, as they would not in Excel's
  "insert row". Formulas elsewhere in the workbook that point into the data are not
  rewritten.
- A new series goes into the plot of the series before it, with the application's automatic
  colour (its template's own `spPr`, point formats and trendlines are dropped). Its data
  goes in the next column (or row) beside the data, which must be empty.
- A chart whose workbook is linked from outside, embedded as an OLE object, missing or
  unreadable is edited in the cache only, with a `ChartDataWarning`; a series added to one
  holds literal data (`c:strLit`/`c:numLit`), since there are no cells for formulas to name.
- A value written over a formula cell removes the formula, and the calculation chain with
  it (Excel rebuilds it). A cell holding a *shared* formula other cells use is refused.
- Removing a series closes the gap only when the series sit side by side; otherwise its
  cells are cleared and a table over them keeps the (now empty, generically named) column.
- Touching a table rewrites a malformed `ref` clean (`A1:D4'`, as one fixture's generator
  writes it, becomes `A1:D4`) -- Excel would repair it otherwise.
- Not edited: number formats, axis scaling, chart type, data labels, colours (all on the raw
  chart part), bubble sizes of a new series (left empty, as a literal), `c:extLst` data-label ranges.
- A SmartArt text edit patches the cached drawing only where the shape that shows the node
  is found through its `presOf` connection *and* its paragraphs are verifiably the texts of
  the nodes it presents; otherwise -- and after any node is added or removed -- the cached
  drawing is dropped. PowerPoint lays the diagram out again (measured, below); pptx2svg and
  other readers that draw the cache show an empty frame until PowerPoint has saved the deck.
- New SmartArt nodes get no presentation points; PowerPoint makes them (measured). The
  full-state SVG can add, remove and retext nodes but not move one or change its level.
- Whether PowerPoint's Edit Data shows what the workbook now holds is argued from the files,
  not observed: PowerPoint's AppleScript dictionary has no chart or chart-data objects, and
  Excel was not driven. See "Phase E4" for the evidence.

**From E5, new shapes and connectors:**

- Elbow and curved connectors are routed by rules measured from 180 PowerPoint-made routes
  (see "Phase E5"). 169 of the 177 distinct ones are reproduced exactly; the other eight
  differ in which side of an overlapping shape a detour takes, or are layouts with both
  ends level where PowerPoint's own route crosses a shape. Every route still starts and
  ends on its sites.
- PowerPoint's minimum extent of 12,700 EMU for an elbow's frame is mirrored, so an end
  level with the other can sit up to one point off its site -- as in PowerPoint's files.
- Connector geometry composes group offsets and scales but not group rotation or flips (as
  `Shape.slide_bounds`). A connector re-routes when a shape it is attached to moves,
  resizes, rotates, flips, changes preset or adjust values through the API or the
  full-state SVG -- a typed attribute, or a changed frame or geometry in the shape's raw
  XML. Moving the connector itself leaves it attached and where it was put.
- A connector cannot attach to a shape whose `cNvPr@id` is duplicated on the slide (the
  Google Slides case): `stCxn@id` could not say which one.
- New shapes get no `a16:creationId` and tables no `p14:modId`: PowerPoint writes random
  GUIDs there, this library writes its durable id stamp, and output stays deterministic.
  Table column and row ids are `10001...` rather than random numbers.
- A new shape's text language is the presentation's default text language
  (`defaultTextStyle`); PowerPoint writes its own editing language.
- A text box keeps the height it is given; PowerPoint grows an `spAutoFit` box to its text
  when it next lays the text out.
- PowerPoint's AppleScript makes connectors and shapes the way VBA's `AddShape`/
  `AddConnector` do; a line drawn by hand in the UI may be styled differently (UI-drawn
  connectors could not be measured without GUI scripting).
- pptx2svg draws a new shape's outline from the theme (`lnRef`) but not its fill
  (`fillRef`) or text colour (`fontRef`) when the theme's style entry uses `phClr` -- the
  same for shapes PowerPoint itself inserted. A pptx2svg finding; PowerPoint fills them.

**From E5, new decks:**

- Names are English where PowerPoint writes its UI language: the layouts ("Title and
  Content"), placeholders ("Content Placeholder 2"), prompts, the theme ("Office Theme"),
  the `app.xml` groups and an untitled slide. The default language is `en-US` unless given;
  the date field's cached text is the creation date as M/D/YYYY.
- Left out: the thumbnail, `a16:creationId`/`p14:creationId` on the master, layouts and
  their placeholders, and the theme-family extension. PowerPoint adds none of them back on
  save (measured) and needs none.
- At another size the master and layouts follow PowerPoint's rule; PowerPoint's own float
  rounding puts some of its frames one EMU from ours.
- Changing the size of a deck with slides follows PowerPoint's "Ensure Fit" (measured),
  except that an auto-fit text box keeps its scaled height (PowerPoint measures its text
  again) and table text is scaled like other text, where PowerPoint set 18 pt cell text to
  10 pt at 4:3; an `endParaRPr` gets the scaled size where PowerPoint drops the element.
- From a template: the notes master stays (PowerPoint drops one no slide uses when it saves,
  with its theme and any Windows printer settings); sections and custom shows go with the
  slides; title, author, last-modified-by, revision and dates start again while the
  template's other properties (company, category...) stay. A macro-enabled template is
  refused.
- `modified` is not stamped on save; set it when that is wanted. `app.xml` names this
  library as the application.
- There is no chart builder, so a deck made from nothing gets charts only by copying a
  chart part in (the round-trip acceptance has none).
- Drafting from an outline places what no placeholder takes on a grid in the master's body
  area, which can overlap a filled placeholder: positions are a draft for the shape API and
  the render loop. A table is as tall as its rows at PowerPoint's 370,840 EMU default, at
  most its frame; a picture in a picture placeholder is written with the placeholder's frame
  and a crop (PowerPoint's own form for that is unprobed); code is Consolas, a choice
  (PowerPoint has no code style).
- A notes master added to a template is given a copy of the slide master's theme -- what
  PowerPoint does for its own Office theme (measured: byte for byte); for another theme
  unprobed. At a size other than 16:9 its slide image is the slide's shape fitted into the
  16:9 box, also unprobed.

---

## The format-neutral core: ooxml-edit

The layer a docx editor shares with this one is its own package now,
[ooxml-edit](https://github.com/uvrt/ooxml-edit) (`ooxml_edit`, lxml only), extracted from
`src/pptx_agent/core/` with its git history and depended on from `pyproject.toml`:

| Module | Role |
| --- | --- |
| `ooxml_edit.opc` | the OPC package: parts, relationships, content types, lossless save, snapshots, packages inside packages, adding, copying and reaping parts |
| `ooxml_edit.xml` | namespace registry, `qn`, schema-ordered insertion with rank groups, `remove` |
| `ooxml_edit.history` | undo / redo / nested batches over any snapshottable package |
| `ooxml_edit.stamp` | durable ids frozen into an `extLst` extension |

**`pptx_agent.core` is a shim.** `pptx_agent.core.opc`, `.xml`, `.history` and `.stamp` are
ooxml-edit's modules themselves, registered under the old names, so existing imports keep
working and the namespace and child-order registries stay one. Code here imports
`ooxml_edit` directly; `tests/test_core_shim.py` checks both. The switch changed no output:
the whole suite passes, every fixture's open+save is byte-identical per part, and the E0-E5
edit sets (with undo and redo, and E5's new decks) write the same bytes per part as before.

PresentationML lives on top: `oxml/xml.py` registers the `p:`/`a:` namespaces and child
orders, `oxml/package.py` subclasses the package with `presentation_part()`/`slide_parts()`,
and `edit/ids.py` says *where* a shape's stamp goes. ooxml-edit's `tests/test_neutrality.py`
keeps that package free of any format's vocabulary -- PowerPoint's, Word's and Excel's -- and
of any import beyond the standard library, lxml and itself; its own tests build small
synthetic packages rather than ship decks.

E4 added one thing to the core: a package inside the package
(`OpcPackage.open_embedded`/`replace_embedded`), which is genuinely neutral -- an embedded
workbook in a deck and in a Word document are the same thing.

**Charts and diagrams now live in `ooxml_edit.charts`** (ooxml-edit 0.2.0), an optional
subpackage the core never imports, moved there with the history of `edit/workbook.py`,
`edit/chart.py` and `edit/diagram.py`: a `.docx` embeds the same chart part, workbook and
SmartArt data model, so the editing is shared. What is left here is the adapter:

- `edit/chart.py`'s `graphic_host(shape)` describes a slide's graphic frame as a
  `GraphicHost` -- the frame, the slide part whose relationships name the chart or diagram,
  the deck's undo step (`batch` + `checkpoint`), the shape id, `application="PowerPoint"`,
  `document="deck"` and `lang="en-US"` for the text it creates -- and `Chart` and `Diagram`
  take a shape resolver as before; `edit/workbook.py` re-exports the shared module.
- `edit/text.py` imports the paragraph rewriting back from `ooxml_edit.charts.dmltext`.
- The full-state SVG reads and validates chart and diagram JSON with
  `ooxml_edit.charts.model`, and applies it with `apply_chart_model`/`apply_diagram_model`;
  their `ChartModelError` becomes `FullStateError` with the same message.
- `ooxml_edit.charts.namespaces` owns the child orders of `a:p`, `a:r`, `a:br`, `a:fld` and
  every chart and diagram element; `oxml/xml.py` no longer registers copies, and the
  registry after import is the same as before.

The move changed no output: every fixture with a chart or SmartArt, through the E4
acceptance edits, each chart and diagram edit on its own, their refusals and cache-only
warnings, and the full-state SVG round trips, writes the same bytes per part, undo bytes,
warnings, error messages, SVGs and models as before.

**Why not ooxml-common.** That package is standard-library-only by design (pptx2svg's install
stays dependency-free), and this layer is lxml-only by necessity — `xml.etree` rewrites
namespace prefixes and breaks `mc:Ignorable`, which is exactly what an editor must never do.
ooxml-common reads and renders; ooxml-edit edits losslessly.

---

## Phase E1 — the semantic API

**Done.** The bulk of what an agent needs.

- **Text** (`edit/text.py`): `TextFrame` / `Paragraph` / `Run` facades over `a:txBody`, for
  shapes and table cells alike. Addressing is positional (`"256.5/p1/r0"`,
  `"256.7/cell2,0/p0/r1"`) and re-resolved per call through `Document.resolve` — runs have no
  identity in OOXML, and pretending otherwise would be a lie the caller pays for later. A
  facade therefore survives undo, and names a *position*, not a run.
- **`set_text` keeps mixed formatting.** E0 gave every paragraph the first run's formatting.
  Now the new text is diffed against the old — paragraphs, then characters — and every new
  character takes the formatting of the old character it replaces or follows. Unchanged
  paragraphs are not touched at all; short coincidental matches are folded into the edit so a
  complete retype still behaves like E0.
- **Colours** (`edit/color.py`): `"accent1"`, `"#4472C4"`, `"accent1 lumMod=75% lumOff=25%"` or
  `Color.theme("accent1", lum_mod=0.75)`. Theme colours are written as `<a:schemeClr>` with
  their modifiers, never resolved; colours read from a file (`sysClr`, `prstClr`...) round-trip
  verbatim.
- **Fill and stroke** (`edit/fill.py`): solid, gradient (linear or path), image and no fill on
  shapes, groups and cells; `LineFormat` for width, colour, dash, cap, arrowheads, on shape
  outlines and cell borders alike. A new fill replaces the old member of the
  `EG_FillProperties` choice rather than adding a second one.
- **Image fills** needed media insertion, which turned out small enough to do here rather than
  wait for E2: `OpcPackage.add_part` / `declare_content_type` / `add_relationship` in the core,
  `PresentationPackage.add_image` (de-duplicating by content) on top. Undo snapshots now cover
  raw parts too, so undoing an image fill removes the media, the relationship and the content
  type again — byte-identical.
- **Tables** (`edit/table.py`): cell text through the same text API; per-cell fill and
  borders; row heights and column widths, keeping the frame's extent in step; `merge`/`split`;
  and row/column insert and delete that adjust `gridSpan`, `rowSpan`, `hMerge` and `vMerge`
  correctly — inserting inside a merge lengthens it, at its edge leaves it alone, and deleting a
  merge's top row or left column hands the origin's content on. The merge pattern written is
  PowerPoint's own. `tests/test_table.py` checks every operation against an independent model
  of where the regions belong, including a 40-seed randomised sweep (~80 edits that cross a
  merge), and checks grid consistency after each step.
- **Groups**: children were already addressable (the id index descends groups); now
  `parent_group`, `children`, `child_offset`/`child_extent`, `slide_bounds`, and
  `move_by(..., space="slide")` for moves in slide units through a scaled group. Moving or
  resizing a child re-fits the group the way PowerPoint does — child space becomes the
  children's bounding box and the frame moves with it — so nothing else on the slide moves.
  `Slide.group(shapes)` creates a group (needed to test all this, since no fixture has one).
- **Element order**: every insertion goes through the schema-order tables, which now cover
  `spPr`, `grpSpPr`, `rPr`/`defRPr`/`endParaRPr`, `pPr`, `a:p`, `ln` and the six cell borders,
  `gradFill`, `blipFill`, `tbl`/`tblPr`/`tr`/`tc`/`tcPr` and the shape tree. Repeating choices
  (a paragraph's runs and breaks; a shape tree's shapes) share a rank, so inserting never
  reorders siblings. `tests/test_validity.py` checks the order of every edited slide.

**Acceptance.** Representative E1 edits (run formatting, a theme gradient, outlines with
arrowheads, an image fill, a table row inserted across a merge, a moved group child) applied
to every fixture: PowerPoint exports all six without a repair prompt, and the edits are
visible in the PDFs. `pytest -m oracle` repeats it (`test_powerpoint_opens_semantically_edited_decks`).

**Left for E2 and later**, because each changes the package's structure or needs relationship
bookkeeping beyond adding one: new picture shapes, ungrouping, hyperlinks on runs, reaping the
relationships an edit orphans (all done in E2). Not yet exposed but purely part-local, so cheap
whenever needed:
text-body properties (autofit, insets, vertical anchor), East Asian and complex-script
typefaces, character spacing and highlight, cell margins, effects.

## Phase E2 — structural operations

**Done.** Everything that changes the package, not just a part.

- **Package operations in the core** (`core/opc.py`, format-neutral): `remove_part` (with its
  relationships part and `Override`), `remove_relationship`, `add_external_relationship`,
  `copy_part` (follows relationships, sharing or copying each by a caller's predicate, keeping
  relationship ids so copied XML needs no rewriting), `release` and `reap`. Removal is part
  of the undo snapshot: a removed part's zip entry keeps its place, and raw changes are
  replayed in the order they happened, so undo *and redo* are byte-identical.
- **Reaping proves before it removes.** `release(part, ids)` drops a relationship only if its
  id appears in no attribute of the part at all (deliberately blunt: `r:id`, `r:embed`,
  `o:relid`... all count). `reap(candidates)` gathers what the candidates lead to, then drops
  from that set every part any relationship *from outside it* targets — from any part in the
  package, reachable or not, the root included — until stable. What is left is referenced only
  from within (a slide and its notes point at each other) and goes, with relationships parts
  and overrides. Shapes call it through `Slide._watch`/`_settle`: relationship ids the part
  referenced before an edit and not after are released. So deleting a picture or chart,
  replacing an image or an image fill, removing a hyperlink, or deleting a run or table row
  that carried one all clean up after themselves — and media another slide still shows stays.
- **Slides** (`edit/slides.py`, `Document`/`Slide`): `add_slide(layout, index=)`,
  `delete_slide`, `duplicate_slide(index=)`, `move_slide`, `Slide.delete/duplicate/move_to`,
  `Document.layouts`. Deleting removes the `sldIdLst` entry, the section and custom-show
  entries, the relationship, the part, its rels, its notes, its charts and their workbooks,
  and the overrides; hyperlinks elsewhere that jumped to it are removed (as PowerPoint does),
  and a reference that cannot be unlinked safely is refused. Duplicating shares the layout,
  media and linked slides and copies notes (re-pointed at the copy), charts with their
  workbooks, diagrams and anything unknown; it renews `p14:creationId`. New `sldId`s are one
  above the largest in use and at least 256. Sections are kept a partition of the slide list
  in order (no fixture has sections, so the tests add them).
- **Pictures**: `Slide.add_picture(image, left, top, width=, height=)` (natural size at
  96 dpi, aspect kept when one side is given, E1's de-duplicating `add_image`),
  `Shape.replace_image`, `Shape.image_part`.
- **Ungroup**: `Shape.ungroup()` composes each child's transform with the group's — child
  space scale, then the group's flips and rotation about its centre; a reflection reverses a
  child's own rotation — and hands a `grpFill` child the group's fill. Pixel-identical in
  pptx2svg for rotated, flipped and scaled groups.
- **Hyperlinks on runs**: `Run.hyperlink` (a `Hyperlink`), `set_hyperlink(url | slide,
  tooltip=)`, `remove_hyperlink()`; works in table cells too.

**Acceptance.** `tests/test_structure.py` runs each of nine operations (delete, duplicate,
reorder, add from layout, insert picture, replace image, ungroup, hyperlink, delete shapes) on
every fixture through four gates: every validity check (now including no orphaned parts, no
`Override` for a missing part, and a consistent slide list), edit → save → reopen → read back,
undo-all to the original bytes and redo-all to the edited bytes, and a pptx2svg render with
the right slide count, order and `data-pptx-id`s. In PowerPoint, every fixture with all the
operations applied (`acceptance_edits`) exports with one page per slide and no repair prompt,
and the rasterised pages show the changes: reordered and duplicated slides, the layout
slide's title, replaced pictures, the ungrouped rotated pair, styled hyperlinks, and a marker
picture on every page (`test_powerpoint_opens_structurally_edited_decks`).

## Phase E3 — full-state SVG

**Done.** `Slide.render_svg(full_state=True)` writes it, `Document.apply_svg(svg)` (or
`Slide.apply_svg`) reads it back. Code in `pptx_agent/fullstate/` — outside the core, because
the vocabulary is PowerPoint's: `model.py` (the vocabulary, read from XML), `emit.py`,
`apply.py`, `safe.py` (parsing untrusted input).

The two deliberate differences from pptx-svg held up:

1. **No resolved-value defect.** Every typed value is read from this library's *unresolved*
   model — the same XML the E1 facades read — so a theme fill is
   `data-ooxml-fill-scheme="accent1"` and nothing in the vocabulary can spell a resolved
   theme colour. Applied back, it goes through `Shape.set_solid_fill` and stays `a:schemeClr`.
2. **A raw-XML floor.** Each group carries `data-ooxml-xml`, so anything the vocabulary does
   not know (effects, 3-D, custom geometry, text-body properties...) survives the trip and can
   be edited there.

### How an SVG is laid out

pptx2svg draws the slide; the emitter adds attributes, nothing else, so the SVG draws exactly
as before — the full-state SVG with `data-ooxml-*` removed is canonically identical to the
plain one, and resvg renders byte-identical PNGs from both, on every fixture slide
(`tests/test_fullstate_emit.py`). Every shape in the slide's id index gets its group: a shape
pptx2svg does not draw (hidden) gets an empty `<g data-ooxml-unrendered="1">` in its z-order
place, so "missing from the SVG" can only ever mean "deleted". A group's children are nested
in its `<g>`, as pptx2svg draws them.

```xml
<svg ... data-ooxml-vocabulary="pptx-agent/1" data-ooxml-slide-id="257"
     data-ooxml-slide-part="ppt/slides/slide2.xml" data-ooxml-layout="ppt/slideLayouts/slideLayout1.xml"
     data-ooxml-theme="ppt/theme/theme1.xml" data-ooxml-slide-cx="12192000" data-ooxml-slide-cy="6858000">
  <g data-pptx-id="257.2" data-pptx-path="0" transform="translate(28, 45)"
     data-ooxml-kind="shape" data-ooxml-name="Shape 0" data-ooxml-cnvpr-id="2"
     data-ooxml-x="266700" data-ooxml-y="428625" data-ooxml-cx="57150" data-ooxml-cy="381000"
     data-ooxml-rot="0" data-ooxml-flip-h="0" data-ooxml-flip-v="0" data-ooxml-geom="rect"
     data-ooxml-fill="solid" data-ooxml-fill-scheme="accent2" data-ooxml-fill-mods="lumMod=75000"
     data-ooxml-line="set"
     data-ooxml-text='{"p":[{"algn":"l","c":[{"t":"Revenue grew ","sz":1800},{"t":"12%","b":true}]}]}'
     data-ooxml-xml="PHA6c3AgeG1sbnM6cD0i...">
    <rect .../>
  </g>
</svg>
```

### The vocabulary

All attributes are prefixed `data-ooxml-`. Lengths are EMU, angles degrees (exact decimals of
OOXML's 60000ths), colours unresolved. *Editable* attributes are applied through the API
named; *read-only* ones describe the shape and a change to them is refused.

| Attribute | On | Value | Applied through |
| --- | --- | --- | --- |
| `vocabulary` | root | `pptx-agent/1` — the reader refuses SVG without it | — |
| `slide-id`, `slide-index` | root | `sldId`; 0-based position (informational) | — (selects the slide) |
| `slide-part`, `layout`, `layout-name`, `master`, `theme`, `theme-name` | root | part names, `cSld@name`, theme name | checked: a mismatch is refused |
| `slide-cx`, `slide-cy` | root | slide size | checked |
| `kind` | shape | `shape`, `picture`, `connector`, `group`, `graphic_frame` | read-only |
| `name` | shape | `cNvPr@name` | `Shape.name` |
| `cnvpr-id`, `ph-type`, `ph-idx` | shape | raw id; placeholder type and index | read-only |
| `x`, `y`, `cx`, `cy` | shape | EMU, or `inherit` (from the layout) | `Shape.left`/`top`/`width`/`height` |
| `rot` | shape | degrees, `0` when absent | `Shape.rotation` |
| `flip-h`, `flip-v` | shape | `0`/`1` | `Shape.flip_h`/`flip_v` |
| `ch-x`, `ch-y`, `ch-cx`, `ch-cy` | group | child space | read-only (re-fitting owns it) |
| `geom` | sp, pic, cxnSp | `ST_ShapeType` preset, or `custom` | `Shape.preset` |
| `adj` | sp, pic, cxnSp with a preset | the explicit adjust values in document order, `adj1=30000 adj2=70000`; `""` for none (E5) | `Shape.adjustments` -- compared by effective value, so a default written out is no change; when `geom` changes and `adj` is left as it was emitted, it is the old preset's and ignored |
| `fill` | sp, pic, cxnSp, grpSp | `inherit`, `none`, `solid`, `gradient`, `image`, `pattern`, `group` | `fill = None`/`"none"`, `set_solid_fill`, `set_gradient_fill`, `set_image_fill` |
| `fill-scheme` / `-rgb` / `-sys` / `-prst` | solid fill | `accent1` / `4472C4` / `windowText` / `red` (one of them) | with `fill` |
| `fill-mods` | solid fill | modifiers in document order, raw: `lumMod=75000 lumOff=25000` | with `fill` |
| `fill-stops`, `fill-angle`, `fill-path` | gradient | `0 scheme:accent1;100000 rgb:FFFFFF`; degrees; `circle`/`rect`/`shape` | with `fill` |
| `fill-image`, `fill-pattern` | image, pattern fill | part name; preset (pattern is read-only) | with `fill` |
| `line` | sp, pic, cxnSp | `inherit` (no `a:ln`) or `set` | `line.clear()` |
| `line-w`, `line-fill`, `line-scheme`..., `line-mods` | outline | EMU; `none`/`solid`; colour as for fills | `LineFormat.width`/`visible`/`color` |
| `line-dash`, `line-cap`, `line-head`, `line-tail` | outline | preset; `rnd`/`sq`/`flat`; `triangle lg lg` | `LineFormat.dash`/`cap`/`head`/`tail` |
| `image` | picture | the media part shown | `Shape.replace_image` (with that part's bytes) |
| `graphic`, `chart`, `diagram` | graphic frame | `table`/`chart`/`diagram`/`ole`/`other`; chart part; diagram data part | read-only |
| `chart-data` | chart frame | JSON, below (E4) | `Chart.set_title`/`set_axis_title`/`set_legend`, `add_series`/`remove_series`, `add_category`/`remove_category`, `Series.set_name`/`set_values`, `Chart.set_category` |
| `diagram-nodes` | SmartArt frame | JSON, below (E4) | `Diagram.set_text`/`add_node`/`remove_node` |
| `text` | sp | JSON, below | `TextFrame.set_text`, `Paragraph.segment`, `Paragraph`/`Run` setters, `Run.set_hyperlink` |
| `table` | table frame | JSON, below | `Table.insert_row`/`delete_row`/`insert_column`/`delete_column`, `merge`/`split`, sizes, cell text/fill/borders |
| `cxn-begin`, `cxn-end` | connector | `<shape id> <site>` (`257.4 3`): a shape on the slide or a new one in the SVG, and its connection site; `none` when free (E5) | `Shape.connect`/`disconnect`, after every other edit; the connector is then routed |
| `rels` | shape | JSON: `[{"id","type","target","external"?}]` | re-points `r:*` ids (below) |
| `xml` | shape | base64 of the shape's OOXML (a group: its shell) | replaces the element |
| `unrendered` | shape | `1` for an empty placeholder group | — |

Colours inside JSON and gradient stops use one string: `scheme:accent1 lumMod=75000`,
`rgb:C00000`, `sys:windowText`, `prst:red`.

**Text** (`data-ooxml-text`) is `{"p": [paragraph...]}`. A paragraph has `c`, its content in
order — runs `{"t": "...", ...}`, fields `{"t", "fld": "slidenum"}`, line breaks
`{"br": true}` — and optionally `algn` (raw `ST_TextAlignType`), `lvl`, `marL`, `indent`,
`spcBef`/`spcAft`/`lnSpc` (`pts:600` in 1/100 pt or `pct:150000`), and `bu` (`"none"`,
`"picture"`, or `{"char": "•" | "num": "arabicPeriod", "start", "font", "color", "size"}`). A
run may have `b`, `i` (booleans), `u`, `strike` (raw), `sz` (1/100 pt), `latin`, `color`,
`fill` (a non-solid text fill, read-only), and `link` (`{"url"}` or `{"slide": sldId}`, with
`tip`). A key that is absent means *inherited*, exactly as in the file. Run indexes `r<n>` are
E1's: runs and fields count, breaks do not.

**Tables** (`data-ooxml-table`) are `{"cols": [w...], "rows": [h...], "merges": [[top, left,
bottom, right]...], "cells": [[cell...]...]}`, the full grid, hidden cells included. A cell may
have `text` (as above), `fill` (as for shapes, keyed `kind`, `scheme`...) and the borders
`left`, `right`, `top`, `bottom`, `diagonal_down`, `diagonal_up` (as for outlines, keyed
`kind`, `w`, `fill`, `scheme`...).

**Charts** (`data-ooxml-chart-data`, E4) are `{"types": ["bar"], "title": "...",
"axis_titles": {"category": null, "value": "..."}, "legend": true, "format": "General",
"categories": ["Q1", ...], "series": [{"name": "Revenue", "values": [3980, 4120, null]}]}`:
every key present, one value per category in every series, numbers finite or `null` for a
blank. `types` and `format` are read-only; a title or axis title of `null` means none (an
axis the chart lacks is absent from `axis_titles`). On the way back, series are lined up
with the chart's by name and categories by label, like table rows: missing ones are removed,
new ones added, the rest renamed, relabelled and revalued -- through the E4 edits, so the
workbook moves with the cache.

**Diagrams** (`data-ooxml-diagram-nodes`, E4) are `{"layout": "urn:...layout/vList2",
"nodes": [{"id": "{GUID}", "lvl": 0, "t": "Goals"}, {"id": "...", "lvl": 1, "t": "..."}]}`,
the nodes depth first; `t` is the node's text with `\n` between paragraphs. A node missing
from the list is removed (with its children, which must be missing too), a node without an
`id` is added where the list puts it (its parent is the nearest node above it one level
up), and changed text is set. Existing nodes cannot move or change level; `layout` is
read-only.

**Why JSON for text and tables, not child elements.** Text and tables are trees, and a
structured encoding would have to live *somewhere* in the SVG. Not on pptx2svg's `<tspan>`s:
they are lines after layout, not runs — wrapping splits a run, empty paragraphs leave no
trace, and a table cell's text is several `<text>`s — so there is no one-to-one place to hang
a run's properties. A parallel tree of invisible child elements would be new SVG content
that every renderer has to be trusted to ignore, and would break the property that only
attributes are added. One JSON attribute per shape keeps the drawing untouched, mirrors the
positional addresses agents already use (`p1/r0`), and is parsed by a standard, bounded JSON
parser rather than a second XML layer.

### Relationships: by reference, not embedded

`data-ooxml-xml` holds a shape's relationship ids (`r:embed`, `r:id`) verbatim, and
`data-ooxml-rels` says what each means: type, and target part (or external address). The
targets themselves are *not* embedded. An image's bytes would double the size of a picture
slide; a chart is a chain of parts (chart, workbook, colours, style) and a diagram five, none
of which fits an attribute faithfully. So the SVG is self-contained *relative to its
package*: it applies to its own deck, or any copy of it. On the way back, an id that still
means the described target in the slide is kept; otherwise the slide's existing relationship
to that target is reused, or one is added. An internal target must be a part of the package,
or the SVG is refused. A *new* shape's charts and diagrams are copied (as when a slide is
duplicated) while media, layouts' and slides' links are shared.

### Reading it back

`Document.apply_svg(svg, *, add_new=True, delete_missing=False, limits=None)`:

1. **Refuse what is not ours** — no vocabulary marker (plain pptx2svg output, Figma,
   Illustrator: `NotFullStateSvg`), an unknown slide, or a slide whose part, layout, master,
   theme or size differ. Everything is read and validated before anything changes.
2. **Deletions**, only with `delete_missing=True`: shapes on the slide whose id the SVG does
   not have. The default leaves them, so an SVG cut down to the shapes of interest is safe. A
   group whose child is still in the SVG is refused (that would be an ungroup).
3. **The raw floor**: a shape whose `data-ooxml-xml` differs from its XML is replaced by it,
   relationships re-pointed and any released ones reaped. A group's XML is its *shell* (the
   group without its child shapes, which are their own groups), so replacing it keeps them.
4. **Additions** (`add_new=True`): XML under an id the slide does not know (or none) becomes a
   new shape, placed where the SVG nests and orders it; a taken `cNvPr@id` is renumbered and
   a copied durable id dropped.
5. **Typed edits**: each attribute that differs from the document is applied through the API
   in the table. They are compared with the document *before* step 3, so changing only a
   shape's raw XML is not reverted by the stale typed attributes travelling with it, and a
   typed change on top of a raw change applies on top. Absent means "no opinion"; `inherit`
   removes an explicit value. Connections come last, once every shape is where the SVG
   puts it; then a connector added through the SVG, one whose raw ends changed, and those
   attached to a shape whose raw frame or geometry changed are routed again.

All of it is one undo step, an SVG that matches the document changes no byte (not even the
undo stack), and a failure part-way rolls everything back. It returns an `ApplyReport`:
`replaced`, `edited` (shape → features), `added` (SVG id → new id), `deleted`, `ignored`.

### Security

The SVG is untrusted input with XML inside it, and both layers get the same treatment
(`fullstate/safe.py`): any `<!DOCTYPE` or `<!ENTITY` is refused before parsing; lxml parses
with `resolve_entities=False`, `no_network=True`, `load_dtd=False`, `huge_tree=False`, and a
surviving entity reference is refused too; the SVG (64 MB), each base64 attribute (16 MB,
checked before decoding), each JSON attribute (8 MB), the shape count (20,000), a chart's
values (1,000,000), a diagram's nodes (10,000) and a shape's adjust values (64) are
bounded (`Limits`); base64 is decoded strictly; the shape XML must be a `p:sp`/`p:pic`/`p:cxnSp`/
`p:grpSp`/`p:graphicFrame` (a group's without child shapes); every typed value is validated
field by field against the schema's ranges and enumerations; and relationships may not point
at layouts, masters, notes or themes, or at parts the package does not have.
`tests/test_fullstate_security.py` shows XXE (in the SVG and inside `data-ooxml-xml`) and a
billion-laughs payload refused, the parser not expanding either even past the pre-check, and
the document untouched after every refusal.

**Acceptance.** For every fixture: emit every slide, apply to a fresh copy, save — every part
byte-identical. Edits made only in the SVG (a theme fill, a run, a position, a table cell,
rows and merges, split and linked runs, outlines and gradients) read back after save and
reopen; theme colours stay `a:schemeClr`; undo gives the original bytes; the validity checks
pass. A shadow added only in `data-ooxml-xml` arrives canonically identical. In PowerPoint,
every fixture with the acceptance edits exports with no repair prompt and one page per slide,
the marker square added from raw XML is on every page, and the retyped run and table cell are
in the PDF's text (`test_powerpoint_opens_svg_edited_decks`).

**Not in scope:** importing arbitrary SVG from Figma or Illustrator. That is a different project
— `path`→`custGeom`, `text`→`txBody`, no theme semantics — and inherently lossy.

**Left for later.** Slide-level state (background, transitions, timing, notes) and z-order
changes through the SVG; the refused cases under "Known approximations" that have an
obvious E1-style edit (pattern fills, custom dashes, percentage spacing); and the text
properties E1 does not expose yet, which can join the JSON the moment they do. (Charts and
SmartArt joined the vocabulary in E4.)

## Phase E4 — charts and SmartArt

**Done.** `Shape.chart` and `Shape.diagram`, first in `edit/chart.py`, `edit/workbook.py`
and `edit/diagram.py` and now in ooxml-edit's `ooxml_edit.charts` behind a small adapter
(see "The format-neutral core"); both in the full-state SVG.

```python
chart = deck.shape("257.29").chart
chart.categories, [s.name for s in chart.series], chart.series[0].values
chart.chart_types, chart.title, chart.axis_title("value"), chart.number_format
chart.series[0].set_value(2, 12.4)          # cache, cell -- one undo step
chart.series[1].name = "Gross margin"       # cache, header cell, table column name
chart.set_category(0, "Q1 FY25")
chart.add_category("Q4", [12.1, 43.2, 7.9]) # lines move, formulas and the table grow
chart.add_series("Net", [7.4, 7.5, 7.7, 7.9])
chart.remove_series("Operating margin")     # later columns close up, the table shrinks
chart.set_title("Margins"); chart.set_axis_title("value", "%"); chart.set_legend(False)

diagram = deck.shape("257.3").diagram
[(n.level, n.text) for n in diagram.nodes]
diagram.set_text(1, "Much faster edits")    # data model, and the cached drawing if exact
diagram.node(0).add_child("Smaller files")  # cached drawing dropped; PowerPoint re-lays out
diagram.remove_node(diagram.nodes[-1])      # with its children and presentation points
```

### Charts: cache and workbook together

Every data edit writes the chart part's caches (`c:strCache`/`c:numCache`: `ptCount`, one
`c:pt idx=` per point, gaps for blanks) **and** the embedded workbook, which is an OPC
package inside the package, opened with the core (`OpcPackage.open_embedded`) and edited
with lxml alone (`edit/workbook.py`): numbers as numbers, text as shared strings when the
workbook has a shared-string table (counts kept right) and inline strings when it does not,
cells in row and column order, `dimension` and row `spans` kept true. Adding or removing a
category or series also moves the cells of the chart's lines, rewrites the series formulas
(`c:f`, keeping their sheet quoting and `$` style), keeps per-point formatting (`c:dPt`,
`c:dLbl`) on its point, and keeps an Excel table over the data (`xl/tables`) in step: its
`ref` and `autoFilter` grow and shrink, a column inserted or removed inside it gets or loses
its `tableColumn`, and every column is named after its header cell, uniquely. A value written
over a formula drops the formula and the calculation chain. Untouched workbook parts keep
their bytes, entry order and compression.

A chart whose workbook is linked, an OLE object, missing or unreadable is edited in the cache
alone, with a `ChartDataWarning` saying Edit Data will show the old values. A layout the
structural edits cannot follow is refused with `ChartDataError` (see Known approximations).
Undo is byte for byte, the workbook included: it is one raw part, in the snapshot like any
other.

### SmartArt: what PowerPoint does with the cached drawing (measured)

Measured on PowerPoint for Mac 16 with hand-made variants of a deck PowerPoint had saved,
each exported to PDF through the oracle script and, where noted, saved again as `.pptx`:

| Variant | PDF shows | Saved again |
| --- | --- | --- |
| Data model text changed, cached drawing left stale | the **new** text | drawing rewritten with the new text |
| Cached drawing's text changed, data model left alone | the **old** (data model) text | drawing rewritten from the data model |
| Cached drawing removed (relationship and `dataModelExt` too) | the new text, laid out | a new drawing written |
| A node added with no presentation points, drawing removed or stale | the new node, laid out | presentation points and drawing written |
| A node removed with its presentation points, drawing removed or stale | the diagram without it | -- |
| No presentation points and no drawing at all (a hand-written data model) | the diagram, laid out | everything written |

So PowerPoint **ignores the cached drawing and re-lays every diagram out from its data
model on open**: a stale drawing never shows in PowerPoint, and a missing one is regenerated.
Other readers (pptx2svg among them) draw only the cache. Hence the approach: edit the data
model; patch the drawing where that is exact -- the shape that shows a node is found through
the node's `presOf` connection to a presentation point, whose `modelId` the drawing's `dsp:sp`
carries, and its paragraphs are rewritten only when they are, verifiably, the texts of the
nodes it presents in `destOrd` order (in "Vertical Bullet List" a child's text is one
paragraph of the parent's `childText` shape); otherwise drop the drawing. Adding and removing
nodes turned out to be tractable for the same reason: a node needs only its point, its two
transition points and a parent-of connection, and PowerPoint supplies the rest.

The fixture for this, `tests/fixtures/powerpoint-smartart.pptx`, was written by PowerPoint:
a hand-made data model for "Basic Block List" and "Vertical Bullet List" (the layout and
style definitions in that *input* came from PowerPoint's own resources, since a stub layout
definition makes PowerPoint fall back to Basic Block List), opened and saved by PowerPoint,
which wrote the presentation points, drawing and definitions.

### Edit Data: argued from the files

PowerPoint's AppleScript dictionary has no chart or chart-data objects, so its own view of
an embedded workbook cannot be asked for, and Excel was not driven. The evidence instead:

* the tests open every edited deck's workbook independently (`tests/xlsx.py`, not the
  library's reader) and check that every formula's cells hold exactly what its cache holds,
  that tables cover the data and are named after their headers, and that shared-string
  counts are right -- after every edit, on every chart in the corpus;
* PowerPoint, saving an edited deck again, kept every chart's caches and formulas as written
  (adding only a `formatCode` to a doughnut's cache) and left the embedded workbooks byte
  for byte -- it accepts both as they are and does not rewrite the workbook on open;
* what Edit Data opens is that workbook, so with the cells, formulas and tables agreeing it
  shows the edited data.

Opening an edited workbook in Excel (or Edit Data by hand) would close the gap.

### Acceptance

`tests/test_chart.py` runs thirteen edits (a value, a blank, a whole series, a category
label, a series name, categories added first and last and removed, series added first and
last and removed, titles, the legend) on all six charts in the corpus -- bar, line, doughnut
and radar -- through edit, save, reopen and read back; the independent workbook check; the
deck validity checks plus schema order in the chart part; and undo to the original bytes and
redo to the edited ones. pptx2svg draws the new values. `tests/test_diagram.py` does the same
for SmartArt text (including a bullet child in its parent's shape and a drawing that already
disagreed), adding and removing nodes. `tests/test_fullstate_charts.py` covers the SVG: both
attributes emitted, edits through them, identity, and refusals (bad numbers, `NaN`, wrong
lengths, read-only fields, unknown or reordered nodes, size limits), with the document
untouched after each. In PowerPoint (`test_powerpoint_opens_chart_and_smartart_edits`), every
deck with a chart or a diagram, with all of it applied -- a value far above each axis, a
series and a category renamed and added, a category removed, each title set through the SVG,
SmartArt text edited through the API and the SVG, a node added, one removed -- exports
unprompted, one page per slide; the new names, labels and node text are in the PDF's text,
the removed node is not, and each value axis reaches the new value. The E1-E3 oracle tests
pass on the new SmartArt deck too.

## Phase E5 — authoring

**Done.** New slides from a layout and new picture shapes landed in E2. Shapes: code in
`edit/creating.py` (the API and connector upkeep), `edit/authoring.py` (the XML, measured),
`edit/connectors.py` (routing), `edit/presets.py` with the compiled `edit/preset_sites.py`.
New decks: `edit/blank.py` (the parts, measured), `edit/deck.py` (templates, language,
slide size) and `edit/properties.py` (`core.xml`, `app.xml`).

```python
slide = deck.slides[0]                       # or a group: group.add_shape(...), child space
box = slide.add_shape("roundRect", 914400, 914400, 1828800, 914400, text="Plan",
                      adjustments={"adj": 30000}, fill="accent2", line={"width": 19050})
goal = slide.add_shape("ellipse", 5486400, 2743200, 1371600, 1371600, text="Ship")
box.adjustments["adj"] = 10000               # by name, raw units; del/reset for defaults
box.connection_sites                         # [(x, y), ...] on the slide, stCxn@idx order
arrow = slide.add_connector("elbow", (box, 3), (goal, 2), line={"tail": "triangle"})
goal.move_by(dy=-914400)                     # the connector is re-routed, still attached
arrow.begin_connection, arrow.end_connection # (shape, site)
arrow.connect(end=(other, 0)); arrow.disconnect("begin"); arrow.reroute()
slide.add_textbox(914400, 4572000, 3657600, 369332, "Notes")
table = slide.add_table(3, 4, 914400, 5029200, 7315200, 1097280).table   # then E1's API
box.bring_forward(); goal.send_backward()
```

Every creation is one undo step (text, fill, outline and adjust values included), byte for
byte; gets a `cNvPr@id` free on the slide (and free as a local id, so the new shape's id is
`<sldId>.<cNvPr@id>`); PowerPoint's name for it; and this library's durable id stamp, written
with its own `pa:` prefix. A shape added to a group goes in the group's child space and the
group re-fits, as when a child moves. PowerPoint does not put tables in groups, so
`group.add_table` is refused. `Shape.connection_sites` and the router read the preset's
`cxnLst`, which ooxml-common does not carry (it compiles the specification's adjustments,
guides and paths for drawing); `tools/derive_connection_sites.py` compiles each preset's
adjustment defaults, sites and the guides they need from the same SHA-verified
`presetShapeDefinitions.xml`, without redistributing it. A custom geometry's own `cxnLst`
is evaluated too.

### What PowerPoint writes for a new shape (measured)

Measured on PowerPoint for Mac 16 through its AppleScript -- `make new shape`/`text box`/
`connector`/`shape table`, `begin connect`/`end connect`, `adjustment_value`, `z order` --
saving each deck and reading the XML.

| What | PowerPoint writes | Written here |
| --- | --- | --- |
| Autoshape | `p:sp`, empty `p:cNvSpPr`, `a:prstGeom` with `a:avLst` | same |
| Its style | `lnRef idx=2` (`accent1` + `shade 15000`), `fillRef idx=1`, `effectRef idx=0` (`accent1`), `fontRef minor` (`lt1`) | same |
| Its text body | `<a:bodyPr rtlCol="0" anchor="ctr"/>` (default insets), one `<a:pPr algn="ctr"/>` paragraph ending in `endParaRPr lang` | same, the deck's default language |
| Adjust values | once any is set, *all* of the preset's, defaults filled in, in its order; a `star5` with `adj` alone is repaired on open | same |
| Text box | `txBox="1"`, `a:noFill`, no `p:style`, `<a:bodyPr vert="horz" rtlCol="0"><a:spAutoFit/>`, plain paragraph | same, plus `wrap="square"` (the default, written by UI text boxes) |
| Connector | `p:cxnSp`; when attached `<a:cxnSpLocks/><a:stCxn/><a:endCxn/>`; style `lnRef idx=2`, `fillRef idx=0`, `effectRef idx=1` (`accent1`), `fontRef minor` (`tx1`); `bentConnector3` with `adj1` written out | same |
| Table | `graphicFrameLocks noGrp="1"`, `tblPr firstRow="1" bandRow="1"`, style `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` (Medium Style 2 - Accent 1), even columns and rows, `a16:colId`/`rowId`, empty `tcPr` | same; ids `10001...` instead of random |
| Names | type and `cNvPr@id - 1`: "TextBox 3" is id 4. English names from PowerPoint's own table (`strman_f401`: "Rectangle", "Oval", "Straight Arrow Connector", "Elbow Connector", "Curved Connector", "TextBox", "Table", "Flowchart: Decision"...); presets it cannot insert are "Shape" | same |
| Ids | `a16:creationId` GUID on every new shape, `p14:modId` on a table | this library's stamp instead |
| Bring forward / send backward | one step past the next (previous) shape, overlapping or not | same |
| On re-save | an attached connector's frame, `stCxn` and `endCxn` kept as written | -- |

### Connectors: where PowerPoint bends them (measured)

180 elbow connectors were attached between a rectangle and a second rectangle (or an
ellipse) in fourteen arrangements -- apart, touching, overlapping in one axis or both, on
every side -- for every pairing of sites, plus moves and resizes; then read back. A curved
connector between the same sites always got exactly the elbow's frame and adjustments, so
one router serves both (`curvedConnector<n>` for `bentConnector<n>`). On a move or resize
PowerPoint keeps the sites and routes again (it picks the nearest sites only for "Reroute
Connectors"), which is what this library does on every geometry change made through it.

* An end leaves its site along the site's angle (`a:cxn@ang`, turned and mirrored with
  the shape); sites are the preset's `cxnLst`, so a rectangle has four and an ellipse
  eight, in `idx` order.
* Facing ends with room between them: a Z, bent halfway between the *ends* (`adj1`
  50000). At right angles, the end ahead and on the side it faces away from: an L
  (`bentConnector2`).
* Otherwise the route steps out a fixed 228,600 EMU (0.25 in, whatever the shape size), or
  to the middle of the gap when the end shape lies wholly ahead of the start shape; and it
  arrives beyond the end shape the same way, between the shapes when there is a gap on the
  side the end faces. A run that would cross the end shape's span clears it. Ends facing the
  same way make a U, round the end (or start) shape when it is in the way.
* The frame's rotation is the start's direction (0 right, 90 down, 180 left, 270 up) and
  the flips put the far corner on the end; PowerPoint spells two vertical cases `rot=270
  flipH` and `rot=90 flipH flipV`, an end level with the start takes the side the unflipped
  rotation gives, and extents are at least 12,700 EMU. Straight connectors are
  `straightConnector1` with flips only.

The 177 distinct routes are kept in `tests/connector_cases.py`; 169 are written exactly as
PowerPoint wrote them (frame, rotation, flips and adjustments within 2 EMU), and the eight
others are listed under "Known approximations".

### A new deck

```python
deck = Document.new()                          # 16:9; or size="4:3", or (width, height)
deck = Document.new(template="brand.potx", title="Q3", author="Ada", language="nl-BE")
[layout.name for layout in deck.layouts]       # Title Slide, Title and Content, ...
cover = deck.add_slide("Title Slide")          # placeholders inherit from layout and master
cover.shapes[0].set_text("Q3 review")
deck.title, deck.author, deck.created, deck.modified, deck.language, deck.slide_size
deck.set_slide_size("4:3")                     # scaled as PowerPoint's Ensure Fit scales
deck.save("q3.pptx"); deck.save_as_template("q3.potx")   # or save("q3.potx")
```

Every part is written here from measured values; nothing is copied from PowerPoint's
template. Measured on PowerPoint for Mac 16.106 by making a new presentation through its
AppleScript and saving it -- at 16:9, switched to 4:3 and to a banner, with a slide added
per layout and text typed into every placeholder, and with slides switched between sizes:

| What | PowerPoint writes | Written here |
| --- | --- | --- |
| Parts | presentation, presProps, viewProps, tableStyles, one master, eleven layouts, one theme, core and app properties, a thumbnail; no notes master, no slides | the same, without the thumbnail |
| Size | 12,192,000 x 6,858,000, no `type`; notes 6,858,000 x 9,144,000 | the same; `"4:3"` is 9,144,000 x 6,858,000, `screen4x3` |
| Theme | Office: dk2 `0E2841`, lt2 `E8E8E8`, accents `156082` `E97132` `196B24` `0F9ED5` `A02B93` `4EA72E`, links `467886`/`96607D`; Aptos Display over Aptos, with per-script faces; 1, 1.5 and 2 pt lines | the same values |
| Master | title (838200, 365125) 10515600 x 1325563; body (838200, 1825625) 10515600 x 4351338; date, footer, number at y 6356350. `titleStyle` 44 pt major font at 90%; `bodyStyle` 28/24/20/18 pt, `•` in Arial, `marL` 228,600 + 457,200 a level, `indent` -228,600, 10 then 5 pt before; `otherStyle` and `defaultTextStyle` 18 pt | the same |
| Layouts | Title Slide, Title and Content, Section Header, Two Content, Comparison, Title Only, Blank, Content with Caption, Picture with Caption, Title and Vertical Text, Vertical Title and Text; footers `idx` 10-12 | the same types, indexes, sizes, frames and list styles |
| Another size | master and layout frames stretched axis by axis; `sz`, `marL`, `indent`, `defTabSz`, `spcPts` times the smaller ratio, rounded; insets, percentages and `defaultTextStyle` kept (4:3: titles 33 pt) | the same rule |
| Resizing slides | content scaled uniformly by the smaller ratio and centred; explicit sizes scaled, text outside placeholders that inherited its size written scaled; line widths kept | the same (see Known approximations) |
| Slides added | `sldId` from 256, a layout relationship, no notes; `app.xml` counts slides, words and paragraphs and lists a title per slide | the same, `app.xml` as the deck is written |
| Saved again | no part added or removed; an empty `p15:sldGuideLst`; `endParaRPr` dropped where it repeats the run; field caches, `app.xml` names, zoom and `lastModifiedBy` rewritten | -- |

A deck of a slide per layout with the same text in every placeholder, made by PowerPoint
and by this library, exports to the same pixels at 144 dpi, at 16:9 and at 4:3, and every
placeholder resolves within two EMU of PowerPoint's.

**From a template** the deck keeps the template's masters, layouts, themes and notes and
handout masters, and loses its slides as `delete_slide` loses one -- notes, charts, media
and relationships reaped -- and then its sections, custom shows and thumbnail. A `.potx`'s
main part (`presentationml.template.main`) becomes a presentation. Creating a deck either
way is the base state: there is nothing to undo, and undoing everything after returns to it
byte for byte.

**Metadata.** `title`, `author`, `created`, `modified` and `language` are properties, each
an undoable edit (`core.xml` is added when a deck has none; the language is the
`defaultTextStyle`'s and the masters' `otherStyle`'s, and new slides and shapes take it).
`app.xml` is brought up to date as the deck is written (see Known approximations).

### Acceptance

`tests/test_authoring.py` runs each creator (autoshape with text, fill, outline and adjust
values; text box; the three connector kinds attached and then a shape moved; a table edited
with E1's row, column and merge edits; a shape in a group; z-order steps) through edit,
save, reopen and read back, the validity checks with schema order, and undo to the original
bytes and redo, and builds an acceptance slide with all of it on every fixture -- valid,
round-tripped, undone, drawn by pptx2svg with the API's ids, and its full-state SVG applied
to a fresh copy changing no byte. All 187 presets are made and drawn. A shape added through
the full-state SVG from raw XML takes typed attributes (position, fill, preset, text) on top.
`tests/test_connectors.py` holds the router to PowerPoint's routes and every connector to
its sites after moves, resizes, rotations, flips, preset and adjust-value changes, inside
scaled groups, and when re-connected, disconnected or orphaned by a deletion. In PowerPoint
(`test_powerpoint_opens_authored_shapes`) every fixture with the acceptance slide exports
unprompted with one page per slide; the rasterised page is red just outside every site a
connector ends on, after one shape was moved and the other resized; the new text is in the
PDF; and the deck saved again by PowerPoint keeps every `stCxn`/`endCxn`.

`tests/test_newdeck.py` pins a new deck to the measurements above (layouts, placeholders,
frames, text styles, theme; 4:3 and banner layouts; resizing), and builds a whole deck from
nothing -- slides on seven layouts with every placeholder filled, shapes, an elbow
connector, a table, a picture and the acceptance slide -- through the gates: the validity
checks on every part (schema order and orphans included), save and reopen with nothing
lost, undo to the new deck and redo, and every slide's full-state SVG applied to a fresh
copy changing no byte. pptx2svg draws each placeholder where its layout or master puts it,
in the theme's faces at the master's sizes (0.75 of them at 4:3). Templates: every fixture
as one, a `.potx` in and out. Connections and adjust values in the SVG
(`tests/test_authoring.py`, `tests/test_fullstate_security.py`): emitted, changed, attached
to a shape added in the same SVG, followed through a raw-XML move, and refused when
malformed, unknown or out of range, with the document untouched. In PowerPoint
(`test_powerpoint_opens_new_decks`) a deck from nothing at 16:9 and 4:3 and from a template
-- a fixture `.pptx`, and a `.potx` PowerPoint saved -- built up the same way exports
unprompted with one page per slide and its text, and saved again gains no part and keeps
every slide, layout, text and connection; `test_a_new_deck_matches_the_one_powerpoint_makes`
holds this library's new deck to PowerPoint's own, frame by frame, style by style and
pixel by pixel.

## Phase E6 — Markdown outline

**Done.** A deck is mostly visual, so Markdown is not its source of truth here: it is a
reading aid, cheaper for an agent than renders, and a drafting aid. Code in
`outline/markdown.py` (the dialect, a normalised AST, inline text written so it reads back),
`outline/read.py` (`to_outline`), `outline/draft.py` (`insert_outline`) and `edit/notes.py`
(notes pages and the notes master, measured); the format is specified in
`pptx_agent.outline`'s docstring.

```python
print(deck.to_outline())                     # or slides=[1, 2], ["s:257"]; ids=False; notes=False
new = Document.new(template="brand.potx")
new.insert_outline("""
# Quarterly review

Prepared for the board

# Where we are

- Revenue grew **12%**
  - Operating margin `11.9%`
- See [the report](https://example.org/q3)

Notes:

Open with the headline number.
""")                                         # at=, layout_map={"obj": "Agenda"}, images="figures/"
```

gives a Title Slide (title and subtitle) and a Title and Content slide with its notes, which
`to_outline` reads back as

```markdown
<!-- s:256 layout: Title Slide -->
# Quarterly review <!-- 256.2 -->

<!-- 256.3 ph: subTitle 1 -->
Prepared for the board

<!-- s:257 layout: Title and Content -->
# Where we are <!-- 257.2 -->

<!-- 257.3 ph: obj 1 -->
- Revenue grew **12%**
  - Operating margin `11.9%`
- See [the report](https://example.org/q3)

Notes:

Open with the headline number.
```

**Reading** (`Document.to_outline(slides=None, *, ids=True, notes=True)`) changes no byte.
The dialect and the id comments are docx-agent's: CommonMark with GFM tables and
strikethrough, read by markdown-it-py; `<!-- id ... -->` on the line above the block it
names, at the end of the line in a heading. A slide is a `#` heading -- its title
placeholder's text -- under its id, `hidden` if it is, and its layout's name. Each
placeholder and text shape is a block under its shape id and what it is (`ph: <type>
<idx>`, `empty`, `table`, `picture`, `chart`, `smartart`; the id alone for text outside
placeholders), in z-order, groups flattened. A paragraph showing a bullet -- its own or one
inherited through the layout, master and presentation, as pptx2svg resolves it -- is a list
item at its level (numbered ones ordered); the runs' own bold, italic, strikethrough,
monospace face and web or mail links are Markdown marks. A table is a GFM table whose cells
are `<table id>/cell<row>,<column>`, as in docx-agent; a picture `![alt text](media part)`;
a chart one line (`[chart: column; title: ...; series: ...; categories: ...]`); SmartArt a
list of its node text; speaker notes follow a `Notes:` line. Text is escaped so it reads back
as text, and inline Markdown is parsed back before it is used, dropping a mark CommonMark
cannot place rather than changing a character (docx-agent's self-check). `ids=False` writes
plain Markdown, a slide's blocks kept apart by `---`.

**Drafting** (`Document.insert_outline(md, *, at=None, layout_map=None, images=None)`)
reads the same format; a hand-written outline needs no comments. `#` starts a slide and its
heading goes to the title placeholder. Blocks fall into units -- a shape comment or `---`
(the column marker) starts one; a table or an image is one of its own -- and `Notes:` on a
line of its own starts the speaker notes. The layout is the deck's own:

| The slide holds | Role (`ST_SlideLayoutType`) |
| --- | --- |
| a layout the slide's comment names, which the deck has | that layout |
| nothing (an empty heading) | Blank (`blank`) |
| only its heading | Section Header (`secHead`); Title Slide (`title`) when it will be first |
| one text unit | Title and Content (`obj`); Title Slide, the text its subtitle, when first and not a list |
| two text units (`---`, or exactly two lists in one unit, split after the first) | Two Content (`twoObj`) |
| tables or pictures and no text | Title Only (`titleOnly`) |
| one text unit and one table or picture | Two Content |
| anything else | Title and Content |

A role is found by layout type, then PowerPoint's English name, then by the placeholders a
layout has, in the first master -- so a localised template ("Titel en object") and Google
Slides' (`tx`, `twoColTx`) work. `layout_map` overrides the choice, keyed by role, English
name or layout name. A unit naming a placeholder (`ph: body 1`) takes it, one the layout
lacks being added; the others fill free placeholders in the layout's order (text the body,
content and subtitle ones; a table a table or content one; a picture a picture or content
one, cropped to fill a picture placeholder and taking its `p:ph`, as PowerPoint inserts
one); what is left, and a unit whose comment names no placeholder, becomes a text box, table
(E5's `add_table`) or picture on a grid in the master's body area. Placeholders nothing went
into are removed unless the outline names them, and the shapes are stacked in the outline's
order. Bullets keep the placeholder's inherited ones where they agree; a paragraph in a
bulleted placeholder gets `buNone`, a list item where nothing bullets one gets PowerPoint's
`•` in Arial or `arabicPeriod` numbering. Runs carry the deck's language; code is Consolas.
Pictures come from `images`: a directory, a mapping, a callable, or the `Document` an
outline was read from. Charts and SmartArt are not drafted (an `OutlineWarning` says so).
The whole insertion is one undo step.

**The round trip** -- `to_outline`, `insert_outline` into a new deck from the same template,
`to_outline` again -- gives the same outline, up to ids (slide, shape and so cell ids), the
media part names in picture sources, and chart and SmartArt blocks, which are not drafted.
Guaranteed: slide order, layouts, hidden slides, the placeholder each text is in (and empty
placeholders the outline names), text, paragraph and list structure and levels, numbering,
bold, italic, strikethrough, code and web links, line breaks, tables' cell text and marks,
pictures and alt text, and speaker notes. Not carried, because the outline does not say
them: positions and sizes, fonts, colours and other run formatting, bullet characters and
numbering schemes (a numbered list comes back `arabicPeriod`), indents and alignment, empty
paragraphs, a level that skips one (written one deeper than its predecessor), slide-jump
links, table styles, merges and header-row settings, shapes without text, connectors and
group structure. Without ids the outline is stable too, the layout and placeholders then
chosen again by the rules.

### Speaker notes (measured)

The deck from nothing has no notes master, so drafting notes may add one. Measured on
PowerPoint for Mac 16 through its AppleScript: a new presentation, a slide, text set on its
notes page's notes placeholder, saved.

| What | PowerPoint writes | Written here |
| --- | --- | --- |
| Notes page | `sldImg` (no `idx`, `noGrp noRot noChangeAspect`), notes `body idx="1"` (whatever the master's `idx`), `sldNum sz="quarter"` at the master's `idx` with a `slidenum` field; no frames; relationships to the notes master and the slide, and from the slide | same |
| Notes master | `bgRef idx=1001` `bg1`; header and date 2,971,800 x 458,788 along the top; slide image at (685,800, 1,143,000) 5,486,400 x 3,086,100 with a 1 pt black outline; notes at (685,800, 4,400,550) 5,486,400 x 3,600,450 prompting five levels; footer and number 2,971,800 x 458,787 along the bottom; 12 pt, right-hand ones right-aligned, bottom ones anchored at the bottom; `notesStyle` nine 12 pt levels, `marL` 457,200 a level; listed in `notesMasterIdLst` | same, English prompts |
| Its theme | a part of its own, byte for byte the slide master's | a copy of the slide master's theme |

### Acceptance

`tests/test_outline.py`: reading changes no byte on every fixture and every id it writes
resolves; the outline of each construct (lists and levels, marks, links, tables with their
cell ids, pictures, charts, SmartArt, notes) is pinned; text that looks like Markdown reads
back as text, and inline marks read back or are dropped, never the text. Drafting: the
layout rules, by type in a localised template and in Google Slides' and placeholder-less
templates, overridden by a named layout and `layout_map`; bullets, numbering, marks, links,
the deck's language; a table, a picture fitted or cropped into a picture placeholder;
notes with the notes master as measured, and none added to a template that has one; `at`;
refusals that change nothing; one undo step. The gates on a deck drafted from a hand-written
outline into a deck from nothing, and on every fixture's own outline drafted into a new deck
from that fixture: the validity checks, save and reopen with nothing lost, undo to the
undrafted deck byte for byte and redo, every slide's full-state SVG applied to a fresh copy
changing no byte, and a pptx2svg render with each run's text in the shape it was put in.
The round trip is stable on every fixture, with ids and without, and again from the drafted
deck. In PowerPoint (`test_powerpoint_opens_drafted_decks`) each of those eight drafted
decks exports unprompted, with one page per slide and every shape's text on its page; and
saved again by PowerPoint (`test_powerpoint_keeps_a_drafted_deck_as_drafted`) each reads as
the same outline -- layouts, placeholders, text, lists, marks, links, tables, pictures and
notes.

**Shared with docx-agent.** The dialect, the id-comment conventions and the inline writer
are docx-agent's, re-implemented here so that neither repository changed. What could be
shared, as a small `ooxml-markdown` package next to ooxml-edit (Markdown is not an OPC
concern, and ooxml-edit stays lxml-only): the parser configuration (`parser()`, with
footnotes an option), the normalised AST (`parse`, `normalise_inline`, `differences`), the
inline writer with its escaping and self-check (`inline_markdown`, `escape`, `destination`),
and the id-comment helpers (`comment`, `comment_text`). The format-specific halves --
docx-agent's style map and readers, this outline's slides and placeholders -- stay where they
are.

## Usability (end-to-end pilot)

A Sonnet agent did a real task with pptx-agent as a black box -- restate Q3 revenue,
operating profit and net profit through a Japanese results deck
(`real-financial-report.pptx`), in its text, tables and charts, with every derived figure
-- using only the README, `help()` and `dir()`. It got every number right, but found the
API late and worried most about a wrong index silently writing a plausible number. What
changed:

* **A usage guide**, [docs/USAGE.md](docs/USAGE.md): what is editable (tables and charts
  fully, a chart's cache and workbook together), and a short example per common task.
  `tests/test_usage.py` runs every example and checks the output each one shows.
* **Labels, strictly.** `Table.cell_by_label`, `find_row`, `find_column`;
  `chart.series["name"]`, `category_index`, `Series.set_value("Q3", ...)`,
  `Series.value`, `items`. A label that names nothing or more than one thing raises
  `LabelError` listing the candidates; full-width forms and case are forgiven only when
  nothing matches exactly, and only when that is unique.
* **The outline is Markdown, escaped.** `to_outline` says so, and escapes a line's start
  only where CommonMark would read a block (`11.9%` and `+0.6pt` are no longer escaped;
  docx-agent's writer still escapes them, harmlessly -- both read back the same).
  `outline_blocks()` and `find_text()` give every block's raw text with its address
  (SmartArt nodes resolve as `<id>/node<k>`). `set_text` and the paragraph and run
  setters warn with `MarkdownEscapeWarning` when new text has an escape or `**bold**`
  markup that the old text lacked; the check is narrow (no warning for Windows paths,
  `\\server`, `3*4**2`), and the text is still written as given.
* **`ids=False` separators.** The pilot saw `---` appear between two blocks after an edit.
  The separators already depended on structure only; the `---` was `difflib`'s autojunk
  treating frequent lines as junk. A test pins both.
* **`Document.validate()`**: the structural checks as `Problem(code, part, detail)`, as
  in docx-agent. **`Chart.number_formats`**: what decides precision on screen (data labels,
  value axis), since `number_format` -- the values' cached format -- does not.
* **Docstrings**: every public class and method has a one-line example
  (`tests/test_docstrings.py`); `TableCell.text` says it writes.
* **The task again**, as `examples/restate_figures.py`, from the guide alone: 132 lines
  against the pilot's 199, every table and chart edit by label, every unchanged figure read
  from the deck. Its test asserts the answer key; with `PPTX_AGENT_PILOT_OUTPUT` it compares
  every text block and chart value with the pilot's own output (equal).

Left: `Shape.text` still reads a line break as `"\n"` (E0's reading, pinned by a test;
`text_frame.text` and the blocks give `"\v"`). The pilot's other notes were about the task
(which derived figures to recompute, the rounding convention), not the library.

## Usability (full trial)

Twelve PowerPoint tasks -- two runs each, agents using only the README, USAGE.md, `help()`
and `dir()` -- succeeded ten times out of twelve. Both failures were silent, and the
friction logs agreed on the rest. What changed:

* **Templates.** A `.potx` opened with `Document.open` and saved as `.pptx` kept the
  template content type, which PowerPoint refuses outright; `validate()` was clean. `save`
  now writes the type the extension needs (`.pptx`/`.pptm` a presentation, `.potx`/`.potm`
  a template), `save_as_template` refuses a `.pptx` name, `validate(target=...)` reports a
  `package-type` mismatch, and `Document.open` warns (`TemplateOpenedWarning`). Oracle:
  the trial's file, made again the same way, opens unprompted.
* **Fit, as PowerPoint shows it.** `normAutofit` text is shrunk only when PowerPoint edits
  it; on opening, it draws the stored `fontScale` -- none, for anything a generator wrote --
  so an overflowing body draws full size. pptx2svg shrank it (being changed there), and an
  agent that checked its render shipped an overflowing slide. `Shape.text_fit()` lays the
  text out with ooxml-common's measurement over pptx2svg's resolved model, at the stored
  scale: the height needed and available, the overflow, effective sizes and line counts,
  the autofit mode and scale. An overflow counts once it passes the inkless bottom fifth of
  the last line box. `Document.overflows()` adds overlaps with titles and filled
  placeholders, and shapes past the slide's edge. Oracle: six cases at and around the
  boundary agree with PowerPoint's PDF (a wider probe agreed on 40 of 40). It uses
  ooxml-common's private layout helpers (`_estimate_text_height`, `_text_area`...); a
  public "measure a text body" function there is proposed.
* **Speaker notes**: `Slide.notes` and `notes_frame`, `"<sldId>/notes"` addresses in
  `resolve`, `find_text` and `outline_blocks` (kind `notes`), the address in the outline,
  `duplicate_slide(..., notes=...)`; a first note writes the notes page and master E6
  measured. Oracle: PowerPoint saves the notes back unchanged.
* **`insert_outline`** drafts free tables and pictures below the title's actual bounds (the
  trial's `TITLE_ONLY` title sat inside the master body); a table frame's width and height
  rescale its columns and rows in proportion, as dragging it in PowerPoint does.
* **Introspection**: `Layout.placeholders` (type, idx, name, bounds, default sizes per
  level) and a read-only `Layout.shapes`; `Placeholder(type, idx)`; `deck.theme` (colours,
  fonts), `Color.resolve`; `Run.effective_size`/`effective_font`, `Paragraph.effective`,
  held to pptx2svg's resolution on every fixture; `TextFrame.autofit` and `font_scale`;
  `Shape.image_size` and `replace_image(keep=..., anchor=...)`;
  `Shape.connection_site("right")`, the site order documented, `LineFormat.start`/`end`;
  `Chart.workbook_values()` (proposed for `ooxml_edit.charts`).
* **One way to name a slide**: a `Slide`, an `sldId` or `"s:<sldId>"` everywhere;
  `move_slide(slide, to=...)`; errors that say which numbers are 1-based and which are
  positions. Edit calls return what they edited.

---

## Suggested order

```
E0 (done) ──┬─▶ E1 (done)          ──┬─▶ E4 (done)
            ├─▶ E2 (done)          ──┤
            └─▶ E3 (done)            └─▶ E5 (done) ──▶ E6 (done)
```

E1 and E2 are independent and both unblock real agent work; E3 turned out to lean on both, since
every typed attribute it reads back is applied as an E1 or E2 edit. E4 needs E1's text API (and
extends E3's vocabulary). E5 needs E1 and E2. E6 needs E5's new decks (it drafts into one)
and E1's text API.

---

## Testing strategy

Three layers, in decreasing order of authority and increasing order of availability:

1. **The PowerPoint oracle** (`pytest -m oracle`) — drives the real application via
   `pptx2svg/tools/powerpoint_export_pdf.applescript` and exports to PDF. The only authoritative
   answer to "will it open?". A deck with E0 and E1 edits exports; every fixture with the
   E1 edits applied exports; a damaged deck is rejected (so a pass means something); and a
   good deck still exports straight afterwards; every fixture with the E2 structural
   operations exports with the right page count, and (with `pypdfium2` installed) the
   rasterised pages are checked for a marker picture put on every slide; every fixture with
   edits made only through its full-state SVGs exports with the right page count, a red
   marker square added from raw XML on every page, and the retyped run and table cell in
   the PDF's text; every fixture with a chart or SmartArt, with E4's edits applied, exports
   with the right page count and the new chart labels, series names, node text and values
   in the PDF; every fixture with E5's acceptance slide exports with the right page count,
   connectors meeting their shapes after a move (red just outside every site, on the
   rasterised page) and the new text in the PDF -- and, saved again by PowerPoint
   (`oracle.save_as_pptx`), keeps every connector's `stCxn`/`endCxn`; a new deck from
   nothing at 16:9 and 4:3 and from a template (a fixture, and a `.potx` PowerPoint saved,
   `oracle.save_as_potx`) exports with one page per slide and, saved again, gains no part;
   and a new deck matches the one PowerPoint makes itself (`oracle.new_powerpoint_deck`),
   frame by frame and pixel by pixel; a deck drafted from a hand-written outline, and every
   fixture's own outline drafted into a new deck from that fixture, exports with one page per
   slide and every shape's text on its page, and saved again by PowerPoint reads as the same
   outline. From a checkout that
   is not a sibling of
   `pptx2svg` (a git worktree, say), point `PPTX2SVG_ORACLE_SCRIPT` at the script.

   Getting this reliable took more than calling `osascript`, and the findings are in
   `tests/oracle.py`:

   | Symptom | Cause | Fix |
   | --- | --- | --- |
   | Every export fails with −9074, including known-good files | A damaged file raised an app-modal repair dialog; PowerPoint refuses all file access until it is cleared | Dismiss it, then retry |
   | The export script cannot dismiss the dialog itself | It is blocked inside `open` and never regains control | Recovery must run in a separate process |
   | Escape does not close the dialog | Only a real button click does | Click Cancel, matched across localisations (`Annuleren` on a Dutch install, `Cancel` on English), falling back to button 1 of an untitled two-button sheet |
   | `count of presentations` says the app is healthy when it is not | A wedged PowerPoint answers `0` while still refusing every file | Never treat it as sufficient evidence of recovery |
   | Restarting clears the wedge, then it returns | PowerPoint reopens the document it was killed over and raises the dialog again | Prefer dismissal; restart is only the no-Accessibility fallback |
   | The first export after a restart fails | `open` sent while the app is still launching is refused instantly with −9074 | Relaunch explicitly and poll until it answers |
   | `open` hangs until the AppleEvent times out (no dialog on screen) | The path was made a `POSIX file` *inside* the `tell application` block | Coerce it before the `tell`, as the export script does |

   Requires **both** macOS permissions: Automation (to drive PowerPoint) and Accessibility (to
   click the dialog). Without Accessibility it still works, but each rejected file costs a
   PowerPoint restart. Both paths must be under `$HOME` — the sandbox rejects anything else
   with −9074.
2. **Structural validity** (`tests/test_validity.py`) — well-formedness, no dangling
   relationships, every part content-typed, no `Override` for a missing part, no orphaned
   part, a consistent slide list, children in schema order, zip intact, still renderable —
   all after the full E0 + E1 + E2 edit set on every fixture, after each E2 operation on
   its own (`tests/test_structure.py`), and after each E5 creator and E5's acceptance slide
   (`tests/test_authoring.py`), and on every part of a new deck, from nothing or a
   template (`tests/test_newdeck.py`), and on decks drafted from outlines
   (`tests/test_outline.py`). Runs everywhere and
   catches the failure modes that actually cause repair prompts.
   Charts are checked one level down as well: the embedded workbook of every edited chart is
   opened independently (`tests/xlsx.py`) and its cells, formulas, tables and shared strings
   must agree with the chart's caches (`tests/test_chart.py`).
3. **Round-trip fidelity** (`tests/test_roundtrip.py`) — byte-identity for untouched parts,
   canonical-form equality for edited ones, and undo back to the original bytes after the
   whole E1 edit set, both as one batch and step by step. The gate everything else rests on.
   The full-state SVG has its own (`tests/test_fullstate_apply.py`): emit every slide, apply
   to a fresh copy, and every part is byte-identical.

The fixture corpus is deliberately mixed-provenance: PowerPoint, Google Slides and Keynote
exports, and since E4 one deck PowerPoint itself wrote from a hand-made SmartArt data model
(`powerpoint-smartart.pptx`) -- no other fixture has SmartArt. The duplicate-`cNvPr@id` case that drove the id design came from a Google Slides deck,
and would not have appeared in a PowerPoint-only corpus.

---

## Non-goals

- Rendering. That is `pptx2svg`'s job, and this project depends on it rather than duplicating it.
- Importing arbitrary SVG (see E3); the reader refuses SVG it did not write.
- Macros and VBA.
- Pixel-exact PowerPoint reproduction — the fidelity target lives in `pptx2svg`'s own roadmap.
