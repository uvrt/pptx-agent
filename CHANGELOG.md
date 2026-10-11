# Changelog

pptx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail.

## Unreleased

- An SVG picture's fallback PNG is drawn from the SVG at the size PowerPoint reads from its
  root (ooxml-common 0.12.0's `svgimage.with_pixel_size`), in place of this module's own
  rewrite: the same pixels for `px`, `pt`, `pc`, `in`, `cm` and `mm`, and now PowerPoint's
  for an `em` (12 px, or the root's font-size), a percentage or no size (the `viewBox`'s
  size; 500 px without one, where the placed size was used), and a Latin-1 or UTF-16 SVG
  is re-encoded rather than refused. The root's other attributes and the SVG's content are
  no longer re-serialised through lxml.
- Rendering in a daemonic process (a Celery prefork worker's child; ooxml-edit 0.17.0):
  the slide renderer is a registered worker task (`@worker_task`), so ooxml-edit's subprocess
  workers run it there -- in parallel, with deadlines that kill -- instead of in-process.
- Slide master and layout editing (production feedback round 3, G5), measured in PowerPoint
  for Mac 16. `deck.masters` / `deck.master()` give `Master` objects (`id` `M0`, `name`,
  `part_path`, `layouts`, `slides`, `shapes`); `Layout` gets `id` (`L2`), `master`, `type`,
  `slides`, `placeholder("body")`, `rename`, `duplicate` and `delete`. `Layout.shapes` and
  `Master.shapes` are now the editable `Shape` a slide has (they were read-only
  `LayoutShape` summaries, still importable), addressed `L<n>.<id>` / `M<n>.<id>` where `n`
  is the `sldLayoutId` / `sldMasterId` less 2147483648 -- ids PowerPoint keeps on save, as it
  keeps a slide's `sldId`; `deck.shape`, `deck.resolve` and the text API take them. Text,
  run and paragraph formatting, fill, position, delete, replace_image, and `add_shape`,
  `add_textbox`, `add_picture` (SVG too), `add_connector`, `add_table` on layouts and
  masters. `shape.set_prompt(text)` / `shape.prompt`: a layout placeholder's prompt, the
  first level written and `hasCustomPrompt="1"` set, as typing in Slide Master view does
  (typing any text into a layout placeholder sets it too; formatting does not).
  `deck.find_text(..., scope="slides" | "layouts" | "masters" | "all")` and
  `outline_blocks(scope=)` (`TextBlock.part`; a placeholder holding only a field is left
  out). `deck.set_footer(text, slide_number=, date="auto" | text | None, slides=,
  skip_title=)` as Insert > Header & Footer writes it (Apply to All and Apply), and
  `add_slide` then gives a new slide the layout's footers, as New Slide does.
  `add_layout`, `duplicate_layout`, `rename_layout`, `delete_layout(reassign_to=)` as Slide
  Master view does them (a layout in use is refused, as PowerPoint disables it).
  `shape.inherited_by()` lists what inherits from a layout or master placeholder; deleting
  one that something inherits from is refused. `deck.affected_slides(target)`.
- Tools: `ppt_set_text`, `ppt_format_text`, `ppt_set_shape`, `ppt_arrange` take `L2.5` /
  `M0.7`, and the add tools `slide: L2` / `M0`; every result lists the slides an edit
  `affects` (`checks.affects`) and checks them. `find_text` / `replace_text` take decks'
  `stories`: `layouts`, `masters`, `all` (ooxml-edit 0.16.0); a replace that finds nothing on
  the slides says when the text is on a layout or master. `describe` lists the masters, the
  layouts' ids, their placeholders' ids and custom prompts, and the text on them that is not
  a placeholder. The deferred `ppt_edit_master` (`set_prompt`, `set_footer`, `add_layout`,
  `duplicate_layout`, `rename_layout`, `delete_layout`). Undo by slide (`scope`) finds a
  change to the slide's layout or master and refuses it as `entangled`
  (`DocumentFormat.undo_shared`, ooxml-edit 0.16.0). Every definition 11,779 estimated
  (guard 11,800), the core 3,440 (about 5,091 counted); the strict set unchanged. Golden p17:
  a template made ready for 2027 in Dutch.
- SVG pictures (production feedback round 3, I1): `slide.add_picture` takes an SVG and
  stores it as PowerPoint 2016 and later store an inserted SVG, measured in PowerPoint for
  Mac 16: the PNG in `a:blip`, the SVG behind its `asvg:svgBlip` extension (`xmlns:asvg` on
  the element), a `ppt/media/imageN.svg` part (`image/svg+xml`) behind an ordinary image
  relationship, the picture named "Graphic N" with its aspect unlocked. The PNG fallback is
  `fallback=` or the SVG drawn at the placed size and `fallback_dpi` (300) by pptx2svg (an
  `ImportError` names `pptx-agent[png]` without it); the native size is the SVG's width and
  height, else its viewBox, at 96 dpi. The SVG is parsed as untrusted data first (no DTD,
  no entities, at most 16 MiB), stored once, and kept as given. `picture.vector` reads it,
  `picture.set_vector(svg | None)` attaches, replaces or removes it (one undo step), and
  `picture.replace_image(png, svg=...)` (or an SVG `image`) keeps both images in step. The
  outline marks the picture `picture svg`, the geometry `image.svg`. `ppt_add_picture`
  takes `image/svg+xml` blobs, to insert and to replace, and makes the fallback. In
  PowerPoint the result opens unprompted, draws the vector, and offers the Graphics Format
  tab (Graphics Fill, Convert to Shape).
- Align and distribute in the tools (round 3, A5): `ppt_arrange` actions `align` (`edge`,
  `to`: `shapes`, `slide` or `content_area`) and `distribute` (`axis`, `within`: `shapes`
  or `content_area`, optional `gap`) on its targets, through `pptx_agent.edit.arrange`; the
  content area is `Slide.content_area`. Each is one undo step and returns the boxes. For
  house rules that require equal spacing (GUIDANCE.md). The every-definition guard rises
  to 11,400 (11,309 estimated, +196), the core unchanged; ppt_arrange is now sent
  non-strict (eight optional parameters) and ppt_add_table strict in its place. Golden
  p16: the Halden approach slide, four phases with SVG icons spaced across the content
  area, 7 calls.
- Chart formatting (production feedback round 3, G1-G4; ooxml-edit 0.15.0): a chart's axis
  scale (a 1-5 maturity axis by 1, categories top-down for a ranked bar chart), gridlines,
  series and point colours, lines and markers, a column series drawn as a line (target
  markers over actual columns), chart text, and points labelled with their own text --
  optionally from the workbook, as PowerPoint's Value From Cells -- through `Chart` and
  `edit_chart`'s new actions; `slide.add_chart` takes `radar_filled` and `bubble`,
  `category_order`, `gridlines` and `bubble_scale`. A series and a point are found by name
  and label as `series_named` and `category_index` find them. The every-definition guard
  rises to 11,200 (11,113 estimated, +646), the core unchanged. Golden p15: a maturity gap
  chart and an effort-against-impact bubble matrix labelled R1-R6, 10 calls.
- The outline lists each shape once: `slide.shapes` lists a group and then its members, and
  `to_outline` (and `ppt_read_slides` `outline`, `outline_blocks`, `find_text`,
  `replace_text`) walked the groups as well, so every grouped shape came twice. A group's
  blocks now stand between `<!-- 256.5 group "Agenda item 1" -->` and `<!-- /256.5 -->`
  (docx-agent's closing form), nested as the groups are; `insert_outline` reads the marks
  as nothing. The other readers of `slide.shapes` (fit, collisions, layout facts, design
  facts, comments) already met each shape once, and a test now holds them to the same facts
  grouped or not. Production feedback round 3 (D3).
- Empty placeholders: `slide.empty_placeholders()` names the placeholders that hold
  nothing -- no text, only whitespace or only empty paragraphs, and no field; one that took
  a picture, table, chart or SmartArt holds it; date, footer, slide number and header left
  out; titles and subtitles counted. Measured in PowerPoint 16: an empty one shows "Click to
  add text" (or "...title", the content icons) while editing and no prompt in the export; a
  whitespace-only one a stray bullet, exported too; several empty paragraphs nothing; a
  placeholder's fill or outline is drawn either way. `slide.delete_empty_placeholders()`
  deletes them in one undo step, keeping an empty title unless `titles=True` and one that
  draws a fill or outline. The tools report them as `empty_placeholders` (`shape`, `type`,
  `idx`) in every changing call's checks and in `check` (with `fit`, the default), and
  `save_document` lists those left under `unresolved`; `ppt_manage_slides`
  `delete_empty_placeholders` deletes a slide's (A3).
- Table styles and Table Design flags: `deck.table_styles` (the deck's own from
  `tableStyles.xml`, then PowerPoint's 74 built-in styles, each `{id, name, builtin}`, the
  default marked), `table.style` (read, and set by name or id), `table.flags` and
  `table.set_flags(first_row=, last_row=, first_column=, last_column=, banded_rows=,
  banded_columns=)`; `add_table(style=)` takes a name too, and refuses an id neither the
  deck nor PowerPoint has. The tools: `ppt_add_table`'s `style`, `ppt_edit_table`'s
  `set_style` and `set_flags`, a table's `style` and `flags` in its results, and
  `describe`'s `theme.table_styles` (the default, the deck's own, how the built-in names
  go). Every definition 10,269 -> 10,467 estimated tokens; the core unchanged (A4). A
  golden, p14, builds a chart slide on a content layout and deletes its empty body, and a
  table in the template's own style without banding.
- `render` and `check` name slides by id as well as number: `slides` takes `s:257`, a
  `$ref` (the dispatcher replaces it) or a number from 1 (also written `"2"`, as a strict
  decoder writes it now that the items are strings), so a loop renders the slide it read
  even after another loop added or moved slides. Each image is still labelled with both,
  `slide 3 (s:257)`. A shape's address or an id the deck lacks is refused on `slides[i]`
  (`invalid_arguments` / `not_found`, listing the ids). Production feedback round 3 (A2).
- `render` rasterises, and `check` measures, a snapshot with the deck's lock let go: the
  bytes and version are taken under the lock, the worker draws them (or a copy of the deck
  opened from them is measured and validated) while other calls may edit the deck, and the
  images and facts are cached under the snapshot's version, which the result reports. An
  edit to slide 2 no longer waits for a render of slide 1 (the `a6` repro with worker
  processes: 333 ms -> about 60 ms; its own time alone). Inside a batch, unchanged. Needs
  ooxml-edit 0.14.0 (A6).
- Python 3.14 and 3.15: CI runs the suite on both, on Linux, macOS and Windows, and the
  classifiers declare them. `requires-python` stays `>=3.10`. Python 3.14 made
  `forkserver` Linux's default start method (`fork` before); the worker pool names `spawn`
  itself, so nothing changes, and a test now runs the render tool, with the application's
  font folders, in a worker under every start method the platform has. CPython 3.14's
  Windows builds deflate with zlib-ng, whose bytes differ from zlib's for the same entries
  (both valid; what a package holds is unchanged), so each golden transcript also records
  its packages' `content_sha256` (every entry's name, metadata and bytes, a chart's
  embedded workbook by its own entries; `tests/zip_content.py`), checked everywhere, and
  the archives' `sha256` is checked where deflate is zlib's.
- The application's own font folders, for the tools and the library: `Toolbox(font_dirs=...)`
  (or a session's own, ooxml-edit 0.13) sets `Document.font_dirs` on every deck the
  session opens or makes, and every render -- in the worker process too, the folders
  resolved where the toolbox runs and handed over -- and every measurement of text (fit,
  overflow, near-wrap, a table's row heights, the per-call facts, `save_document`'s and
  the design facts, `measure_text`) uses them: a face there is drawn with and, where
  pptx2svg's tables do not measure its family as itself, measured from. Without them
  `OOXML_FONT_DIRS` is read. The folders key `render`'s and `check`'s caches. No tool
  definition or prompt changes. Needs ooxml-common 0.8.0, ooxml-edit 0.13.0 and a
  pptx2svg with `ConvertOptions.font_dirs`.
- A table whose rows grow past the slide is reported. A row's stored height is only a
  minimum; `drawn_bounds` used the stored heights, so a 12-row table of wrapping text cut
  off at the slide's bottom passed every check. Rows are now measured with pptx2svg's
  table layout (`pptx2svg.table_row_heights`), an empty paragraph drawn as a line, and
  held to PowerPoint's PDF (`tools/table_rows_probe.py`, five tables, every row within
  0.1 pt). `Table.drawn_row_heights`, `Table.rows_fitting()` (`RowsFit`: how many rows fit
  above a line, to paginate a table) and an `off_slide` fact with `rows_past` and
  `rows_fit` in every changing call, `check` and `save_document`.
- `render` says which text its image leaves out for want of a font (`missing_glyphs`:
  face, script, sample; pptx2svg's `glyphs-missing`), with a note that the deck is
  unchanged. The README installs `pptx2svg-fonts` and says what it holds.
- Golden p5 replays a committed input (`tests/fixtures/generated/trial`) instead of
  drawing it with the installed Pillow.
- `undo` takes `scope`, a slide (`256`, `s:256`): only that slide's latest change is undone,
  with what the slide owns (notes, charts, media), and later changes on other slides stay;
  refused with `entangled` when the change shares a part with a later one (two added slides
  share the presentation part). For parallel loops, one per slide. Needs ooxml-edit 0.12.0.
- `edit_chart` `add` makes a radar chart (`chart_type: "radar"`), as PowerPoint inserts one:
  lines in the theme's accents, the legend at the top (measured on Office for Mac 16).
  Without `position`, a new chart's legend is where the application puts it. Needs
  ooxml-edit 0.11.0.
- Golden p13 (a vendor comparison on a slide): a radar chart from a CSV.

## 0.0.1 -- 2026-10-08 (initial public release)

Developed 2026-09-12 and 2026-10-01 to 2026-10-08:

- **Document core:** lossless OOXML editing with stable shape ids, now on
  [`ooxml-edit`](https://github.com/uvrt/ooxml-edit); undo, redo and batches.
- **Editing API:** text with formatting, fills and outlines in theme colours, slides,
  pictures, groups, connectors routed as PowerPoint routes them, tables, charts with
  their cache and embedded workbook in step (`slide.add_chart`), SmartArt nodes,
  threaded comments, new decks from nothing or from a template, Markdown outlines.
- **Feedback:** rendering through pptx2svg with matching ids, a full-state SVG that
  applies back, fit and collision checks, text measured before building, layout helpers
  and data-to-position scales, design facts.
- **Agent tools** (`pptx_agent.tools`) on `ooxml_edit.tools`, documented in
  `SUPPORTED.md` and `GUIDANCE.md`, with golden transcripts of the trial tasks.
