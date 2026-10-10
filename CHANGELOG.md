# Changelog

pptx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail.

## Unreleased

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
