# Changelog

pptx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail.

## Unreleased

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
