# What works today

Moved from the README. How to do each of these, with an example the tests run, is in
[USAGE.md](USAGE.md); what the agent tools support is in
[SUPPORTED.md](../src/pptx_agent/tools/SUPPORTED.md).

| | |
| --- | --- |
| Open, save, byte-identical round trip; templates (`.potx`) in and out | ✅ |
| Stable ids, collision handling, stamping | ✅ |
| Position, size, rotation (incl. placeholder inheritance); where a shape is drawn, rotation, flips and a connector's route included | ✅ |
| Text: replace keeping mixed formatting, paragraphs, runs, bullets that hang, hyperlinks; insets, vertical anchor, wrapping, autofit | ✅ |
| Speaker notes: read, write, find, address (`<sldId>/notes`) | ✅ |
| Fills and outlines in theme colours; the theme's colours and fonts, resolved and set; which colours do which job, and their tints | ✅ |
| Tables: cell text and formatting, rows, columns, merges, cells by label; rows as drawn, grown to fit their text, and how many fit on the slide (to split a long table) | ✅ |
| Charts: values, categories, series, titles, legend, with cache and embedded workbook together; series and points by label | ✅ |
| SmartArt: node text, adding and removing nodes | ✅ |
| Shapes, pictures, connectors that stay attached, groups, z-order | ✅ |
| Comments: PowerPoint's threaded comments on a slide or a shape -- read, add, reply, resolve, reopen, edit, delete | ✅ |
| Slides: add from a layout, duplicate, reorder, delete; titles, and a slide by its title; layouts and their placeholders | ✅ |
| Draft slides from a Markdown outline; read a deck as one | ✅ |
| Fit: does the text fit, as PowerPoint will show it, and how close a line is to wrapping; collisions (text over text, lines through text, and on request label boxes that overlap); measuring text before building; effective font sizes | ✅ |
| One text spec -- paragraphs, runs, bullets, frame -- that builds a shape's text and measures it the same way; the box height its text needs; a slide's content area | ✅ |
| Agent tools (`pptx_agent.tools`): the deck through tool calls for a model, on `ooxml_edit.tools` -- lengths in points, many shapes per call, refs, facts after every edit (with layout facts about the shapes an edit touched); golden transcripts do the trial tasks with the tools alone. What the tools support: [SUPPORTED.md](../src/pptx_agent/tools/SUPPORTED.md); for the application's thinking layer: [GUIDANCE.md](../src/pptx_agent/tools/GUIDANCE.md) | ✅ |
| Layout help in the library: PowerPoint's Align and Distribute on what is drawn; shapes in a row, column or grid; labels beside their anchors within a stated distance, avoiding collisions; data-to-position scales (numbers, dates, bands) with ticks; copying shapes and groups to another slide or deck, repeated, with new text (no longer tools: no model called them) | ✅ |
| Design facts, no verdicts: palette, colour-coded sets and legends, empty regions, alignment and near-misses, shape vocabulary, text sizes, lines over text; colours that are not theme colours | ✅ |
| A compact SVG view of a slide for a model (pptx2svg's agent view, with these ids); SVG authoring in the library (`edit.svgprofile`; the experimental `ppt_draw` tool was removed) | ✅ |
| Undo / redo / batch | ✅ |
| Render to SVG and PNG with matching ids; a full-state SVG that applies back | ✅ |

Tables and charts are fully editable. Lengths are EMU everywhere (914,400 per inch;
`Pt(6)` says 6 points as EMU), font sizes points. What is measured, approximated or still to
come is in [ROADMAP.md](../ROADMAP.md).
