# What the deck tools support

The tools in `pptx_agent.tools`, as a model sees them. Lengths are points, colours theme names
(`accent1`, `tx1 lumMod=75%`) or `#RRGGBB`. Every changing call is one undo step and returns
facts about the slides it touched (see [GUIDANCE.md](GUIDANCE.md)).

## Supported

| Feature | Tool / action | Notes |
|---|---|---|
| Open a deck or template; a new deck, from nothing or from a template | `open_document`, `new_document` | inputs are blobs the application registered; a template's slides are dropped from a new deck |
| Save as a deck, a template or a text outline | `save_document` (`pptx`, `potx`, `outline`) | refuses new validation problems; lists overflows and collisions still there |
| Close a document | `close_document` | the session holds a few documents at a time |
| The deck at a glance | `describe` | slides with ids, titles, layouts and content areas; theme colours, fonts, roles and tints; layouts and their placeholders; comment counts |
| Read slides | `ppt_read_slides` (`outline`, `geometry`, `svg`) | outline Markdown with ids; every shape's box, fill, outline and text; a compact SVG view |
| Read a text input (CSV, Markdown, plain text) | `read_blob` | a page at a time |
| Find and replace text | `find_text`, `replace_text` | slides, notes, tables and SmartArt; `expect: one` or `all` |
| Look at slides | `render` | PNG, at most 4 a call; says which text the image leaves out for want of a font (`missing_glyphs`) |
| Fit, collisions, validation, design facts | `check` | overflows, collisions (`boxes` too), shapes off the slide, problem facts, design facts, validation |
| Undo and redo | `undo` | one call is one step; refs come back with the state |
| Several calls as one step | `batch` | in order, all or nothing, checks once at the end |
| Slides from a layout, with title, body and notes | `ppt_add_slide` | the deck's own layouts and placeholders |
| Slides from a Markdown outline | `ppt_draft_slides` | into the deck's layouts |
| Duplicate, move, delete a slide; find one by title | `ppt_manage_slides` | a duplicate does not take its original's comments |
| Text of shapes, table cells and speaker notes | `ppt_set_text` | plain text keeps formatting; the text spec states runs, bullets, levels and the frame |
| Run, paragraph and frame formatting | `ppt_format_text` | bold, italic, underline, size, font, colour; alignment, levels, bullets, spacing; insets, anchor, wrap, autofit; table cells too |
| Autoshapes and text boxes | `ppt_add_shape` | every preset; `measure: true` sizes text first and adds nothing; `fit_height` |
| Position, size, rotation, flips, fill, gradient, outline, geometry | `ppt_set_shape` | only the fields given change |
| Connectors that stay attached | `ppt_add_connector` (`straight`, `elbow`, `curved`) | glued to a shape's side, or between points; arrowheads |
| Z-order, group, ungroup, duplicate, delete | `ppt_arrange` | |
| Pictures: insert, replace | `ppt_add_picture` | from an image blob; replacing keeps the frame, height or width |
| Tables: new; cells, rows, columns, merges, widths, heights, fill, borders | `ppt_add_table`, `ppt_edit_table` | cells by row and column or by their labels |
| Charts: new from data; read; values, categories, series, titles, axis titles, legend, data labels, gap width | `edit_chart` | column, stacked column, bar, stacked bar, line, pie, scatter, radar (lines, as PowerPoint inserts one; not the filled or marker radar); the embedded workbook changes with the chart |
| SmartArt: node text, adding and removing nodes | `edit_smartart` | diagrams already in the deck; PowerPoint lays the diagram out again on opening |
| The theme's colours and fonts | `ppt_set_theme` | every theme reference in the deck follows |
| Document properties | `set_properties` | title, author, language, subject |
| Comments: threads on a slide or a shape, replies, resolving | `ppt_comments` (`list`, `add`, `reply`, `resolve`, `reopen`, `edit`, `delete`) | PowerPoint's modern (threaded) comments |
| Design facts of a slide | `ppt_design_facts` | palette, sets of like shapes and legends, empty regions, alignment, vocabulary, text sizes, lines over text: facts, no verdicts |
| Layout facts after an edit | `checks.layout` on changing calls | near-alignment, uneven gaps, outlier text sizes, labels far from their markers, each with its exact fix |

## Not supported

- **Animations and transitions.**
- **SmartArt creation:** build the graphic with shapes and connectors (`ppt_add_shape`,
  `ppt_add_connector`); existing SmartArt can be edited.
- **Icons** (PowerPoint's icon library).
- **Video and audio.**
- **Sections** (grouping slides into named sections).
- **Hyperlinks:** no links between slides, and no listing or editing of links. (A web address
  can be put on a run of text through the text spec's `hyperlink`; nothing more.)
- **Action buttons** and other click or hover actions.
- **Slide master and layout editing:** using the deck's existing layouts is supported;
  changing a master or a layout is not (the theme's colours and fonts can be set).
- The older, non-threaded comments (`ppt/comments/comment<n>.xml`) are not read or written.
