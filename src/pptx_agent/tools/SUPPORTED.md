# What the deck tools support

The tools in `pptx_agent.tools`, as a model sees them. Lengths are points, colours theme names
(`accent1`, `tx1 lumMod=75%`) or `#RRGGBB`. Every changing call is one undo step and returns
facts about the slides it touched (see [GUIDANCE.md](GUIDANCE.md)).

## Supported

| Feature | Tool / action | Notes |
|---|---|---|
| Open a deck or template; a new deck, from nothing or from a template | `open_document`, `new_document` | inputs are blobs the application registered; a template's slides are dropped from a new deck |
| Save as a deck, a template or a text outline | `save_document` (`pptx`, `potx`, `outline`) | refuses new validation problems; lists overflows, collisions and empty placeholders still there |
| Close a document | `close_document` | the session holds a few documents at a time |
| The deck at a glance | `describe` | slides with ids, titles, layouts and content areas; theme colours, fonts, roles and tints; the deck's table styles (its default, its own, how the built-in names go); masters and layouts with their ids (`M0`, `L2`), placeholders (and their ids and custom prompts) and the text on them that is not a placeholder (a legal line); comment counts |
| Read slides | `ppt_read_slides` (`outline`, `geometry`, `svg`) | outline Markdown with ids; every shape's box, fill, outline and text; a compact SVG view |
| Read a text input (CSV, Markdown, plain text) | `read_blob` | a page at a time |
| Find and replace text | `find_text`, `replace_text` | slides, notes, tables and SmartArt; `expect: one` or `all`; `stories: layouts`, `masters` or `all` reach the text on layouts and masters (a footer year, a legal line, a prompt), addresses `L2.5`, `M0.7` |
| Look at slides | `render` | PNG, at most 4 a call; says which text the image leaves out for want of a font (`missing_glyphs`) |
| Fit, collisions, validation, design facts | `check` | overflows, collisions (`boxes` too), shapes off the slide, empty placeholders ("Click to add text"), problem facts, design facts, validation |
| Undo and redo | `undo` | one call is one step; refs come back with the state; undo by slide (`scope`) refuses a change to a layout or master the slide shows (`entangled`: other slides show it too) |
| Several calls as one step | `batch` | in order, all or nothing, checks once at the end |
| Slides from a layout, with title, body and notes | `ppt_add_slide` | the deck's own layouts and placeholders |
| Slides from a Markdown outline | `ppt_draft_slides` | into the deck's layouts |
| Duplicate, move, delete a slide; find one by title; delete a slide's empty placeholders | `ppt_manage_slides` | a duplicate does not take its original's comments; `delete_empty_placeholders` keeps an empty title |
| Text of shapes, table cells and speaker notes | `ppt_set_text` | plain text keeps formatting; the text spec states runs, bullets, levels and the frame |
| Run, paragraph and frame formatting | `ppt_format_text` | bold, italic, underline, size, font, colour; alignment, levels, bullets, spacing; insets, anchor, wrap, autofit; table cells too |
| Autoshapes and text boxes | `ppt_add_shape` | every preset; `measure: true` sizes text first and adds nothing; `fit_height` |
| Position, size, rotation, flips, fill, gradient, outline, geometry | `ppt_set_shape` | only the fields given change |
| Connectors that stay attached | `ppt_add_connector` (`straight`, `elbow`, `curved`) | glued to a shape's side, or between points; arrowheads |
| Z-order, group, ungroup, duplicate, delete | `ppt_arrange` | |
| Align and distribute | `ppt_arrange` (`align`, `distribute`) | align an edge or centre to the targets' bounds, the slide or the content area; distribute with equal gaps between the outermost targets or across the content area, or with a fixed gap; one undo step, the boxes returned |
| Pictures: insert, replace | `ppt_add_picture` | from an image blob; replacing keeps the frame, height or width |
| SVG pictures | `ppt_add_picture` with an `image/svg+xml` blob | stored as PowerPoint 2016 and later store an inserted SVG: the vector, which PowerPoint can recolour, and a PNG fallback made for it; insert and replace |
| Tables: new; cells, rows, columns, merges, widths, heights, fill, borders; style and Table Design flags | `ppt_add_table`, `ppt_edit_table` (`set_style`, `set_flags`) | cells by row and column or by their labels; a style by name or id, the deck's own (a template's) or any of PowerPoint's 74 built-in styles; header row, total row, first and last column, banded rows and columns |
| Charts: new from data; read; values, categories, series, titles, axis titles, legend, data labels, gap width, axis scales, gridlines, series and point colours and markers, a column series as a line, chart text, point labels | `edit_chart` | column, stacked column, bar, stacked bar, line, pie, scatter, radar (lines), filled radar, bubble, as PowerPoint inserts each; point labels from the workbook as PowerPoint's Value From Cells; the embedded workbook changes with the chart |
| SmartArt: node text, adding and removing nodes | `edit_smartart` | diagrams already in the deck; PowerPoint lays the diagram out again on opening |
| The theme's colours and fonts | `ppt_set_theme` | every theme reference in the deck follows |
| Slide master and layout shapes: text, formatting, fill, position, size, delete, pictures; new text boxes, shapes, pictures (SVG too), connectors and tables on them | `ppt_set_text`, `ppt_format_text`, `ppt_set_shape`, `ppt_arrange`, `ppt_add_shape`, `ppt_add_picture`, `ppt_add_connector`, `ppt_add_table` with addresses `L2.5`, `M0.7` (`slide: L2` or `M0` to add) | the same edits a slide's shapes take; every result lists the slides it `affects` and checks them; a slide's own placeholder text and overrides stay; deleting a layout or master placeholder that slides or layouts inherit from is refused, with the list |
| A layout placeholder's prompt text ("Typ hier uw bevinding") | `ppt_edit_master` `set_prompt` | the first level (`\n`: the next ones), formatting and levels kept, `hasCustomPrompt` set, as typing it in Slide Master view does; shown in an empty placeholder while editing, never in a show or export |
| Header & Footer: footer text, slide number, date (updating or fixed), not on title slides | `ppt_edit_master` `set_footer` | Apply to All (masters, every layout and every slide, as PowerPoint writes it) or Apply (`slides`); turning one off removes it from the slides |
| Layouts: add (a custom layout), duplicate, rename, delete | `ppt_edit_master` `add_layout`, `duplicate_layout`, `rename_layout`, `delete_layout` | as Slide Master view does; a layout in use is not deleted unless its slides move to another (`reassign_to`); names stay unique on a master; ids (`L2`) are the file's own and PowerPoint keeps them |
| Document properties | `set_properties` | title, author, language, subject |
| Comments: threads on a slide or a shape, replies, resolving | `ppt_comments` (`list`, `add`, `reply`, `resolve`, `reopen`, `edit`, `delete`) | PowerPoint's modern (threaded) comments |
| Design facts of a slide | `ppt_design_facts` | palette, sets of like shapes and legends, empty regions, alignment, vocabulary, text sizes, lines over text: facts, no verdicts |
| Layout facts after an edit | `checks.layout` on changing calls | near-alignment, uneven gaps, outlier text sizes, labels far from their markers, each with its exact fix |

## Not supported

- **Animations and transitions.**
- **SmartArt creation:** build the graphic with shapes and connectors (`ppt_add_shape`,
  `ppt_add_connector`); existing SmartArt can be edited.
- **Icons** (PowerPoint's icon library): no icon library is provided. An icon the
  application brings as SVG can be inserted (SVG pictures, above).
- **Video and audio.**
- **Sections** (grouping slides into named sections).
- **Hyperlinks:** no links between slides, and no listing or editing of links. (A web address
  can be put on a run of text through the text spec's `hyperlink`; nothing more.)
- **Action buttons** and other click or hover actions.
- **Masters themselves:** a master's shapes, its layouts and Header & Footer can be edited
  (above); adding or deleting a slide master, the master's text styles (the levels' sizes and
  bullets) and backgrounds are not. Text fit is measured on the slides a layout or master
  edit affects, not on the layout itself.
- The older, non-threaded comments (`ppt/comments/comment<n>.xml`) are not read or written.
