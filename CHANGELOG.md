# Changelog

pptx-agent has not been released to PyPI. The public repository starts from a single
snapshot commit; the development history before it is summarised here.
[`ROADMAP.md`](ROADMAP.md) has the detail.

## Unreleased

- The tightest spacings on a slide, measured (round 12, Q1): `check` with `spacing` (pt,
  e.g. `8`; ooxml-edit 0.56.0) adds `spacing`, nearest first and at most 12 a slide
  (`spacing_left_out` counts the rest) -- gaps `between` two neighbours facing on an
  `axis` (each shape's nearest, within the same container), `inside` a filled or outlined
  shape (the content's `padding` to each `edge`, `also` others as near), and `clearance`
  to the footer, slide number and date placeholders or past the layout's `content_area`
  (`gap` negative). Text is measured where it is drawn (the collision measurement's
  extent), shapes with rotation and groups resolved; overlaps (collisions), full-slide
  backgrounds, lines, connectors and a shape's own insets are left out. `Slide.spacing`
  in the library (boxes only without pptx2svg). Measurements, no verdicts; GUIDANCE says to
  cite the gap. The core 3,628 -> 3,664 estimated tokens, every definition 13,064 ->
  13,099 (guard 13,100); a reviewer's core 2,298 -> 2,333 (guard 2,335), all 4,546 ->
  4,581 (guard 4,585); the strict set unchanged.

- Table row heights in a run of slides (round 11, P1-P3). A row's height is a minimum, so
  `ppt_add_table`'s box `h` is an upper bound: each row takes its text's height
  (`Table.fit_rows`, measured with pptx2svg's table layout) and the frame their sum, and
  the summary says what the table takes -- `it takes 96 pt of the 276 pt box, each row its
  text's height (header row 19.2 pt); fill_box: true shares the rest among the body rows`.
  `fill_box: true` shares the room left among the body rows, the header keeping its text's
  height; rows taller than the box run past it as before, the summary saying by how much
  and naming `continue`. A header row is no longer stretched to a fifth of a tall box.
  `insert_row` takes its `like` row's height and `delete_row` takes its row's out of the
  frame, as they did; `set_heights`/`set_widths` take `rows`/`columns` to set only those;
  every table result lists `heights`. `ppt_set_shape`'s `h` on a table warns that its rows
  were rescaled, naming the header row's new height. `ppt_compare_slides` compares matched
  tables inside: the header row's drawn height, fill and text size, the body rows' height
  (with their row counts as context, so a three-row and a six-row table agree), column
  widths, cell insets and borders -- every differing table with its fix
  (`{"tool": "ppt_edit_table", "calls": [{"target", "action": "set_heights", "rows": [0],
  "sizes": [19.2]}]}`, or `ppt_format_text` items); a table's box height is no longer
  compared. Every definition 13,010 -> 13,064 estimated tokens (guard 13,070), the core
  unchanged (3,628); a reviewer's core 2,298 (guard 2,300), all 4,546 (guard 4,550); the
  strict set unchanged. Goldens p1 and p14 re-recorded: their tables' rows now hug their text.

- A contact sheet whose tiles do not fit the 2,576 px edge is drawn with narrower tiles,
  not refused (round 10, O4; ooxml-edit 0.55.0): too wide, the widest that fit the columns
  (two columns of the call's 1,280 px: 1,276 each); too tall, the widest for the slides'
  shape, before anything is drawn. `sheet.shrunk` (`{"from", "why"}`) and the summary say
  so; only tiles that would not fit even at 160 px are refused.
- `ppt_edit_table`'s `set_cells` takes `row_label` with `col`, and `row` with `col_label`
  (round 10, O6): each side by label or by position; both for one side, or neither, is
  refused naming the field.

- The render worker calls pptx2svg's public `render_svgs`, `rasterise_svgs` and
  `uses_host_fonts` (pptx2svg's round-10 API) instead of its private `_render`,
  `_rasterise` and `_host_fonts`; the same calls, the same PNGs. Rasterising itself got
  cheaper in pptx2svg and ooxml-common 0.25.0 (round 10, item 3): an effect's filter region
  fits its reach (a full-slide shadow drew four slides of blur), fewer whole-SVG scans per
  slide, the font mapping indexed; and a font folder holding both Aptos and Aptos Display
  draws body text in Aptos again.

- `render` draws a call's slides over every worker at once, and keeps each image by its
  content (round 10, O1 and O3; ooxml-edit 0.52.0, ooxml-common 0.24.0): the slides not
  held are split into as many chunks as there are workers (`render_views`, one per
  worker, a chunk's deck parse its fixed cost), and each image is keyed by its SVG and
  everything else its pixels depend on (`raster_key`, `raster_fingerprint`: every font
  file the rasteriser may be handed with its time and size, the embedded faces by digest,
  whether glyphs are checked, the renderers' code), kept in the toolbox's memory store
  (`Toolbox.images`, byte-bounded, every session's and view's). A repeat of an unchanged
  deck is answered from the store with no worker asked; after an edit, the workers draw
  the SVGs again and rasterise only the slides whose key changed. `cached` lists the
  slides not drawn by the call, as before; a contact sheet reuses the tiles `render`
  drew and says which (`cached_tiles`). PNGs are `render_slides`' byte for byte.
- The worker opens the deck's bytes with pptx2svg directly: no `Document.open` and
  `to_bytes` per call; the main process makes a version's bytes once
  (`DocumentEntry.snapshot`, `check` too).
- A deck's embedded faces are decoded once: when a deck embeds fonts and no worker has
  decoded them yet, one chunk is drawn first, alone, and the faces it decoded go to the
  other chunks (and later calls) with their tasks (`Toolbox.decoded_fonts`), so no
  worker decodes them again. Memory only.
- The image budget is asked before anything is drawn (`call.check_images`), so a refused
  render or sheet costs nothing. `FORMAT.warm_up` (`render.warm_up`) warms each new worker
  -- the imports, the Office font index, resvg's system-font load -- outside any call's
  deadline. Definitions are unchanged.

- `render` with `sheet` (round 9, N1; ooxml-edit 0.49.0): up to 30 slides, layouts or
  masters as one image, tiled in order, each labelled with its number and id
  (`3 · s:258`; `· hidden` after a hidden slide's; a layout `L2 Title Only`) -- a contact
  sheet. `{"columns": 4, "width": 480}`; `{}` is 4 columns of 360 px (1,480 px wide).
  `sheet.width` is each tile's, 160-1280 (the call's `width` stands in); `columns` 1-8.
  Refused, naming the bound and what to do: more than 30, a sheet over 2,576 px wide (the
  widest tile that fits) or tall (how many slides fit, before anything is drawn). The tiles
  are drawn in one worker call at the tile width, from the render cache where there; the
  sheet is cached whole. `missing_glyphs` and `template_note` as for a render; render
  without `sheet` and more than 4 slides now says `sheet` takes 30. Definitions: +64 to the
  core (3,563 -> 3,627, about 5,368 counted) and every definition (12,944 -> 13,008; guard
  13,010); a reviewer's core guard 2,225 -> 2,285, all 4,475 -> 4,535; strict set unchanged.
- `edit_chart` (round 9, N2; ooxml-edit 0.48.0): a `scale` without `axis` -- in
  `set_axis_scale` or `add` -- is the value axis where it is the chart's one axis that takes
  bounds and units (column, bar, line, radar, dot); a scatter's x and y, a date axis or a
  secondary value axis must still be named, and the refusal names them, as it does for a
  scale of only `reverse` or `number_format`. Every "needs" refusal names what to give:
  `set_value needs category: 'A', 'B' or 'C' (or a number from 0 to 2)` (the first 12 and
  how many more), the series, the chart's axes, what `add` needs (`chart_type`'s kinds,
  `data`'s and `box`'s shape). Definitions unchanged.
- `ppt_format_text` on a table (round 9, N3): a table formats every cell, as `ppt_edit_table
  format_cells` does by default, where it was refused "holds no text"; `256.7/row0` and
  `256.7/col1` format one row's or column's cells (one undo step a call). The `target`
  description says so: +9 to every definition (12,935 -> 12,944; guard 12,940 -> 12,945),
  none to the core (the tool is deferred); strict set unchanged.
- What a name that is no deck tool means (round 9, N5; ooxml-edit 0.47.0): "no tool
  'ppt_add_chart': use edit_chart with action add", in a batch's op too, with the best few
  tools in `valid_options` instead of every one. ooxml-edit reads the tools' names and
  actions; `TOOL_HINTS` (`FORMAT.tool_hints`) adds the deck's names it misses or settles
  wrongly: deleting a shape, chart, picture or table is `ppt_arrange delete`, moving or
  resizing one `ppt_set_shape` (not `ppt_manage_slides move`), a slide's title or notes
  `ppt_set_text` (not `edit_chart set_title`), a text box `ppt_add_shape`, an image
  `ppt_add_picture`, a line `ppt_add_connector`, a slide number `ppt_edit_master
  set_footer`. A batch op may carry `comment`/`note` (N4), kept in the call log. Data and
  tests only: definitions and the strict set unchanged.
- Layouts and masters read as slides (round 8, M3): `ppt_read_slides` takes `L2` and `M0`
  in `slides`, as `render` does, beside slide numbers and ids. `compact` lists their shapes
  under their addresses (`L2.6`) with a `placeholder` column (type and idx: `title`,
  `content 1`, `dt 10`; slides' placeholders in the same read labelled too), sizes inherited
  from the master; `geometry` gives each placeholder's `idx` in such a read. `outline` and
  `svg` are of slides and refuse a layout, naming `describe` and `render`; an unknown key
  is `not_found` and names the layouts and masters there are (render's too).
- Renamed shapes in a compared run (round 8, M4): `ppt_compare_slides` joins name groups
  that are each on some slides only, and gives a name group a shape left on a slide without
  it, by kind and nearest box (overlapping or within 36 pt, closest pair first, each once);
  a group whose names differ lists `name` as its first difference (the majority's,
  `other` the rest). A renamed shape far from its place is still `missing_on` and `only`.
- Definitions: `ppt_read_slides`' `slides` description adds "; layouts L2, masters M0"
  (core 3,556 -> 3,563 estimated, about 5,274 counted by the proxy) and
  `ppt_compare_slides`' description "name," (deferred): every definition 12,926 -> 12,935,
  the guard 12,930 -> 12,940; a reviewer's core 2,212 -> 2,220 (guard 2,215 -> 2,225),
  every definition 4,460 -> 4,469 (guard 4,465 -> 4,475). The strict set is unchanged.
- Several agents on one deck (round 8, M1 and M2; ooxml-edit 0.46.0): a reviewer's and a
  builder's views of one shared session at once (`session.view(roles=..., context=...)`),
  each held to its own roles, in a batch too; a reviewer's prompt with the role line
  decided over the tools sent (`system_prompt(tools=...)`). Tests and the tools README
  only: no deck tool or definition changed.
- Tool roles (round 7, item 9; ooxml-edit 0.45.0): the deck tools are tagged for the
  additive roles `review`, `resolve`, `accept` and `edit` (`Toolbox(roles=...)`,
  `toolbox.session(roles=...)`). `ppt_comments` takes its actions' roles from
  `shared.COMMENT_ACTION_ROLES` -- list, add, reply `review`; resolve, reopen, delete (and a
  reply with `resolve: true`) `resolve`; editing a comment's text every role's (the user's
  decision: its author is not known) -- and joins the core of a role set without `edit`,
  as `ppt_compare_slides` does; `ppt_manage_slides find_by_title` is `review`'s; measuring
  with `ppt_add_shape` or `ppt_add_table` stays with the tool, `edit`. No deck tool settles
  tracked changes, so `accept` adds nothing to a deck toolbox. The default, every role,
  sends the definitions byte for byte as before (core 3,556, every definition 12,926
  estimated). A reviewer (`["review"]`) is sent 15 tools, the core 2,204 estimated (about
  3,262 counted by the proxy), every definition 4,452; `["review", "resolve"]` 2,212 and
  4,460 -- guards `REVIEW_CORE_GUARD` 2,215 and `REVIEW_ALL_GUARD` 4,465
  (tests/test_tools_roles.py). Claude's strict set for a reviewer: `ppt_comments` and
  `ppt_manage_slides` (find_by_title only); the default's and `["review", "edit"]`'s are
  unchanged. The online check counts both reviewer cores and sends their strict tools.
- A run of like slides compared (round 7, L3): `ppt_compare_slides` (deferred, `ppt_slides`), library
  `Document.compare_slides` and `pptx_agent.edit.compare`, matches the slides' corresponding
  shapes -- placeholders by type and idx, then a unique name, then kind and nearest box --
  and lists per shape where the box, text size per level, text colour and typeface, fill
  or outline differ from the majority: near-misses (2 pt; a colour within 8 of 255 per
  channel) with a one-call `fix`, larger differences apart, where each value comes from
  (`slide`, `layout`, `master`) when the sources differ, and shapes missing on some slides
  or found on few. Measurement only. `ppt_read_slides` compact gains a `size` column, the
  effective size most of a shape's text is drawn at. The PROMPT names the tool after a run
  of like slides is built.
- The loaded core's budget is now 6,500 tokens as count-tokens counts them (raised from 6,000
  by the user on 2026-10-17); the offline proxy and the online check use it.
- A finished slide onto another layout (round 7, L2): `ppt_manage_slides` action
  `set_layout` with `layout` (`L3` or its name), library `Slide.set_layout`, does what
  PowerPoint's Layout menu does (measured on 53 made-up slides, round 7, A1-A9 and B2):
  empty placeholders go and the layout's are added empty (A6's form, the layout
  placeholder's name as on a new slide; PowerPoint writes "<kind> N" in its interface
  language); filled ones match by kind -- title or centred title; content, body or
  subtitle, the same type first; pic and tbl placeholders take no text; a picture a picture
  placeholder, else a content one -- with both sides sorted by idx and paired in order,
  never by name; a matched one takes the layout's type and idx (no `sz`) and *loses its own
  box*, also on the same layout again, its text, fill, outline and formatting kept; a
  picture moved to a picture placeholder (its own again too) gets a crop that fills it from
  the whole image, centred, replacing its own (an empty `a:srcRect` when the shapes agree);
  one moved into a content placeholder loses its crop and is drawn at its own size -- its
  pixels at the whole dpi its `pHYs` or JPEG density states, else 144 dpi (measured on a
  Retina Mac) -- shrunk to fit keeping its shape, never enlarged, centred, with no frame
  when it fills the box exactly; a table takes a table or content placeholder (either way),
  at the box's place and width, its columns and cells' left and right insets (91440 when
  unstated) scaled alike, its rows kept; a chart takes any content, table or picture
  placeholder and its box; dates, footers and slide numbers take the layout's of their
  kind keeping `sz`, empty ones too, and are never added (an empty one with no place goes);
  content with no place stays, idx 4294967295, its box written on it with x snapped to the
  nearer side (a table or chart its size kept); the placeholders gather into one block where the first was, in the layout's order,
  orphans last; other shapes and every relationship but the layout's untouched. The result
  lists where each placeholder went, the ones added and removed, the orphans and why
  (`no title placeholder left on L8 'Body'`) and the boxes dropped. Measured further on the
  sample and 46 more slides (round 7, A2 of the follow-up). One undo step. Every
  definition +36 estimated tokens (guard 12,725 -> 12,760), the core unchanged; the strict
  set changes on Claude: ppt_set_theme is sent non-strict, edit_smartart strict.
- A layout gets placeholders of its own (round 7, L1): `ppt_edit_master` action
  `add_placeholder`, with `type` (`body`, `content`, `pic`, or `title` on a layout without
  one), `box` (points) and `text` (the name; default `Text Placeholder <n>` and the like),
  writes a placeholder each slide on the layout fills -- boxes in fixed places for a run of
  like slides. Its `idx` is unique on the layout: the largest there plus one, at least 13,
  inferred from Office's bundled layouts (date 10, footer 11, slide number 12, the ones they
  add 13, 14, ...); PowerPoint's own choice could not be measured, and a duplicate idx,
  which PowerPoint accepts silently, would leave slides two placeholders to inherit from.
  The result gives its address (`L12.6`), `idx`, name and box, and the slides on the layout
  that lack it (`slides_without`). As in PowerPoint (measured, round 7, B) slides made on
  the layout afterwards get it empty, without a frame; slides already there do not, unless
  listed in `slides` -- then each gets an empty copy as PowerPoint's re-apply of the layout
  writes one (type and idx, empty `spPr`, one empty paragraph; a picture's without a text
  body) -- or later with the placeholder as `target` and `slides`. `ppt_add_slide` lists an
  added placeholder's `idx` and name; `ppt_add_picture` with an empty picture placeholder as
  `target` fills it: the picture keeps its `p:ph`, id and name, cropped to fill its box
  (drafting's form; PowerPoint's own is unprobed). Move, resize, rename, prompt or delete it
  as any layout shape (`L12.6`); table and chart placeholders are not added. A new slide's
  picture or table placeholder now has no text body and a content placeholder no
  `type="obj"`, as PowerPoint writes them. Library: `Layout.add_placeholder`,
  `Slide.add_layout_placeholder`, `Shape.insert_picture`. Every definition +132 estimated
  tokens (guard 12,595 -> 12,725), the core unchanged; the strict set unchanged.
- Date names for a region (round 7, L5; ooxml-edit 0.44.0, ooxml-common 0.19.0): a table's
  or a chart's `date_format` and a text fact's date pattern with month or day names
  (`d mmmm yyyy`) were refused in a locale such as en-BE or fr-BE; they now take the
  language's names, as Office shows them (measured in Excel for Mac, 2026-10-17: en-BE
  English, fr-BE French, de-CH German), the locale's separators and day order unchanged.
  A chart keeps the region's id where Excel was seen storing one (`[$-80C]` for fr-BE),
  else its language's (`[$-409]` for en-BE), and the render draws those names. A language
  with no measured names (pt-BR) is refused as before; Spanish regions are not the
  language's (measured): es-ES, es-MX (`[$-80A]`) and es have their own, others are refused.
  Tool definitions unchanged.
- A size set on a slide that overrides its layout's is a check fact (round 7, L4): `check`
  and the checks after a change list `overrides`, one entry per placeholder whose slide
  states a text size -- a run's `sz`, its paragraph's `a:defRPr` or the shape's own list
  style -- other than the one its layout gives it (the layout placeholder's list style,
  the master placeholder's or the master's text style, by the cascade `effective_size`
  reads): `{"shape": "261.2", "overrides": "size", "set": 26, "layout": 28}`, `set` a
  list (largest first) where the shape sets several sizes, and `levels: [{"level", "set",
  "layout"}]` where its levels take different layout sizes; with an `overrides_note`.
  Measurement, not a fault: `save_document` does not call it unresolved. Sizes are
  compared unscaled -- a stored `normAutofit fontScale` is PowerPoint's shrink, applied
  to both alike, not a size the slide sets. Not listed: a size equal to the layout's (it
  draws the same; a deck written by tools that restate every size would list every
  shape), a shape that is not a placeholder, text whose size only the presentation's
  default or PowerPoint's 18 pt gives, whitespace-only runs. Each entry carries its `fix`,
  the `ppt_format_text` item `{"target": "261.2", "size": 0}`: `size: 0` (new) removes the size set where it is
  given -- on a run its `sz`; on a paragraph its runs', end mark's and `a:defRPr`'s; on a
  shape or cell every paragraph's and its own list style's levels -- so the layout's (or
  master's, or the table's) applies; sizes from 1 to 400 as before, below 1 refused.
  (`size: null` cannot say it: an optional `null` is read as not given, as the OpenAI
  strict adapters send every unused option.)
  Library: `Slide.size_overrides`, `pptx_agent.edit.effective.size_overrides`,
  `size_overrides_json`, `slide_size`, `template_size`, `Paragraph.clear_size`,
  `TextFrame.clear_size`. `ppt_format_text`'s `size` takes 0, described: every definition
  +15 estimated tokens (12,574 -> 12,589, guard 12,595), the core unchanged (3,553); the
  strict set unchanged.
- An empty `items` list is a no-op with a warning, not a refusal (round 7, L7; ooxml-edit
  0.43.0's `Result.nothing_given`): `ppt_add_shape`, `ppt_set_text`, `ppt_format_text`,
  `ppt_set_shape`, `ppt_add_connector` and `ppt_comments` (other than `list`) given
  `items: []` change nothing, record no undo step and warn `items is empty: nothing was
  added` (written, formatted, changed, measured for `ppt_add_shape` with `measure`) --
  what a model sends after trimming a batch. `ppt_add_shape` still resolves its `slide`
  first, so a wrong one is still named; `ppt_comments` without `items` still needs them.
  The five arrays lose `minItems: 1`: the core -8 estimated tokens (3,561 -> 3,553),
  every definition -18 (12,592 -> 12,574); the guard stays 12,595. `ppt_arrange`'s
  `targets` keeps its `minItems: 1` (an action on no shapes is a mistake, not a leftover).
- `batch` runs a tool the application registered in the toolbox with these (round 7, L6;
  ooxml-edit 0.43.0), tested here with a made-up one; one the application dispatches
  itself is refused in the op saying to call it on its own.

- A chart title that states no font is a check finding (round 6, item 11; ooxml-common
  0.18.0): `check`, the checks after a change and `save_document`'s `unresolved` list
  `chart_title_font`, one entry per chart -- `{"chart": "259.4", "fix": {"action":
  "set_text_style", "target": "259.4", "part": "title", "style": {"font": "+mn-lt",
  "size": 18.62, "bold": false}}}` -- with a `chart_title_font_note`: "chart title states
  no font: PowerPoint draws it in Arial 18 pt". Measured on PowerPoint 16 for Mac,
  2026-10-16: a rich title with no paragraph `a:defRPr` and no face on its runs is drawn
  ArialMT, regular, at the run's size or 18 pt, whatever the theme's font and the
  chart-wide `c:txPr`; pptx2svg now draws it so. The fix is the one `edit_chart` call that
  states it: `+mn-lt`, the theme's body font as a reference, not a name; `bold: false`,
  because the call writes paragraph defaults and PowerPoint draws a title with those
  **bold** unless `b="0"` is stated (measured: `+mn-lt` and 18.62 pt without `b` came out
  Aptos bold 19 pt; with `b="0"` regular); and, where the title states no size, 18.62 pt,
  the size `edit_chart add` writes for a new title (ooxml-edit's `POWERPOINT_LOOK`; its
  titles state `b="0"`). A title that states a size is listed without one. Not listed: a
  title with paragraph defaults (the theme's face already), a new chart's title, an
  automatic title (the minor face, bold: measured), a title naming a face, a deleted one.
  Nothing is changed by itself. Library: `pptx_agent.edit.chart_text.title_font_fix`,
  `unstated_title_fonts`. Tool definitions unchanged.
- `render` takes layouts and masters (round 6, R11; ooxml-edit 0.41.0, pptx2svg's
  `ConvertOptions(templates=...)`): `slides: ["L2", "M0", 1]`, the addresses
  ppt_edit_master and the add tools take, in the order given. A layout is drawn as a new
  slide on it starts -- its background, the master's shapes unless `showMasterSp` hides
  them, its own shapes -- with every placeholder a grey dashed box holding its prompt in
  its text style and tagged with its type and `idx`, and a banner naming it; a master
  likewise with its own placeholders. A wireframe of pptx2svg's, chosen for a designer:
  PowerPoint draws no empty placeholder outside editing, and Slide Master view is not
  measured. The image is labelled `layout L2 'Title and Content'`, the result carries a
  `template_note`, and the cache keys a layout by its key. `describe` already lists the
  layouts with names and placeholders. In the library: `Document.render_svg/render_png`
  take `"L2"`, `"M0"`, a `Layout` or a `Master` among `slides`, and `Layout.render_png()`,
  `Master.render_svg()` draw one. Definitions: render's `slides` names the forms and
  `slides`/`pages` stop repeating the tool's "at most 4": the core +1 estimated token
  (3,560 -> 3,561), every definition unchanged (12,592); the guard stays 12,595. render's
  schema types are unchanged.
- A dot chart (round 6, R10; ooxml-edit 0.40.0): `edit_chart` `add` with `chart_type:
  "dot"` draws one marker per category on the value axis, no line -- averages on a fixed
  1-5 scale with their values is one call (`axis: "value"`, `scale: {"min": 1, "max": 5}`,
  `data_labels: true`). It is a native line chart, so PowerPoint edits it as one: each
  series' line off and a marker on (a 5 pt circle in the series colour, the marker
  PowerPoint gives a new scatter series, measured); `colors` colour the markers.
  `edit_chart read`, `ppt_read_slides`' compact view and the outline report any line chart
  whose series all have no line and a marker as `dot` (`Chart.is_dot`); `set_values` and
  `add_series` keep it one. Categories run along the bottom: the schema has no horizontal
  line chart, and a scatter chart has no category labels. The `chart_type` enum gains
  `dot` (+2 estimated tokens to every definition, 12,590 -> 12,592; the guard stays
  12,595; edit_chart is sent non-strict).
- A chart's dates read as dates (round 6; ooxml-edit 0.39.0): `edit_chart read` lists
  categories kept as serials under a date code as ISO dates (`2026-03-02`, not `46083`)
  with `category_format`, the code that shows them; a category is named by that date
  (`category: "2026-03-02"`), and one that is none is refused listing the dates. The
  outline's chart summary lists them so too. ISO, not the code's text: a model reads it
  whatever the locale, and it is the form `from` dates take. `render` draws them as
  PowerPoint does (pptx2svg on ooxml-common 0.17.0: `5 januari 2026`). No definition
  changes.
- The chart text check sizes an automatic title (a `c:title` with no text, which
  PowerPoint shows as the series' name) as PowerPoint draws it: its own `c:txPr` size,
  else 1.2 times the chart-wide size rounded to the point (8, 10, 11 and 24 pt give 10,
  12, 13 and 29), else 18 pt -- measured on PowerPoint 16 for Mac, 2026-10-16. It was
  sized as a rich title (the master's text styles, 18 pt). A rich title with no size
  stays 18 pt whatever the chart-wide size (measured).

- Date formats (round 6, R9; ooxml-edit 0.38.0, `ooxml_edit.dates`). Excel's date tokens
  (`d dd ddd dddd m mm mmm mmmm mmmmm yy yyyy`, either case), everything else written as
  given; month and day names measured in Excel for Mac, 2026-10-16, for en-US, en-GB,
  nl-NL, nl-BE, de-DE, fr-FR, es-ES and it-IT; another locale refuses a name token and
  takes the numeric ones.
  - A reference in any text tool takes a date pattern: `{{d:events/#1/when:d mmmm yyyy}}`
    is `2 March 2026`, `2 maart 2026` in nl-BE; `:dd/mm/yyyy` is `02/03/2026` in every
    locale (fixed text). A date pattern on a non-date and a number format on a date are
    refused, naming the other. The prompt says so in one clause.
  - `ppt_add_table` and `ppt_edit_table`'s `set_cells` from take `date_format`: the dates
    of the columns placed in that pattern; refused without `from`, or when no column placed
    holds a date.
  - `edit_chart` add from takes `date_format`: the category dates are kept as dates (serials)
    in the pattern's Excel code for the session's locale (`[$-813]d\ mmm\ yyyy`), in the
    chart and in its workbook, on a text axis (one band per date); PowerPoint shows the
    code's `/` as the viewer's date separator and the names by the prefix (measured).
    `set_values` and `add_series` from match a date column to such categories. Not yet
    drawn by pptx2svg: `render` shows the serials (46083), as `edit_chart read` lists them.
  - Without a pattern nothing changes: dates are written in the locale's day order.
  - Definitions: +84 estimated tokens to every definition (12,506 -> 12,590), none to the
    core (3,560); the guard rises to 12,595. The strict set is unchanged.

- Slide numbers by reference (round 6, R8): `{{slide:s:262}}` -- or `{{slide:$results}}`,
  a ref to a slide -- in any text tool's text (`ppt_set_text`, `ppt_add_shape`, a table's
  `data` and `set_cells`, `replace_text`'s `replace`) writes that slide's number, kept
  current for agenda and divider slides. The deck's own reference (ooxml-edit 0.37.0,
  `DocumentFormat.own_facts`), kept as a text fact like a dataset value's
  (`UVRT_TEXT_FACTS`, `source` `slide:s:262`, `slide` 262), listed in `data.facts`.
  - The number is the one the slide's slide-number field shows: its position plus
    `p:presentation@firstSlideNum` minus one (as pptx2svg measured the field), hidden
    slides counted as the field counts them.
  - Kept current: at the end of every outermost `Document.batch` -- every tool call, and
    every library call that adds, deletes, moves or duplicates slides -- the deck compares
    its slide order with the one it last numbered for and, when it changed, rewrites the
    numbers that moved, in the same undo step (undo takes both back). The text is
    rewritten as a text tool writes it, so the run keeps its formatting, and the other
    facts in the text move with it. The checks list each (`slide_numbers`: `shape`,
    `slide`, `text`, `was`). `Document.update_slide_numbers()` rewrites them all;
    `Document.save`, `save_as_template` and `save_document` call it first (its changes in
    `save_document`'s `slide_numbers`), so a deck PowerPoint reordered is caught on save.
  - A reference whose slide was deleted keeps its last number and is marked `gone`, so a
    slide later given the same id (a reopened deck hands out a deleted top id again) is
    never taken for it; `check` and `save_document` list it as `missing_slides` ("slide
    s:262 no longer exists: the text keeps its last number, 3"), counted in `check`'s
    summary.
  - Refused, listing the slides: `{{slide:3}}` (says which slide is number 3), a shape
    (`{{slide:257.2}}`), a slide that is not there, a document handle or a number format.
  - Not kept current: speaker notes and SmartArt nodes (they keep no facts), shapes copied
    to another deck (their slide facts are dropped; the text stays), and an edited number
    (its fact is dropped as any edited fact is). PowerPoint has no field that shows another
    slide's number, so in a deck edited in PowerPoint the number is plain text and stays
    stale until this library next changes its slide order or saves it.
  - The prompt fragment names the syntax in one sentence; no definition changes (every
    definition 12,506, core 3,560 estimated).

- A minimum text size per kind, tables grouped, chart text checked (round 6, R1). 380 of
  402 `small_text` entries in a 46-slide deck were table cells the brief allowed smaller.
  - `deck_format(min_text_size=...)` also takes a mapping, `{"body": 12, "table": 10,
    "chart": 10}`: `body` the text of shapes, text boxes and placeholders, `table` table
    cells', `chart` a chart's title, axis titles, axis labels, data labels and legend. A
    kind the mapping leaves out (or gives `None`) is not checked; an unknown kind, an
    empty mapping or a size outside 1 to 400 is refused when the format is made, naming
    the kinds. One number is every kind's minimum, as before -- charts included now.
  - `small_text` lists a table once: `{"table": "257.4", "cells": 30, "size": 9.0,
    "min": 10, "from": "the run"}` -- how many cells are below, the smallest size and where
    it comes from (the smallest, since it is what has to rise). A chart once, each part's
    smallest size, the part as `edit_chart`'s `set_text_style` takes it: `{"chart":
    "257.5", "min": 10, "parts": [{"part": "axis_labels", "size": 8.0, "from": "the
    chart's text properties"}]}`. A shape's entries stay one per size and source. The
    note names each kind's minimum, `(body 12 pt, table 10 pt, chart 10 pt)`, and says how
    a table's and a chart's entry read; `check`, the checks after every change and
    `save_document`'s `unresolved` all list it so.
  - Chart text sizes from the chart part (`pptx_agent.edit.chart_text`), the cascade
    pptx2svg draws with (`ooxml_common.chart.layout`): axis labels, the legend and data
    labels take the innermost `c:txPr` (a point's label, its series', the chart group's;
    the axis', the legend's), then the chart space's, then 10 pt; a title or axis title
    is DrawingML text, resolved as a text box's (run, paragraph, its `c:txPr`, its list
    style, the master's `otherStyle`, the presentation's default, 18 pt). Only what
    PowerPoint shows is checked: a stated title, axes not deleted and with tick labels,
    data labels that show something, a legend. A new PowerPoint chart's labels are
    11.97 pt, so a 12 pt chart minimum lists every new chart.
  - `Slide.text_below(points)` takes the mapping too and lists chart text;
    `SmallText` gains `kind`, `shape` (a cell's table) and `part` (a chart's).
    `effective.text_minimums`, `effective.small_text_json`.
  - Definitions unchanged: it is the application's setting.

- Slides by id wherever slides are named (round 6, R6). `ppt_read_slides slides:[278]`
  said "no slide 278": slide lists took positions only, render and check ids too.
  - Every `slides` list takes a number from 1, `s:<id>` or a `$ref`: `ppt_read_slides`,
    `ppt_edit_master` (set_footer), `ppt_comments` (list), and from ooxml-edit 0.36.0
    `find_text` and `replace_text`, as `render` and `check` did. `ppt_comments`' list
    read a bare number as an id before; it is a position now, as in every list.
  - A number past the last slide that is a slide's id is refused with the way to say it,
    never taken for that slide: "no slide 278 (the deck has 12); slide id 278 is slide 4:
    write 's:278' or 4". `s:4` when 4 is no id but a position: "slide 4 is s:259: write
    's:259' or 4". A `slide` argument (where a bare number is an id, `s:256`'s short
    form) given a position: "no slide id 2 ...; slide 2 is s:257: write 's:257'". `at` and
    `to` (positions to put a slide at) given an id: "position 278; from 1 to 13; slide id
    278 is slide 4: ...". `tools.common.id_hint`.
  - Definitions: `ppt_read_slides`' slides are strings with `orInteger` (as render's),
    the core +18 estimated (3,542 -> 3,560), every definition +29 (12,477 -> 12,506); the
    guard rises to 12,510.
- Slide ids near the top of the range (round 6, R12; measured with PowerPoint for Mac on
  2026-10-16): once 2,147,483,647 is taken, new slide ids continue from 256, the lowest
  free, exactly as PowerPoint does -- a deck whose slide is 2147483641 gets 2147483642 to
  2147483647, then 256, 257, ... -- and no existing slide is renumbered. The order was
  already this; what changed is that an id past the top is also never one handed out or
  deleted earlier in the session (it used to be the lowest not in use, so a deleted 256
  came back): the in-session rule holds there too. PowerPoint hands a deleted id out
  again; this does not while the deck is open.

- Forgiving the model's small mistakes (round 6, R2, R3, R4, R7), each read with a line in
  the result's `warnings` and no change to any definition:
  - Text and paragraphs together (`ppt_set_text`, `ppt_add_shape`, measuring too): when
    the paragraphs' text -- runs joined, paragraphs by `\n`, as they are written -- is
    `text`, the paragraphs are used ("items[0]: text and paragraphs say the same;
    paragraphs used -- give one of them"); when it is not, the refusal says where they part
    ("paragraph 1: text has 'Plan A', paragraphs[0] 'Plan B'", or the paragraph counts).
    `tools.common.one_text`, `paragraphs_text`.
  - `ppt_format_text` reads an item's `top` as `space_before` and `bottom` as
    `space_after` ("items[11].top read as space_before"): an item has no other top or
    bottom (the frame's are `insets.top` and `anchor`).
  - From ooxml-edit 0.35.0, tested here: a ref named `recNow` or `scale-note` is kept as
    `rec_now` / `scale_note`, `$recNow` read as `$rec_now` (refused when the folded name
    already names another shape); `functions.ppt_set_text` as an op's `tool` in a batch;
    a key one level off says where it belongs (`items[0].box.preset_name`: "preset_name
    belongs one level up, on items[0], not in box").

- SVG recolouring as PowerPoint writes it (round 5, T20b). A second round of measurements
  on PowerPoint for Mac (2026-10-15) covered what T20 left out; `Shape.recolor` now writes
  what PowerPoint writes in each case:
  - Every theme slot's class: `MsftOfcThm_<slot>_Fill_v2` for `MainDark1`, `MainLight1`,
    `MainDark2`, `MainLight2`, `Accent1`-`Accent6`, `Hyperlink`, `FollowedHyperlink`,
    `Text1`, `Background1`, `Text2` and `Background2` (`dk1`... are the Main slots,
    `tx1`/`text1`... the Text/Background ones), so every theme colour keeps its link; a
    tint or shade is still its value alone (not measurable by script).
    `svgpicture.THEME_SLOTS` and `theme_class()` replace `THEME_CLASSES`.
  - Shapes with no fill are filled too (outline icons included; the warning and the
    refusal of an outline-only SVG are gone, and `RecolorWarning` with them), and every
    shape gets PowerPoint's stroke attributes, its own stroke kept.
  - `recolor(line=, line_width=)`: PowerPoint's Graphics Outline -- a `_Stroke_v2` class
    (every slot measured) and rule and `stroke=` on every shape that has a stroke, fills
    left alone; a colour adds no stroke to a shape without one. A weight is converted to
    the SVG's units at the frame's size and strokes the shapes that had none in Text 1,
    PowerPoint's default outline (`MsftOfcThm_Text1_Stroke_v2`); with a colour too, in
    that colour (the two measured rules composed). A transparency, `"#FF6600 alpha=50%"`,
    is `stroke-opacity`, drawn once in the PNG (PowerPoint's looked doubled). With a fill,
    one style block and both classes.
    `ppt_set_shape`'s `line` `color` and `width` on an SVG picture do it (`dash` or
    arrowheads with them are refused; alone, or `color: none`, they are the picture's
    border as before); on a bitmap `line` stays its border.
  - An SVG's style sheet and `style` attributes are replaced, as PowerPoint replaces them,
    every declaration carried over as a presentation attribute of the element it applied
    to (so strokes draw the same); a selector beyond tags, classes, ids and descendant or
    child combinators, an at-rule, or a property with no attribute is refused. This
    replaces T20's fill in each shape's `style`.
  - A transparency, `"accent2 alpha=50%"`, is `fill-opacity="0.5"` (refused before); the
    PNG fallback draws it once (PowerPoint's own fallback looked to apply it twice for an
    RGB colour).
  - A picture sharing its SVG with others on the slide is split as PowerPoint splits it:
    the recoloured one keeps its relationships and parts, the others get copies (T20 gave
    the recoloured one new parts); shared with another slide, it still gets its own.
  - Gradient fills become the solid fill (the gradient stays, unreferenced); `use`
    elements and a `symbol`'s shapes are filled.
  - Tools: `ppt_add_picture`'s description says `fill and line` recolour an SVG (+3
    estimated tokens for every definition, 12,474 -> 12,477, within the 12,480 guard); the
    core and the strict set are unchanged.

- Brand assets and recolouring (round 5, T20; ooxml-edit 0.34.0). Production shipped two
  tools of its own over a brand icon set (about 2,800 SVGs with a keyword index), and
  `Picture.replace_image(svg=)` had no tool, so a highlighted icon was inserted again.
  - `Shape.recolor(color)`: an SVG picture's shapes filled with one colour as PowerPoint's
    Graphics Fill does it, measured on PowerPoint for Mac (2026-10-15): the SVG rewritten
    in place (`fill` and `fill-opacity="1"` on every shape, no-fill and `currentColor` ones
    too; for accent 2 also PowerPoint's class and rule, `MsftOfcThm_Accent2_Fill_v2`, which
    keeps the theme link), the PNG fallback drawn again, the `p:pic` kept. A theme colour
    is resolved against the slide's theme and colour map; other slots' classes are not
    measured, so their value alone is written. Outline shapes (fill `none`) keep their
    colours (`svgpicture.RecolorWarning`); a bitmap, an outline-only SVG and a colour with
    alpha are refused. Parts another picture shares are not touched: this one gets its own.
  - Tools: `ppt_set_shape`'s `fill` on an SVG picture recolours it -- PowerPoint's own name
    for a graphic's recolour; on a bitmap picture it stays the picture's background fill,
    as before (PowerPoint's Fill of a raster picture; only a graphic has Graphics Fill).
    `none` and `inherit` work as before. No new parameter: the core is unchanged.
    `ppt_add_picture`'s description says so (+9 estimated tokens for every definition,
    12,465 -> 12,474; the guard rises from 12,470 to 12,480).
  - `deck_format(assets=catalogue)`: an application's asset catalogue (`FolderCatalogue`
    over a folder and a JSON or CSV index, or any object with `search` and `get`). With
    one the toolbox offers `ppt_find_asset` (deferred, `ppt_objects`; ids, names and
    keywords for words) and `ppt_add_picture` takes `asset` in place of `image`, to insert
    or replace; without one neither is sent (with one: +141 estimated tokens, all
    deferred; the core unchanged).

- Ids are never handed out twice while a deck is open (round 5, T9). Deleting shape
  `256.3` and adding one gave the new shape `256.3`; deleting slide 257 and adding one gave
  the new slide 257 -- so a reviewer's finding or an agent's note about the old one silently
  pointed at the new one.
  - Shapes: PowerPoint's rule, measured with PowerPoint for Mac on 2026-10-15 -- the lowest
    `cNvPr@id` from 2 that the slide has not had since the deck was opened (every id on it
    then, every id handed out or deleted since). Of 2..6, delete 6 and add: 7, not 6; of 2,
    3 and 7, add four: 4, 5, 6 and 8. Layouts and masters likewise.
  - Slides: one above the highest `sldId` seen since the deck was opened, at most
    2,147,483,647 (ECMA-376 `ST_SlideId`), then the lowest free from 256 (round 6, R12,
    above). This differs on purpose from PowerPoint, which reuses a deleted top `sldId` in
    the session (measured the same day).
  - The memory is the `Document`'s, outside what undo restores: an id undone, or added in a
    batch or tool-call savepoint that rolled back, stays taken; a measuring copy
    (`copy_of`) carries it on. Nothing is written to the file: after a save and reopen only
    the ids in it count, as in PowerPoint. The definitions and prompt are unchanged.

- Figures in text keep their source (round 5, T4; ooxml-edit 0.33.0). Tables and charts
  kept where their values came from, text did not: a card saying "Average 2.2 across 41
  items" had no record of the 2.2, and builders expanded their own placeholders first.
  - `ppt_set_text` (text and runs), `ppt_add_shape`, `ppt_add_table`'s `data`,
    `ppt_edit_table`'s `set_cells` and `replace_text`'s `replace` take references,
    `{{d:by_topic/door/mean}}` or `{{t:checks/N-1/rating}}`, `:0.0` after one to round:
    the value is written in the session's locale (`2,2` in nl-BE) and the result's
    `data.facts` lists each with its sources. `\{{` writes one literally; `{{name}}`
    stays as it is; a reference to no single value refuses the call, changing nothing.
  - The deck keeps each value as a fact -- the record's fields (`source`, `format`,
    `text`, `document`, `version`, `sources`) and its place, `p` and `at` (a table's
    `cell` too) -- in a `UVRT_TEXT_FACTS` tag of the shape's tags part
    (`p:nvSpPr/p:nvPr/p:custDataLst/p:tags`; a table's frame, beside its own record).
    `Document.provenance()` and `session.provenance(doc)` list them (`facts`) while their
    characters are their text: an edit keeps the facts it does not touch, moved, and drops
    the ones it does; a table's row and column inserts and deletes move or drop its cells'
    facts; and reading skips a fact whose text is no longer where it says, whoever edited
    it. Undo takes them with the text; copies take their own. Speaker notes and SmartArt
    take the values but keep no facts.
  - Not measured: that PowerPoint keeps a tag on a `p:sp` (it was measured on a graphic
    frame, 2026-10-14). The prompt: one sentence on references in the deck fragment (197
    characters, about 57 estimated tokens). The definitions are unchanged (core 3,542,
    every definition 12,465; guards as they were); the strict set is unchanged.

- A chart added with its look in one call (round 5, T2; ooxml-edit 0.32.0). Styling one
  new chart took a builder six turns. `edit_chart`'s `add`, which took the frame (`box`),
  scale, gridlines, legend position and number format already, also takes `data_labels:
  true`, `colors` (each series' colour in order, a pie's slices'; filled, lined or as
  markers as each kind is drawn) and `style` with `part` (the text: font, size, bold,
  colour). The chart is byte for byte the one `add` and the separate `show_data_labels`,
  `format_series` (`format_point` for a pie) and `set_text_style` calls write, tested
  for column with a line series, line, scatter, pie and filled radar; one undo step, and a
  refused look adds no chart. The definitions: 62 estimated tokens in every definition
  (12,403 -> 12,465; the guard 12,405 -> 12,470), none in the core (3,542). The strict set
  is unchanged.

- Dense slides read in brief (round 5, T10). Reviewers read slides of up to 118 shapes, and
  `ppt_read_slides` with `geometry` gave every text in full.
  - `detail: "compact"`: one row per shape under a header, `columns` = id, kind, box,
    fill, text, group; back to front, as geometry orders them. The kind is geometry's,
    with a graphic frame named `table`, `chart` or `smartart`; the box `[x, y, w, h]` in
    whole points; the fill as drawn (resolved hex, else `gradient`, `image`, `pattern`);
    the text's first eight words, then `… (23 words)`. A table gives `3×4: ` and its first
    cell's words, a chart `bar, 2 series`, a group `3 shapes`; a group's members name
    it in the last column, so tile grids read as tiles. Pages of 300 rows.
  - Measured on a made-up 111-shape slide (24 three-shape tiles, 8 connectors, a table, a
    chart, 4 long text boxes): geometry 29,647 characters over two pages, compact 8,480 in
    one (29%). The header-and-rows encoding is 23% smaller than the same records as
    objects (8,282 against 10,724 characters), so rows it is.
  - The definitions: `compact` in the `detail` enum and nine words of description (the
    svg view's "compact" dropped so the word means one thing) add 16 estimated tokens to
    the core (3,526 -> 3,542, about 5,242 counted by the proxy) and to every definition
    (12,387 -> 12,403; the guard 12,390 -> 12,405). The strict set is unchanged.

- Tile and card grids in one call (round 5, T5; the user brought grid back as a
  `ppt_arrange` action, TOOLS-ROADMAP D13 amended). Builders drew grids by hand -- one
  of 41 tiles -- computing every box, the main source of their layout mistakes.
  - `ppt_arrange` `action: "grid"`: the targets, in the order given, row by row in
    `columns` cells across (rows follow, or `rows`), `gap` points apart (`gap_y` between
    rows, if different), filling `box` or, without one, the slide's content area
    (`Slide.content_area`, as `align` and `distribute` take it). Each shape is sized to
    its cell; a picture keeps its aspect, as large as fits, centred in its cell. The
    result lists each target's new box and the area.
  - `action: "flow"`: the targets keep their size and wrap left to right in cells as large
    as the largest, at most `columns` across.
  - With `$ref`s from `ppt_add_shape` in the same `batch`, a 40-tile grid is one call and
    one undo step. Refused: `columns` under 1; gaps that leave no room for the cells,
    more targets than `rows` x `columns`, or flowed rows past the area's foot, each saying
    what to change; targets on different slides; grid parameters on other actions.
  - Library: `edit.layout.grid` takes `fit="fill"` (shapes to the cell, pictures by
    aspect) and `edit.layout.flow` is new; refusals say what to change.
  - The definitions: `gap_y`, `columns`, `rows` and `box` on the deferred, non-strict
    `ppt_arrange` add 162 estimated tokens to every definition (12,225 -> 12,387; the
    guard 12,230 -> 12,390), none to the core (3,526). The strict set is unchanged.
- A long table continued on the next slides, and fewer rows of a source (round 5, T3;
  ooxml-edit 0.31.0, xlsx-agent's row ranges). A builder wanted rows 1-4 of a 42-row
  table on one slide: `rows: 5` was refused ("from needs 42") and a row range was not an
  address.
  - `from` takes a row range: `t:items/#1:#9` (rows 1 to 9, from 1, the last included;
    the workbook's grammar) or `d:by_region/#10:` (row 10 to the last; ooxml-edit), each
    value with its sources.
  - `rows` fewer than the source is refused saying how to place fewer: "rows is 10, but
    from has 40 rows and a header: place fewer with d2/t:items/#1:#9, continue: true to
    go on over the next slides, or leave rows out (nothing is cut)", the row range as
    `valid_options`.
  - `ppt_add_table` with `continue: true` places every row: the pages measured as
    `measure: true` measures them, the first on the target slide, each next one on a new
    slide right after it, on the target's layout, its title the target's title text as it
    is (no "(cont.)": the model may retitle), its other empty placeholders deleted; the
    same box, the header row repeated. Each table's rows are set to their measured
    heights (a box's height shared evenly among uneven rows would draw a page taller than
    measured). One undo step. The result lists `pages` -- slide, table, source rows,
    height and `from` -- and `created` the new slides and tables; each table keeps the
    record of its own rows. Refused with `rows` (every row is placed), on a layout or
    master, and with no width.
  - A table record's `first_row` is the table row of the first source row placed, the
    n-th placed at `first_row + n - 1`: source row n when every row is placed, as before.
  - The definitions: `continue` on the deferred, non-strict `ppt_add_table` adds 29
    estimated tokens to every definition (12,196 -> 12,225; the guard 12,200 -> 12,230),
    none to the core (3,526). The strict set is unchanged.
- Measuring a table before placing it (round 5, T13): an agent no longer places a table
  to learn it needs four slides.
  - `ppt_add_table` with `measure: true` adds nothing -- no version, no undo step -- and
    returns each row's height as PowerPoint draws it (`heights`, the header first), the
    table's `height`, whether it `fits`, and its `pages`: the rows below the header, from 1
    (with `from`, the source rows), that fit each page with the header row repeated, and
    each page's table height (`[{"rows": "1-9", "h": 309.6}, {"rows": "10-18", ...}]`). A
    page is the box's `h`, or the slide's content area when `h` is 0; a row too tall for a
    page gets one to itself and is listed in `too_tall`. It takes everything an add takes:
    `from`, `from_columns`, `header`, literal `data`, `style`.
  - `ppt_add_table` takes `size` (every cell's text size, points, empty lines too) and
    `widths` (column widths, points), adding and measuring alike.
  - `Slide.measure_table(data, width=, height=, style=, size=, widths=, header=)` is the
    library call, a `TableMeasure` (`heights`, `pages` as ranges of table rows,
    `page_heights`, `fits`, `too_tall`); `pptx_agent.edit.fit.paginate` splits row heights
    into pages; `Table.set_text_size(points)` sizes a table's text. The table is built on a
    copy of the deck by the same code an add uses (`edit.fit.build_table`) and measured as
    `Table.drawn_row_heights` measures a placed table, so a table placed with a box `h` of 0
    is drawn at exactly the measured heights.
  - The definitions: `measure`, `size` and `widths` on the deferred, non-strict
    `ppt_add_table` add 77 estimated tokens to every definition (12,119 -> 12,196; the guard
    12,120 -> 12,200), none to the core (3,526). The strict set is unchanged.
- Text colour as drawn and inherited text sizes against a minimum (round 5, T12), none of
  it in the definitions: the tool definitions, their estimated tokens and the strict set
  are unchanged.
  - `Run.effective_color`: the run's colour resolved to RGB through the chain
    `effective_size` reads -- run, paragraph, the shape's list style, the layout and
    master placeholders, the master's `titleStyle`/`bodyStyle`/`otherStyle`, the
    presentation's default -- and, for text outside a placeholder, a shape style's
    `fontRef` colour and a table cell's table-style text colour (by region: header and
    total rows, first and last columns, bands, corners), each where pptx2svg measured it
    to sit (ECMA-376 Part 1, 21.1.2.3, 20.1.4.1.17, 20.1.4.2). A theme colour goes
    through the slide's colour map (`clrMap`, `clrMapOvr`) and theme with its modifiers
    applied, as `Color.resolve` does; a hyperlink with no colour of its own is the
    theme's `hlink`; `None` for text in no solid colour. `Paragraph.effective.color` is
    the first run's.
  - PowerPoint's 74 built-in table styles' text colours are in
    `pptx_agent.edit.table_text_colors`, compiled from pptx2svg's measured catalogue by
    `tools/derive_table_text_colors.py` (PowerPoint does not write a built-in style into
    `tableStyles.xml`): a default table's header reads white, its body `dk1`. A deck's
    own style in `tableStyles.xml` comes first.
  - `Run.effective_source("size" | "color")` says where the value comes from: `"the
    run"`, `"layout 'Title and Content' content placeholder"`, `"the master's
    bodyStyle"`, `"table style 'Medium Style 2 - Accent 1' firstRow"`, `"... at 90%
    autofit"`. `Slide.text_below(points)` lists the shapes and table cells whose text is
    drawn smaller, with size and source.
  - The application's minimum text size, `deck_format(min_text_size=12)` in place of
    `FORMAT` (`Toolbox(TOOLS, formats=[deck_format(min_text_size=12)], ...)`): `check`,
    the checks after every change and `save_document` report `small_text` --
    `{"shape": "256.3", "size": 9.0, "from": "layout 'Title and Content' content
    placeholder"}` -- so a placeholder that inherits 9 pt from its layout is reported, not
    only text that states its size. App configuration, never a tool argument; without
    it, nothing is reported.

- `describe` lists the session's datasets (round 5, T11; ooxml-edit 0.30.0), none of it in
  the definitions: a deck's `describe` result gets `datasets` -- each one's id, columns,
  `types`, `rows` (a count) and `from` (`d1/t:Sales`, `d:by_rep`, or `application` for one
  the application kept), `stale`/`closed` when true -- so a figure placed from `d:<id>` is
  seen to have a source. Never a value; no key when the session has none.
- Dates from a source written one way (round 5, T18b; ooxml-edit 0.29.0), none of it in the
  definitions. `ppt_add_table` and `ppt_edit_table` `set_cells` from a source write a `date`, `datetime` or `time` column's ISO
  8601 text -- a workbook's date column through `t:`, a derive's dates, a dataset labelled
  `date` -- in the session locale's day order, as a Python date already was (`1/5/2026`
  en-US, `5/1/2026` nl-BE), so one deck no longer mixes `1/3/2026` and `2026-01-20`;
  unlabelled ISO text is written as it is. `edit_chart` categories from such a column are
  written the same way, and `set_values` matches its rows to categories written ISO (an
  older deck) or in the locale. The kept provenance names cells, never values:
  unchanged. Needs ooxml-edit 0.29.0.

- Fewer wasted turns on names a model naturally writes (round 5, T6-T8; ooxml-edit
  0.27.0), none of it in the definitions: the estimated tokens are unchanged (core 3,526,
  every definition 12,119 by the guard's count) and so is the strict set.
  - Colours: PowerPoint's names for the theme colours are read as the scheme names --
    `text1`, `text2`, `background1`, `background2` as `tx1`, `tx2`, `bg1`, `bg2`, and
    Word's `dark1`...`light2`, `hyperlink`, `followedHyperlink` -- in any case and with a
    space before the number (`"Text 1"`, `"Accent 1 lumMod=75%"`), wherever a colour is
    taken (`Color.parse`, `Color.theme`, so every tool) and in charts. A colour that is
    none still lists the scheme names, and says Office's are read as them.
  - `autofit: "shrink"` is read as `normal` and `"resize"` as `shape`, with a warning
    (`items[0].frame.autofit 'shrink' read as 'normal'`): `normal` is `a:normAutofit`,
    whose `fontScale` ECMA-376 Part 1, 21.1.2.1.3, describes as decreasing the font size
    "in order to auto-fit text within a bounding box" (PowerPoint's "Shrink text on
    overflow"); `shape` is `a:spAutoFit`, 21.1.2.1.4 ("Resize shape to fit text"). An
    autofit close to a mode says "did you mean", as every enum now does.
  - `ppt_add_table` and `ppt_edit_table` read `columns` given as names (a list of
    strings that are not numbers) as `from_columns`, as `edit_chart` names them;
    `edit_chart` reads `from_columns` as `columns` and, adding, `slide` as `target`.
  - `text` with `paragraphs: []` (or `text: ""` with paragraphs) is no longer "both": the
    empty one is left out (`common.given_text`).
  - Refusals say what would do: an unknown argument close to a real one says "did you
    mean"; `key` on `ppt_set_text` and the other setters says keys are for the tools that
    create something and the setter is safe to repeat; a ref name with capitals or `-`
    says the rule in words and a name that would do (`catChart`: `cat_chart`).

- `batch` applies its ops one by one by default (ooxml-edit 0.26.0, round 5 T1): a failing
  op -- a wrong shape id, an enum the schema refuses, a `$ref` that names nothing -- is
  skipped with its own error and the others still apply as one undo step; `data` lists
  each op's outcome with `applied` and `failed` counts, and the summary names the ops to
  resend. A `$ref` a failed op would have made says which op failed. `atomic: true` keeps
  the all-or-none batch. The deck's `restored` hook (`_reset_caches`) now also runs after a
  failed op is rolled back inside the batch. The definitions grow by 22 estimated tokens
  (core 3,504 -> 3,526, about 5,218 counted by the proxy; every definition 12,096 ->
  12,119, guard 12,100 -> 12,120); the strict set is unchanged.
- Source addresses in the deck: what a table's or chart's `from` says is now kept in the
  deck, not only in the result (ooxml-edit 0.25.0's `ooxml_edit.provenance` record: the
  source without the session's document handle, the source document's name, its version,
  and per column -- or category and series -- the rows placed and their sources). It goes
  in the graphic frame's tags part, the place PowerPoint for Mac kept through open, edit
  and save (measured 2026-10-14): `p:nvGraphicFramePr/p:nvPr/p:custDataLst/p:tags r:id` ->
  a slide relationship of type `.../relationships/tags` -> `ppt/tags/tagN.xml`
  (`application/vnd.openxmlformats-officedocument.presentationml.tags+xml`) holding
  `<p:tagLst><p:tag name="UVRT_PROVENANCE" val="{json}"/></p:tagLst>`; the name is upper
  case, as PowerPoint's object model writes tag names. `p:custDataLst` goes in its schema
  place (before `p:extLst`); other tags on the frame stay, and a tags part shared with
  another frame is copied before it is written. Written by `ppt_add_table` and
  `ppt_edit_table`'s `set_cells` with `from` (`call.keep_provenance`) and by `edit_chart`
  with `from` (the format's `store_provenance`; the chart part keeps its own copy), in the
  call's undo step; read by `Session.provenance(doc)` and the new `Document.provenance()`
  -> `{frame address: record}` (`pptx_agent.edit.provenance`; a record that does not read
  is skipped with a `ProvenanceWarning`; a chart frame without the tag falls back on its
  chart part's record). A record is never left untrue: a cell of a table's block written
  without a source (`set_cells`, `ppt_set_text`, `replace_text`, `merge`) drops that
  column's entry, the last one the record; a row inserted or deleted above the block moves
  it, one inside it (its header too) drops the record, one below changes nothing; a column
  inserted or deleted moves the entries and drops the deleted column's. `edit_chart`'s own
  drops (ooxml-edit) drop the tag too, and with the last tag the part and its relationship
  go. Copies are their own: `Shape.duplicate` (`ppt_arrange` duplicate) now copies a
  frame's per-shape parts -- tags, and also a chart (with its workbook) and a diagram,
  which a duplicate used to share with the original -- as `copy_shapes` and
  `duplicate_slide` already did; deleting a shape or a slide leaves no tags part behind.
  No tool definition changes (12,096 estimated tokens, as before); `strict_problems` stays empty; the
  goldens are unchanged. ooxml-edit 0.25.0 is the floor.
- Tables from data (X1): `ppt_add_table` and `ppt_edit_table`'s `set_cells` take `from` --
  a dataset `d:<id>` or another open document's table `t:<name>`, as ooxml-edit 0.22.0's
  `edit_chart` does -- with `from_columns` (the source's columns by name, in order; default
  all) and `header` (the header row's texts; `[]` for none; default the column names). Every
  source row is placed; numbers are written as chart category labels are (General, the
  session locale's decimal separator), dates as the source gives them. `ppt_add_table`'s
  `rows` and `columns` are optional with `from` (the table fits it); `set_cells` writes
  from row `index`. A source too big for a table (200 rows with the header, 50 columns,
  5,000 cells) or for the table written into is refused, never cut. The result's
  `data.from` gives the source, its document and version, the header's and first row's
  table rows, and per placed column its table column, source rows and their sources;
  nothing of it is saved in the deck. Every definition grows by 196 estimated tokens
  (11,897 -> 12,093; the guard 11,900 -> 12,100), the core not at all; ppt_add_table, now
  with nine optional parameters, is sent non-strict and ppt_set_theme strict in its place.
  ooxml-edit 0.22.0 is the floor.

- The user's locale (ooxml-common 0.15.0's `Locale`, ooxml-edit 0.19.0's
  `Toolbox(locale=...)`): `Document.locale` and `Document.new(locale=...)`, set from the
  session for every deck that joins it. Every render passes it to pptx2svg
  (`ConvertOptions(locale=...)`) -- `Document.render_svg`/`render_png`, the agent view, the
  `render` tool's worker process (a subprocess worker too; the locale keys its cache) and
  the text measurement behind `check`. A new chart's `c:lang` and text, a chart title's
  runs, a new deck's language and new text in a deck that states no language take the
  locale's language where they took en-US. en-US without one, never the server's: the
  output is unchanged. No tool definition changes.
- A new date field's cached text (Header & Footer's `date="auto"`) is what PowerPoint draws
  for the field's format and language -- pptx2svg's measured `DATE_FORMATS` -- where it was
  a guess (`M/D/YYYY` for en-US, `D/M/YYYY` for every other language): en-US `datetime1`
  is `M/d/yy` (`1/5/27`), nl-NL `05-01-2027`; a language without a table takes the
  locale's, else en-US's, as pptx2svg draws it. Only the cache changes; PowerPoint redraws
  the field.
- A date scale's tick labels (`DateScale.ticks`) name months and weekdays themselves
  instead of through `strftime`, whose names followed the process's `LC_TIME`: English, as
  before under the C locale, or with `language=` the names pptx2svg measured from
  PowerPoint's date fields (nl, de, fr). PowerPoint's own date axes follow its interface
  language, which no file or locale here states, so nothing picks a language for them.
- ooxml-edit 0.18.0 makes workbooks a shared-tool kind, and the shared definitions every
  deck request carries grow with it (`describe`, `find_text`, `new_document`'s `kind`,
  `save_document`'s `format`): the core 3,440 -> 3,504 estimated (about 5,186 counted by
  the proxy, within 6,000), every definition 11,844; the guard rises to 11,900. A deck
  still saves as pptx, potx or outline: a workbook format is refused with those three.
- A placeholder holding only whitespace now renders with the stray bullet PowerPoint draws
  for it (ooxml-common 0.13.0), so the render shows what `empty_placeholders` reports;
  one of empty paragraphs still renders nothing.
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
