# Using pptx-agent

pptx-agent edits a PowerPoint deck in place, through a semantic API, and renders it back to
a picture. A typical agent loop is:

1. **look**: `deck.to_outline()` (Markdown, cheap) or `slide.render_png()` (pixels);
2. **find**: shape ids from the outline, raw text from `deck.outline_blocks()` or
   `deck.find_text()`, table cells and chart points by their labels;
3. **change**: `set_text`, `cell.text = ...`, `series.set_value(...)`, fills, shapes, slides;
4. **check**: `deck.validate()`, `deck.overflows()` (does the text fit, as PowerPoint will
   show it? does anything collide?), render again, read the outline again;
5. **save**: `deck.save("out.pptx")`.

Everything you do not touch survives byte for byte. Lengths are EMU, font sizes points:
see [Units](#units-emu-everywhere-pt-for-points) below.

## What you can edit

| | |
| --- | --- |
| Text | replace keeping mixed formatting, paragraphs, runs, bullets that hang, numbering, spacing, hyperlinks; sizes and fonts as drawn; autofit, insets, vertical anchor, wrapping |
| Speaker notes | read, write (keeping formatting), find, address as `<sldId>/notes`; a duplicate with new notes |
| Tables | **fully editable**: cell text and formatting, fills, borders, rows and columns, sizes, merges; cells by row and column label |
| Charts | **fully editable**: values, categories, series, titles, axis titles, legend; series by name and points by category label; the cached data and the embedded workbook (what Edit Data opens) always change together |
| SmartArt | node text, adding and removing nodes |
| Shapes | position, size, rotation, flips, preset geometry and adjustments, fills and outlines in theme colours, z-order, duplicate, delete, groups; new autoshapes, text boxes, tables, pictures and connectors that stay attached; a picture's native size, its image replaced |
| Slides | add from a layout, duplicate, reorder, delete; titles, and a slide by its title; a layout's placeholders; draft whole slides from Markdown |
| Decks | new from nothing or from a template; title, author, language, slide size; the theme's colours and fonts, read and set; which colours do which job, and their tints |
| Checks | `validate()` (what PowerPoint would repair or refuse), `overflows()` and `text_fit()` (what will not fit, as PowerPoint draws it), `collisions()` (text over text, lines through text), `measure_text()` (before building) |
| Round trips | a full-state SVG that can be edited anywhere and applied back |

Every example below is run by `tests/test_usage.py`, from a directory holding
`results.pptx` (the Japanese results deck `tests/fixtures/real-financial-report.pptx`) and
`smartart.pptx` (`tests/fixtures/powerpoint-smartart.pptx`).

## Common tasks

### Units: EMU everywhere, `Pt` for points

```python
from pptx_agent import Document, Pt

deck = Document.new()
box = deck.add_slide("Blank").add_textbox(Pt(72), Pt(72), Pt(288), Pt(36), "Units")
paragraph = box.text_frame.paragraph(0)
paragraph.space_after = Pt(6)                      # a length: EMU, and Pt(6) is 76,200
paragraph.run(0).size = 14                         # a font size: points
print(box.left, paragraph.space_after, paragraph.run(0).size)   # -> 914400 76200 14.0
try:
    paragraph.space_before = 6                     # 6 EMU: points meant, refused
except ValueError as error:
    print(type(error).__name__)                    # -> ValueError
```

This is the one place units are stated; everything else follows it. `Inches(1)` and
`Cm(2.54)` are EMU too (914,400), like `Pt`; `to_pt(914400)` is `72.0`.

| What | Unit |
| --- | --- |
| Positions and sizes (`left`, `width`, `bounds`), insets, margins and indents, `space_before`/`space_after`, line widths, `TextFit` heights, overflow amounts | **EMU**: 914,400 per inch, 12,700 per point (`EMU_PER_INCH`, `EMU_PER_POINT`). `Pt(6)` *is* the `int` 76,200, so it goes wherever a length does |
| Font sizes: `run.size`, `effective_size`, `TextFit.sizes`, `measure_text(size=)` | **points** (a `Pt` given here is read as its points) |
| `line_spacing` | a multiple of single spacing (`1.2`); `line_spacing_points` is points |
| `rotation` | degrees, clockwise |
| Colour modifiers in a colour string (`lumMod=75%`) | percent |

A paragraph spacing over 1,000 pt, or under a hundredth of a point, is refused (`ValueError`
naming `Pt`): it is points written as EMU. One under a point that is not a `Pt` draws a
`UnitWarning`.

### Open a deck, read it, edit by id

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
print(deck.to_outline(slides=[1]))      # Markdown; each block under <!-- its shape id -->

for shape in deck.slides[0].shapes:     # or walk the shapes
    print(shape.id, shape.kind, shape.name, shape.text[:20])

kpi = deck.shape("256.29")              # ids are "<slide id>.<shape id>"
print(kpi.text)                         # -> 11.9%
kpi.set_text("12.1%")                   # the figure stays bold, same font and colour
deck.save("out.pptx")
```

Ids never change when shapes or slides move. A shape id with `#` (`257.3#5`) is one the
deck itself numbered twice; use it as it is.

> **The outline is Markdown: do not paste its text into `set_text`.** `to_outline` escapes
> Markdown punctuation (`a\*b`, `\- item`) and marks bold as `**...**`; `set_text` writes
> what it is given, so the backslashes and asterisks would land on the slide. It warns
> with `MarkdownEscapeWarning` when it sees them. For text to edit from, read `shape.text`
> or use `outline_blocks()` / `find_text()` below: they return the raw text.

### Find text across a deck

```python
import re
from pptx_agent import Document

deck = Document.open("results.pptx")
for block in deck.find_text("4,285"):   # titles, text shapes, table cells, SmartArt nodes
    print(block.slide, block.kind, block.address, block.text)
    # 1 text 256.13 4,285 / 2 cell 257.3#5/cell1,1 4,285億円 / 3 cell 258.4#3/cell5,1 4,285億円
    target = deck.resolve(block.address)
    target.text = block.text.replace("4,285", "4,310")        # formatting kept

percentages = deck.find_text(re.compile(r"\d+\.\d%"))          # a compiled pattern searches
everything = deck.outline_blocks()                              # every block, raw, in order
```

A block's `text` is raw: paragraphs joined by `"\n"`, a line break as `"\v"`, nothing
escaped. It is what the object's `text` setter takes back unchanged.

### Text with mixed formatting

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
box = deck.shape("256.5")
box.set_text("対象期間：2025年10月1日〜2025年12月31日\n発表日：2026年2月6日")
# set_text diffs old against new, so only the changed characters change, and they take the
# formatting of the characters they replace.  "\n" separates paragraphs, "\v" breaks a line.

paragraph = box.text_frame.paragraph(1)            # or deck.resolve("256.5/p1")
paragraph.segment(["発表日：", "2026年2月6日"])    # split into runs, text unchanged
paragraph.run(1).format(bold=True, color="accent2")
paragraph.add_run("（予定）", size=10, italic=True)
paragraph.alignment = "r"
box.text_frame.add_paragraph("速報値", like=0)
deck.resolve("256.5/p2/r0").set_hyperlink("https://example.org/ir")
print([run.text for run in paragraph.runs])        # -> ['発表日：', '2026年2月6日', '（予定）']
```

Addresses: `"256.5/p1"` is a paragraph, `"256.5/p1/r0"` a run (0-based, re-resolved
each time, so they name positions).

`format(...)` on a run, a paragraph or a whole frame -- and the `**formatting` of
`add_run` and `add_paragraph` -- takes run properties only: `bold`, `italic`, `strike`,
`underline`, `size` (points), `typeface` (or `font`), `color`, `hyperlink`. Any other key
raises `TypeError` listing these. Paragraph properties (`alignment`, `level`,
`margin_left`, `indent`, `space_before`, `space_after`, `line_spacing`, bullets) are set
on the paragraph itself.

### Text frames: insets, anchor, wrapping, bullets that hang

```python
from pptx_agent import Document, Pt

deck = Document.new()
slide = deck.add_slide("Blank")
card = slide.add_shape("rect", 914400, 914400, 2743200, 1371600, text="Diagnose")
frame = card.text_frame
print(frame.autofit, frame.wrap, frame.insets, frame.anchor)
# -> none True (91440, 45720, 91440, 45720) middle
frame.insets = (Pt(6), Pt(6), Pt(6), Pt(6))        # (left, top, right, bottom); None inherits
frame.anchor = "top"                               # "top", "middle", "bottom" ("t", "ctr", "b")
label = slide.add_textbox(914400, 2743200, 2743200, 457200, "Exactly this box",
                          autofit="none")
print(label.text_frame.autofit, label.text_frame.anchor)   # -> none top
notes = slide.add_textbox(914400, 3657600, 4572000, 914400, "Kick-off\nSteerCo 1")
for paragraph in notes.text_frame.paragraphs:
    paragraph.set_bullet()                         # hangs, as PowerPoint hangs one
print(notes.text_frame.paragraph(0).margin_left, notes.text_frame.paragraph(0).indent)
# -> 285750 -285750
notes.text_frame.format(size=14, color="accent1")  # run formatting, every run
```

What each new shape's text frame starts with -- all of it settable on `text_frame`
(`autofit`, `wrap`, `insets`, `set_insets(left=...)`, `anchor`):

| Made by | autofit | wrap | insets (l, t, r, b) | anchor | alignment |
| --- | --- | --- | --- | --- | --- |
| `add_shape` | `"none"`: text may overflow | on | 91440, 45720, 91440, 45720 | `"middle"` | centred |
| `add_textbox` | `"shape"` by default: PowerPoint resizes the box to its text when it lays it out, so `height` is where it starts. `autofit="none"` keeps the box exactly as given; `"normal"` shrinks text on overflow once PowerPoint edits it | on | 91440, 45720, 91440, 45720 | `"top"` | left |
| `add_table` (each cell) | -- (a row grows to fit) | on | 91440, 45720, 91440, 45720 (the cell's margins) | `"top"` | left |
| `add_slide` (placeholders) | inherited from the layout and master: a body is usually `"normal"` | inherited | inherited | inherited (a title is usually `"middle"`) | inherited |

`set_bullet()` hangs the bullet the way PowerPoint does when it bullets a paragraph:
`margin_left` = h and `indent` = -h, with h stepping with the font size -- 171,450 EMU up
to 12 pt, 285,750 up to 18, 342,900 up to 25, 457,200 up to 35, and so on (measured;
`pptx_agent.edit.text.BULLET_HANGING`). A paragraph that already hangs -- a placeholder
whose list style indents it -- keeps its indents; `hanging=False` leaves them alone, an
`int` sets h in EMU.

### Tables

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
table = deck.shape("257.3#5").table
print(table.rows, table.columns)                   # -> 6 5
print(table.column_labels())   # -> ['項目', '当期実績', '前年同期', '増減額', '増減率']
print(table.row_labels())      # -> ['項目', '売上高', '売上総利益', '営業利益', '経常利益', '当期純利益']

cell = table.cell_by_label("営業利益", "当期実績")  # by what the table says
print(cell.address, cell.text)                     # -> 257.3#5/cell3,1 512億円
cell.text = "520億円"                              # cell.text reads AND writes; bold stays bold
table.cell(1, 1).text = "4,310億円"                # or by position: (row, column), 0-based

table.insert_row(like=5)                           # a new last row formatted like row 5
table.cell(6, 0).text = "EBITDA"
table.delete_column(2)
table.cell(6, 1).fill = "accent1 lumMod=20% lumOff=80%"
table.cell(6, 1).set_border("bottom", width=12700, color="accent1")
table.merge(6, 2, 6, 3)

rows = table.rows_fitting()                        # as drawn: rows grow to fit their text
print(rows.count, len(rows.heights), rows.past)    # -> 7 7 0
```

A row's stored height is a minimum: PowerPoint grows a row to fit its text, and the rows
below move down -- past the slide's bottom edge, where they are cut off, while the stored
heights still add up to a frame that fits. `drawn_row_heights` is each row as PowerPoint
draws it (the renderer's table layout, held to PowerPoint's PDF); `drawn_bounds` and
`overflows()` use it, and `rows_fitting()` (the example's last lines) says how many rows
fit above the slide's bottom edge, or any other line. `overflows()` reports a table that
runs past it as `off_slide` with `detail="rows"`, `rows` (the first and last row past the
edge, 1-based) and `rows_fit`: keep `rows_fit` rows on this slide and move the rest to a
table on the next.

Label lookups are strict: a label that names no row or column, or more than one, raises
`LabelError` listing the labels there are. Whitespace is collapsed; only when nothing
matches exactly are full-width forms and case ignored (`"売上高(億円)"` finds
`"売上高（億円）"`), and that must be unique too. The header row and label column are 0
unless you say otherwise: `cell_by_label(row, column, header_row=1, label_column=2)`.

### Charts

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
chart = deck.shape("257.25").chart
print(chart.chart_type, chart.categories, chart.series_names)
# -> bar ['Q1', 'Q2', 'Q3'] ['売上高（億円）', '営業利益（億円）']

revenue = chart.series["売上高（億円）"]           # by name (or chart.series[0])
print(revenue.value("Q3"), dict(revenue.items()))  # -> 4285 {'Q1': 3980, 'Q2': 4120, 'Q3': 4285}
revenue.set_value("Q3", 4310)                      # by category label (or set_value(2, ...))
chart.series["営業利益（億円）"].set_values([465, 488, 520])

chart.add_category("Q4", {"売上高（億円）": 4400})  # other series get a blank
chart.set_category("Q4", "Q4（予想）")
chart.add_series("純利益（億円）", [300, 315, 333, None])
chart.set_title("四半期業績")
chart.set_axis_title("value", "億円")
chart.set_legend(True, "b")
print(chart.number_formats)  # -> {'values': 'General', 'data_labels': None, 'value_axis': 'General'}
```

**The cache and the workbook move together.** A chart stores its numbers twice: in the
chart part (what PowerPoint draws) and in its embedded workbook (what Edit Data opens).
Every edit writes both, so the slide and the spreadsheet never disagree; formulas, cell
formats and a table over the data follow inserted and removed categories and series. A
chart without a workbook is edited in its cache only, with a `ChartDataWarning`.

Values are stored exactly as given. What decides how many decimals the slide shows is the
format of whatever displays them -- data labels and the value axis -- which
`chart.number_formats` reports; `data_labels` is `None` when no label shows a number,
and then a value's precision only moves a bar or a point. `chart.number_format` is the
values' own cached format (the workbook's, e.g. `General`).

Series names and category labels are as strict as table labels (`LabelError`). An `int`
is always a position; to address a numeric category by label use
`chart.category_index(2025)`.

### SmartArt

```python
from pptx_agent import Document

deck = Document.open("smartart.pptx")
diagram = deck.shape("256.3").diagram
print(diagram.texts)                               # -> ['Plan', 'Build', 'Ship']
diagram.set_text(1, "Build it")                    # by position, or diagram.nodes[1].text = ...
node = diagram.add_node("Measure")                 # at the top level, formatted like a sibling
node.add_child("Weekly")
diagram.remove_node(0)
print([(n.text, n.level) for n in diagram.nodes])
# -> [('Build it', 0), ('Ship', 0), ('Measure', 0), ('Weekly', 1)]
```

PowerPoint lays a diagram out again from its data when it opens the deck, so adding or
removing a node drops the cached drawing and PowerPoint redraws it.

### Fills and outlines, with theme colours

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
shape = deck.shape("256.3")
shape.fill = "accent1"                             # a theme colour stays a theme colour
shape.fill = "accent1 lumMod=75% lumOff=25%"       # with PowerPoint's tint and shade
shape.fill = "#1F4E79"                             # or RGB; None = inherit; "none" = no fill
shape.set_gradient_fill([(0, "accent1"), (1, "accent1 lumMod=50%")], angle=90)
shape.line.color = "accent2"
shape.line.width = 19050                           # 1.5 pt
shape.line.dash = "dash"
print(shape.fill.kind, shape.line.color)                   # -> gradient accent2
```

Colours are `"accent1"`...`"accent6"`, `"tx1"`, `"bg1"`, `"dk2"`, `"lt2"`, `"hlink"`...,
`"#RRGGBB"`, or one of those followed by modifiers (`lumMod=`, `lumOff=`, `tint=`,
`shade=`, `alpha=`, in percent). They are written unresolved, as PowerPoint writes them.

### Shapes, connectors and groups

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
slide = deck.slides[3]
box = slide.add_shape("roundRect", 914400, 914400, 1828800, 914400, text="Plan",
                      adjustments={"adj": 30000})
goal = slide.add_shape("ellipse", 5486400, 914400, 1371600, 1371600, text="Ship")
arrow = slide.add_connector("elbow", box.connection_site("right"), (goal, "left"),
                            line={"end": "triangle"})
goal.move_by(dy=914400)                            # the connector re-routes, still attached
note = slide.add_textbox(914400, 2286000, 3657600, 369332, "Notes")
group = slide.group([box, goal, arrow])
group.move_by(dx=457200)
copy = note.duplicate(dx=0, dy=457200)
copy.bring_to_front()
group.ungroup()
note.delete()
print(arrow.begin_connection[1], arrow.end_connection[1])   # -> 3 2
```

A connector runs from `begin` to `end`; its `start` arrowhead (OOXML's `head`) is at the
begin, its `end` arrowhead (`tail`) at the end -- so an arrow pointing at the shape it ends
on is `line={"end": "triangle"}`. Name a site by its side, `"top"`, `"right"`, `"bottom"`
or `"left"` (as drawn, rotation included), or by its index in `shape.connection_sites`,
numbered as PowerPoint numbers them: a rectangle's, a rounded rectangle's, a diamond's and
the flowchart boxes' are 0 top, 1 left, 2 bottom, 3 right; an ellipse's eight start at the
top and go counter-clockwise (2 left, 4 bottom, 6 right).
`slide.add_table(rows, columns, left, top, width, height)` and
`slide.add_picture("logo.png", left, top, width=...)` add the rest.

`pptx_agent.PRESETS` lists every preset name with its adjustment names
(`PRESETS["roundRect"]` is `("adj",)`, `PRESETS["rightArrow"]` `("adj1", "adj2")`).
`shape.fill` is the explicit fill only -- `None` for a new `add_shape` shape, which its
shape style fills -- and `shape.effective_fill` the fill as drawn (`accent1` there).
`Shape.duplicate` gives a copied group's members fresh ids too, and glues a connector inside
the copy to the copies of its shapes.

`shape.kind` is one of five, and a property a kind does not have reads `None` (`[]` for
`paragraphs`, `""` for `text`) rather than raising, so a loop over every shape needs no
kind test:

| `kind` | has |
| --- | --- |
| `"shape"` (autoshape, text box) | `text_frame`, `text`, `paragraphs`, `fill`, `line`, `preset`, `adjustments`, `connection_sites`, `text_fit()`, `fit_height()` |
| `"picture"` | `image_part`, `image_size`, `line`, `replace_image()` |
| `"connector"` | `line`, `route` (its drawn polyline), `begin_connection`, `end_connection` |
| `"group"` | `children`, `child_offset`, `child_extent` |
| `"graphic_frame"` | one of `table`, `chart`, `diagram` |

Every kind has `id`, `name`, `left`, `top`, `width`, `height` and `bounds` (the same four
as one tuple, as a layout's shapes spell it), `drawn_bounds` (where it is drawn on the
slide, rotation, flips, groups and a connector's route included), `rotation` and the flips.

### Slides: add from a layout, duplicate, reorder

```python
from pptx_agent import Document

deck = Document.new()                              # PowerPoint's own eleven layouts
print([layout.name for layout in deck.layouts])
cover = deck.add_slide("Title Slide")
cover.shapes[0].set_text("Q3 results")             # placeholders inherit from the layout
body = deck.add_slide("Title and Content")
body.shapes[1].set_text("Revenue grew 11%\nMargin 12.1%")
copy = body.duplicate(notes="Same figures, new owner")   # notes, charts, SmartArt copied
deck.move_slide(copy, to=0)                        # to a 0-based position; ids do not change
deck.delete_slide("s:257")                         # with everything only it used
print([slide.slide_id for slide in deck.slides])   # -> [258, 256]
cover.title = "Q3 results, restated"               # the title placeholder's text
print(deck.slide_titled("q3 results, restated").index)   # -> 1
```

A slide is named the same way everywhere: a `Slide`, its `sldId` (`257`) or `"s:257"`, as
the outline writes it. Positions are 0-based in `deck.slides[i]`, `move_slide(..., to=i)` and
`insert_outline(at=i)`; the slide *numbers* `to_outline`, `outline_blocks`, `find_text` and
the renderers take are 1-based, as PowerPoint numbers slides -- so pass a `Slide` or an
`"s:"` id when in doubt. `slide.index` is the 0-based position, read or called
(`slide.index()`). `slide.title` reads and writes the title placeholder's text (`None` when
the slide has none); `deck.slide_titled(text)` finds the one slide with that title, matched
as a table label is, and raises `LabelError` listing the titles otherwise.

### Draft slides from Markdown (`insert_outline`)

```python
from pptx_agent import Document

deck = Document.new()
deck.insert_outline("""\
# Q3 results

- Revenue **4,310** 億円
  - up 10.7% on last year
- Operating margin 12.1%

| Segment | Revenue |
| --- | --- |
| Digital | 1,842 |

Notes:

Lead with the margin.
""")
print(deck.to_outline())
```

`#` starts a slide; lists, tables, pictures and `Notes:` go into the deck's own layouts and
placeholders. Charts and SmartArt are not drafted (`OutlineWarning`). Whole insertion: one
undo step.

### Speaker notes

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
slide = deck.slides[0]
slide.notes = "Lead with the margin.\nThen the outlook."   # a notes page is made if needed
print(slide.notes_frame.paragraph(1).text)         # -> Then the outlook.
(block,) = deck.find_text("Lead with")             # notes are searched too
print(block.kind, block.address)                   # -> notes 256/notes
deck.resolve("256/notes/p0").text = "Lead with the 12.1% margin."   # formatting kept
copy = deck.duplicate_slide("s:256", notes="Repeat the headline.")
print(copy.notes)                                  # -> Repeat the headline.
deck.undo()                                        # the copy and its notes: one step
```

`slide.notes` reads and writes the notes as raw text (paragraphs joined by `"\n"`);
`slide.notes_frame` is a `TextFrame`, with paragraphs and runs like a shape's. A slide that
had no notes gets a notes page, and the deck the notes master PowerPoint adds with it. In
the outline the notes follow `Notes:`, under `<!-- 256/notes -->`.

### Does the text fit? (`text_fit`, `overflows`)

```python
from pptx_agent import Document

deck = Document.new()
points = "".join(f"- Point {k}\n" for k in range(1, 21))
(slide,) = deck.insert_outline("# Lessons learned\n\n" + points)
body = slide.shapes[1]
fit = body.text_fit()
print(fit.overflows, fit.autofit, fit.font_scale, fit.sizes[0])   # -> True normal 1.0 (28.0,)
print(fit.needed > fit.available, fit.overflow == fit.needed - fit.available)   # -> True True
print([problem.kind for problem in deck.overflows()])   # -> ['text']
body.set_text("\n".join(f"Point {k}" for k in range(1, 9)))
print(body.text_fit().overflows, deck.overflows())   # -> False []
```

**`overflows` is not `needed > available`.** The bottom of a text's last line box draws no
ink, so `overflows` allows a fifth of the tallest line (line spacing included), and never
less than 1,270 EMU (0.1 pt): a box sized tightly to one line is not reported, a line that
does not fit is. `overflow` is always `needed - available` (0 when it fits), so a `TextFit`
can say `overflows=False` with `overflow > 0`; `fit.slack` is the allowance, and
`overflows` is `overflow > slack`.

**A line can also be close to wrapping.** `fit.margin_to_wrap` is the least room, EMU, any
line leaves before the right edge of its text area, and `fit.near_wrap` says it is under
0.16 pt (`WRAP_MARGIN`). `text_fit` kerns as PowerPoint does -- with a face's legacy `kern`
table only -- but PowerPoint draws each glyph up to about 0.05 pt off its advance, so a
line within 0.16 pt of the edge may break in PowerPoint where `text_fit` keeps it whole:
give it room. Measured on 290 one-word boxes (`tools/wrap_boundary_probe.py`): 260 verdicts
agree, and every line PowerPoint wrapped that the measurement kept had less than 0.16 pt to
spare.

**Check fit with `text_fit`, not with a render.** A body set to shrink text on overflow
(`autofit` `"normal"`) is only shrunk when PowerPoint *edits* it; the file then stores the
scale, and on opening PowerPoint draws exactly what is stored. A body this library or any
generator filled stores none, so PowerPoint draws it at full size -- overflowing -- even if a
render shrinks it to fit. `text_fit` lays the text out the way pptx2svg measures it, at the
stored scale, and its verdicts agree with PowerPoint's own PDF in the oracle tests. Fix an
overflow by splitting the text, or say what PowerPoint should draw:
`frame.font_scale = 0.8`. `deck.overflows()` also lists collisions (`"overlap"`, next) and
shapes drawn past the slide's edge (`"off_slide"`, by their `drawn_bounds`: a rotated
shape's turned frame, a connector's route).

### Collisions: text over text, lines through text (`collisions`)

```python
from pptx_agent import Document

deck = Document.new()
slide = deck.add_slide("Blank")
bar = slide.add_shape("rect", 914400, 1828800, 3657600, 457200, text="Savings tracking")
slide.add_textbox(1143000, 1828800, 914400, 457200, "Q1")       # on the bar: layering
line = slide.add_connector("straight", (2743200, 914400), (2743200, 3657600),
                           line={"dash": "dash"})                  # drawn over the bar
first = slide.add_textbox(914400, 4114800, 1828800, 457200, "Kick-off")
second = slide.add_textbox(1371600, 4114800, 1828800, 457200, "SteerCo 1")
print([(p.detail, p.shape, p.other) for p in slide.collisions()])
# -> [('text', '256.6', '256.5'), ('line', '256.4', '256.2')]
line.send_to_back()                                # behind the bar now: hidden, fine
second.left = 3200400
print(slide.collisions())                          # -> []
elbow = slide.add_connector("elbow", (bar, "bottom"), (first, "top"))
print(elbow.route[0], elbow.drawn_bounds == elbow.bounds)   # -> (2743200.0, 2286000.0) False
```

`slide.collisions()` is the `"overlap"` items of `deck.overflows()` for one slide, each an
`Overflow` with a `detail`:

- `"text"`: two shapes holding text overlap -- their drawn boxes, or, for a text box with
  neither fill nor outline, the text itself (placed by its alignment and anchor);
- `"line"`: a line or connector crosses a text area; `amount` is the length crossed, EMU.
  A line behind an opaque shape is hidden by it and not reported;
- `"placeholder"`: a shape crosses what a title or another placeholder holds -- its text,
  or a picture, chart or table placeholder's frame.

`slide.collisions(boxes=True)` (and `deck.overflows(boxes=True)`) also reports `"box"`: two
text shapes whose boxes overlap although their text does not -- labels side by side whose
frames run into each other. A marker line drawn in front of bars is reported where it
crosses their text; sent behind them it is not, so there is no need to split it.

Layering is not a collision: a shape lying wholly on an opaque shape behind it (a label on a
bar, a title on a banner, text on a picture). A group's members are compared with each
other and the rest of the slide, never with their own group. Overlaps under a point are
ignored. `connector.route` is the connector's drawn polyline in slide EMU, from its begin to
its end, and `drawn_bounds` the box around it: PowerPoint writes an elbow that leaves
downwards with a rotated frame, so its `bounds` can reach far past the slide.

### Measure text before building (`measure_text`, `fit_height`)

```python
from pptx_agent import Document, measure_text

deck = Document.new()
m = measure_text("Order received", size=18, width=1371600, deck_or_shape=deck)
print(m.lines, m.height, m.near_wrap)              # -> ('Order', 'received') 640080 False
slide = deck.add_slide("Blank")
box = slide.add_textbox(914400, 914400, 1371600, m.height, "Order received", autofit="none")
print(box.text_fit().needed == m.height, box.text_fit().overflows)   # -> True False
label = slide.add_textbox(914400, 2743200, 650000, 450000, "Pass")
# PowerPoint breaks it "Pas / s", and so does text_fit
print(label.text_fit().lines)                      # -> (2,)
label.width = 652000
fit = label.text_fit()
print(fit.lines, fit.margin_to_wrap, fit.near_wrap)   # -> (1,) 1316 True
label.width += 2 * 12700
print(label.text_fit().near_wrap)                  # -> False
long = "Agree the scope with the steering committee before the pilot starts"
print(measure_text(long, width=1828800, deck_or_shape=box).lines)
# -> ('Agree the scope', 'with the steering', 'committee', 'before the pilot', 'starts')
box.set_text(long)
box.height = box.fit_height()                      # exactly as tall as its text needs
print(box.text_fit().overflows)                    # -> False
```

`measure_text(text, *, font=None, size, width, line_spacing=None, deck_or_shape=None)`
lays the text out exactly as `text_fit` lays a shape's out, before any shape exists:
`lines`, `breaks` (where each line starts in the text), `height` (EMU, the insets
included -- what a box needs), `widest` and `margin_to_wrap`. `width` is the box's, insets
included (`insets=` overrides PowerPoint's defaults). A deck or slide supplies its theme's
body font; a shape its font, size, line spacing, insets and wrapping, so
`measure_text(new_text, width=shape.width, deck_or_shape=shape)` says whether new text
will fit before it is set. `shape.fit_height()` is the height a shape needs for its text at
its present width.

### One text spec: build and measure the same text (`TextSpec`, `fit_box`)

```python
from pptx_agent import Document, ParagraphSpec, Pt, RunSpec, TextSpec, fit_box, measure_text

deck = Document.new()
slide = deck.add_slide("Title Only")
left, top, width, height = slide.content_area      # the band under the title, EMU
spec = TextSpec([
    ParagraphSpec([RunSpec("Discover", bold=True, size=16)], align="center"),
    ParagraphSpec([RunSpec("Interviews, data audit")], bullet="bullet", level=1),
], anchor="top")
m = measure_text(spec, width=Pt(180), preset="roundRect", deck_or_shape=slide)
box = slide.add_shape("roundRect", left, top, Pt(180), m.box_height, text=spec)
fit = box.text_fit()
print(fit.needed == m.height, fit.needed <= fit.available)   # -> True True
print(fit_box(spec, Pt(180), preset="roundRect", deck_or_shape=slide) == m.box_height)   # -> True
```

A `TextSpec` is a shape's whole text as one value: paragraphs of runs, each run with its
own `bold`, `italic`, `underline`, `strike`, `size` (points), `font`, `color` and
`hyperlink`; each paragraph's `align`, `bullet` (`"none"`, `"bullet"`, `"number"`),
`level`, `space_before`/`space_after` (EMU) and `line_spacing`; and the frame's `insets`,
`anchor`, `wrap` and `autofit`. `set_text`, `add_shape(text=...)` and
`add_textbox(text=...)` build it as one undo step; `measure_text(spec, ...)` builds it in a
shape of the named `preset` (or `"textbox"`), or `like=` an existing shape, on a copy of
the deck, and measures it as `text_fit` does -- so what is measured is what is built,
default insets, bold and bullets included. `box_height` (and `fit_box`) is the least
height at which the text fits; for a geometry whose text area narrows as it grows (a
chevron) it is found by halving. `bold=True` on plain-text `measure_text` measures bold
too: pass it whenever the text will be bold.

`slide.content_area` is where the layout puts content: its body placeholders together, or
the band under the title and above any footer the layout draws.

### Agent tools: a deck through tool calls (`pptx_agent.tools`)

```python
from ooxml_edit.tools import Toolbox
from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

deck = Document.new()
deck.add_slide("Title Only").title = "Plan"
with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
    session = toolbox.session()
    session.open(deck.to_bytes(), "plan.pptx")              # "d1"
    made = toolbox.dispatch(session, "ppt_add_shape", {"doc": "d1", "slide": "s:256", "items": [
        {"preset": "roundRect", "box": {"x": 66, "y": 160, "w": 180, "h": 60}, "text": "Plan",
         "ref": "plan"},
        {"preset": "roundRect", "box": {"x": 300, "y": 160, "w": 180, "h": 60}, "text": "Ship",
         "ref": "ship"}]})
    print(made.refs)                                       # -> {'plan': '256.3', 'ship': '256.4'}
    linked = toolbox.dispatch(session, "ppt_add_connector", {"doc": "d1", "items": [
        {"kind": "straight", "from": {"shape": "$plan"}, "to": {"shape": "$ship"},
         "line": {"end": "triangle"}}]})
    print(linked.checks["collisions"], linked.checks["validate"]["new"])   # -> [] []
    saved = toolbox.dispatch(session, "save_document",
                             {"doc": "d1", "name": "plan.pptx", "format": "pptx"})
    print(saved.data["size"] > 0, len(session.take_outputs()))   # -> True 1
```

`pptx_agent.tools` is the deck side of `ooxml_edit.tools`: the tools a model edits decks
with, no code and no files. Lengths are points at this boundary. Tools that make many
things take `items[]`, an item may name what it makes (`"ref": "plan"`) and any later
target may be `"$plan"`; `batch` runs calls to different tools as one, all or nothing.
Every changing call returns `checks` for the slides it touched -- overflows, collisions,
shapes off the slide, lines close to wrapping, fits within the allowance (stated), and
new validation problems -- and, when there are any, `layout`: at most five facts about the
shapes the call touched (a box 1-3 pt off the line its like neighbours are on, one uneven gap
in a row of like boxes, a size unlike its like boxes', a label far from its marker compared
with the others), each with the exact call that resolves it, such as
`ppt_set_shape 256.6 y=150` (`pptx_agent.edit.feedback`). Facts, not rules: a difference may
be intended.

### Review comments (`ppt_comments`, `Document.add_comment`)

```python
import datetime as dt
from ooxml_edit.tools import Toolbox
from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

deck = Document.new()
slide = deck.add_slide("Title Only")
slide.title = "Plan"
figure = slide.add_shape("rect", 914400, 1828800, 2743200, 914400)
figure.text = "2,140"
thread = deck.add_comment(figure, "Is this the audited figure?", author="Dana Reviewer",
                          date=dt.datetime(2026, 10, 7, 9, 0, tzinfo=dt.timezone.utc))
print(thread.target, thread.initials, thread.done)          # -> 256.3 DR False
with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
    session = toolbox.session()
    session.open(deck.to_bytes(), "plan.pptx")
    toolbox.dispatch(session, "ppt_comments", {
        "doc": "d1", "action": "reply", "author": "Sam Author",
        "items": [{"comment": thread.address, "text": "Yes, the September close.",
                   "resolve": True}]})
    listed = toolbox.dispatch(session, "ppt_comments", {"doc": "d1", "action": "list"})
    print(listed.data[0]["done"], listed.data[0]["replies"][0]["author"])  # -> True Sam Author
```

PowerPoint's modern comments: a thread on a slide (`s:256`) or a shape (`256.3`), with
replies, resolved or open, in the parts PowerPoint for Mac reads and writes back
(`ppt/comments/modernComment_*.xml`, `ppt/authors.xml`). `ppt_comments` takes
`word_comments`' actions -- list, add, reply, resolve, reopen, edit, delete -- many at once.
A comment's address is `c:` and eight hex digits. A duplicated slide does not take its
original's comments; deleting a slide removes them.

Positions in the tools are the model's own, in points. The layout library is still there for
code -- `pptx_agent.edit.arrange` (align, distribute), `pptx_agent.edit.layout` (stack,
column, grid, place_labels), `pptx_agent.edit.scales` and `Slide.copy_shapes` -- but its
tools (`ppt_layout`, `ppt_scale`, `ppt_copy`, `ppt_align`) were removed after trial 3, where no
model called them. `ppt_design_facts` reports the palette, colour-coded sets and legends,
empty regions, alignment, shape vocabulary, text sizes and lines over text -- facts, no
verdicts (`Slide.design_facts`).

### Layouts and their placeholders

```python
from pptx_agent import Document

deck = Document.new()
layout = deck.layout("Title and Content")
title, body = layout.placeholders[:2]
print(title.type, title.idx, title.bounds)         # -> title None (838200, 365125, 10515600, 1325563)
print(body.type, body.idx, body.sizes[:3])         # -> None 1 (28.0, 24.0, 20.0)
slide = deck.add_slide(layout)
content = next(s for s in slide.shapes if s.placeholder == (None, 1))
print(content.placeholder)                         # -> Placeholder(type=None, idx=1)
print([shape.kind for shape in layout.shapes][:2])   # -> ['shape', 'shape']
```

`layout.placeholders` lists what a slide made from the layout gets: each placeholder's
`type` (`None` is a content placeholder, `"obj"`), `idx` (what a slide's placeholder matches
it by), `name`, `bounds` (`left, top, width, height` in EMU) and the default text size of
each level. `shape.placeholder` is the same `(type, idx)` named tuple. `layout.shapes` are
read-only: layouts are edited in PowerPoint's slide master view.

### Theme colours, fonts and formatting as drawn

```python
from pptx_agent import Color, Document

deck = Document.open("results.pptx")
print(deck.theme.colors["accent1"], deck.theme.colors["tx1"])   # -> #4472C4 #000000
print(deck.theme.fonts.major, deck.theme.fonts.minor)   # -> Calibri Light Calibri
print(Color.parse("accent1 lumMod=75%").resolve(deck))   # -> #2F5597
box = deck.shape("256.29")
box.fill = "accent2"
print(box.fill.color, box.fill.color.resolve(box))   # -> accent2 #ED7D31
run = box.text_frame.paragraph(0).run(0)
print(run.size, run.effective_size, run.effective_font)   # -> 18.0 18.0 Noto Sans JP
print(box.text_frame.paragraphs[0].effective.alignment)   # -> left
print(box.text_frame.autofit, box.text_frame.font_scale)   # -> none 1.0
```

A colour stays a theme reference in the file; `resolve` says what it looks like today,
through the slide's colour map, with every modifier applied as PowerPoint composes them.
`effective_size` and `effective_font` follow the inheritance -- run, paragraph, list styles,
layout and master placeholders, the master's text styles, the presentation -- resolve
`+mn-lt` to the theme's font, and apply a stored autofit scale. `autofit` is `"none"`,
`"normal"` (shrink text on overflow) or `"shape"` (resize shape to fit text), and settable.

### Theme: change its colours and fonts; which colours do which job

```python
from pptx_agent import Document

deck = Document.new()
roles = deck.theme.roles
print(roles.primary, roles.secondary, roles.neutral, roles.highlight)   # -> dk2 accent1 lt2 accent2
print(roles.source["primary"])                     # -> convention
print([tint.color for tint in deck.theme.tints("accent1")][:2])
# -> ['accent1 lumMod=20% lumOff=80%', 'accent1 lumMod=40% lumOff=60%']
theme = deck.theme.set_colors({"accent1": "#0B6E79", "dk2": "#0B2545"})
theme = theme.set_fonts(major="Georgia", minor="Arial")
print(theme.colors["accent1"], theme.fonts.major)  # -> #0B6E79 Georgia
box = deck.add_slide("Blank").add_shape("rect", 914400, 914400, 914400, 914400)
box.fill = deck.theme.tints(roles.neutral)[1].color
```

`set_colors` sets slots (`dk1`...`folHlink`, or `tx1`/`bg2`... through the colour map) to
RGB, and `set_fonts` the heading (`major`) and body (`minor`) faces; every reference in the
deck follows, each call is one undo step, and each returns the theme as it now is (a
`Theme` read earlier keeps what it read).

**Choosing colours like the template's own.** `theme.roles` names the slots acting as
`primary` (the colour the template carries its identity in: titles, bars, rules),
`secondary` (its partner), `neutral` (what boxes and panels are tinted from) and
`highlight` (kept for the one thing that must stand out), with `colors` (RGB), `source`
(`"master"` or `"convention"`) and `usage`. The rule: count every theme colour the master
and its layouts use themselves -- fills, outlines, text, the background and the title
style, `tx2`/`bg2` mapped to their slots, `dk1`/`lt1` and the hyperlink colours left out;
`primary` is the most used of `dk2` and `accent1`-`accent6`, `secondary` the next,
`highlight` the most used remaining accent (ties go to `dk2`, then the accents in order).
What a template does not show, Office's convention fills: `dk2` primary, `accent1`
secondary, `accent2` highlight; `neutral` is `lt2`. Restraint reads better than a rainbow:
one carrier colour, tints of it and of the neutral for boxes, the highlight once.

`theme.tints(slot)` gives a colour's recommended variants, lightest first, as theme
references (`Tint(name, color, hex)`): lighter 80%, 60%, 40% and darker 25%, 50% for a
mid-tone (`lumMod`/`lumOff` 20/80, 40/60, 60/40, 75, 50); lighter 90%...10% for a colour
darker than 20% luminance, darker 10%...90% for one lighter than 80%. `theme.ramps` has
them for `dk2`, `lt2` and every accent.

### Pictures: native size, and a new image without distortion

```python
from pptx_agent import Document

deck = Document.new()
slide = deck.add_slide("Blank")
logo = slide.add_picture(deck.slides[0].render_png(width=200), 10972800, 228600,
                         width=731520, height=731520)
print(logo.image_size.width, logo.image_size.height)   # -> 200 113
wide = slide.render_png(width=300, height=120)
right = logo.left + logo.width
logo.replace_image(wide, keep="height", anchor="top_right")   # height and corner stay
print(logo.height, logo.width, logo.left + logo.width == right)   # -> 731520 1828800 True
```

`image_size` is the stored image's pixels and the resolution it states (`dpi`, `None` when it
states none). `replace_image` keeps the frame by default; `keep="height"` or `"width"` lets
the other side follow the new image's proportions, `"none"` places it at its pixel size, and
`anchor` is the point of the frame that stays put.

### Full-state SVG

```python
from lxml import etree
from pptx_agent import Document

deck = Document.open("results.pptx")
svg = deck.slides[0].render_svg(full_state=True)  # the slide's OOXML rides in data-ooxml-*
tree = etree.fromstring(svg.encode())
group = tree.find(".//{*}g[@data-pptx-id='256.29']")
group.set("data-ooxml-x", str(int(group.get("data-ooxml-x")) + 914400))   # move it right
report = deck.apply_svg(etree.tostring(tree))      # typed edits, raw XML, new shapes: one undo
print(report.edited)                               # -> {'256.29': ['x']}
```

The vocabulary (`data-ooxml-x`, `-fill`, `-text`, `-table`, `-chart-data`...) is in
ROADMAP.md, "Phase E3". An SVG is untrusted input: one this library did not write, or one
with a DTD, an entity or invalid data, is refused before anything changes.

### New decks and templates

```python
from pptx_agent import Document

deck = Document.new(size="16:9", title="Plan", author="Ada", language="ja-JP")
deck.add_slide("Title Slide").shapes[0].set_text("計画")
deck.save_as_template("brand.potx")                # or deck.save("brand.potx")

branded = Document.new(template="brand.potx", title="Q4")  # its masters, layouts, theme
branded.add_slide("Title and Content")
print(branded.slide_size, len(branded.slides), branded.title)   # -> (12192000, 6858000) 1 Q4
branded.save("q4.pptx")
```

`Document.new(template=...)` starts a deck from a template's masters, layouts and theme,
without its slides. `Document.open("brand.potx")` opens the template *itself*, and warns
(`TemplateOpenedWarning`). Either way, `save` writes what the extension says -- `.pptx` a
presentation, `.potx` a template -- because PowerPoint refuses a file whose extension and
declared type disagree; `deck.validate(target="out.pptx")` names such a mismatch.

### Render, validate, save

```python
from pathlib import Path
from pptx_agent import Document

deck = Document.open("results.pptx")
deck.shape("256.29").set_text("12.1%")
png = deck.slides[0].render_png(width=1280)        # bytes; needs pptx-agent[png]
Path("slide1.png").write_bytes(png)
svg = deck.slides[0].render_svg()                  # each shape's <g data-pptx-id="256.29">

problems = deck.validate()                         # what PowerPoint would repair
print([str(p) for p in problems])                  # this deck arrives with three
deck.save("out.pptx")
```

`validate()` returns `Problem(code, part, detail)` values; a deck may arrive with some, so
compare before and after an edit. No edit adds one. Only PowerPoint itself can say for
certain that a deck opens: the test suite's opt-in oracle drives it.

### Undo and batches

```python
from pptx_agent import Document

deck = Document.open("results.pptx")
with deck.batch():                                 # one undo step; rolled back on an exception
    deck.shape("256.13").set_text("4,310")
    deck.shape("257.3#5").table.cell_by_label("売上高", "当期実績").text = "4,310億円"
deck.undo()
print(deck.shape("256.13").text)                   # -> 4,285
deck.redo()
```

## Errors and warnings

| | When |
| --- | --- |
| `LabelError` (a `KeyError`) | a row, column, series or category label, or a slide title (`slide_titled`), names nothing or more than one; the message lists the labels, and `label`, `what`, `candidates` and `matches` hold them as data |
| `MarkdownEscapeWarning` | `set_text` or a `text` setter got text that looks copied from `to_outline` (`\*`, `\-` at a line's start, `11\.`, `**bold**`) that the old text did not have; it is still written as given |
| `ChartDataWarning` | a chart has no embedded workbook, so only its cache changed |
| `ChartDataError` | a chart edit that cannot be made (the wrong number of values, the last series) |
| `OutlineWarning` | `insert_outline` left a chart or SmartArt block out |
| `TemplateOpenedWarning` | `Document.open` was given a template: it edits the template itself; a deck *from* it is `Document.new(template=...)` |
| `FullStateError` / `NotFullStateSvg` | `apply_svg` refused an SVG; nothing changed |
| `UnitWarning` | a paragraph spacing under a point that is not a `Pt`: probably points written as EMU (under a hundredth of a point, or over 1,000 pt, it is a `ValueError`) |
| `TypeError` from `format` | a formatting key no run property has; the message lists them |
| `KeyError` / `IndexError` | an id, address or position that does not exist |
