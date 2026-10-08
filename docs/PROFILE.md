# The SVG authoring profile (library only; the `ppt_draw` tool was removed)

You draw a slide graphic as **SVG**, in the small vocabulary below, and one call turns it
into **native, editable PowerPoint objects**: preset shapes (chevrons, rounded rectangles...),
real text frames with wrapping, insets and anchors, connectors glued to shapes, groups, and
theme colours. It is not a picture: everything stays editable in PowerPoint.

**Status.** The converter, `pptx_agent.edit.svgprofile.apply_svg_graphic`, is a library call.
Its tool, the experimental `ppt_draw`, was removed from `pptx_agent.tools` after trial 3: SVG
authoring never beat shape calls (the spike, arm C and trial 3's draw arm), so models build
graphics with `ppt_add_shape` and `ppt_add_connector`. The text below describes the profile as
the tool took it; the converter takes the same SVG. The `ppt_scale` references are historical:
that tool was removed too.

```json
{"tool": "ppt_draw", "arguments": {"doc": "d1", "slide": "s:257", "svg": "<svg viewBox=...>...</svg>"}}
{"tool": "ppt_draw", "arguments": {"doc": "d1", "slide": "s:257", "svg": "<g id=...>...</g>", "measure": true}}
{"tool": "ppt_draw", "arguments": {"doc": "d1", "slide": "s:257", "svg": "...", "replace": "lead"}}
```

- `describe` gives each slide's content area (points), the theme's colours, roles and tints.
- The result gives `ids` (SVG id -> the shape address it became), `text_fits` (every text
  shape, measured as PowerPoint lays it out), `warnings` (approximations and repairs made), and
  the usual `checks`: overflows and collisions for the slide (`[]` is the goal).
- Errors are `invalid_arguments` naming the element (its id, or its tag and number: `<rect> #3`
  is the third `<rect>`) and line. When an error is returned, **nothing was changed**. One call
  is one undo step.
- To fix a graphic, `undo` and draw it again, or redraw one top-level group in place with
  `replace: "ID"`: only the element with that id is taken from the SVG (same viewBox), the old
  group is deleted and the new one takes its place in the z-order. Connectors outside the group
  that were glued into it come loose (the result lists them), so keep a connector and the boxes
  it joins in the same group.
- `box` {x, y, w, h} (points on the slide) places the viewBox somewhere other than the content
  area.

## Coordinates and units

- The root `<svg>` needs a `viewBox`. The viewBox is mapped onto `box`, which defaults to the
  slide's content area. As in SVG, `preserveAspectRatio` defaults to `xMidYMid meet`: one uniform
  scale, centred. **Use `viewBox="0 0 W H"` with the content area's width and height in points**:
  then one unit is one point and nothing is distorted.
- **Points, never scaled:** `font-size`, `stroke-width`, `data-inset`, `data-space-before/after`.
  Write bare numbers, or `pt`/`px` (both are read as points). `%`, `em` and `in` are not accepted.
- `transform` on any element: `translate`, `scale`, `rotate(a [cx cy])` (and `matrix` without skew).
  A rotated rectangle becomes a rotated shape. A negative scale becomes a flip.

## Shapes

| SVG | Becomes |
|---|---|
| `<rect x y width height>` | a rectangle; with `rx`, a rounded rectangle |
| `<rect ... data-preset="chevron" data-adj-adj="30000">` | that PowerPoint preset in that box (`homePlate`, `chevron`, `roundRect`, `diamond`, `ellipse`, `triangle`, `rightArrow`, `flowChartProcess`... any preset name), with its adjust values in raw DrawingML units |
| `<path d=... data-preset="homePlate">` | the preset, in the path's bounding box (the path itself is only a preview) |
| `<ellipse cx cy rx ry>`, `<circle cx cy r>` | an oval |
| `<path d>`, `<polygon points>`, `<polyline points>` without `data-preset` | a freeform (custom geometry). Supports M L H V C S Q T A Z, absolute and relative |
| `<line x1 y1 x2 y2>` | a straight connector line, not attached to any shape |
| `<g>` | a group, if it holds two or more objects (nested groups stay nested) |

**Prefer presets.** A chevron, a pentagon, a rounded box or a diamond drawn as a free path is
harder to edit in PowerPoint. Use `data-preset`.

**Adjustment names** (`data-adj-NAME`, raw DrawingML units: 100000 = 100% of the shorter side
unless noted; a wrong name is an error that lists the right ones):

| Preset | Adjustments (default) | What they do |
|---|---|---|
| `chevron`, `homePlate` | `adj` (50000) | depth of the point |
| `roundRect` | `adj` (16667) | corner radius; `rx` on a `<rect>` sets it for you |
| `round2SameRect` | `adj1` (16667), `adj2` (0) | top corners, bottom corners |
| `snip1Rect`, `plaque`, `foldedCorner` | `adj` (16667) | the cut or fold |
| `triangle` | `adj` (50000) | apex position, 0 = left, 100000 = right |
| `parallelogram`, `trapezoid` | `adj` (25000) | slant |
| `hexagon` | `adj` (25000), `vf` (115470) | point depth |
| `octagon` | `adj` (29289) | corner cut |
| `plus`, `donut`, `can`, `cube` | `adj` (25000) | arm width, ring thickness, top depth |
| `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `leftRightArrow`, `upDownArrow`, `notchedRightArrow` | `adj1` (50000), `adj2` (50000) | shaft thickness, head length |
| `wedgeRectCallout` | `adj1` (-20833), `adj2` (62500) | tail tip x, y (from the centre, in 1/100000 of width/height) |
| `wedgeRoundRectCallout` | `adj1`, `adj2`, `adj3` (16667) | as above, plus corner radius |
| `leftBracket`, `rightBracket` | `adj` (8333) | curl |
| `leftBrace`, `rightBrace` | `adj1` (8333), `adj2` (50000) | curl, position of the point |
| `pie`, `arc`, `chord` | `adj1`, `adj2` (angles, 60000 = 1 degree, clockwise from 3 o'clock) | start and end angle |
| `blockArc` | `adj1`, `adj2` (angles), `adj3` (25000) | start, end, thickness |
| `rect`, `ellipse`, `diamond`, `flowChartProcess`, `flowChartDecision`, `flowChartMerge` | none | |

**Names:** `id="..."` (or `data-name`) becomes the shape's name in PowerPoint's selection
pane, and its key in the `ids` that `ppt_draw` returns. Ids must be unique.

**Z-order** is document order: later elements are drawn on top.

## Colours

`fill` and `stroke` take:
- **theme tokens:** `accent1`...`accent6`, `dk1`, `lt1`, `dk2`, `lt2`, `tx1`, `tx2`, `bg1`, `bg2`, `hlink`;
- **a tint or a shade:** `accent1/tint40` (40% lighter) or `dk2/shade25` (25% darker), in steps of 1..90;
- **`#RRGGBB`** (hard-coded RGB: avoid it when the template's colours will do);
- **`none`.**

They are written as theme references, so they follow the template.
- CSS colour names (`red`, `white`) are rejected; use `bg1`/`lt1` for white and `tx1`/`dk1` for black.
- `opacity`, `fill-opacity` and `stroke-opacity` become transparency.
- `<linearGradient>` in `<defs>`, used as `fill="url(#id)"`, with theme-token stops, becomes a gradient fill.

`describe` gives the theme's colours, which ones carry the template's identity (`roles`:
`primary`, `secondary`, `neutral`, `highlight`) and their recommended tints in this syntax.

Defaults, as in SVG:
- **fill:** a shape with no `fill` (on it or on an ancestor `<g>`) is filled `tx1` (dark). Say `fill="none"` for no fill.
- **stroke:** no `stroke` means no outline.
- **lines:** a `<line>` with no stroke is drawn `tx1`.

**Stroke:**
- `stroke-width` is in points.
- `stroke-dasharray` maps to the nearest PowerPoint dash; a warning says which. Or name the dash directly with `data-dash="dash|sysDash|sysDot|lgDash|dashDot"`.
- Arrowheads: `data-arrow-end="triangle"` and `data-arrow-start=...` (`triangle`, `stealth`, `diamond`, `oval`, `arrow`). SVG `marker` elements are not supported.

## Text: always in a box

Text is never a free-floating baseline string. Each text block is one PowerPoint text frame, which
wraps inside its box. There are three forms:

**1. A shape holding text:** a `<g data-text-box>` with exactly one shape (`rect`, `ellipse`, or a
`path`/`rect` with `data-preset`) and one `<text>`. The shape gives the geometry and fill; the text goes inside it.

```xml
<g id="phase1" data-text-box="" data-anchor="middle" data-inset="4 8">
  <rect data-preset="homePlate" x="0" y="0" width="215" height="44" fill="dk2"/>
  <text font-size="14" fill="bg1" text-anchor="middle">
    <tspan font-weight="bold">Diagnose</tspan>
    <tspan font-size="11">Weeks 1–3</tspan>
  </text>
</g>
```

A shape element may also contain the `<text>` directly: `<rect ...><text>…</text></rect>`.

**2. A plain text box:** `<text x y data-width data-height>`. Here **x, y are the box's top-left
corner**, not a baseline. It is anchored to the top by default. Optional `data-fill`, `data-stroke`
and `data-stroke-width` give it a fill or outline.

```xml
<text id="src" x="0" y="388" data-width="900" data-height="14" font-size="9" fill="tx1/tint40">
  Source: Halden Group ERP spend data FY2025
</text>
```

**3.** Text anywhere else (a `<text>` without `data-width`/`data-height`) is an error.

**Paragraphs and runs:**
- **Paragraphs:** each direct `<tspan>` child of `<text>` is a paragraph. A `<text>` without
  tspans is one paragraph. Don't put loose text between paragraph tspans.
- **Runs:** `<tspan>`s nested inside a paragraph are runs with their own `font-weight`,
  `font-style`, `fill`, `font-size` and `text-decoration="underline"`.
- **Whitespace** collapses as in SVG, so line breaks in your source don't matter. `x`/`y`/`dy` on
  tspans are ignored: PowerPoint lays out the lines.
- **Bullets:** `<tspan data-bullet="">` gives a bullet paragraph (`•`), and `data-bullet="–"` a
  custom character. The bullet hangs as PowerPoint hangs it. `data-level="1"` indents it a level.
- **Spacing:** `data-space-before`/`data-space-after` (pt) and `data-line-spacing="1.1"` (a multiple),
  on a paragraph tspan or on the `<text>` (for all its paragraphs).

**Text attributes** inherit from `<g>` → `<text>` → `<tspan>`:

| Attribute | Default |
|---|---|
| `font-size` (pt) | **12** |
| `font-weight` (`bold` or ≥ 600) | |
| `font-style` | |
| `fill` (the text colour) | **`tx1`** |
| `font-family` | the theme's body font; `major` gives the heading font, `minor` the body font |

**The text frame**, on the `g`, the shape or the `<text>` (when more than one has it, the most
specific wins: the `<text>`, then the shape, then the `g`):

| Attribute | Values | Default |
|---|---|---|
| `text-anchor` | `start`/`middle`/`end` → left/centre/right | left |
| `data-align` | `left|center|right|justify`; on a paragraph tspan, overrides `text-anchor` | |
| `data-anchor` | vertical: `top|middle|bottom` | `middle` in a shape, `top` in a plain text box |
| `data-inset` | in pt: `"6"` (all sides), `"4 8"` (vertical horizontal) or `"l t r b"` | PowerPoint's 7.2 pt left/right, 3.6 pt top/bottom |
| `data-wrap` | `"none"` turns wrapping off | wrapping on |

Text boxes never auto-fit: the box is exactly what you draw. With no `data-inset`, PowerPoint's
default insets apply (7.2 pt left and right, 3.6 pt top and bottom), and they count: a 12 pt line
needs 14.4 pt + 7.2 pt = 21.6 pt of height.

### Measuring

`ppt_draw` with `measure: true` lays text out **exactly as drawing will build it**: the same
element, so the same insets, sizes, bold runs, bullets, spacing, wrap and shape geometry -- and
changes nothing. Write the box, measure it, then use the same markup in the graphic.

The `svg` can be just the element(s), e.g. a `<g data-text-box>...</g>` or a
`<text x y data-width data-height>`; a fragment is placed in a viewBox where one unit is one point.
A whole `<svg>` is measured in its own viewBox (connectors are skipped). For each text shape (by
id) it returns `fits`, `needed` (lines + spacing + top/bottom insets, pt), `available`, `lines`
(per paragraph) and `height_to_fit`: the least `height` (in SVG units) that fits at that width.
`fits` allows the last line's empty bottom (it draws no ink), so it can be true while `needed`
is a little over `available`.

## Connectors that stay attached

```xml
<line data-from="lead" data-to="ws1" stroke="tx1/tint50" stroke-width="1"/>
<path d="M0 0" data-from="sc" data-to="pmo" data-from-side="right" data-to-side="left"
      data-connector="elbow" stroke="dk2" data-arrow-end="triangle"/>
```

- `data-from`/`data-to` name shapes by `id`, either earlier or later in the document; a group can't be named.
- The line is glued to those shapes' connection sites: move a box in PowerPoint and the connector follows.
- **Sites:** `data-from-side`/`data-to-side` (`top|right|bottom|left`) pick them. Without them:
  a target wholly **below** the source (its top under the source's bottom) is joined source
  bottom -> target top, however far to the side it is (org charts, trees); wholly above, top ->
  bottom; otherwise, wholly to the right or left, side to side. Give one side and the other is
  still chosen automatically.
- **Kind:** `data-connector` is `straight`, `elbow` or `curved`. Without it, the connector is straight
  when the two sites line up and elbow otherwise.
- **Coordinates:** a glued line's `x1..y2` (or `d`) are ignored.
- **Order:** put connectors **before** the boxes they join, so the boxes are drawn on top of them.

## Data positions: scales

Proportional placement -- a date on a timeline, a score on an axis, a row of a Gantt -- is
where hand-written coordinates went wrong. Declare a scale once with `ppt_scale` (the same
scales the shape tools use) and write data values instead of numbers:

```xml
<g data-scale-x="$time" data-scale-y="$rows">
  <rect id="b1" x="@2026-11-16" data-x2="@2026-12-18" y="@Category strategy"
        data-y2="@Category strategy" fill="accent1"/>
  <line x1="@2027-01-22" x2="@2027-01-22" y1="@Spend & baseline|start"
        y2="@Change & capability|end" stroke="tx1" data-dash="dash"/>
  <ellipse cx="@2027-02-12" cy="@Negotiations" rx="6" ry="6" fill="accent2"/>
</g>
```

- Inside an element with `data-scale-x="$name"` (or `data-scale-y`), a position whose value
  starts with `@` is a data value of that scale: a number (linear), a date `YYYY-MM-DD`, or a
  band name.
- Position attributes: `x`, `cx`, `x1`, `x2` (and the `y` ones), plus `data-x2`/`data-y2` (a
  `rect`'s or boxed `text`'s right or bottom edge: it sets the width or height) and
  `data-cx`/`data-cy` (its centre, with `width`/`height` given).
- A date or a band has a start, a centre and an end. `x`/`y` mean its start, `data-x2`/`data-y2`
  its end, centres and line ends its centre; add `|start`, `|center` or `|end` to say otherwise.
  So `x="@2026-11-16" data-x2="@2026-12-18"` spans both days, inclusive.
- A scale's `from`/`to` are **slide points**, as in the shape tools; the converter maps them into
  your viewBox. Data values cannot sit under a `transform`.
- `ppt_scale` with `ticks` returns the positions and labels of gridlines or month headers.

## Escaping

The SVG is XML, but you need not escape `&` by hand: a bare `&` that does not start a reference
(`R&D`, `Spend & baseline`) is taken as a literal ampersand, and a `<` followed by a space, digit or
`=` (`< 5%`) as a literal `<`. `&amp;`, `&lt;`, `&gt;`, `&quot;`, `&apos;` and numeric references
(`&#8211;`) work as usual; HTML names (`&ndash;`, `&nbsp;`, `&rarr;`) become their characters. A
`warnings` line says what was repaired. Characters such as `–`, `•`, `→` can also be typed
directly (UTF-8).

## Not in the profile (returns an error naming the element)

- `filter`, `mask`, `clipPath`/`clip-path`, `pattern`, `radialGradient`, `image`, `use`, `symbol`;
- `<style>` and `class`;
- `foreignObject`, `script`, `marker`/`marker-*`, `textPath`, `switch`, `a`;
- skewed presets or text;
- CSS colour names;
- a DTD or entities.

`style="..."` with plain presentation properties (`fill:…; stroke-width:…`) is accepted.

## Known limits of this prototype

- **`render`** (pptx2svg) does not draw the **outline of freeform (custom-geometry) shapes**, a
  renderer bug (LR5). PowerPoint draws them. Fills, presets, lines and connectors render
  correctly.
- **Rotation:** a rotated group's members are rotated individually; the group itself is not rotated.
- **Opacity** on a `<g>` is multiplied into each member's colours.
- **Not supported:** tables, charts and pictures, and per-character baseline shifts.
