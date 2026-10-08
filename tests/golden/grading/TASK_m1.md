# Task: build the prioritisation matrix (2×2) slide for the Halden proposal

You are an analyst at Meridian Partners, a strategy consultancy. The engagement manager has sent you this brief.

> We go to Halden Group on Wednesday with the procurement proposal. The draft deck is in `input/halden-proposal-draft.pptx`. It is in our template: a cover and two content slides that only have titles. Slide 3 needs to become the prioritisation matrix: our usual 2×2, ease of implementation across, savings potential up, the four quadrants named, and the ten wave-1 candidate categories plotted as bubbles at their scores, coloured by spend type with a legend. Build it from shapes, not a chart, an image or SmartArt, so we can drag bubbles in the meeting. The scores below come from the diagnostic; plot them exactly, because the CFO will compare them with the backup. The layout is up to you, but it has to look like something we would put in front of a CEO. Leave slide 2 alone.

## Client situation (for context; it does not go on the slide)
Halden Group is a Nordic building-materials manufacturer with EUR 2.1 billion revenue and six plants. The diagnostic scored the top spend categories on ease of implementation and annual savings potential, each from 0 to 10.

## Content for slide 3

**Action title** (use this wording, or a sharper version of it that is still one full sentence of at least 8 words):
Four quick wins combine high savings with easy implementation and should lead wave 1

**Axes:** a horizontal axis and a vertical axis, each a straight line running exactly from score 0 to score 10 (arrowheads allowed), meeting at score (0, 0) at the bottom left. Their titles (as written):
- horizontal: `Ease of implementation (0 = hard, 10 = easy)`
- vertical: `Savings potential (0 = low, 10 = high)` (it may be rotated to run up the axis)

Both axes use the same score scale along their own length: score s sits at s/10 of the axis length from the origin. The axes may have different lengths. Tick labels are optional.

**Quadrants** (split at score 5 on both axes; each label lies inside its quadrant):

| Quadrant | Label (as written) |
|---|---|
| top right (easy, high savings) | Quick wins |
| top left (hard, high savings) | Strategic bets |
| bottom right (easy, low savings) | Fill-ins |
| bottom left (hard, low savings) | Deprioritise |

**Initiatives** (one bubble each, centred on its scores; its label as written, in or next to the bubble):

| Initiative (label) | Ease | Savings | Spend type |
|---|---|---|---|
| Cement & binders | 7.5 | 8.5 | Direct materials |
| Steel reinforcement | 3.5 | 9.0 | Direct materials |
| Packaging | 8.0 | 6.5 | Direct materials |
| Energy contracts | 2.5 | 7.0 | Indirect |
| MRO spares | 8.5 | 3.5 | Indirect |
| IT & telecoms | 6.0 | 2.0 | Indirect |
| Professional services | 9.0 | 5.5 | Indirect |
| Inbound freight | 6.5 | 7.5 | Logistics |
| Warehousing | 2.0 | 3.0 | Logistics |
| Fleet leasing | 4.0 | 1.5 | Logistics |

**Legend:** the three spend types, `Direct materials`, `Indirect` and `Logistics`, each with a colour swatch.

**Source line** (at the foot of the slide, small):
Source: Meridian Partners diagnostic, scores agreed with category owners in week 3.

## Hard requirements (these will be checked)
1. Output: `work/halden-proposal-matrix.pptx`, still 3 slides. Slides 1 and 2 are unchanged. Slide 3 keeps its title placeholder, which now holds the action title. The title fits its placeholder.
2. **Axes:** one vertical and one horizontal straight line (a connector or line shape) meeting at the origin, each spanning exactly score 0 to score 10 (within 0.05 in). Their titles are on the slide as written.
3. **Bubbles:** every initiative is **one oval** (preset `ellipse`), all ten the same size (at least 0.2 in across). Each is centred on its scores **within 0.1 in**, measured against the axes you draw.
4. **Labels:** each initiative's label (exactly as written) is either the bubble's own text or a text shape whose centre is within 1.0 in of the bubble's centre.
5. **Quadrants:** the four quadrant labels are on the slide as written, each lying wholly inside its quadrant of the plot area.
6. **Colour by spend type:** all bubbles of a spend type share one fill, and the three types have three different fills. The legend shows each type's name next to (within 0.3 in) a swatch shape filled like that type's bubbles.
7. All the text above is on the slide exactly as written.
8. **Theme colours only:** every fill, outline and text colour you set is a theme colour (`accent1`...`accent6`, `tx1`, `tx2`, `bg1`, `bg2`, `dk1`, `dk2`, `lt1`, `lt2`, with tints or shades allowed). Do not use hard-coded RGB.
9. **Legible:** bubble labels and quadrant labels are at least 9 pt, and nothing on the slide is smaller than 8 pt.
10. **Nothing overflows or collides:** `deck.overflows()` is empty (this includes text over text and lines through text).
11. **Inside the margins:** everything you add lies between 0.4 in and 12.933 in horizontally, and between 1.4 in (below the title rule) and 7.0 in (above the footer band with the "MERIDIAN PARTNERS" wordmark) vertically. The slide is 13.333 × 7.5 in.
12. `deck.validate()` stays clean, and PowerPoint must open the file without a repair prompt.
