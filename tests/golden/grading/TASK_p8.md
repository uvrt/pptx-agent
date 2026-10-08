# Task: build the engagement plan (Gantt) slide for the Halden proposal

You are an analyst at Meridian Partners, a strategy consultancy. The engagement manager has sent you this brief.

> We go to Halden Group on Wednesday with the procurement proposal. The draft deck is in `input/halden-proposal-draft.pptx`. It is in our template: a cover and two content slides that only have titles. Slide 3 needs to become the engagement plan: a Gantt-style timeline from November to March, with the workstreams as rows, bars for the activities, the SteerCo milestones and the board update. Build it from shapes, not a table, an image or a chart, so we can move bars in the meeting. The phases must match the approach page (Diagnose, Design, Negotiate, Embed). The dates below are agreed with the client; get them exactly right, because their PMO will check. Leave slide 2 alone; someone else is doing it. The layout is up to you, but it has to look like something we would put in front of a CEO.

## Client situation (for context; it does not go on the slide)
Halden Group is a Nordic building-materials manufacturer with EUR 2.1 billion revenue and six plants. We are proposing a 16-week procurement programme. It starts on Monday 2 November 2026 and pauses for the holidays from 21 December 2026 to 1 January 2027 (programme weeks do not count the break). It ends on Friday 5 March 2027.

## Content for slide 3

**Action title** (use this wording, or a sharper version of it that is still one full sentence of at least 8 words):
Signed wave-1 savings by mid-February, with four SteerCo decisions keeping the programme on track

**Time axis:** five month columns with these exact header texts: `Nov 2026`, `Dec 2026`, `Jan 2027`, `Feb 2027`, `Mar 2027`. Week ticks or week numbers under the months are optional.

**Phase bands** (across the top, above the workstreams; each band spans exactly its dates and shows the phase name):

| Phase | From | To (inclusive) |
|---|---|---|
| Diagnose | Mon 2 Nov 2026 | Fri 20 Nov 2026 |
| Design | Mon 23 Nov 2026 | Fri 18 Dec 2026 |
| Negotiate | Mon 4 Jan 2027 | Fri 12 Feb 2027 |
| Embed | Mon 15 Feb 2027 | Fri 5 Mar 2027 |

**Workstreams (rows, in this order) and their activity bars:**

| Workstream | Activity (bar text) | From | To (inclusive) |
|---|---|---|---|
| Spend & baseline | Spend cube | 2 Nov 2026 | 20 Nov 2026 |
| Spend & baseline | Savings tracking | 4 Jan 2027 | 5 Mar 2027 |
| Category strategy | Category deep dives | 16 Nov 2026 | 18 Dec 2026 |
| Category strategy | Wave-2 preparation | 25 Jan 2027 | 26 Feb 2027 |
| Negotiations | Wave-1 RFQs | 4 Jan 2027 | 22 Jan 2027 |
| Negotiations | Supplier negotiations | 25 Jan 2027 | 12 Feb 2027 |
| Operating model | Target model design | 30 Nov 2026 | 18 Dec 2026 |
| Operating model | Transition and hand-over | 15 Feb 2027 | 5 Mar 2027 |
| Change & capability | Category manager coaching | 11 Jan 2027 | 26 Feb 2027 |

**Milestones** (a diamond or triangle marker on the date, with its label next to it):

| Milestone label | Date |
|---|---|
| Kick-off | 2 Nov 2026 |
| SteerCo 1: savings hypothesis | 20 Nov 2026 |
| SteerCo 2: wave plan approved | 18 Dec 2026 |
| SteerCo 3: first contracts signed | 12 Feb 2027 |
| Final SteerCo | 5 Mar 2027 |

**Marker line:** a vertical dashed line on **22 Jan 2027**, running down through the workstream rows, labelled `Board update`.

**Footnote** (at the foot of the slide, small):
Note: programme weeks exclude the holiday break, 21 Dec 2026 to 1 Jan 2027. Dates as agreed with the Halden PMO on 9 October 2026.

## Hard requirements (these will be checked)
1. Output: `work/halden-proposal-plan.pptx`, still 3 slides. Slides 1 and 2 are unchanged. Slide 3 keeps its title placeholder, which now holds the action title. The title fits its placeholder.
2. **The date scale:** the five month headers are text shapes in one row, left to right, edge to edge, with no gaps. A date's position inside its month is proportional to the day: day 1 starts at the month's left edge and the last day ends at its right edge. Month columns may be equal-width or proportional to their number of days. Positions are checked against the month headers you draw.
3. **Bars:** every activity is **one rectangle or rounded rectangle whose own text is the activity name** (exactly as in the table). Its left edge is at the start of its *From* day and its right edge at the end of its *To* day, each within 0.1 inch. All bars have the same height. The bars of a workstream sit in that workstream's row, and rows are evenly spaced.
4. **Workstream labels:** each workstream name is a text shape to the left of the plot area, vertically centred on its row (within 0.1 inch of the row's bars).
5. **Phase bands:** one shape per phase whose text is the phase name. Its left and right edges match its dates within 0.1 inch (a chevron or pentagon's bounding box counts). All four bands are in one row above the bars.
6. **Milestones:** five marker shapes (preset `diamond`, `flowChartDecision`, `triangle` or `flowChartMerge`), each centred on its date within 0.1 inch, with its label text on the slide.
7. **Marker line:** a vertical line (a connector or line shape, dashed) at 22 Jan 2027 within 0.1 inch, spanning at least from the first workstream row to the last, with the label `Board update`.
8. **Colour:** one colour per workstream, or one per phase, applied consistently. If the colours carry meaning, add a legend. Use **theme colours only** (`accent1`...`accent6`, `tx1`, `tx2`, `bg1`, `bg2`, tints and shades allowed), with no hard-coded RGB.
9. **Legible:** nothing is smaller than 8 pt, and bar texts and workstream labels are at least 9 pt.
10. **Nothing overflows or collides:** `deck.overflows()` is empty, and no two shapes that hold text overlap.
11. **Inside the margins:** everything you add lies between 0.4 in and 12.933 in horizontally, and between 1.4 in (below the title rule) and 7.0 in (above the footer band with the "MERIDIAN PARTNERS" wordmark) vertically. The slide is 13.333 × 7.5 in.
12. `deck.validate()` stays clean, and PowerPoint must open the file without a repair prompt.
