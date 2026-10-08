# Task: build the programme governance (org chart) slide for the Halden proposal

You are an analyst at Meridian Partners, a strategy consultancy. The engagement manager has sent you this brief.

> We go to Halden Group on Wednesday with the procurement proposal. The draft deck is in `input/halden-proposal-draft.pptx`. It is in our template: a cover and two content slides that only have titles. Slide 2 needs to become the programme governance page. It's a classic org chart: the Steering Committee at the top, the Programme Lead under it with the PMO beside her, and the four workstreams below, each with its lead and its roles. Build it from shapes and real connectors, not SmartArt, a table or a picture, so we can move boxes in the meeting and the lines follow. Use the names and wording below exactly; the client lead has confirmed them. The layout is up to you, but it has to look like something we would put in front of a CEO. Leave slide 3 alone.

## Client situation (for context; it does not go on the slide)
Halden Group is a Nordic building-materials manufacturer with EUR 2.1 billion revenue and six plants. Procurement is decentralised. The CEO wants clear accountability: one steering committee, one programme lead, and named workstream leads from the business.

## Content for slide 2

**Action title** (use this wording, or a sharper version of it that is still one full sentence of at least 8 words):
A lean governance structure gives the CEO one steering committee and four accountable workstream leads

**Units** (each is one box: its name and its detail line in the same shape):

| Unit | Name (as written) | Detail (as written) |
|---|---|---|
| Steering Committee | Steering Committee | Group CEO (chair), CFO, COO and Meridian partner; meets monthly |
| Programme Lead | Programme Lead | Ingrid Solberg, Group Procurement Director |
| PMO | PMO | Savings tracking, risks and reporting; meets weekly |

**Workstreams** (left to right in this order; each workstream box holds the workstream name and its lead line):

| # | Workstream | Lead line | Roles (top to bottom, each its own box) |
|---|---|---|---|
| 1 | Spend & Baseline | Lead: Erik Dahl, Finance | Spend data analyst; Plant controllers (6) |
| 2 | Category Strategy | Lead: Maria Berg, Procurement | Category managers (8); Technical specialists; Meridian category experts |
| 3 | Negotiations | Lead: Jonas Lind, Procurement | Negotiation team; Legal counsel |
| 4 | Operating Model & Change | Lead: Sofie Nyberg, HR | HR business partner; Change manager; Training coordinator |

**Footnote** (at the foot of the slide, small):
Note: names as proposed by Halden; roles to be confirmed at SteerCo 1.

## Hard requirements (these will be checked)
1. Output: `work/halden-proposal-governance.pptx`, still 3 slides. Slides 1 and 3 are unchanged. Slide 2 keeps its title placeholder, which now holds the action title. The title fits its placeholder.
2. **Boxes:** the Steering Committee, the Programme Lead, the PMO and each workstream are each **one shape** holding its name and its detail (or lead) line. Each role is **its own shape** holding exactly the role text.
3. **Hierarchy:** the Steering Committee box is above the Programme Lead box, with the Programme Lead box's centre within the Steering Committee box's width. The PMO box is **beside** the Programme Lead box: to its left or right, not overlapping it, its vertical centre within the Programme Lead box's height. The four workstream boxes are in **one row** below the Programme Lead (same top, same width and height, equal gaps, in the order above). Each workstream's roles are **below its workstream box**, within its column (each role box lies horizontally within the workstream box's left and right edges, ±0.05 in), in the order given, top to bottom.
4. **Connectors:** real PowerPoint connectors **attached at both ends** (they follow the boxes when moved) join: the Steering Committee and the Programme Lead; the Programme Lead and the PMO; the Programme Lead and each of the four workstreams. Connectors to the roles are optional.
5. All the text above is on the slide exactly as written.
6. **Theme colours only:** every fill, outline and text colour you set is a theme colour (`accent1`...`accent6`, `tx1`, `tx2`, `bg1`, `bg2`, `dk1`, `dk2`, `lt1`, `lt2`, with tints or shades allowed). Do not use hard-coded RGB.
7. **Legible:** names and detail lines are at least 10 pt, role texts at least 9 pt, and nothing on the slide is smaller than 8 pt.
8. **Nothing overflows or collides:** `deck.overflows()` is empty (this includes text over text and lines through text).
9. **Inside the margins:** everything you add lies between 0.4 in and 12.933 in horizontally, and between 1.4 in (below the title rule) and 7.0 in (above the footer band, which holds the "MERIDIAN PARTNERS" wordmark) vertically. The slide is 13.333 × 7.5 in.
10. `deck.validate()` stays clean, and PowerPoint must open the file without a repair prompt.
