# Task: build the "phased approach" slide for the Halden proposal

You are an analyst at Meridian Partners, a strategy consultancy. The engagement manager has sent you this brief.

> We go to Halden Group on Wednesday with the procurement proposal. The draft deck is in `input/halden-proposal-draft.pptx`. It is in our template: a cover and two content slides that only have titles. Slide 2 needs to become our standard "phased approach" page. You know the format: chevrons across the top, the activities in a column under each phase, the deliverables in a box at the bottom of each column. Please build it from shapes, not as a table or SmartArt, so we can still tweak it in the meeting. Use the content below exactly as written, because the partner has already agreed the wording with the client lead. The layout is up to you, but it has to look like something we would put in front of a CEO. Leave slide 3 alone; someone else is doing the plan.

## Client situation (for context; it does not go on the slide)
Halden Group is a Nordic building-materials manufacturer with EUR 2.1 billion revenue and six plants. Its EBITDA margin has fallen from 14% to 9% in two years, mostly because of input costs. Procurement is decentralised: each plant buys for itself, and there is no spend visibility across the group. The CEO wants a 16-week programme that delivers signed savings, not just a diagnostic, and leaves behind a procurement function that can keep them.

## Content for slide 2

**Action title** (use this wording, or a sharper version of it that is still one full sentence of at least 8 words):
A 16-week programme takes Halden from a spend baseline to signed savings and a procurement function that sustains them

**Phases, in this order** (each chevron shows the phase name and its weeks):

| # | Phase | Weeks |
|---|---|---|
| 1 | Diagnose | Weeks 1–3 |
| 2 | Design | Weeks 4–7 |
| 3 | Negotiate | Weeks 8–13 |
| 4 | Embed | Weeks 14–16 |

**Key activities** (bullets in the column under each phase):
- Diagnose:
  - Build the group spend cube from ERP data for all six plants
  - Interview 25 category owners and plant managers
  - Benchmark prices for the top 15 categories
  - Size the savings potential by category
- Design:
  - Prioritise eight wave-1 categories
  - Write category strategies and negotiation plans
  - Design the target procurement operating model
- Negotiate:
  - Run RFQs and supplier negotiations for wave 1
  - Track savings in a weekly war room
  - Prepare the wave-2 categories
  - Coach category managers on the job
- Embed:
  - Hand over to the new procurement organisation
  - Set up procurement KPIs and governance
  - Agree the wave-2 roadmap

**Deliverables** (in a visually distinct box at the foot of each column, under a "Deliverables" label):
- Diagnose: Spend cube and cost baseline; Savings hypothesis by category
- Design: Eight category strategies; Target operating model; Wave plan approved by the SteerCo
- Negotiate: Signed supplier agreements; Savings tracker validated by Finance
- Embed: Savings sign-off; Procurement KPI dashboard; 12-month roadmap

**Source line** (at the foot of the slide, small):
Source: Halden Group ERP spend data FY2025; Meridian Partners procurement benchmarks. Timeline indicative, subject to data access in week 1.

## Hard requirements (these will be checked)
1. Output: `work/halden-proposal-approach.pptx`, still 3 slides. Slides 1 and 3 are unchanged. Slide 2 keeps its title placeholder, which now holds the action title. The title fits its placeholder.
2. The four phases are **four chevron shapes** (preset `chevron`, or `homePlate` for the first one). They sit left to right in phase order, in one row with the same top, the same size and even spacing. Each chevron's own text is the phase name and its weeks, for example "Diagnose" and "Weeks 1–3" (two lines in the one shape).
3. Each phase's activities are in **one text shape** under its chevron. All four activity columns have the same width and the same gaps between them, and each column is centred under its chevron (within 0.1 inch).
4. Each phase's deliverables are in **one filled shape** under its activity column, the same width as the column. All four deliverable boxes are aligned in one row with the same top and height. The word "Deliverables" appears as a label, either inside each box or once as a row label.
5. All the text above is on the slide exactly as written. Bullets may be real bullets or the text alone.
6. **Theme colours only:** every fill, outline and text colour you set is a theme colour (`accent1`...`accent6`, `tx1`, `tx2`, `bg1`, `bg2`, with tints or shades allowed). Do not use hard-coded RGB.
7. **Legible:** activity text is at least 10 pt, and nothing on the slide is smaller than 8 pt.
8. **Nothing overflows or collides:** `deck.overflows()` is empty, and no two shapes that hold text overlap. The one exception is neighbouring chevrons, which may nest into each other.
9. **Inside the margins:** everything you add lies between 0.4 in and 12.933 in horizontally, and between 1.4 in (below the title rule) and 7.0 in (above the footer band, which holds the "MERIDIAN PARTNERS" wordmark) vertically. The slide is 13.333 × 7.5 in.
10. `deck.validate()` stays clean, and PowerPoint must open the file without a repair prompt.
