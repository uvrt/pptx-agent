# Task: a chart slide and a table slide in the template's own style

`input/business-review-2026.pptx` is a four-slide business review on the company template. The finance team has sent the year's revenue in `input/quarterly-revenue-2026.csv`: one row per quarter, one column per region, in EUR million.

Add two slides after "Results by region" (they become slides 4 and 5; "Next steps" moves to slide 6):

- **Slide 4, on the "Title and Content" layout:** a clustered column chart of the CSV (the quarters as categories, in order, one series per region with the CSV's values exactly) in the slide's content area below the title, and a one-line takeaway title. Nothing may be left that PowerPoint shows as "Click to add text": no empty placeholder on the slide.
- **Slide 5:** a table of each region's full-year revenue (the sum of its four quarters, one decimal, e.g. `61.4`) with a header row `Region`, `Revenue 2026 (EUR m)`, one row per region in the CSV's order, in the **template's own table style** (not one of PowerPoint's built-in styles), with its **header row on and banded rows off**. A one-line title.
- Nothing overlaps or leaves the slide, and the other slides do not change.

Save the result as `work/business-review-2026.pptx`.
