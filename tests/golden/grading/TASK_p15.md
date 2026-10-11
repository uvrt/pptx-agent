# Task: chart the maturity gap and the recommendations

`input/maturity-assessment.pptx` is a four-slide maturity assessment of warehouse operations. The panel's scores are in `input/maturity-scores.csv` (one row per capability: the current score and the target, on a scale from 1 to 5), and its recommendations in `input/recommendations.csv` (an id, a name, and the effort, impact and priority each was rated).

Add two slides after "Approach" ("Next steps" moves to slide 6):

- **Slide 4, the maturity gap:** the current scores as columns in the deck's second text colour (`tx2`), with the lowest-scoring capability highlighted in an accent colour; the targets as diamond markers on the same axes, with no line between them. The value axis runs from 0 to 5 in steps of 1, with no gridlines; all the chart's text is 10 pt. The capabilities in the CSV's order, the CSV's values exactly, and Edit Data agrees with what the chart draws.
- **Slide 5, effort against impact:** a bubble chart with each recommendation at its effort (x) and impact (y), sized by its priority, and labelled with its id (R1 to R6) beside the bubble. Both axes from 0 to 5, titled "Effort" and "Impact".
- Each slide has **a one-line takeaway title** (a finding, not a label).
- The charts fill the content area below their titles; nothing overlaps or leaves the slide.
- The other slides do not change.

Save the result as `work/maturity-assessment.pptx`.
