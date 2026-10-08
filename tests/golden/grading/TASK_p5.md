# Task: rebrand a deck

`input/customer-review.pptx` was made with the old brand. The new brand's colours are already the deck's **theme colours**, but the deck uses the old colours as hard-coded RGB values, and it shows the old logo.

1. **Colours.** Replace every hard-coded old brand colour with the theme colour it now corresponds to, wherever it is used (shape fills, outlines, text colour, table cells):

   | old (hard-coded) | new (theme colour) |
   |---|---|
   | `#8B1E3F` | Accent 1 |
   | `#F2A541` | Accent 2 |
   | `#3C3C3B` | Text 1 (dark 1) |

   The result must use the theme colours themselves (so a later theme change carries through), not their RGB values.

2. **Logo.** The old logo (the maroon square, top right of every slide) becomes `input/new-logo.png`. Keep each logo's position (top right corner) and height; let its width follow the new logo's proportions so it is not distorted, and keep it on the slide. The team photo on the last slide is not a logo and stays.

Change nothing else. Save the result as `work/customer-review.pptx`.
