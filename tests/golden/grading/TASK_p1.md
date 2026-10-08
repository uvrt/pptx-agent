# Task: build a deck from an outline, on the company template

`input/outline.txt` is the outline of an 8-slide steering-committee deck, with speaker notes for every slide. `input/company-template.potx` is the company's PowerPoint template (its layouts have their own names).

Build the deck on the template:

- Exactly 8 slides, in the outline's order, made from the **template's own layouts**:
  - slide 1 from the layout `TITLE`;
  - slide 4 (the big number) from `BIG_NUMBER`: the number "38%" big, with its sentence under it;
  - slide 5 (two columns) from `TITLE_AND_TWO_COLUMNS`, one column per side;
  - slide 6 (the table) from `TITLE_ONLY`, with the figures as a real table;
  - every other slide from `TITLE_AND_BODY`, the bullets as real bullets (keep the outline's sub-bullets as a second level).
- Each slide's title and text as in the outline, word for word.
- Each slide's **speaker notes** as given in the outline.
- Nothing may overflow its placeholder or run off the slide, and nothing may be left over from the template that is not part of the outline (no empty or leftover placeholders showing prompt text).

Save the result as `work/project-harbour.pptx`.
