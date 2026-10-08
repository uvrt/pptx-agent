# Inputs of the full end-to-end trial and the graphics spike

| File | SHA-256 |
| --- | --- |
| `company-template.potx` | `b6be5df5cd7eaed8cee3b796769358f94a078a6a8a5508cb828692d9efc0341b` |
| `pilot-retrospective.pptx` | `ea2c5453e264914b79f4286ed5bd13b1e8426687424b16ff42f75b438b19b03f` |
| `halden-proposal-draft.pptx` | `9b2639c49317af6b6820c2bfd2dfaeb41c37772a116bda318b9a33be116446fe` |
| `business-review-2026.pptx` | `a7718991dbfe36bbcaac5a41dac40dda58d15e5e7d94af1c43d0375607c52952` |
| `quarterly-revenue-2026.csv` | `977e11638b60fdeb12c706935d93e440aa1f9c25044d81e3959316c6a9175b78` |
| `wms-vendor-selection.pptx` | `7e8d7582c1922d5baba257d9265bb66503fe5fb9c5c1a82b4773ff29d979d8ac` |
| `wms-vendor-scores.csv` | `1bb99ba07ed6c04a0db06307b77ed25f9c18fb13ae27cfd92689157e3f4fee63` |
| `customer-review.pptx` | `697f2724d7aa28ed6f745072a7f7559577d71f856fb259e4f6119b939019ca81` |
| `new-logo.png` | `78522c7cc2628493168f347e1c24fc285c051d6a9c12717f632db23a70332b10` |

**Written by pptx-agent** through its own API, for the full end-to-end trial (twelve tasks,
two runs each, agents using only the public API and its documentation), and committed as
the trial used them:

- `company-template.potx` is task P1's template: `Document.new(template=...)` from
  `tests/fixtures/real-basic-theme.pptx` (a Google Slides export; its eleven layouts keep
  their own names, `TITLE`, `TITLE_ONLY`, `BIG_NUMBER`...), titled "Harbour template",
  author "PMO", saved with `save_as_template`. It reproduces two findings: a template opened
  with `Document.open` and saved as `.pptx` kept the template content type, which PowerPoint
  refuses; and `insert_outline` put a drafted table over the `TITLE_ONLY` layout's title,
  which sits lower than the layout's body.
- `pilot-retrospective.pptx` is task P3's deck: `Document.new()` and `insert_outline` of a
  four-slide outline whose slide 3, "Lessons learned", has five topics of three points each
  in one body placeholder -- more than fits. The Office layout's body is `normAutofit`
  with no stored `fontScale`, so PowerPoint draws it at full size and it overflows.
- `halden-proposal-draft.pptx` is the draft deck of trial 2's design tasks (p7, p8) and the
  graphics spike's (p7, p8, o1, m1): `Document.new()` with an invented "Meridian" theme
  (Arial, its own colours) and three slides, a cover and two title-only slides. The spike
  derived o1's and m1's decks from it by retitling one slide (`tests/golden_inputs.py`).

- `business-review-2026.pptx` and `quarterly-revenue-2026.csv` are T4's chart task p9:
  `Document.new()` with an invented theme (six accents, Arial), `insert_outline` of a
  four-slide review, and the year's revenue by region and quarter, in EUR million.
- `wms-vendor-selection.pptx` and `wms-vendor-scores.csv` are the radar-chart task p13:
  `Document.new()` with an invented theme (six accents, Arial), `insert_outline` of a
  four-slide vendor selection, and an evaluation panel's scores (1 to 5) for three vendors
  on six criteria.
- `customer-review.pptx` and `new-logo.png` are the rebrand task p5's deck and new logo,
  written by `p5_draw` in `tests/golden_inputs.py` (`python tests/golden_inputs.py --draw
  <root> p5-rebrand`, Pillow 12.3): `Document.new()`, `insert_outline` of a four-slide
  customer review, then the old brand -- a red title colour, bands, an amber callout, a
  red table header and an old logo on every slide, and a team photo. The logos and the
  photo are flat shapes Pillow draws (the logo's word in Pillow's default font, as
  pixels; no font file is carried). The golden replayed whatever the installed Pillow
  drew until they were committed; these are the bytes its transcript was recorded with.

The text is invented. Licence: this repository's, MIT -- with one exception.
`company-template.potx` derives from `tests/fixtures/real-basic-theme.pptx`, a Google
Slides export from pptx-glimpse's fixtures (MIT licensed, third-party; see
[`../../README.md`](../../README.md)), and carries that deck's theme, masters and layouts.
It also embeds the deck's fonts as `ppt/fonts/*.fntdata`: Lato and Raleway (regular,
bold, italic, bold italic), which are licensed under the SIL Open Font License 1.1, not
under MIT. The other files embed no font.

They sit one level below `fixtures/`, so the corpus the other suites hold to every fixture
(`tests/conftest.py`, `fixture_paths`) does not take them in.
