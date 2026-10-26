# Contributing

## Running the tests

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main" \
            "ooxml-edit @ git+https://github.com/uvrt/ooxml-edit@main" \
            "pptx2svg[png] @ git+https://github.com/uvrt/pptx2svg@main"
pip install -e .[dev,png]       # or, with sibling checkouts: pip install -e . -e ../pptx2svg
python -m pytest -q
```

Over two thousand tests run on the decks in `tests/fixtures/` (PowerPoint, Google Slides,
Keynote and other generators) and on decks the tests build. Some tests are **local-only**
and skip cleanly elsewhere, including on CI:

- **PowerPoint oracle** (`python -m pytest -m oracle`, opt-in): drives the real
  application, the only authoritative answer to "will it open?" and "does it fit?". It
  needs macOS with PowerPoint, the export script from a pptx2svg checkout next to this one
  (or `PPTX2SVG_ORACLE_SCRIPT`), inputs and outputs under your home directory, and the
  Automation and Accessibility permissions; see [ROADMAP.md](ROADMAP.md).
- **Provider tests** (`-m provider`): call a model provider's API; skipped without
  `ANTHROPIC_API_KEY`.
- A few tests read local trial or spike output and skip where it is absent.

## What never goes into the repository

- **No Microsoft font file**, in any form, and no other font data beyond what a
  documented third-party fixture already carries.
- **No Office output**: no PDF or raster PowerPoint exported.
- **No third-party document** without a licence that allows redistribution, recorded in
  [tests/fixtures/README.md](tests/fixtures/README.md) or a `PROVENANCE.md` beside it.
- No API keys or other secrets.

## Pull requests

Open pull requests against `main`. CI runs the suite on Linux, macOS and Windows, Python
3.10 to 3.15, with the siblings installed from their `main` branches.

## Versions

`pyproject.toml`'s version is the one [CHANGELOG.md](CHANGELOG.md)'s top version section
names, and `pptx_agent.__version__` the same; CI's `build` job checks it
(`tools/check_version.py`). Each version that reaches `main` is tagged `v<version>` on that
commit by the tag workflow (`.github/workflows/tag.yml`). Raise the version with its
CHANGELOG section; changes since the last one go under an `Unreleased` heading above it.

## Dependencies

The siblings are not on PyPI yet: install them first, from sibling checkouts
(`pip install -e ../ooxml-common -e ../ooxml-edit -e ../pptx2svg`) or from their
repositories, as CI does. `pyproject.toml` names versions, never URLs, since PyPI refuses
a distribution whose dependencies name one. Each floor is the first version with what
this package uses; what each earlier version added is in that package's CHANGELOG.

- `ooxml-edit>=0.68.0`: the format-neutral editing core (the OPC package, ordered
  insertion, undo, id stamping, the tool layer, and in `ooxml_edit.charts` the chart,
  embedded-workbook and SmartArt editing), extracted from this repository with its
  history. 0.62.0 gives public names to what was imported private (`opc.part_for_rels`,
  `xml.child_ranks`, `charts.chart.series_of`; review, N2); 0.63.0 the one lookup for
  `edit_chart`'s series and category and `edit_smartart`'s node (`charts.actions.resolve_*`;
  review, N4); 0.67.0 the handlers' work itself, the same in a Word document
  (`charts.actions.apply_data_action`, `add_plan`, `chart_json`..., `Chart.number_formats`, and
  the table `charts.cases` the tests replay; review, S10); 0.68.0 an SVG picture and its
  recolouring, shared with docx-agent (`ooxml_edit.pictures`; round 17, P6).
- `ooxml-common>=0.31.0`: the reading core pptx2svg is built on, a dependency rendering
  or not: the package kinds (validate, saving a template as a deck) and the text
  measurement come from it. 0.27.0 carries `drawingml.text_chain`, the inherited text
  chain shared with pptx2svg (`edit/effective.py`); 0.28.0 `fonts.embedded.forget_decoded`,
  which the parallel-render test empties the decoded-face cache with (review, N2); 0.31.0
  `drawingml.table_styles` and `drawingml.theme`, the built-in table styles and the Office
  theme a new deck is written with (review, S3).
- `lxml>=4.9`: the XML tree every edit is made in.
- `markdown-it-py>=3.0`, `mdit-py-plugins>=0.4`: the Markdown outline in docx-agent's
  dialect (CommonMark with GFM tables and strikethrough; the plugins carry docx-agent's
  footnotes, which a slide has no place for, so the two stay one dialect).
- `pptx2svg>=0.2.0` (`render`, `png` and `dev` extras): rendering, optional so the edit
  layer can be used headless. 0.2.0 is pptx2svg's first tagged version and carries
  everything used here (`table_row_heights`, `convert_pptx_to_agent_svg`,
  `ConvertOptions.templates`), so nothing checks for an older one.
