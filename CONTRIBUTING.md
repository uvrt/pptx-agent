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
