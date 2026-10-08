# pptx-agent

[![CI](https://github.com/uvrt/pptx-agent/actions/workflows/ci.yml/badge.svg)](https://github.com/uvrt/pptx-agent/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

An AI-editable PowerPoint layer for Python: inspect a deck, change it through a semantic API,
render it back to an image, repeat.

The agent never writes XML or SVG. It looks at pixels, understands structure through inspection,
mutates a semantic model, and looks again.

To give a model the deck as tool calls (Claude, or OpenAI's Responses and Chat Completions
APIs): [using the tools with a model provider](src/pptx_agent/tools/README.md).

## Install

Not on PyPI yet. Python 3.10+. The siblings install from git:

```bash
pip install "ooxml-common @ git+https://github.com/uvrt/ooxml-common@main" \
            "ooxml-edit @ git+https://github.com/uvrt/ooxml-edit@main" \
            "pptx2svg[png] @ git+https://github.com/uvrt/pptx2svg@main" \
            "pptx2svg-fonts @ git+https://github.com/uvrt/pptx2svg@main#subdirectory=packages/pptx2svg-fonts"
pip install "pptx-agent @ git+https://github.com/uvrt/pptx-agent@main"
```

Editing needs only lxml, `ooxml-common` and `ooxml-edit`; rendering is delegated to
[`pptx2svg`](https://github.com/uvrt/pptx2svg) (the `render` and `png` extras).
**`pptx2svg-fonts` is the fonts the PNGs are drawn with** (about 11 MB): Carlito, Arimo,
Tinos and Cousine, metric-compatible stand-ins for Calibri, Arial, Times New Roman and
Courier New; Caladea for Cambria; Noto Sans JP for Japanese and Chinese text; Lato and
Raleway. They are under the SIL Open Font License 1.1, with their licence texts in the
package; its code is MIT; no Microsoft font is in it. Without it a render draws with the
machine's own fonts, and text none of them can draw -- every CJK character, on a server
with no CJK font -- is left out of the image: `render` then says so (`missing_glyphs`),
and pptx2svg warns.

## Example

```python
from pptx_agent import Document

deck = Document.open("deck.pptx")
slide = deck.slides[0]

for shape in slide.shapes:
    print(shape.id, shape.kind, shape.name, shape.left, shape.top)

slide.shape("256.3").set_text("Q3 results")
slide.shape("256.5").move_by(dx=914400)      # one inch right

png = slide.render_png(width=1280)           # what the agent looks at next
deck.save("edited.pptx")
```

Ids are stable (`"<sldId>.<cNvPr id>"`, never positions), and the rendered SVG tags every
shape with the same id, so an agent can point at what it sees. Open and save reproduces
every part of a deck byte for byte.

## What is supported

- **Library:** open and save (`.pptx`, `.potx`), text with formatting, notes, fills,
  outlines and theme colours, tables, charts (with their embedded workbooks), SmartArt
  nodes, shapes, pictures, connectors, groups, comments, slides and layouts, Markdown
  outlines in and out, fit and collision checks, layout helpers, design facts, undo and
  redo, SVG and PNG rendering. The full table: [docs/FEATURES.md](docs/FEATURES.md).
- **Agent tools** (`pptx_agent.tools`): the deck through tool calls for a model, on
  `ooxml_edit.tools`. What they support and do not:
  [SUPPORTED.md](src/pptx_agent/tools/SUPPORTED.md); guidance for the application's
  thinking layer: [GUIDANCE.md](src/pptx_agent/tools/GUIDANCE.md).
- Lengths are EMU everywhere (914,400 per inch), font sizes points. What is measured,
  approximated or still to come is in [ROADMAP.md](ROADMAP.md).

## Status

Pre-release (0.0.1). The API is exercised end to end by golden transcripts of trial
tasks; it may still change. PowerPoint oracle tests are local-only (macOS with
PowerPoint) and skip elsewhere, including CI.

## Documentation

- [docs/USAGE.md](docs/USAGE.md) -- how to do each thing, with examples the tests run
- [docs/FEATURES.md](docs/FEATURES.md) -- what works today
- [docs/DESIGN.md](docs/DESIGN.md) -- the three layers, live XML views, stable ids, prior art
- [docs/PROFILE.md](docs/PROFILE.md) -- the SVG profile
- [tools/README.md](src/pptx_agent/tools/README.md), [SUPPORTED.md](src/pptx_agent/tools/SUPPORTED.md), [GUIDANCE.md](src/pptx_agent/tools/GUIDANCE.md) -- the agent tools
- [ROADMAP.md](ROADMAP.md) -- the working plan and what has been measured
- [CHANGELOG.md](CHANGELOG.md), [CONTRIBUTING.md](CONTRIBUTING.md) (running the tests)

## Family

- [pptx2svg](https://github.com/uvrt/pptx2svg) -- renders PowerPoint (`.pptx`) slides to SVG and PNG.
- [docx2svg](https://github.com/uvrt/docx2svg) -- renders Word (`.docx`) documents to SVG, page by page.
- [ooxml-common](https://github.com/uvrt/ooxml-common) -- the format-neutral reading, DrawingML, fonts and text metrics both renderers share.
- [ooxml-edit](https://github.com/uvrt/ooxml-edit) -- lossless, undoable editing of OOXML packages, shared by both agent layers.
- [pptx-agent](https://github.com/uvrt/pptx-agent) (this repo) -- an AI-editable PowerPoint layer: inspect, edit, re-render.
- [docx-agent](https://github.com/uvrt/docx-agent) -- an AI-editable Word layer: inspect, edit (optionally as tracked changes), re-render.

## License

MIT, see [LICENSE](LICENSE). Six sample decks in `tests/fixtures/` come from
[pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse) (MIT) and are not covered by
it; see [tests/fixtures/README.md](tests/fixtures/README.md).
