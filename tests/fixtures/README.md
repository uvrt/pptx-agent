# Test fixtures

| File | Produced by | Origin |
| --- | --- | --- |
| `real-basic-theme.pptx` | Google Slides | pptx-glimpse, third-party (below) |
| `real-financial-report.pptx` | hand-authored | pptx-glimpse, third-party (below) |
| `real-product-page.pptx` | hand-authored | pptx-glimpse, third-party (below) |
| `sample.pptx` | PowerPoint | pptx-glimpse, third-party (below) |
| `sample-issue-387.pptx` | PowerPoint | pptx-glimpse, third-party (below) |
| `authoring-integration.pptx` | python-pptx | pptx-glimpse, third-party (below) |
| `powerpoint-smartart.pptx` | PowerPoint | ours: PowerPoint's save of a hand-made SmartArt data model (`tests/test_diagram.py`) |
| `generated/trial/*` | pptx-agent | ours, written through this project's API, except as [`generated/trial/PROVENANCE.md`](generated/trial/PROVENANCE.md) says |

## Provenance of the third-party decks

The six decks marked third-party are byte-identical copies of the fixtures of the
sibling renderer [pptx2svg](https://github.com/uvrt/pptx2svg) (`tests/fixtures/`), which
copied them from [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse)'s
`shared-fixtures/` directory (MIT licensed, Copyright (c) Hiroki Sakabe). They are
renderer and editor inputs, not part of this project's own work, and are not covered by
this repository's licence (see [`LICENSE`](../../LICENSE)). pptx2svg's
`tests/fixtures/README.md` and `FIXTURES-README.md` (pptx-glimpse's own listing) describe
each one in full.

| File | SHA-256 |
| --- | --- |
| `real-basic-theme.pptx` | `ce894bc9a8cf021d4a6ab01717e005617dce01056ecd839a60a9f56551d8ad27` |
| `real-financial-report.pptx` | `0e06d2b33bbd48d512300f12327e97821949edd8b7109ad47bc682a18aac181f` |
| `real-product-page.pptx` | `694134ba611d3ee3e7895926d5dfa800d9aeac67253609aa1dbc171c1d9dca97` |
| `sample.pptx` | `44957b1e1cfec8637b10dcf0837050e97afb2d22065bfb676335e1806189f142` |
| `sample-issue-387.pptx` | `bf40a70d9876ed29a0d34c0550b1821b23c76ac04b2c269ad1638ba66341b066` |
| `authoring-integration.pptx` | `a246408b231b66897145b898cc8de0f614ad5a8aa541961840baafb010bc4f4b` |

`real-basic-theme.pptx` embeds the Lato and Raleway faces Google Slides exported with it
(`ppt/fonts/*.fntdata`), licensed under the SIL Open Font License 1.1. No other font data
is in this repository, and no Microsoft font in any form.

Tests that need no fixture of a particular origin hold to every deck here
(`tests/conftest.py`); `generated/` sits one level down so that corpus does not take it in.
