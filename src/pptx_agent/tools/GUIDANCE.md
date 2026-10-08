# Guidance for the application's thinking layer (decks)

`pptx_agent.tools` gives a model the means to edit a deck and the facts about the result.
It ships no judgement: no house style, no palette rules, no density limits. What a good slide
is for *your* users is the application's to say, in its own prompt and in its own review
pass. This page is for the developer of that layer: how to use the facts the tools return,
how to run a review pass, which model to route to, and where house rules go. The numbers
come from trial 3 (Sonnet 5.5, 46 runs), its fix round and the Haiku 5.5 trial (46 runs);
the tool roadmap in `ooxml-edit` (`docs/TOOLS-ROADMAP.md`) has the detail.

## The facts the tools return

Every changing call returns `checks` for the slides it touched. All of them are facts about
the deck as PowerPoint will show it, with shape addresses:

| Fact | Where | What it says |
|---|---|---|
| `overflows` | every changing call, `check` | text that does not fit its box: `needed`, `available`, `overflow` in points |
| `within_allowance` | every changing call | over by less than the empty bottom of the last line: not an overflow (`allowance_note` says why) |
| `near_wrap` | every changing call | a line with under 0.16 pt to spare: PowerPoint may break it where the measurement does not |
| `collisions` | every changing call, `check` | text over text, a line through text (`crossing`, pt), boxes overlapping (`check` with `boxes: true`) |
| `off_slide` | every changing call, `check` | a shape past the slide's edge, by how much; for a table whose rows grew past the bottom, `rows_past` (first and last, 1-based) and `rows_fit` -- where to split it (`rows_note`) |
| `layout` | changing calls that touched shapes | at most five: a box 1-3 pt off the line its like neighbours are on; one uneven gap in a row of like boxes; a paragraph size unlike its like boxes'; a label much farther from its marker than the others of its kind. Each carries `fix`, the exact call that makes it agree (`ppt_set_shape 256.6 y=150`) |
| `validate` | every changing call, `check` | validation problems the call added (`new`) or removed (`fixed`) against the deck as opened |
| `missing_glyphs` | `render` | text the image leaves out because the renderer has no font for it (`face`, `script`, `sample`); the deck is unchanged and PowerPoint draws it (`missing_glyphs_note`): install `pptx2svg-fonts` where the tools run |
| `unresolved` | `save_document` | the overflows, collisions and off-slide shapes still in the deck when it was saved (the save goes ahead) |
| design facts | `ppt_design_facts`, `check` with `design` | palette and theme roles, sets of like shapes with their colours and whether a legend covers them, the largest empty regions, alignment lines and near-misses, shape vocabulary, text sizes, lines over text |
| problem facts | `check` with `facts` | colours that are not theme colours; wrap margins of titles and text boxes |

How to use them:

- **Fit and collisions are what the deck is; a render is not.** A model told to read `checks`
  after every change resolves most overflows itself. Trial 3's one structural miss on save
  (a line crossing bar text) was listed in the call's result and in `check`, and the model
  saved anyway; Haiku 5.5 did the same twice and then reported "no collisions". So check
  `save_document`'s `unresolved` in the application, not only in the model's prompt (see
  "Gates" below).
- **Layout facts are measured against the slide's own shapes.** They say that a box is 2 pt
  off a line three like boxes are on, not that it should be on it. Most of the time it should;
  sometimes the offset is the point (an indented sub-item). Tell the model to apply the `fix`
  unless the difference is intended, and say so in its reply. On 85 trial decks the facts
  fired three times in all (two labels far from their bubbles, one header 1 pt larger than its
  like boxes'), and never on the fixture decks with every shape touched: they are rare by
  design, so when one appears it is worth a look.
- **Design facts are inputs to your rules.** `ppt_design_facts` reports that a set of four
  like boxes uses four accent hues and that no legend covers them; whether that is a
  "rainbow" to fix is your house rule. Ask for it on slides with a graphic, before saving.
  Sonnet 5.5 called it in 13 of 16 graphics runs when the house rules asked; Haiku 5.5 in 6
  of 16.

## A review pass

A review pass is the application looking at the result before it accepts it. It runs
outside the model's tool loop: the application calls the tools itself, through
`toolbox.dispatch`, and decides what to do with what it finds.

```python
def review(toolbox, session, doc, numbers, house_rules, model_call):
    """Facts and renders for the slides the task changed (numbers from 1), judged by the
    application's own rules in its own model call."""
    facts = toolbox.dispatch(session, "check", {"doc": doc, "slides": numbers,
                                                "include": ["fit", "collisions", "validate", "facts"]})
    deck = toolbox.dispatch(session, "describe", {"doc": doc}).data
    ids = [deck["slides"][n - 1]["id"] for n in numbers]
    design = [toolbox.dispatch(session, "ppt_design_facts", {"doc": doc, "slide": sid}).data
              for sid in ids]
    images = toolbox.dispatch(session, "render", {"doc": doc, "slides": numbers[:4]}).images
    return model_call(images=images, facts={"check": facts.data, "design": design},
                      rules=house_rules)
```

- **When:** after the model reports it is done, and before the application hands the file
  on. One pass, and at most one more after the model has fixed what it found: in the trials
  a second render rarely changed anything a first one had not.
- **What to send the reviewer:** the renders (`render`, 1280 px: about 1,200 tokens a slide on
  Claude), the `check` facts and the design facts as JSON, the task's brief, and your house
  rules. Ask for findings with shape addresses, each tied to a fact or a rule, not general
  impressions. The reviewer can be the same model or a cheaper one; give it no tools.
- **What to do with findings:** send them back to the editing model as one user turn ("The
  review found: ... Fix these, or say why a finding does not apply.") and let it work through
  the tools again.
- `check`'s `include: ["app"]` is reserved for a critique hook the application registers;
  no hook is registered for decks today (`check` says so in its `notes`). Run the pass
  yourself as above.

### Gates

The tools never refuse a save over a fit or layout fact: they report. If your application
must not hand on a deck with an overflow or a line through text, gate it yourself:

```python
saved = toolbox.dispatch(session, "save_document", {"doc": doc, "name": name, "format": "pptx"})
left = (saved.data or {}).get("unresolved")
if left:
    # Re-prompt with the facts, or refuse the output: a plain listing is not enough for
    # every model (Haiku 5.5 saved over reported collisions and said there were none).
    ...
```

Validation is already a gate: `save_document` refuses a deck with new validation problems,
and only the application can allow them (`Toolbox(allow_new_problems=True)`).

## Model routing

| Work | Model | Evidence (main arm, per run) |
|---|---|---|
| Word documents (w1-w9) | **Haiku 5.5** | 18/18 succeed, every one graded 10; USD 0.006 against Sonnet 5.5's 0.113 (about 5%) |
| Deck edits without a designed graphic (p1-p6: a deck from an outline and template, restating figures and charts, splitting a slide, a flow chart, a rebrand, a restructure) | **Haiku 5.5** | 12/12, every one graded 10; USD 0.008 against 0.147 |
| Designed graphics from a brief (a phased approach, a Gantt, an org chart, a 2x2 matrix) | **Sonnet 5.5** | Sonnet 6/8 succeed, mean grade 9.13, USD 0.27; Haiku 4/8, mean 7.75, USD 0.016 |

- **Why graphics go to Sonnet:** Haiku misread a design requirement as optional (one colour
  for every workstream, 3 of 4 p8 runs), saved over collisions it had been told about, and
  skipped the house rules more often. Its geometry was exact; its judgement was not.
- **A routing signal:** the brief asks for a new graphic built from shapes (a timeline,
  chart-like diagram, org chart, matrix, process with more than a few steps), or the edit
  creates more than a handful of shapes on one slide. Everything else -- text, tables,
  charts from data, slides from layouts, theme and rebrand work, comments -- goes to Haiku.
- **Haiku first, Sonnet on failure** is plausible for graphics behind a hard gate (unresolved
  facts, failed checks), but the trials did not test it.
- Both models found deferred tools through tool search (Sonnet 65 of 66 searches, Haiku every
  search); use the default `toolbox.definitions("anthropic")`.

## House rules: in the application's prompt

House rules -- a palette, when a legend is needed, minimum text sizes, how much empty space
is too much -- go in the application's own guidance, appended after the shipped fragments.
They are written against the facts the tools report, so the model (and your reviewer) can
check them. The documented example, from `pptx_agent.tools` (documentation only; nothing
like it ships):

```python
HOUSE_RULES = '''House rules (Acme):
- Before saving a slide with a graphic, call ppt_design_facts on it.
- If a set of like shapes uses more than two accent hues and has no legend, recolour
  it to tints of one accent, or add a legend if the colours carry meaning.
- Leave no empty region larger than a quarter of the content area.
- Body text is at least 12 pt.'''
system = toolbox.system_prompt(extra=HOUSE_RULES)
```

- Keep rules **checkable**: each names a fact (`ppt_design_facts`' groups and legends, empty
  regions, text sizes) and what to do about it. "Make it look professional" gives the model
  nothing to measure.
- Keep them **short**: the trial graphics tasks used the four lines above, and Sonnet followed
  them in 13 of 16 runs.
- Put the same rules in the reviewer's brief, so editing and review judge alike.
- Do not put house rules in tool descriptions or patch them into the libraries: the tools
  stay the same for every application, and a rule belongs to one.
