# Using the deck tools with a model provider

The tools in `pptx_agent.tools` are defined once, in a provider-neutral canonical schema
([`ooxml_edit.tools`](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)). A
`Toolbox` turns them into each provider's definitions, runs the model's calls with
`toolbox.dispatch(session, name, arguments)`, and turns the results back into that
provider's messages. What the tools do: [SUPPORTED.md](SUPPORTED.md); how to use the facts
they return: [GUIDANCE.md](GUIDANCE.md).

| Call | Sends |
| --- | --- |
| `definitions("anthropic")` | every tool, strict on as many as Claude's limits allow (writing tools first); the non-core groups deferred behind the BM25 tool-search tool, `cache_control` on the last core tool |
| `definitions("openai-responses")` | the Responses API (GPT-6): the core as functions, each other group a deferred `namespace`, and `tool_search` |
| `definitions("openai-chat", groups=[...])` | Chat Completions, which has no tool search: the core and the groups named; `toolbox.allowed_tools([...], provider="openai-chat")` narrows a turn's calls without changing `tools` |

**Status.** This repository's own model trials ran on Claude. A production user ran these
tools live on Azure OpenAI's Responses API (October 2026: the default
`definitions("openai-responses")`, `store=False` with
`include=["reasoning.encrypted_content"]`, images in the outputs and in a user message), and
every task completed with the loop below as written. `items += response.output` passes the
reasoning items (with their `encrypted_content`) and a hosted tool search's
`tool_search_call` / `tool_search_output` back as they came; only `function_call` items are
dispatched, by their bare `name` (a deferred tool's group arrives separately, as
`namespace`). The OpenAI definitions and results are also checked offline against OpenAI's
documented rules (`ooxml_edit.tools.adapters.openai_problems`, `openai_input_problems`, and
ooxml-edit's Responses round-trip test). Chat Completions has not been run live.

## Groups

| Group | Tools |
| --- | --- |
| `core` (always loaded) | `open_document`, `new_document`, `save_document`, `undo`, `find_text`, `replace_text`, `render`, `check`, `describe`, `ppt_read_slides`, `ppt_set_text`, `ppt_set_shape`, `batch` |
| `shared_misc` | `close_document`, `read_blob`, `set_properties`, `edit_chart`, `edit_smartart` |
| `ppt_text` | `ppt_format_text`: runs, paragraphs, frames, table cells, a whole table or a row or column |
| `ppt_graphics` | `ppt_add_shape`, `ppt_add_connector`, `ppt_arrange`, `ppt_design_facts` |
| `ppt_objects` | `ppt_add_picture`, `ppt_add_table`, `ppt_edit_table`; `ppt_find_asset` with an asset catalogue configured |
| `ppt_review` | `ppt_comments` |
| `ppt_slides` | `ppt_add_slide`, `ppt_draft_slides`, `ppt_manage_slides`, `ppt_set_theme`, `ppt_edit_master`, `ppt_compare_slides` |

## Setup (every provider)

```python
from ooxml_edit.tools import Toolbox
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS)   # toolbox.close() when done
session = toolbox.session()
session.open(deck_bytes, name="q3-review.pptx")             # the model calls it "d1"
system = toolbox.system_prompt(extra=HOUSE_RULES)          # your guidance after the tools'
task = "In d1, make the title of slide 1 read 'Q3 review', then save it as q3.pptx."
```

**Your own fonts.** Faces kept in a folder of the application's own (licensed fonts the
system does not search) are named once, as configuration -- never by the model:

```python
toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=["/srv/app/fonts"])
session = toolbox.session()                          # the toolbox's folders
other = toolbox.session(font_dirs=["/srv/b/fonts"])  # or a session's own
```

Every render (in the worker process too) and every measurement of text -- `check`'s fit,
overflow and near-wrap, a table's row heights, the facts each changing call and
`save_document` return, the design facts -- then finds those faces, searched before the
system's folders. Without `font_dirs` the environment variable `OOXML_FONT_DIRS` (folders
separated by `os.pathsep`) is read where the toolbox runs; `font_dirs=[]` reads neither.
No tool definition or prompt changes. A deck opened outside the tools takes the same
setting as `Document.font_dirs`.

**The user's locale.** The regional settings a deck is drawn and edited in -- an
`ooxml_common.locales.Locale` -- are configuration too, per toolbox or per session:

```python
from ooxml_common.locales import Locale

dutch = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-NL", list_separator=";")
toolbox = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, locale=dutch)
other = toolbox.session(locale=None)                 # en-US
```

Every render (in the worker process too) passes it to pptx2svg -- a chart's decimal and
thousands separators, the language a date field with none is drawn in -- and the deck
edits in it: a new deck's language, the language of new text in a deck that states none,
a new chart's `c:lang`, a new date field's cached text. Without one it is en-US, never the
server's own. No tool definition or prompt changes. A deck outside the tools takes it as
`Document.locale` (or `Document.new(locale=...)`).

**Your brand assets.** An icon set or logo library of the application's own -- a folder of
SVGs (or bitmaps) and an index of their keywords -- is configuration as well, a setting of
the format:

```python
from ooxml_edit.tools import FolderCatalogue
from pptx_agent.tools import deck_format

icons = FolderCatalogue("/srv/app/brand-icons")      # its index.json or index.csv
toolbox = Toolbox(TOOLS, formats=[deck_format(assets=icons)], groups=GROUPS)
```

With it the model gets `ppt_find_asset` (deferred, in `ppt_objects`): words in, at most 10
(up to 50) `{id, name, keywords}` out, best first; and `ppt_add_picture` takes `asset`, one
of those ids, in place of `image` -- to insert, or on `target` to replace a picture's image.
The bytes go from the catalogue into the deck; the model never sees a path. Without a
catalogue neither is sent (with one: about 141 estimated tokens more, all deferred). The
index is a JSON list, `[{"id": "arrow-up", "file": "arrows/arrow-up.svg", "name": "Arrow
up", "keywords": ["arrow", "up"]}]`, or a CSV with the columns `id,file,name,keywords`
(keywords separated by `;`); `file` is required and stays in the folder, `id` defaults to
its name without the extension ([`ooxml_edit.tools.assets`](https://github.com/uvrt/ooxml-edit/blob/main/src/ooxml_edit/tools/assets.py)).
An object of your own with `search(query, limit)` and `get(id) -> (bytes, MIME type)` works
the same: a search service, a database, a DAM.

An inserted SVG icon is recoloured with `ppt_set_shape`'s `fill` (`"accent2"`, `"#1F4E79"`,
`"accent2 alpha=50%"`) and outlined with its `line` colour and width, as PowerPoint's
Graphics Fill and Outline do it (`Shape.recolor`); on a bitmap picture `fill` and `line` stay
its background fill and border, as in PowerPoint.

**Roles: a reviewer's tools.** A toolbox or session loads roles, additive toolsets
(ooxml-edit 0.45.0, [TOOLS.md, "Roles"](https://github.com/uvrt/ooxml-edit/blob/main/docs/TOOLS.md)):
`review` reads, checks, compares, renders and adds and replies to comments; `resolve`
resolves, reopens and deletes them; `edit` makes every other change (and `accept`, which
accepts and rejects tracked changes, has no deck tool). Editing a comment's text is every
role's. The default is all four, the toolbox as without roles.

```python
reviewer = Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, roles=["review"])
tools = reviewer.definitions("anthropic")       # 15 tools; ppt_comments list/add/reply/edit
system = reviewer.system_prompt(extra=HOUSE_RULES)
session = reviewer.session()                    # ppt_set_text here is refused
second_pass = toolbox.session(roles=["review", "resolve"])   # per session, same toolbox
```

A reviewer's core is open, save, `describe`, `ppt_read_slides`, `find_text`, `render`,
`check`, `batch`, `ppt_compare_slides` and `ppt_comments` (2,212 estimated tokens against
the full core's 3,563); edit_chart offers only `read`, `ppt_manage_slides` only
`find_by_title`. A call outside the session's roles -- in a `batch` too, op by op -- is
refused naming the role it needs; a reviewer reports each finding with the call that would
fix it, for a builder (`roles=["review", "edit"]`, or the default) to apply in one batch.

**Spacing in a review.** A reviewer judging gaps by eye passes the tight ones; `check` with
`spacing` measures them. `{"doc": "d1", "slides": [1], "spacing": 8}` adds `spacing`, the
gaps below 8 pt nearest first -- `{"between": ["257.3", "257.4"], "axis": "y", "gap": 4.0}`
(a heading 4 pt under a panel), `{"inside": "257.5", "shape": "257.7", "edge": "bottom",
"padding": 4.0}` (a chip 4 pt from its card's edge), `{"clearance": "footer", "shape":
"257.5", "gap": 6.5}` -- text measured where it is drawn, not by its box. Measurements only:
put the threshold in the house rules and have the reviewer cite the entry.

Several agents on one deck at once share one session -- the deck, its undo history, refs
and outputs -- each through a view with roles of its own (ooxml-edit 0.46.0); a view
narrows the session's roles, never widens them, and its `context` reaches the handlers of
the application's own tools as `call.caller`:

```python
session = toolbox.session()                     # every role
reviewer = session.view(roles=["review"], context={"agent": "reviewer-1"})
builder = session.view(roles=["review", "edit"], context={"agent": "builder-1"})
toolbox.dispatch(reviewer, "ppt_set_text", args)   # refused: needs the edit role
sent = toolbox.definitions("anthropic", roles=["review"])
system = toolbox.system_prompt(roles=["review"], tools=sent)   # the line follows `sent`
```

**What a call changed, for a live view.** Every result carries `result.touched`
(ooxml-edit 0.57.0), for the application and never in what the model reads: what the call
changed in the deck's own addresses, from the bytes that changed -- so every tool, `undo`
and `redo` included, and nothing a call only read (a `like`, the shape or slide a copy was
made from):

```python
{"doc": "d1", "slide": "s:258", "shapes": ["258.7", "258.9"]}       # shapes redrawn
{"doc": "d1", "slide": "s:258", "shapes": ["258.12"], "created": ["258.12"]}
{"doc": "d1", "slide": "s:258", "shapes": ["258.3"], "removed": ["258.3"]}
{"doc": "d1", "slide": "s:258", "whole": True}       # background, layout, transition, hidden
{"doc": "d1", "slide": "s:258", "notes": True}       # speaker notes; "comments": True
{"doc": "d1", "slide": "s:300", "created": ["s:300"]}               # or "removed"
{"doc": "d1", "layout": "L2", "shapes": ["L2.2"], "affects": ["s:257", "s:259"]}
{"doc": "d1", "master": "M0", "shapes": ["M0.7"], "affects": [...every slide on it]}
{"doc": "d1", "theme": "M0", "affects": [...]}      # the theme of master M0
{"doc": "d1", "deck": ["order"], "affects": ["s:257", "s:258"]}     # slides renumbered
{"doc": "d1", "part": "ppt/..."}                     # anything else, never left out
```

A shape is listed when its own XML changed (a group's without its members), when it moved
into or out of a group, or when it moved among its siblings (the fewest that explain the
new order); a part it draws from -- a chart, its workbook, a picture's image -- counts as a
change to it. A layout's or master's `affects` are the slides built on it, which show the
change; they are not listed again. `deck` also says `size` and `text_styles` (every slide
affected), `sections`, `table_styles` (the slides with tables), `document_properties`, and
the relationship name of any other deck-level part (`authors`, `notesMaster`). A `batch`
has the union, and `result.touched_by_op` each op's (`{0: [...], 2: [...]}`). The library
gives the same from a watch: `with deck.history.watch() as change: ...`, then
`pptx_agent.edit.touched.touched(deck, change)`.

### Tables from data

`ppt_add_table` and `ppt_edit_table`'s `set_cells` take a table's text from a source, as
`edit_chart` takes a chart's (ooxml-edit 0.22.0): `from` is a dataset (`d:<id>`) or a table of
another open document (`t:<name>`), `from_columns` picks its columns by name, in the order
placed (default: all), and every row is placed. A header row of the column names goes
above them unless `header` gives its texts, or `[]` for none. Numbers are written as a
chart's category labels write them (General, the session locale's decimal separator),
dates as the source gives them, a blank as an empty cell. `ppt_add_table` sizes the table
to fit; `set_cells` writes from row `index` (default 0), column 0, and refuses a source
the table has no room for. Nothing is cut: past 200 rows, 50 columns or 5,000 cells the call
is refused (`limit`).

```json
{"doc": "d1", "slide": "s:256", "box": {"x": 40, "y": 100, "w": 880, "h": 300},
 "from": "d:checks", "from_columns": ["check", "rating", "due"]}
```

The result's `data.from` says where every placed value came from -- the source, its
document and version, `header_row` and `first_row` (table rows from 0; source row `n`,
from 1, is table row `first_row + n - 1`), and per column its table column, source column,
source rows as runs and their sources:

```json
{"source": "d:checks", "doc": "d2", "version": 3, "header_row": 0, "first_row": 1,
 "columns": [{"col": 1, "column": "rating", "rows": "1-12",
              "sources": ["Checks!C2:C13"]}, ...]}
```

The deck keeps it too, in the table's frame (the source without the session's document
handle, and the source document's name added), as `edit_chart` with `from`
keeps a chart's -- a `UVRT_PROVENANCE` tag in the graphic frame's tags part
(`p:nvPr/p:custDataLst/p:tags` -> `ppt/tags/tagN.xml`), which PowerPoint keeps through
open, edit and save. It is read back after a save with `session.provenance(doc)` or
`Document.open(...).provenance()` -> `{frame address: record}`. A table edit keeps the
record true or drops what it no longer says: a cell of the block written without a source
(`set_cells`, `ppt_set_text`, `replace_text`, `merge`) drops its column; a row inserted or
deleted above the block moves it, inside it drops the record; a column inserted or deleted
moves the columns' entries (and drops the deleted one's). Copies of the frame, its slide
or its shapes have their own copy of the tags part; deleting the frame or the slide
deletes it.

A figure in text keeps its source the same way. The text tools (`ppt_set_text`,
`ppt_add_shape`, `ppt_add_table`'s `data`, `set_cells`, `replace_text`'s `replace`) take a
reference to one value, `{{d:by_topic/door/mean}}` or `{{t:checks/N-1/rating}}` (`#3` names
a row by position; `:0.0`, `:#,##0`, `:0%` after it format a number), and write the value in
the session's locale; the result lists each (`data.facts`: `from`, `text`, `sources`). The
deck keeps them in a `UVRT_TEXT_FACTS` tag of the shape's tags part (a table's frame for its
cells), each with its place -- `p`, `at`, a table's `cell` -- and `Document.provenance()`
lists them under the shape's address (`facts`) while their text is where they say: an edit
drops the facts it touches and moves the rest. `\{{` writes a reference literally.

A slide's number is a reference too: `{{slide:s:262}}` (or `{{slide:$results}}`, a ref to a
slide) writes the number that slide's slide-number field shows -- its position plus
`firstSlideNum` minus one, hidden slides counted -- and keeps it as a fact (`source`
`slide:s:262`, `slide` 262). After every change to the slide order (adding, deleting,
moving, duplicating slides; a batch at its end) the deck rewrites the numbers that moved,
in the same undo step, and the checks list them (`slide_numbers`: `shape`, `slide`,
`text`, `was`); `save_document` and `Document.save` rewrite them all first, which catches a
deck PowerPoint reordered. A reference whose slide was deleted keeps its last number, is
marked `gone` (a later slide with the same id is never taken for it) and is listed as
`missing_slides` by `check` and `save_document`. `{{slide:3}}` is refused with the slide
that is number 3. PowerPoint has no field that shows another slide's number: in a deck
edited in PowerPoint the numbers are plain text and stay stale until this library next
saves it (`pptx_agent.edit.slidenumbers`).

### Measuring a table before placing it

`ppt_add_table` with `measure: true` builds nothing: it takes the same arguments as an add
(`from`, `from_columns`, `header` or literal `data`, `style`, and `size` -- every cell's
text size -- and `widths` -- the column widths -- which an add takes too) and returns each
row's height as PowerPoint draws it, the table's height, whether it fits, and the pages its
rows split into with the header row repeated on each. A page's room is the box's `h`, or the
slide's content area height when `h` is 0; `pages` number the rows below the header from 1
(with `from`, its source rows) and give each page's table height, header included:

```json
{"heights": [28.8, 28.8, 50.4, ...], "height": 1022.4, "page_height": 342.63,
 "fits": false, "header": 28.8,
 "pages": [{"rows": "1-9", "h": 309.6}, {"rows": "10-18", "h": 331.2}, ...]}
```

The table is built on a copy of the deck and measured as `Table.drawn_row_heights`
measures a placed one, so a table placed with these arguments and a box `h` of 0 is drawn
at exactly these heights; the deck is not changed (no version, no undo step). A row too
tall for a page under the header gets a page of its own and is listed in `too_tall`. The
library call is `Slide.measure_table(...)`, a `TableMeasure`.

### A long table on the next slides, and fewer rows

`ppt_add_table` with `continue: true` places every row: it measures the pages as `measure`
does, places the first on the target slide and each next one on a new slide right after
it -- on the target slide's layout, its title the target's title text (as it is; retitle
them with `ppt_set_text` if you like), its other empty placeholders deleted -- in the same
box, the header row repeated. Each table's rows are set to their measured heights, so a
page is drawn as measured rather than stretched by the box's height shared evenly among
rows. One call, one undo step. The result lists the pages, and each table keeps the record
of its own rows (`first_row` is the table row of the page's first source row):

```json
{"pages": [{"slide": "s:256", "table": "256.4", "rows": "1-9", "h": 309.6, "from": {...}},
           {"slide": "s:259", "table": "259.4", "rows": "10-18", "h": 331.2, "from": {...}}]}
```

### Row heights: the box is an upper bound

A row's height is a minimum (PowerPoint draws a row whose text needs more taller), so
`ppt_add_table` sets each row to its text's height and the frame to their sum; the box's
`h` is the most the table may take, and the summary says how much it does: `Added a 5x3
table 256.3: it takes 96 pt of the 276 pt box, each row its text's height (header row
19.2 pt); fill_box: true shares the rest among the body rows`. With `fill_box: true` the
body rows share the box (`its body rows share the 276 pt box (header row 19.2 pt)`); the
header keeps its text's height. Rows taller together than the box run past it, as they
always did: `it takes 340 pt, 64 pt past the 276 pt box, each row its text's height;
continue: true goes on over the next slides`. `insert_row` copies its `like` row's height
and `delete_row` takes its row's out of the frame; no other row changes. `set_heights` with
`rows` sets only those rows (`"rows": [0], "sizes": [24]`), `set_widths` with `columns`
only those columns. `ppt_set_shape`'s `h` on a table still rescales every row in
proportion and warns, naming the header row's new height.

To place fewer rows, name them: a row range of the source, `t:items/#1:#9` (rows 1 to 9,
from 1, the last included) or `d:by_region/#10:` (row 10 to the last); each value keeps
its sources. `rows` smaller than the source is refused saying so: nothing is cut.

The provider SDKs read their keys from the environment (`ANTHROPIC_API_KEY`,
`OPENAI_API_KEY`). When the loop ends, `session.take_outputs()` holds what `save_document`
wrote, as bytes.

## Claude (Messages API)

```python
import anthropic

client = anthropic.Anthropic()
tools = toolbox.definitions("anthropic")
messages = [{"role": "user", "content": task}]
while True:
    response = client.messages.create(model="claude-sonnet-5-5", max_tokens=8000,
                                      system=system, tools=tools, messages=messages)
    messages.append({"role": "assistant", "content": response.content})
    calls = [block for block in response.content if block.type == "tool_use"]
    if not calls:
        break
    results = [(c.id, toolbox.dispatch(session, c.name, c.input)) for c in calls]
    messages.append(toolbox.render_results("anthropic", results))   # tool_result blocks
```

## OpenAI (Responses API)

```python
from openai import OpenAI

client = OpenAI()
tools = toolbox.definitions("openai-responses")
items = [{"role": "user", "content": task}]
while True:
    response = client.responses.create(model="gpt-6", instructions=system, tools=tools,
                                       input=items)
    items += response.output
    calls = [item for item in response.output if item.type == "function_call"]
    if not calls:
        break
    results = [(c.call_id, toolbox.dispatch(session, c.name, c.arguments)) for c in calls]
    items += toolbox.render_results("openai-responses", results)    # function_call_output items
```

## OpenAI (Chat Completions)

```python
from openai import OpenAI

client = OpenAI()
tools = toolbox.definitions("openai-chat", groups=["ppt_graphics", "ppt_objects"])
messages = [{"role": "system", "content": system}, {"role": "user", "content": task}]
while True:
    reply = client.chat.completions.create(model="gpt-6", tools=tools,
                                           messages=messages).choices[0].message
    messages.append(reply)
    if not reply.tool_calls:
        break
    results = [(c.id, toolbox.dispatch(session, c.function.name, c.function.arguments))
               for c in reply.tool_calls]
    messages += toolbox.render_results("openai-chat", results)
```

A Chat Completions tool message is text only, so `render_results` follows the tool
messages with one user message carrying the images (a `render`'s PNGs), each labelled with
its call. `dispatch` takes the arguments as a dict or as the JSON string OpenAI sends, and
never raises for the model's mistakes: an error comes back as a result the model can read,
naming the valid options where there are some.
