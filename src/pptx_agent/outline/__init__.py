"""The Markdown outline (E6): a reading view of a deck, and slides drafted from Markdown.

A deck is mostly visual, so Markdown is not its source of truth here: it is a reading aid,
cheaper for an agent than renders, and a drafting aid.  :meth:`Document.to_outline` writes
it; :meth:`Document.insert_outline` drafts slides from it into the deck's own layouts and
styles.  Positions, arrangement, charts and diagrams stay with the shape API and the render
loop.

The dialect is docx-agent's -- CommonMark plus GFM tables and strikethrough, read with
markdown-it-py (:mod:`.markdown`) -- and so are the id comments: ``<!-- id ... -->`` on the
line above the block it names, an id at the end of the line it names in a heading.

**The format** (``to_outline``)::

    <!-- s:256 layout: Title and Content -->
    # Where we are <!-- 256.2 -->

    <!-- 256.3 ph: obj 1 -->
    - Revenue grew **12%**
      - Operating margin `11.9%`
    - See [the report](https://example.org/q3)

    <!-- 256.5 -->
    A text box

    <!-- 256.7 table -->
    | Metric | Q3 |
    | --- | --- |
    | Revenue | 4,285 |

    <!-- 256.8 picture -->
    ![Alt text](ppt/media/image1.png)

    <!-- 256.9 chart -->
    [chart: column; title: Sales; series: 2025, 2026; categories: Q1, Q2, Q3]

    <!-- 256.10 smartart -->
    - Plan
      - Detail

    Notes:

    What the speaker says.

* **A slide** is a ``#`` heading: its title placeholder's text (paragraphs and line breaks
  as ``<br>``), the title's shape id at the end, and above it the slide's id, ``hidden``
  for a hidden slide, and its layout's name.  A slide with no title placeholder has an
  empty heading, ``#``.
* **A block** per placeholder and per shape with text, in z-order, groups flattened, under
  a comment with the shape id and what it is: ``ph: <type> <idx>`` for a placeholder (``obj``
  when the type is omitted), ``empty`` for an empty placeholder, ``table``, ``picture``,
  ``chart`` or ``smartart``; a shape's own id alone for text outside placeholders.
  Connectors and shapes without text are left out.
* **Text**: a paragraph showing a bullet (its own, or one it inherits from the layout,
  master or presentation, as pptx2svg resolves it) is a list item at its level -- a
  numbered one an ordered item -- and any other a paragraph; a line break is a hard break.
  Runs that are bold, italic, struck through or in a monospace face (Consolas, Courier
  New, Menlo...) are ``**``, ``*``, ``~~`` and code; a run linking to a web or mail
  address is a link.  These are the runs' own properties: what a style makes bold stays
  the style's.  Text is escaped so it reads back as text.
* **A table** is a GFM table, its first row the header; a cell's paragraphs and line
  breaks are ``<br>``.  As in docx-agent, its cells are named by the table's id:
  ``256.7/cell<row>,<column>``.  A cell a merge covers is empty.
* **A picture** is ``![alt text](source)``: the alt text is ``cNvPr@descr``, the source the
  media part (``ppt/media/image1.png``), or a linked picture's address.
* **A chart** is one line: its type (``column``, ``bar``, ``line``, ``pie``...; a combination
  joined by ``+``), title, series and categories.  **SmartArt** is a list of its node text,
  nested by level.
* **Speaker notes** follow a ``Notes:`` line, to the next slide; a paragraph of the slide
  that says only "Notes:" is written ``Notes\\:``.

``ids=False`` writes no comments: the blocks of a slide are then kept apart by ``---``.

**Drafting** (``insert_outline``) reads the same format, so an outline can be edited and
inserted, and a hand-written one needs none of the comments:

* ``#`` starts a slide; the heading goes to the title placeholder (to a text box in the
  title's place on a layout without one).  ``##`` and deeper headings inside a slide are
  paragraphs.  Text before the first ``#`` is refused.
* Blocks go into **units**: a shape comment starts one, and so does ``---`` (the column
  marker); consecutive blocks without either are one unit.  A unit is a text unit; a
  table or an image is a unit of its own.  ``Notes:`` on a line of its own starts the
  speaker notes: everything after it, to the next ``#``.
* **The layout**, when the slide's comment names one the deck has, is that one.  Otherwise
  the content chooses a role:

  ============================================  =======================================
  The slide holds                               Role (``ST_SlideLayoutType``)
  ============================================  =======================================
  nothing (an empty heading)                    Blank (``blank``)
  only its heading                              Section Header (``secHead``); Title
                                                Slide (``title``) when it will be first
  one text unit                                 Title and Content (``obj``); Title Slide
                                                when first and not a list
  two text units (``---`` or exactly two        Two Content (``twoObj``)
  lists in one unit, split after the first)
  tables or pictures, no text                   Title Only (``titleOnly``)
  one text unit and one table or picture        Two Content
  anything else                                 Title and Content
  ============================================  =======================================

  A role is found by layout type, then by PowerPoint's English name, then by the
  placeholders a layout has (a title and one, two or no text placeholders), in the first
  master -- so a localised or Google Slides template works.  ``layout_map`` overrides the
  choice: keys are roles, their English names or layout names; values a layout of the deck.
* **Placeholders**: a unit naming one (``ph: body 1``) takes it -- one the layout lacks is
  added, inheriting from the master by type; the other units fill free placeholders in the
  layout's order: text the body, content and subtitle ones, a table a table or content one,
  a picture a picture or content one.  A unit whose comment names no placeholder, and what
  finds no free one, becomes a text box, table or picture in the master's body area, on a
  grid.  A placeholder nothing went into is removed (an empty placeholder draws nothing),
  unless the outline names it.  The shapes are stacked in the outline's order.
* **Text** keeps the placeholder's inherited bullets where they agree; a paragraph in a
  bulleted placeholder gets ``buNone``, a list item where nothing bullets one gets
  PowerPoint's own bullet (``•`` in Arial) or numbering (``arabicPeriod``).  Runs carry the
  deck's language; code is Consolas; a link with a scheme is an external hyperlink.
* **Pictures** are fitted into their frame; in a picture placeholder they fill it, cropped,
  and take the placeholder's ``p:ph`` -- as a table put into a placeholder does.
* **Notes** get a notes page, and the deck a notes master when it has none
  (:mod:`pptx_agent.edit.notes`, measured).
* Charts and SmartArt are not drafted: their blocks are left out with an
  :class:`OutlineWarning`.  The whole insertion is one undo step.

**The round trip** -- ``to_outline``, ``insert_outline`` into a new deck from the same
template (``images=`` the source deck), ``to_outline`` again -- gives the same outline up to:

* ids (slide and shape ids, and so the cell ids), and the media part names in picture
  sources;
* charts and SmartArt, which are not drafted.

Everything else is guaranteed: slide order, layouts, hidden slides, which placeholder each
text is in (and empty placeholders the outline names), text, paragraph and list structure
and levels, numbering, the marks and links above, line breaks, tables' cell text and marks,
pictures and alt text, speaker notes.  Not carried, because the outline does not say them:
positions and sizes, fonts and colours, run formatting other than the marks above
(underline, size...), bullet characters and numbering schemes (a numbered list is
``arabicPeriod``), paragraphs' indents and alignment, empty paragraphs, a level that skips
one (written one deeper than its predecessor), slide-jump links, table styles, merges and
header-row settings, shapes without text, connectors, group structure, and anything else
on the slide.  With ``ids=False`` the outline is stable too, but the layout and placeholder
of each block are chosen again by the rules above.
"""

from .draft import OutlineWarning, ROLES, insert_outline
from .read import to_outline

__all__ = ["OutlineWarning", "ROLES", "insert_outline", "to_outline"]
