"""Comments: PowerPoint's modern (threaded) comments on a slide or on a shape.

A thread is a comment with its replies, on a slide or on one shape of it; a thread can be
resolved and reopened.  This is the comment model of PowerPoint for Microsoft 365 and
PowerPoint for Mac 16 (the ``p188`` namespace).  The older comments
(``ppt/comments/comment<n>.xml`` with ``commentAuthors.xml``) are neither read nor written.

What the package holds, measured on PowerPoint for Mac 16.89: a deck with hand-written
comments, opened unprompted and saved again through its AppleScript; what it kept and how
it wrote it back:

* **The authors** of every comment and reply: ``ppt/authors.xml``
  (``application/vnd.ms-powerpoint.authors+xml``), related from the presentation by
  ``http://schemas.microsoft.com/office/2018/10/relationships/authors``;
  ``p188:authorLst`` of ``p188:author`` with ``id`` (a braced GUID), ``name``, ``initials``,
  ``userId`` and ``providerId``, in no order that means anything (two re-saves wrote them
  back reversed).
* **One comments part per slide**: ``ppt/comments/modernComment_<sldId hex>_0.xml``
  (``application/vnd.ms-powerpoint.comments+xml``; PowerPoint renamed a part to this form,
  the slide id in hexadecimal and ``0``), related from the slide by
  ``http://schemas.microsoft.com/office/2018/10/relationships/comments``, and named again in
  the slide's own ``p:extLst``: a ``p:ext`` with URI ``{6950BFC3-D8DA-4A85-94F7-54DA5524770B}``
  holding ``p188:commentRel r:id``, written before the slide's ``p14:creationId``.
* **A thread**: ``p188:cm`` with ``id``, ``authorId``, ``status="resolved"`` when resolved,
  and ``created`` (written back without fractional seconds when they are zero, with a ``Z``
  when it had one); then its anchor -- ``pc:sldMkLst`` (``pc:docMk``, ``pc:sldMk cId sldId``)
  for the slide, or ``ac:deMkLst`` (the same two, then the moniker of the shape's kind --
  ``ac:spMk``, ``ac:picMk``, ``ac:cxnSpMk``, ``ac:graphicFrameMk`` for a table or chart,
  ``ac:grpSpMk`` for a group -- with ``id``, the shape's ``cNvPr`` id, and ``creationId``,
  the shape's ``a16:creationId`` or, when it has none, the nil GUID) for a shape -- ``p188:replyLst`` of ``p188:reply`` (``id``, ``authorId``, ``created``,
  ``p188:txBody``), and the comment's own ``p188:txBody`` (``a:bodyPr``, ``a:lstStyle``,
  paragraphs).  ``cId`` is the slide's ``p14:creationId`` (``0`` when it has none: kept as
  written).  ``p188:pos`` is optional (a thread without one was kept).  Threads are written
  back in order of ``created``.

The moniker per kind, measured: PowerPoint for Mac 16.106, commenting through its UI on a shape,
picture, connector, table, chart, group and a shape in a group, wrote ``ac:spMk``,
``ac:picMk``, ``ac:cxnSpMk``, ``ac:graphicFrameMk`` (table and chart alike), ``ac:grpSpMk``
and, for the shape in the group, its own ``ac:spMk`` alone (no group chain) -- [MS-ODRAWXML]
2.29.3's CT_DrawingElementMonikerList.  None of those shapes had an ``a16:creationId``:
PowerPoint wrote ``creationId`` as the nil GUID and added none to the shape (nor a
``p14:creationId`` to the slide: ``cId="0"``).  An ``ac:spMk`` on a chart (what this module
wrote before) put the marker at the chart, but selecting the thread did not select the chart.
Reading accepts every kind.

PowerPoint kept threads whose author, slide creation id or shape it could not find, so it
does not check them; this module writes only threads that resolve.  What PowerPoint writes
when *it* adds a comment -- its own ``created`` form, ``p188:pos`` -- could not be measured
(no AppleScript command adds one, and the screen was not available to the UI): ``created`` is
written as UTC with a ``Z``, as docx-agent writes Word's dates, and no ``p188:pos``.

Ids are GUIDs made from the content and the clock (``uuid5``), so the same calls give the
same bytes; a comment's address is ``c:`` and the first eight hex digits of its GUID
(``c:3F2A9C1B``), and the full GUID is accepted too.  A duplicated slide does not take its
original's comments (they are a review of that slide); deleting a slide removes its comments.

For example::

    thread = deck.add_comment("256.3", "Is this the Q3 figure?", author="Dana Reviewer")
    deck.reply_to_comment(thread.address, "Yes, from the September close.", author="Sam")
    deck.resolve_comment(thread.address)
"""

from __future__ import annotations

import datetime as _dt
import uuid
from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Iterable

from ..oxml.xml import Element, find, findall, make, parse_xml, qn, register_namespaces, \
    serialize

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document, Shape, Slide

P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
register_namespaces({"p188": P188, "pc": PC, "ac": AC})

REL_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
REL_AUTHORS = "http://schemas.microsoft.com/office/2018/10/relationships/authors"
CT_COMMENTS = "application/vnd.ms-powerpoint.comments+xml"
CT_AUTHORS = "application/vnd.ms-powerpoint.authors+xml"
#: The slide's ``p:ext`` naming its comments part.
COMMENT_REL_EXT_URI = "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"
_SLIDE_CREATION_URI = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
_SHAPE_CREATION_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NIL_GUID = "{00000000-0000-0000-0000-000000000000}"
_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
#: The namespace of the ids made here (uuid5).
_IDS = uuid.UUID("6b1f5c1e-3a52-4f0b-9d4e-7a0c2b9e5d11")


# -- what a read gives -------------------------------------------------------------------------


@dataclass
class Reply:
    """A reply in a thread."""

    id: str                     #: the GUID, braced
    author: str
    initials: str
    created: str
    text: str

    @property
    def address(self) -> str:
        return "c:" + self.id.strip("{}")[:8]


@dataclass
class Comment:
    """A thread: the comment, where it is, and its replies."""

    id: str                     #: the GUID, braced
    slide_id: int
    shape: str | None           #: the shape's address, or ``None`` for the slide
    author: str
    initials: str
    created: str
    text: str
    done: bool
    replies: list[Reply] = field(default_factory=list)

    @property
    def address(self) -> str:
        """``c:`` and the GUID's first eight hex digits."""
        return "c:" + self.id.strip("{}")[:8]

    @property
    def target(self) -> str:
        """What it is on: the shape's address, or ``s:<sldId>``."""
        return self.shape or f"s:{self.slide_id}"


# -- the parts ---------------------------------------------------------------------------------


def _authors_part(package, create: bool = False) -> str | None:
    presentation = package.presentation_part()
    related = [p for p in package.related_parts_of_type(presentation, REL_AUTHORS)
               if package.has_part(p)]
    if related or not create:
        return related[0] if related else None
    part = "ppt/authors.xml" if not package.has_part("ppt/authors.xml") else \
        package.unused_part_name("ppt/authors{n}.xml")
    root = f'<p188:authorLst xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p188="{P188}"/>'
    package.add_part(part, _DECL + root.encode(), CT_AUTHORS, override=True)
    package.add_relationship(presentation, REL_AUTHORS, part)
    return part


def _comments_part(package, slide_part: str) -> str | None:
    related = [p for p in package.related_parts_of_type(slide_part, REL_COMMENTS)
               if package.has_part(p)]
    return related[0] if related else None


def _slide_extensions(package, slide_part: str) -> Element:
    root = package.tree(slide_part)
    extensions = find(root, "p:extLst")
    if extensions is None:
        extensions = make("p:extLst")
        root.append(extensions)
    return extensions


def _new_comments_part(package, slide_part: str, slide_id: int) -> str:
    name = f"ppt/comments/modernComment_{slide_id:X}_0.xml"
    part = name if not package.has_part(name) else \
        package.unused_part_name(f"ppt/comments/modernComment_{slide_id:X}_{{n}}.xml")
    root = f'<p188:cmLst xmlns:a="{_A}" xmlns:r="{_R}" xmlns:p188="{P188}"/>'
    package.add_part(part, _DECL + root.encode(), CT_COMMENTS, override=True)
    rel_id = package.add_relationship(slide_part, REL_COMMENTS, part)
    extensions = _slide_extensions(package, slide_part)
    extension = make("p:ext", uri=COMMENT_REL_EXT_URI)
    reference = _declared("p188:commentRel", "p188")
    reference.set(qn("r:id"), rel_id)
    extension.append(reference)
    extensions.insert(0, extension)          # measured: before the slide's creationId
    package.mark_dirty(slide_part)
    return part


def drop_comments(package, slide_part: str) -> None:
    """Remove a slide's comments: its relationship, its ``p:ext`` and (when nothing else
    holds it) the part.  For a duplicated slide, whose copy is not reviewed."""
    root = package.tree(slide_part)
    for extension in findall(root, "p:extLst/p:ext"):
        if extension.get("uri") == COMMENT_REL_EXT_URI:
            extension.getparent().remove(extension)
            package.mark_dirty(slide_part)
    extensions = find(root, "p:extLst")
    if extensions is not None and len(extensions) == 0:
        root.remove(extensions)
    ids = [rel.id for rel in package.relationships(slide_part).values()
           if rel.type == REL_COMMENTS]
    if ids:
        package.release(slide_part, ids)


def _declared(tag: str, prefix: str, **attributes: str) -> Element:
    """``make``, declaring ``prefix`` on the element itself, as PowerPoint writes these."""
    from lxml import etree

    element = etree.Element(qn(tag), nsmap={prefix: {"p188": P188, "pc": PC, "ac": AC}[prefix]})
    for name, value in attributes.items():
        element.set(name, value)
    return element


def _creation_id(element: Element | None, uri: str, tag: str, attribute: str) -> str | None:
    if element is None:
        return None
    for extension in element.iter(qn("p:ext"), qn("a:ext")):
        if extension.get("uri") != uri:
            continue
        node = extension.find(qn(tag))
        if node is not None and node.get(attribute):
            return node.get(attribute)
    return None


#: The moniker of each kind of drawing element ([MS-ODRAWXML] 2.29.3: CT_DrawingElementMonikerList
#: is ``pc:docMk``, ``pc:sldMk`` and one of these), by the shape element's local name.
MONIKERS = {"sp": "ac:spMk", "grpSp": "ac:grpSpMk", "graphicFrame": "ac:graphicFrameMk",
            "cxnSp": "ac:cxnSpMk", "pic": "ac:picMk", "contentPart": "ac:inkMk"}
_MONIKER_TAGS = frozenset(qn(tag) for tag in MONIKERS.values())


def _moniker_of(shape: "Shape") -> str:
    from ..oxml.xml import local_name

    kind = local_name(shape._element)
    if kind not in MONIKERS or _cnv_id(shape) is None:
        raise ValueError(f"cannot anchor a comment to {shape.id} ({kind}): it has no "
                         "drawing-element id; comment on its slide instead")
    return MONIKERS[kind]


# -- reading -----------------------------------------------------------------------------------


def _text_of(body: Element | None) -> str:
    if body is None:
        return ""
    paragraphs = []
    for paragraph in findall(body, "a:p"):
        parts = []
        for node in paragraph.iter(qn("a:t"), qn("a:br")):
            parts.append("\v" if node.tag == qn("a:br") else (node.text or ""))
        paragraphs.append("".join(parts))
    return "\n".join(paragraphs)


def _authors(package) -> dict[str, tuple[str, str]]:
    part = _authors_part(package)
    if part is None:
        return {}
    return {node.get("id"): (node.get("name") or "", node.get("initials") or "")
            for node in findall(package.tree(part), "p188:author")}


def read_comments(document: "Document", slides: "Iterable[Slide] | None" = None) -> list[Comment]:
    """Every thread on ``slides`` (default all), slide by slide, in document order."""
    package = document.package
    authors = _authors(package)
    out: list[Comment] = []
    for slide in document.slides if slides is None else slides:
        part = _comments_part(package, slide.part_path)
        if part is None:
            continue
        shapes = {}
        for shape in _all_shapes(slide):
            shapes.setdefault(_cnv_id(shape), shape.id)
        for node in findall(package.tree(part), "p188:cm"):
            out.append(_thread(node, slide, shapes, authors))
    return out


def _cnv_id(shape: "Shape") -> str | None:
    from .ids import cnv_pr

    properties = cnv_pr(shape._element)
    return None if properties is None else properties.get("id")


def _all_shapes(slide: "Slide"):
    """Every shape of the slide, groups and their members included, once each."""
    out, seen = [], set()

    def walk(items):
        for shape in items:
            if shape._element in seen:
                continue
            seen.add(shape._element)
            out.append(shape)
            if shape.kind == "group":
                walk(shape.children)

    walk(slide.shapes)
    return out


def _marker(anchor: Element) -> Element | None:
    """The drawing-element moniker of a ``ac:deMkLst``: its last ``ac:`` child that is one of
    :data:`MONIKERS` (any kind, and an ``ac:grpSpMk`` met before it, are accepted)."""
    found = None
    for child in anchor:
        if child.tag in _MONIKER_TAGS:
            found = child
    return found


def _thread(node: Element, slide: "Slide", shapes: dict, authors) -> Comment:
    shape = None
    marker = None
    for anchor in node:
        if anchor.tag == qn("ac:deMkLst"):
            marker = _marker(anchor)
    if marker is not None:
        shape = shapes.get(marker.get("id"))
    name, initials = authors.get(node.get("authorId"), ("", ""))
    replies = []
    for reply in findall(node, "p188:replyLst/p188:reply"):
        who, short = authors.get(reply.get("authorId"), ("", ""))
        replies.append(Reply(reply.get("id"), who, short, reply.get("created") or "",
                             _text_of(find(reply, "p188:txBody"))))
    return Comment(node.get("id"), slide.slide_id, shape, name, initials,
                   node.get("created") or "", _text_of(find(node, "p188:txBody")),
                   node.get("status") == "resolved", replies)


# -- writing -----------------------------------------------------------------------------------


def _stamp(date: "_dt.datetime | None") -> str:
    moment = date or _dt.datetime.now(_dt.timezone.utc)
    if moment.tzinfo is not None:
        moment = moment.astimezone(_dt.timezone.utc)
    return moment.strftime("%Y-%m-%dT%H:%M:%SZ")


def initials_of(name: str) -> str:
    """``Dana Reviewer`` -> ``DR``: the first letter of each word, at most three."""
    letters = [word[0] for word in name.replace("-", " ").split() if word[:1].isalpha()]
    return "".join(letters[:3]).upper() or name[:1].upper()


def _ensure_author(package, name: str, initials: str | None) -> str:
    part = _authors_part(package, create=True)
    root = package.tree(part)
    for node in findall(root, "p188:author"):
        if node.get("name") == name:
            return node.get("id")
    identifier = "{" + str(uuid.uuid5(_IDS, f"author:{name}")).upper() + "}"
    author = make("p188:author", id=identifier, name=name,
                  initials=initials or initials_of(name), userId=name, providerId="None")
    root.append(author)
    package.mark_dirty(part)
    return identifier


def _used_ids(document: "Document") -> set[str]:
    used = set()
    for thread in read_comments(document):
        used.add(thread.id.strip("{}")[:8].upper())
        used.update(reply.id.strip("{}")[:8].upper() for reply in thread.replies)
    return used


def _new_id(document: "Document", seed: str) -> str:
    used = _used_ids(document)
    counter = 0
    while True:
        value = str(uuid.uuid5(_IDS, f"{seed}:{counter}")).upper()
        if value[:8] not in used:
            return "{" + value + "}"
        counter += 1


def _body(text: str, language: str) -> Element:
    body = make("p188:txBody")
    body.append(make("a:bodyPr"))
    body.append(make("a:lstStyle"))
    for line in str(text).split("\n"):
        paragraph = make("a:p")
        pieces = line.split("\v")
        for index, piece in enumerate(pieces):
            if index:
                paragraph.append(make("a:br"))
            if piece:
                run = make("a:r")
                run.append(make("a:rPr", lang=language))
                node = make("a:t")
                node.text = piece
                run.append(node)
                paragraph.append(run)
        if not len(paragraph):
            paragraph.append(make("a:endParaRPr", lang=language))
        body.append(paragraph)
    return body


def _locate(document: "Document", identifier: str):
    """The ``(part, element, thread element)`` an address or GUID names."""
    key = str(identifier).strip()
    if key.startswith("c:"):
        key = key[2:]
    key = key.strip("{}").upper()
    if not key:
        raise KeyError(f"no comment {identifier!r}")
    package = document.package
    found = []
    for slide in document.slides:
        part = _comments_part(package, slide.part_path)
        if part is None:
            continue
        for thread in findall(package.tree(part), "p188:cm"):
            for node in [thread, *findall(thread, "p188:replyLst/p188:reply")]:
                guid = (node.get("id") or "").strip("{}").upper()
                if guid == key or guid[:8] == key:
                    found.append((part, node, thread, slide))
    if not found:
        known = [c.address for c in read_comments(document)]
        raise KeyError(f"no comment {identifier!r}; the deck has "
                       f"{', '.join(known[:20]) or 'none'}")
    if len(found) > 1:
        raise KeyError(f"{identifier!r} names {len(found)} comments: give the full GUID")
    return found[0]


class CommentOps:
    """The comment methods of :class:`~pptx_agent.Document`."""

    def comments(self: "Document", slide: "Slide | int | str | None" = None) -> list[Comment]:
        """The threads on a slide, or on every slide, in document order.

        For example::

            [c.text for c in deck.comments() if not c.done]
        """
        slides = None if slide is None else [self._as_slide(slide, "comments")]
        return read_comments(self, slides)

    def comment(self: "Document", identifier: str) -> Comment:
        """The thread an address (``c:3F2A9C1B``) or GUID names (a reply's: its thread).

        For example::

            deck.comment("c:3F2A9C1B").replies
        """
        _, _, thread, slide = _locate(self, identifier)
        guid = thread.get("id")
        return next(c for c in read_comments(self, [slide]) if c.id == guid)

    def add_comment(self: "Document", target: "str | Shape | Slide", text: str, *,
                    author: str, initials: str | None = None,
                    date: "_dt.datetime | None" = None) -> Comment:
        """A new thread on a slide (``s:256``) or a shape (``256.3``), by ``author`` at
        ``date`` (now, UTC, by default).  ``text`` may hold ``\\n`` between paragraphs.

        For example::

            deck.add_comment("s:257", "Add the source of these numbers.", author="Dana")
        """
        from .document import Shape, Slide

        if not str(text).strip():
            raise ValueError("a comment needs text")
        if not str(author or "").strip():
            raise ValueError("a comment needs an author")
        if isinstance(target, Slide):
            slide, shape = target, None
        elif isinstance(target, Shape):
            slide, shape = target._slide, target
        elif str(target).startswith("s:") or str(target).isdigit():
            slide, shape = self.slide(str(target) if str(target).startswith("s:")
                                      else f"s:{target}"), None
        else:
            shape = self.shape(str(target))
            slide = shape._slide
        package = self.package
        from .authoring import default_language

        language = default_language(package)
        created = _stamp(date)
        with self.batch():
            self.history.checkpoint()
            author_id = _ensure_author(package, author, initials)
            part = _comments_part(package, slide.part_path) or \
                _new_comments_part(package, slide.part_path, slide.slide_id)
            identifier = _new_id(self, f"cm:{slide.slide_id}:{shape.id if shape else ''}:"
                                       f"{author}:{created}:{text}")
            thread = make("p188:cm", id=identifier, authorId=author_id, created=created)
            slide_root = package.tree(slide.part_path)
            creation = _creation_id(find(slide_root, "p:extLst"), _SLIDE_CREATION_URI,
                                    "p14:creationId", "val") or "0"
            if shape is None:
                anchor = _declared("pc:sldMkLst", "pc")
                anchor.append(make("pc:docMk"))
                anchor.append(make("pc:sldMk", cId=creation, sldId=str(slide.slide_id)))
            else:
                anchor = _declared("ac:deMkLst", "ac")
                anchor.append(_declared("pc:docMk", "pc"))
                anchor.append(_declared("pc:sldMk", "pc", cId=creation,
                                        sldId=str(slide.slide_id)))
                marker = make(_moniker_of(shape), id=_cnv_id(shape))
                shape_creation = _creation_id(shape._element, _SHAPE_CREATION_URI,
                                              "a16:creationId", "id")
                # Measured: PowerPoint writes the nil GUID for a shape without a creation id
                # (and adds none to the shape).
                marker.set("creationId", shape_creation or _NIL_GUID)
                anchor.append(marker)
            thread.append(anchor)
            thread.append(_body(text, language))
            root = package.tree(part)
            # Measured: threads in order of created.
            later = [node for node in findall(root, "p188:cm")
                     if (node.get("created") or "") > created]
            if later:
                later[0].addprevious(thread)
            else:
                root.append(thread)
            package.mark_dirty(part)
            self._reset_caches()
        return self.comment(identifier)

    def reply_to_comment(self: "Document", identifier: str, text: str, *, author: str,
                         initials: str | None = None,
                         date: "_dt.datetime | None" = None) -> Reply:
        """A reply at the end of the thread (a reply's address answers its thread).

        For example::

            deck.reply_to_comment("c:3F2A9C1B", "Fixed.", author="Sam Author")
        """
        if not str(text).strip():
            raise ValueError("a reply needs text")
        if not str(author or "").strip():
            raise ValueError("a reply needs an author")
        package = self.package
        from .authoring import default_language

        language = default_language(package)
        created = _stamp(date)
        with self.batch():
            self.history.checkpoint()
            part, _, thread, _ = _locate(self, identifier)
            author_id = _ensure_author(package, author, initials)
            reply_id = _new_id(self, f"reply:{thread.get('id')}:{author}:{created}:{text}")
            replies = find(thread, "p188:replyLst")
            if replies is None:
                replies = make("p188:replyLst")
                body = find(thread, "p188:txBody")
                if body is not None:
                    body.addprevious(replies)
                else:
                    thread.append(replies)
            reply = make("p188:reply", id=reply_id, authorId=author_id, created=created)
            reply.append(_body(text, language))
            replies.append(reply)
            package.mark_dirty(part)
            self._reset_caches()
        return next(r for r in self.comment(reply_id).replies if r.id == reply_id)

    def resolve_comment(self: "Document", identifier: str) -> Comment:
        """Mark a thread resolved (``status="resolved"``).

        For example::

            deck.resolve_comment("c:3F2A9C1B").done      # True
        """
        return self._set_status(identifier, "resolved")

    def reopen_comment(self: "Document", identifier: str) -> Comment:
        """Open a resolved thread again.

        For example::

            deck.reopen_comment("c:3F2A9C1B")
        """
        return self._set_status(identifier, None)

    def _set_status(self: "Document", identifier: str, status: str | None) -> Comment:
        with self.batch():
            self.history.checkpoint()
            part, _, thread, _ = _locate(self, identifier)
            if status is None:
                thread.attrib.pop("status", None)
            else:
                # Measured attribute order: id, authorId, status, created.
                attributes = dict(thread.attrib)
                thread.attrib.clear()
                for key in ("id", "authorId"):
                    if key in attributes:
                        thread.set(key, attributes.pop(key))
                thread.set("status", status)
                attributes.pop("status", None)
                for key, value in attributes.items():
                    thread.set(key, value)
            self.package.mark_dirty(part)
            self._reset_caches()
        return self.comment(thread.get("id"))

    def edit_comment(self: "Document", identifier: str, text: str) -> Comment:
        """Replace a comment's or a reply's text.

        For example::

            deck.edit_comment("c:3F2A9C1B", "Is this the audited figure?")
        """
        if not str(text).strip():
            raise ValueError("a comment needs text")
        from .authoring import default_language

        language = default_language(self.package)
        with self.batch():
            self.history.checkpoint()
            part, node, thread, _ = _locate(self, identifier)
            old = find(node, "p188:txBody")
            new = _body(text, language)
            if old is not None:
                old.addprevious(new)
                node.remove(old)
            else:
                node.append(new)
            self.package.mark_dirty(part)
            self._reset_caches()
        return self.comment(thread.get("id"))

    def delete_comment(self: "Document", identifier: str) -> None:
        """Delete a thread with its replies, or one reply.  A slide left without comments
        loses its comments part.

        For example::

            deck.delete_comment("c:3F2A9C1B")
        """
        package = self.package
        with self.batch():
            self.history.checkpoint()
            part, node, thread, slide = _locate(self, identifier)
            if node is thread:
                thread.getparent().remove(thread)
            else:
                replies = node.getparent()
                replies.remove(node)
                if not len(replies):
                    thread.remove(replies)
            package.mark_dirty(part)
            if not findall(package.tree(part), "p188:cm"):
                drop_comments(package, slide.part_path)
            self._reset_caches()


__all__ = ["AC", "COMMENT_REL_EXT_URI", "CT_AUTHORS", "CT_COMMENTS", "Comment", "CommentOps",
           "MONIKERS", "P188", "PC", "REL_AUTHORS", "REL_COMMENTS", "Reply", "drop_comments",
           "initials_of", "read_comments"]
