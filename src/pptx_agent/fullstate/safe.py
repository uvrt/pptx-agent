"""Parsing untrusted full-state SVG: no entities, no network, bounded sizes.

An SVG handed back to :meth:`~pptx_agent.Document.apply_svg` may have been through an agent,
a browser or a stranger's hands, and it carries XML inside XML (``data-ooxml-xml``, base64).
Both layers go through the same rules:

* **No document type declarations at all.**  A ``<!DOCTYPE`` or ``<!ENTITY`` anywhere is
  refused before the parser sees the bytes, which closes external entities (XXE) and entity
  expansion ("billion laughs") alike.  Nothing this library writes ever needs one.
* **A parser that would not expand them anyway**: lxml with ``resolve_entities=False``,
  ``no_network=True``, ``load_dtd=False`` and ``huge_tree=False`` -- defence in depth should
  the pre-check ever be bypassed, and a check afterwards that no entity reference survived.
* **Size limits** on the whole document, on each base64 attribute (checked before decoding)
  and on each JSON attribute, and a cap on the number of shapes.
* **Strict base64** (``validate=True``): anything outside the alphabet is an error, not
  silently dropped.
"""

from __future__ import annotations

import base64
import binascii
import json
import re
from dataclasses import dataclass
from typing import Any

from lxml import etree

from ..oxml.xml import Element


class FullStateError(ValueError):
    """A full-state SVG could not be applied: malformed, unsafe, stale or not expressible."""


class NotFullStateSvg(FullStateError):
    """The SVG was not written by this library's full-state emitter.

    Importing arbitrary SVG (from Figma, Illustrator...) is out of scope by design: it has no
    theme semantics and no OOXML to fall back on, so any "import" would be a lossy guess.
    """


class UnsafeInput(FullStateError):
    """The input tried something no full-state SVG needs: a DTD, an entity, excessive size."""


@dataclass(frozen=True)
class Limits:
    """Bounds on untrusted input.  The defaults are far above any real slide."""

    #: The whole SVG, in bytes.
    svg_bytes: int = 64 * 1024 * 1024
    #: One ``data-ooxml-xml`` attribute, in base64 characters (checked before decoding).
    xml_attribute: int = 16 * 1024 * 1024
    #: One JSON-valued attribute (text, table, relationships), in characters.
    json_attribute: int = 8 * 1024 * 1024
    #: Shapes (elements carrying ``data-pptx-id`` or ``data-ooxml-xml``) in one SVG.
    shapes: int = 20_000
    #: Values in one chart (categories and series, counted as a grid).
    chart_points: int = 1_000_000
    #: Nodes in one SmartArt diagram.
    diagram_nodes: int = 10_000
    #: Adjust values in one ``data-ooxml-adj`` (a preset has at most eight).
    adjustments: int = 64


DEFAULT_LIMITS = Limits()

_DECLARATION = re.compile(rb"<!\s*(DOCTYPE|ENTITY)", re.IGNORECASE)


def _parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        load_dtd=False,
        dtd_validation=False,
        huge_tree=False,
        remove_blank_text=False,
        collect_ids=False,
    )


def parse_untrusted(data: bytes | str, *, limit: int, what: str) -> Element:
    """Parse ``data`` as XML under the rules above, or raise :class:`UnsafeInput`."""
    if isinstance(data, str):
        data = data.encode("utf-8")
    if len(data) > limit:
        raise UnsafeInput(f"{what} is {len(data)} bytes; the limit is {limit}")
    if _DECLARATION.search(data):
        raise UnsafeInput(f"{what} contains a document type or entity declaration")
    try:
        root = etree.fromstring(data, _parser())
    except etree.XMLSyntaxError as error:
        raise FullStateError(f"{what} is not well-formed XML: {error}") from None
    tree = root.getroottree()
    if tree.docinfo.doctype or tree.docinfo.internalDTD is not None:
        raise UnsafeInput(f"{what} declares a document type")
    for _ in root.iter(etree.Entity):
        raise UnsafeInput(f"{what} contains an entity reference")
    return root


def decode_base64(value: str, *, limit: int, what: str) -> bytes:
    """Strict base64, length-checked before any decoding happens."""
    if len(value) > limit:
        raise UnsafeInput(f"{what} is {len(value)} characters; the limit is {limit}")
    try:
        return base64.b64decode(value.encode("ascii"), validate=True)
    except (binascii.Error, UnicodeEncodeError, ValueError):
        raise FullStateError(f"{what} is not valid base64") from None


def encode_base64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def load_json(value: str, *, limit: int, what: str, kind: type) -> Any:
    """Parse a JSON attribute, length-checked, and insist on the top-level type."""
    if len(value) > limit:
        raise UnsafeInput(f"{what} is {len(value)} characters; the limit is {limit}")
    try:
        loaded = json.loads(value)
    except (ValueError, RecursionError):
        raise FullStateError(f"{what} is not valid JSON") from None
    if not isinstance(loaded, kind):
        raise FullStateError(f"{what} must be a JSON {kind.__name__}")
    return loaded


def dump_json(value: Any) -> str:
    """Compact, deterministic JSON: the same document always emits the same SVG."""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"), sort_keys=False)
