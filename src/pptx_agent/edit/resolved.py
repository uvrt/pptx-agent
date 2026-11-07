"""What pptx2svg is handed to measure or draw a deck: its parts in memory, and one cache
of resolved slides for the whole process.

Fit facts, the live view, contrast and renders all ask pptx2svg for slides of the deck as
it is now, many times a deck.  Two things made each ask cost a whole deck:

* the deck was written as a zip (:meth:`Document.to_bytes`) and read back -- every part
  serialised, compressed and unpacked -- where pptx2svg needs a handful of parts;
* every slide was parsed, and the asked-for ones resolved, again.

Now pptx2svg reads the parts it needs straight from the open package
(:func:`source_of`, :func:`pptx2svg.open_parts`), and resolved slides are kept, by the
bytes they were resolved from, in one :class:`pptx2svg.ModelCache` per process
(:func:`model_cache`): after an edit, only the slides whose part or whose layout, master,
theme, pictures or charts changed are resolved again.  The same model either way, so the
same facts and pictures; with a pptx2svg without these (before 0.3.8) the zip round trip
is kept.
"""

from __future__ import annotations

import threading
from collections.abc import Iterator, Mapping
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:  # pragma: no cover
    from .document import Document

_lock = threading.Lock()
_cache: Any = None


def model_cache() -> Any:
    """The process's :class:`pptx2svg.ModelCache`, made on first use; ``None`` with a
    pptx2svg that has none.  Bounded (pptx2svg's defaults: 512 resolved slides, 256
    parsed layouts, masters and themes), shared by every deck and thread of the process."""
    global _cache
    import pptx2svg

    factory = getattr(pptx2svg, "ModelCache", None)
    if factory is None:                        # pragma: no cover - pptx2svg before 0.3.8
        return None
    with _lock:
        if _cache is None:
            _cache = factory()
        return _cache


def convert_options(**options: Any) -> Any:
    """``pptx2svg.ConvertOptions(**options)`` with the process's cache."""
    import pptx2svg

    cache = model_cache()
    if cache is not None:
        options.setdefault("cache", cache)
    return pptx2svg.ConvertOptions(**options)


def source_of(document: "Document") -> Any:
    """The deck as pptx2svg reads it -- the bytes :meth:`Document.to_bytes` would give,
    part by part, each serialised when pptx2svg asks for it."""
    import pptx2svg

    opener = getattr(pptx2svg, "open_parts", None)
    if opener is None:                         # pragma: no cover - pptx2svg before 0.3.8
        return document.to_bytes()
    return opener(_Parts(document))


class _Parts(Mapping):
    """``{part name: bytes}`` of an open deck, as its zip would hold them.

    :meth:`Document.to_bytes` writes two parts in place of the package's own
    (:func:`pptx_agent.edit.deck.write_replacements`): ``[Content_Types].xml`` when the
    main part is to be retyped -- never for a deck kept as it is, which is what is drawn --
    and ``docProps/app.xml`` with its slide counts and titles brought up to date.  Of
    ``app.xml`` pptx2svg reads only ``AppVersion``, which that leaves as it is, so the
    package's own is handed over: refreshing it reads the deck as it was opened."""

    def __init__(self, document: "Document") -> None:
        from .deck import wanted_content_type

        package = document.package
        self._package = package
        self._names = frozenset(package.part_names)
        self._read: dict[str, bytes] = {}
        retyped = wanted_content_type(package, template=None, macro=None)
        if retyped is not None:                # pragma: no cover - a kept deck is never
            self._read["[Content_Types].xml"] = package.content_types_with(
                package.presentation_part(), retyped)

    def __getitem__(self, name: str) -> bytes:
        found = self._read.get(name)
        if found is not None:
            return found
        if name not in self._names:
            raise KeyError(name)
        data = self._package.read(name)
        if data is None:
            raise KeyError(name)
        self._read[name] = data
        return data

    def __contains__(self, name: object) -> bool:
        return name in self._names

    def __iter__(self) -> Iterator[str]:
        return iter(self._names)

    def __len__(self) -> int:
        return len(self._names)
