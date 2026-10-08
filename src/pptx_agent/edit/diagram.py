"""SmartArt: the text of a diagram's nodes, and adding or removing nodes.

The editing itself is :mod:`ooxml_edit.charts.diagram`, where a diagram's data model and
its cached drawing are edited for any document that holds one.  A slide's diagram is reached
through the same :func:`~pptx_agent.edit.chart.graphic_host` as a chart, and keeps the
policy measured on PowerPoint for Mac 16 (see ROADMAP.md, Phase E4): PowerPoint lays every
diagram out again from its data model when it opens a deck, so a cached drawing that cannot
be kept exactly in step is dropped, and PowerPoint regenerates it.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Callable

from ..oxml import xml as _oxml  # noqa: F401  (the slide's vocabulary is registered first)
from ooxml_edit.charts import diagram as _diagram
from ooxml_edit.charts.diagram import (  # noqa: F401  (re-exported)
    DATA_MODEL_EXT_URI,
    DGM_NS,
    DSP_NS,
    REL_DIAGRAM_DRAWING,
    TEXT_TYPES,
    DiagramNode,
    diagram_model,
    new_model_id,
)

from .chart import graphic_host

if TYPE_CHECKING:  # pragma: no cover
    from .document import Shape


class Diagram(_diagram.Diagram):
    """The SmartArt in a graphic frame.  Re-resolved from the document on every call::

        diagram = deck.shape("256.3").diagram
        diagram.texts                         # ['Plan', 'Build', 'Ship']
        diagram.set_text(1, "Build it")       # or diagram.nodes[1].text = "Build it"
        diagram.add_node("Measure").add_child("Weekly")
        diagram.remove_node(0)

    Nodes are listed depth first, in order (``diagram.nodes[0]`` is the first top-level
    node), and addressed by their ``modelId``, which never changes.  A
    :class:`DiagramNode` has ``text`` (settable), ``level``, ``parent``, ``children``,
    ``add_child(text)`` and ``remove()``; ``deck.resolve("256.3/node1")`` is node 1.
    """

    def __init__(self, resolve: Callable[[], "Shape"]) -> None:
        super().__init__(lambda: graphic_host(resolve()), on_inexact_drawing="drop")

    @property
    def nodes(self) -> list[DiagramNode]:
        """Every node, depth first::

            [(node.text, node.level) for node in diagram.nodes]
        """
        return super().nodes

    @property
    def texts(self) -> list[str]:
        """Each node's text, depth first::

            diagram.texts                     # ['Plan', 'Build', 'Ship']
        """
        return super().texts

    def node(self, which) -> DiagramNode:
        """A node by position (in :attr:`nodes`) or by ``modelId``::

            diagram.node(0).text
        """
        return super().node(which)

    def set_text(self, which, text: str) -> "Diagram":
        """Set a node's text -- by position, ``modelId`` or node -- in the data model and
        the cached drawing::

            diagram.set_text(1, "Build it")
        """
        return super().set_text(which, text)

    def add_node(self, text: str = "", *, parent=None, index: int | None = None) -> DiagramNode:
        """A new node under ``parent`` (default: the top level), at ``index`` among its
        siblings (default: last), formatted like a sibling.  PowerPoint lays the diagram
        out again when it opens the deck::

            diagram.add_node("Measure")
        """
        return super().add_node(text, parent=parent, index=index)

    def remove_node(self, which) -> "Diagram":
        """Delete a node and everything under it::

            diagram.remove_node(0)
        """
        return super().remove_node(which)


__all__ = ["Diagram", "DiagramNode", "diagram_model"]
