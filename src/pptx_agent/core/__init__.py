"""The format-neutral core, now the separate package ``ooxml_edit`` -- kept here as a shim.

The OPC package, the XML helpers and ordered insertion, undo history and id stamping moved,
with their history, to `ooxml-edit <https://github.com/uvrt/ooxml-edit>`_, which a docx
editor shares.  ``pptx_agent.core.opc``, ``.xml``, ``.history`` and ``.stamp`` are those
modules themselves, not copies: importing either name gives the same module object, so the
namespace and child-order registries a format layer fills are one registry.  New code
imports ``ooxml_edit`` directly.
"""

import sys as _sys

from ooxml_edit import history, opc, stamp, xml

for _module in (history, opc, stamp, xml):
    _sys.modules[f"{__name__}.{_module.__name__.rpartition('.')[2]}"] = _module

__all__ = ["history", "opc", "stamp", "xml"]
