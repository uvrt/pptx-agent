"""``pptx_agent.core`` is ooxml-edit, under its old name.

The format-neutral core moved to the ooxml-edit package, which keeps its own neutrality test.
What is left here is a shim, and it must be the same modules -- not copies -- or the
namespace and child-order registries the PresentationML layer fills would be two.
"""

from __future__ import annotations

import ast
from pathlib import Path

import ooxml_edit.history
import ooxml_edit.opc
import ooxml_edit.stamp
import ooxml_edit.xml

SRC = Path(__file__).parents[1] / "src" / "pptx_agent"


def test_the_old_names_are_the_same_modules():
    import pptx_agent.core
    from pptx_agent.core import history, opc, stamp, xml

    assert (history, opc, stamp, xml) == (
        ooxml_edit.history, ooxml_edit.opc, ooxml_edit.stamp, ooxml_edit.xml)
    import pptx_agent.core.opc as dotted

    assert dotted is ooxml_edit.opc
    assert pptx_agent.core.xml.NAMESPACES is ooxml_edit.xml.NAMESPACES


def test_the_presentation_layer_registers_with_ooxml_edit():
    from pptx_agent.oxml.package import PresentationPackage

    assert issubclass(PresentationPackage, ooxml_edit.opc.OpcPackage)
    assert ooxml_edit.xml.NAMESPACES["p"].endswith("/presentationml/2006/main")
    assert "p:spTree" in ooxml_edit.xml.CHILD_ORDER


def test_the_package_itself_imports_ooxml_edit_not_the_shim():
    for module in sorted(SRC.rglob("*.py")):
        if module.parent.name == "core":
            continue
        for node in ast.walk(ast.parse(module.read_text(encoding="utf-8"))):
            if isinstance(node, ast.ImportFrom):
                name = ("." * node.level) + (node.module or "")
                assert "core" not in name.split("."), f"{module.relative_to(SRC)}: {name}"
