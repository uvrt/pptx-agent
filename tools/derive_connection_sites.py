#!/usr/bin/env python3
"""Compile every preset's adjustments and connection sites out of ECMA-376's
``presetShapeDefinitions.xml``.

A preset's definition lists where connectors may attach -- ``a:cxnLst``, one ``a:cxn`` per
site, with a position (two guide names or literals) and the angle a connector leaves it at.
``stCxn``/``endCxn@idx`` count into this list, so placing a connector's ends needs it.
ooxml-common compiles the same file's adjustments, guides and paths (for drawing) but not
its sites, and pptx-agent edits without depending on it, so this compiles the sites here,
each with only the guides its positions and angles need.

    python3 tools/derive_connection_sites.py --source presetShapeDefinitions.xml
    python3 tools/derive_connection_sites.py --check --source ...   # exit 1 if stale

**The source is not redistributed here**, for the reason pptx2svg's
``tools/derive_preset_geometry.py`` gives (and where it says to obtain it): ECMA-376 Part 1,
5th edition, ``OfficeOpenXML-DrawingMLGeometries.zip`` -> ``presetShapeDefinitions.xml``.
``--source`` takes the inner ``.zip`` or the extracted ``.xml``; its SHA-256 is verified
before anything is read.  ``upArrow``, which the file does not define, mirrors
``downArrow`` top to bottom, guide for guide (``SUPPLEMENTS``), as pptx2svg's tool does.
Every preset is listed, for its adjustments; the connectors and ``chartPlus``,
``chartStar``, ``chartX`` and ``funnel`` have no sites.
"""

from __future__ import annotations

import argparse
import hashlib
import io
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE.parent / "src" / "pptx_agent" / "edit" / "preset_sites.py"
DRAWINGML_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
SOURCE_SHA256 = "2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939"
A = "{%s}" % DRAWINGML_NS

SUPPLEMENTS = {
    "upArrow": f"""<upArrow xmlns:a="{DRAWINGML_NS}">
      <a:avLst><a:gd name="adj1" fmla="val 50000"/><a:gd name="adj2" fmla="val 50000"/></a:avLst>
      <a:gdLst>
        <a:gd name="maxAdj2" fmla="*/ 100000 h ss"/>
        <a:gd name="a1" fmla="pin 0 adj1 100000"/>
        <a:gd name="a2" fmla="pin 0 adj2 maxAdj2"/>
        <a:gd name="dy2" fmla="*/ ss a2 100000"/>
        <a:gd name="y2" fmla="+- t dy2 0"/>
      </a:gdLst>
      <a:cxnLst>
        <a:cxn ang="3cd4"><a:pos x="hc" y="t"/></a:cxn>
        <a:cxn ang="cd2"><a:pos x="l" y="y2"/></a:cxn>
        <a:cxn ang="cd4"><a:pos x="hc" y="b"/></a:cxn>
        <a:cxn ang="0"><a:pos x="r" y="y2"/></a:cxn>
      </a:cxnLst>
    </upArrow>""",
}

HEADER = '''"""Every ECMA-376 preset's adjustments, and its connection sites (``a:cxnLst``) with the
guides they need.

**Generated file -- do not edit.**  Regenerate with::

    python3 tools/derive_connection_sites.py --source presetShapeDefinitions.xml

Source: ECMA-376 Part 1, 5th edition (December 2016), electronic addendum
``OfficeOpenXML-DrawingMLGeometries.zip`` -> ``presetShapeDefinitions.xml``
SHA-256 ``{sha}``, plus ``upArrow`` (``SUPPLEMENTS`` in the tool).

Each entry is ``name -> (adjustments, guides, sites)``: the ``a:avLst`` defaults as
``(name, value)``, the ``a:gdLst`` guides the sites depend on as ``(name, formula)`` in
evaluation order, and one ``(angle, x, y)`` per site in ``idx`` order (none for a preset
without a ``cxnLst``), each a guide name or a literal.  :mod:`pptx_agent.edit.presets`
evaluates them.
"""

# fmt: off
PRESET_SITES: dict = {{
'''


def read_source(path: Path) -> bytes:
    data = path.read_bytes()
    if path.suffix == ".zip":
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            data = archive.read("presetShapeDefinitions.xml")
    found = hashlib.sha256(data).hexdigest()
    if found != SOURCE_SHA256:
        raise SystemExit(f"presetShapeDefinitions.xml has SHA-256 {found}, expected {SOURCE_SHA256}")
    return data


def _value(fmla: str) -> int:
    match = re.fullmatch(r"val (-?\d+)", fmla.strip())
    if match is None:
        raise SystemExit(f"an adjustment default is not a literal: {fmla!r}")
    return int(match.group(1))


def compile_shape(shape: ET.Element) -> tuple:
    sites_node = shape.find(f"{A}cxnLst")
    adjustments = tuple((gd.get("name"), _value(gd.get("fmla")))
                        for gd in shape.findall(f"{A}avLst/{A}gd"))
    guides = [(gd.get("name"), " ".join(gd.get("fmla").split()))
              for gd in shape.findall(f"{A}gdLst/{A}gd")]
    sites = []
    for cxn in [] if sites_node is None else sites_node.findall(f"{A}cxn"):
        pos = cxn.find(f"{A}pos")
        sites.append((cxn.get("ang"), pos.get("x"), pos.get("y")))
    # Keep only the guides a site reaches, through any chain of guides.
    defined = {name: formula for name, formula in guides}
    needed: set[str] = set()
    pending = [token for site in sites for token in site]
    while pending:
        token = pending.pop()
        if token in defined and token not in needed:
            needed.add(token)
            pending.extend(defined[token].split()[1:])
    kept = tuple((name, formula) for name, formula in guides if name in needed)
    return adjustments, kept, tuple(sites)


def compile_sites(data: bytes) -> dict[str, tuple]:
    root = ET.fromstring(data)
    out: dict[str, tuple] = {}
    for shape in root:
        name = shape.tag.split("}")[-1]
        if name in out:
            continue  # upDownArrow is defined twice, identically
        out[name] = compile_shape(shape)
    for name, text in SUPPLEMENTS.items():
        if name in out:
            raise SystemExit(f"{name} is in the specification now; drop it from SUPPLEMENTS")
        out[name] = compile_shape(ET.fromstring(text))
    return dict(sorted(out.items()))


LINE_LIMIT = 96


def _render_tuple(items: tuple, indent: int) -> str:
    """``items`` as a tuple literal, one line when it fits, else packed lines below."""
    pad = " " * indent
    one_line = f"{pad}{items!r},\n"
    if len(one_line) <= LINE_LIMIT + 1:
        return one_line
    out, line = [f"{pad}(\n"], pad + "    "
    for item in items:
        piece = f"{item!r}, "
        if len(line) + len(piece.rstrip()) > LINE_LIMIT and line.strip():
            out.append(line.rstrip() + "\n")
            line = pad + "    "
        line += piece
    out.append(line.rstrip() + "\n")
    out.append(f"{pad}),\n")
    return "".join(out)


def render(sites: dict) -> str:
    lines = [HEADER.format(sha=SOURCE_SHA256)]
    for name, (adjustments, guides, points) in sites.items():
        lines.append(f"    {name!r}: (\n")
        for part in (adjustments, guides, points):
            lines.append(_render_tuple(part, 8))
        lines.append("    ),\n")
    lines.append("}\n# fmt: on\n")
    return "".join(lines).replace("'", '"')


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    text = render(compile_sites(read_source(args.source)))
    if args.check:
        if TARGET.read_text() != text:
            print(f"{TARGET} is stale", file=sys.stderr)
            return 1
        return 0
    TARGET.write_text(text)
    print(f"wrote {TARGET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
