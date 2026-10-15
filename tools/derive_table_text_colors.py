#!/usr/bin/env python3
"""Compile the text colour of each region of PowerPoint's built-in table styles out of
pptx2svg's measured catalogue (``pptx2svg.parse.table_styles_builtin``).

PowerPoint does not write a built-in style's definition into ``ppt/tableStyles.xml``, not
even for the style a table in the deck uses, so a cell's inherited text colour cannot be
read from the deck.  pptx2svg measured all 74 built-in styles from PowerPoint 16 itself
(its ``tools/derive_table_styles.py`` and ``tools/derive_themed_table_styles.py``); this
copies the one property ``Run.effective_color`` needs -- each region's ``a:tcTxStyle``
colour -- so pptx-agent reads it without importing pptx2svg at run time.

    python3 tools/derive_table_text_colors.py            # needs pptx2svg installed
    python3 tools/derive_table_text_colors.py --check    # exit 1 if stale
"""

from __future__ import annotations

import argparse
import sys
import textwrap
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE.parent / "src" / "pptx_agent" / "edit" / "table_text_colors.py"

HEADER = '''"""The text colour of each region of PowerPoint's built-in table styles.

**Generated file -- do not edit.**  Regenerate with::

    python3 tools/derive_table_text_colors.py

Source: pptx2svg's catalogue of the 74 built-in styles, measured from PowerPoint 16 (its
``tools/derive_table_styles.py`` and ``tools/derive_themed_table_styles.py``).  PowerPoint
does not write these definitions into a deck's ``tableStyles.xml``.

Each entry is ``GUID -> {region: colour}``: the region an ``a:tcTxStyle`` sits in
(``wholeTbl``, ``firstRow``... ECMA-376 Part 1, 20.1.4.2) and its colour as
:meth:`Color.parse <pptx_agent.Color.parse>` reads it.
"""

from __future__ import annotations

BUILTIN_TEXT_COLORS: dict[str, dict[str, str]] = {
'''


def _spec(color) -> str:
    parts = [color.scheme]
    for transform in getattr(color, "transforms", None) or []:
        parts.append(f"{transform.kind}={transform.value}")
    return " ".join(parts)


def generate() -> str:
    from pptx2svg.parse.source import TABLE_STYLE_REGIONS
    from pptx2svg.parse.table_styles_builtin import BUILTIN_TABLE_STYLES, builtin_table_style

    lines = [HEADER]
    for guid in BUILTIN_TABLE_STYLES:
        style = builtin_table_style(guid)
        regions = {}
        for element, attribute in TABLE_STYLE_REGIONS:
            region = getattr(style, attribute)
            color = region.text.color if region is not None and region.text is not None else None
            if color is not None:
                regions[element] = _spec(color)
        body = ", ".join(f'"{name}": "{spec}"' for name, spec in regions.items())
        entry = f'    "{guid.upper()}": {{{body}}},\n'
        if len(entry) > 100:
            wrapped = textwrap.fill(body, 91, initial_indent=" " * 8, subsequent_indent=" " * 8)
            entry = f'    "{guid.upper()}": {{\n{wrapped}}},\n'
        lines.append(f"    # {style.name}\n{entry}")
    lines.append("}\n")
    return "".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("--check", action="store_true", help="exit 1 if the module is stale")
    arguments = parser.parse_args()
    text = generate()
    if arguments.check:
        current = TARGET.read_text(encoding="utf-8") if TARGET.exists() else ""
        if current != text:
            print(f"{TARGET} is stale; run {Path(__file__).name}", file=sys.stderr)
            return 1
        return 0
    TARGET.write_text(text, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
