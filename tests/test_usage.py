"""``docs/USAGE.md``: every example runs, so the guide cannot rot.

Each ```python block runs on its own, in a fresh directory holding ``results.pptx`` and
``smartart.pptx`` (copies of two fixtures), with warnings as errors.  A ``print(...)``
whose line ends in ``# -> text`` must print exactly that text.
"""

from __future__ import annotations

import re
import shutil
import warnings
from pathlib import Path

import pytest

from conftest import FIXTURE_DIR

USAGE = Path(__file__).resolve().parent.parent / "docs" / "USAGE.md"
_BLOCK = re.compile(r"^### (.+?)\n.*?```python\n(.*?)```", re.DOTALL | re.MULTILINE)
#: ``print(...)  # -> text``, the comment on the same line or the next.
_CHECKED = re.compile(r"^([ \t]*)print\((.*)\)\s*# -> (.*)$", re.MULTILINE)


def _examples() -> list[tuple[str, str]]:
    return [(title, code) for title, code in _BLOCK.findall(USAGE.read_text(encoding="utf-8"))]


def _checked(code: str) -> str:
    return _CHECKED.sub(lambda m: f"{m.group(1)}_check({m.group(2)}, expected={m.group(3)!r})",
                        code)


def _check(*values, expected: str) -> None:
    printed = " ".join(str(value) for value in values)
    assert printed == expected.strip(), f"printed {printed!r}, the guide says {expected!r}"


def test_the_guide_has_the_common_tasks():
    titles = [title for title, _ in _examples()]
    assert len(titles) == 27, titles
    for wanted in ("Units", "Text frames", "Collisions", "Measure text", "change its colours",
                   "Open a deck", "Find text", "mixed formatting", "Tables", "Charts",
                   "SmartArt", "Fills and outlines", "connectors and groups", "Slides",
                   "insert_outline", "Speaker notes", "text fit", "Layouts", "Theme",
                   "Pictures", "Full-state SVG", "New decks and templates",
                   "Render, validate, save", "Undo", "One text spec", "Agent tools",
                   "Review comments"):
        assert any(wanted in title for title in titles), wanted


@pytest.mark.parametrize("title,code", _examples(), ids=[t for t, _ in _examples()])
def test_the_example_runs(title, code, tmp_path, monkeypatch, capsys):
    if any(name in code for name in ("render_png", "render_svg", "text_fit", "overflows",
                                     "collisions", "measure_text", "fit_height",
                                     "rows_fitting")):
        pytest.importorskip("pptx2svg")
    shutil.copy(FIXTURE_DIR / "real-financial-report.pptx", tmp_path / "results.pptx")
    shutil.copy(FIXTURE_DIR / "powerpoint-smartart.pptx", tmp_path / "smartart.pptx")
    monkeypatch.chdir(tmp_path)
    source = _checked(code)
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        exec(compile(source, f"USAGE.md: {title}", "exec"),
             {"_check": _check, "__name__": "usage"})
    assert source.count("_check(") == len(_CHECKED.findall(code))
