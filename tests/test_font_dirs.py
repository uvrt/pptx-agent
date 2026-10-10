"""The application's own font folders through the tool layer: ``Toolbox(font_dirs=...)``,
a session's own, or ``OOXML_FONT_DIRS`` -- measured with (fit, overflow) and drawn with,
in the worker process too.

Production: an application keeps its licensed faces in a folder the system does not
search; the tools rendered and measured as though the face were missing.  The face here is
an open one -- the ``pptx2svg-fonts`` bundle's Cousine, relabelled in a temporary folder as
a family nothing else answers to, so no system copy can stand in for it.
"""

from __future__ import annotations

import json
import multiprocessing
import os
import sys
import warnings
from pathlib import Path

import pytest

from ooxml_common.fonts import bundle_dir
from ooxml_common.fonts.office import FONT_DIRS_ENV
from ooxml_common.fonts.sfnt import relabel
from ooxml_edit.tools import Toolbox

from pptx_agent import Document, ParagraphSpec, RunSpec, TextSpec
from pptx_agent.tools import FORMAT, GROUPS, TOOLS

pptx2svg = pytest.importorskip("pptx2svg")
BUNDLE = bundle_dir()
FAMILY = "Fontdirs Probe Mono"
TEXT = "The quick brown fox jumps over the lazy dog, again and again and again."

pytestmark = pytest.mark.skipif(BUNDLE is None, reason="needs pptx2svg-fonts' files")


@pytest.fixture(autouse=True)
def _no_environment(monkeypatch):
    monkeypatch.delenv(FONT_DIRS_ENV, raising=False)


@pytest.fixture
def folder(tmp_path) -> Path:
    target = tmp_path / "app-fonts" / "probe"
    target.mkdir(parents=True)
    data = relabel((BUNDLE / "Cousine-Regular.ttf").read_bytes(), FAMILY, bold=False, italic=False)
    (target / "FontdirsProbeMono-Regular.ttf").write_bytes(data)
    return tmp_path / "app-fonts"


@pytest.fixture
def deck_bytes() -> bytes:
    """One slide, one text box in the probe's face, as tall as its text measures without
    the face (it fits) -- and too short for it measured as the face (it overflows)."""
    deck = Document.new()
    slide = deck.add_slide()
    box = slide.add_textbox(457200, 457200, 3200400, 457200, autofit="none")
    box.set_text(TextSpec([ParagraphSpec([RunSpec(TEXT, font=FAMILY, size=18)])]))
    box.height = box.fit_height()
    assert not deck.overflows()
    return deck.to_bytes()


def test_the_document_measures_with_its_folder(deck_bytes, folder, monkeypatch):
    deck = Document.open(deck_bytes)
    assert not deck.overflows()
    deck.font_dirs = (str(folder),)
    assert deck.overflows(), "measured as the (wider) monospaced face, the text overflows"
    deck.font_dirs = None
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    assert Document.open(deck_bytes).overflows()
    deck.font_dirs = ()
    assert not deck.overflows()


def _box(**options) -> Toolbox:
    return Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, **options)


def _check_and_render(box: Toolbox, deck_bytes: bytes, **session_options):
    session = box.session(**session_options)
    session.open(deck_bytes, "deck.pptx")
    check = box.dispatch(session, "check", {"doc": "d1", "include": ["fit"]})
    render = box.dispatch(session, "render", {"doc": "d1", "slides": [1], "width": 640})
    assert check.ok and render.ok, (check.error, render.error)
    return check.data.get("overflows") or [], render.images[0].data


def test_the_toolbox_reaches_the_checks_and_the_worker(deck_bytes, folder, monkeypatch):
    direct = Document.open(deck_bytes)
    direct.font_dirs = (str(folder),)
    drawn = direct.render_png([1], width=640)[0]

    with _box(workers=1) as plain:
        overflows, unconfigured = _check_and_render(plain, deck_bytes)
    assert overflows == [] and unconfigured != drawn

    with _box(workers=1, font_dirs=[folder]) as configured:
        overflows, image = _check_and_render(configured, deck_bytes)
        assert overflows and image == drawn          # rendered in the worker, with the folder
        # A session overrides the toolbox: [] is none (no environment variable either).
        overflows, image = _check_and_render(configured, deck_bytes, font_dirs=[])
        assert overflows == [] and image == unconfigured

    # The environment variable, read where the toolbox runs, reaches the worker too.
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    with _box(workers=1) as from_env:
        overflows, image = _check_and_render(from_env, deck_bytes)
    assert overflows and image == drawn


def test_in_process_rendering_sees_the_folder(deck_bytes, folder):
    direct = Document.open(deck_bytes)
    direct.font_dirs = (str(folder),)
    with _box(workers=0, font_dirs=[folder]) as box:
        _, image = _check_and_render(box, deck_bytes)
    assert image == direct.render_png([1], width=640)[0]


def test_the_definitions_and_prompt_do_not_change(folder):
    with _box() as plain, _box(font_dirs=[folder]) as configured:
        for provider in ("anthropic", "openai-responses"):
            assert json.dumps(plain.definitions(provider)) == json.dumps(configured.definitions(provider))
        assert plain.system_prompt() == configured.system_prompt()


# Python 3.14 made forkserver Linux's default start method (fork before).  The toolbox's
# pool names spawn itself; under each method this platform has, the render and the folders
# handed to it cross to the worker the same.  fork only on Linux: macOS' system libraries
# are not safe to fork with threads running.
_METHODS = [method for method in multiprocessing.get_all_start_methods()
            if method != "fork" or sys.platform.startswith("linux")]


@pytest.mark.parametrize("method", _METHODS)
def test_the_worker_sees_the_folder_under(method, deck_bytes, folder):
    direct = Document.open(deck_bytes)
    direct.font_dirs = (str(folder),)
    drawn = direct.render_png([1], width=640)[0]
    with warnings.catch_warnings():
        # fork() with threads running is a DeprecationWarning since 3.12; this test asks.
        warnings.simplefilter("ignore", DeprecationWarning)
        with _box(workers=1, start_method=method, font_dirs=[folder]) as box:
            overflows, image = _check_and_render(box, deck_bytes)
            assert not box.pool.in_process
    assert overflows and image == drawn
