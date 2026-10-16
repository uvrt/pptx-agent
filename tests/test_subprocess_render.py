"""Rendering in a daemonic process (a Celery prefork worker's child): multiprocessing may not
start workers there, so ooxml-edit's pool renders in subprocess workers, which run only
registered tasks -- the deck's renderer is one."""

from __future__ import annotations

import logging
import multiprocessing
import os

import pytest

from ooxml_edit.tools import SubprocessPool, Toolbox
from ooxml_edit.tools.worker import task_reference

from pptx_agent import Document
from pptx_agent.tools import FORMAT, GROUPS, TOOLS
from pptx_agent.tools.render import render_slides, render_slides_reporting

pytest.importorskip("pptx2svg")
pytest.importorskip("resvg_py")


def _deck() -> bytes:
    deck = Document.new()
    deck.add_slide("Title Only").title = "Rendered elsewhere"
    return deck.to_bytes()


def test_the_renderers_are_registered_tasks():
    assert task_reference(render_slides_reporting) == (
        "pptx_agent.tools.render:render_slides_reporting")
    assert task_reference(render_slides) == "pptx_agent.tools.render:render_slides"


def test_a_subprocess_worker_renders_what_this_process_renders(caplog):
    data = _deck()
    expected = render_slides(data, [1], 320, [])
    with caplog.at_level(logging.WARNING, "ooxml_edit.tools.worker"):
        with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS,
                     runner=SubprocessPool(1), font_dirs=[]) as box:
            session = box.session()
            session.open(data, "deck.pptx")
            result = box.dispatch(session, "render", {"doc": "d1", "slides": [1], "width": 320})
            pids = box.pool.pids()
    assert result.ok, result.error
    assert result.images[0].data == expected[0]
    assert pids and os.getpid() not in pids
    assert not [r for r in caplog.records if "not a registered" in r.getMessage()]


def test_a_subprocess_worker_renders_a_layout_too():
    # A layout's key travels to the worker as a slide number does (round 6, R11).
    data = _deck()
    expected = render_slides(data, ["L2", 1], 320, [])
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS,
                 runner=SubprocessPool(1), font_dirs=[]) as box:
        session = box.session()
        session.open(data, "deck.pptx")
        result = box.dispatch(session, "render", {"doc": "d1", "slides": ["L2", 1],
                                                  "width": 320})
    assert result.ok, result.error
    assert [image.data for image in result.images] == expected


def render_in_a_daemonic_process(queue, data: bytes) -> None:
    try:
        with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS, font_dirs=[]) as box:
            session = box.session()
            session.open(data, "deck.pptx")
            result = box.dispatch(session, "render", {"doc": "d1", "slides": [1], "width": 320})
            queue.put({"ok": result.ok, "runner": box.pool.runner, "pids": box.pool.pids(),
                       "pid": os.getpid(),
                       "image": result.images[0].data if result.ok else None,
                       "error": result.error.message if result.error else None})
    except BaseException as exc:  # noqa: BLE001 -- reported to the test
        queue.put({"ok": False, "error": f"{type(exc).__name__}: {exc}"})


def test_a_daemonic_process_renders_in_subprocess_workers():
    data = _deck()
    context = multiprocessing.get_context("spawn")
    queue = context.Queue()
    process = context.Process(target=render_in_a_daemonic_process, args=(queue, data),
                              daemon=True)
    process.start()
    try:
        got = queue.get(timeout=240)
    finally:
        process.join(30)
    assert got["ok"], got["error"]
    assert got["runner"] == "subprocesses"
    assert got["pids"] and got["pid"] not in got["pids"]
    assert got["image"] == render_slides(data, [1], 320, [])[0]
