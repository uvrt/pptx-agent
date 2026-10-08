"""Golden transcripts: the trial's PowerPoint tasks p1-p6, the graphics spike's p7, p8, o1
and m1 (positions computed by the caller, in points, since the layout tools went), T4's p9 (a
chart slide from a CSV), p10 (a review in comments), p11 (SmartArt updated in place) and
p12 (a rebrand through the theme), done with the tools alone and replayed.

Each transcript in ``golden/transcripts`` is a reference solution as tool calls -- what a
model would send, with no code and no files -- recorded with what every call returned.
Replaying it on a fresh session (fixed clock) must give the same results, the same output
bytes, and an output that passes the task's own check from the trial (``golden/grading``).
This is the regression test for "the tools can do every trial task without Python".

The inputs are rebuilt by the trial's builder (``golden_inputs.py``) from this repository's
fixtures, and must hash as they did when the transcripts were recorded.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from ooxml_edit.tools import Toolbox

from pptx_agent.tools import FORMAT, GROUPS, TOOLS

import golden_inputs

HERE = Path(__file__).resolve().parent
GOLDEN = HERE / "golden"
FIXTURES = HERE / "fixtures"
CLOCK = dt.datetime(2026, 10, 7, 12, 0, tzinfo=dt.timezone.utc)
TRANSCRIPTS = sorted((GOLDEN / "transcripts").glob("*.json"))


def _expectation(result) -> dict:
    data = {"ok": result.ok}
    if not result.ok:
        data["error"] = result.error.code
        return data
    for key in ("created", "removed"):
        value = getattr(result, key)
        if value:
            data[key] = value
    if result.refs:
        data["refs"] = result.refs
    return data


@pytest.fixture(scope="module")
def inputs(tmp_path_factory):
    root = tmp_path_factory.mktemp("golden-inputs")
    for task in golden_inputs.BUILDERS:
        if task == "p5-rebrand":
            try:
                import PIL  # noqa: F401
            except ImportError:
                continue
        golden_inputs.build(task, root)
    return root


@pytest.mark.parametrize("path", TRANSCRIPTS, ids=lambda path: path.stem)
def test_a_golden_transcript_replays_to_a_passing_check(path, inputs, tmp_path):
    transcript = json.loads(path.read_text(encoding="utf-8"))
    task = transcript["task"]
    with Toolbox(TOOLS, formats=[FORMAT], groups=GROUPS) as toolbox:
        session = toolbox.session(clock=lambda: CLOCK)
        for item in transcript["inputs"]:
            source = (FIXTURES if item.get("root") == "fixtures" else inputs / task) / item["file"]
            if not source.exists():
                pytest.skip(f"{source.name} was not built (Pillow missing?)")
            data = source.read_bytes()
            assert hashlib.sha256(data).hexdigest() == item["sha256"], \
                f"{item['file']} is not the input the transcript was recorded with"
            if item["as"] == "document":
                session.open(data, source.name)
            else:
                session.add_blob(data, source.name)
        for index, step in enumerate(transcript["calls"]):
            result = toolbox.dispatch(session, step["tool"], step["arguments"])
            assert _expectation(result) == step["expect"], (index, step["tool"], result.to_json())
        (output,) = session.take_outputs()
    assert output.name == transcript["output"]["name"]
    assert hashlib.sha256(output.data).hexdigest() == transcript["output"]["sha256"]

    deck = tmp_path / output.name
    deck.write_bytes(output.data)
    check = GOLDEN / "grading" / f"check_{task.split('-')[0]}.py"
    proc = subprocess.run([sys.executable, str(check), str(deck)], capture_output=True,
                          text=True, env={**os.environ, "GOLDEN_INPUTS": str(inputs)},
                          timeout=600)
    report = json.loads(proc.stdout)
    failed = [r for r in report["results"] if not r["ok"]]
    assert not failed, failed
    assert report["passed"] == report["total"] == transcript["check"]["total"]


def test_the_goldens_cover_every_task_in_few_calls():
    tasks = {json.loads(p.read_text())["task"].split("-")[0]: json.loads(p.read_text())
             for p in TRANSCRIPTS}
    assert sorted(tasks) == ["m1", "o1", "p1", "p10", "p11", "p12", "p2", "p3", "p4", "p5",
                             "p6", "p7", "p8", "p9"]
    for transcript in tasks.values():
        assert len(transcript["calls"]) <= 10
        assert all(step["tool"] in {tool.name for tool in TOOLS} for step in transcript["calls"])
