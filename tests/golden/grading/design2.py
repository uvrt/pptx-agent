"""Shared checks for the spike's new tasks (o1, m1): public pptx_agent API only.

Like design.common_checks, but collisions come from deck.overflows() (which knows layering:
a label on a bubble, text on a panel), and the margin test uses drawn_bounds (a rotated axis
title, a connector's route)."""
from __future__ import annotations

from common import norm
from design import (IN, SAFE, EPS, is_theme, is_title, leaf_shapes, run_sizes, shape_colors,
                    text_shapes, unchanged, ntext)


def dbox(sh):
    try:
        l, t, w, h = sh.drawn_bounds
    except Exception:  # noqa: BLE001
        l, t, w, h = sh.slide_bounds
    return l, t, l + w, t + h


def sbox(sh):
    l, t, w, h = sh.slide_bounds
    return l, t, l + w, t + h


def centre(b):
    return (b[0] + b[2]) / 2, (b[1] + b[3]) / 2


def gap(a, b):
    """Distance between two boxes (0 when they touch or overlap)."""
    dx = max(0, max(a[0], b[0]) - min(a[2], b[2]))
    dy = max(0, max(a[1], b[1]) - min(a[3], b[3]))
    return (dx * dx + dy * dy) ** 0.5


def spread(values):
    return max(values) - min(values) if values else 0


def common_checks2(c, deck, src, slide_no: int, others, old_title: str):
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    c.check("still 3 slides", len(deck.slides) == 3, len(deck.slides))
    c.run("other slides unchanged", lambda: unchanged(deck, src, others))
    slide = deck.slides[slide_no - 1]
    titles = [sh for sh in slide.shapes if is_title(sh)]
    title = titles[0] if titles else None
    t = norm(title.text) if title is not None else ""
    c.check("title placeholder holds an action title (>= 8 words, changed)",
            title is not None and t != old_title and len(t.split()) >= 8, t)
    c.run("title fits its placeholder", lambda: (title is not None and not title.text_fit().overflows,
                                                 title and title.text_fit()))
    ov = deck.overflows()
    c.check("deck.overflows() empty (text, overlaps, lines through text, off-slide)", not ov,
            [str(o)[:160] for o in ov])
    W, H = deck.slide_size
    leaves = leaf_shapes(slide)
    outside = [(sh.id, sh.name, dbox(sh)) for sh in leaves
               if dbox(sh)[0] < 0 or dbox(sh)[1] < 0 or dbox(sh)[2] > W or dbox(sh)[3] > H]
    c.check("every shape inside the slide (as drawn)", not outside, outside)
    unsafe = []
    for sh in leaves:
        if is_title(sh):
            continue
        b = dbox(sh)
        if b[0] < SAFE[0] - EPS or b[1] < SAFE[1] - EPS or b[2] > SAFE[2] + EPS or b[3] > SAFE[3] + EPS:
            unsafe.append((sh.id, sh.name, norm(sh.text or "")[:30], [round(v / IN, 2) for v in b]))
    c.check("everything added inside the margins (0.4-12.933 in, 1.4-7.0 in)", not unsafe, unsafe)
    small = [(sh.id, norm(sh.text)[:25], s) for sh in text_shapes(slide) for s in run_sizes(sh) if s < 8]
    c.check("no text under 8 pt (effective_size)", not small, small)
    rgb = [(sh.id, sh.name, k, v) for sh in leaves for k, v in shape_colors(sh) if not is_theme(v)]
    c.check("theme colours only (no hard-coded RGB)", not rgb, rgb)
    return slide


def every_text(slide) -> str:
    return " ".join(ntext(sh.text) for sh in text_shapes(slide))
