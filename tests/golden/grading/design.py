"""Geometry and style helpers shared by the p7/p8 design checks (public pptx_agent API only)."""
from __future__ import annotations

import re

from common import norm


def all_shapes(slide):
    """Every shape once, groups and their members, de-duplicated by id."""
    out, seen = [], set()
    stack = list(slide.shapes)
    while stack:
        sh = stack.pop(0)
        if sh.id in seen:
            continue
        seen.add(sh.id)
        out.append(sh)
        try:
            stack[0:0] = list(sh.children or [])
        except Exception:  # noqa: BLE001
            pass
    return out

IN = 914400
SAFE = (int(0.4 * IN), int(1.4 * IN), int(12.933 * IN), int(7.0 * IN))  # left, top, right, bottom
EPS = int(0.01 * IN)
THEME = {"accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "tx1", "tx2", "bg1", "bg2", "dk1",
         "dk2", "lt1", "lt2", "hlink", "folHlink", "phClr"}


def ntext(text: str) -> str:
    return norm(text).replace("–", "-").replace("—", "-").replace("‑", "-")


def box(sh):
    l, t, w, h = sh.slide_bounds
    return l, t, l + w, t + h


def leaf_shapes(slide):
    return [sh for sh in all_shapes(slide) if sh.kind != "group"]


def text_shapes(slide):
    return [sh for sh in leaf_shapes(slide) if sh.kind in ("shape", "placeholder") and norm(sh.text or "")]


def is_title(sh) -> bool:
    return bool(sh.placeholder) and sh.placeholder[0] in ("title", "ctrTitle")


def color_name(c) -> str | None:
    if c is None:
        return None
    s = str(c).strip()
    return s or None


def is_theme(c) -> bool:
    s = color_name(c)
    if s is None:
        return True
    return s.split()[0] in THEME


def shape_colors(sh):
    """Every colour the shape states: fill, gradient stops, outline, run colours."""
    out = []
    fill = getattr(sh, "fill", None)
    if fill is not None:
        kind = getattr(fill, "kind", None)
        if kind == "solid":
            out.append(("fill", color_name(fill.color)))
        elif kind == "gradient":
            for stop in getattr(fill, "stops", None) or []:
                col = stop[1] if isinstance(stop, tuple) else getattr(stop, "color", None)
                out.append(("gradient", color_name(col)))
    try:
        ln = sh.line
        if ln is not None and ln.color is not None:
            out.append(("line", color_name(ln.color)))
    except Exception:  # noqa: BLE001
        pass
    if sh.kind in ("shape", "placeholder") and (sh.text or ""):
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.color is not None:
                    out.append(("text", color_name(r.color)))
    return [(k, v) for k, v in out if v]


def run_sizes(sh):
    sizes = []
    for p in sh.text_frame.paragraphs:
        for r in p.runs:
            if (r.text or "").strip():
                sizes.append(r.effective_size)
    return sizes


def overlaps(a, b, tol=EPS) -> bool:
    return min(a[2], b[2]) - max(a[0], b[0]) > tol and min(a[3], b[3]) - max(a[1], b[1]) > tol


def fill_key(sh) -> str:
    fill = getattr(sh, "fill", None)
    if fill is None:
        return "inherited"
    if getattr(fill, "kind", None) == "solid":
        return str(fill.color)
    return str(getattr(fill, "kind", fill))


def unchanged(deck, src, numbers) -> bool:
    return all(deck.to_outline(slides=[i]).split("\n", 1)[1] == src.to_outline(slides=[i]).split("\n", 1)[1]
               for i in numbers)


def common_checks(c, deck, src, slide_no: int, others, old_title: str, extra_ok=lambda a, b: False):
    """The checks p7 and p8 share.  Returns the slide."""
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
    c.check("deck.overflows() empty", not ov, [str(o)[:160] for o in ov])
    W, H = deck.slide_size
    outside = [(sh.id, sh.name, box(sh)) for sh in leaf_shapes(slide)
               if box(sh)[0] < 0 or box(sh)[1] < 0 or box(sh)[2] > W or box(sh)[3] > H]
    c.check("every shape inside the slide", not outside, outside)
    unsafe = []
    for sh in leaf_shapes(slide):
        if is_title(sh):
            continue
        b = box(sh)
        if b[0] < SAFE[0] - EPS or b[1] < SAFE[1] - EPS or b[2] > SAFE[2] + EPS or b[3] > SAFE[3] + EPS:
            unsafe.append((sh.id, sh.name, norm(sh.text or "")[:30], [round(v / IN, 2) for v in b]))
    c.check("everything added inside the margins (0.4-12.933 in, 1.4-7.0 in)", not unsafe, unsafe)
    tb = [(sh, box(sh)) for sh in text_shapes(slide)]
    hits = []
    for i in range(len(tb)):
        for j in range(i + 1, len(tb)):
            a, b = tb[i], tb[j]
            if overlaps(a[1], b[1]) and not extra_ok(a[0], b[0]):
                hits.append((norm(a[0].text)[:25], norm(b[0].text)[:25]))
    c.check("no two text-bearing shapes overlap", not hits, hits)
    small = [(sh.id, norm(sh.text)[:25], s) for sh in text_shapes(slide) for s in run_sizes(sh) if s < 8]
    c.check("no text under 8 pt (effective_size)", not small, small)
    rgb = [(sh.id, sh.name, k, v) for sh in leaf_shapes(slide) for k, v in shape_colors(sh) if not is_theme(v)]
    c.check("theme colours only (no hard-coded RGB)", not rgb, rgb)
    return slide


def find_exact(shapes, text):
    t = ntext(text)
    return [sh for sh in shapes if ntext(sh.text) == t]


def lines_of(sh):
    return [ntext(p.text) for p in sh.text_frame.paragraphs]


def has_all(sh, items) -> bool:
    body = ntext(sh.text)
    return all(ntext(i) in body for i in items)


def strip_bullet(s: str) -> str:
    return re.sub(r"^[•\-\*•▪■◦·]\s*", "", s)
