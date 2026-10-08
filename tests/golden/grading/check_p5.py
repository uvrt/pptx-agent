"""P5: hard-coded old brand colours -> theme colours; old logo -> new logo (aspect kept)."""
import hashlib
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, all_shapes, main  # noqa: E402

from PIL import Image  # noqa: E402
from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p5-rebrand/input/customer-review.pptx"
NEW_LOGO = INPUTS / "p5-rebrand/input/new-logo.png"
MAP = {"#8B1E3F": {"accent1"}, "#F2A541": {"accent2"}, "#3C3C3B": {"tx1", "dk1"}}


def cstr(x):
    if x is None:
        return None
    x = getattr(x, "color", x)
    return None if x is None else str(x)


def base(s):
    s = s or ""
    return s.split()[0] if s else ""


def inventory(deck):
    """Every colour use: (where, kind, value)."""
    uses = []
    for i, slide in enumerate(deck.slides):
        for sh in all_shapes(slide):
            if sh.kind in ("shape", "connector"):
                if sh.fill is not None:
                    uses.append((sh.id, "fill", cstr(sh.fill)))
                try:
                    if sh.line.color is not None:
                        uses.append((sh.id, "line", cstr(sh.line.color)))
                except Exception:  # noqa: BLE001
                    pass
            if sh.kind == "shape" or sh.kind == "connector":
                try:
                    for p in sh.text_frame.paragraphs:
                        for r in p.runs:
                            if r.color is not None:
                                uses.append((f"{sh.id}/{p.address}/{r.text[:12]}", "text", cstr(r.color)))
                except Exception:  # noqa: BLE001
                    pass
            if sh.has_table:
                t = sh.table
                for r in range(t.rows):
                    for k in range(t.columns):
                        cell = t.cell(r, k)
                        if cell.fill is not None:
                            uses.append((f"{sh.id}/cell{r},{k}", "cell", cstr(cell.fill)))
    return uses


def body(c, out: Path):
    deck = Document.open(out)
    src = Document.open(INPUT)
    c.check("validate() clean", not deck.validate(), [str(p) for p in deck.validate()])
    c.check("still 4 slides", len(deck.slides) == 4, len(deck.slides))
    before = inventory(src)
    after = {(w, k): v for w, k, v in inventory(deck)}
    old_left = [(w, k, v) for (w, k), v in after.items() if v and base(v).upper() in MAP]
    c.check("no old brand RGB left anywhere", not old_left, old_left[:8])
    wrong = []
    for w, k, v in before:
        if base(v).upper() not in MAP:
            continue
        got = after.get((w, k))
        if got is None or base(got) not in MAP[base(v).upper()]:
            wrong.append((w, k, v, got))
    c.check("each old colour now the mapped theme colour (same place)", not wrong, wrong[:8])
    # logos
    new_hash = hashlib.sha256(NEW_LOGO.read_bytes()).hexdigest()
    nw, nh = Image.open(NEW_LOGO).size
    with zipfile.ZipFile(out) as z, zipfile.ZipFile(INPUT) as zs:
        src_pics = {sh.id: (sh, hashlib.sha256(zs.read(sh.image_part)).hexdigest())
                    for s in src.slides for sh in s.shapes if sh.kind == "picture"}
        old_logo_hash = max(set(h for _, h in src_pics.values()),
                            key=lambda h: sum(1 for _, x in src_pics.values() if x == h))
        W, H = deck.slide_size
        logos_ok, photo_ok, geo = [], [], []
        for i, s in enumerate(deck.slides):
            pics = [sh for sh in s.shapes if sh.kind == "picture"]
            hashes = [(sh, hashlib.sha256(z.read(sh.image_part)).hexdigest()) for sh in pics]
            new = [sh for sh, h in hashes if h == new_hash]
            old = [sh for sh, h in hashes if h == old_logo_hash]
            logos_ok.append(len(new) == 1 and not old)
            olds = [sh for sid, (sh, h) in src_pics.items() if h == old_logo_hash and sid.split(".")[0] ==
                    str(s.slide_id)]
            if new and olds:
                n, o = new[0], olds[0]
                ratio = n.width / n.height
                geo.append({
                    "slide": i + 1,
                    "height_kept": abs(n.height - o.height) <= 12700,
                    "aspect_ok": abs(ratio - nw / nh) / (nw / nh) < 0.03,
                    "top_kept": abs(n.top - o.top) <= 25400,
                    "top_right": abs((n.left + n.width) - (o.left + o.width)) <= 0.25 * 914400 or
                    (n.left + n.width) > W * 0.85,
                    "on_slide": n.left >= 0 and n.left + n.width <= W,
                })
            if i == 3:
                photo = [sh for sh, h in hashes if h not in (new_hash, old_logo_hash)]
                photo_ok.append(len(photo) == 1)
        c.check("every slide shows the new logo once and no old logo", all(logos_ok) and len(logos_ok) == 4,
                logos_ok)
        c.check("team photo kept", all(photo_ok) and photo_ok, photo_ok)
        for key in ("height_kept", "aspect_ok", "top_kept", "top_right", "on_slide"):
            c.check(f"logo geometry: {key}", geo and all(g[key] for g in geo) and len(geo) == 4,
                    [(g["slide"], g[key]) for g in geo])
    # nothing else changed: outlines identical apart from the media names
    import re
    strip = lambda md: re.sub(r"!\[[^\]]*\]\([^)]*\)", "![]()", md)  # noqa: E731
    same = strip(deck.to_outline()) == strip(src.to_outline())
    c.check("text and structure unchanged", same)


if __name__ == "__main__":
    main("p5-rebrand", body)
