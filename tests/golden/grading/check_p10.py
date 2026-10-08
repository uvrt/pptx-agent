"""P10: a review left as PowerPoint comments: two new threads where the brief puts them, a reply
that resolves the analyst's thread, the old thread untouched, and no slide changed.

Read from the package with lxml, not through pptx-agent's comment API, so the check does not
trust what it checks."""
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # common.py beside it
from common import INPUTS, main, norm  # noqa: E402

from lxml import etree  # noqa: E402

from pptx_agent import Document  # noqa: E402

INPUT = INPUTS / "p10-review-comments" / "input" / "harbor-lane-q2-review.pptx"
P188 = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
NS = {"p188": P188, "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "pc": "http://schemas.microsoft.com/office/powerpoint/2013/main/command",
      "ac": "http://schemas.microsoft.com/office/drawing/2013/main/command",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
REL_COMMENTS = "http://schemas.microsoft.com/office/2018/10/relationships/comments"
REVIEWER = "Joost de Vries"


def text_of(node) -> str:
    body = node.find("p188:txBody", NS)
    return "\n".join("".join(t.text or "" for t in p.iter(f"{{{NS['a']}}}t"))
                     for p in body.findall("a:p", NS)) if body is not None else ""


def read(path: Path):
    """Slides (sldId -> part), each slide's threads, the authors, the content types."""
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    targets = {r.get("Id"): "ppt/" + r.get("Target") for r in rels}
    slides = [(int(e.get("id")), targets[e.get(f"{{{NS['r']}}}id")])
              for e in pres.find("p:sldIdLst", NS)]
    authors = {}
    for r in rels:
        if r.get("Type").endswith("/authors"):
            for a in etree.fromstring(z.read("ppt/" + r.get("Target"))).findall("p188:author", NS):
                authors[a.get("id")] = a.get("name")
    threads = {}
    shapes = {}
    for sid, part in slides:
        root = etree.fromstring(z.read(part))
        shapes[sid] = {c.get("id"): "".join(t.text or "" for t in c.getparent().getparent()
                                            .iter(f"{{{NS['a']}}}t"))
                       for c in root.iter(f"{{{NS['p']}}}cNvPr")}
        rel_path = part.replace("slides/", "slides/_rels/") + ".rels"
        threads[sid] = []
        if rel_path not in names:
            continue
        for r in etree.fromstring(z.read(rel_path)):
            if r.get("Type") == REL_COMMENTS:
                target = "ppt/comments/" + r.get("Target").split("/")[-1]
                threads[sid] = etree.fromstring(z.read(target)).findall("p188:cm", NS)
                ext = [e for e in root.iter(f"{{{NS['p']}}}ext")
                       if e.get("uri") == "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}"]
                threads[sid + 0.5] = bool(ext) and ext[0][0].get(f"{{{NS['r']}}}id") == r.get("Id")
    return z, slides, threads, authors, shapes


def slide_text(z, part) -> str:
    root = etree.fromstring(z.read(part))
    for ext in list(root.iter(f"{{{NS['p']}}}ext")):
        if ext.get("uri") == "{6950BFC3-D8DA-4A85-94F7-54DA5524770B}":
            ext.getparent().remove(ext)
    tree = root.find("p:cSld", NS)
    return etree.tostring(tree, method="c14n").decode()


def body(c, out: Path):
    new = {str(p) for p in Document.open(out).validate()} - \
        {str(p) for p in Document.open(INPUT).validate()}
    c.check("no new validation problems", not new, sorted(new))
    z, slides, threads, authors, shapes = read(out)
    z0, slides0, threads0, _, _ = read(INPUT)
    c.check("no slide changed", [slide_text(z, p) for _, p in slides]
            == [slide_text(z0, p) for _, p in slides0], "slide content differs")
    c.check("each comments part named in its slide's extLst", all(
        threads.get(sid + 0.5, True) for sid, _ in slides if threads.get(sid)))
    every = [cm for sid, _ in slides for cm in threads[sid]]
    c.check("four threads: two new, two kept", len(every) == 4, len(every))
    c.check("every author is in the authors part", all(
        cm.get("authorId") in authors and all(r.get("authorId") in authors
                                              for r in cm.iter(f"{{{P188}}}reply"))
        for cm in every), authors)
    c.check("the reviewer is one author", list(authors.values()).count(REVIEWER) == 1, authors)
    by = lambda cm: authors.get(cm.get("authorId"))  # noqa: E731

    s1, s2, s3 = (sid for sid, _ in slides)
    figure = [sid for sid, text in shapes[s1].items() if norm(text) == "2,140"]
    on_figure = [cm for cm in threads[s1] if by(cm) == REVIEWER
                 and cm.find("ac:deMkLst/ac:spMk", NS) is not None
                 and cm.find("ac:deMkLst/ac:spMk", NS).get("id") in figure]
    c.check("a comment on the 2,140 figure, by the reviewer", len(on_figure) == 1,
            [text_of(cm) for cm in threads[s1]])
    c.check("it asks to confirm against the audited close", on_figure and
            "audited Q2 close" in text_of(on_figure[0]), on_figure and text_of(on_figure[0]))
    marker = on_figure and on_figure[0].find("ac:deMkLst/pc:sldMk", NS)
    c.check("the shape anchor names slide 1", marker is not None and marker is not False
            and marker.get("sldId") == str(s1))
    on_slide3 = [cm for cm in threads[s3] if by(cm) == REVIEWER
                 and cm.find("pc:sldMkLst/pc:sldMk", NS) is not None]
    c.check("a comment on slide 3 as a whole", len(on_slide3) == 1
            and on_slide3[0].find("pc:sldMkLst/pc:sldMk", NS).get("sldId") == str(s3),
            [text_of(cm) for cm in threads[s3]])
    c.check("it asks for the source line", on_slide3 and
            "Source: management accounts, Q2 FY2026" in text_of(on_slide3[0]))
    mia = [cm for cm in threads[s2] if by(cm) == "Mia Jansen"]
    c.check("Mia's thread is resolved", len(mia) == 1 and mia[0].get("status") == "resolved")
    replies = mia[0].findall("p188:replyLst/p188:reply", NS) if mia else []
    c.check("with the reviewer's reply", len(replies) == 1 and by(replies[0]) == REVIEWER
            and "net margin in Q1" in text_of(replies[0]), [text_of(r) for r in replies])
    old_now = [etree.tostring(cm, method="c14n") for cm in threads[s1] if by(cm) != REVIEWER]
    old_then = [etree.tostring(cm, method="c14n") for cm in threads0[s1]]
    c.check("the old thread is unchanged", old_now == old_then)
    created = [cm.get("created") for cm in on_figure + on_slide3] + \
        [r.get("created") for r in replies]
    c.check("dates are ISO date-times", all(created) and all(
        re.match(r"^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?Z?$", d) for d in created), created)
    ids = [cm.get("id") for cm in every] + [r.get("id") for cm in every
                                             for r in cm.iter(f"{{{P188}}}reply")]
    c.check("ids are unique braced GUIDs", len(set(ids)) == len(ids) and all(
        re.match(r"^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$", i) for i in ids), ids)


if __name__ == "__main__":
    main("p10-review-comments", body)
