"""Made-up decks for the layout-change tests (round 7, L2): the blank deck's master and
theme with layouts and slides of our own, written as XML so each case states exactly the
placeholders PowerPoint was measured on.  Boxes are in points on a 960 x 540 pt slide."""

from __future__ import annotations

import io
import re
import struct
import zipfile
import zlib

from pptx_agent import Document

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
GROUP = ('<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
         '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/>'
         '<a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_LAYOUT = "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"
CT_SLIDE = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
PT = 12700


def xfrm(box) -> str:
    x, y, w, h = (round(v * PT) for v in box)
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'


def ph(sid, name, kind=None, idx=None, box=None, text=None, *, size=None, extra="",
       layout=False):
    """A placeholder ``p:sp``: on a layout with a prompt, on a slide empty (``text`` None)
    or holding ``text``."""
    attributes = (f' type="{kind}"' if kind else "") + extra + (
        f' idx="{idx}"' if idx is not None else "")
    if layout:
        body = f'<a:p><a:r><a:rPr lang="en-US"/><a:t>{name} prompt</a:t></a:r></a:p>'
    elif text is None:
        body = '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    else:
        sz = f' sz="{size}"' if size else ""
        body = f'<a:p><a:r><a:rPr lang="en-US"{sz} dirty="0"/><a:t>{text}</a:t></a:r></a:p>'
    tx = "" if (kind == "pic" and not layout and text is None) else \
        f"<p:txBody><a:bodyPr/><a:lstStyle/>{body}</p:txBody>"
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr><a:spLocks '
            f'noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{attributes}/></p:nvPr></p:nvSpPr>'
            f'<p:spPr>{xfrm(box) if box else ""}</p:spPr>{tx}</p:sp>')


def rect(sid, name, box, text) -> str:
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr/>'
            f'</p:nvSpPr><p:spPr>{xfrm(box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/>'
            f'<a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>{text}</a:t></a:r>'
            '</a:p></p:txBody></p:sp>')


def pic(sid, name, rid, box=None, placeholder="", crop="") -> str:
    """A picture; ``placeholder`` its ``p:ph`` (a picture in a placeholder), ``crop`` the
    attributes of its ``a:srcRect``."""
    lock = ' noGrp="1"' if placeholder else ""
    rect_ = f"<a:srcRect {crop}/>" if crop else ""
    geometry = "" if placeholder else '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvPicPr><a:picLocks'
            f'{lock} noChangeAspect="1"/></p:cNvPicPr><p:nvPr>{placeholder}</p:nvPr></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}"/>{rect_}<a:stretch><a:fillRect/></a:stretch>'
            f'</p:blipFill><p:spPr>{xfrm(box) if box else ""}{geometry}</p:spPr></p:pic>')


def png(width=4, height=4, dpi=None) -> bytes:
    """A one-colour PNG; ``dpi`` writes a ``pHYs`` chunk."""
    raw = b"".join(b"\x00" + b"\xc8\x1e\x1e" * width for _ in range(height))

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    density = b""
    if dpi:
        per_metre = round(dpi / 0.0254)
        density = chunk(b"pHYs", struct.pack(">IIB", per_metre, per_metre, 1))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2,
                                                            0, 0, 0))
            + density + chunk(b"IDAT", zlib.compress(raw))
            + chunk(b"IEND", b""))


def frame(sid, name, placeholder, box, *, chart=False, columns=(200, 200), rows=(40, 40),
          margins="") -> str:
    """A table (``columns`` and ``rows`` in points, ``margins`` the cells' ``tcPr``
    attributes) or, with ``chart``, a chart frame (its chart ``rId3``)."""
    if chart:
        data = ('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" '
                'r:id="rId3"/></a:graphicData>')
    else:
        grid = "".join(f'<a:gridCol w="{w * PT}"/>' for w in columns)
        cell = ('<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/>'
                f'<a:t>{{}}</a:t></a:r></a:p></a:txBody><a:tcPr {margins}/></a:tc>')
        body = "".join(f'<a:tr h="{h * PT}">' + "".join(
            cell.format(f"{chr(65 + c)}{r + 1}") for c in range(len(columns))) + "</a:tr>"
            for r, h in enumerate(rows))
        data = ('<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
                f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{body}'
                '</a:tbl></a:graphicData>')
    x, y, w, h = (round(v * PT) for v in box)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{name}"/>'
            '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            f'<p:nvPr>{placeholder}</p:nvPr></p:nvGraphicFramePr><p:xfrm><a:off x="{x}" '
            f'y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm><a:graphic>{data}</a:graphic>'
            '</p:graphicFrame>')


CHART = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c='
         '"http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.'
         'openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:layout/><c:barChart>'
         '<c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:idx val="0"/><c:order '
         'val="0"/><c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt '
         'idx="1"><c:v>5</c:v></c:pt></c:numLit></c:val></c:ser><c:axId val="1"/><c:axId '
         'val="2"/></c:barChart><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"'
         '/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>'
         '<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete'
         ' val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx></c:plotArea></c:chart>'
         '</c:chartSpace>')


def title(name, box=(50, 20, 860, 80)) -> str:
    return ph(2, f"{name} Title", "title", None, box, layout=True)


#: The measured layouts (oracle build_l2*.py), by our own short names.
LAYOUTS = {
    "LA": [title("LA"), ph(3, "LA Body", None, 1, (50, 120, 860, 380), layout=True)],
    "LB": [title("LB", (60, 30, 840, 70)),
           ph(3, "LB Body2", None, 2, (100, 150, 760, 340), layout=True)],
    "LB2": [title("LB2"), ph(3, "LB2 Body5", None, 5, (100, 150, 760, 340), layout=True)],
    "LC": [title("LC"), ph(3, "LC Left", None, 1, (50, 120, 420, 380), extra=' sz="half"',
                           layout=True),
           ph(4, "LC Right", None, 2, (490, 120, 420, 380), extra=' sz="half"', layout=True)],
    "LD": [title("LD"), ph(3, "LD Picture", "pic", 13, (200, 120, 560, 380), layout=True)],
    "LE": [title("LE"), ph(3, "Custom Notes Box", "body", 14, (50, 400, 860, 100),
                           layout=True)],
    "LF": [title("LF", (50, 200, 860, 100))],
    "LG": [ph(3, "LG Body", None, 1, (50, 50, 860, 450), layout=True)],
    "LI": [ph(3, "LI Body", None, 1, (50, 20, 860, 300), layout=True),
           title("LI", (50, 400, 860, 100))],
    "LJ": [title("LJ"), ph(3, "Other", None, 8, (50, 120, 420, 380), layout=True),
           ph(4, "Notes Area", None, 7, (490, 120, 420, 380), layout=True)],
    "LJ2": [title("LJ2"), ph(3, "Other", None, 7, (50, 120, 420, 380), layout=True),
            ph(4, "Notes Area", None, 8, (490, 120, 420, 380), layout=True)],
    "LJ3": [title("LJ3"), ph(3, "Zeta", None, 9, (50, 120, 420, 380), layout=True),
            ph(4, "Alpha", None, 3, (490, 120, 420, 380), layout=True)],
    "LK": [title("LK"), ph(3, "LK Text", "body", 1, (70, 140, 820, 360), layout=True)],
    "LL": [ph(2, "LL CtrTitle", "ctrTitle", None, (100, 150, 760, 120), layout=True),
           ph(3, "LL Subtitle", "subTitle", 1, (100, 300, 760, 100), layout=True)],
    "LM": [title("LM"), ph(3, "LM Text5", "body", 5, (50, 120, 420, 380), layout=True),
           ph(4, "LM Content6", None, 6, (490, 120, 420, 380), layout=True)],
    "LN": [title("LN"), ph(3, "LN Table", "tbl", 1, (80, 130, 800, 370), layout=True)],
    "LP": [title("LP"), ph(3, "LP Body1", None, 1, (50, 380, 860, 130), layout=True),
           ph(4, "LP Extra3", None, 3, (50, 120, 860, 240), layout=True)],
    "LQ": [title("LQ"), ph(3, "LQ Content1", None, 1, (50, 120, 420, 380), layout=True),
           ph(4, "LQ Text2", "body", 2, (490, 120, 420, 380), layout=True)],
    "LR": [],
    "LS": [title("LS"), ph(3, "LS Pic1", "pic", 1, (200, 120, 560, 380), layout=True)],
    "LT": [title("LT"), ph(3, "LT Q1", None, 1, (50, 120, 420, 380), extra=' sz="quarter"',
                           layout=True),
           ph(4, "LT Full2", None, 2, (490, 120, 420, 380), layout=True)],
    "LU": [title("LU"), ph(3, "LU Obj1", None, 1, (50, 120, 420, 380), layout=True),
           ph(4, "LU Text2", "body", 2, (490, 120, 420, 380), layout=True)],
    "LV": [title("LV"), ph(3, "LV Pic1", "pic", 1, (50, 120, 420, 380), layout=True),
           ph(4, "LV Obj2", None, 2, (490, 120, 420, 380), layout=True)],
}


def build(slides, layouts=LAYOUTS, media=None) -> Document:
    """A deck whose layouts are ``layouts`` (``{name: [shape xml]}``, in order) and whose
    slides are ``slides``: ``(layout name, [shape xml])``, a picture's image ``rId2``."""
    base = zipfile.ZipFile(io.BytesIO(Document.new().to_bytes()))
    parts = {n: base.read(n) for n in base.namelist() if not n.startswith("ppt/slideLayouts/")}
    names = list(layouts)
    content = parts["[Content_Types].xml"].decode()
    content = re.sub(r'<Override PartName="/ppt/slideLayouts/[^"]+"[^>]*/>', "", content)
    master_rels = parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"].decode()
    master_rels = re.sub(r'<Relationship Id="[^"]+" Type="[^"]+/slideLayout" Target="[^"]+"/>',
                         "", master_rels)
    rels, ids, overrides = "", "", ""
    for k, name in enumerate(names, 1):
        parts[f"ppt/slideLayouts/slideLayout{k}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sldLayout {NS} '
            f'preserve="1"><p:cSld name="{name}"><p:spTree>{GROUP}{"".join(layouts[name])}'
            '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '</p:sldLayout>').encode()
        parts[f"ppt/slideLayouts/_rels/slideLayout{k}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns='
            '"http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '
            f'Id="rId1" Type="{REL}slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
            '</Relationships>').encode()
        rels += (f'<Relationship Id="rIdL{k}" Type="{REL}slideLayout" '
                 f'Target="../slideLayouts/slideLayout{k}.xml"/>')
        ids += f'<p:sldLayoutId id="{2147483648 + k}" r:id="rIdL{k}"/>'
        overrides += (f'<Override PartName="/ppt/slideLayouts/slideLayout{k}.xml" '
                      f'ContentType="{CT_LAYOUT}"/>')
    parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = master_rels.replace(
        "</Relationships>", rels + "</Relationships>").encode()
    master = parts["ppt/slideMasters/slideMaster1.xml"].decode()
    master = re.sub(r"<p:sldLayoutIdLst>.*?</p:sldLayoutIdLst>",
                    f"<p:sldLayoutIdLst>{ids}</p:sldLayoutIdLst>", master, flags=re.S)
    parts["ppt/slideMasters/slideMaster1.xml"] = master.encode()
    presentation_rels = parts["ppt/_rels/presentation.xml.rels"].decode()
    slide_ids = ""
    for k, (layout, shapes) in enumerate(slides, 1):
        parts[f"ppt/slides/slide{k}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:sld {NS}><p:cSld>'
            f'<p:spTree>{GROUP}{"".join(shapes)}</p:spTree></p:cSld><p:clrMapOvr>'
            '<a:masterClrMapping/></p:clrMapOvr></p:sld>').encode()
        image = (f'<Relationship Id="rId2" Type="{REL}image" Target="../media/image1.png"/>'
                 if any("r:embed" in shape for shape in shapes) else "")
        if any("<c:chart " in shape for shape in shapes):
            image += (f'<Relationship Id="rId3" Type="{REL}chart" '
                      f'Target="../charts/chart{k}.xml"/>')
            parts[f"ppt/charts/chart{k}.xml"] = CHART.encode()
            overrides += (f'<Override PartName="/ppt/charts/chart{k}.xml" ContentType='
                          '"application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>')
        parts[f"ppt/slides/_rels/slide{k}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns='
            '"http://schemas.openxmlformats.org/package/2006/relationships"><Relationship '
            f'Id="rId1" Type="{REL}slideLayout" Target="../slideLayouts/slideLayout'
            f'{names.index(layout) + 1}.xml"/>{image}</Relationships>').encode()
        presentation_rels = presentation_rels.replace(
            "</Relationships>", f'<Relationship Id="rIdS{k}" Type="{REL}slide" '
            f'Target="slides/slide{k}.xml"/></Relationships>')
        slide_ids += f'<p:sldId id="{255 + k}" r:id="rIdS{k}"/>'
        overrides += (f'<Override PartName="/ppt/slides/slide{k}.xml" '
                      f'ContentType="{CT_SLIDE}"/>')
    parts["ppt/_rels/presentation.xml.rels"] = presentation_rels.encode()
    presentation = parts["ppt/presentation.xml"].decode()
    presentation = re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", "", presentation, flags=re.S)
    presentation = presentation.replace(
        "</p:sldMasterIdLst>", f"</p:sldMasterIdLst><p:sldIdLst>{slide_ids}</p:sldIdLst>")
    parts["ppt/presentation.xml"] = presentation.encode()
    if any("r:embed" in shape for _, shapes in slides for shape in shapes):
        if 'Extension="png"' not in content:
            overrides = '<Default Extension="png" ContentType="image/png"/>' + overrides
        parts["ppt/media/image1.png"] = media or png()
    parts["[Content_Types].xml"] = content.replace("</Types>", overrides + "</Types>").encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for name, data in parts.items():
            archive.writestr(name, data)
    return Document.open(out.getvalue())


def dump(slide) -> list[str]:
    """The slide's shapes as the oracle's compare files list them: tag, id, name, ``p:ph``
    attributes, own box in points, run sizes, text."""
    from lxml import etree

    out = []
    for shape in slide._sp_tree():
        tag = etree.QName(shape).localname
        if tag in ("nvGrpSpPr", "grpSpPr", "extLst"):
            continue
        cnv = next(shape.iter("{%s}cNvPr" % P))
        ph = next(shape.iter("{%s}ph" % P), None)
        attrs = ",".join(f"{k}={v}" for k, v in ph.attrib.items()) if ph is not None else None
        box = next((n for n in shape.iter("{%s}xfrm" % A, "{%s}xfrm" % P)), None)
        frame = ""
        if box is not None:
            off, ext = box.find("{%s}off" % A), box.find("{%s}ext" % A)
            values = [int(off.get("x")), int(off.get("y")), int(ext.get("cx")),
                      int(ext.get("cy"))]
            frame = " xfrm(" + ",".join(f"{v / PT:g}" for v in values) + ")"
        sizes = [r.get("sz") for r in shape.iter("{%s}rPr" % A) if r.get("sz")]
        text = "".join(t.text or "" for t in shape.iter("{%s}t" % A))
        body = "Y" if shape.find("{%s}txBody" % P) is not None else "N"
        out.append(f"{tag} id={cnv.get('id')} {cnv.get('name')!r}"
                   + (f" ph({attrs})" if ph is not None else "")
                   + frame + (f" sz={sizes}" if sizes else "") + f" tx={body} {text!r}")
    return out


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
