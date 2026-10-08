"""E5's acceptance slide: one of every new kind of shape, built on a fresh slide of a deck.

Shared by the validity tests and the PowerPoint oracle.  Positions are laid out on a
10 x 5.625 inch grid -- the smallest slide in the corpus -- and scaled to the deck's size.
"""

from __future__ import annotations

from pptx_agent import Document

EMU_PER_INCH = 914400
#: The connectors are drawn in this, thick, so a rasterised page shows where they run.
CONNECTOR_RED = "FF0000"
SENTINELS = ("E5BOX", "E5CELL", "E5KID", "E5FROM")


def slide_size(document: Document) -> tuple[int, int]:
    from pptx_agent.oxml.xml import find, get_int

    root = document.package.tree(document.package.presentation_part())
    size = find(root, "p:sldSz")
    return get_int(size, "cx", 9144000), get_int(size, "cy", 5143500)


def acceptance_edits(document: Document) -> dict:
    """Add a slide with presets and adjustments, a text box, three connectors attached
    between two shapes (one shape then moved and the other resized), an edited table and a
    shape added inside a group.  Returns what to look for afterwards."""
    width, height = slide_size(document)
    sx, sy = width / 10.0, height / 5.625

    def at(x: float, y: float, w: float, h: float) -> tuple[int, int, int, int]:
        return (round(x * sx), round(y * sy), round(w * sx), round(h * sy))

    slide = document.add_slide()
    presets = [
        slide.add_shape("roundRect", *at(0.3, 0.3, 1.4, 0.8), text="Rounded",
                        adjustments={"adj": 30000}),
        slide.add_shape("rightArrow", *at(1.9, 0.3, 1.4, 0.8),
                        adjustments={"adj1": 30000, "adj2": 70000}),
        slide.add_shape("star5", *at(3.5, 0.3, 0.9, 0.8)),
        slide.add_shape("donut", *at(4.6, 0.3, 0.8, 0.8)),
        slide.add_shape("flowChartDecision", *at(5.6, 0.3, 1.4, 0.8), text="?"),
    ]
    presets[2].adjustments["adj"] = 15000
    presets[2].bring_forward()
    presets[3].send_backward()

    source = slide.add_shape("rect", *at(0.5, 1.9, 1.6, 0.9), text="E5FROM")
    target = slide.add_shape("ellipse", *at(6.0, 2.2, 1.2, 1.2), text="To")
    red = {"color": CONNECTOR_RED, "width": 38100}
    elbow = slide.add_connector("elbow", (source, 3), (target, 2), line=red)
    curved = slide.add_connector("curved", (source, 2), (target, 4), line=red)
    straight = slide.add_connector("straight", (source, 0), (target, 0),
                                   line=dict(red, tail="triangle"))
    target.move_by(round(-1.0 * sx), round(1.1 * sy))   # the connectors follow
    source.width = round(2.0 * sx)

    slide.add_textbox(*at(7.4, 0.3, 2.3, 0.4), "E5BOX")
    table_frame = slide.add_table(3, 2, *at(7.4, 1.2, 2.3, 1.2))
    table = table_frame.table
    table.cell(0, 0).text = "Quarter"
    table.cell(0, 1).text = "Revenue"
    table.cell(1, 0).text = "E5CELL"
    table.insert_row(2)
    table.merge(2, 0, 3, 0)

    first = slide.add_shape("rect", *at(7.6, 3.6, 0.6, 0.4))
    second = slide.add_shape("rect", *at(8.6, 3.9, 0.6, 0.4))
    group = slide.group([first, second])
    child = group.add_shape("ellipse", *at(7.9, 4.5, 1.5, 0.6), text="E5KID")

    probes = []
    for connector in (elbow, curved, straight):
        ends = [getattr(document.shape(connector.id), f"{end}_connection")
                for end in ("begin", "end")]
        for mine, other in (ends, ends[::-1]):
            probes.append(_probe(mine, other if connector is straight else None))
    return {
        "pages": len(document.slides),
        "page": slide.index,
        "slide_id": slide.slide_id,
        "connectors": [elbow.id, curved.id, straight.id],
        "group": group.id,
        "child": child.id,
        "probes": probes,
        "size": (width, height),
    }


def _probe(connection, toward=None, distance: float = 0.06) -> tuple[float, float]:
    """A point (inches) just outside a site, on the connector leaving it: along the site's
    direction for an elbow or a curve, toward the other end (``toward``) for a line."""
    from pptx_agent.edit.creating import shape_sites

    shape, site = connection
    end = shape_sites(shape)[site]
    if toward is None:
        dx, dy = end.direction
    else:
        other = shape_sites(toward[0])[toward[1]].point
        dx, dy = other[0] - end.point[0], other[1] - end.point[1]
        length = (dx * dx + dy * dy) ** 0.5 or 1.0
        dx, dy = dx / length, dy / length
    x = end.point[0] + dx * distance * EMU_PER_INCH
    y = end.point[1] + dy * distance * EMU_PER_INCH
    return x / EMU_PER_INCH, y / EMU_PER_INCH
