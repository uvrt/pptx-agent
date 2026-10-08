"""Build the inputs of the trial tasks p2-p6, of the spike's graphics tasks p7, p8, o1 and
m1, of p10 (a review in comments), p11 (SmartArt) and p12 (a rebrand through the theme), for
the golden transcripts (test_golden.py).

The trial's own builder (recovered from its coordinator's transcript), run against this
repository's fixtures, writes ``<root>/<task>/input/...``.  p1's template is the committed
``fixtures/generated/trial/company-template.potx`` (built from a deck that embeds fonts, so
it is not rebuilt here), and p3's deck is also committed there; both are read in place.
Pillow draws p5's logos and photo, imported only for p5.  The graphics tasks start from
the committed ``fixtures/generated/trial/halden-proposal-draft.pptx`` (trial 2's draft, as
the spike used it); o1 and m1 retitle one slide of it, as the spike's builder did.  The text
is invented.
"""
from __future__ import annotations

import sys
from datetime import datetime, timezone
from pathlib import Path

from pptx_agent import Document

FIX = Path(__file__).resolve().parent / "fixtures"
DATE = datetime(2026, 9, 28, 9, 0, tzinfo=timezone.utc)
EMU = 914400
INPUTS = Path(".")


def out(task: str) -> Path:
    folder = INPUTS / task / "input"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


# ---------------------------------------------------------------- P2 quarterly update
P2_TEXT = {
    "256.4": "Q2 FY2026\nResults\nSummary",
    "256.5": "Period: 1 April to 30 June 2026\nPublished: 6 August 2026",
    "256.7": "Harbor Lane Logistics N.V.",
    "256.8": "Euronext Amsterdam (ticker: HLL)",
    "256.12": "Revenue", "256.13": "2,140", "256.14": "EUR m", "256.15": "▲ 8.7%",
    "256.16": "Op. profit", "256.17": "257", "256.18": "EUR m", "256.19": "▲ 15.8%",
    "256.20": "Net profit", "256.21": "181", "256.22": "EUR m", "256.23": "▲ 14.6%",
    "256.28": "Operating margin", "256.29": "12.0%", "256.30": "+0.7pt YoY",
    "256.32": "ROE", "256.33": "13.1%", "256.34": "+0.8pt YoY",
    "256.36": "EPS", "256.37": "€2.41", "256.38": "+14.7% YoY",
    "256.40": "Dividend", "256.41": "€0.90", "256.42": "+€0.10",
    "256.43": "All figures in this deck are fictitious.",
    "257.3#1": "Results summary: key figures",
    "257.4": "Q2 FY2026 consolidated results (April to June 2026)",
    "257.6": "Consolidated income statement (main items)",
    "257.9": "Progress against the full-year forecast",
    "257.11": "Revenue", "257.12": "48.6%", "257.13": "Forecast EUR 4,400m",
    "257.15": "Op. profit", "257.16": "49.4%", "257.17": "Forecast EUR 520m",
    "257.19": "Net profit", "257.20": "49.6%", "257.21": "Forecast EUR 365m",
    "257.23": "Revenue and operating profit by quarter",
    "257.27": "Margins by quarter",
    "257.30": "All figures are fictitious.",
    "257.31": "Harbor Lane Logistics | Q2 FY2026 results | p.2",
    "258.3": "Segment analysis",
    "258.4#2": "Results and share of revenue by business segment",
    "258.7": "Segment revenue against the prior year",
    "258.11": "Share of revenue",
    "258.15": "Segment highlights",
    "258.16": ("Road freight: cross-border volumes up 9% on the Rotterdam to Milan corridor\n"
               "Warehousing: two new e-commerce sites opened, in Tilburg and Lyon\n"
               "Air & sea: ocean forwarding margins recovered as rates stabilised\n"
               "Other & adjustments: margin of 10.7%, up from 9.2% a year ago"),
    "258.17": "All figures are fictitious.",
    "258.18": "Harbor Lane Logistics | Q2 FY2026 results | p.3",
}
P2_PL = [
    ["Item", "Q2 actual", "Prior year", "Change", "Change %"],
    ["Revenue", "EUR 2,140m", "EUR 1,968m", "+172m", "+8.7%"],
    ["Gross profit", "EUR 642m", "EUR 600m", "+42m", "+7.0%"],
    ["Operating profit", "EUR 257m", "EUR 222m", "+35m", "+15.8%"],
    ["Profit before tax", "EUR 248m", "EUR 215m", "+33m", "+15.3%"],
    ["Net profit", "EUR 181m", "EUR 158m", "+23m", "+14.6%"],
]
P2_SEG = [
    ["Segment", "Revenue", "Share", "Op. profit", "Margin", "Headcount"],
    ["Road freight", "EUR 980m", "45.8%", "EUR 108m", "11.0%", "5,420"],
    ["Warehousing", "EUR 620m", "29.0%", "EUR 82m", "13.2%", "3,180"],
    ["Air & sea", "EUR 390m", "18.2%", "EUR 51m", "13.1%", "960"],
    ["Other & adjustments", "EUR 150m", "7.0%", "EUR 16m", "10.7%", "410"],
    ["Total", "EUR 2,140m", "100.0%", "EUR 257m", "12.0%", "9,970"],
]


def p2() -> None:
    folder = out("p2-quarterly-update")
    deck = Document.open(FIX / "real-financial-report.pptx")
    deck.delete_slide(deck.slides[3])
    for sid, text in P2_TEXT.items():
        deck.shape(sid).set_text(text)
    widths = {"256.12": 1000000, "256.16": 1000000, "256.20": 1000000, "256.28": 1500000,
              "256.41": 1000000, "256.42": 900000, "257.11": 1000000, "257.15": 1000000,
              "257.19": 1000000, "257.13": 1600000, "257.17": 1600000, "257.21": 1600000}
    for sid, w in widths.items():
        shape = deck.shape(sid)
        centre = shape.left + shape.width // 2
        shape.width = w
        shape.left = centre - w // 2
    for sid, rows in (("257.3#5", P2_PL), ("258.4#3", P2_SEG)):
        table = deck.shape(sid).table
        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                table.cell(r, c).text = value
    c1 = deck.shape("257.25").chart
    for i, cat in enumerate(["Q4 FY25", "Q1 FY26", "Q2 FY26"]):
        c1.set_category(i, cat)
    c1.series[0].name = "Revenue (EUR m)"
    c1.series[1].name = "Op. profit (EUR m)"
    c1.series[0].set_values([2010, 2075, 2140])
    c1.series[1].set_values([231, 240, 257])
    c2 = deck.shape("257.29").chart
    for i, cat in enumerate(["Q4 FY25", "Q1 FY26", "Q2 FY26"]):
        c2.set_category(i, cat)
    for s, name, vals in zip(c2.series, ["Operating margin (%)", "Gross margin (%)", "Net margin (%)"],
                             [[11.5, 11.6, 12.0], [29.9, 29.8, 30.0], [8.2, 8.3, 8.5]]):
        s.name = name
        s.set_values(vals)
    c3 = deck.shape("258.9").chart
    for i, cat in enumerate(["Road freight", "Warehousing", "Air & sea", "Other"]):
        c3.set_category(i, cat)
    c3.series[0].name = "Prior year"
    c3.series[1].name = "Current"
    c3.series[0].set_values([923, 568, 347, 130])
    c3.series[1].set_values([980, 620, 390, 150])
    c4 = deck.shape("258.13").chart
    for i, cat in enumerate(["Road freight", "Warehousing", "Air & sea", "Other"]):
        c4.set_category(i, cat)
    c4.series[0].name = "Share of revenue"
    c4.series[0].set_values([45.8, 29.0, 18.2, 7.0])
    deck.title = "Harbor Lane Logistics Q2 FY2026 results"
    deck.save(folder / "harbor-lane-q2-results.pptx")


# ---------------------------------------------------------------- P3 split an overcrowded slide
P3_OUTLINE = """\
# Pilot retrospective

Customer portal pilot, January to June 2026

# Agenda

- What we set out to do
- Lessons learned
- Next steps

# Lessons learned

- Planning
  - The scope was agreed two weeks after the pilot started
  - Estimates for data migration were 40% too low
  - Weekly planning sessions worked well once they started
- Delivery
  - Releasing every two weeks kept stakeholders engaged
  - Manual testing slowed down the last three releases
  - The rollback plan was never tested
- People
  - Two key developers left in month four
  - Pairing new joiners with experienced staff shortened onboarding
  - Business testers were only available on Fridays
- Tooling
  - The shared test environment was often down
  - Moving the backlog to one board removed duplicate work
  - Monitoring dashboards arrived too late to help
- Customers
  - Pilot customers liked the self-service invoices most
  - Support calls fell by 18% in the pilot group
  - Customers asked for a mobile app in almost every interview

Notes:

Planning: the late scope agreement is the root cause of most of the slippage.

Delivery: the fortnightly releases are the thing to keep; the untested rollback is the thing to fix first.

People: the two departures cost us about six weeks.

Tooling: we will ask for a dedicated test environment in the next phase.

Customers: the 18% fall in support calls is the strongest evidence for a full rollout.

# Next steps

- Agree the scope for the full rollout by 15 November
- Automate the regression tests
- Book business testers for two days a week
"""


def p3() -> None:
    folder = out("p3-split-slide")
    deck = Document.new(title="Pilot retrospective", author="PMO", created=DATE)
    deck.insert_outline(P3_OUTLINE)
    deck.save(folder / "pilot-retrospective.pptx")


# ---------------------------------------------------------------- P4 process diagram
P4_OUTLINE = """\
# Order-to-cash review

Finance operations, October 2026

<!-- layout: Title Only -->
# Order-to-cash process

# Open questions

- Who owns the credit check for new customers?
- Should invoices go out at shipping or at delivery?
"""


def p4() -> None:
    folder = out("p4-process-diagram")
    deck = Document.new(title="Order-to-cash review", author="Finance", created=DATE)
    deck.insert_outline(P4_OUTLINE)
    deck.save(folder / "order-to-cash.pptx")


# ---------------------------------------------------------------- P5 rebrand
OLD_RED, OLD_AMBER, OLD_GREY = "#8B1E3F", "#F2A541", "#3C3C3B"


def logo(path: Path, size, colour, text) -> None:
    from PIL import Image, ImageDraw

    im = Image.new("RGB", size, "white")
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, size[0] - 1, size[1] - 1], fill=colour)
    d.ellipse([size[1] * 0.15, size[1] * 0.15, size[1] * 0.85, size[1] * 0.85], fill="white")
    d.text((size[1] * 0.95, size[1] * 0.4), text, fill="white")
    im.save(path)


def photo(path: Path) -> None:
    from PIL import Image, ImageDraw

    im = Image.new("RGB", (480, 320), (200, 215, 230))
    d = ImageDraw.Draw(im)
    for i, c in enumerate([(90, 110, 140), (140, 100, 80), (80, 130, 90)]):
        d.ellipse([60 + i * 130, 80, 160 + i * 130, 180], fill=c)
        d.rectangle([50 + i * 130, 180, 170 + i * 130, 320], fill=c)
    im.save(path)


P5_OUTLINE = """\
# Kestrel Utilities

Annual customer review 2026

# Service performance

- Supply interruptions fell by 14%
- 96% of new connections made on time
- Complaints answered within five working days

# Customer satisfaction

| Measure | 2025 | 2026 |
| --- | --- | --- |
| Overall satisfaction | 7.4 | 7.9 |
| Billing clarity | 6.8 | 7.5 |
| Contact centre | 7.1 | 7.3 |

# Our field teams

- 1,240 engineers across four regions
- New apprentices: 64 this year
"""


def p5() -> None:
    folder = out("p5-rebrand")
    tmp = folder.parent
    logo(folder / "new-logo.png", (600, 240), (21, 96, 130), "KESTREL")
    old = tmp / "_old-logo.png"
    team = tmp / "_team.png"
    logo(old, (300, 300), (139, 30, 63), "K")
    photo(team)
    deck = Document.new(title="Kestrel annual customer review", author="Communications", created=DATE)
    deck.insert_outline(P5_OUTLINE)
    W, H = deck.slide_size
    s1, s2, s3, s4 = deck.slides
    for s in (s1, s2, s3, s4):
        s.shapes[0].text_frame.paragraph(0).run(0).format(color=OLD_RED)
    band = s1.add_shape("rect", 0, H - EMU // 2, W, EMU // 2)
    band.fill = OLD_RED
    band.line.color = OLD_RED
    bar = s1.add_shape("rect", EMU, int(H * 0.62), 3 * EMU, EMU // 12)
    bar.fill = OLD_AMBER
    bar.send_to_back()
    band.send_to_back()
    for s in (s2, s3, s4):
        top = s.add_shape("rect", 0, 0, W, EMU // 8)
        top.fill = OLD_RED
        top.line.color = OLD_RED
    call = s2.add_shape("roundRect", int(W * 0.68), int(H * 0.62), int(2.6 * EMU), int(0.8 * EMU),
                        text="All targets met")
    call.fill = OLD_AMBER
    call.line.color = OLD_GREY
    call.line.width = 19050
    call.text_frame.paragraph(0).run(0).format(color=OLD_GREY, bold=True)
    table = next(sh for sh in s3.shapes if sh.has_table).table
    for c in range(table.columns):
        table.cell(0, c).fill = OLD_RED
    note = s4.add_textbox(int(0.8 * EMU), int(H * 0.85), int(6 * EMU), int(0.4 * EMU),
                          "Photo: the Northfield crew at the 2026 open day")
    note.text_frame.paragraph(0).run(0).format(color=OLD_GREY, italic=True, size=12)
    s4.add_picture(team, int(W * 0.58), int(H * 0.3), width=int(4 * EMU))
    for s in (s1, s2, s3, s4):
        s.add_picture(old, W - int(1.2 * EMU), int(0.25 * EMU), width=int(0.8 * EMU))
    old.unlink()
    team.unlink()
    deck.save(folder / "customer-review.pptx")


# ---------------------------------------------------------------- P6 restructure
P6_OUTLINE = """\
# Website relaunch: project review

Digital team, October 2026

Notes:

Opening: thank the sponsors.

# Agenda

- Goals
- Timeline
- Budget
- Team
- Risks
- Questions

Notes:

Keep the agenda short.

# Goals

- Faster pages: under two seconds on mobile
- Accessible to WCAG 2.2 AA
- Online sales up 15% in the first year

Notes:

The sales target was agreed with the commercial director.

# Timeline

- Design complete: December 2026
- Beta: March 2027
- Launch: May 2027

Notes:

The beta date depends on the payment provider.

# Budget

| Item | EUR thousand |
| --- | --- |
| Design | 120 |
| Build | 340 |
| Testing | 60 |
| Total | 520 |

Notes:

The budget includes a 10% contingency.

# Team: design

- Lead: Marta Silva
- Two product designers
- One content designer

Notes:

The design team sits with marketing on Tuesdays.

# Risks

- Payment provider migration slips
- Content not ready for launch
- Key staff unavailable over the holidays

Notes:

The payment provider risk is the one to watch.

# Appendix: old figures (2024)

- Traffic 2024: 1.2 million visits
- Conversion 2024: 1.9%

Notes:

These figures are out of date.

# Questions

Thank you

Notes:

Leave ten minutes for questions.
"""


def p6() -> None:
    folder = out("p6-restructure-deck")
    deck = Document.new(title="Website relaunch review", author="Digital team", created=DATE)
    deck.insert_outline(P6_OUTLINE)
    deck.save(folder / "website-relaunch.pptx")


HALDEN = FIX / "generated" / "trial" / "halden-proposal-draft.pptx"


def halden(task: str, retitle: "tuple[int, str] | None" = None):
    """The Halden draft for a graphics task, one slide retitled for o1 and m1."""
    def build() -> None:
        target = out(task) / "halden-proposal-draft.pptx"
        if retitle is None:
            target.write_bytes(HALDEN.read_bytes())
            return
        deck = Document.open(HALDEN)
        deck.slides[retitle[0] - 1].title = retitle[1]
        deck.save(target)
    return build


# ---------------------------------------------------------------- P10 review with comments
def p10() -> None:
    """p2's results deck, with the analyst's open question on the margin chart and an older,
    resolved thread on the title, as a reviewer would find them."""
    p2()
    deck = Document.open(INPUTS / "p2-quarterly-update" / "input" / "harbor-lane-q2-results.pptx")
    old = deck.add_comment("s:256", "Use the full legal name on the cover.", author="Mia Jansen",
                           date=datetime(2026, 9, 21, 14, 5, tzinfo=timezone.utc))
    deck.reply_to_comment(old.address, "Done.", author="Ruben Smit",
                          date=datetime(2026, 9, 22, 8, 40, tzinfo=timezone.utc))
    deck.resolve_comment(old.address)
    deck.add_comment("257.29", "Should the board see all three margins, or is operating margin "
                     "enough?", author="Mia Jansen",
                     date=datetime(2026, 9, 25, 16, 30, tzinfo=timezone.utc))
    deck.save(out("p10-review-comments") / "harbor-lane-q2-review.pptx")


# ---------------------------------------------------------------- P11 SmartArt, P12 theme
def p11() -> None:
    """PowerPoint's own SmartArt deck (a process and a list), with slide titles."""
    deck = Document.open(FIX / "powerpoint-smartart.pptx")
    deck.slides[0].shape("256.2").text = "How we deliver"
    deck.slides[1].shape("257.2").text = "Goals and risks"
    deck.title = "Delivery process"
    deck.save(out("p11-update-smartart") / "delivery-process.pptx")


P12_OUTLINE = """\
# Harbourside Clinics

Patient experience review, autumn 2026

# What patients told us

- Waiting times are the first complaint
- Staff kindness is the first compliment
- Online booking is used by half of new patients

# Three priorities

- Shorter waits in the morning peak
- One booking channel
- Follow-up calls within two days
"""


def p12() -> None:
    """A deck in theme colours throughout -- headline band, two callouts, a three-box row --
    before the brand's new palette and fonts are applied to its theme."""
    deck = Document.new(title="Patient experience review", author="Strategy", created=DATE)
    deck.insert_outline(P12_OUTLINE)
    W, H = deck.slide_size
    s1, s2, s3 = deck.slides
    band = s1.add_shape("rect", 0, H - EMU // 2, W, EMU // 2)
    band.fill = "accent1"
    band.line.visible = False
    call = s2.add_shape("roundRect", int(W * 0.66), int(H * 0.6), int(2.8 * EMU), int(0.8 * EMU),
                        text="74% would recommend us")
    call.fill = "accent2"
    call.line.visible = False
    for n, word in enumerate(("Waits", "Booking", "Follow-up")):
        box = s3.add_shape("rect", int(EMU * (0.8 + 4.1 * n)), int(H * 0.72), int(3.6 * EMU),
                           int(0.7 * EMU), text=word)
        box.fill = "accent1" if n == 0 else "accent1 lumMod=60% lumOff=40%"
        box.line.visible = False
    deck.save(out("p12-theme-rebrand") / "patient-review.pptx")


BUILDERS = {"p2-quarterly-update": p2, "p3-split-slide": p3, "p4-process-diagram": p4,
            "p5-rebrand": p5, "p6-restructure-deck": p6,
            "p7-phased-approach": halden("p7-phased-approach"),
            "p8-engagement-plan": halden("p8-engagement-plan"),
            "o1-governance-chart": halden("o1-governance-chart", (2, "Programme governance")),
            "m1-priority-matrix": halden("m1-priority-matrix", (3, "Prioritisation")),
            "p10-review-comments": p10, "p11-update-smartart": p11,
            "p12-theme-rebrand": p12}


def build(task: str, root: Path) -> Path:
    """Write ``task``'s inputs under ``root/<task>/input``; returns that folder."""
    global INPUTS
    INPUTS = Path(root)
    BUILDERS[task]()
    return INPUTS / task / "input"


if __name__ == "__main__":
    target = Path(sys.argv[1])
    for name in sys.argv[2:] or list(BUILDERS):
        print("built", build(name, target))
