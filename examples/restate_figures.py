"""Restate a quarter's figures throughout a results deck: text, tables and charts.

The end-to-end pilot's task (ROADMAP.md, "Usability"), written again from docs/USAGE.md.
Finance restated Q3 FY2025 revenue, operating profit and net profit in
``tests/fixtures/real-financial-report.pptx``; everything derived from them follows --
year-on-year changes, progress against the full-year forecast, margins, the segment
table's その他・調整 and 合計 rows and every segment's share -- and so do the charts.

Tables and charts are addressed by their labels, so a wrong row cannot be written
silently; every other figure is read from the deck rather than typed in.  Percentages are
rounded half up to one decimal, and point changes are taken between the rounded figures,
as the deck's own +0.6pt is.

    python examples/restate_figures.py tests/fixtures/real-financial-report.pptx out.pptx
"""

from __future__ import annotations

import sys
from decimal import ROUND_HALF_UP, Decimal

from pptx_agent import Document

#: The restated figures, in 億円.
RESTATED = {"売上高": 4310, "営業利益": 520, "当期純利益": 333}


def pct(part, whole) -> Decimal:
    """``part / whole`` as a percentage, half up to one decimal."""
    return (Decimal(part) * 100 / Decimal(whole)).quantize(Decimal("0.1"), ROUND_HALF_UP)


def oku(text: str) -> int:
    """``"4,285億円"`` -> 4285."""
    return int(text.replace("億円", "").replace(",", ""))


def plus(value) -> str:
    return f"+{value}" if value >= 0 else f"{value}"


def replace(deck: Document, address: str, old: str, new: str) -> None:
    """Write ``new`` where the deck says ``old`` -- and refuse if it does not."""
    target = deck.resolve(address)
    assert target.text == old, f"{address} reads {target.text!r}, expected {old!r}"
    target.text = new


def restate(deck: Document) -> None:
    pl = deck.shape("257.3#5").table              # 連結損益計算書
    forecast = deck.shape("259.5#5").table        # 通期業績予想
    segments = deck.shape("258.4#3").table        # セグメント別

    old = {item: oku(pl.cell_by_label(item, "当期実績").text) for item in RESTATED}
    prior = {item: oku(pl.cell_by_label(item, "前年同期").text) for item in RESTATED}
    full_year = {item: oku(forecast.cell_by_label(item, "通期予想").text) for item in RESTATED}
    revenue, operating, net = (RESTATED[k] for k in ("売上高", "営業利益", "当期純利益"))
    gross = oku(pl.cell_by_label("売上総利益", "当期実績").text)
    margin = pct(operating, revenue)
    margin_change = margin - pct(prior["営業利益"], prior["売上高"])

    with deck.batch():                             # one undo step; all or nothing
        # 連結損益計算書: the actual, the change and the rate, row by row.
        for item, value in RESTATED.items():
            change = value - prior[item]
            pl.cell_by_label(item, "当期実績").text = f"{value:,}億円"
            pl.cell_by_label(item, "増減額").text = f"{plus(change)}億"
            pl.cell_by_label(item, "増減率").text = f"{plus(pct(change, prior[item]))}%"

        # The title slide's highlights and KPI, and the progress figures (ids from to_outline).
        for item, figure, rate in (("売上高", "256.13", "256.15"), ("営業利益", "256.17", "256.19"),
                                   ("当期純利益", "256.21", "256.23")):
            value = RESTATED[item]
            replace(deck, figure, f"{old[item]:,}", f"{value:,}")
            old_rate = pct(old[item] - prior[item], prior[item])
            replace(deck, rate, f"▲ {old_rate}%", f"▲ {pct(value - prior[item], prior[item])}%")
        replace(deck, "256.29", "11.9%", f"{margin}%")
        replace(deck, "256.30", "+0.6pt YoY", f"{plus(margin_change)}pt YoY")
        for item, shape in (("売上高", "257.12"), ("営業利益", "257.16"), ("当期純利益", "257.20")):
            replace(deck, shape, f"{pct(old[item], full_year[item])}%",
                    f"{pct(RESTATED[item], full_year[item])}%")

        # セグメント別: the whole difference goes to その他・調整; then the total and shares.
        other = "その他・調整"
        other_revenue = oku(segments.cell_by_label(other, "売上高").text) \
            + revenue - old["売上高"]
        other_operating = oku(segments.cell_by_label(other, "営業利益").text) \
            + operating - old["営業利益"]
        segments.cell_by_label(other, "売上高").text = f"{other_revenue:,}億円"
        segments.cell_by_label(other, "営業利益").text = f"{other_operating:,}億円"
        segments.cell_by_label(other, "利益率").text = f"{pct(other_operating, other_revenue)}%"
        names = [n for n in segments.row_labels()[1:] if n != "合計"]
        sales = {n: oku(segments.cell_by_label(n, "売上高").text) for n in names}
        assert sum(sales.values()) == revenue, "the segments must add up to the new revenue"
        shares = {n: pct(sales[n], revenue) for n in names}
        for name in names:
            segments.cell_by_label(name, "構成比").text = f"{shares[name]}%"
        segments.cell_by_label("合計", "売上高").text = f"{revenue:,}億円"
        segments.cell_by_label("合計", "営業利益").text = f"{operating:,}億円"
        segments.cell_by_label("合計", "利益率").text = f"{margin}%"
        # The total's YoY is operating profit's, as in the original (+16.9%).
        segments.cell_by_label("合計", "YoY").text = \
            f"{plus(pct(operating - prior['営業利益'], prior['営業利益']))}%"

        # Charts: the Q3 points, by series name and category label.
        quarters = deck.shape("257.25").chart
        quarters.series["売上高（億円）"].set_value("Q3", revenue)
        quarters.series["営業利益（億円）"].set_value("Q3", operating)
        margins = deck.shape("257.29").chart
        margins.series["営業利益率（%）"].set_value("Q3", float(margin))
        margins.series["売上総利益率（%）"].set_value("Q3", float(pct(gross, revenue)))
        margins.series["純利益率（%）"].set_value("Q3", float(pct(net, revenue)))
        deck.shape("258.9").chart.series["当期"].set_value("その他", other_revenue)
        # The doughnut's categories are short names: map them to the table's rows.
        doughnut = deck.shape("258.13").chart.series["売上構成"]
        short = {"デジタルソリューション": "デジタルソリューション",
                 "ビジネスPF": "ビジネスプラットフォーム", "グローバル": "グローバル事業",
                 "その他": "その他・調整"}
        doughnut.set_values([float(shares[short[c]]) for c in doughnut.categories])


def main(source: str, target: str) -> None:
    deck = Document.open(source)
    arrived_with = set(deck.validate())            # this deck's generator left a few
    restate(deck)
    added = set(deck.validate()) - arrived_with
    assert not added, added
    deck.save(target)


if __name__ == "__main__":
    main(*sys.argv[1:3])
