"""Plotly charts. Each chart also carries plain data so the PDF can fall back to native charts."""
from __future__ import annotations

from dataclasses import dataclass, field

import pandas as pd

NAVY, TEAL, AMBER = "#1F3A5F", "#2A9D8F", "#E9A23B"


@dataclass
class ChartSpec:
    key: str
    title: str
    labels: list[str]
    series: dict[str, list[float]]
    width: int = 1100
    height: int = 520
    colors: list[str] = field(default_factory=lambda: [NAVY, TEAL])
    percent: bool = False

    def figure(self):
        import plotly.graph_objects as go

        fig = go.Figure()
        horizontal = self.key != "by_department"
        for (name, values), color in zip(self.series.items(), self.colors):
            kwargs = dict(
                x=values if horizontal else self.labels,
                y=self.labels if horizontal else values,
                name=name,
                marker_color=color,
                text=[f"{v:.1f}%" if self.percent else f"{v:,.0f}" for v in values],
                textposition="outside",
            )
            if horizontal:
                kwargs["orientation"] = "h"
            fig.add_trace(go.Bar(**kwargs))
        fig.update_layout(
            title=self.title,
            template="plotly_white",
            barmode="group",
            width=self.width,
            height=self.height,
            margin=dict(l=40, r=60, t=70, b=40),
            showlegend=len(self.series) > 1,
        )
        if horizontal:
            fig.update_yaxes(autorange="reversed", automargin=True)
        if self.percent:
            fig.update_xaxes(range=[0, 115])
        return fig


def build_chart_specs(detail: pd.DataFrame) -> list[ChartSpec]:
    detail = detail.copy()
    detail["Path"] = detail["From Department"] + " -> " + detail["To Department"]

    out = detail.groupby("From Department")["Transfers"].sum()
    inc = detail.groupby("To Department")["Transfers"].sum()
    depts = sorted(set(out.index) | set(inc.index))
    by_dept = ChartSpec(
        "by_department", "Transfers by Department", depts,
        {"Outbound (transferred from)": [float(out.get(d, 0)) for d in depts],
         "Inbound (transferred to)": [float(inc.get(d, 0)) for d in depts]},
    )

    pct = detail.sort_values("Answer %", ascending=False)
    answer = ChartSpec(
        "answer_pct", "Answer Percentage by Transfer Path", pct["Path"].tolist(),
        {"Answer %": pct["Answer %"].astype(float).tolist()},
        height=max(420, 120 + 28 * len(pct)), colors=[TEAL], percent=True,
    )

    top = detail.nlargest(10, "Transfers")
    top10 = ChartSpec(
        "top10", "Top 10 Transfer Paths", top["Path"].tolist(),
        {"Transfers": top["Transfers"].astype(float).tolist()}, colors=[AMBER],
        height=max(420, 120 + 40 * len(top)),
    )
    return [by_dept, answer, top10]
