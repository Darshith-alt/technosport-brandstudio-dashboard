"""Plotly charts in the TechnoSport blue and orange theme.

Colour meaning is the same in every chart:
  dark blue  = done / complete
  mid blue   = outsourced / in editing (on its way)
  orange     = pending / overdue (needs attention)
  dark orange = finished late (TAT)
  grey       = other / not recorded
"""
from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go

# Theme colours (checked for colour-blind separation)
NAVY = "#0B2545"
BLUE_DARK = "#1C5CAB"
BLUE_MID = "#6DA7EC"
BLUE = "#2A78D6"
ORANGE = "#EB6834"
ORANGE_DARK = "#C2410C"
GREY = "#B7B4AB"
GREY_LIGHT = "#D9DDE5"

INK = "#0B1F3A"
INK_2 = "#52607A"
INK_MUTED = "#8A94A6"
GRID = "#E3E7EE"
AXIS = "#C9CFD9"
SURFACE = "#FFFFFF"   # card background - also used for the 2px gap between stacked segments
FONT = "system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif"

STATUS_COLORS = {"Done": BLUE_DARK, "Outsourced": BLUE_MID, "Pending": ORANGE, "Other": GREY, "Not needed": GREY_LIGHT}
STAGE_COLORS = {"Complete": BLUE_DARK, "In editing": BLUE_MID, "Pending shoot": ORANGE}
TAT_VIEW_COLORS = {"Completed within TAT": BLUE_DARK, "Overdue": ORANGE,
                   "Pending · still within TAT": BLUE_MID, "Pending · not shot yet": GREY_LIGHT}
TAT_COLORS = {"On time": BLUE_DARK, "Late": ORANGE_DARK, "Within TAT": BLUE_MID, "Overdue": ORANGE,
              "Done, no edit date": GREY, "Not started": GREY_LIGHT, "Not needed": GREY_LIGHT}
# Label colour inside each fill (white on dark fills, dark ink on light fills)
_LABEL_ON = {BLUE_DARK: "#FFFFFF", ORANGE: "#FFFFFF", ORANGE_DARK: "#FFFFFF", BLUE_MID: INK, GREY: INK,
             GREY_LIGHT: INK, BLUE: "#FFFFFF"}

PLOTLY_CONFIG = {
    "displaylogo": False,
    "modeBarButtonsToRemove": ["zoom", "pan", "select", "lasso2d", "zoomIn", "zoomOut", "autoScale", "resetScale"],
    "toImageButtonOptions": {"format": "png", "scale": 2},
}

_ROW_PX = 36      # height per bar row -> bars about 22px thick
_BAR_GAP = 0.38


def _style(fig: go.Figure, height: int, legend: bool = True) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=8, r=20, t=40 if legend else 8, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family=FONT, size=13, color=INK_2),
        showlegend=legend,
        legend=dict(orientation="h", xref="container", yref="container", xanchor="left", x=0, yanchor="top", y=1,
                    title_text="",
                    font=dict(color=INK_2, size=12), traceorder="normal"),
        hoverlabel=dict(bgcolor="#FFFFFF", bordercolor=AXIS, font=dict(family=FONT, color=INK, size=12)),
        bargap=_BAR_GAP,
        uniformtext=dict(minsize=11, mode="hide"),   # hide a segment label when it doesn't fit
    )
    fig.update_xaxes(gridcolor=GRID, linecolor=AXIS, zeroline=False, automargin=True,
                     tickfont=dict(color=INK_MUTED, size=12), title_font=dict(color=INK_MUTED, size=12))
    fig.update_yaxes(gridcolor=GRID, linecolor=AXIS, zeroline=False, automargin=True,
                     tickfont=dict(color=INK_2, size=12), title_font=dict(color=INK_MUTED, size=12))
    return fig


def _empty(message: str = "No data for the current filters") -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(text=message, showarrow=False, font=dict(color=INK_MUTED, size=13))
    fig.update_xaxes(visible=False)
    fig.update_yaxes(visible=False)
    return _style(fig, 160, legend=False)


def stacked_bars(table: pd.DataFrame, colors: dict, value_name: str = "SKUs") -> go.Figure:
    """Horizontal stacked bars. `table`: index = categories (top to bottom), columns = segments."""
    if table.empty or table.to_numpy().sum() == 0:
        return _empty()
    cats = [str(c) for c in table.index]
    totals = table.sum(axis=1).replace(0, 1)
    fig = go.Figure()
    for seg in table.columns:
        vals = table[seg]
        if vals.sum() == 0:
            continue
        color = colors.get(seg, GREY)
        share = (vals / totals * 100).round(0)
        fig.add_bar(
            y=cats, x=vals.tolist(), name=str(seg), orientation="h",
            marker=dict(color=color, line=dict(color=SURFACE, width=2)),
            text=[f"{v:,}" if v else "" for v in vals],
            textposition="inside", insidetextanchor="middle", textangle=0,
            textfont=dict(color=_LABEL_ON.get(color, INK), size=12),
            customdata=share.tolist(),
            hovertemplate=f"<b>%{{y}}</b><br>{seg}: %{{x:,}} {value_name} (%{{customdata:.0f}}%)<extra></extra>",
        )
    fig.update_layout(barmode="stack")
    fig.update_yaxes(autorange="reversed", showgrid=False, ticks="", ticksuffix="  ")
    fig.update_xaxes(rangemode="tozero", nticks=6)
    return _style(fig, 84 + _ROW_PX * len(cats))


def single_bars(series: pd.Series, value_name: str = "SKUs", muted: tuple = ("Not recorded", "Not set")) -> go.Figure:
    """One-colour horizontal bars, largest on top, value at the bar tip."""
    series = series[series > 0]
    if series.empty:
        return _empty()
    series = series.sort_values(ascending=False)
    colors = [GREY if str(k) in muted else BLUE for k in series.index]
    fig = go.Figure(go.Bar(
        y=[str(k) for k in series.index], x=series.tolist(), orientation="h",
        marker=dict(color=colors),
        text=[f"{v:,}" for v in series], textposition="outside", cliponaxis=False,
        textfont=dict(color=INK_2, size=12),
        hovertemplate=f"<b>%{{y}}</b><br>%{{x:,}} {value_name}<extra></extra>",
    ))
    fig.update_layout(barcornerradius=4)
    fig.update_yaxes(autorange="reversed", showgrid=False, ticks="", ticksuffix="  ")
    fig.update_xaxes(rangemode="tozero", range=[0, series.max() * 1.18], nticks=6)
    return _style(fig, 56 + _ROW_PX * len(series), legend=False)


def weekly_columns(weekly: pd.DataFrame) -> go.Figure:
    """weekly: index = week start (Timestamp), columns = series (e.g. Inwarded, Shot)."""
    if weekly.empty or weekly.to_numpy().sum() == 0:
        return _empty()
    labels = [d.strftime("%d %b") for d in weekly.index]
    palette = {"Inwarded": BLUE_MID, "Shot": BLUE_DARK, "Edited": ORANGE}
    fig = go.Figure()
    for name in weekly.columns:
        fig.add_bar(
            x=labels, y=weekly[name].tolist(), name=name,
            marker=dict(color=palette.get(name, BLUE)),
            hovertemplate=f"Week of %{{x}}<br>{name}: %{{y:,}} SKUs<extra></extra>",
        )
    fig.update_layout(barmode="group", bargap=0.35, bargroupgap=0.08, barcornerradius=4)
    fig.update_xaxes(showgrid=False, type="category", title_text="Week starting (Monday)")
    fig.update_yaxes(title_text="SKUs", rangemode="tozero")
    return _style(fig, 340)


def days_vs_target(avg: pd.Series, target: pd.Series) -> go.Figure:
    """Average days taken per deliverable (bar) with the TAT target (black tick). Same axis: days."""
    avg = avg.dropna()
    if avg.empty:
        return _empty("No finished deliverables with edit dates yet")
    cats = [str(c) for c in avg.index]
    tgt = target.reindex(avg.index)
    colors = [ORANGE_DARK if (pd.notna(t) and a > t) else BLUE_DARK for a, t in zip(avg, tgt)]
    fig = go.Figure()
    fig.add_bar(
        y=cats, x=avg.round(1).tolist(), orientation="h", name="Average days taken",
        marker=dict(color=colors), text=[f"{v:.1f} d" for v in avg], textposition="outside", cliponaxis=False,
        textfont=dict(color=INK_2, size=12),
        customdata=tgt.tolist(),
        hovertemplate="<b>%{y}</b><br>Average: %{x:.1f} days<br>TAT: %{customdata:g} days<extra></extra>",
    )
    fig.add_scatter(
        y=cats, x=tgt.tolist(), mode="markers", name="TAT target",
        marker=dict(symbol="line-ns", size=30, line=dict(width=3, color=INK)),
        hovertemplate="<b>%{y}</b><br>TAT target: %{x:g} days<extra></extra>",
    )
    fig.update_layout(barcornerradius=4)
    fig.update_yaxes(autorange="reversed", showgrid=False, ticks="", ticksuffix="  ")
    top = max(avg.max(), tgt.max() if tgt.notna().any() else 0)
    fig.update_xaxes(rangemode="tozero", range=[0, top * 1.2], nticks=6, title_text="Days")
    return _style(fig, int(96 + _ROW_PX * len(cats) * 1.3))


def due_columns(counts: pd.Series, hover: pd.Series) -> go.Figure:
    """Pending items by TAT due date. counts: index = 'Overdue', dates..., 'Later'."""
    if counts.empty or counts.sum() == 0:
        return _empty("Nothing pending")
    labels = [str(k) for k in counts.index]
    colors = [ORANGE if k == "Overdue" else (GREY if k == "Not shot yet" else BLUE_MID) for k in labels]
    fig = go.Figure(go.Bar(
        x=labels, y=counts.tolist(), marker=dict(color=colors),
        text=[f"{v:,}" for v in counts], textposition="outside", cliponaxis=False,
        textfont=dict(color=INK_2, size=12),
        customdata=hover.reindex(counts.index).fillna("").tolist(),
        hovertemplate="<b>%{x}</b><br>%{y:,} pending<br>%{customdata}<extra></extra>",
    ))
    fig.update_layout(barcornerradius=4, bargap=0.35)
    fig.update_xaxes(showgrid=False, type="category", title_text="TAT due date")
    fig.update_yaxes(title_text="Pending deliverables", rangemode="tozero", range=[0, counts.max() * 1.18])
    return _style(fig, 320, legend=False)


def stacked_columns(table: pd.DataFrame, colors: dict, x_title: str = "", value_name: str = "SKUs") -> go.Figure:
    """Vertical stacked columns. `table`: index = x categories (left to right), columns = segments."""
    if table.empty or table.to_numpy().sum() == 0:
        return _empty()
    xs = [d.strftime("%d %b") if hasattr(d, "strftime") else str(d) for d in table.index]
    fig = go.Figure()
    for seg in table.columns:
        vals = table[seg]
        if vals.sum() == 0:
            continue
        color = colors.get(seg, GREY)
        fig.add_bar(
            x=xs, y=vals.tolist(), name=str(seg),
            marker=dict(color=color, line=dict(color=SURFACE, width=2)),
            text=[f"{v:,}" if v else "" for v in vals], textposition="inside", insidetextanchor="middle",
            textangle=0, textfont=dict(color=_LABEL_ON.get(color, INK), size=12),
            hovertemplate=f"<b>%{{x}}</b><br>{seg}: %{{y:,}} {value_name}<extra></extra>",
        )
    fig.update_layout(barmode="stack", bargap=0.35)
    fig.update_xaxes(showgrid=False, type="category", title_text=x_title)
    fig.update_yaxes(title_text=value_name, rangemode="tozero")
    return _style(fig, 360)
