"""
TechnoSport Brand Studio - Photo / Video Shoot Dashboard (Streamlit)

Run locally:   streamlit run app.py
Data source (first one found is used):
  1. File uploaded in the sidebar ("Use a different file")
  2. Google Drive / Sheets file set in .streamlit/secrets.toml  ([sheet] file_id = ...)
  3. TRACKER_FILE environment variable, or the first .xlsx in the data/ folder
"""
from __future__ import annotations

import hmac
import html
import os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

import charts
import tracker as tk

IST = ZoneInfo("Asia/Kolkata")
APP_DIR = Path(__file__).parent

st.set_page_config(page_title="Brand Studio · Shoot Dashboard", page_icon="📸", layout="wide",
                   initial_sidebar_state="auto")

# ---------------------------------------------------------------------------
# Styling
# ---------------------------------------------------------------------------
st.html(f"""
<style>
  .block-container {{ padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1440px; }}
  .ts-header {{ background: {charts.NAVY}; border-radius: 14px; padding: 18px 24px;
               border-bottom: 4px solid {charts.ORANGE}; display: flex; justify-content: space-between;
               align-items: flex-end; flex-wrap: wrap; gap: 10px 24px; margin-bottom: 14px; }}
  .ts-eyebrow {{ color: #F7B08F; font-size: .78rem; letter-spacing: .08em; text-transform: uppercase; font-weight: 600; }}
  .ts-header h1 {{ color: #FFFFFF; font-size: 1.6rem; font-weight: 650; margin: 2px 0 0; padding: 0; line-height: 1.25; }}
  .ts-meta {{ color: #C9D6EA; font-size: .82rem; text-align: right; line-height: 1.5; }}
  .kpi-grid {{ display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin: 4px 0 16px; }}
  @media (max-width: 900px) {{ .kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }} }}
  .kpi {{ background: #FFFFFF; border: 1px solid rgba(11,37,69,.10); border-radius: 12px; padding: 14px 16px; }}
  .kpi .label {{ color: {charts.INK_2}; font-size: .8rem; font-weight: 500; }}
  .kpi .value {{ color: {charts.NAVY}; font-size: 1.65rem; font-weight: 650; line-height: 1.25; margin-top: 2px; }}
  .kpi .note {{ color: {charts.INK_MUTED}; font-size: .76rem; margin-top: 2px; }}
  .kpi-section {{ color: {charts.NAVY}; font-weight: 650; font-size: .9rem; margin: 6px 0 6px; }}
  .kpi-section span {{ color: {charts.INK_MUTED}; font-weight: 400; font-size: .8rem; }}
  .tat-grid {{ display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-bottom: 16px; }}
  .tat-grid.n1 {{ grid-template-columns: 1fr; }}
  .tat-grid.n2, .tat-grid.n4 {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
  @media (min-width: 1800px) {{ .tat-grid.n4 {{ grid-template-columns: repeat(4, minmax(0, 1fr)); }} }}
  @media (max-width: 1000px) {{ .tat-grid.n3 {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }} }}
  @media (max-width: 700px) {{ .tat-grid {{ grid-template-columns: 1fr !important; }} }}
  .tat-card {{ background: #FFFFFF; border: 1px solid rgba(11,37,69,.10); border-radius: 12px; padding: 14px 16px; }}
  .tat-card .head {{ display: flex; justify-content: space-between; align-items: center; gap: 8px; }}
  .tat-card .title {{ color: {charts.NAVY}; font-weight: 650; font-size: .95rem; }}
  .tat-card .pill {{ background: #E3EEFB; color: #1C4F8F; font-size: .72rem; font-weight: 600; padding: 2px 9px;
                    border-radius: 999px; white-space: nowrap; }}
  .tat-card .total {{ margin-top: 6px; color: {charts.INK_2}; font-size: .82rem; }}
  .tat-card .total b {{ color: {charts.NAVY}; font-size: 1.7rem; font-weight: 650; margin-right: 4px; }}
  .tat-card .mini {{ display: flex; gap: 2px; height: 8px; border-radius: 4px; overflow: hidden; margin: 10px 0 12px;
                    background: #EEF2F7; }}
  .tat-card .mini span {{ display: block; height: 100%; }}
  .tat-card .stats {{ display: flex; flex-direction: column; gap: 6px; }}
  .tat-card .stat {{ background: #F5F7FB; border-radius: 8px; padding: 8px 12px; display: flex;
                    justify-content: space-between; align-items: center; gap: 10px; }}
  .tat-card .stat .k {{ color: {charts.INK}; font-size: .82rem; font-weight: 500; display: flex; gap: 8px; align-items: center; }}
  .tat-card .stat .k i {{ flex: none; width: 9px; height: 9px; border-radius: 2px; display: inline-block; }}
  .tat-card .stat .s {{ color: {charts.INK_MUTED}; font-size: .72rem; margin: 1px 0 0 17px; }}
  .tat-card .stat .v {{ color: {charts.NAVY}; font-size: 1.35rem; font-weight: 650; white-space: nowrap; }}
  .tat-card .stat .v .p, .kpi .value .p {{ color: {charts.INK_2}; font-size: .8rem; font-weight: 500; margin-left: 8px;
                                        background: #E9EEF6; border-radius: 999px; padding: 2px 8px; vertical-align: middle; }}
  .kpi .label i {{ display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 6px; }}
  .tat-table-wrap {{ background: #FFFFFF; border: 1px solid rgba(11,37,69,.10); border-radius: 12px; overflow-x: auto;
                    margin-bottom: 18px; }}
  .tat-table {{ width: 100%; border-collapse: collapse; font-size: .86rem; color: {charts.INK}; }}
  .tat-table th {{ background: #F1F5FB; color: {charts.INK_2}; font-weight: 600; text-align: left; padding: 10px 12px;
                  white-space: nowrap; }}
  .tat-table th i {{ display: inline-block; width: 9px; height: 9px; border-radius: 2px; margin-right: 6px; }}
  .tat-table td {{ padding: 9px 12px; border-top: 1px solid #EEF2F7; white-space: nowrap; }}
  .tat-table .num {{ text-align: right; font-variant-numeric: tabular-nums; }}
  .tat-table td.num {{ font-weight: 650; color: {charts.NAVY}; }}
  .tat-table .pc {{ text-align: left; color: {charts.INK_2}; font-variant-numeric: tabular-nums; width: 1%; }}
  .tat-table .step {{ font-weight: 600; color: {charts.NAVY}; }}
  .tat-table tr.total td {{ background: #F7F9FC; border-top: 2px solid #DDE5F0; }}
  .tat-card .foot {{ color: {charts.INK_MUTED}; font-size: .72rem; margin-top: 8px; }}
  @media (max-width: 640px) {{ .ts-meta {{ text-align: left; }}
                              .block-container {{ padding-top: 3.2rem; }} }}
  .card-title {{ color: {charts.NAVY}; font-weight: 650; font-size: 1rem; margin: 0; }}
  .card-note {{ color: {charts.INK_2}; font-size: .82rem; margin: 2px 0 4px; }}
  [class*="st-key-card"] {{ background: #FFFFFF; }}
  .chip {{ display: inline-block; padding: 4px 10px; margin: 0 6px 6px 0; border-radius: 999px; font-size: .8rem;
           background: #FFF1EA; color: #9A3A12; border: 1px solid #F7C7B1; }}
  .chip b {{ color: {charts.NAVY}; }}
  .section-note {{ color: {charts.INK_2}; font-size: .85rem; margin: 0 0 8px; }}
</style>
""")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def secret(*path, default=None):
    """Read a value from st.secrets without failing when secrets.toml doesn't exist."""
    try:
        node = st.secrets
        for p in path:
            node = node[p]
        return node
    except Exception:
        return default


def password_gate() -> None:
    """Ask for APP_PASSWORD (if set in secrets) before showing anything."""
    expected = secret("APP_PASSWORD")
    if not expected or st.session_state.get("auth_ok"):
        return
    st.markdown(f"<div class='ts-header'><div><div class='ts-eyebrow'>TechnoSport · Brand Studio</div>"
                f"<h1>Shoot Dashboard</h1></div></div>", unsafe_allow_html=True)
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Open dashboard", type="primary"):
            if hmac.compare_digest(pw, str(expected)):
                st.session_state["auth_ok"] = True
                st.rerun()
            st.error("Wrong password.")
    st.stop()


@st.cache_data(ttl=600, show_spinner="Loading the tracker from Google Drive…")
def fetch_from_drive(file_id: str, access: str) -> tuple[bytes, dict, str]:
    if access == "public_link":
        data, meta = tk.download_public(file_id)
    else:
        data, meta = tk.download_with_service_account(file_id, dict(st.secrets["gcp_service_account"]))
    return data, meta, datetime.now(IST).strftime("%d %b %Y, %I:%M %p")


@st.cache_data(show_spinner="Reading the tracker…")
def parse_cached(data: bytes, tracker_name: str, today: str):
    return tk.parse_tracker(data, tk.TRACKERS[tracker_name], today=pd.Timestamp(today).date())


@st.cache_data(show_spinner=False)
def sheets_in(data: bytes) -> list[str]:
    return tk.list_sheets(data)


def load_workbook() -> tuple[bytes, str, str]:
    """Return (xlsx bytes, source label, loaded-at text)."""
    with st.sidebar.expander("Use a different file"):
        upload = st.file_uploader("Upload the tracker (.xlsx)", type=["xlsx"],
                                  help="Only for a quick test. The file stays in this browser session.")
    if upload is not None:
        return upload.getvalue(), f"Uploaded file · {upload.name}", datetime.now(IST).strftime("%d %b %Y, %I:%M %p")

    file_ref = secret("sheet", "file_id") or secret("sheet", "url")
    if file_ref:
        access = str(secret("sheet", "access", default="service_account"))
        if access != "public_link" and secret("gcp_service_account") is None:
            st.error("`[gcp_service_account]` is missing in secrets. Add the service account key, "
                     "or set `access = \"public_link\"` under `[sheet]`.")
            st.stop()
        try:
            data, meta, at = fetch_from_drive(tk.extract_file_id(str(file_ref)), access)
        except Exception as exc:  # show a readable message instead of a stack trace
            st.error(f"Could not load the tracker from Google Drive.\n\n{exc}")
            st.stop()
        label = f"Google Drive · {meta.get('name', 'tracker')}"
        if meta.get("modifiedTime"):
            changed = pd.Timestamp(meta["modifiedTime"]).tz_convert(IST).strftime("%d %b, %I:%M %p")
            label += f" · sheet edited {changed}"
        return data, label, at

    local = os.environ.get("TRACKER_FILE") or next(iter(sorted((APP_DIR / "data").glob("*.xlsx"))), None)
    if local and Path(local).exists():
        path = Path(local)
        return path.read_bytes(), f"Local file · {path.name}", datetime.now(IST).strftime("%d %b %Y, %I:%M %p")

    st.markdown("<div class='ts-header'><div><div class='ts-eyebrow'>TechnoSport · Brand Studio</div>"
                "<h1>Shoot Dashboard</h1></div></div>", unsafe_allow_html=True)
    st.info("No tracker connected yet. Upload the tracker in the sidebar (**Use a different file**), put the "
            ".xlsx in the `data/` folder, or connect the Google Sheet in `.streamlit/secrets.toml` (see README).")
    st.stop()


def fmt_days(v: float) -> str:
    return "–" if pd.isna(v) else f"{v:.1f} days"


def pct(part: int, whole: int) -> str:
    return f"{part / whole * 100:.0f}%" if whole else "–"


def card(key: str, title: str, note: str = ""):
    box = st.container(border=True, key=f"card_{key}")
    box.markdown(f"<p class='card-title'>{html.escape(title)}</p>"
                 + (f"<p class='card-note'>{html.escape(note)}</p>" if note else ""), unsafe_allow_html=True)
    return box


def show_chart(box, fig, key: str, table: pd.DataFrame | None = None) -> None:
    box.plotly_chart(fig, width="stretch", theme=None, config=charts.PLOTLY_CONFIG, key=f"fig_{key}")
    if table is not None and not table.empty:
        with box.expander("View as table"):
            st.dataframe(table, width="stretch", hide_index=False)


def stage_table(df: pd.DataFrame, dim: str, order: list | None = None) -> pd.DataFrame:
    t = pd.crosstab(df[dim], df["stage"]).reindex(columns=tk.STAGE_ORDER, fill_value=0)
    if order:
        t = t.reindex([o for o in order if o in t.index] + [i for i in t.index if i not in order])
    else:
        t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    t.index.name = None
    return t


def multiselect_filter(container, label: str, series: pd.Series, key: str) -> list:
    options = sorted(series.dropna().unique().tolist(), key=lambda x: str(x))
    return container.multiselect(label, options, placeholder="All", key=key)


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
password_gate()

st.sidebar.markdown("### 📸 Brand Studio")
data, source_label, loaded_at = load_workbook()

available = [name for name, cfg in tk.TRACKERS.items() if cfg.sheet in sheets_in(data)]
if not available:
    st.error(f"None of these sheets were found in the file: {', '.join(c.sheet for c in tk.TRACKERS.values())}.")
    st.stop()
tracker_name = st.sidebar.radio("Tracker", available, index=0)
cfg = tk.TRACKERS[tracker_name]
outsourced_ok = st.sidebar.toggle(
    "Count outsourced as complete", value=True,
    help="On: a deliverable marked 'Outsource' counts as handled. Off: only 'Done' counts as complete.")

if st.sidebar.button("Refresh data", icon=":material/refresh:", width="stretch"):
    st.cache_data.clear()
    st.rerun()
st.sidebar.caption(f"**Source:** {source_label}  \n**Loaded:** {loaded_at}  \nGoogle Drive data refreshes every 10 minutes.")

today = datetime.now(IST).date()
try:
    base_df, issues = parse_cached(data, tracker_name, today.isoformat())
except Exception as exc:
    st.error(f"Could not read the '{cfg.sheet}' sheet: {exc}")
    st.stop()
if base_df.empty:
    st.warning(f"The '{cfg.sheet}' sheet has no SKU rows yet.")
    st.stop()
df_all = tk.add_stage(base_df, cfg, outsourced_ok)

# Header
st.markdown(
    f"<div class='ts-header'><div><div class='ts-eyebrow'>TechnoSport · Brand Studio</div>"
    f"<h1>{html.escape(cfg.name)} Dashboard</h1></div>"
    f"<div class='ts-meta'>{html.escape(source_label)}<br>Loaded {html.escape(loaded_at)}</div></div>",
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Filters (one row, above everything they affect)
# ---------------------------------------------------------------------------
with st.container(border=True, key="card_filters"):
    c1, c5, c2, c3, c4 = st.columns([1.3, 1, 1, 1, 1])
    dmin, dmax = df_all["inward_date"].min(), df_all["inward_date"].max()
    has_dates = pd.notna(dmin) and pd.notna(dmax)
    if has_dates:
        dr = c1.date_input("Inward date", value=(dmin.date(), dmax.date()), min_value=dmin.date(),
                           max_value=dmax.date(), format="DD/MM/YYYY", key=f"{tracker_name}_dates")
    f_style = multiselect_filter(c5, "Style code", df_all["style"], f"{tracker_name}_style")
    f_project = multiselect_filter(c2, "Project (shoot type)", df_all["project"], f"{tracker_name}_project")
    f_cat = multiselect_filter(c3, "Super category", df_all["super_category"], f"{tracker_name}_cat")
    f_budget = multiselect_filter(c4, "Budget DPT", df_all["budget"], f"{tracker_name}_budget")
    with st.expander("More filters"):
        m1, m2, m3, m4, m5 = st.columns(5)
        f_vertical = multiselect_filter(m1, "Vertical", df_all["vertical"], f"{tracker_name}_vertical")
        f_ideal = multiselect_filter(m2, "Ideal for", df_all["ideal_for"], f"{tracker_name}_ideal")
        models = pd.Series(sorted({p for v in df_all["model"] for p in tk.split_people(v)}))
        f_model = multiselect_filter(m3, "Model", models, f"{tracker_name}_model")
        stylists = pd.Series(sorted({p for v in df_all["stylist"] for p in tk.split_people(v)}))
        f_stylist = multiselect_filter(m4, "Stylist", stylists, f"{tracker_name}_stylist")
        f_search = m5.text_input("Search SKU / style / colour", key=f"{tracker_name}_search").strip().lower()

mask = pd.Series(True, index=df_all.index)
if has_dates and isinstance(dr, (tuple, list)) and len(dr) == 2:
    start, end = pd.Timestamp(dr[0]), pd.Timestamp(dr[1])
    in_range = df_all["inward_date"].between(start, end)
    full_range = start <= dmin and end >= dmax
    mask &= in_range | (df_all["inward_date"].isna() & full_range)
for values, column in [(f_style, "style"), (f_project, "project"), (f_cat, "super_category"), (f_budget, "budget"),
                       (f_vertical, "vertical"), (f_ideal, "ideal_for")]:
    if values:
        mask &= df_all[column].isin(values)
if f_model:
    mask &= df_all["model"].map(lambda v: bool(set(tk.split_people(v)) & set(f_model)))
if f_stylist:
    mask &= df_all["stylist"].map(lambda v: bool(set(tk.split_people(v)) & set(f_stylist)))
if f_search:
    hay = (df_all["sku"].fillna("") + " " + df_all["style"].fillna("") + " " + df_all["colour"].fillna("")).str.lower()
    mask &= hay.str.contains(f_search, regex=False)
df = df_all[mask]

if df.empty:
    st.info("No SKUs match the filters. Clear a filter to see data.")
    st.stop()

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
n = len(df)
n_complete = int((df["stage"] == "Complete").sum())
n_shot = int(df["shoot_status"].isin(tk.SHOT_STATUSES).sum())
n_wait_shoot = int((df["stage"] == "Pending shoot").sum())
n_editing = int((df["stage"] == "In editing").sum())
oldest_wait = df.loc[df["stage"] == "Pending shoot", "days_since_inward"].max()
tat_rules = {r.deliverable: r for r in cfg.tat}


def tile(label: str, value: str, note: str, tag: str = "", dot: str = "") -> str:
    dot_html = f"<i style='background:{dot}'></i>" if dot else ""
    return (f"<div class='kpi'><div class='label'>{dot_html}{html.escape(label)}</div><div class='value'>{value}</div>"
            f"<div class='note'>{html.escape(note)}</div>{tag}</div>")


status_tiles = [
    tile("SKUs inwarded", f"{n:,}", f"{df['style'].nunique():,} style{'s' if df['style'].nunique() != 1 else ''}"),
    tile("Shot", f"{n_shot:,}", f"{pct(n_shot, n)} of inward"),
    tile("Waiting for shoot", f"{n_wait_shoot:,}",
         f"oldest waiting {int(oldest_wait)} days" if n_wait_shoot and pd.notna(oldest_wait) else "none waiting"),
    tile("In editing", f"{n_editing:,}", "shot, deliverables not all done"),
]


def tat_step_table() -> str:
    """HTML table: one row per TAT step with counts and % of SKUs, plus an all-steps total."""
    rows, tot = [], {"n": 0, "on": 0, "over": 0, "within": 0}
    for i, r in enumerate([r for r in cfg.tat if f"{r.deliverable}__tat_outcome" in df], start=1):
        oc, target = df[f"{r.deliverable}__tat_outcome"], df[f"{r.deliverable}__tat_target"]
        n_rows = len(oc)
        on, over, within = int((oc == "On time").sum()), int((oc == "Overdue").sum()), int((oc == "Within TAT").sum())
        for k, v in (("n", n_rows), ("on", on), ("over", over), ("within", within)):
            tot[k] += v
        start_word = "Shoot" if r.start == "shoot" else "Inward"
        tgt = f"{target.min():g}" if target.min() == target.max() else f"{target.min():g}–{target.max():g}"
        rows.append((f"{i}. {start_word} → {cfg.label(r.deliverable)}", f"{tgt} days", n_rows, on, over, within))
    if not rows:
        return ""

    def cells(n_rows, on, over, within):
        return (f"<td class='num'>{n_rows:,}</td>"
                f"<td class='num'>{on:,}</td><td class='pc'>{pct(on, n_rows)}</td>"
                f"<td class='num'>{over:,}</td><td class='pc'>{pct(over, n_rows)}</td>"
                f"<td class='num'>{within:,}</td><td class='pc'>{pct(within, n_rows)}</td>")

    body = "".join(f"<tr><td class='step'>{html.escape(name)}</td><td>{tat}</td>{cells(*vals)}</tr>"
                   for name, tat, *vals in rows)
    total = (f"<tr class='total'><td class='step'>All steps</td><td></td>"
             f"{cells(tot['n'], tot['on'], tot['over'], tot['within'])}</tr>")
    dot = lambda c: f"<i style='background:{c}'></i>"
    head = ("<tr><th>Step</th><th>TAT</th><th class='num'>SKUs</th>"
            f"<th class='num' colspan='2'>{dot(charts.BLUE_DARK)}Completed within TAT</th>"
            f"<th class='num' colspan='2'>{dot(charts.ORANGE)}Overdue</th>"
            f"<th class='num' colspan='2'>{dot(charts.BLUE_MID)}Pending, still within TAT</th></tr>")
    return (f"<div class='tat-table-wrap'><table class='tat-table'><thead>{head}</thead>"
            f"<tbody>{body}{total}</tbody></table></div>")

n_styles = df["style"].nunique()


def tat_card(d: str) -> str:
    """Card for one deliverable: total, completed within TAT, pending, pending still within TAT."""
    label = cfg.label(d)
    rule = tat_rules.get(d)
    if rule and f"{d}__tat_outcome" in df:
        oc, target, due = df[f"{d}__tat_outcome"], df[f"{d}__tat_target"], df[f"{d}__tat_due"]
        counts = oc.value_counts()
        within, late = int(counts.get("On time", 0)), int(counts.get("Late", 0))
        no_date = int(counts.get("Done, no edit date", 0))
        open_ok, overdue = int(counts.get("Within TAT", 0)), int(counts.get("Overdue", 0))
        not_started = int(counts.get("Not started", 0))
        pending = open_ok + overdue + not_started
        start_word = "Shoot" if rule.start == "shoot" else "Inward"
        tgt = f"{target.min():g}" if target.min() == target.max() else f"{target.min():g}–{target.max():g}"
        pill = f"TAT {tgt} days"
        segments = [(within, charts.BLUE_DARK), (late + no_date, charts.GREY),
                    (open_ok, charts.BLUE_MID), (overdue, charts.ORANGE), (not_started, charts.GREY_LIGHT)]
        n_card = len(df)
        stats = [
            ("Completed within TAT", f"{within:,}", charts.BLUE_DARK, pct(within, n_card)),
            ("Overdue", f"{overdue:,}", charts.ORANGE, pct(overdue, n_card)),
            ("Pending, still within TAT", f"{open_ok:,}", charts.BLUE_MID, pct(open_ok, n_card)),
        ]
        avg = df.loc[oc.isin(["On time", "Late"]), f"{d}__tat_days"].mean()
        foot = []
        if pd.notna(avg):
            foot.append(f"Average {avg:.1f} days")
        if late + no_date:
            foot.append(f"{late + no_date:,} more completed (late or without a date)")
    else:  # tracker without TAT columns: done vs pending only
        status = df[f"{d}__status"]
        done = int(status.isin(["Done", "Outsourced", "Not needed"]).sum())
        pending = len(df) - done
        start_word, pill = "Inward", "No TAT set"
        segments = [(done, charts.BLUE_DARK), (pending, charts.ORANGE)]
        stats = [("Completed", f"{done:,}", charts.BLUE_DARK, pct(done, len(df))),
                 ("Pending", f"{pending:,}", charts.ORANGE, pct(pending, len(df)))]
        avg = df[f"{d}__days"].mean()
        foot = [f"Average {avg:.1f} days from inward"] if pd.notna(avg) else []
    total = len(df)
    bar = "".join(f"<span style='width:{v / total * 100:.2f}%;background:{c}'></span>" for v, c in segments if v)
    stat_html = "".join(
        f"<div class='stat'><div class='k'><i style='background:{c}'></i>{html.escape(k)}</div>"
        f"<div class='v'>{html.escape(v)}<span class='p'>{html.escape(p_)}</span></div></div>"
        for k, v, c, p_ in stats)
    return (
        f"<div class='tat-card'><div class='head'><span class='title'>{start_word} → {html.escape(label)}</span>"
        f"<span class='pill'>{html.escape(pill)}</span></div>"
        f"<div class='total'><b>{total:,}</b>SKU{'s' if total != 1 else ''} · {n_styles:,} style{'s' if n_styles != 1 else ''}</div>"
        f"<div class='mini'>{bar}</div><div class='stats'>{stat_html}</div>"
        + (f"<div class='foot'>{html.escape(' · '.join(foot))}</div>" if foot else "") + "</div>"
    )


tat_steps = [r.deliverable for r in cfg.tat if f"{r.deliverable}__tat_outcome" in df] or list(cfg.deliverables)
tat_cards = [tat_card(d) for d in tat_steps]

kpi_html = (
    f"<div class='kpi-grid' style='grid-template-columns: repeat({len(status_tiles)}, minmax(0, 1fr))'>"
    + "".join(status_tiles) + "</div>"
    + "<div class='kpi-section'>Turnaround vs TAT <span>· SKU counts, calendar days</span></div>"
    + f"<div class='tat-grid n{min(len(tat_cards), 4)}'>"
    + "".join(tat_cards) + "</div>"
    + ("<div class='kpi-section'>TAT summary by step <span>· % = share of SKUs in each step</span></div>"
       + tat_step_table() if cfg.tat else "")
)
st.markdown(kpi_html, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
has_tat = bool(cfg.tat) and any(f"{r.deliverable}__tat_outcome" in df for r in cfg.tat)
n_overdue = int(sum((df[f"{r.deliverable}__tat_outcome"] == "Overdue").sum()
                    for r in cfg.tat if f"{r.deliverable}__tat_outcome" in df))
tab_names = ["Overview"] + ([f"TAT ({n_overdue:,} overdue)"] if has_tat else []) + [
    "Breakdowns", "Team", f"Pending ({n - n_complete:,})", f"Data quality ({len(issues):,})", "All data"]
focus = None
if has_tat:
    focus_options = {cfg.label(r.deliverable): r.deliverable for r in cfg.tat if f"{r.deliverable}__tat_outcome" in df}
    pick = st.segmented_control("Charts show TAT for", list(focus_options), default=list(focus_options)[0],
                                required=True, key=f"{tracker_name}_focus")
    focus = focus_options.get(pick) or next(iter(focus_options.values()))
    focus_label = cfg.label(focus)
TAT_NOTE = "Items completed after the TAT or without an edit date aren't shown."
tabs = st.tabs(tab_names)
tab_overview = tabs[0]
tab_tat = tabs[1] if has_tat else None
tab_break, tab_team, tab_pending, tab_quality, tab_data = tabs[-5:]

with tab_overview:
    left, right = st.columns([1.15, 1])
    if focus:
        with left:
            box = card("pipeline", "TAT status by step",
                       "Completed within TAT, overdue, and pending still within TAT. " + TAT_NOTE)
            t = tk.tat_view_by_deliverable(df, cfg)
            show_chart(box, charts.stacked_bars(t, charts.TAT_VIEW_COLORS), "pipeline", t)
            box = card("weekly", f"{focus_label}: by inward week", "SKUs grouped by the week they were inwarded.")
            t = tk.tat_view_table(df, focus, "inward_week").sort_index()
            tt = t.copy()
            tt.index = [d.strftime("%d %b %Y") for d in tt.index]
            show_chart(box, charts.stacked_columns(t, charts.TAT_VIEW_COLORS, "Inward week (starting Monday)"), "weekly", tt)
        with right:
            box = card("category", f"{focus_label}: by super category")
            t = tk.tat_view_table(df, focus, "super_category")
            show_chart(box, charts.stacked_bars(t, charts.TAT_VIEW_COLORS), "category", t)
            box = card("project", f"{focus_label}: by project (shoot type)")
            t = tk.tat_view_table(df, focus, "project")
            show_chart(box, charts.stacked_bars(t, charts.TAT_VIEW_COLORS), "project", t)
    else:  # tracker without TAT (e.g. Video Shoot): progress by stage
        sl = tk.status_long(df, cfg)
        status_tbl = sl.pivot_table(index="Step", columns="Status", values="SKUs", aggfunc="sum", fill_value=0,
                                    observed=True)
        status_tbl = status_tbl.reindex(columns=[c for c in tk.STATUS_ORDER if c in status_tbl.columns])
        weekly = pd.concat([
            df.groupby("inward_week").size().rename("Inwarded"),
            df.groupby("shoot_week").size().rename("Shot"),
        ], axis=1, sort=True).fillna(0).astype(int).sort_index()
        wt = weekly.copy()
        wt.index = wt.index.strftime("%d %b %Y")
        with left:
            box = card("pipeline", "Pipeline status by step", "SKUs at each step: shoot first, then each deliverable.")
            show_chart(box, charts.stacked_bars(status_tbl, charts.STATUS_COLORS), "pipeline", status_tbl)
            box = card("weekly", "Inwarded vs shot, by week", "SKUs received in the studio and SKUs shot, per week.")
            show_chart(box, charts.weekly_columns(weekly), "weekly", wt)
        with right:
            box = card("category", "Progress by super category", "Complete = every deliverable done.")
            t = stage_table(df, "super_category")
            show_chart(box, charts.stacked_bars(t, charts.STAGE_COLORS), "category", t)
            box = card("project", "Progress by project (shoot type)")
            t = stage_table(df, "project")
            show_chart(box, charts.stacked_bars(t, charts.STAGE_COLORS), "project", t)

if tab_tat is not None:
    with tab_tat:
        defaults = " / ".join(f"{r.default_days:g}" for r in cfg.tat)
        start_txt = ", ".join(f"{'shoot' if r.start == 'shoot' else 'inward'} date for {cfg.label(r.deliverable)}"
                              for r in cfg.tat)
        st.markdown(
            "<p class='section-note'>Calendar days from the start date (" + html.escape(start_txt) + ") to each "
            "deliverable's Edit Date, compared with the TAT columns in the sheet "
            f"(blank TAT cells use {defaults} days). <b>TAT date</b> = start date + TAT days. "
            "<b>Overdue</b> = not done and the TAT date has passed. "
            "<b>Pending · still within TAT</b> = not done and the TAT date hasn't come yet.</p>", unsafe_allow_html=True)
        summary = tk.tat_summary(df, cfg)
        st.dataframe(summary, width="stretch", hide_index=True, column_config={
            "On-time %": st.column_config.ProgressColumn("On-time %", min_value=0, max_value=100, format="%d%%"),
        })
        c1, c2 = st.columns([1, 1.3])
        with c1:
            box = card("tat_avg", "Average days taken vs TAT",
                       "Finished deliverables with an Edit Date. The black line is the TAT target.")
            labels = {r.deliverable: cfg.label(r.deliverable) for r in cfg.tat if f"{r.deliverable}__tat_days" in df}
            avg = pd.Series({labels[d]: df.loc[df[f"{d}__tat_outcome"].isin(["On time", "Late"]), f"{d}__tat_days"].mean()
                             for d in labels})
            target = pd.Series({labels[d]: df[f"{d}__tat_target"].median() for d in labels})
            show_chart(box, charts.days_vs_target(avg, target), "tat_avg",
                       pd.DataFrame({"Average days taken": avg.round(1), "TAT (days)": target}))
        pend = tk.tat_pending(df, cfg, today)
        with c2:
            box = card("tat_due", "When pending items are due",
                       "Pending deliverables by TAT date (start date + TAT days). Orange = overdue.")
        if pend.empty:
            show_chart(box, charts.due_columns(pd.Series(dtype=int), pd.Series(dtype=str)), "tat_due")
        else:
            today_ts = pd.Timestamp(today)
            horizon = today_ts + pd.Timedelta(days=14)

            def bucket(row) -> str:
                if pd.isna(row["TAT due date"]):
                    return "Not shot yet"
                if row["Days left"] < 0:
                    return "Overdue"
                if row["TAT due date"] > horizon:
                    return "Later"
                return row["TAT due date"].strftime("%d %b")

            pend["_bucket"] = pend.apply(bucket, axis=1)
            dated = sorted({b for b in pend["_bucket"] if b not in ("Overdue", "Later", "Not shot yet")},
                           key=lambda x: pd.Timestamp(f"{x} {today_ts.year}"))
            order = (["Overdue"] if (pend["_bucket"] == "Overdue").any() else []) + dated + \
                [b for b in ("Later", "Not shot yet") if (pend["_bucket"] == b).any()]
            counts = pend["_bucket"].value_counts().reindex(order)
            hover = pend.groupby("_bucket")["Deliverable"].apply(
                lambda s: " · ".join(f"{k} {v}" for k, v in s.value_counts().items()))
            show_chart(box, charts.due_columns(counts, hover), "tat_due",
                       pd.crosstab(pend["_bucket"], pend["Deliverable"]).reindex(order))
            pend = pend.drop(columns="_bucket")

        box = card("tat_list", "Pending deliverables with TAT due dates",
                   "Most urgent first. Row numbers match the Google Sheet.")
        if pend.empty:
            box.success("Nothing pending.")
        else:
            view = box.segmented_control("Show", ["All pending", "Overdue", "Due in next 7 days", "Not shot yet"],
                                         default="All pending", key=f"{tracker_name}_tat_view")
            shown_rows = pend
            if view == "Overdue":
                shown_rows = pend[pend["Days left"] < 0]
            elif view == "Due in next 7 days":
                shown_rows = pend[pend["Days left"].between(0, 7)]
            elif view == "Not shot yet":
                shown_rows = pend[pend["TAT due date"].isna()]
            box.dataframe(shown_rows, width="stretch", hide_index=True, height=460, column_config={
                "TAT due date": st.column_config.DateColumn(format="DD-MM-YYYY"),
                "Start date": st.column_config.DateColumn(format="DD-MM-YYYY"),
                "TAT (days)": st.column_config.NumberColumn(format="%d"),
                "Days left": st.column_config.NumberColumn(format="%d", help="Negative = days past the TAT date"),
            })
            box.download_button("Download list (CSV)", shown_rows.to_csv(index=False).encode("utf-8-sig"),
                                file_name=f"{cfg.name.lower().replace(' ', '_')}_tat_due_{today:%Y%m%d}.csv",
                                mime="text/csv", icon=":material/download:")

with tab_break:
    def breakdown(dim: str, order: list | None = None) -> pd.DataFrame:
        if focus:
            return tk.tat_view_table(df, focus, dim, order)
        return stage_table(df, dim, order)

    colors = charts.TAT_VIEW_COLORS if focus else charts.STAGE_COLORS
    prefix = f"{focus_label}: by" if focus else "Progress by"
    c1, c2 = st.columns(2)
    with c1:
        all_verticals = breakdown("vertical")
        shown = all_verticals.head(15)
        note = (f"The {len(shown)} largest of {len(all_verticals)} verticals. The table has all of them."
                if len(all_verticals) > len(shown) else "")
        box = card("vertical", f"{prefix} vertical", note)
        show_chart(box, charts.stacked_bars(shown, colors), "vertical", all_verticals)
    with c2:
        box = card("budget", f"{prefix} budget DPT (GT / MT)")
        t = breakdown("budget", order=["GT", "MT", "GT + MT", "Not set"])
        show_chart(box, charts.stacked_bars(t, colors), "budget", t)
        box = card("ideal", f"{prefix} ideal for")
        t = breakdown("ideal_for")
        show_chart(box, charts.stacked_bars(t, colors), "ideal", t)

    if focus:
        box = card("styles", f"{focus_label}: by style",
                   "One row per style code. Sorted by the most overdue SKUs.")
        v = tk.tat_view(df, focus)
        styles = pd.DataFrame({"style": df["style"], "vertical": df["vertical"], "project": df["project"],
                               "inward": df["inward_date"], "group": v})
        table = styles.groupby("style", dropna=False).agg(
            Vertical=("vertical", "first"), Project=("project", "first"),
            SKUs=("group", "size"), First_inward=("inward", "min"))
        for g in tk.TAT_VIEW_ORDER:
            table[g] = styles.assign(hit=styles["group"].eq(g)).groupby("style", dropna=False)["hit"].sum().astype(int)
        table = table.loc[:, [c for c in table.columns if c not in tk.TAT_VIEW_ORDER or table[c].sum() > 0]]
        sort_cols = [c for c in ["Overdue", "SKUs"] if c in table.columns]
        table = table.sort_values(sort_cols, ascending=False).reset_index().rename(
            columns={"style": "Style code", "First_inward": "First inward", "SKUs": "Total SKUs"})
        table = table[[c for c in table.columns if c != "First inward"] + ["First inward"]]
        box.dataframe(table, width="stretch", hide_index=True, height=min(38 + 35 * len(table), 460),
                      column_config={"First inward": st.column_config.DateColumn(format="DD-MM-YYYY")})
    else:
        box = card("styles", "Progress by style", "One row per style code. Sorted by the most SKUs still open.")
        styles = df.groupby("style", dropna=False).agg(
            Vertical=("vertical", "first"), Project=("project", "first"), SKUs=("sku", "size"),
            Shot=("shoot_status", lambda s: int(s.isin(tk.SHOT_STATUSES).sum())),
            Complete=("stage", lambda s: int((s == "Complete").sum())),
            Inward=("inward_date", "min"),
        )
        styles["Open"] = styles["SKUs"] - styles["Complete"]
        styles["% complete"] = (styles["Complete"] / styles["SKUs"] * 100).round(0)
        styles = styles.sort_values(["Open", "SKUs"], ascending=False).reset_index().rename(columns={"style": "Style code"})
        box.dataframe(
            styles, width="stretch", hide_index=True, height=min(38 + 35 * len(styles), 460),
            column_config={
                "% complete": st.column_config.ProgressColumn("% complete", min_value=0, max_value=100, format="%d%%"),
                "Inward": st.column_config.DateColumn("First inward", format="DD-MM-YYYY"),
            },
        )

with tab_team:
    shot_df = df[df["shoot_status"].isin(tk.SHOT_STATUSES)]

    def people_counts(frame: pd.DataFrame, column: str) -> pd.Series:
        names = frame[column].map(lambda v: tk.split_people(v) or ["Not recorded"]).explode()
        return names.value_counts()

    c1, c2 = st.columns(2)
    with c1:
        box = card("crew", f"SKUs shot by {cfg.crew_label.lower()}")
        s = people_counts(shot_df, "crew")
        show_chart(box, charts.single_bars(s), "crew", s.rename("SKUs shot").to_frame())
        box = card("stylist", "SKUs shot by stylist", "A SKU styled by two stylists counts for both.")
        s = people_counts(shot_df, "stylist")
        show_chart(box, charts.single_bars(s), "stylist", s.rename("SKUs shot").to_frame())
    with c2:
        box = card("model", "SKUs shot by model", "A SKU shot with two models counts for both.")
        s = people_counts(shot_df, "model")
        show_chart(box, charts.single_bars(s), "model", s.rename("SKUs shot").to_frame())
        rows = []
        for d in cfg.deliverables:
            done = df[df[f"{d}__status"] == "Done"]
            for editor in done[f"{d}__editor"].fillna("Not recorded"):
                rows.append((editor, cfg.label(d)))
        ed = pd.DataFrame(rows, columns=["Editor", "Deliverable"])
        box = card("editor", "Deliverables marked Done, by editor",
                   "Uses the Editor Name next to each deliverable. 'Not recorded' = Done with no editor name.")
        s = ed["Editor"].value_counts() if not ed.empty else pd.Series(dtype=int)
        tbl = pd.crosstab(ed["Editor"], ed["Deliverable"]) if not ed.empty else None
        show_chart(box, charts.single_bars(s, value_name="deliverables"), "editor", tbl)

with tab_pending:
    open_df = df[df["stage"] != "Complete"].copy()
    st.markdown("<p class='section-note'>Every SKU that isn't fully complete, oldest first. "
                "Row numbers match the Google Sheet.</p>", unsafe_allow_html=True)
    counts = {"Shoot": int((open_df["stage"] == "Pending shoot").sum())}
    for d in cfg.deliverables:
        counts[cfg.label(d)] = int(open_df["waiting_for"].str.contains(cfg.label(d), regex=False).sum())
    st.markdown("".join(f"<span class='chip'>Waiting for {html.escape(k)}: <b>{v:,}</b></span>" for k, v in counts.items()),
                unsafe_allow_html=True)
    show = st.segmented_control("Show", ["All open", "Pending shoot", "In editing"], default="All open",
                                key=f"{tracker_name}_pending_view")
    if show in ("Pending shoot", "In editing"):
        open_df = open_df[open_df["stage"] == show]
    open_df = open_df.sort_values(["days_since_inward", "sheet_row"], ascending=[False, True])
    pending_view = pd.DataFrame({
        "Row in sheet": open_df["sheet_row"],
        "SKU": open_df["sku"],
        "Style": open_df["style"],
        "Waiting for": open_df["waiting_for"],
        "Days since inward": open_df["days_since_inward"],
        "Inward date": open_df["inward_date"],
        "Stage": open_df["stage"],
        "Colour": open_df["colour"],
        "Vertical": open_df["vertical"],
        "Project": open_df["project"],
        "Budget": open_df["budget"],
        "Model": open_df["model"],
        "Stylist": open_df["stylist"],
    })
    st.dataframe(pending_view, width="stretch", hide_index=True, height=520,
                 column_config={"Waiting for": st.column_config.TextColumn(width="large"),
                                "Inward date": st.column_config.DateColumn(format="DD-MM-YYYY"),
                                "Days since inward": st.column_config.NumberColumn(format="%d")})
    st.download_button("Download pending list (CSV)", pending_view.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{cfg.name.lower().replace(' ', '_')}_pending_{today:%Y%m%d}.csv",
                       mime="text/csv", icon=":material/download:")

with tab_quality:
    st.markdown("<p class='section-note'>Rows in the sheet that need a fix. This list covers the whole "
                f"'{html.escape(cfg.sheet)}' sheet (filters don't apply). The dashboard already reads around "
                "these, but fixing them keeps the numbers right.</p>", unsafe_allow_html=True)
    if issues.empty:
        st.success("No data issues found.")
    else:
        summary = issues.groupby(["Severity", "Issue"]).size().rename("Rows").reset_index()
        st.dataframe(summary, width="stretch", hide_index=True)
        pick = st.selectbox("Show rows for", ["All issues"] + summary["Issue"].tolist(), key=f"{tracker_name}_issue")
        detail = issues if pick == "All issues" else issues[issues["Issue"] == pick]
        st.dataframe(detail, width="stretch", hide_index=True, height=420)
        st.download_button("Download issue list (CSV)", detail.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"{cfg.name.lower().replace(' ', '_')}_data_issues_{today:%Y%m%d}.csv",
                           mime="text/csv", icon=":material/download:")

with tab_data:
    export = pd.DataFrame({
        "Row in sheet": df["sheet_row"], "SKU": df["sku"], "Style code": df["style"], "Colour": df["colour"],
        "Size": df["size"], "Vertical": df["vertical"], "Super category": df["super_category"],
        "Ideal for": df["ideal_for"], "Project": df["project"], "Budget DPT": df["budget"],
        "Inward date": df["inward_date"], "Shoot status": df["shoot_status"], "Shoot date": df["shoot_date"],
        cfg.crew_label: df["crew"], "Stylist": df["stylist"], "Model": df["model"],
    })
    for d in cfg.deliverables:
        label = cfg.label(d)
        export[f"{label} status"] = df[f"{d}__status"]
        export[f"{label} edit date"] = df[f"{d}__edit_date"]
        export[f"{label} editor"] = df[f"{d}__editor"]
        export[f"{label} days from inward"] = df[f"{d}__days"]
        if f"{d}__tat_outcome" in df:
            export[f"{label} TAT (days)"] = df[f"{d}__tat_target"]
            export[f"{label} TAT result"] = df[f"{d}__tat_outcome"]
            export[f"{label} days over TAT"] = df[f"{d}__tat_days_over"]
            export[f"{label} actual (sheet)"] = df[f"{d}__tat_sheet_actual"]
    for r in cfg.tat:
        if r.deliverable not in cfg.deliverables and f"{r.deliverable}__tat_outcome" in df:  # e.g. Shoot
            label = cfg.label(r.deliverable)
            export[f"{label} TAT (days)"] = df[f"{r.deliverable}__tat_target"]
            export[f"{label} TAT result"] = df[f"{r.deliverable}__tat_outcome"]
            export[f"{label} days over TAT"] = df[f"{r.deliverable}__tat_days_over"]
            export[f"{label} actual (sheet)"] = df[f"{r.deliverable}__tat_sheet_actual"]
    if "tat_status_sheet" in df:
        export["TAT Status (sheet)"] = df["tat_status_sheet"]
    export["Stage"] = df["stage"]
    export["Waiting for"] = df["waiting_for"]
    export["Days since inward"] = df["days_since_inward"]
    export["Inward → shoot (days)"] = df["days_inward_to_shoot"]
    export["Shoot → last edit (days)"] = df["days_shoot_to_edit"]
    export["Inward → last edit (days)"] = df["days_inward_to_last_edit"]
    export["Number of assets"] = df["assets"]
    date_cols = {c: st.column_config.DateColumn(format="DD-MM-YYYY") for c in export.columns if "date" in c.lower()}
    st.markdown(f"<p class='section-note'>{len(export):,} rows after filters, cleaned (status spelling, dates, "
                "budget and names made consistent).</p>", unsafe_allow_html=True)
    st.dataframe(export, width="stretch", hide_index=True, height=560, column_config=date_cols)
    st.download_button("Download filtered data (CSV)", export.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{cfg.name.lower().replace(' ', '_')}_{today:%Y%m%d}.csv",
                       mime="text/csv", icon=":material/download:")
