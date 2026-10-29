"""
Load and clean the Brand Studio master tracker (Photo Shoot / Video Shoot sheets).

This file has no Streamlit code, so it can be tested on its own (see tests/).

What it handles in the raw sheet:
- Repeated "Edit Date" / "Editor Name" headers: each deliverable (Master Poster, PDF, ...)
  uses the two columns just before it as its edit date and editor.
- Dates typed as real dates or as text ("23/10/2026"), always read day-first.
- Status spelling: Done / done / DONE -> Done, Outsource -> Outsourced, blank -> Pending.
- Budget DPT: "GT, MT", "MT,GT", "MT   " -> "GT + MT" / "MT".
- Stylist / model teams: "Sriya/Mustafa" and "Mustafa/Sriya" -> "Mustafa / Sriya".
- Lists rows that need fixing in the sheet (see build_issues).
"""
from __future__ import annotations

import io
import math
import re
from dataclasses import dataclass
from datetime import date, datetime

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Tracker settings - edit here if the sheet layout changes
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TatRule:
    """Turnaround target for one deliverable."""
    deliverable: str        # status column header, e.g. "Master Poster"
    target_col: str         # sheet column with the TAT target in days, e.g. "TAT Poster"
    actual_col: str         # sheet column with the actual days (only used to cross-check)
    variance_col: str       # sheet column with actual - TAT (kept in the data export)
    default_days: float     # used when the TAT cell is blank
    start: str = "inward"   # TAT counts from "inward" date or "shoot" date


@dataclass(frozen=True)
class TrackerConfig:
    name: str                       # shown in the app
    sheet: str                      # sheet (tab) name in the workbook
    crew_col: str                   # "Photographer" / "VideoGrapher" column header
    crew_label: str                 # label used in the app
    deliverables: tuple             # status column headers, in pipeline order
    labels: tuple = ()              # optional (header, display label) pairs
    # Optional: which deliverables each budget needs. None = every SKU needs all of them.
    # Example: (("GT", ("Master Poster", "PDF")), ("MT", ("MT catalog Image",)))
    required_by_budget: tuple | None = None
    tat: tuple = ()                 # TatRule per deliverable (empty = no TAT tab)
    tat_status_col: str = "TAT Status"

    def label(self, deliverable: str) -> str:
        return dict(self.labels).get(deliverable, deliverable)


TRACKERS = {
    "Photo Shoot": TrackerConfig(
        name="Photo Shoot",
        sheet="Photo Shoot",
        crew_col="Photographer",
        crew_label="Photographer",
        deliverables=("Master Poster", "PDF", "MT catalog Image"),
        labels=(("MT catalog Image", "MT Catalog Image"),),
        # TAT targets come from the TAT columns; blank cells use default_days.
        # Days are calendar days to the deliverable's Edit Date, counted from the inward date
        # (start="inward") or the shoot date (start="shoot"). MT Catalog Image counts from the shoot date.
        tat=(
            TatRule("Shoot", "Shoot TAT", "Actual Shoot", "Variance Shoot", 2, start="inward"),
            TatRule("Master Poster", "TAT Poster", "Actual Poster", "Variance Poster", 7, start="inward"),
            TatRule("PDF", "TAT PDF", "Actual PDF", "Variance PDF", 15, start="inward"),
            TatRule("MT catalog Image", "TAT MT", "Actual MT", "Variance MT", 3, start="shoot"),
        ),
    ),
    "Video Shoot": TrackerConfig(
        name="Video Shoot",
        sheet="Video Shoot",
        crew_col="VideoGrapher",
        crew_label="Videographer",
        deliverables=("GT Videos", "MT Videos"),
    ),
}

# Column headers the loader looks for (case and extra spaces don't matter).
FIELD_HEADERS = {
    "location": ["location"],
    "inward_date": ["inward date", "inward"],
    "sku": ["sku"],
    "style": ["style code", "style"],
    "brand": ["brand"],
    "colour": ["colour", "color"],
    "size": ["size"],
    "vertical": ["vertical name", "vertical"],
    "pairing": ["pairing style"],
    "super_category": ["super category"],
    "ideal_for": ["ideal for"],
    "project": ["project (shoot type)", "project", "shoot type"],
    "budget_raw": ["budget dpt", "budget"],
    "shoot_raw": ["shoot date"],
    "stylist_raw": ["stylist"],
    "model_raw": ["model name", "model"],
    "studio_asst": ["studio asst", "studio assistant"],
    "sequence": ["shoot sequence"],
    "assets_raw": ["number of assets", "assets"],
    "duplicate_raw": ["duplicate"],
    "content_status": ["content status"],
    "remarks": ["remarks"],
    "catalog_comment": ["catalog comment"],
}

STATUS_ORDER = ["Done", "Outsourced", "Pending", "Other", "Not needed"]
TAT_ORDER = ["On time", "Late", "Within TAT", "Overdue", "Done, no edit date", "Not started", "Not needed"]
STAGE_ORDER = ["Complete", "In editing", "Pending shoot"]
SHOT_STATUSES = ("Done", "Outsourced")

_DONE_WORDS = {"done", "completed", "complete", "yes", "delivered", "ok", "y"}
_DATE_FORMATS = (
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%y", "%d-%m-%y",
    "%d-%b-%Y", "%d %b %Y", "%d-%b-%y", "%d %B %Y", "%d/%m/%Y %H:%M:%S",
)

# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------


def _norm_header(h) -> str:
    return "" if _blank(h) else re.sub(r"\s+", " ", str(h)).strip().lower()


def _blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and math.isnan(v):
        return True
    if v is pd.NaT:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    return False


def clean_text(v) -> str | None:
    """Strip and collapse spaces. Whole numbers stored as floats lose their '.0'."""
    if _blank(v):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (datetime, date)):
        return pd.Timestamp(v).strftime("%d-%m-%Y")
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def parse_date(v) -> tuple[pd.Timestamp, str]:
    """Return (date, kind). kind: 'date', 'text_date', 'blank' or 'text' (not a date)."""
    if _blank(v):
        return pd.NaT, "blank"
    if isinstance(v, (pd.Timestamp, datetime, date)):
        ts = pd.Timestamp(v).normalize()
        return (ts, "date") if 2015 <= ts.year <= 2100 else (pd.NaT, "text")
    if isinstance(v, (int, float, np.integer, np.floating)):
        if 20000 <= float(v) <= 80000:  # Excel serial number
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(v)), "date"
        return pd.NaT, "text"
    s = re.sub(r"\s+", " ", str(v)).strip()
    for fmt in _DATE_FORMATS:
        try:
            ts = pd.Timestamp(datetime.strptime(s, fmt)).normalize()
        except ValueError:
            continue
        if 2015 <= ts.year <= 2100:
            return ts, "text_date"
    return pd.NaT, "text"


def norm_status(v) -> str:
    s = clean_text(v)
    if s is None:
        return "Pending"
    low = s.lower()
    if low in _DONE_WORDS or low.startswith("done"):
        return "Done"
    if "outsourc" in low:
        return "Outsourced"
    if low in {"pending", "-", "wip", "in progress"}:
        return "Pending"
    if low in {"na", "n/a", "not needed", "not required"}:
        return "Not needed"
    return "Other"


def norm_budget(v) -> str:
    tokens = set(re.findall(r"[A-Z]+", str(v).upper())) & {"GT", "MT"} if not _blank(v) else set()
    if tokens == {"GT", "MT"}:
        return "GT + MT"
    if tokens:
        return tokens.pop()
    return "Not set"


def norm_people(v) -> str | None:
    """'sriya/Mustafa ' -> 'Mustafa / Sriya'. Numbers are treated as not a name."""
    s = clean_text(v)
    if s is None or re.fullmatch(r"[\d.\s]+", s):
        return None
    parts = [p.strip().title() for p in re.split(r"[/,&+]| and ", s) if p.strip()]
    return " / ".join(sorted(dict.fromkeys(parts))) if parts else None


def split_people(v) -> list[str]:
    return [p for p in v.split(" / ") if p] if isinstance(v, str) else []


def extract_file_id(url_or_id: str) -> str:
    """Accept a full Google Sheets / Drive link or a bare file ID."""
    s = (url_or_id or "").strip()
    m = re.search(r"/d/([A-Za-z0-9_-]{20,})", s) or re.search(r"[?&]id=([A-Za-z0-9_-]{20,})", s)
    return m.group(1) if m else s


def week_start(ts: pd.Series) -> pd.Series:
    return ts - pd.to_timedelta(ts.dt.weekday, unit="D")


# ---------------------------------------------------------------------------
# Download from Google Drive
# ---------------------------------------------------------------------------
_XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_GSHEET_MIME = "application/vnd.google-apps.spreadsheet"


def download_with_service_account(file_id: str, service_account_info: dict) -> tuple[bytes, dict]:
    """Download the tracker as .xlsx using a Google service account (keeps the sheet private).

    Works for a native Google Sheet and for an .xlsx file stored in Drive.
    The sheet must be shared (Viewer) with the service account's client_email.
    """
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2 import service_account

    creds = service_account.Credentials.from_service_account_info(
        dict(service_account_info), scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    session = AuthorizedSession(creds)
    base = f"https://www.googleapis.com/drive/v3/files/{file_id}"
    meta_resp = session.get(base, params={"fields": "name,mimeType,modifiedTime", "supportsAllDrives": "true"}, timeout=60)
    if meta_resp.status_code in (403, 404):
        raise RuntimeError(
            f"Google Drive returned {meta_resp.status_code} for the tracker. Share the sheet with "
            f"{service_account_info.get('client_email', 'the service account email')} as Viewer, "
            "and check that the Google Drive API is enabled in the Google Cloud project."
        )
    meta_resp.raise_for_status()
    meta = meta_resp.json()
    if meta.get("mimeType") == _GSHEET_MIME:
        resp = session.get(f"{base}/export", params={"mimeType": _XLSX_MIME}, timeout=120)
    else:
        resp = session.get(base, params={"alt": "media", "supportsAllDrives": "true"}, timeout=120)
    resp.raise_for_status()
    return resp.content, meta


def download_public(file_id: str) -> tuple[bytes, dict]:
    """Download a tracker shared as 'Anyone with the link - Viewer'. Anyone with the link can see the data."""
    import requests

    urls = [
        f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx",
        f"https://drive.google.com/uc?export=download&id={file_id}",
    ]
    for url in urls:
        resp = requests.get(url, timeout=120)
        if resp.ok and resp.content[:2] == b"PK":  # .xlsx files are zip archives
            return resp.content, {"name": "Shared Google Sheet"}
    raise RuntimeError(
        "Could not download the sheet from the public link. Set sharing to 'Anyone with the link - Viewer', "
        "or use the service account option (recommended)."
    )


# ---------------------------------------------------------------------------
# Parse one tracker sheet
# ---------------------------------------------------------------------------


def list_sheets(xlsx_bytes: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(xlsx_bytes)).sheet_names


def parse_tracker(xlsx_bytes: bytes, cfg: TrackerConfig, today: date | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (clean rows, issues). One clean row per SKU row in the sheet."""
    today = pd.Timestamp(today or date.today()).normalize()
    raw = pd.read_excel(io.BytesIO(xlsx_bytes), sheet_name=cfg.sheet, header=None, dtype=object)
    if raw.empty:
        raise ValueError(f"Sheet '{cfg.sheet}' is empty.")

    headers = [_norm_header(h) for h in raw.iloc[0].tolist()]
    body = raw.iloc[1:]

    def find(names: list[str]) -> int | None:
        for n in names:
            if n in headers:
                return headers.index(n)
        return None

    pos = {f: find(names) for f, names in FIELD_HEADERS.items()}
    pos["crew_raw"] = find([cfg.crew_col.lower()])
    if pos["sku"] is None:
        raise ValueError(f"Sheet '{cfg.sheet}' has no 'SKU' column in row 1.")

    deliv_pos = {}
    for d in cfg.deliverables:
        i = find([d.lower()])
        if i is None:
            deliv_pos[d] = (None, None, None)
            continue
        date_i = i - 2 if i >= 2 and headers[i - 2].startswith("edit date") else None
        editor_i = i - 1 if i >= 1 and headers[i - 1].startswith("editor") else None
        deliv_pos[d] = (i, date_i, editor_i)

    def col(i):
        return body.iloc[:, i] if i is not None else pd.Series([None] * len(body), index=body.index, dtype=object)

    sku = col(pos["sku"]).map(clean_text)
    keep = sku.notna()
    body = body[keep]
    sku = sku[keep]

    df = pd.DataFrame(index=body.index)
    df["sheet_row"] = body.index + 1  # row number as seen in Google Sheets / Excel
    df["sku"] = sku
    for f in ["location", "style", "brand", "colour", "size", "vertical", "pairing", "super_category",
              "ideal_for", "project", "studio_asst", "sequence", "content_status", "remarks", "catalog_comment"]:
        df[f] = col(pos[f]).map(clean_text)
    for f in ["vertical", "super_category", "ideal_for", "project", "size"]:
        df[f] = df[f].fillna("Not set")

    df["budget"] = col(pos["budget_raw"]).map(norm_budget)
    df["budget_raw"] = col(pos["budget_raw"]).map(clean_text)

    inward = col(pos["inward_date"]).map(parse_date)
    df["inward_date"] = pd.to_datetime(inward.map(lambda t: t[0]))
    inward_kind = inward.map(lambda t: t[1])

    crew = col(pos["crew_raw"]).map(norm_people)
    df["crew"] = crew
    shoot_cells = col(pos["shoot_raw"])
    shoot = shoot_cells.map(parse_date)
    df["shoot_date"] = pd.to_datetime(shoot.map(lambda t: t[0]))
    shoot_kind = shoot.map(lambda t: t[1])
    shoot_text = shoot_cells.map(clean_text)
    outsourced_shoot = shoot_text.fillna("").str.lower().str.contains("outsourc") | crew.fillna("").str.lower().str.contains("outsourc")
    df["shoot_status"] = np.select(
        [df["shoot_date"].notna(), outsourced_shoot, shoot_kind.eq("blank")],
        ["Done", "Outsourced", "Pending"],
        default="Other",
    )
    df["shoot_raw"] = shoot_text

    df["stylist"] = col(pos["stylist_raw"]).map(norm_people)
    df["model"] = col(pos["model_raw"]).map(norm_people)
    stylist_text = col(pos["stylist_raw"]).map(clean_text)
    model_text = col(pos["model_raw"]).map(clean_text)

    assets = pd.to_numeric(col(pos["assets_raw"]), errors="coerce")
    df["assets"] = assets
    dup_col = pd.to_numeric(col(pos["duplicate_raw"]), errors="coerce")

    issues: list[dict] = []

    def add_issue(row_idx, field, value, issue, severity="High"):
        issues.append({
            "Row in sheet": int(df.at[row_idx, "sheet_row"]),
            "SKU": df.at[row_idx, "sku"],
            "Column": field,
            "Value": "" if _blank(value) else str(value),
            "Issue": issue,
            "Severity": severity,
        })

    # Deliverables
    for d in cfg.deliverables:
        i, date_i, editor_i = deliv_pos[d]
        status = col(i).map(norm_status) if i is not None else pd.Series("Pending", index=body.index)
        status_text = col(i).map(clean_text)
        edit_cells = col(date_i)
        edit = edit_cells.map(parse_date)
        edit_date = pd.to_datetime(edit.map(lambda t: t[0]))
        edit_kind = edit.map(lambda t: t[1])
        editor_cells = col(editor_i)
        editor_as_date = editor_cells.map(parse_date)
        editor_is_date = editor_as_date.map(lambda t: t[1] in ("date", "text_date"))
        editor = editor_cells.map(norm_people).where(~editor_is_date, None)
        # A date typed in the Editor Name column is used as the edit date when that is blank.
        edit_date = edit_date.fillna(pd.to_datetime(editor_as_date.map(lambda t: t[0])).where(editor_is_date))
        df[f"{d}__status"] = status
        df[f"{d}__status_text"] = status_text
        df[f"{d}__edit_date"] = edit_date
        df[f"{d}__editor"] = editor

        label = cfg.label(d)
        for idx in df.index:
            if status[idx] == "Other":
                add_issue(idx, label, status_text[idx], "Status not recognised (use Done / Outsource / blank)", "Low")
            if editor_is_date[idx]:
                add_issue(idx, f"{label} - Editor Name", clean_text(editor_cells[idx]), "Date typed in the Editor Name column")
            if edit_kind[idx] == "text" and "outsourc" not in (clean_text(edit_cells[idx]) or "").lower():
                add_issue(idx, f"{label} - Edit Date", clean_text(edit_cells[idx]), "Not a valid date")
            if edit_kind[idx] == "text_date":
                add_issue(idx, f"{label} - Edit Date", clean_text(edit_cells[idx]), "Date typed as text (works, but format the column as Date)", "Low")
            if pd.notna(edit_date[idx]) and pd.notna(df.at[idx, "shoot_date"]) and edit_date[idx] < df.at[idx, "shoot_date"]:
                add_issue(idx, f"{label} - Edit Date", edit_date[idx].strftime("%d-%m-%Y"), "Edit date is before the shoot date")
            if status[idx] == "Done" and df.at[idx, "shoot_status"] == "Pending":
                add_issue(idx, label, status_text[idx], "Marked Done but the shoot date is blank")

    # Row-level checks
    dup_mask = df["sku"].duplicated(keep=False)
    rows_by_sku = df[dup_mask].groupby("sku")["sheet_row"].apply(lambda s: ", ".join(map(str, s)))
    for idx in df.index:
        if dup_mask[idx]:
            add_issue(idx, "SKU", f"{df.at[idx, 'sku']} (rows {rows_by_sku[df.at[idx, 'sku']]})", "SKU appears more than once", "Low")
        if inward_kind[idx] == "blank":
            add_issue(idx, "inward date", None, "Inward date is blank")
        elif inward_kind[idx] == "text":
            add_issue(idx, "inward date", clean_text(col(pos["inward_date"])[idx]), "Not a valid date")
        elif inward_kind[idx] == "text_date":
            add_issue(idx, "inward date", clean_text(col(pos["inward_date"])[idx]), "Date typed as text (works, but format the column as Date)", "Low")
        if shoot_kind[idx] == "text_date":
            add_issue(idx, "Shoot date", shoot_text[idx], "Date typed as text (works, but format the column as Date)", "Low")
        elif df.at[idx, "shoot_status"] == "Other":
            add_issue(idx, "Shoot date", shoot_text[idx], "Not a valid date")
        if pd.notna(df.at[idx, "shoot_date"]) and pd.notna(df.at[idx, "inward_date"]) and df.at[idx, "shoot_date"] < df.at[idx, "inward_date"]:
            add_issue(idx, "Shoot date", df.at[idx, "shoot_date"].strftime("%d-%m-%Y"), "Shoot date is before the inward date")
        if df.at[idx, "vertical"] == "Not set":
            add_issue(idx, "Vertical Name", None, "Vertical Name is blank", "Low")
        if pd.notna(stylist_text[idx]) and pd.isna(df.at[idx, "stylist"]):
            add_issue(idx, "Stylist", stylist_text[idx], "Number in the Stylist column")
        if pd.notna(model_text[idx]) and pd.isna(df.at[idx, "model"]):
            add_issue(idx, "Model Name", model_text[idx], "Number in the Model Name column")
        a, dcol = assets[idx], dup_col[idx]
        if pd.notna(a) and pd.notna(dcol) and a == dcol and a > 2:
            add_issue(idx, "Number of Assets", int(a), "Same number as the Duplicate column - looks like a serial number", "Low")

    # Days from inward date to each deliverable's edit date (top KPI tiles)
    for d in cfg.deliverables:
        days = (df[f"{d}__edit_date"] - df["inward_date"]).dt.days
        for idx in df.index[days < 0]:
            add_issue(idx, f"{cfg.label(d)} - Edit Date", df.at[idx, f"{d}__edit_date"].strftime("%d-%m-%Y"),
                      "Edit date is before the inward date")
        df[f"{d}__days"] = days.where(days >= 0)

    # The shoot step can have its own TAT (inward date -> shoot date), handled like a deliverable
    df["Shoot__status"] = df["shoot_status"]
    df["Shoot__status_text"] = df["shoot_raw"]
    df["Shoot__edit_date"] = df["shoot_date"]
    df["Shoot__editor"] = df["crew"]

    # TAT (turnaround time) per deliverable
    for rule in cfg.tat:
        d = rule.deliverable
        if f"{d}__status" not in df:
            continue
        target_cells = col(find([rule.target_col.lower()]))
        target = pd.to_numeric(target_cells, errors="coerce")
        for idx in df.index[target.isna() & target_cells.map(lambda v: not _blank(v))]:
            add_issue(idx, rule.target_col, clean_text(target_cells[idx]), "TAT is not a number")
        target = target.fillna(rule.default_days)
        start = df["shoot_date"] if rule.start == "shoot" else df["inward_date"]
        tat_days = (df[f"{d}__edit_date"] - start).dt.days
        tat_days = tat_days.where(tat_days >= 0)
        open_days = (today - start).dt.days
        status = df[f"{d}__status"]
        has_days = tat_days.notna()
        outcome = np.select(
            [has_days & (tat_days <= target), has_days, status.eq("Not needed").to_numpy(),
             status.isin(["Done", "Outsourced"]).to_numpy(), (open_days > target).to_numpy(), open_days.notna().to_numpy()],
            ["On time", "Late", "Not needed", "Done, no edit date", "Overdue", "Within TAT"],
            default="Not started",
        )
        df[f"{d}__tat_target"] = target
        df[f"{d}__tat_start"] = start
        df[f"{d}__tat_due"] = start + pd.to_timedelta(target, unit="D")   # date it must be done by
        df[f"{d}__tat_days"] = tat_days
        df[f"{d}__tat_open_days"] = open_days.where(~has_days)
        df[f"{d}__tat_outcome"] = outcome
        df[f"{d}__tat_days_over"] = np.where(outcome == "Late", tat_days - target,
                                             np.where(outcome == "Overdue", open_days - target, np.nan))
        sheet_actual = pd.to_numeric(col(find([rule.actual_col.lower()])), errors="coerce")
        df[f"{d}__tat_sheet_actual"] = sheet_actual
        df[f"{d}__tat_sheet_variance"] = pd.to_numeric(col(find([rule.variance_col.lower()])), errors="coerce")
        mismatch = sheet_actual.notna() & tat_days.notna() & (sheet_actual != tat_days)
        for idx in df.index[mismatch]:
            add_issue(idx, rule.actual_col, f"sheet {sheet_actual[idx]:g} · from dates {tat_days[idx]:g}",
                      "Actual days in the sheet don't match the dates", "Low")
    if cfg.tat:
        df["tat_status_sheet"] = col(find([cfg.tat_status_col.lower()])).map(clean_text)

    issues_df = pd.DataFrame(issues, columns=["Row in sheet", "SKU", "Column", "Value", "Issue", "Severity"])
    if not issues_df.empty:
        issues_df = issues_df.sort_values(["Severity", "Row in sheet"]).reset_index(drop=True)

    df["inward_week"] = week_start(df["inward_date"])
    df["shoot_week"] = week_start(df["shoot_date"])
    df["days_inward_to_shoot"] = (df["shoot_date"] - df["inward_date"]).dt.days
    df.loc[df["days_inward_to_shoot"] < 0, "days_inward_to_shoot"] = np.nan
    edit_cols = [f"{d}__edit_date" for d in cfg.deliverables]
    df["last_edit_date"] = df[edit_cols].max(axis=1)
    df["days_shoot_to_edit"] = (df["last_edit_date"] - df["shoot_date"]).dt.days
    df.loc[df["days_shoot_to_edit"] < 0, "days_shoot_to_edit"] = np.nan
    df["days_inward_to_last_edit"] = (df["last_edit_date"] - df["inward_date"]).dt.days
    df.loc[df["days_inward_to_last_edit"] < 0, "days_inward_to_last_edit"] = np.nan
    df["days_since_inward"] = (today - df["inward_date"]).dt.days
    return df.reset_index(drop=True), issues_df


# ---------------------------------------------------------------------------
# Stage (depends on the "count outsourced as complete" switch)
# ---------------------------------------------------------------------------


def required_deliverables(cfg: TrackerConfig, budget: str) -> tuple:
    if not cfg.required_by_budget:
        return cfg.deliverables
    rules = dict(cfg.required_by_budget)
    needed = set()
    for token in budget.split(" + "):
        needed.update(rules.get(token, ()))
    return tuple(d for d in cfg.deliverables if d in needed) or cfg.deliverables


def add_stage(df: pd.DataFrame, cfg: TrackerConfig, outsourced_counts_as_done: bool = True) -> pd.DataFrame:
    """Add 'stage' (Complete / In editing / Pending shoot) and 'waiting_for'."""
    out = df.copy()
    ok = {"Done", "Not needed"} | ({"Outsourced"} if outsourced_counts_as_done else set())

    # Mark deliverables this SKU doesn't need (only when required_by_budget is set).
    if cfg.required_by_budget:
        for d in cfg.deliverables:
            not_needed = ~out["budget"].map(lambda b: d in required_deliverables(cfg, b))
            out.loc[not_needed & out[f"{d}__status"].eq("Pending"), f"{d}__status"] = "Not needed"
            if f"{d}__tat_outcome" in out:
                open_rows = not_needed & out[f"{d}__tat_outcome"].isin(["Within TAT", "Overdue", "Not started"])
                out.loc[open_rows, f"{d}__tat_outcome"] = "Not needed"
                out.loc[open_rows, f"{d}__tat_days_over"] = np.nan

    done_all = pd.Series(True, index=out.index)
    for d in cfg.deliverables:
        done_all &= out[f"{d}__status"].isin(ok)
    shot = out["shoot_status"].isin(SHOT_STATUSES)
    out["stage"] = np.select([done_all, ~shot], ["Complete", "Pending shoot"], default="In editing")

    def waiting(row) -> str:
        items = []
        if row["shoot_status"] not in SHOT_STATUSES and row["stage"] != "Complete":
            items.append("Shoot")
        for d in cfg.deliverables:
            st = row[f"{d}__status"]
            if st not in ok:
                text = row[f"{d}__status_text"]
                items.append(f"{cfg.label(d)} ({text})" if st == "Other" and isinstance(text, str) else cfg.label(d))
        return ", ".join(items)

    out["waiting_for"] = out.apply(waiting, axis=1)
    out.loc[out["stage"] == "Complete", "waiting_for"] = ""
    return out


def status_long(df: pd.DataFrame, cfg: TrackerConfig) -> pd.DataFrame:
    """Counts per pipeline step and status: Shoot, then each deliverable."""
    rows = [("Shoot", s, n) for s, n in df["shoot_status"].value_counts().items()]
    for d in cfg.deliverables:
        rows += [(cfg.label(d), s, n) for s, n in df[f"{d}__status"].value_counts().items()]
    out = pd.DataFrame(rows, columns=["Step", "Status", "SKUs"])
    out["Step"] = pd.Categorical(out["Step"], ["Shoot"] + [cfg.label(d) for d in cfg.deliverables], ordered=True)
    out["Status"] = pd.Categorical(out["Status"], STATUS_ORDER, ordered=True)
    return out.sort_values(["Step", "Status"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# TAT tables
# ---------------------------------------------------------------------------


def tat_outcomes(df: pd.DataFrame, cfg: TrackerConfig) -> pd.DataFrame:
    """Rows = deliverables, columns = TAT outcome, values = SKUs."""
    rows = {}
    for rule in cfg.tat:
        col = f"{rule.deliverable}__tat_outcome"
        if col in df:
            rows[cfg.label(rule.deliverable)] = df[col].value_counts()
    t = pd.DataFrame(rows).T.fillna(0).astype(int)
    return t.reindex(columns=[c for c in TAT_ORDER if c in t.columns])


def tat_summary(df: pd.DataFrame, cfg: TrackerConfig) -> pd.DataFrame:
    """One row per deliverable, using the same groups as the KPI cards."""
    out = []
    for rule in cfg.tat:
        d = rule.deliverable
        if f"{d}__tat_outcome" not in df:
            continue
        oc, target, days = df[f"{d}__tat_outcome"], df[f"{d}__tat_target"], df[f"{d}__tat_days"]
        finished = oc.isin(["On time", "Late"])
        lo, hi = target.min(), target.max()
        overdue, within, not_started = (oc == "Overdue").sum(), (oc == "Within TAT").sum(), (oc == "Not started").sum()
        row = {
            "Deliverable": cfg.label(d),
            "TAT (days)": "–" if pd.isna(lo) else (f"{lo:g}" if lo == hi else f"{lo:g}–{hi:g}"),
            "Counted from": "Shoot date" if rule.start == "shoot" else "Inward date",
            "Completed within TAT": int((oc == "On time").sum()),
            "Pending": int(overdue + within + not_started),
            "Overdue": int(overdue),
            "Pending · still within TAT": int(within),
        }
        if rule.start == "shoot":
            row["Pending · not shot yet"] = int(not_started)
        row["On-time %"] = round((oc == "On time").sum() / finished.sum() * 100) if finished.sum() else None
        row["Avg days taken"] = round(days[finished].mean(), 1) if finished.any() else None
        out.append(row)
    t = pd.DataFrame(out)
    if "Pending · not shot yet" in t:
        t["Pending · not shot yet"] = t["Pending · not shot yet"].fillna(0).astype(int)
        cols = list(t.columns)
        cols.insert(cols.index("Pending · still within TAT") + 1, cols.pop(cols.index("Pending · not shot yet")))
        t = t[cols]
    return t


def tat_exceptions(df: pd.DataFrame, cfg: TrackerConfig) -> pd.DataFrame:
    """One row per late or overdue deliverable, worst first."""
    parts = []
    for rule in cfg.tat:
        d = rule.deliverable
        if f"{d}__tat_outcome" not in df:
            continue
        m = df[f"{d}__tat_outcome"].isin(["Late", "Overdue"])
        if not m.any():
            continue
        sub = df[m]
        parts.append(pd.DataFrame({
            "Row in sheet": sub["sheet_row"],
            "SKU": sub["sku"],
            "Style": sub["style"],
            "Deliverable": cfg.label(d),
            "Result": sub[f"{d}__tat_outcome"].map({"Late": "Late (done)", "Overdue": "Overdue (still open)"}),
            "TAT (days)": sub[f"{d}__tat_target"],
            "Days taken / open": sub[f"{d}__tat_days"].fillna(sub[f"{d}__tat_open_days"]),
            "Days over TAT": sub[f"{d}__tat_days_over"],
            "Inward date": sub["inward_date"],
            "Shoot date": sub["shoot_date"],
            "Edit date": sub[f"{d}__edit_date"],
            "Editor": sub[f"{d}__editor"],
            "Project": sub["project"],
        }))
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    out["_open"] = out["Result"].str.startswith("Overdue")
    return out.sort_values(["_open", "Days over TAT"], ascending=[False, False]).drop(columns="_open").reset_index(drop=True)


PENDING_TAT = ("Within TAT", "Overdue", "Not started")


def tat_pending(df: pd.DataFrame, cfg: TrackerConfig, today: date | None = None) -> pd.DataFrame:
    """Every pending deliverable with its TAT due date, most urgent first."""
    today = pd.Timestamp(today or date.today()).normalize()
    parts = []
    for rule in cfg.tat:
        d = rule.deliverable
        if f"{d}__tat_outcome" not in df:
            continue
        sub = df[df[f"{d}__tat_outcome"].isin(PENDING_TAT)]
        if sub.empty:
            continue
        due = sub[f"{d}__tat_due"]
        left = (due - today).dt.days

        def describe(n, not_started):
            if not_started:
                return "Not shot yet" if rule.start == "shoot" else "No start date"
            if n < 0:
                return f"Overdue by {-int(n)} day{'s' if n != -1 else ''}"
            if n == 0:
                return "Due today"
            return f"Due in {int(n)} day{'s' if n != 1 else ''}"

        parts.append(pd.DataFrame({
            "Row in sheet": sub["sheet_row"],
            "SKU": sub["sku"],
            "Style": sub["style"],
            "Deliverable": cfg.label(d),
            "TAT due date": due,
            "Status": [describe(n, pd.isna(n)) for n in left],
            "Days left": left,
            "TAT (days)": sub[f"{d}__tat_target"],
            "Counted from": "Shoot date" if rule.start == "shoot" else "Inward date",
            "Start date": sub[f"{d}__tat_start"],
            "Vertical": sub["vertical"],
            "Project": sub["project"],
        }))
    if not parts:
        return pd.DataFrame()
    out = pd.concat(parts, ignore_index=True)
    return out.sort_values(["TAT due date", "Row in sheet"], na_position="last").reset_index(drop=True)


# The three groups used on the cards and in every chart (plus "not shot yet" for MT)
TAT_VIEW = {
    "On time": "Completed within TAT",
    "Overdue": "Overdue",
    "Within TAT": "Pending · still within TAT",
    "Not started": "Pending · not shot yet",
}
TAT_VIEW_ORDER = list(TAT_VIEW.values())


def tat_view(df: pd.DataFrame, d: str) -> pd.Series:
    """Chart group per SKU for deliverable d. Completed late / without an edit date -> NaN (not shown)."""
    return df[f"{d}__tat_outcome"].map(TAT_VIEW)


def tat_view_table(df: pd.DataFrame, d: str, dim: str, order: list | None = None) -> pd.DataFrame:
    """Rows = values of `dim`, columns = TAT groups, values = SKUs."""
    v = tat_view(df, d)
    m = v.notna()
    if not m.any():
        return pd.DataFrame()
    t = pd.crosstab(df.loc[m, dim], v[m]).reindex(columns=[c for c in TAT_VIEW_ORDER if c in set(v[m])], fill_value=0)
    if order:
        t = t.reindex([o for o in order if o in t.index] + [i for i in t.index if i not in order])
    elif not pd.api.types.is_datetime64_any_dtype(t.index):
        t = t.loc[t.sum(axis=1).sort_values(ascending=False).index]
    t.index.name = None
    t.columns.name = None
    return t


def tat_view_by_deliverable(df: pd.DataFrame, cfg: TrackerConfig) -> pd.DataFrame:
    rows = {}
    for rule in cfg.tat:
        d = rule.deliverable
        if f"{d}__tat_outcome" in df:
            rows[cfg.label(d)] = tat_view(df, d).value_counts()
    if not rows:
        return pd.DataFrame()
    t = pd.DataFrame(rows).T.fillna(0).astype(int)
    return t.reindex(columns=[c for c in TAT_VIEW_ORDER if c in t.columns])
