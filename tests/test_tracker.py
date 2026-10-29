"""Tests for the tracker loader and a smoke test of the whole app.

They build a small fake tracker, so no real data is needed (safe for GitHub).
Run:  pytest -q
"""
from __future__ import annotations

import io
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook

import tracker as tk

ROOT = Path(__file__).resolve().parents[1]

PHOTO_HEADERS = [
    "Location", "inward date", "SKU", "Style Code", "Brand", "Colour", "Size", "Vertical Name", "Pairing Style",
    "Super Category", "Ideal For", "Project (Shoot Type)", "Budget DPT", "Photographer", "Shoot date", "Stylist",
    "Model Name", "Studio Asst", "Edit Date", "Editor Name", "Master Poster", "Edit Date", "Editor Name", "PDF",
    "Edit Date", "Editor Name", "MT catalog Image", "Number of Assets", "Touch Point", "Duplicate",
    "Content Status", "Remarks", "Catalog Comment",
]


def _row(sku, inward, shoot, mp, pdf, mt, **kw):
    r = dict.fromkeys(range(len(PHOTO_HEADERS)))
    r.update({0: "BLR", 1: inward, 2: sku, 3: kw.get("style", "P1/01"), 4: "Technosport", 5: "BLACK", 6: "L",
              7: kw.get("vertical", "T-SHIRT"), 9: "Mens Apparels", 10: "Mens", 11: "Regular Catalog",
              12: kw.get("budget", "GT, MT"), 13: kw.get("crew", "Abishek"), 14: shoot,
              15: kw.get("stylist", "Sriya/Mustafa"), 16: kw.get("model", "Julio"), 17: "Ashok",
              18: kw.get("mp_date"), 19: kw.get("mp_editor"), 20: mp,
              21: kw.get("pdf_date"), 22: kw.get("pdf_editor"), 23: pdf,
              24: kw.get("mt_date"), 25: kw.get("mt_editor"), 26: mt, 27: kw.get("assets"), 29: kw.get("dup")})
    return [r[i] for i in range(len(PHOTO_HEADERS))]


def make_workbook() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Photo Shoot"
    ws.append(PHOTO_HEADERS)
    d = datetime
    rows = [
        # complete, real dates, lower-case status
        _row("SKU1", d(2026, 10, 6), d(2026, 10, 8), "done", "Done", "Done", mp_editor="suvin",
             mt_date=d(2026, 10, 13), mt_editor="Suvin", assets=2),
        # shoot date typed as text, outsourced deliverables
        _row("SKU2", d(2026, 10, 6), "09/10/2026", "Outsource", "Outsource", "Done", stylist="Mustafa/Sriya",
             mp_date="Outsource", mp_editor="Outsource"),
        # not shot yet
        _row("SKU3", d(2026, 10, 19), None, None, None, None, budget="MT   ", model=None, stylist=None),
        # date typed in the Editor Name column, odd status
        _row("SKU4", d(2026, 10, 13), d(2026, 10, 13), "Done", "Done", "Pending on boy", pdf_editor=d(2026, 10, 22)),
        # duplicate SKU + assets that look like a serial number
        _row("SKU1", d(2026, 10, 13), d(2026, 10, 13), None, None, None, assets=20, dup=20),
    ]
    for r in rows:
        ws.append(r)
    ws.append([None] * 24 + [d(2026, 10, 13)])  # junk row without SKU -> ignored
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@pytest.fixture(scope="module")
def parsed():
    df, issues = tk.parse_tracker(make_workbook(), tk.TRACKERS["Photo Shoot"], today=datetime(2026, 10, 27).date())
    return df, issues


def test_rows_and_dates(parsed):
    df, _ = parsed
    assert len(df) == 5  # junk row dropped
    assert list(df["sheet_row"]) == [2, 3, 4, 5, 6]
    assert df.loc[1, "shoot_date"] == pd.Timestamp("2026-10-09")  # text date read day-first
    assert df.loc[0, "days_inward_to_shoot"] == 2


def test_status_and_names(parsed):
    df, _ = parsed
    assert df.loc[0, "Master Poster__status"] == "Done"
    assert df.loc[1, "PDF__status"] == "Outsourced"
    assert df.loc[3, "MT catalog Image__status"] == "Other"
    assert df.loc[0, "Master Poster__editor"] == "Suvin"
    assert df.loc[0, "stylist"] == df.loc[1, "stylist"] == "Mustafa / Sriya"
    assert df.loc[0, "budget"] == "GT + MT" and df.loc[2, "budget"] == "MT"
    assert df.loc[2, "shoot_status"] == "Pending"
    # Date in the Editor Name column becomes the edit date
    assert df.loc[3, "PDF__edit_date"] == pd.Timestamp("2026-10-22")
    assert pd.isna(df.loc[3, "PDF__editor"])


def test_stage_switch(parsed):
    df, _ = parsed
    cfg = tk.TRACKERS["Photo Shoot"]
    on = tk.add_stage(df, cfg, outsourced_counts_as_done=True)
    off = tk.add_stage(df, cfg, outsourced_counts_as_done=False)
    assert list(on["stage"]) == ["Complete", "Complete", "Pending shoot", "In editing", "In editing"]
    assert off.loc[1, "stage"] == "In editing"
    assert on.loc[3, "waiting_for"] == "MT Catalog Image (Pending on boy)"
    assert on.loc[2, "waiting_for"].startswith("Shoot")


def test_issues(parsed):
    _, issues = parsed
    kinds = set(issues["Issue"])
    assert "Date typed in the Editor Name column" in kinds
    assert "SKU appears more than once" in kinds
    assert "Status not recognised (use Done / Outsource / blank)" in kinds
    assert "Same number as the Duplicate column - looks like a serial number" in kinds
    assert "Date typed as text (works, but format the column as Date)" in kinds


def test_required_by_budget():
    cfg = tk.TrackerConfig(
        name="Photo Shoot", sheet="Photo Shoot", crew_col="Photographer", crew_label="Photographer",
        deliverables=("Master Poster", "PDF", "MT catalog Image"),
        required_by_budget=(("GT", ("Master Poster", "PDF")), ("MT", ("MT catalog Image",))),
    )
    df, _ = tk.parse_tracker(make_workbook(), cfg, today=datetime(2026, 10, 27).date())
    out = tk.add_stage(df, cfg)
    assert out.loc[2, "Master Poster__status"] == "Not needed"  # MT-only SKU


def test_helpers():
    assert tk.extract_file_id("https://docs.google.com/spreadsheets/d/1tog-cqkNHR-nOuVKOVv45HY5CGeeJQGz/edit#gid=0") \
        == "1tog-cqkNHR-nOuVKOVv45HY5CGeeJQGz"
    assert tk.parse_date("23/10/2026")[0] == pd.Timestamp("2026-10-23")
    assert tk.parse_date(46318)[1] == "date"  # Excel serial number
    assert tk.parse_date("Outsource")[1] == "text"
    assert tk.norm_people("20") is None


def test_app_runs(tmp_path, monkeypatch):
    """Run the whole Streamlit app headless on the fake tracker."""
    from streamlit.testing.v1 import AppTest

    f = tmp_path / "tracker.xlsx"
    f.write_bytes(make_workbook())
    monkeypatch.setenv("TRACKER_FILE", str(f))
    monkeypatch.chdir(ROOT)
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=60).run()
    assert not at.exception, [e.value for e in at.exception]
    assert [t.label for t in at.tabs][0] == "Overview"
    at.toggle[0].set_value(False).run()
    assert not at.exception


TAT_HEADERS = ["Shoot TAT", "TAT Poster", "TAT PDF", "TAT MT", "Actual Shoot", "Actual Poster", "Actual PDF",
               "Actual MT", "Variance Poster", "Variance PDF", "Variance MT", "TAT Status"]


def make_tat_workbook() -> bytes:
    """New sheet layout: TAT columns after Touch Point (the sample rows the team shared)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Photo Shoot"
    headers = PHOTO_HEADERS[:PHOTO_HEADERS.index("Touch Point") + 1] + TAT_HEADERS
    ws.append(headers)
    d = datetime
    base = _row("MTP990OGY", d(2026, 8, 4), d(2026, 8, 5), "Done", "Done", "Done",
                mp_date=d(2026, 8, 12), mp_editor="Suvin", pdf_date=d(2026, 8, 26), pdf_editor="Suvin",
                mt_date=d(2026, 8, 14), mt_editor="Suvin")[:len(headers) - len(TAT_HEADERS)]
    ws.append(base + [2, 7, 15, 3, 1, 8, 22, 9, 1, 7, 6, None])
    # Open SKU: inward 13 Oct, nothing edited yet, blank TAT cells -> defaults 7 / 15 / 3
    open_row = _row("OPEN1", d(2026, 10, 13), d(2026, 10, 13), None, None, None)[:len(headers) - len(TAT_HEADERS)]
    ws.append(open_row + [None] * len(TAT_HEADERS))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_tat():
    cfg = tk.TRACKERS["Photo Shoot"]
    today = datetime(2026, 10, 27).date()
    df, issues = tk.parse_tracker(make_tat_workbook(), cfg, today=today)
    # Inward -> deliverable days
    assert df.loc[0, "Master Poster__days"] == 8
    assert df.loc[0, "PDF__days"] == 22
    # MT counts from the shoot date (5 Aug -> 14 Aug = 9), same as the sheet's Actual MT
    assert df.loc[0, "MT catalog Image__tat_days"] == 9
    assert "Actual days in the sheet don't match the dates" not in set(issues["Issue"])
    # Shoot TAT: inward 4 Aug -> shoot 5 Aug = 1 day, within 2
    assert df.loc[0, "Shoot__tat_days"] == 1
    assert df.loc[0, "Shoot__tat_outcome"] == "On time"
    assert df.loc[1, "Shoot__tat_outcome"] == "On time"      # inward and shoot on 13 Oct
    # Against TAT 7 / 15 / 3
    assert df.loc[0, "Master Poster__tat_outcome"] == "Late"
    assert df.loc[0, "Master Poster__tat_days_over"] == 1
    assert df.loc[0, "PDF__tat_days_over"] == 7
    # Open SKU, inward + shoot 13 Oct: TAT dates 20 Oct (Poster), 28 Oct (PDF), 16 Oct (MT)
    assert df.loc[1, "Master Poster__tat_outcome"] == "Overdue"
    assert df.loc[1, "PDF__tat_outcome"] == "Within TAT"
    assert df.loc[1, "PDF__tat_due"] == pd.Timestamp("2026-10-28")
    assert df.loc[1, "MT catalog Image__tat_outcome"] == "Overdue"
    summary = tk.tat_summary(df, cfg)
    assert list(summary["Deliverable"]) == ["Shoot", "Master Poster", "PDF", "MT Catalog Image"]
    assert list(summary["Completed within TAT"]) == [2, 0, 0, 0]
    assert list(summary["Overdue"]) == [0, 1, 0, 1]
    assert list(summary["Pending · still within TAT"]) == [0, 0, 1, 0]
    view = tk.tat_view_by_deliverable(df, cfg)
    assert view.loc["PDF", "Pending · still within TAT"] == 1
    pending = tk.tat_pending(df, cfg, today)
    assert len(pending) == 3
    assert pending.iloc[-1]["Status"] == "Due in 1 day"   # PDF, due 28 Oct


def test_tat_from_inward_date():
    base = tk.TRACKERS["Photo Shoot"]
    rules = tuple(tk.TatRule(r.deliverable, r.target_col, r.actual_col, r.variance_col, r.default_days, start="inward")
                  for r in base.tat)
    cfg = tk.TrackerConfig(name=base.name, sheet=base.sheet, crew_col=base.crew_col, crew_label=base.crew_label,
                           deliverables=base.deliverables, labels=base.labels, tat=rules)
    df, issues = tk.parse_tracker(make_tat_workbook(), cfg, today=datetime(2026, 10, 27).date())
    assert df.loc[0, "MT catalog Image__tat_days"] == 10
    flagged = issues[issues["Issue"] == "Actual days in the sheet don't match the dates"]
    assert list(flagged["Column"]) == ["Actual MT"]   # sheet says 9 (from shoot date)
