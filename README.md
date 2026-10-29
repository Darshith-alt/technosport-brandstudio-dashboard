# Brand Studio – Shoot Dashboard

Streamlit dashboard for the **Master Tracker – Brand Studio** sheet (Photo Shoot and Video Shoot tabs).
Blue and orange TechnoSport theme. Reads the Google Sheet directly, so it updates when the team updates the sheet.

**What it shows**

- **KPIs:** SKUs inwarded, shot, waiting for shoot, in editing
- **Turnaround vs TAT cards:** Inward → Shoot, Inward → Master Poster, Inward → PDF, Shoot → MT Catalog Image. Each shows total SKUs and
  styles, completed within TAT, overdue (TAT date passed) and pending still within TAT, each with its %
- **TAT summary by step table:** below the cards, one row per step plus an All steps total, with Completed within TAT,
  Overdue and Pending still within TAT as counts and % of SKUs
- **TAT tab:** results per deliverable, average days vs TAT, pending items by TAT due date, and a downloadable
  list of pending items with their TAT due dates
- **Overview:** pipeline status by step (Shoot, Master Poster, PDF, MT Catalog Image), progress by super category and project, inwarded vs shot per week
- **Breakdowns:** progress by vertical, budget DPT (GT / MT), ideal for, and a style-code table
- **Team:** SKUs shot by photographer, model and stylist; deliverables done by editor
- **Pending:** every open SKU, oldest first, with what it is waiting for (downloadable)
- **Data quality:** rows in the sheet that need fixing, with row numbers
- **All data:** the cleaned, filtered data (downloadable)
- **Filters:** inward date, project, super category, budget, vertical, ideal for, model, stylist, SKU search
- **Video Shoot** tracker uses the same screens (switch in the sidebar)

---

## Files

| File | What it does |
|---|---|
| `app.py` | The dashboard (layout, filters, KPIs, tabs) |
| `tracker.py` | Reads and cleans the sheet, downloads it from Google Drive, finds data issues |
| `charts.py` | Charts and colours |
| `.streamlit/config.toml` | Blue and orange theme |
| `.streamlit/secrets.toml.example` | Template for the Google connection and password |
| `tests/test_tracker.py` | Automatic tests (use fake data, safe for GitHub) |
| `.github/workflows/tests.yml` | Runs the tests on GitHub every time you push |
| `data/` | Put the tracker .xlsx here for local testing. It is never uploaded to GitHub |

---

## 1. Run it on your computer

**Using Jupyter / Anaconda?** In a notebook saved inside this folder, run these in three cells:

```python
import sys; print(sys.version)          # must be 3.11 or newer
%pip install -U streamlit plotly openpyxl google-auth requests
!streamlit run app.py                     # opens http://localhost:8501; press Stop (■) to close
```

Restart the kernel after the install cell. If Python is older than 3.11, open **Anaconda Prompt** and use a separate environment:
`conda create -n dashboard python=3.12 -y`, `conda activate dashboard`, `cd` into this folder,
`pip install -r requirements.txt`, `streamlit run app.py`.

**Without Jupyter:**

Needs Python 3.11 or newer (3.12 recommended).

```bash
cd photo-shoot-dashboard
python -m venv .venv
# Windows:  .venv\Scripts\activate
# Mac/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

Download the tracker (File → Download → Microsoft Excel) and save it in the `data/` folder, then:

```bash
streamlit run app.py
```

The dashboard opens at http://localhost:8501. You can also upload a file from the sidebar (**Use a different file**).

Run the tests (optional):

```bash
pip install -r requirements-dev.txt
pytest -q
```

---

## 2. Connect the live Google Sheet (recommended: service account)

This keeps the sheet private. Only the dashboard can read it.

1. Open https://console.cloud.google.com → create a project (e.g. `brand-studio-dashboard`).
2. **APIs & Services → Library →** enable **Google Drive API**.
3. **IAM & Admin → Service accounts → Create service account** (any name, no roles needed).
4. Open the service account → **Keys → Add key → Create new key → JSON**. A .json file downloads. Keep it private.
5. Open the tracker in Google Sheets → **Share** → add the service account's `client_email`
   (looks like `name@project.iam.gserviceaccount.com`) as **Viewer**.
6. Copy `.streamlit/secrets.toml.example` to `.streamlit/secrets.toml` and fill it in:
   - `file_id`: the tracker link or its ID
   - `[gcp_service_account]`: copy each value from the .json key. Keep `private_key` on one line with the `\n` as in the file.
   - `APP_PASSWORD`: optional password screen. Delete the line to turn it off.
7. Run `streamlit run app.py` again. The sidebar should say **Source: Google Drive · …**

Works whether the tracker is a Google Sheet or an .xlsx file stored in Drive.
Data is cached for 10 minutes; **Refresh data** in the sidebar reloads it immediately.

> If Google Cloud says "Service account key creation is disabled", your Workspace admin blocks keys.
> Ask IT to allow it for this project, or use the simpler option below.

**Simpler option (less private):** set the sheet to *Anyone with the link → Viewer* and in `secrets.toml` use
`access = "public_link"` (no `[gcp_service_account]` needed). Anyone who has the sheet link can then see the data.

---

## 3. Put the code on GitHub

1. On GitHub: **New repository** → name `brand-studio-dashboard` → **Private** → Create.
2. On your computer, in the project folder:

```bash
git init
git add .
git commit -m "Brand Studio shoot dashboard"
git branch -M main
git remote add origin https://github.com/<your-username>/brand-studio-dashboard.git
git push -u origin main
```

`.gitignore` keeps `secrets.toml` and the `data/*.xlsx` file off GitHub.
Use the commands above (or GitHub Desktop), not drag-and-drop on the GitHub website: the website upload ignores `.gitignore`.

After the push, open the **Actions** tab on GitHub. The **Tests** workflow should show a green tick.

---

## 4. Publish on the web (Streamlit Community Cloud, free)

1. Go to https://share.streamlit.io and sign in with GitHub. Allow access to private repos when asked.
2. Click **Create app** (top right) → **Yup, I have an app**.
   - Repository: `<your-username>/brand-studio-dashboard`, Branch: `main`, Main file path: `app.py`
   - Optional: pick a short **App URL**, e.g. `ts-brand-studio`
   - **Advanced settings:** Python version 3.12, and paste the full contents of your `secrets.toml` into **Secrets** → Save
3. **Deploy.** You get a link like `https://ts-brand-studio.streamlit.app`.
4. Every `git push` to `main` updates the live app automatically.
5. **Who can see it:** an app deployed from a private repo is private by default. Use **Share** on the app to invite
   people by email (they sign in with Google or an email link). The free plan allows one private app.
   `APP_PASSWORD` adds a second lock if you want one.

---

## How the numbers work

| Term | Meaning |
|---|---|
| SKUs inwarded | Rows in the sheet with a SKU (after filters) |
| Shot | Shoot date filled, or shoot marked Outsource |
| Waiting for shoot | No shoot date yet |
| Deliverable status | `Done` / `done` → Done · `Outsource` → Outsourced · blank → Pending · anything else → Other |
| Complete | Every deliverable is Done (and Outsourced, when **Count outsourced as complete** is on). Used in the Overview charts and the Pending tab |
| In editing | Shot, but not every deliverable is complete |
| Inward → shoot | Shoot date minus inward date, average in days |
| Shoot → last edit | Latest edit date of the SKU's deliverables minus shoot date (only SKUs with edit dates) |
| Inward → Master Poster / PDF / MT Catalog Image | That deliverable's Edit Date minus inward date, average in days |
| Inward → all edits done | Latest Edit Date minus inward date, for fully complete SKUs |
| TAT target | From *Shoot TAT / TAT Poster / TAT PDF / TAT MT* in the sheet. Blank cells use 2 / 7 / 15 / 3 days |
| TAT start | Shoot, Master Poster and PDF count from the inward date; MT Catalog Image counts from the shoot date |
| TAT date | Start date + TAT days (the date the deliverable should be done by) |
| Pending, still within TAT | Not done yet and the TAT date hasn't passed |
| On time / Late | Finished (has an Edit Date) within / after the TAT |
| Overdue | Not finished and the days since inward are already past the TAT |
| Done, no edit date | Marked Done or Outsource but the Edit Date is blank, so the days can't be measured |

Dates are read as **day/month/year**, whether typed as a date or as text.

---

## Change things later

- **TAT targets and start date:** `tat=(...)` under Photo Shoot in `tracker.py`. Each line has the TAT column names, the
  default days for blank cells, and `start="inward"`. Change it to `start="shoot"` to count that deliverable from
  the shoot date instead.

- **Colours:** `charts.py` (top of file) and `.streamlit/config.toml`.
- **Deliverables / sheet names:** `TRACKERS` in `tracker.py`. The loader finds each deliverable's *Edit Date* and
  *Editor Name* as the two columns just before it, so keep that order in the sheet.
- **Deliverables that depend on budget:** if, for example, GT SKUs don't need the MT Catalog Image, set
  `required_by_budget` in `tracker.py`:

  ```python
  required_by_budget=(("GT", ("Master Poster", "PDF")), ("MT", ("Master Poster", "MT catalog Image"))),
  ```

  Deliverables a SKU doesn't need then show as *Not needed* and don't block *Complete*.

## Keep the sheet dashboard-friendly

- Keep headers in row 1 and don't rename columns.
- Format date columns as **Date** (Format → Number → Date).
- Use `Done`, `Outsource`, or leave blank in status columns.
- Put names only in Editor / Stylist / Model columns (the **Data quality** tab lists rows that break this).
