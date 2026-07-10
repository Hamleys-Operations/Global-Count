# Hamleys Global Count (GC) Dashboard

A 100% client-side, GitHub Pages–ready dashboard for monitoring daily
Global Count activity across Hamleys India stores. No backend, no database —
just HTML, CSS and JavaScript, powered by SheetJS. Purely tabular (no charts)
by design — KPI cards plus RM / ROM / SD / Date wise Completion / Store wise GC
/ Raw Data tables.

---

## 1. Folder Structure

```
Hamleys-GC-Dashboard/
│
├── index.html            → Main executive dashboard
├── admin.html             → Admin Generator (Excel → JSON)
├── css/
│   └── style.css          → Hamleys-branded theme (red / white / gold)
├── js/
│   ├── app.js              → Data loading, filters, KPIs, tables, exports
│   └── excel.js             → SheetJS Excel→JSON engine (Admin page)
├── assets/
│   └── logo.png            → Hamleys logo (used on every page)
├── data/
│   ├── gc-data.json         → Current Global Count dataset (dashboard reads this)
│   └── store-mapping.json   → Store ↔ RM / ROM / SD master mapping
└── README.md
```

---

## 2. One-Time Setup — Upload to GitHub

1. Create a new repository on GitHub, e.g. `Hamleys-GC-Dashboard`.
2. Upload **all files and folders exactly as-is**, keeping the structure above
   (drag the whole `Hamleys-GC-Dashboard` folder contents into the repo, or use
   `git add . && git commit -m "Initial dashboard" && git push`).
3. Do not rename `data/gc-data.json` or `data/store-mapping.json` — the app
   looks for these exact paths.

## 3. Enable GitHub Pages

1. In your repository, go to **Settings → Pages**.
2. Under **Build and deployment**, set **Source** to `Deploy from a branch`.
3. Choose branch `main` (or `master`) and folder `/ (root)`. Click **Save**.
4. GitHub will publish your site at:
   `https://<your-username>.github.io/Hamleys-GC-Dashboard/`
5. Wait 1–2 minutes for the first deployment, then open the URL.

---

## 4. Daily Workflow — Uploading Today's Excel

You only need to upload **the new day's rows** each time (e.g. just today's
date) — you do **not** need to re-upload the full history. The Admin
Generator fetches what's already published and merges your upload into it.

1. Open your published site and click the **⚙ Admin** button (top-right of
   the dashboard header), or go directly to `.../admin.html`.
2. Leave **"Merge with already-published data"** checked (this is the
   default) — it fetches the current live `data/gc-data.json`, adds your
   new rows, and leaves every previously-published row untouched. If you
   upload a store/date combination that's already published, that row is
   *corrected* with your new values instead of being duplicated. Uncheck
   this only if you deliberately want to replace the entire dashboard with
   just this one file.
3. Drag and drop (or browse to) today's **Global Count Excel** file into the
   upload box.
   - The **first row** of the **first sheet** is used as column headers.
   - **Every column is preserved** — nothing is hardcoded, so extra columns
     added later will still flow through to the Raw Data tab automatically.
   - The generator also defends against workbooks with a stale/undersized
     sheet dimension (a common cause of only *some* rows being picked up) by
     recomputing the true used range from the sheet's actual cells before
     converting — no rows are silently dropped.
4. The generator will: read the file → convert to JSON → validate required
   columns (Store Code, Store Name, Date, No. of SKU's Counted) → auto-fill
   any missing RM / ROM / SD using `store-mapping.json` → merge with the
   already-published data → build the final dataset.
5. Review the stats cards (now reflecting the **full** merged dataset) and
   the 15-row preview table.
6. Click **⬇ Download Updated gc-data.json**.
7. Replace the existing `data/gc-data.json` file in your GitHub repository
   with the downloaded file (upload it via the GitHub web UI — "Add file →
   Replace this file" — or `git add data/gc-data.json && git commit && git push`).
8. Refresh the live dashboard (or click the **⟳ Refresh** button) — the new
   data appears everywhere (KPIs, filters, tables) **with no code changes
   required**.

### Optional: Skip the manual re-upload (Auto-Publish to GitHub)

In Admin → **Section 2**, you can save a GitHub Personal Access Token
(fine-grained, scoped to *this repo only*, permission **Contents: Read and
write**) together with your username/repo/branch. After generating the JSON,
click **☁ Auto-Publish to GitHub** and the file is committed directly —
no manual download/upload needed. The token is stored only in your browser's
local storage and is sent only to GitHub's official API.

---

## 5. Updating the Store ↔ RM / ROM / SD Mapping

Only needed when stores open/close or reporting lines change:

1. In Admin → **Section 3**, drop the updated **SD ROM Store Mapping** Excel
   file (columns: Store Code, Name, RM, SD, ROM — order doesn't matter).
2. Click **⬇ Download store-mapping.json** and replace
   `data/store-mapping.json` in your repo the same way as above.
3. Re-generate today's `gc-data.json` afterwards so it picks up the refreshed
   mapping.

---

## 6. What's on the Dashboard

- **Header** — logo, title, last-updated date, Refresh, Admin, Dark/Light mode.
- **KPI cards** — Total Global Count, Total Stores, Stores Uploaded, RM/ROM/SD
  counts, Average/Highest/Lowest GC, Completion %.
- **Filters** — Date From / Date To (calendar-range pickers), Month, Financial
  Year (Apr–Mar, e.g. "FY 2026-27"), RM, ROM, SD, Store search, Reset.
- **Tabs**:
  - **RM Summary** — RM, Total Stores, GC Stores, Completion %, Total Net Qty
    Diff, Total Net MAP Value Diff.
  - **ROM Summary** — same column layout, grouped by ROM.
  - **SD Summary** — same column layout, grouped by SD. Any SD name that is
    also a known ROM name is automatically excluded here (that person is
    acting as the ROM for that store, not a distinct SD), so names never
    repeat between the ROM and SD tables.
  - **Date wise Completion** — Date, Total Stores, GC Stores, Completion %,
    one row per calendar date on which a Global Count was recorded.
  - **Store wise GC** — one row per Global Count entry, in the same column
    shape as the source Excel: Store Code, Store Name, Date, No. of SKU's
    Counted, Shortage Qty, Shortage MAP Value, Excess Qty, Excess MAP Value,
    Total Net Qty Difference, Total Net MAP Value Difference, Global Count
    Done by, Global Count Validated by, Shortage Qty Moved to 2997 ?, Reason
    for Short or Excess Qty — with search, column sort, pagination and
    Export Excel.
  - **Raw Data** — every column from your Excel, dynamically rendered, with
    search, column sort, pagination, Export Excel/CSV and Print.
- Auto-refresh every 5 minutes, animated KPI counters, responsive layout for
  mobile/tablet/desktop. No charts — this build is intentionally table-first.

---

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| Dashboard shows "Could not load data/gc-data.json" | Make sure `data/gc-data.json` exists at that exact path in your repo and GitHub Pages has finished deploying (check the **Actions** tab). |
| Admin page says "Missing required column(s)" | Your Excel must contain columns matching *Store Code*, *Store Name*, a *Date* column and a *SKU/Global Count* column (names are matched flexibly, e.g. "No. of SKU's Counted"). |
| RM/ROM/SD show as "Unassigned" | The store code in your daily file isn't present in `data/store-mapping.json`. Update the mapping (Section 5) and re-generate. |
| Changes don't show after upload | Hard-refresh the browser (Ctrl/Cmd+Shift+R) or click **⟳ Refresh** — GitHub Pages / your browser may cache the JSON briefly. |
| Auto-Publish fails with 401/403 | Your Personal Access Token is invalid, expired, or lacks **Contents: Read and write** permission on the repo. |
| Export buttons don't work | Check your internet connection — SheetJS loads from a CDN (`cdnjs.cloudflare.com`). No local install is required. |
| Logo missing | Confirm `assets/logo.png` was uploaded and the path/case matches exactly (GitHub Pages is case-sensitive). |
| Dashboard breaks after an update | `index.html` and `js/app.js` change together — always upload **all** files from this folder in the same commit, 