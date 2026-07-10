# Hamleys Global Count (GC) Dashboard

A 100% client-side, GitHub Pages–ready executive dashboard for monitoring daily
Global Count activity across Hamleys India stores. No backend, no database —
just HTML, CSS and JavaScript, powered by Chart.js and SheetJS.

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
│   ├── charts.js            → Chart.js wrapper (all chart types + PNG export)
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

1. Open your published site and click the **⚙ Admin** button (top-right of
   the dashboard header), or go directly to `.../admin.html`.
2. Drag and drop (or browse to) today's **Global Count Excel** file into the
   first upload box.
   - The **first row** of the **first sheet** is used as column headers.
   - **Every column is preserved** — nothing is hardcoded, so extra columns
     added later will still flow through to the Raw Data tab automatically.
3. The generator will: read the file → convert to JSON → validate required
   columns (Store Code, Store Name, Date, No. of SKU's Counted) → auto-fill
   any missing RM / ROM / SD using `store-mapping.json` → build the final
   dataset.
4. Review the stats cards and the 15-row preview table.
5. Click **⬇ Download Updated gc-data.json**.
6. Replace the existing `data/gc-data.json` file in your GitHub repository
   with the downloaded file (upload it via the GitHub web UI — "Add file →
   Replace this file" — or `git add data/gc-data.json && git commit && git push`).
7. Refresh the live dashboard (or click the **⟳ Refresh** button) — the new
   data appears everywhere (KPIs, filters, charts, tables) **with no code
   changes required**.

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
- **Filters** — Date, Month, Year, RM, ROM, SD, Region (auto-derived from
  store name), Store search, Reset.
- **Tabs** — Overview, RM Summary, ROM Summary, SD Summary, Store Summary,
  Daily Trends, Monthly Trends, Raw Data.
- **Charts** — bar, line, pie, doughnut, area, horizontal bar, top/bottom 10,
  heat-map style calendar — built with Chart.js, animated, with tooltips and
  PNG download on every chart.
- **Raw Data tab** — every column from your Excel, dynamically rendered,
  with search, column sort, pagination, Export Excel/CSV and Print.
- Auto-refresh every 5 minutes, animated KPI counters, responsive layout for
  mobile/tablet/desktop.

---

## 7. Troubleshooting

| Symptom | Fix |
|---|---|
| Dashboard shows "Could not load data/gc-data.json" | Make sure `data/gc-data.json` exists at that exact path in your repo and GitHub Pages has finished deploying (check the **Actions** tab). |
| Admin page says "Missing required column(s)" | Your Excel must contain columns matching *Store Code*, *Store Name*, a *Date* column and a *SKU/Global Count* column (names are matched flexibly, e.g. "No. of SKU's Counted"). |
| RM/ROM/SD show as "Unassigned" | The store code in your daily file isn't present in `data/store-mapping.json`. Update the mapping (Section 5) and re-generate. |
| Changes don't show after upload | Hard-refresh the browser (Ctrl/Cmd+Shift+R) or click **⟳ Refresh** — GitHub Pages / your browser may cache the JSON briefly. |
| Auto-Publish fails with 401/403 | Your Personal Access Token is invalid, expired, or lacks **Contents: Read and write** permission on the repo. |
| Charts look broken/blank | Check your internet connection — Chart.js and SheetJS load from a CDN (`cdnjs.cloudflare.com`). No local install is required. |
| Logo missing | Confirm `assets/logo.png` was uploaded and the path/case matches exactly (GitHub Pages is case-sensitive). |

---

## 8. Technical Notes

- Pure static site: **HTML + CSS + JavaScript only** — no Node.js, no npm
  build step, no server.
- External libraries loaded via CDN: **Chart.js 4.4**, **SheetJS (xlsx) 0.18**.
- All computation (KPIs, aggregations, growth %, completion %, rankings)
  happens client-side in `js/app.js` against the JSON produced by the Admin
  page — the same JSON structure works regardless of how many columns your
  Excel export contains.
- Dark/Light mode and GitHub publish settings persist via `localStorage`.

Built for Hamleys India Retail Operations.
