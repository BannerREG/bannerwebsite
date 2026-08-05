# PROJECT.md — Banner Rent Comp Dashboard

> **This file is the single source of truth for this repository.** It is written for both humans and AI coding agents (Claude Code, Codex, Gemini CLI, Cursor, Windsurf, etc.). Assume any AI reading this has no memory of prior sessions — everything needed to work productively should be here or linked from here.

---

## 1. Executive Summary

**What it is.** A single-page, client-side web application ("Rent Comp Dashboard") that lets Banner Real Estate Group's team upload competitor apartment-property rent rolls (Yardi and CoStar exports), describe a subject property, and instantly generate a full rent-comparison analysis: weighted-average tables, charts, rent-growth trends, a suggested unit mix, suggested pricing, an interactive map, and a fully-formatted, chart-ready Excel export.

**Why it exists.** Before this tool, comp analysis for multifamily development deals was done manually in Excel — copying rent-roll data out of Yardi PDFs/exports, retyping it, and hand-building pivot tables and charts for every new deal. That was slow and error-prone. This app automates the entire pipeline: upload → normalize → compare → export, in minutes instead of hours.

**Who it serves.** Internal users at Banner Real Estate Group (analysts, developers, underwriters) evaluating multifamily development or acquisition opportunities. It is not a public product — it is gated behind a shared password and treated as an internal tool.

**Long-term vision.** Become the team's default first step for any new deal's rent comp analysis — a single shared, always-up-to-date tool (not a spreadsheet template passed around by email) that every analyst uses the same way, with saved comparisons visible to the whole team via a shared Google Drive.

**MVP vision (achieved).** Upload Yardi/CoStar comp files, enter a subject property, get weighted-average comparison tables and charts, and export everything to a polished Excel workbook with real (not pasted-image) native charts.

**What success looks like.** An analyst can go from "here are 3–5 comp exports" to "here is a client-ready Excel comp package" in under 10 minutes, with no manual data entry beyond describing the subject property, and the resulting numbers are trustworthy enough to put in front of a lender or investment committee without a manual double-check.

---

## 2. Product Goals

### Primary goals
- Eliminate manual re-typing of Yardi/CoStar rent-roll data.
- Produce consistent, weighted-average comp math every time (no more "which analyst's spreadsheet formula do we trust").
- Make comparisons easy to save, share, and revisit — across the whole team, not just one person's laptop.
- Produce a client/lender-ready Excel export with **native, editable** charts (not screenshots).
- Give data-driven starting points for two judgment calls that are normally pure guesswork: what unit mix to build, and what rents to set.

### Non-goals
- This is **not** a full real-estate underwriting/proforma model (no debt sizing, no IRR, no waterfall — see the separate `multifamily-proforma` project for that, which is unrelated to this repo).
- Not a multi-tenant SaaS product — it is a single internal tool for one organization.
- Not designed to store or process personally identifiable information beyond property addresses.
- No real user-account system — access control is a single shared password, not per-user authentication.

### Future roadmap (not committed, directional only)
- Possible: per-user accounts instead of a single shared password.
- Possible: richer map-based comp discovery (e.g., pull comps by drawing a radius).
- Possible: direct API integration with Yardi (no more manual file export) if that becomes feasible.
- Possible: historical trend tracking across saved comparisons over time (e.g., "how has this submarket's $/SF moved over the last year across all our saved comps").

### Problems being solved
1. Manual data entry from Yardi/CoStar exports into ad hoc spreadsheets.
2. Inconsistent bedroom-type bucketing between analysts (e.g., is "1x1 Alcove" a Studio or its own category?).
3. No shared, versioned place to find a past comp analysis — everything lived in individual email attachments.
4. Unit mix and pricing recommendations were pure gut-feel, not tied to what the actual comp data showed.
5. Charts in Excel exports were static pasted images that couldn't be tweaked by the recipient.

### Core user experience
Setup tab (upload + describe subject) → Dashboard (charts) → Tables (detailed comparisons) → Map (visualize locations) → Rent Growth (trends over time) → Save/Export. See [Section 4.3](#43-application-structure-tabs) for the full tab breakdown.

---

## 3. Current Development Stage

**Current phase:** Actively maintained, feature-complete MVP+. The core pipeline (upload → compare → export) has been stable for a while; recent work has been incremental feature additions and polish (dark mode, mapping, unit-mix suggestions) driven directly by user feedback in real usage.

### Completed milestones
- Yardi "Property Detail – Rental Rate History" file parsing (multi-sheet: rent + $/SF).
- CoStar "Unit Mix" export parsing (different layout, auto-detected).
- Weighted-average comparison tables (Summary + per-bedroom-type detail).
- Dashboard charts (rent, $/SF, unit mix distribution, unit count, avg SF, rent premium/discount, bubble chart).
- Rent Growth (%) trend charts, indexed to each comp's first tracked month.
- Native Excel export — every chart is a real OOXML chart object (hand-built XML, not an image), not just a data dump.
- Save/load comparisons, synced to a shared Google Drive folder via a Google Apps Script backend (so saves are visible to the whole team, not just one browser).
- Share-by-link (compressed URL) and JSON import/export of a single comparison.
- Data-driven "Suggest Pricing" (benchmarks subject rents to the top comp(s) per bedroom type).
- Data-driven "Suggest Unit Mix" (derives a recommended unit count + SF per bedroom type from comp data — see [Section 4.6](#46-suggest-unit-mix-algorithm)).
- Interactive comp map (Leaflet + OpenStreetMap tiles, geocoded via U.S. Census Bureau with an OpenStreetMap/Nominatim fallback, proxied through the same Apps Script backend).
- Site-wide password gate (client-side only — see [Section 7 — Secrets Management](#secrets-management) for the important caveat).
- Dark mode toggle, verified for text/background contrast across every tab.

### Work in progress
- None actively mid-flight as of this writing. The repository is in a stable state between feature requests.

### Upcoming milestones
- None formally scheduled. This project is driven by ad hoc feature requests from actual usage, not a fixed backlog. See [Section 11 — Backlog](#11-backlog) for known candidate improvements.

### Known blockers
- None blocking current functionality.

### Open questions
- **Ownership transition:** the original developer (a former intern) is transitioning this project to a company-owned GitHub account, Google Drive folder, and Apps Script project. As of this writing, some of that transfer may still be in progress — **verify current ownership/access of all three systems before assuming they are all under a permanent company-controlled account.**
- Whether per-user authentication is ever worth the complexity vs. the current single-shared-password model.

---

## 4. Technical Architecture

### 4.1 Overall architecture

This is a **single static HTML file with no build step and no backend server**. There is no `package.json`, no bundler, no framework, no server-side code owned by this repo. All logic — parsing, math, chart rendering, Excel-file generation — runs in the browser.

```
┌─────────────────────────────────────────────────────────────────┐
│  Browser (client)                                                │
│  yardi-compare/index.html  — HTML + CSS + vanilla JS, one file   │
│                                                                    │
│  • Parses uploaded .xlsx/.xlsm/.csv files client-side (SheetJS)  │
│  • Holds all state in memory (properties[], subjectProperty)     │
│  • Renders Chart.js charts, Leaflet map                          │
│  • Builds .xlsx exports client-side (raw OOXML + ExcelJS/JSZip)  │
└───────────────┬───────────────────────────────┬──────────────────┘
                │ fetch()                        │ fetch()
                ▼                                ▼
┌───────────────────────────────┐   ┌──────────────────────────────┐
│ Google Apps Script Web App     │   │ External free APIs           │
│ (owned separately from this    │   │ (called directly from the    │
│  repo — see Section 4.5)       │   │  browser, no proxy)          │
│                                 │   │                               │
│ • list / save / delete saved   │   │ • U.S. Census Bureau          │
│   comparisons (Drive JSON      │   │   geocoder (via the Apps      │
│   files)                       │   │   Script proxy — CORS-        │
│ • geocode action (proxies to   │   │   blocked for direct browser  │
│   Census Bureau, since Census  │   │   calls)                      │
│   has no CORS support)         │   │ • Nominatim/OpenStreetMap     │
│                                 │   │   geocoder (fallback,         │
│ Backed by a shared Google      │   │   CORS-enabled, called        │
│ Drive folder (JSON files, one  │   │   directly)                   │
│ per saved comparison)          │   │ • OpenStreetMap tile server   │
└───────────────────────────────┘   │   (Leaflet map tiles)         │
                                     └──────────────────────────────┘
```

### 4.2 Major components (all inside `index.html`)

| Component | Responsibility |
|---|---|
| File parsers | `parseYardiWorkbook()`, `parseCoStarWorkbook()`, `isCoStarWorkbook()` — read uploaded spreadsheets via SheetJS (`XLSX.read`), normalize into a common `property` shape. |
| Unit-type normalization | `normalizeUnitType()` — maps messy raw strings ("1x1 Alcove", "One Bedroom", etc.) to a canonical bedroom-type bucket (Studio, 1 BR Alcove, 1 BR, 2 BR Alcove, 2 BR, 3 BR, 4 BR, 5 BR). |
| Aggregation | `getAggregatedForType()`, `getFloorplansForType()`, `getAllProperties()` — unit-weighted averaging of rent/SF/$-per-SF across floorplans and properties. |
| Dashboard rendering | `renderDashboard()`, `renderCharts()`, `renderSummaryStats()` — Chart.js bar/bubble/line charts. |
| Tables rendering | `renderTables()`, `renderSummaryTable()`, `renderUnitTypeTables()` — side-by-side comp tables per bedroom type, sorted high-to-low by $/SF, with a synthesized "COMP AVERAGE" row. |
| Rent Growth | `renderRentGrowth()`, `pctGrowthSeries()`, `zeroLineDataset()` — % rent growth over time, indexed to each comp's first tracked month. |
| Suggest Pricing | `suggestPricing()` — benchmarks the subject's rents to the top-priced comp(s) per bedroom type, with market-condition adjustments (occupancy, year-built) and a written rationale. |
| Suggest Unit Mix | `suggestUnitMix()` — see [Section 4.6](#46-suggest-unit-mix-algorithm). |
| Excel export | `exportToExcel()` plus a large block of hand-rolled OOXML generation (`_buildChartXml()`, `_buildDrawingXml()`, `_chartSerXml()`, `_bubbleSerXml()`) layered on top of ExcelJS/JSZip, because ExcelJS alone cannot produce every native chart type this export needs. |
| Comp Map | `plotCompsOnMap()`, `ensureCompMap()`, `geocodeAddress()`/`geocodeViaCensusProxy()`/`geocodeViaNominatim()` — Leaflet map, geocoding with Census-first/Nominatim-fallback. |
| Save/Share | `saveComparisonToDrive()`, `loadComparisonsFromDrive()`, `deleteComparisonFromDrive()`, `shareComparison()`, `exportComparisonJson()`/`importComparisonFile()` — Drive sync plus local-only fallback (`STORAGE_KEY = 'rentCompSaved'` in `localStorage`) and shareable compressed-URL links (via `lz-string`). |
| Auth gate | `submitAuthGate()` — a simple client-side password check gating the whole page. |
| Dark mode | `toggleDarkMode()`, `applyChartTheme()`, `initDarkModeUI()` — CSS-custom-property-driven theme switch, persisted in `localStorage`. |

### 4.3 Application structure (tabs)

The whole app is one page with a tab bar; each tab is a `.tab-content` div toggled by `switchTab()`:

1. **Guide** — static onboarding/help content (default active tab on load).
2. **Setup** — upload comps (drag-and-drop Yardi/CoStar files), enter the subject property and its unit mix, "Suggest Unit Mix" and "Suggest Pricing" buttons, "Apply & Compare".
3. **Dashboard** — summary stat cards + 8 Chart.js charts (rent, $/SF, unit mix %, unit count, avg SF, rent premium/discount $ and %, bubble chart, unit-mix distribution).
4. **Tables** — property summary table + one detailed comparison table per bedroom type (Studio, 1 BR, 2 BR, ...), each with a synthesized COMP AVERAGE row.
5. **Map** — Leaflet map with a pin per loaded property + the subject; click "Plot on Map" to geocode and render.
6. **Rent Growth** — % rent-growth-over-time charts, by bedroom type across comps and by property across its own unit types.
7. **Saved** — list of saved comparisons (merged from Drive + any local-only ones), sortable, with load/share/export/delete actions.

### 4.4 Data model (in-memory only — not persisted as a formal schema)

```js
// A "property" (comp or subject) — see parseYardiWorkbook() / applySubject()
{
  id, name, address, market, submarket,
  totalUnits, yearBuilt, occupancy,      // occupancy is % or null
  unitTypes: [                            // one entry per floorplan/row
    {
      unitType,                           // raw string, e.g. "1x1"
      units, sqft, marketRent,
      rentPerSqft,
      monthlyRents: { "Jan '24": 1850, ... },   // Yardi only — for Rent Growth
      monthlyPerSF:  { "Jan '24": 2.44,  ... }   // Yardi only, from sheet 2
    }, ...
  ],
  monthHeaders: [...]                     // ordered list of month labels present
}
```
Global state: `properties` (array of comps), `subjectProperty` (single object or `null`), `charts` (live Chart.js instances, destroyed/recreated on re-render).

### 4.5 Backend: Google Apps Script + Drive

There is **no backend code in this repository.** The "backend" is a Google Apps Script Web App project that lives on Google's infrastructure (script.google.com), tied to a specific Google account, entirely separate from this git repo. It backs a shared Google Drive folder used as the datastore for saved comparisons.

- **Endpoint:** stored in `DRIVE_SYNC_URL` in `index.html`.
- **Auth:** a shared-secret token (`DRIVE_SYNC_TOKEN` in `index.html`, `SHARED_SECRET` in the Apps Script) passed as a query param / POST body field. This is a basic access gate, **not real authentication** — anyone with the token can read/write the shared Drive folder.
- **Actions implemented in the Apps Script's `doGet`/`doPost`:**
  - `doGet?action=list&token=...` — list saved comparisons (reads `comp_*.json` files from the Drive folder).
  - `doGet?action=delete&id=...&token=...` — trash a saved comparison's file.
  - `doGet?action=geocode&address=...&token=...` — proxy an address to the U.S. Census Bureau geocoder (`geocoding.geo.census.gov`) and return `{lat, lon}` or `null`. Exists purely because the Census API doesn't send CORS headers, so the browser can't call it directly.
  - `doPost` (body: `{token, comparison}`) — create/overwrite `comp_<id>.json` in the Drive folder.
- **Because this backend is not in the repo:** any AI agent working on backend-affecting changes must ask the user for the current Apps Script source (or have them paste it) before proposing edits — do not assume its current contents. See the commit history / prior conversation logs for the exact `doGet`/`doPost` implementation as of the last known-good state.

### 4.6 Suggest Unit Mix algorithm

This is one of the more intricate pieces of business logic in the app (`suggestUnitMix()`), worth understanding in detail before modifying:

1. **Target total units** = average of the loaded comps' own `totalUnits` values (no manual entry).
2. **Pool** every floorplan from every comp into one list of `{bedroomType, sqft, units, $/SF}`.
3. **Bucket** by (bedroom type, SF rounded to nearest 25) so near-duplicate sizes cluster together.
4. **Bedroom-type significance filter** — a type is kept if it clears an absolute bar (≥5 units AND ≥5% of pooled units) **OR** appears in at least half of the loaded comps (added specifically so a type that's a small-but-consistent share of most comps — e.g. always ~3% studios — isn't dropped just because no single comp has "enough" of them). A true one-off (present in only one comp with a handful of units) is excluded.
5. **Mix %** = the *average of each comp's own mix %*, not a pooled/unit-weighted average — so a 300-unit comp doesn't dominate the recommended mix more than a 40-unit comp.
6. **SF pick per surviving type** = the SF bucket with the best unit-weighted $/SF, **restricted to buckets within ±5% of that type's own average SF** (computed the same way the Tables tab's "COMP AVERAGE" row computes it — average of each comp's own weighted-average SF for that type). This constraint exists because $/SF alone favors small "outlier" units (the "small-unit premium"), which would otherwise win purely for being small.
7. **Distribute** the target total across surviving types via largest-remainder rounding, so counts always sum exactly to the target.
8. Rebuilds the subject's unit-mix table with the result (rents left blank) and shows a rationale panel explaining every inclusion/exclusion.

If asked to modify this algorithm, preserve the ordering above — an earlier bug class in this code came specifically from computing per-type SF recommendations *before* the type-level significance filter ran, producing rationale text for types that were later excluded from the table.

### 4.7 Deployment

- **Hosting:** GitHub Pages, serving directly from the `main` branch of `github.com/manderson-source/bannerwebsite`. No build/CI step — GitHub Pages serves `yardi-compare/index.html` as-is. A push to `main` typically goes live within 1–2 minutes.
- **Local dev:** `.claude/launch.json` defines a `yardi-compare` launch config that starts a minimal PowerShell-based static file server (`HttpListener`) on `http://localhost:3000`, serving files out of the `yardi-compare/` folder. There is no `npm run dev` — this is a Windows/PowerShell-specific local server used for testing inside this environment.
- **No environment variables / no `.env` file** — the few secrets that exist (`DRIVE_SYNC_TOKEN`, `SITE_PASSWORD`) are hardcoded directly in `index.html` because this is a static site with no server-side runtime to source secrets from. See [Secrets Management](#secrets-management) below for why this matters.

### 4.8 Third-party / external dependencies (all via CDN, no package manager)

| Library | Purpose |
|---|---|
| SheetJS (`xlsx`) | Parse uploaded Yardi/CoStar `.xlsx`/`.xlsm`/`.csv` files client-side. |
| Chart.js | All dashboard and rent-growth charts. |
| ExcelJS | Base of the Excel export (workbook/sheet/cell writing); native charts are layered on top via hand-built OOXML (see 4.2). |
| JSZip | Used by the Excel export pipeline for zipping the `.xlsx` package. |
| lz-string | Compresses comparison JSON into shareable URLs (`Share Link` feature). |
| Leaflet | Interactive map on the Map tab. |
| Leaflet + OpenStreetMap tile server | Map basemap tiles (`{s}.tile.openstreetmap.org`) — free, no API key. |
| Nominatim (OpenStreetMap) | Fallback geocoder, called directly from the browser (CORS-enabled). |
| U.S. Census Bureau Geocoder | Primary/preferred geocoder for U.S. addresses (street-number accurate), reached only via the Apps Script proxy (no CORS support for direct browser calls). |
| Google Fonts | `Libre Baskerville` (serif headings). |

There is intentionally **no npm, no bundler, no framework** — this is a deliberate simplicity choice (see [Decision Log](#8-decision-log)), not an oversight.

### 4.9 Folder structure

```
Desktop/Project/                    ← git repo root (github.com/manderson-source/bannerwebsite)
├── .claude/
│   └── launch.json                 ← local dev server config (PowerShell static file server)
├── PROJECT.md                      ← this file
├── rent_comp_builder.py            ← legacy/standalone predecessor script (see 4.10) — NOT used by the live site
└── yardi-compare/
    └── index.html                  ← the entire application (HTML + CSS + JS, ~3,200 lines, one file)
```
Sibling folders on disk under `Desktop/Project/` (`deal-tracker/`, `multifamily-proforma/`, `workpulse/`) are **separate, unrelated projects** that happen to live next to this repo on the local machine — they are not part of this git repository and out of scope for this document.

### 4.10 `rent_comp_builder.py` — legacy artifact

A standalone Python script (uses `openpyxl`) that was an earlier, CLI-based version of the same idea: read Yardi files from a local `./comps/` folder, hardcode a subject property at the top of the file, and generate a similar Excel workbook. It predates `index.html` and is **not used by the live site** and not invoked by anything in `index.html`. It is tracked in git for historical reference only. Do not extend it — all active development happens in `index.html`. Flag to the user (don't just delete it) if its continued presence in the repo ever seems worth revisiting.

### 4.11 Data flow (typical session)

```
User drags Yardi/CoStar files onto Setup tab
        │
        ▼
handleFiles() → isCoStarWorkbook()? → parseCoStarWorkbook() ─┐
                              └─ else → parseYardiWorkbook() ─┤
                                                                ▼
                                                    properties.push(parsed)
                                                    renderPropertyCards()
        │
        ▼
User fills in subject property fields (+ optionally "Suggest Unit Mix" / "Suggest Pricing")
        │
        ▼
applySubject() → subjectProperty = {...} → switchTab('dashboard')
        │
        ▼
renderDashboard() / renderTables() / renderRentGrowth() / renderMapTab()
   (all pure functions of properties[] + subjectProperty — re-run on demand,
    no persistent client-side database, everything lives in memory until
    Save or Export)
        │
        ├─→ exportToExcel()               → downloads a .xlsx file
        ├─→ saveComparisonToDrive()        → POST to Apps Script → Drive JSON file
        └─→ shareComparison()              → lz-string-compressed URL
```

---

## 5. Engineering Principles

Given this is a single static file with no framework, some conventional "clean architecture" rules don't apply literally — but the underlying intent still does:

- **Business logic must stay independent of any specific chart/export library's quirks.** Aggregation and normalization functions (`getAggregatedForType`, `normalizeUnitType`, the Suggest Unit Mix pipeline) operate on plain JS objects and arrays — they don't reach into Chart.js or ExcelJS internals. Chart/export code consumes their output, never the reverse.
- **External integrations go through small, named wrapper functions**, not ad hoc `fetch()` calls scattered through the codebase — e.g. `geocodeViaCensusProxy()` / `geocodeViaNominatim()` are separate, swappable functions behind the single `geocodeAddress()` entry point; `saveComparisonToDrive()` / `loadComparisonsFromDrive()` / `deleteComparisonFromDrive()` are the only functions that talk to the Apps Script backend.
- **Prefer composition over cleverness.** E.g. the dark-mode fix history in this repo (see Decision Log) repeatedly favored adding a clearly-named new CSS variable over reusing an existing one for two different purposes — reuse without a clear single responsibility caused real, hard-to-spot bugs (navy-on-navy text, mismatched row backgrounds).
- **Avoid premature optimization.** This app re-renders full tables/charts from scratch on most state changes rather than doing incremental DOM diffing — deliberately simple, and fine at the data volumes this tool handles (a handful of comps, dozens of floorplans).
- **No automated test suite exists.** Correctness is currently verified by hand-constructed test data run through the live app in a browser (see [Section 6](#6-ai-development-workflow) for the expected verification workflow) and checked against manually-computed expected values. This is a known gap, not a deliberate choice — see [Section 11 — Backlog](#11-backlog).
- **Write readable code over clever code.** Functions favor explicit, sequential steps with comments explaining *why* a step exists (especially in `suggestUnitMix()` and the Excel OOXML builders), not terse one-liners.
- **Document non-obvious assumptions inline**, especially around data quirks (e.g. Yardi's blank-most-recent-month fallback logic in `parseYardiWorkbook()`, or the CSS-transition timing gotcha noted in dark-mode testing).
- **Never introduce a breaking change to the saved-comparison data shape without a migration path** — comparisons already saved in the shared Drive folder must keep loading correctly after any change to how `applySubject()`/`loadComparison()` structure data.

---

## 6. AI Development Workflow

**Before writing code:**
1. Read this file (`PROJECT.md`) in full.
2. Skim recent `git log` — this repo has no separate CHANGELOG or issue tracker; commit messages *are* the project history and are written to be detailed (multi-paragraph, explaining root cause / fix / verification). Recent commits are the most reliable source of "what was just changed and why."
3. Locate the relevant function(s) in `index.html` via `Grep`, not by assuming file layout from memory — this file is large (~3,200 lines) and grows.
4. If the change touches the Apps Script backend, **ask the user for its current source** — it is not in this repo (see [Section 4.5](#45-backend-google-apps-script--drive)).

**Before changing code:**
- State your understanding of the current behavior and the requested change in plain language before editing.
- Call out assumptions explicitly, especially about data shapes (e.g. "I'm assuming `p.unitTypes` always has `sqft > 0` when `units > 0`") — this codebase has several defensive filters (`u.units > 0 && u.sqft > 0 && u.marketRent > 0`) precisely because that assumption doesn't always hold in real uploaded data.
- For anything beyond a small, obviously-scoped fix, briefly state the plan and tradeoffs before implementing, especially for anything touching `suggestUnitMix()`, the Excel OOXML builders, or the color/theme system — all three have a history of subtle regressions from well-intentioned changes.
- Wait for explicit user go-ahead before large refactors or before adding new external dependencies/services.

**While implementing:**
- Preserve the existing code style: vanilla JS, `function` declarations for top-level logic, arrow functions for callbacks, CSS custom properties (`var(--x)`) for all themeable colors — never hardcode a color that should adapt to dark mode.
- After any CSS/JS edit, verify brace/paren balance before testing (this repo has no linter or bundler to catch a stray unmatched brace before runtime) — a quick PowerShell character-count check has been the established method in this project's history:
  ```powershell
  $t = Get-Content -Raw "yardi-compare\index.html"
  ($t.ToCharArray()|?{$_-eq'{'}).Count; ($t.ToCharArray()|?{$_-eq'}'}).Count
  ($t.ToCharArray()|?{$_-eq'('}).Count; ($t.ToCharArray()|?{$_-eq')'}).Count
  ```
- Start the local dev server (`.claude/launch.json` → `yardi-compare` config, `http://localhost:3000`) and manually verify the change in a browser before considering it done. The password gate requires `localStorage.setItem('bannerAuthOk','true')` before the page is interactive in an automated test session.
- **Known testing gotcha:** CSS transitions can cause `getComputedStyle()` to report a mid-transition value if checked synchronously right after a state change (e.g. right after toggling dark mode). Either wait ~300–400ms, or temporarily inject `* { transition: none !important; }` before asserting on computed styles.
- Always reset any test data injected into `properties`/`subjectProperty` (or reload the page) before finishing, so the live app isn't left in a test state.

**After completing work:**
- Update this file (`PROJECT.md`) if architecture, data flow, or the backend contract changed.
- Update the Decision Log (Section 8) if a non-obvious tradeoff was made.
- Write a detailed, multi-paragraph commit message in the established style of this repo's history: what the problem was, what the fix does, and how it was verified — future readers (human or AI) rely on commit messages as the primary changelog.
- Only commit/push when the user explicitly asks — do not assume approval to commit carries forward from a previous request.
- Summarize what changed and what (if anything) is left to do.

---

## 7. Repository Conventions

**Naming conventions**
- JS functions: `camelCase`, verb-first (`renderDashboard`, `suggestUnitMix`, `geocodeViaCensusProxy`).
- CSS custom properties: `--kebab-case` (`--navy-text`, `--row-accent`, `--input-bg`).
- Constants: `UPPER_SNAKE_CASE` for true constants (`DRIVE_SYNC_URL`, `UNITMIX_MIN_SHARE`), `camelCase` for mutable module-level state (`properties`, `subjectProperty`, `charts`).

**Folder conventions**
- One feature = functions grouped together in `index.html` under a `// ── Section Name ──` comment banner (e.g. `// ── Comp Map ──`, `// ── Dark Mode ──`, `// ── Saved Comparisons ──`). When adding a new feature area, follow this same banner-comment convention rather than scattering related functions across the file.

**Branch naming**
- This repo has worked directly on `main` throughout its history (no observed feature-branch workflow). If a branching workflow is introduced, document it here and update this section.

**Commit message style**
- Multi-paragraph, present-tense summary line under ~70 characters, followed by a blank line and a detailed body covering: what was wrong (if a fix), what changed, and how it was verified. Every commit in this repo's history since AI-assisted development began ends with `Co-Authored-By: Claude Sonnet 4.6 <noreply@anthropic.com>` — continue that pattern for AI-assisted commits.

**Testing strategy**
- No automated tests exist. Verification is manual: inject synthetic test data via the browser console against the running local dev server, assert on computed values/styles via `javascript_exec`-style tooling, then reset state before finishing. See [Section 6](#6-ai-development-workflow).

**Code formatting**
- No formatter/linter is configured. Match existing style exactly (2-space indent in the `<style>` block and JS, dense single-line CSS rules, minimal whitespace inside object literals).

**Documentation standards**
- Comments explain *why*, not *what* — this file's own Engineering Principles apply to the codebase itself. Avoid restating what a line obviously does.

**Logging standards**
- Minimal — a handful of `console.error()` calls exist for Drive sync failures. There is no structured logging or telemetry. User-facing feedback goes through the `notify(message, 'success'|'error')` toast function, not `console.log`.

**Error handling philosophy**
- Fail soft, tell the user. Network calls (Drive sync, geocoding) are wrapped in `try/catch` and degrade gracefully (e.g. geocoding falls back from Census → Nominatim; Drive load failure falls back to local-only saved comparisons) rather than throwing uncaught errors. User-facing errors go through `notify(..., 'error')`.

**Configuration management**
- None — there is no config file or environment-variable system. The few configurable values (`DRIVE_SYNC_URL`, `SITE_PASSWORD`, tuning constants like `UNITMIX_MIN_SHARE`) are top-level `const`s directly in `index.html`.

**Secrets management**
- <a name="secrets-management"></a>**Important caveat:** `DRIVE_SYNC_TOKEN` and `SITE_PASSWORD` are plain-text constants embedded in `index.html`, which is served as a public static file. Anyone who views the page source can read both values. This is a deliberate, accepted tradeoff for an internal tool with no sensitive financial/PII data at stake beyond property addresses — but it means **this is not real security**, only a basic deterrent against casual/accidental access. Do not treat this pattern as a template for anything handling genuinely sensitive data, and do not "fix" it by inventing a fake client-side security layer — if stronger access control is ever needed, that requires an actual authenticated backend, which is out of scope for the current architecture.

---

## 8. Decision Log

| Date (approx.) | Decision | Alternatives considered | Why this won | Tradeoffs |
|---|---|---|---|---|
| Early | Single static HTML file, no framework/bundler | React/Vue SPA with a build step; small backend (Node/Flask) | Zero infrastructure to host or maintain; GitHub Pages deploys it for free with a `git push`; matched the actual complexity needed | Harder to modularize as the file grows (~3,200 lines in one file); no type safety; no automated testing harness |
| Mid | Native OOXML chart-building layered on ExcelJS, instead of pasting chart images into the export | Export static chart images via `canvas.toDataURL()` (simpler); switch to a different Excel library with native chart support | Recipients can edit/re-theme the exported charts in Excel itself — a genuine product requirement (client/lender deliverable), not a nice-to-have | Significantly more complex export code (hand-rolled XML generation) that's harder to modify safely |
| Mid | Google Apps Script + Drive folder as the "backend," instead of a real server/database | Firebase/Supabase; a small hosted backend (would cost money, need hosting/auth); staying local-storage-only (no team sharing) | Free, already-available Google Workspace infrastructure; "good enough" for internal team file sharing; no new infra to provision or pay for | Backend logic lives outside this git repo (no version control on it from here); Apps Script has execution quotas; CORS limitations required a proxy pattern for the Census geocoder |
| Mid | Client-side password gate instead of real auth | No gate at all (relying on GitHub Pages URL obscurity); real per-user auth (OAuth) | Cheap, "good enough" deterrent for an internal tool the team already trusts each other on; real auth was overkill for the actual risk | Not real security — trivially bypassable by viewing page source (documented explicitly in Section 7) |
| Recent | Census Bureau geocoder (via Apps Script proxy) as primary, Nominatim as fallback, instead of Nominatim alone | Nominatim/OpenStreetMap only (simpler, no proxy needed); Google Maps Geocoding API (requires billing account + API key) | Nominatim's accuracy for apartment buildings not individually mapped in OSM was visibly poor when zoomed in; Census is free, no API key, and street-number accurate for U.S. addresses; Google Maps would've required setting up billing for a team that wanted to avoid that | Added a dependency on the Apps Script backend being kept in sync (CORS forced routing through it); Census only covers U.S. addresses (Nominatim fallback covers the rest) |
| Recent | Suggest Unit Mix: SF picked per bedroom type must be within ±5% of that type's own average SF, not just the best $/SF | Pure best-$/SF ranking (original implementation) | Best-$/SF alone systematically favored small "outlier" units (small-unit $/SF premium), producing unrealistic SF suggestions that didn't reflect the type's actual typical size | Slightly more complex selection logic; requires computing a second average (matching the Tables tab's own average-SF method, for consistency) |
| Recent | Dark mode: split `--navy` into `--navy` (background role) and a new `--navy-text` (heading/label role) rather than overriding `--navy` directly | Override `--navy` itself for dark mode | `--navy` was already reused for two conflicting roles (dark background AND navy heading text on light cards); overriding it directly would have broken every dark-navy background (header, buttons, table headers) | Required a global find/replace of every `color: var(--navy)` usage to the new variable — a larger diff than a naive fix, but avoided a much worse hidden regression |
| Recent | Dark mode row backgrounds: introduced `--row-accent`, separate from `--warm-gray` (the page background), for in-card subtle highlights | Reuse `--warm-gray` for both roles | `--warm-gray` needed to go *darker* than the page's cards (to keep cards "popping" against the page, in both themes) — but reused inside a table row, that same darkening direction made the "highlight" column render *darker* than its own row instead of lighter, looking like a jarring hole rather than a highlight | Same "split the overloaded variable" pattern as `--navy-text` above — a recurring lesson in this codebase's theming work |
| Recent | Property-name table cell: removed its own standalone background entirely (previously `var(--row-accent)`) | Give it a background that matches whatever the row needs (subject / totals / plain) | A per-cell background always paints over a row-level (`<tr>`) background in browsers — so as long as the name cell has *any* explicit background of its own, it will always create a seam against sky-pale (COMP AVERAGE) or transparent (plain rows) row backgrounds. Removing it entirely lets whatever the row's actual background is show through uniformly. | None significant — the font-weight/font-family styling that mattered was independent of the background and was preserved |

---

## 9. Project Roadmap

This project does not follow formal phase gates (no PM/roadmap tooling is in use) — but retroactively, and for planning future work, it maps onto:

**Phase 0 — Discovery (complete)**
- Goal: prove a Yardi-export-to-Excel-comp-package pipeline was feasible.
- Deliverable: `rent_comp_builder.py`, a local CLI script.
- Exit criteria: could produce a usable Excel output from real Yardi files. *(Met — see Section 4.10.)*

**Phase 1 — MVP web app (complete)**
- Goal: move the pipeline into a shareable, no-install web app.
- Deliverables: `index.html` with Yardi/CoStar upload, subject-property entry, comparison tables, dashboard charts, native Excel export.
- Exit criteria: a real analyst could use it end-to-end for a real deal without the original developer's help. *(Met.)*

**Phase 2 — Team collaboration (complete)**
- Goal: make comparisons shareable across the team, not stuck on one browser.
- Deliverables: Google Drive sync via Apps Script, share links, JSON import/export, password gate.
- Exit criteria: saved comparisons visible to teammates other than the creator. *(Met.)*

**Phase 3 — Decision support (complete)**
- Goal: turn the tool from "just shows you the data" into "helps you decide."
- Deliverables: Suggest Pricing, Suggest Unit Mix, Rent Growth trends, interactive Map.
- Exit criteria: both suggestion features ship with a written, defensible rationale a user can sanity-check (not a black box). *(Met.)*

**Phase 4 — Ownership transition & hardening (in progress)**
- Goal: this tool survives the original developer's internship ending.
- Deliverables: GitHub repo, Drive folder, and Apps Script project all transferred to company-controlled accounts (see [Section 3 — Open Questions](#3-current-development-stage)); this `PROJECT.md`, written specifically so a future AI agent (or new team member) can pick the project up cold.
- Exit criteria: someone other than the original developer can make a change to this project, end to end, using only what's in this repo plus a normal AI coding assistant.

**Phase 5 — Scaling (not started / speculative)**
- Candidate goals: automated tests, a real backend if per-user auth or heavier data volume is ever needed, direct Yardi API integration.
- No committed deliverables — see [Backlog](#11-backlog) "Could Have" / "Won't Have Yet" for candidates.

---

## 10. Known Risks

| Risk | Category | Mitigation |
|---|---|---|
| No automated tests — regressions rely on manual browser verification | Technical | Every change in this project's history has been manually verified against synthetic test data before shipping; this is documented as the required workflow in Section 6, but it is inherently less reliable than a real test suite. |
| Apps Script backend lives outside version control (not in this repo) | Technical | Document its contract thoroughly here (Section 4.5); ask the user to paste current source before making backend changes; consider mirroring it into this repo as a `.gs` file if it changes further. |
| Shared-secret token and password are plain-text in a public file | Security | Accepted risk for an internal tool with low-sensitivity data (see Section 7); do not extend this pattern to anything more sensitive. |
| Client-side password gate is trivially bypassable (view source) | Security | Same as above — explicitly documented as "not real security," not a bug to silently "fix" with more client-side theater. |
| Single point of failure: one Google account currently owns the Drive folder + Apps Script project | Business / Continuity | Ownership transition to a company-controlled account is the active Phase 4 goal (Section 9) — confirm this has actually completed before assuming continuity. |
| No formal backup of saved comparisons beyond Google Drive's own versioning/trash | Technical | Google Drive's built-in trash/version history is the only safety net today; no separate backup process exists. |
| Nominatim (OpenStreetMap) fallback usage — free tier has a usage policy (max ~1 req/sec) | Technical / Compliance | Already respected in code (`geocodeAddress()` only throttles when it actually falls back to Nominatim); don't remove that throttle if touching this code. |
| Single developer/AI-pair-programmed codebase, ~3,200 lines in one file, no modularization | Technical | This document exists specifically to offset that risk for future contributors; keep it updated as the file grows. |

---

## 11. Backlog

**Must Have**
- Keep this `PROJECT.md` current as features are added or architecture changes.
- Confirm/complete the ownership transfer of GitHub repo, Drive folder, and Apps Script project to company-controlled accounts (see Section 3).

**Should Have**
- Some form of automated regression testing, even lightweight (e.g. a script that loads known-good synthetic data and asserts on key computed values), given the growing complexity of `suggestUnitMix()` and the Excel export.
- Mirror the Apps Script backend source into this repo (e.g. as a documented `.gs` file, even if not deployed from here) so it's version-controlled and reviewable like everything else.

**Could Have**
- Per-user authentication instead of a single shared password, if the team grows or access needs to be individually revocable.
- Splitting `index.html` into multiple files, if it grows meaningfully beyond its current size and a build step becomes worth the added complexity it currently avoids.
- Historical trend tracking across saved comparisons (e.g., "how has this submarket moved over time").

**Won't Have Yet**
- Full underwriting/proforma modeling (debt, IRR, waterfalls) — explicitly out of scope; that's a different tool (`multifamily-proforma`, a separate, unrelated project).
- A real multi-tenant backend/database — not justified by current scale (single internal team).
- Direct Yardi API integration — not evaluated for feasibility yet; manual export remains the input method.

---

## 12. Definition of Done

Before considering any change to this repository complete:

- [ ] The change works when manually verified in a browser against the local dev server (not just "looks right in the diff").
- [ ] Brace/paren balance verified after any JS/CSS edit (no linter exists to catch this automatically — see Section 6).
- [ ] No new hardcoded colors introduced outside the CSS custom-property system (breaks dark mode otherwise).
- [ ] If the change affects saved-comparison data shape, existing saved comparisons in Drive still load correctly.
- [ ] If the change affects the Apps Script contract, the exact required backend change was communicated clearly to the user (since it can't be edited from this repo).
- [ ] Test data injected into the running app for verification has been cleared / the page reloaded before finishing.
- [ ] This `PROJECT.md` updated if architecture, data flow, or the backend contract changed.
- [ ] Commit message follows the established detailed style (Section 7) — problem, fix, verification — if the user asked for a commit.
- [ ] No known critical bug introduced (verified manually, since there's no automated suite to lean on).

---

## 13. AI Guardrails

Future AI agents working in this repository must **never**:
- Rewrite `index.html` wholesale or restructure it into multiple files/a build system without being explicitly asked — this is a deliberate architectural choice (see Decision Log), not an oversight to "fix."
- Introduce a bundler, framework, or npm dependency tree unless explicitly requested — the zero-build-step property is a feature of this project.
- Assume the Apps Script backend's current source without being given it — it is not in this repo.
- "Fix" the password gate or shared-secret token by inventing a fake stronger client-side security layer — that would create a false sense of security without solving the actual problem (which requires a real backend, out of scope today).
- Hardcode a color instead of using the existing CSS custom-property theme system — this has caused real, subtle dark-mode bugs in this project's history (see Decision Log) and is easy to get visually-plausible-but-wrong.
- Touch `rent_comp_builder.py` as if it's part of the live product — it's a legacy artifact, not connected to `index.html`.
- Commit or push without being explicitly asked to in that turn.
- Guess at unstated requirements for ambiguous product decisions (e.g., exact color choices, exact threshold constants like `UNITMIX_MIN_SHARE`) — ask, or clearly flag the assumption made and why.

Instead, they should:
- Ask clarifying questions when a request is ambiguous, especially around anything touching money-adjacent logic (pricing suggestions, unit mix suggestions) where a wrong silent assumption could mislead a real deal decision.
- Prefer small, incremental, reviewable changes over large rewrites.
- Document assumptions explicitly, in the response to the user and, where relevant, as code comments.
- Preserve backward compatibility with existing saved-comparison data whenever possible.
- Verify changes by hand in the browser (Section 6) before declaring them done — this repo has no other safety net.

---

## 14. Repository Context

**Important files**
- `yardi-compare/index.html` — the entire application. Start here for essentially any task.
- `PROJECT.md` — this file.
- `rent_comp_builder.py` — legacy, not live (Section 4.10).
- `.claude/launch.json` — local dev server definition.

**Important directories**
- `yardi-compare/` — contains only `index.html`. No subfolders, no assets folder (all icons are inline SVG; all libraries load from CDN).

**Configuration files**
- None in the traditional sense. `.claude/launch.json` is the closest thing (dev server config, not app config).

**CI/CD**
- None. Deployment is implicit: GitHub Pages serves whatever is on `main` directly. There is no build, test, or lint step in the deploy path.

**Scripts**
- None (no `package.json`, no npm scripts). `rent_comp_builder.py` is a standalone script but is legacy/unused by the live site.

**Developer setup**
1. Clone `github.com/manderson-source/bannerwebsite`.
2. No install step needed — it's a static HTML file.
3. To preview locally, use the `.claude/launch.json` `yardi-compare` config (starts a PowerShell static file server on port 3000), or simply open `yardi-compare/index.html` directly in a browser (most features work; Drive sync and geocoding still hit real network endpoints either way since there's no local mock).

**Local development**
- `http://localhost:3000` via the launch config above.
- The password gate will block interaction until `localStorage.setItem('bannerAuthOk', 'true')` is set (matching the real `SITE_PASSWORD` check in the deployed app) — see `submitAuthGate()`.

**Testing commands**
- None automated. Manual verification via browser console / DevTools, as described in Section 6.

**Deployment commands**
```bash
git add <changed files>
git commit -m "..."
git push origin main
```
GitHub Pages picks up the change automatically within roughly 1–2 minutes. There is no separate deploy command or pipeline.

---

## 15. Future AI Handoff

### Instructions for Future AI Agents

- Treat this `PROJECT.md` as the repository's source of truth. If it and the code disagree, **say so explicitly** rather than silently trusting one over the other — then update whichever is stale once you've confirmed which is correct.
- Never assume unstated requirements, especially anything touching the two "judgment call" features (Suggest Pricing, Suggest Unit Mix) — small threshold/weighting changes there can materially change what a real analyst sees as a recommendation for a real deal.
- Prefer incremental, reviewable changes. This is a single ~3,200-line file maintained by AI-pair-programming sessions one feature request at a time — large unrequested refactors make future diffs (and future AI sessions) harder to reason about, not easier.
- Explain your reasoning *before* implementing anything non-trivial. The established pattern in this project's history is: understand → verify assumptions → (for larger changes) propose a plan → implement → verify manually in-browser → commit with a detailed message.
- Keep this document synchronized with the code. If you add a feature, a tab, an external dependency, or change the data model, update the relevant section here in the same turn if the user is asking for the work to be "done," not just implemented.
- If uncertain — about a product decision, a threshold value, the current state of the (out-of-repo) Apps Script backend, or anything else — **ask the user** rather than inventing plausible-sounding behavior. This codebase's history shows that plausible-but-wrong assumptions (e.g., about which CSS variable was safe to reuse for a second purpose) have caused real, hard-to-spot bugs more than once.
