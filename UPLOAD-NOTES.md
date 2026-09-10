# Receiving Dashboard — incident mode update

Upload these five files to the root of `aftabz-lab/receiving-dashboard-shwapno`,
replacing the existing ones. `xlsx-lite.js` is new. Nothing else in the repo
changes — `snapshot.json`, `organization.js`, `kpi-typography.css`, the
workflows and the scripts folder all stay exactly as they are.

| File | Status |
| --- | --- |
| `index.html` | replaced |
| `app.js` | replaced |
| `powerbi.js` | replaced |
| `styles.css` | replaced |
| `xlsx-lite.js` | **new** |

## What changed

**Over/Under switch drives the whole view.** Changing the "Receiving incident"
filter (or the Over/Under buttons on the exceptions tab) now re-renders the KPI
cards, the division table, the management-focus signals and every open
management table in the selected mode.

**KPI cards swap automatically.** A new "Under-receiving value" card sits beside
the existing one. In Over mode the Under value and Under incident cards are
hidden; in Under mode the Over value and Over incident cards are hidden. Seven
cards are visible at a time, so the KPI grid is now seven columns wide.

**Division table.** The "Over value" column becomes "Under value", including its
sort key, sort status line and CSV header.

**Management tables 1–4.** Titles and measure columns follow the mode, so Table 4
reads "Under Receiving By Category" with Under Receiving, Under Receiving Value,
Under Receiving Score and Under incident columns.

**Clickable incident numbers.** On the user-incident table, both the incident
count and the incident-% cells are links. Clicking one opens Table 4 filtered to
that outlet *and* that user code, so you land on the category list behind that
exact number. Because the user table spans every business division, this drill
keeps the same cross-division scope instead of falling back to the saved page's
single division.

**Per-column search.** A search box sits under every column header on the
management tables. Boxes combine (all must match), tolerate commas — typing
`1488` finds `1,488` — and Escape clears one. Typing no longer steals focus,
because the header only rebuilds when the columns themselves change.

**Two-sheet export.** "Visible CSV" on the user-incident table now produces a
real `.xlsx` workbook, because a CSV file cannot hold two sheets:

- Sheet 1 — the visible user rows, in the order and filtering you left them in.
- Sheet 2 — the category-level incident rows behind exactly those outlet/user
  pairs, each tagged with its outlet code, outlet name, RHO, Zonal and user code
  so the two sheets join cleanly.

Both sheets follow the selected mode, so Under receiving exports Under sheets.
Every other export on the dashboard is still a CSV, unchanged. The workbook is
built in the browser with no external library or CDN.

## One thing to check on the Power BI side

The published model exposes **Over Receiving Value**, but no **Under Receiving
Value** measure appears anywhere in the report or the snapshot. The dashboard
now detects this at runtime:

- If the mirrored Under measures exist, everything works symmetrically.
- If they do not, the Under-receiving value card shows `—` with the note "Under-receiving
  value is not published in the source model", the division table shows `—` in
  that column, and the Under versions of Tables 1–4 simply drop those columns and
  sort by Under incidents instead. Nothing errors out.

If you add `Under Receiving Value` (and optionally `Under Receiving`, `Under
Receiving Score`, `Under Receiving Score Icon`) to the model and republish, the
dashboard picks them up on the next load with no further changes here.
