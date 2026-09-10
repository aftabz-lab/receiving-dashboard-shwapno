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

## The Under measures in your model

Confirmed from the report's Under Receiving page: the model publishes
**Under Receiving** and **Under Receiving Score**, but no **Under Receiving
Value** and no **Under Receiving Score Icon**. The dashboard now matches that
exactly:

- The Under tables show `Under Receiving` and `Under Receiving Score`, with no
  empty value or status columns, laid out like the published page.
- `Under Receiving` is shown to two decimals, because the source returns a
  negative fractional quantity (-1,051.82).
- The Under tables sort by `Under Receiving Score` descending, as the published
  page does. The Over tables still sort by `Over Receiving Value`.
- The pic 3 card is titled **Under-receiving units** and shows the Under
  Receiving quantity, since there is no monetary measure to show. The division
  column reads **Under units** to match.

Everything is resolved at runtime, so if `Under Receiving Value` is ever added
to the model and republished, the card retitles itself to "Under-receiving
value", switches to taka formatting and the value column reappears on the Under
tables, with no further changes here.
