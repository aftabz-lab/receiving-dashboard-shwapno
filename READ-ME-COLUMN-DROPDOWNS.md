# Column dropdowns on the outlet table — receiving-dashboard-shwapno

Three files change: `index.html`, `app.js`, `styles.css`. Nothing else is
touched — no fonts, colours, number formats, KPI logic, Power BI queries,
snapshot rules or workflows. `styles.css` is appended to only; the existing
sheet is untouched above the new block.

## Upload

- Repository home → **Add file** → **Upload files**
- Drag in `index.html`, `app.js` and `styles.css`
- **Commit changes** → Commit directly to `main`
- Wait about a minute, then hard-refresh the dashboard (Ctrl+F5). The `?v=`
  numbers were bumped so browsers fetch the new files.

## What was added

A second header row on the outlet table carries one dropdown per selectable
column, sitting directly under the column name:

| Column | Dropdown |
|---|---|
| # | none — it is the row number |
| Outlet | All outlets, then every outlet in the current view |
| Division | All divisions |
| RHO | All RHOs |
| Zonal | All Zonals |
| Received → Under incidents | none — these are measures, not selectable lists |

How they behave:

- **They cascade.** Pick a division and the RHO and Zonal lists narrow to what
  exists inside it. The list you are standing in always keeps its full choice,
  so you can switch division without clearing anything first.
- **They combine.** Division + RHO + Zonal + Outlet all apply together.
- **They are instant.** Selections narrow the rows already on screen and never
  re-query Power BI, so there is no wait and no load on the report.
- **They follow everything else.** The outlet count chip, the three exception
  cards above the table, the Visible CSV export and the sort — column sort,
  Shift-click multi-sort and the Sort options panel — all work on the selected
  rows.
- **The first option clears it.** "All divisions" and so on. A dropdown in use
  is outlined and tinted in the petrol accent so an active selection is obvious.
- **Reset all clears them** along with the rest of the filters.
- The row stays docked under the sticky column headers when the table scrolls.

The global filter panel is unchanged and still does what it always did: it
changes the Power BI query itself. These dropdowns are a quick local narrowing
of the table in front of you — use whichever suits the moment.

## Verified before shipping

Loaded in a headless browser against your real `snapshot.json` with the Zone
Distribution mapping in place, and 27 checks were run: a dropdown on every named
column and none on the row-number or measure columns, each list populated,
cascading confirmed against the rendered rows, two- three- and four-way
combinations, the count chip following along, highlight on and off, clearing
restoring the full table, Reset all clearing the selections, and the sort panel
and Refresh button from the previous package still working. All 27 passed.

Two things the tests caught and fixed on the way: outlet labels were printing
the code twice ("B004 — B004 Barishal Sadar Road") because outlet names in this
data already start with the code, and Reset all was clearing the selection in
memory but leaving the old value showing in the dropdown until the next refresh.
