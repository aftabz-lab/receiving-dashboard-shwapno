# Sort options + Refresh button — receiving-dashboard-shwapno

Three files change: `index.html`, `app.js`, `styles.css`. Nothing else is
touched. No fonts, no colours, no number formats, no filters, no KPI logic, no
Power BI queries, no snapshot rules, no workflows. `app.js` loses exactly one
line (the old refresh click handler); everything else is added.

## Upload

- Repository home → **Add file** → **Upload files**
- Drag in `index.html`, `app.js` and `styles.css`
- **Commit changes** → Commit directly to `main`
- Wait about a minute for Pages to rebuild, then hard-refresh the dashboard
  (Ctrl+F5). The `?v=` numbers on `app.js` and `styles.css` were bumped, so
  browsers pull the new files instead of their cached copies.

## 1. The Refresh button

It was never dead. Clicking it re-downloads `snapshot.json` with a cache-buster,
reloads the Zone Distribution organisation data, and asks Power BI directly
whether a newer refresh exists — and if it does, it rebuilds the view.

The problem was that it told you nothing. The busy state only ran on a cold
start, so on a normal click the spinner never span, and because Power BI usually
publishes once or twice a day the numbers and the header line were identical
before and after. A working button and a dead button looked exactly the same.

Now a click:

- spins the icon and disables the button while it works
- shows **Checking Power BI · Looking for a newer shared snapshot…**
- ends with either the new data loaded, or **Already up to date · checked
  16 Sept, 17:40 GMT+6**, which reverts to the normal line after six seconds
- is ignored if you click it again while the first check is still running

So it is worth keeping: it is the only way to pull a new snapshot between the
automatic 15-minute checks, and it is what you press after the hourly Action
commits a new `snapshot.json`.

## 2. Sort options on the outlet table

A **Sort options** button now sits in the table header next to the outlet count.
It opens a small panel with three levels:

| Level | Control |
|---|---|
| Sort by | dropdown with all 11 sortable columns + ↑/↓ toggle |
| Then by | dropdown, or "Not used" + ↑/↓ toggle |
| Then by | dropdown, or "Not used" + ↑/↓ toggle |

- The column list is read from the table headers themselves, so it stays correct
  if a column is ever added or renamed.
- A level only unlocks once the level above it is in use, and the same column
  cannot be chosen twice.
- **Reset** returns to the default sort for the current exception type (over
  value, under incidents or stock days).
- The existing chip keeps reading **Sorted by stock days ↓, then division ↑** so
  the active sort is always visible at a glance.
- Closes on **Done**, on a click outside, or on Escape.

Everything that worked before still works: clicking a header sorts by it,
clicking it again flips direction, and Shift-clicking a second header adds a
level. The panel and the headers stay in sync in both directions.

The same panel can be added to the Division table or the drill-down table later
by repeating the small block in `index.html` with `region` or `detail` in place
of `outlet`; the engine behind it is already shared.

## Verified before shipping

The changed dashboard was loaded in a headless browser against your real
`snapshot.json` and 29 checks were run: panel builds with three levels and all
11 columns, dropdown changes actually reorder the rendered table, direction
toggles flip it back, duplicate columns are rejected, header clicks and
Shift-clicks still work and still sync the dropdowns, reset restores the focus
default, Done/outside-click/Escape all close, and the Refresh button shows its
busy state, its message, and its "already up to date" result. All 29 passed.
