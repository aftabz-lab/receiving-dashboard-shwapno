# Hourly snapshot fix — receiving-dashboard-shwapno

Four files change. Nothing else in the repository is touched: no HTML, no CSS, no
fonts, no number formats, no filter logic, no DAX, no snapshot schema. All Power
BI query logic in `powerbi.js` below line 5 is byte-identical to what you have
now.

| File | What changed |
|---|---|
| `.github/workflows/update-snapshot.yml` | Hourly schedule, 3 retries, rebase-safe push, heartbeat, run summary |
| `.github/workflows/refresh-snapshot.yml` | Same hardening, plus Force and one-off link inputs |
| `scripts/update-snapshot.mjs` | Heartbeat file, force/stale rebuild, truncation guard fixed |
| `powerbi.js` | First 4 lines only: report key read from the link, old values as fallback |

## Upload steps

**1. The two workflow files** — edit them in place, it is the most reliable way.

- Open `https://github.com/aftabz-lab/receiving-dashboard-shwapno/blob/main/.github/workflows/update-snapshot.yml`
- Click the pencil (Edit this file)
- Select all, delete, paste the contents of `.github/workflows/update-snapshot.yml` from this package
- **Commit changes** → Commit directly to `main`
- Repeat for `.github/workflows/refresh-snapshot.yml`

**2. `powerbi.js` and `scripts/update-snapshot.mjs`**

- Repository home → **Add file** → **Upload files**
- Drag in `powerbi.js` and the `scripts` folder from this package
- **Commit changes** → Commit directly to `main`

## Verify (2 minutes)

- **Actions** tab → **Refresh shared Power BI snapshot** → **Run workflow** → Run
- The run should end green. Open it and read the **Summary** box: it prints the
  heartbeat, e.g. `"result": "unchanged"` or `"result": "updated"`.
- A new file `snapshot-status.json` appears in the repository. From now on it is
  rewritten every hour, so the commit list is the proof the hourly job is alive.
- Check **Settings → Pages** shows *Deploy from a branch → main → / (root)*.

## The two knobs

Both sit at the top of `.github/workflows/update-snapshot.yml`.

```yaml
MAX_SNAPSHOT_AGE_HOURS: "6"
COMMIT_HEARTBEAT: "true"
```

- `MAX_SNAPSHOT_AGE_HOURS` — a snapshot older than this is rebuilt even when
  Power BI reports no new refresh. Set it to `"1"` for a full rebuild every hour,
  or `"0"` to rebuild only when Power BI actually changes. The hourly *check*
  runs either way, so any Power BI change is picked up within the hour.
- `COMMIT_HEARTBEAT` — set to `"false"` if you would rather see a commit only
  when `snapshot.json` itself changes.

## When the Power BI link is replaced

Two ways, no code edit needed:

- **One run only:** Actions → **Refresh snapshot (manual)** → Run workflow →
  paste the new link into *embed_url* → Run.
- **Permanently:** Settings → Secrets and variables → Actions → **Variables** →
  New repository variable → name `PBI_EMBED_URL`, value = the new
  `https://app.powerbi.com/view?r=...` link. Every run uses it from then on.
  Add `PBI_API_ROOT` too only if the new report sits on a different Power BI
  cluster.

With no variable set, the link compiled into `powerbi.js` is used, exactly as
today.

## What each fix addresses

1. **Push race.** The old workflow ran `git push` with no rebase. Any manual
   "Add files via upload" landing during a run made the push non-fast-forward,
   the run failed and that hour's snapshot was thrown away. The push now rebases
   onto `main` and retries up to five times, keeping the fresh snapshot and every
   other file from `main`.
2. **Permanent freeze risk.** The script aborted whenever article `2402081` was
   absent from the report window, which would have preserved a stale snapshot
   indefinitely and silently. It now fails only when the article list has also
   shrunk by more than 10 percent — a truncated partition, which is what that
   check was really guarding against. Otherwise it logs a warning and publishes.
3. **Transient Power BI errors.** One 429 or 503 killed the run. Each run now
   makes three attempts with 90-second and 180-second backoff.
4. **No visibility.** A failing job looked identical to "Power BI has not
   changed". `snapshot-status.json` now records `lastCheckedAt`, `result`,
   `powerBiRefreshedAt`, `reportWindow`, row counts, and `lastFailureRun` with a
   direct link to the failed run.
5. **Dropped schedule ticks.** GitHub delays or drops cron ticks under load. The
   hourly tick at :05 is backed by a catch-up at :35; when nothing changed that
   second run finishes in seconds.
