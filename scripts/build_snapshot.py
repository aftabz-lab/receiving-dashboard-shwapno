#!/usr/bin/env python3
"""
Turn data/raw/*.csv into the snapshot the dashboard actually reads.

Output layout under data/snapshot/:

    manifest.json              snapshot id, date range, shard list, dims hash
    dims.json                  outlet / category / division / movement lookups
    daily/YYYY-MM-DD.json.gz   one gzipped shard per day, pre-aggregated
    detail/YYYY-MM.csv.gz      row-level detail, fetched only on drill-down

The KPI cards never read row-level data. A 31-day range is ~30 small shards,
which the browser fetches in parallel and caches. Changing the date after
that is an in-memory slice with zero network cost.
"""

import csv
import gzip
import hashlib
import json
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONFIG = json.loads((ROOT / "scripts" / "config.json").read_text())
RAW_DIR = ROOT / "data" / "raw"
SNAP_DIR = ROOT / "data" / "snapshot"

MAP = CONFIG["snapshot"]
DATE_FORMATS = ["%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%d-%b-%Y", "%d/%m/%Y", "%m/%d/%Y"]


def parse_date(value):
    if not value:
        return None
    text = str(value).strip()
    if text.endswith("Z"):
        text = text[:-1]
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text[:len(fmt) + 8].strip(), fmt).date().isoformat()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text).date().isoformat()
    except ValueError:
        return None


def to_number(value):
    if value in (None, "", "None"):
        return 0.0
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return 0.0


class Dims:
    """Interns repeated strings into small integer ids."""

    def __init__(self):
        self.tables = defaultdict(dict)

    def id_of(self, kind, value):
        value = (value or "").strip() or "(blank)"
        table = self.tables[kind]
        if value not in table:
            table[value] = len(table)
        return table[value]

    def export(self):
        return {kind: [v for v, _ in sorted(t.items(), key=lambda kv: kv[1])]
                for kind, t in self.tables.items()}


def main():
    source = RAW_DIR / MAP["source_csv"]
    if not source.exists():
        sys.exit(f"ERROR: {source} not found. Run pbi_extract.py --extract first.")

    dims = Dims()
    buckets = defaultdict(lambda: defaultdict(float))
    counts = defaultdict(int)
    detail_rows = defaultdict(list)

    dim_cols = MAP["dimensions"]        # {"outlet": "Outlet Name", ...}
    measure_cols = MAP["measures"]      # {"receipts": "Receipt Value", ...}
    keep_detail = MAP.get("detail_columns", [])

    with source.open(encoding="utf-8", newline="") as fh:
        reader = csv.DictReader(fh)
        header = reader.fieldnames or []
        for needed in [MAP["date_column"], *dim_cols.values(), *measure_cols.values()]:
            if needed not in header:
                sys.exit(f"ERROR: column '{needed}' missing from {source.name}. "
                         f"Available: {header}")

        for row in reader:
            day = parse_date(row.get(MAP["date_column"]))
            if not day:
                continue

            key = (day, *[dims.id_of(k, row.get(c)) for k, c in dim_cols.items()])
            bucket = buckets[key]
            for name, column in measure_cols.items():
                bucket[name] += to_number(row.get(column))
            counts[key] += 1

            if keep_detail:
                detail_rows[day[:7]].append([row.get(c, "") for c in keep_detail])

    if not buckets:
        sys.exit("ERROR: no dated rows produced. Check 'date_column' in config.json.")

    # ---- write daily shards -------------------------------------------------
    SNAP_DIR.mkdir(parents=True, exist_ok=True)
    (SNAP_DIR / "daily").mkdir(exist_ok=True)
    for old in (SNAP_DIR / "daily").glob("*.json.gz"):
        old.unlink()

    dim_keys = list(dim_cols.keys())
    measure_keys = list(measure_cols.keys())
    by_day = defaultdict(list)

    for key, bucket in buckets.items():
        day, ids = key[0], list(key[1:])
        by_day[day].append(ids + [round(bucket[m], 4) for m in measure_keys] + [counts[key]])

    shards = []
    for day, rows in sorted(by_day.items()):
        payload = {"date": day, "cols": dim_keys + measure_keys + ["rows"], "data": rows}
        path = SNAP_DIR / "daily" / f"{day}.json.gz"
        with gzip.open(path, "wt", encoding="utf-8") as fh:
            json.dump(payload, fh, separators=(",", ":"))
        shards.append({"date": day, "groups": len(rows), "bytes": path.stat().st_size})

    # ---- write monthly detail ----------------------------------------------
    if keep_detail:
        (SNAP_DIR / "detail").mkdir(exist_ok=True)
        for old in (SNAP_DIR / "detail").glob("*.csv.gz"):
            old.unlink()
        for month, rows in sorted(detail_rows.items()):
            with gzip.open(SNAP_DIR / "detail" / f"{month}.csv.gz", "wt",
                           encoding="utf-8", newline="") as fh:
                writer = csv.writer(fh)
                writer.writerow(keep_detail)
                writer.writerows(rows)

    # ---- dims and manifest --------------------------------------------------
    dims_payload = dims.export()
    dims_text = json.dumps(dims_payload, separators=(",", ":"), ensure_ascii=False)
    (SNAP_DIR / "dims.json").write_text(dims_text, encoding="utf-8")

    days = sorted(by_day.keys())
    manifest = {
        "snapshot_id": datetime.utcnow().strftime("%Y%m%dT%H%M%SZ"),
        "built_at_utc": datetime.utcnow().isoformat() + "Z",
        "min_date": days[0],
        "max_date": days[-1],
        "dimensions": dim_keys,
        "measures": measure_keys,
        "dims_hash": hashlib.sha1(dims_text.encode()).hexdigest()[:12],
        "total_source_rows": sum(counts.values()),
        "shards": shards,
        "detail_months": sorted(detail_rows.keys()) if keep_detail else [],
    }
    (SNAP_DIR / "manifest.json").write_text(
        json.dumps(manifest, indent=2), encoding="utf-8")

    total_bytes = sum(s["bytes"] for s in shards)
    print(f"snapshot {manifest['snapshot_id']}: {len(shards)} days, "
          f"{manifest['total_source_rows']} source rows, "
          f"{total_bytes/1024:.0f} KB of shards")


if __name__ == "__main__":
    main()
