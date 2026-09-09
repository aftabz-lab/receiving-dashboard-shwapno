#!/usr/bin/env python3
"""
Extract full tables from a Power BI "publish to web" report.

This talks to the same public querydata endpoint the embed iframe uses,
authenticated only by the resource key that is already inside the public URL.
No credentials are required and none are stored.

Usage:
    python pbi_extract.py --discover
        Prints every table and column the report's model exposes.
        Run this first, then fill in scripts/config.json.

    python pbi_extract.py --extract
        Pulls every configured table in full (paged) to data/raw/<table>.csv

Fragility warning: the resource key rotates whenever the report is
republished. When that happens the script exits non-zero with a clear
message and the workflow will fail loudly instead of serving stale data.
"""

import argparse
import base64
import csv
import json
import os
import re
import sys
import time
import urllib.parse
from pathlib import Path

import requests

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "scripts" / "config.json"
RAW_DIR = ROOT / "data" / "raw"

PAGE_SIZE = 30000          # max rows Power BI returns per query window
MAX_PAGES = 400            # safety stop: 12m rows
TIMEOUT = 120


# --------------------------------------------------------------------------
# Session discovery
# --------------------------------------------------------------------------

def parse_embed_url(url: str) -> dict:
    """Decode the ?r= token into resource key / tenant / cluster number."""
    qs = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)
    if "r" not in qs:
        sys.exit("ERROR: embed URL has no ?r= token.")
    raw = qs["r"][0]
    raw += "=" * (-len(raw) % 4)
    token = json.loads(base64.b64decode(raw).decode())
    return {
        "resource_key": token["k"],
        "tenant": token.get("t"),
        "cluster_no": token.get("c"),
    }


def discover_backend(session: requests.Session, resource_key: str, embed_url: str) -> str:
    """
    Find the wabi-* backend host for this report.

    Tried in order:
      1. the documented clusterdetails service
      2. the backend URL embedded in the view page HTML
    """
    try:
        r = session.get(
            "https://api.powerbi.com/powerbi/globalservice/v201606/clusterdetails",
            headers={"X-PowerBI-ResourceKey": resource_key},
            timeout=TIMEOUT,
        )
        if r.ok:
            url = r.json().get("clusterUrl")
            if url:
                return url.rstrip("/")
    except requests.RequestException:
        pass

    r = session.get(embed_url, timeout=TIMEOUT)
    r.raise_for_status()
    m = re.search(r"https://wabi-[a-z0-9\-]+\.analysis\.windows\.net", r.text)
    if m:
        return m.group(0)

    sys.exit(
        "ERROR: could not discover the Power BI backend cluster.\n"
        "Open the report in a browser, check the Network tab for a request to\n"
        "wabi-<region>.analysis.windows.net, and set 'backend_url' in config.json."
    )


def fetch_model(session, backend, resource_key):
    """Get the report id, model id and the model schema (tables + columns)."""
    url = f"{backend}/public/reports/modelsAndExploration?preferReadOnlySession=true"
    r = session.get(url, headers={"X-PowerBI-ResourceKey": resource_key}, timeout=TIMEOUT)
    if r.status_code in (401, 403, 404):
        sys.exit(
            f"ERROR: report rejected the resource key (HTTP {r.status_code}).\n"
            "The report was most likely republished, which rotates the key.\n"
            "Get the current 'Publish to web' link and update PBI_EMBED_URL."
        )
    r.raise_for_status()
    payload = r.json()

    models = payload.get("models") or []
    if not models:
        sys.exit("ERROR: no model returned by modelsAndExploration.")
    model = models[0]

    exploration = payload.get("exploration") or {}
    report_id = exploration.get("report", {}).get("objectId") or payload.get("reportId", "")

    return {
        "model_id": model.get("id"),
        "db_name": model.get("dbName"),
        "report_id": report_id,
        "schema": model.get("model", {}),
    }


def list_schema(model_info):
    """Flatten the model schema into {table: [columns]}."""
    out = {}
    for tbl in model_info["schema"].get("tables", []):
        name = tbl.get("name")
        cols = [c.get("name") for c in tbl.get("columns", []) if not c.get("isHidden")]
        measures = [m.get("name") for m in tbl.get("measures", [])]
        out[name] = {"columns": cols, "measures": measures}
    return out


# --------------------------------------------------------------------------
# Query construction
# --------------------------------------------------------------------------

def build_query(entity, columns, measures, model_id, report_id, restart_tokens=None):
    src = {"SourceRef": {"Source": "t"}}
    select, projections = [], []

    for i, col in enumerate(columns):
        select.append({"Column": {"Expression": src, "Property": col}, "Name": f"t.{col}"})
        projections.append(i)

    for j, mea in enumerate(measures):
        select.append({"Measure": {"Expression": src, "Property": mea}, "Name": f"t.{mea}"})
        projections.append(len(columns) + j)

    window = {"Count": PAGE_SIZE}
    if restart_tokens:
        window["RestartTokens"] = restart_tokens

    command = {
        "SemanticQueryDataShapeCommand": {
            "Query": {
                "Version": 2,
                "From": [{"Name": "t", "Entity": entity, "Type": 0}],
                "Select": select,
            },
            "Binding": {
                "Primary": {"Groupings": [{"Projections": projections}]},
                "DataReduction": {"DataVolume": 4, "Primary": {"Window": window}},
                "Version": 1,
            },
            "ExecutionMetricsKind": 1,
        }
    }

    return {
        "version": "1.0.0",
        "queries": [
            {
                "Query": {"Commands": [command]},
                "QueryId": "",
                "ApplicationContext": {
                    "DatasetId": model_id,
                    "Sources": [{"ReportId": report_id, "VisualId": ""}],
                },
            }
        ],
        "cancelQueries": [],
        "modelId": model_id,
    }


# --------------------------------------------------------------------------
# DSR decoding
#
# Power BI returns a compressed row set: values are de-duplicated into
# dictionaries, repeated cells are elided via a bitmask (R) and nulls via
# another bitmask (Ohm). This rebuilds real rows from that.
# --------------------------------------------------------------------------

def decode_dsr(result, n_fields):
    ds = result["results"][0]["result"]["data"]["dsr"]["DS"][0]
    value_dicts = ds.get("ValueDicts", {})
    descriptors = ds.get("S", [])

    dict_names = []
    for d in descriptors:
        dict_names.append(d.get("DN"))
    while len(dict_names) < n_fields:
        dict_names.append(None)

    rows, previous = [], [None] * n_fields

    for container in ds.get("PH", []):
        for item in container.get("DM0", []):
            compressed = item.get("C", [])
            repeat_mask = item.get("R", 0)
            null_mask = item.get("\u00d8", 0)

            row, cursor = [], 0
            for i in range(n_fields):
                if null_mask >> i & 1:
                    value = None
                elif repeat_mask >> i & 1:
                    value = previous[i]
                else:
                    value = compressed[cursor] if cursor < len(compressed) else None
                    cursor += 1
                    dn = dict_names[i]
                    if dn and isinstance(value, int):
                        table = value_dicts.get(dn, [])
                        if 0 <= value < len(table):
                            value = table[value]
                row.append(value)
            previous = row
            rows.append(row)

    restart = ds.get("RT")
    return rows, restart


# --------------------------------------------------------------------------
# Extraction
# --------------------------------------------------------------------------

def extract_table(session, backend, resource_key, model_info, spec):
    entity = spec["table"]
    columns = spec.get("columns", [])
    measures = spec.get("measures", [])
    n_fields = len(columns) + len(measures)

    url = f"{backend}/public/reports/querydata?synchronous=true"
    headers = {
        "X-PowerBI-ResourceKey": resource_key,
        "Content-Type": "application/json;charset=UTF-8",
        "Accept": "application/json, text/plain, */*",
    }

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    out_path = RAW_DIR / f"{entity.replace(' ', '_')}.csv"

    restart, total, page = None, 0, 0
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        writer.writerow(columns + measures)

        while page < MAX_PAGES:
            body = build_query(
                entity, columns, measures,
                model_info["model_id"], model_info["report_id"],
                restart_tokens=restart,
            )
            resp = session.post(url, headers=headers, json=body, timeout=TIMEOUT)
            if resp.status_code == 429:
                time.sleep(10)
                continue
            if not resp.ok:
                sys.exit(f"ERROR: querydata failed for '{entity}' "
                         f"(HTTP {resp.status_code}): {resp.text[:400]}")

            rows, restart = decode_dsr(resp.json(), n_fields)
            for row in rows:
                writer.writerow(row)
            total += len(rows)
            page += 1
            print(f"  {entity}: page {page}, +{len(rows)} rows, {total} total", flush=True)

            if not restart or not rows:
                break

    print(f"  -> {out_path} ({total} rows)")
    return total


# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--discover", action="store_true", help="print model tables and columns")
    ap.add_argument("--extract", action="store_true", help="pull configured tables in full")
    args = ap.parse_args()

    config = json.loads(CONFIG_PATH.read_text()) if CONFIG_PATH.exists() else {}
    embed_url = os.environ.get("PBI_EMBED_URL") or config.get("embed_url")
    if not embed_url:
        sys.exit("ERROR: set PBI_EMBED_URL or 'embed_url' in scripts/config.json")

    token = parse_embed_url(embed_url)
    session = requests.Session()
    session.headers["User-Agent"] = "Mozilla/5.0 (snapshot-builder)"

    backend = config.get("backend_url") or discover_backend(session, token["resource_key"], embed_url)
    print(f"backend: {backend}")

    model_info = fetch_model(session, backend, token["resource_key"])
    print(f"model:   {model_info['model_id']}  report: {model_info['report_id']}")

    if args.discover or not args.extract:
        schema = list_schema(model_info)
        print("\n--- tables and columns ---")
        print(json.dumps(schema, indent=2, ensure_ascii=False))
        (ROOT / "data").mkdir(exist_ok=True)
        (ROOT / "data" / "model_schema.json").write_text(
            json.dumps(schema, indent=2, ensure_ascii=False), encoding="utf-8"
        )
        print("\nSaved to data/model_schema.json. Copy the tables and columns you\n"
              "need into the 'tables' list in scripts/config.json, then run --extract.")
        return

    specs = config.get("tables", [])
    if not specs:
        sys.exit("ERROR: config.json has no 'tables' configured. Run --discover first.")

    grand = 0
    for spec in specs:
        print(f"extracting {spec['table']} ...")
        grand += extract_table(session, backend, token["resource_key"], model_info, spec)
    print(f"done: {grand} rows across {len(specs)} tables")


if __name__ == "__main__":
    main()
