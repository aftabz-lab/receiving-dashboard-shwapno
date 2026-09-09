import { readFile, writeFile, rename } from "node:fs/promises";
import { PowerBIDataClient } from "../powerbi.js";

const filters = {
  days: 30,
  masterCategory: "all",
  category: "all",
  region: "all",
  articleNo: "all",
};

const client = new PowerBIDataClient();
const core = await client.load(filters, { section: "core" });
let previous = null;
try {
  previous = JSON.parse(await readFile(new URL("../snapshot.json", import.meta.url), "utf8"));
} catch {}
if (previous?.ready && previous.sourceTimestamp === core.sourceTimestamp && previous.range?.endExclusive === core.range?.endExclusive) {
  console.log(`Power BI has not changed since ${previous.snapshotGeneratedAt}; keeping the existing snapshot.`);
  process.exit(0);
}
const supporting = await client.load(filters, { section: "options" });
const outletGroups = [];
for (const masterCategory of core.scope.masterCategories) {
  const result = await client.load({ ...filters, masterCategory }, { section: "outlets" });
  outletGroups.push(...(result.outlets || []));
}
const number = value => Number.isFinite(Number(value)) ? Number(value) : 0;
const sum = (rows, field) => rows.reduce((total, row) => total + number(row[field]), 0);
const outletMap = new Map();
for (const row of outletGroups) {
  const code = String(row.OutletCode || "").trim();
  if (!code) continue;
  const current = outletMap.get(code) || { ...row, Sales: 0, Receiving: 0, Inventory: 0, OverValue: 0, OverIncidents: 0, UnderIncidents: 0 };
  for (const field of ["Sales", "Receiving", "Inventory", "OverValue", "OverIncidents", "UnderIncidents"]) {
    current[field] = number(current[field]) + number(row[field]);
  }
  current.StockDay = current.Sales ? current.Inventory / (current.Sales / filters.days) : null;
  outletMap.set(code, current);
}
const outlets = [...outletMap.values()];
const receiving = sum(core.trend || [], "Receiving");
const sales = sum(core.trend || [], "Sales");
const inventory = sum(outlets, "Inventory");
const fallbackKpi = {
  Receiving: receiving,
  Sales: sales,
  Inventory: inventory,
  LatestStock: inventory,
  StockDay: sales ? inventory / (sales / filters.days) : null,
  OverValue: sum(outlets, "OverValue"),
  OverIncidents: sum(outlets, "OverIncidents"),
  UnderIncidents: sum(outlets, "UnderIncidents"),
  OverIncidentPct: null,
  UnderIncidentPct: null,
  ActiveOutlets: outlets.filter(row => number(row.Inventory) > 0).length,
};
const snapshot = {
  ...core,
  ...supporting,
  outlets,
  kpis: core.kpis?.length ? core.kpis : [fallbackKpi],
  ready: true,
  snapshotVersion: 1,
  snapshotGeneratedAt: new Date().toISOString(),
};

for (const key of ["kpis", "trend", "categories", "regions", "outlets", "categoryOptions", "articleOptions", "outletOptions"]) {
  if (!Array.isArray(snapshot[key])) throw new Error(`Snapshot is missing ${key}.`);
}
if (!snapshot.kpis.length || !snapshot.range || !snapshot.trend.length) throw new Error("Power BI returned an empty snapshot; the previous snapshot was preserved.");

const temporaryPath = new URL("../snapshot.next.json", import.meta.url);
const destinationPath = new URL("../snapshot.json", import.meta.url);
await writeFile(temporaryPath, JSON.stringify(snapshot));
await rename(temporaryPath, destinationPath);
console.log(`Snapshot created at ${snapshot.snapshotGeneratedAt}.`);
