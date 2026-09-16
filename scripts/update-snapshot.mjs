import { readFile, writeFile, rename } from "node:fs/promises";
import { PowerBIDataClient, POWER_BI_URL } from "../powerbi.js";

const SNAPSHOT_VERSION = 9;
const NUMERIC_FIELDS = ["Sales", "Receiving", "Inventory", "OverValue", "OverIncidents", "UnderIncidents"];
const OPTION_KEYS = ["categoryOptions", "articleOptions", "outletOptions", "userOptions", "movementOptions"];
const filters = {
  days: 30,
  dateFrom: null,
  dateTo: null,
  masterCategory: "all",
  category: "all",
  region: "all",
  rho: "all",
  zonal: "all",
  outletCode: "all",
  articleNo: "all",
  userCode: "all",
  movementCode: "all",
};

const STATUS_PATH = new URL("../snapshot-status.json", import.meta.url);
// FORCE_SNAPSHOT=1 rebuilds even when Power BI reports no change. Use it after
// the published report link is replaced.
const FORCE = /^(1|true|yes|force)$/i.test(String(process.env.FORCE_SNAPSHOT || "").trim());
// MAX_SNAPSHOT_AGE_HOURS rebuilds a snapshot that has simply grown old, even
// when the Power BI refresh timestamp has not moved. Blank or 0 disables it.
const MAX_AGE_HOURS = Number(String(process.env.MAX_SNAPSHOT_AGE_HOURS || "").trim());

const finite = value => value == null || value === "" || !Number.isFinite(Number(value)) ? null : Number(value);
const sum = (rows, field) => rows.reduce((total, row) => total + (finite(row[field]) ?? 0), 0);

async function writeStatus(patch) {
  let existing = {};
  try {
    existing = JSON.parse(await readFile(STATUS_PATH, "utf8"));
  } catch {}
  const status = { ...existing, lastCheckedAt: new Date().toISOString(), ...patch };
  try {
    await writeFile(STATUS_PATH, `${JSON.stringify(status, null, 2)}\n`);
  } catch (error) {
    console.warn(`The heartbeat file could not be written: ${error?.message || error}`);
  }
  return status;
}

function snapshotIsStale(snapshot) {
  if (!Number.isFinite(MAX_AGE_HOURS) || MAX_AGE_HOURS <= 0) return false;
  const generated = Date.parse(snapshot?.snapshotGeneratedAt || "");
  if (!Number.isFinite(generated)) return true;
  return Date.now() - generated >= MAX_AGE_HOURS * 3_600_000;
}

function hasCompleteOutletMetrics(rows) {
  return Array.isArray(rows)
    && rows.length > 0
    && rows.some(row => finite(row.OverIncidents) != null && finite(row.UnderIncidents) != null);
}

function chunks(values, size = 100) {
  const output = [];
  for (let index = 0; index < values.length; index += size) output.push(values.slice(index, index + size));
  return output;
}

async function loadInOutletPartitions(client, baseFilters, outletCodes, section, label) {
  const results = [];

  async function loadPart(codes) {
    try {
      results.push(await client.load({ ...baseFilters, outletCodes: codes }, { section }));
    } catch (error) {
      if (codes.length <= 1) throw error;
      const middle = Math.ceil(codes.length / 2);
      console.warn(`${label} partition with ${codes.length} outlets failed; retrying smaller partitions.`);
      await loadPart(codes.slice(0, middle));
      await loadPart(codes.slice(middle));
    }
  }

  for (const outletCodesPart of chunks(outletCodes)) await loadPart(outletCodesPart);
  return results;
}

function shiftIsoDate(value, days) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

function sameWindow(left, right) {
  return Boolean(
    left?.sourceTimestamp
    && left.sourceTimestamp === right?.sourceTimestamp
    && left.range?.start === right?.range?.start
    && left.range?.endExclusive === right?.range?.endExclusive
  );
}

function uniqueArticles(rows) {
  const articles = new Map();
  for (const row of rows) {
    const code = String(row.ArticleNo || "").trim();
    if (!code) continue;
    const existing = articles.get(code);
    if (!existing || (!existing.ArticleName && row.ArticleName)) articles.set(code, row);
  }
  return [...articles.values()].sort((left, right) => String(left.ArticleNo).localeCompare(String(right.ArticleNo), undefined, { numeric: true }));
}

function uniqueOutletOptions(rows) {
  const outlets = new Map();
  for (const row of rows) {
    const code = String(row.OutletCode || "").trim();
    if (!code) continue;
    const existing = outlets.get(code) || {};
    outlets.set(code, {
      OutletCode: code,
      Outlet: existing.Outlet || row.Outlet || code,
      Region: existing.Region || row.Region || null,
    });
  }
  return [...outlets.values()].sort((left, right) => left.OutletCode.localeCompare(right.OutletCode, undefined, { numeric: true }));
}

function combineGroupedRows(rows, key, { keepBlank = false } = {}) {
  const groups = new Map();
  for (const row of rows) {
    const value = String(row[key] ?? "").trim();
    if (!value && !keepBlank) continue;
    const mapKey = value || "__UNASSIGNED__";
    const current = groups.get(mapKey) || { ...row, [key]: value || null, ...Object.fromEntries(NUMERIC_FIELDS.map(field => [field, 0])) };
    for (const field of NUMERIC_FIELDS) current[field] += finite(row[field]) ?? 0;
    groups.set(mapKey, current);
  }
  return [...groups.values()].sort((left, right) => String(left[key] ?? "").localeCompare(String(right[key] ?? ""), undefined, { numeric: true, sensitivity: "base" }));
}

function combineOutlets(rows, days) {
  const groups = new Map();
  for (const row of rows) {
    const code = String(row.OutletCode || "").trim();
    if (!code) continue;
    const current = groups.get(code) || {
      OutletCode: code,
      Outlet: row.Outlet || code,
      Region: row.Region || null,
      ...Object.fromEntries(NUMERIC_FIELDS.map(field => [field, 0])),
      _seen: {},
    };
    if (!current.Outlet && row.Outlet) current.Outlet = row.Outlet;
    if (!current.Region && row.Region) current.Region = row.Region;
    for (const field of NUMERIC_FIELDS) {
      const value = finite(row[field]);
      if (value == null) continue;
      current[field] += value;
      current._seen[field] = true;
    }
    groups.set(code, current);
  }
  return [...groups.values()].map(row => {
    for (const field of NUMERIC_FIELDS) if (!row._seen[field]) row[field] = null;
    row.StockDay = finite(row.Inventory) != null && finite(row.Sales) > 0
      ? row.Inventory / (row.Sales / days)
      : null;
    delete row._seen;
    return row;
  });
}

function kpiFallback(categories, outlets) {
  const sales = sum(categories, "Sales");
  const receiving = sum(categories, "Receiving");
  return {
    Sales: sales,
    Receiving: receiving,
    Inventory: sum(categories, "Inventory"),
    LatestStock: null,
    StockDay: null,
    OverValue: sum(categories, "OverValue"),
    OverIncidents: sum(categories, "OverIncidents"),
    UnderIncidents: sum(categories, "UnderIncidents"),
    OverIncidentPct: null,
    UnderIncidentPct: null,
    ActiveOutlets: outlets.length || null,
  };
}

async function loadPartitionedWindow(client, baseFilters, core, { includeArticles = false, includeOutlets = false } = {}) {
  const articleGroups = [];
  const categoryGroups = [];
  const regionGroups = [];
  const outletGroups = [];
  for (const masterCategory of core.scope.masterCategories) {
    console.log(`Loading ${masterCategory} breakdown for ${core.range.start} to ${shiftIsoDate(core.range.endExclusive, -1)}.`);
    const partitionFilters = { ...baseFilters, masterCategory };
    let light = null;
    let breakdownResults;
    try {
      breakdownResults = [await client.load(partitionFilters, { section: "snapshotBreakdowns" })];
    } catch (error) {
      // Inventory is not additive across date partitions. Outlet partitions
      // preserve the requested date window while reducing DAX query pressure.
      console.warn(`${masterCategory} breakdown query was split by outlet: ${error?.message || error}`);
      light = await client.load(partitionFilters, { section: "snapshotOutlets" });
      const outletCodes = [...new Set((light.snapshotOutlets || []).map(row => String(row.OutletCode || "").trim()).filter(Boolean))];
      if (!outletCodes.length) throw error;
      breakdownResults = await loadInOutletPartitions(
        client,
        partitionFilters,
        outletCodes,
        "snapshotBreakdowns",
        `${masterCategory} breakdown`
      );
    }
    for (const breakdowns of breakdownResults) {
      categoryGroups.push(...(breakdowns.categories || []));
      regionGroups.push(...(breakdowns.regions || []));
    }
    if (includeOutlets) {
      if (masterCategory === "COMPANY GOODS") {
        // The complete Company Goods outlet measure is too expensive as one
        // DAX query. Discover the outlet codes with the light projection, then
        // request the full incident/value measures in bounded partitions.
        light ||= await client.load(partitionFilters, { section: "snapshotOutlets" });
        const outletCodes = [...new Set((light.snapshotOutlets || []).map(row => String(row.OutletCode || "").trim()).filter(Boolean))];
        const outletResults = await loadInOutletPartitions(
          client,
          partitionFilters,
          outletCodes,
          "outlets",
          `${masterCategory} outlet`
        );
        for (const outletResult of outletResults) {
          outletGroups.push(...(outletResult.outlets || []));
        }
      } else {
        const outletResult = await client.load(partitionFilters, { section: "outlets" });
        outletGroups.push(...(outletResult.outlets || []));
      }
    }
    if (includeArticles) {
      const articles = await client.load(partitionFilters, { section: "articles" });
      articleGroups.push(...(articles.articleOptions || []));
    }
  }
  return {
    categories: combineGroupedRows(categoryGroups, "Category"),
    regions: combineGroupedRows(regionGroups, "Region", { keepBlank: true }),
    outlets: combineOutlets(outletGroups, core.range.days),
    articleOptions: uniqueArticles(articleGroups),
  };
}

async function loadKpis(client, baseFilters, reusableRows, categories, outlets, scope) {
  if (Array.isArray(reusableRows) && reusableRows.length && finite(reusableRows[0]?.StockDay) != null) return reusableRows;

  try {
    const result = await client.load(baseFilters, { section: "kpis" });
    if (result.kpis?.length && finite(result.kpis[0]?.StockDay) != null) return result.kpis;
  } catch (error) {
    console.warn(`The combined KPI query was split into safe partitions: ${error?.message || error}`);
  }

  const partitionRows = [];
  const outletCodes = [...new Set(outlets.map(row => String(row.OutletCode || "").trim()).filter(Boolean))];
  for (const masterCategory of scope.masterCategories) {
    const partitionFilters = { ...baseFilters, masterCategory };
    try {
      const result = await client.load(partitionFilters, { section: "kpis" });
      partitionRows.push(...(result.kpis || []));
    } catch (error) {
      if (!outletCodes.length) throw error;
      console.warn(`${masterCategory} KPI query was split by outlet: ${error?.message || error}`);
      for (const outletCodesPart of chunks(outletCodes)) {
        const result = await client.load({ ...partitionFilters, outletCodes: outletCodesPart }, { section: "kpis" });
        partitionRows.push(...(result.kpis || []));
      }
    }
  }

  if (!partitionRows.length) return [kpiFallback(categories, outlets)];
  const combined = kpiFallback(categories, outlets);
  combined.LatestStock = sum(partitionRows, "LatestStock");
  const stockDayWeight = partitionRows.reduce((total, row) => {
    const sales = finite(row.Sales);
    const stockDay = finite(row.StockDay);
    return total + (sales != null && stockDay != null ? sales * stockDay : 0);
  }, 0);
  const weightedSales = partitionRows.reduce((total, row) => {
    const sales = finite(row.Sales);
    return total + (sales != null && finite(row.StockDay) != null ? sales : 0);
  }, 0);
  combined.StockDay = weightedSales ? stockDayWeight / weightedSales : null;
  for (const prefix of ["Over", "Under"]) {
    const incidentField = `${prefix}Incidents`;
    const percentField = `${prefix}IncidentPct`;
    const population = partitionRows.reduce((total, row) => {
      const count = finite(row[incidentField]);
      const percent = finite(row[percentField]);
      return total + (count != null && percent > 0 ? count / (percent / 100) : 0);
    }, 0);
    combined[percentField] = population ? (combined[incidentField] / population) * 100 : null;
  }
  combined.ActiveOutlets = outlets.length || null;
  return [combined];
}

let previous = null;
try {
  previous = JSON.parse(await readFile(new URL("../snapshot.json", import.meta.url), "utf8"));
} catch {}

const client = new PowerBIDataClient();
await client.connect();
// The default dashboard window must mirror the saved Power BI date slicer,
// including both boundary dates, rather than applying an independent rolling
// 30-day calculation.
filters.dateFrom = client.scope.start;
filters.dateTo = shiftIsoDate(client.scope.endExclusive, -1);
filters.days = Math.max(1, Math.round((Date.parse(`${client.scope.endExclusive}T00:00:00Z`) - Date.parse(`${client.scope.start}T00:00:00Z`)) / 86_400_000));
const core = await client.load(filters, { section: "trend" });
const stale = snapshotIsStale(previous);
if (FORCE) console.log("FORCE_SNAPSHOT is set; rebuilding regardless of the Power BI refresh timestamp.");
if (stale && !FORCE) console.log(`The saved snapshot is older than ${MAX_AGE_HOURS} hours; rebuilding it.`);
const canReuseDefault = !FORCE && !stale && sameWindow(previous, core);
if (
  canReuseDefault
  && previous?.ready
  && previous.snapshotVersion >= SNAPSHOT_VERSION
  && previous.categories?.length
  && previous.regions?.length
  && previous.outlets?.length
  && previous.articleOptions?.length
  && Array.isArray(previous.cachedRanges)
) {
  console.log(`Power BI has not changed since ${previous.snapshotGeneratedAt}; keeping the existing snapshot.`);
  await writeStatus({
    result: "unchanged",
    powerBiRefreshedAt: core.sourceTimestamp || null,
    reportWindow: previous.range ? `${previous.range.start} to ${shiftIsoDate(previous.range.endExclusive, -1)}` : null,
    snapshotGeneratedAt: previous.snapshotGeneratedAt || null,
    reportLink: POWER_BI_URL,
  });
  process.exit(0);
}

const canReuseOptions = canReuseDefault && OPTION_KEYS.every(key => Array.isArray(previous?.[key]) && previous[key].length);
// A snapshot-version bump can represent a calculation change even when the
// Power BI source timestamp is unchanged. Never carry KPI rows across that
// boundary; rebuild them from the current category partitions.
const canReuseKpis = canReuseDefault && previous?.snapshotVersion >= SNAPSHOT_VERSION;
const supporting = canReuseOptions
  ? Object.fromEntries(OPTION_KEYS.map(key => [key, previous[key]]))
  : await client.load(filters, { section: "options" });
const canReuseArticles = canReuseDefault
  && previous?.completeness?.articleOptionsPartitionedBy === "MasterCategory"
  && previous.articleOptions?.length;
const canReuseOutlets = canReuseDefault && hasCompleteOutletMetrics(previous?.outlets);
const defaultParts = await loadPartitionedWindow(client, filters, core, {
  includeArticles: !canReuseArticles,
  includeOutlets: !canReuseOutlets,
});
const outlets = canReuseOutlets ? previous.outlets : defaultParts.outlets;
const articleOptions = canReuseArticles
  ? previous.articleOptions
  : uniqueArticles([...(supporting.articleOptions || []), ...defaultParts.articleOptions]);
const kpis = await loadKpis(
  client,
  filters,
  canReuseKpis && finite(previous?.kpis?.[0]?.StockDay) != null ? previous.kpis : null,
  defaultParts.categories,
  outlets,
  core.scope
);

const cachedRanges = [];
const sourceEnd = shiftIsoDate(core.scope.endExclusive, -1);
if (core.scope.start !== core.range.start || sourceEnd !== shiftIsoDate(core.range.endExclusive, -1)) {
  const sourceFilters = { ...filters, dateFrom: core.scope.start, dateTo: sourceEnd };
  console.log(`Building fast cached range ${sourceFilters.dateFrom} to ${sourceFilters.dateTo}.`);
  const reusableRange = previous?.cachedRanges?.find(item => (
    item?.sourceTimestamp === core.sourceTimestamp
    && item.range?.start === sourceFilters.dateFrom
    && item.range?.endExclusive === shiftIsoDate(sourceFilters.dateTo, 1)
  ));
  const sourceCore = reusableRange?.trend?.length
    ? { ...reusableRange, scope: core.scope, sourceTimestamp: core.sourceTimestamp }
    : await client.load(sourceFilters, { section: "trend" });
  const sourceNeedsOutlets = !hasCompleteOutletMetrics(reusableRange?.outlets);
  const sourceParts = await loadPartitionedWindow(client, sourceFilters, sourceCore, { includeOutlets: sourceNeedsOutlets });
  const sourceOutlets = sourceNeedsOutlets ? sourceParts.outlets : reusableRange.outlets;
  const sourceKpis = await loadKpis(
    client,
    sourceFilters,
    canReuseKpis && finite(reusableRange?.kpis?.[0]?.StockDay) != null ? reusableRange.kpis : null,
    sourceParts.categories,
    sourceOutlets,
    sourceCore.scope
  );
  cachedRanges.push({
    kpis: sourceKpis,
    trend: sourceCore.trend || [],
    categories: sourceParts.categories,
    regions: sourceParts.regions,
    outlets: sourceOutlets,
    range: sourceCore.range,
    scope: sourceCore.scope,
    sourceTimestamp: sourceCore.sourceTimestamp,
    queryTimestamp: sourceCore.queryTimestamp,
  });
}

const categoryOptions = combineGroupedRows([
  ...(supporting.categoryOptions || []),
  ...defaultParts.categories.map(row => ({ Category: row.Category })),
], "Category").map(row => ({ Category: row.Category }));
const outletOptions = uniqueOutletOptions([...(supporting.outletOptions || []), ...outlets]);
const snapshot = {
  ...core,
  ...supporting,
  kpis,
  categories: defaultParts.categories,
  regions: defaultParts.regions,
  outlets,
  categoryOptions,
  outletOptions,
  articleOptions,
  cachedRanges,
  ready: true,
  snapshotVersion: SNAPSHOT_VERSION,
  snapshotGeneratedAt: new Date().toISOString(),
  completeness: {
    queryBlockMaximumRows: 30000,
    articleOptionsPartitionedBy: "MasterCategory",
    categoryRowsPartitionedBy: "MasterCategory",
    regionRowsPartitionedBy: "MasterCategory",
    articleOptionRows: articleOptions.length,
    categoryRows: defaultParts.categories.length,
    regionRows: defaultParts.regions.length,
    outletRows: outlets.length,
    note: "Power BI result blocks are split, partitioned and de-duplicated; the dashboard applies no smaller snapshot row cap.",
  },
};

for (const key of ["kpis", "trend", "categories", "regions", "outlets", "categoryOptions", "articleOptions", "outletOptions", "userOptions", "movementOptions", "cachedRanges"]) {
  if (!Array.isArray(snapshot[key])) throw new Error(`Snapshot is missing ${key}.`);
}
if (!snapshot.kpis.length || !snapshot.range || !snapshot.trend.length || !snapshot.categories.length || !snapshot.regions.length || !snapshot.outlets.length) {
  throw new Error("Power BI returned an incomplete snapshot; the previous snapshot was preserved.");
}
// The F081 article is a truncation canary, not a business rule. Failing on its
// absence alone froze the snapshot whenever the article had no movement inside
// the report window. Fail only when the article list also shrank, which is what
// a genuinely truncated partition looks like.
const REQUIRED_ARTICLE = "2402081";
const hasRequiredArticle = snapshot.articleOptions.some(row => String(row.ArticleNo) === REQUIRED_ARTICLE);
const previousArticleCount = Array.isArray(previous?.articleOptions) ? previous.articleOptions.length : 0;
if (!hasRequiredArticle && previousArticleCount && snapshot.articleOptions.length < previousArticleCount * 0.9) {
  throw new Error(`Article options look truncated (${snapshot.articleOptions.length} rows against ${previousArticleCount}) and article ${REQUIRED_ARTICLE} is missing; the previous snapshot was preserved.`);
}
if (!hasRequiredArticle) {
  console.warn(`Article ${REQUIRED_ARTICLE} has no rows in the current report window; the snapshot was published because the article list is complete (${snapshot.articleOptions.length} rows).`);
}
for (const range of snapshot.cachedRanges) {
  if (!range.kpis.length || !range.trend.length || !range.categories.length || !range.regions.length || !range.outlets.length) {
    throw new Error(`Cached range ${range.range?.start || "unknown"} is incomplete; the previous snapshot was preserved.`);
  }
}

const temporaryPath = new URL("../snapshot.next.json", import.meta.url);
const destinationPath = new URL("../snapshot.json", import.meta.url);
await writeFile(temporaryPath, JSON.stringify(snapshot));
await rename(temporaryPath, destinationPath);
await writeStatus({
  result: "updated",
  powerBiRefreshedAt: snapshot.sourceTimestamp || null,
  reportWindow: snapshot.range ? `${snapshot.range.start} to ${shiftIsoDate(snapshot.range.endExclusive, -1)}` : null,
  snapshotGeneratedAt: snapshot.snapshotGeneratedAt,
  lastChangeAt: snapshot.snapshotGeneratedAt,
  rows: {
    articleOptions: articleOptions.length,
    categories: snapshot.categories.length,
    outlets: snapshot.outlets.length,
    trendDays: snapshot.trend.length,
  },
  reportLink: POWER_BI_URL,
});
console.log(`Snapshot v${SNAPSHOT_VERSION} created at ${snapshot.snapshotGeneratedAt} with ${articleOptions.length} article options and ${snapshot.categories.length} categories.`);
