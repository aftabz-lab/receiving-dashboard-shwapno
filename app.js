import { PowerBIDataClient, POWER_BI_URL } from "./powerbi.js?v=20260910-7";
import { loadOrganizationSnapshot, normalizeOutletCode } from "./organization.js?v=20260909-4";
import { downloadWorkbook } from "./xlsx-lite.js?v=20260910-7";

const AUTO_REFRESH_MS = 15 * 60 * 1000;
const DETAIL_CACHE_MS = 5 * 60 * 1000;
const DASHBOARD_CACHE_KEY = "receiving-dashboard-shared-snapshot-v4";
const FILTER_CACHE_NAME = "receiving-dashboard-filter-snapshots-v3";
const SHARED_SNAPSHOT_URL = "./snapshot.json";
const DETAIL_ROW_LIMIT = Number.POSITIVE_INFINITY;
const DATALIST_RENDER_LIMIT = 250;
const DEFAULT_FILTERS = Object.freeze({
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
  poNumber: "all",
  movementCode: "all",
});

const nf = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });
const compactNf = new Intl.NumberFormat("en-GB", { notation: "compact", maximumFractionDigits: 1 });
const percentNf = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
const longDateNf = new Intl.DateTimeFormat("en-US", { weekday: "long", year: "numeric", month: "long", day: "numeric", timeZone: "UTC" });

const state = {
  client: new PowerBIDataClient(),
  organization: null,
  organizationPromise: null,
  sharedSnapshot: null,
  data: null,
  filters: { ...DEFAULT_FILTERS },
  activeView: "overview",
  exceptionFocus: "over",
  incidentMode: "over",
  outletSearch: "",
  rhoSuggestions: [],
  zonalSuggestions: [],
  outletSuggestions: [],
  articleSuggestions: [],
  userSuggestions: [],
  movementSuggestions: [],
  sorts: {
    region: { key: "OverValue", direction: "desc" },
    outlet: { key: "OverValue", direction: "desc" },
    detail: { key: "OverValue", direction: "desc" },
  },
  detailRows: [],
  detailColumns: [],
  detailCache: new Map(),
  detailMetric: "OverValue",
  detailContext: null,
  detailTableNumber: 1,
  detailDrillValue: null,
  detailRowFilters: {},
  detailFollowsIncidentScope: false,
  detailPinnedDivision: null,
  detailWidened: false,
  detailRelaxMovement: false,
  detailColumnFilters: {},
  detailColumnFilterTimer: 0,
  detailHeaderSignature: "",
  detailMissingColumns: [],
  detailExportBusy: false,
  underValue: { key: "", status: "idle", kind: "value", kpi: null, categories: new Map(), regions: new Map() },
  detailSourceSearch: null,
  detailSearchTimer: 0,
  detailSearch: "",
  loadSequence: 0,
  loadController: null,
  backgroundStatusTimer: 0,
  detailSequence: 0,
  nextRefreshAt: 0,
  visibleRows: { region: [], outlet: [], detail: [] },
};

const el = id => document.getElementById(id);

const dom = {
  main: el("main-content"),
  loadingBar: el("loading-bar"),
  refreshButton: el("refresh-dashboard"),
  retryButton: el("retry-dashboard"),
  resetButton: el("reset-filters"),
  themeButton: el("theme-toggle"),
  themeLabel: el("theme-label"),
  filterToggle: el("filter-toggle"),
  filtersPanel: el("filters-panel"),
  periodFilter: el("period-filter"),
  fromDateFilter: el("from-date-filter"),
  toDateFilter: el("to-date-filter"),
  masterCategoryFilter: el("master-category-filter"),
  categoryFilter: el("category-filter"),
  regionFilter: el("region-filter"),
  rhoFilter: el("rho-filter"),
  zonalFilter: el("zonal-filter"),
  rhoOptions: el("rho-options"),
  zonalOptions: el("zonal-options"),
  outletFilter: el("outlet-filter"),
  articleFilter: el("article-filter"),
  userFilter: el("user-filter"),
  poFilter: el("po-filter"),
  movementFilter: el("movement-filter"),
  incidentFilter: el("incident-filter"),
  outletOptions: el("outlet-options"),
  articleOptions: el("article-options"),
  userOptions: el("user-options"),
  movementOptions: el("movement-options"),
  cascadeNote: el("cascade-note"),
  organizationStatus: el("organization-status"),
  activeFilterSummary: el("active-filter-summary"),
  activeFilterCount: el("active-filter-count"),
  statusDot: el("status-dot"),
  connectionStatus: el("connection-status"),
  sourceFreshness: el("source-freshness"),
  errorPanel: el("error-panel"),
  errorMessage: el("error-message"),
  overValueCard: el("kpi-card-over-value"),
  underValueCard: el("kpi-card-under-value"),
  overIncidentCard: el("kpi-card-over-incidents"),
  underIncidentCard: el("kpi-card-under-incidents"),
  overviewView: el("overview-view"),
  exceptionsView: el("exceptions-view"),
  managementSignals: el("management-signals"),
  trendChart: el("trend-chart"),
  categoryChart: el("category-chart"),
  regionTable: el("region-table-body"),
  outletTable: el("outlet-table-body"),
  exceptionSummary: el("exception-summary"),
  outletSearch: el("outlet-search"),
  detailDialog: el("detail-dialog"),
  detailLoading: el("detail-loading"),
  detailError: el("detail-error"),
  detailContent: el("detail-content"),
  detailSummary: el("detail-summary"),
  detailSearch: el("detail-search"),
  detailTabs: el("management-table-tabs"),
  detailMeasureNote: el("detail-measure-note"),
  detailTableGrid: el("detail-table-grid"),
  detailHead: el("detail-table-head"),
  detailTable: el("detail-table-body"),
};

el("open-powerbi").href = POWER_BI_URL;
el("error-powerbi-link").href = POWER_BI_URL;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function finite(value) {
  if (value == null || (typeof value === "string" && !value.trim())) return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function exact(value) {
  const number = finite(value);
  return number == null ? "—" : nf.format(number);
}

function compact(value) {
  const number = finite(value);
  return number == null ? "—" : compactNf.format(number);
}

function signedCompact(value) {
  const number = finite(value);
  if (number == null) return "—";
  if (number === 0) return "0";
  return `${number > 0 ? "+" : "−"}${compactNf.format(Math.abs(number))}`;
}

function signedExact(value) {
  const number = finite(value);
  if (number == null) return "—";
  if (number === 0) return "0";
  return `${number > 0 ? "+" : "−"}${nf.format(Math.abs(number))}`;
}

function percentage(value) {
  const number = finite(value);
  return number == null ? "—" : `${percentNf.format(number)}%`;
}

function bdt(value) {
  const number = finite(value);
  if (number == null) return "—";
  const absolute = Math.abs(number);
  const sign = number < 0 ? "−" : "";
  if (absolute >= 10_000_000) return `${sign}৳${(absolute / 10_000_000).toFixed(2)} Cr`;
  if (absolute >= 100_000) return `${sign}৳${(absolute / 100_000).toFixed(1)} L`;
  if (absolute >= 1_000) return `${sign}৳${(absolute / 1_000).toFixed(1)}K`;
  return `${sign}৳${nf.format(absolute)}`;
}

function bdtExact(value) {
  const number = finite(value);
  if (number == null) return "—";
  return `${number < 0 ? "−" : ""}৳${nf.format(Math.abs(number))}`;
}

function toDate(value) {
  if (value instanceof Date) return value;
  const numeric = typeof value === "string" && /^\d+$/.test(value) ? Number(value) : value;
  const date = new Date(numeric);
  return Number.isNaN(date.getTime()) ? null : date;
}

function longDate(value) {
  const date = toDate(value);
  return date ? longDateNf.format(date) : "—";
}

function endInclusive(endExclusive) {
  const date = new Date(`${endExclusive}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() - 1);
  return date;
}

function dateRangeLabel(range) {
  if (!range?.start || !range?.endExclusive) return "—";
  const start = new Date(`${range.start}T00:00:00Z`);
  const end = endInclusive(range.endExclusive);
  const sameYear = start.getUTCFullYear() === end.getUTCFullYear();
  const first = new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    ...(sameYear ? {} : { year: "numeric" }),
    timeZone: "UTC",
  }).format(start);
  const last = new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", year: "numeric", timeZone: "UTC" }).format(end);
  return `${first} – ${last}`;
}

function isoDate(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "")) ? String(value) : null;
}

function inclusiveEndIso(range) {
  const end = range?.endExclusive ? endInclusive(range.endExclusive) : null;
  return end && !Number.isNaN(end.getTime()) ? end.toISOString().slice(0, 10) : "";
}

function syncDateInputs(data) {
  if (!data?.range) return;
  const visibleFrom = state.filters.dateFrom || data.range.start || "";
  const visibleTo = state.filters.dateTo || inclusiveEndIso(data.range);
  const latestTo = inclusiveEndIso({ endExclusive: data.scope?.endExclusive || data.range.endExclusive });
  dom.fromDateFilter.value = visibleFrom;
  dom.toDateFilter.value = visibleTo;
  dom.fromDateFilter.max = latestTo;
  dom.toDateFilter.max = latestTo;
  dom.periodFilter.value = state.filters.dateFrom && state.filters.dateTo ? "custom" : String(state.filters.days);
}

function dateTick(value) {
  const date = toDate(value);
  return date ? new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", timeZone: "UTC" }).format(date) : "";
}

function dhakaDateTime(value) {
  const date = toDate(value);
  if (!date) return "time unavailable";
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
    timeZone: "Asia/Dhaka",
    timeZoneName: "short",
  }).format(date);
}

function setText(id, value, title = "") {
  const node = el(id);
  // A missing element must never stop the rest of the dashboard rendering,
  // for instance when a cached page is paired with a newer script.
  if (!node) return;
  node.textContent = value;
  if (title) node.title = title;
  else node.removeAttribute("title");
}

function uniqueSorted(values) {
  return [...new Set(values.map(value => String(value ?? "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
}

function filtersAreDefault() {
  return Object.entries(DEFAULT_FILTERS).every(([key, value]) => state.filters[key] === value);
}

function saveDashboardCache(data, { force = false } = {}) {
  if (!force && !filtersAreDefault()) return;
  try {
    localStorage.setItem(DASHBOARD_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), data }));
  } catch {}
}

function restoreDashboardCache() {
  try {
    const cached = JSON.parse(localStorage.getItem(DASHBOARD_CACHE_KEY) || "null");
    if (!cached?.data || !cached.savedAt) return false;
    state.sharedSnapshot = cached.data;
    state.data = cached.data;
    renderAll(state.data);
    dom.statusDot.className = "status-dot is-loading";
    dom.connectionStatus.textContent = "Last saved snapshot";
    dom.sourceFreshness.textContent = `Displayed instantly · checking for an update…`;
    return true;
  } catch {
    return false;
  }
}

async function fetchSharedSnapshot() {
  const response = await fetch(SHARED_SNAPSHOT_URL, { cache: "no-store" });
  if (!response.ok) throw new Error(`Shared snapshot could not be loaded (${response.status}).`);
  const data = await response.json();
  if (!data?.ready || !Array.isArray(data.kpis) || !data.range) {
    throw new Error("The first shared snapshot has not been created yet.");
  }
  return data;
}

function cacheHash(value) {
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash ^= value.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function filteredSnapshotRequest(filters, sourceTimestamp) {
  const key = cacheHash(JSON.stringify({ filters, sourceTimestamp: sourceTimestamp || "unknown" }));
  return new Request(new URL(`./cached-filter-${key}.json`, location.href).href);
}

async function readFilteredSnapshot(filters, sourceTimestamp) {
  if (!("caches" in window)) return null;
  try {
    const cache = await caches.open(FILTER_CACHE_NAME);
    const response = await cache.match(filteredSnapshotRequest(filters, sourceTimestamp));
    if (!response) return null;
    const record = await response.json();
    return record?.sourceTimestamp === sourceTimestamp && record?.data?.range ? record.data : null;
  } catch {
    return null;
  }
}

async function writeFilteredSnapshot(filters, sourceTimestamp, data) {
  if (!("caches" in window) || !data?.range) return;
  try {
    const cache = await caches.open(FILTER_CACHE_NAME);
    const body = JSON.stringify({ sourceTimestamp, savedAt: Date.now(), data });
    await cache.put(filteredSnapshotRequest(filters, sourceTimestamp), new Response(body, { headers: { "Content-Type": "application/json" } }));
  } catch (error) {
    console.warn("Filtered snapshot cache could not be updated", error);
  }
}

function snapshotMatchesCurrentFilters(snapshot) {
  if (!snapshot?.range) return false;
  const nonDateKeys = ["masterCategory", "category", "region", "rho", "zonal", "outletCode", "articleNo", "userCode", "poNumber", "movementCode"];
  if (nonDateKeys.some(key => state.filters[key] !== DEFAULT_FILTERS[key])) return false;
  if (state.filters.dateFrom && state.filters.dateTo) {
    return state.filters.dateFrom === snapshot.range.start && state.filters.dateTo === inclusiveEndIso(snapshot.range);
  }
  return Number(state.filters.days) === Number(snapshot.range.days);
}

function embeddedRangeSnapshot(snapshot) {
  if (!Array.isArray(snapshot?.cachedRanges)) return null;
  const nonDateKeys = ["masterCategory", "category", "region", "rho", "zonal", "outletCode", "articleNo", "userCode", "poNumber", "movementCode"];
  if (nonDateKeys.some(key => state.filters[key] !== DEFAULT_FILTERS[key])) return null;
  let start = state.filters.dateFrom;
  let end = state.filters.dateTo;
  if (!start || !end) {
    end = inclusiveEndIso({ endExclusive: snapshot.scope?.endExclusive || snapshot.range.endExclusive });
    const endExclusive = nextIsoDate(end);
    const date = new Date(`${endExclusive}T00:00:00Z`);
    date.setUTCDate(date.getUTCDate() - Number(state.filters.days || 30));
    start = date.toISOString().slice(0, 10);
  }
  const cached = snapshot.cachedRanges.find(item => item?.range?.start === start && inclusiveEndIso(item.range) === end);
  return cached ? {
    ...snapshot,
    ...cached,
    cachedRanges: snapshot.cachedRanges,
    categoryOptions: snapshot.categoryOptions || [],
    articleOptions: snapshot.articleOptions || [],
    outletOptions: snapshot.outletOptions || [],
    userOptions: snapshot.userOptions || [],
    movementOptions: snapshot.movementOptions || [],
  } : null;
}

function coreWithSnapshotOptions(core, snapshot) {
  return {
    ...core,
    // Filtered breakdowns must never fall back to the unfiltered snapshot.
    // An empty result is the correct result for the active selection.
    categories: Array.isArray(core?.categories) ? core.categories : [],
    regions: Array.isArray(core?.regions) ? core.regions : [],
    outlets: core?.outlets?.length ? core.outlets : (snapshot?.outlets || []),
    categoryOptions: snapshot?.categoryOptions || [],
    articleOptions: snapshot?.articleOptions || [],
    outletOptions: snapshot?.outletOptions || [],
    userOptions: snapshot?.userOptions || [],
    movementOptions: snapshot?.movementOptions || [],
  };
}

function addSelectOptions(select, values, allLabel, selected) {
  const normalized = uniqueSorted(values);
  select.replaceChildren(new Option(allLabel, "all"));
  normalized.forEach(value => select.add(new Option(value, value)));
  if (selected !== "all" && !normalized.includes(selected)) {
    select.add(new Option(`${selected} · no current match`, selected));
  }
  select.value = selected;
}

function fillDatalist(datalist, values, limit = DATALIST_RENDER_LIMIT) {
  const fragment = document.createDocumentFragment();
  values.slice(0, limit).forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    fragment.append(option);
  });
  datalist.replaceChildren(fragment);
}

function refreshSearchDatalist(datalist, suggestions, query = "") {
  const needle = query.trim().toLocaleLowerCase();
  const matches = needle
    ? suggestions.filter(value => value.toLocaleLowerCase().includes(needle))
    : suggestions;
  fillDatalist(datalist, matches);
}

function organizationForCode(code) {
  return state.organization?.byOutlet?.get(normalizeOutletCode(code)) || null;
}

function enrichOutlet(row) {
  const organization = organizationForCode(row.OutletCode);
  return {
    ...row,
    Division: organization?.Division || row.Region || "Unassigned / HO",
    RHO: organization?.RHO || "Not mapped",
    Zonal: organization?.Zonal || "Not mapped",
    RHOShort: organization?.RHOShort || "",
    ZonalShort: organization?.ZonalShort || "",
    Area: organization?.Area || "",
    Format: organization?.Format || "",
    Gap: (finite(row.Receiving) ?? 0) - (finite(row.Sales) ?? 0),
  };
}

function plainOutlet(row) {
  return row?.Outlet || row?.OutletName || row?.OutletCode || "Unassigned / HO";
}

function plainRegion(row) {
  return row?.Division || row?.Region || "Unassigned / HO";
}

function currentOrganizationRows(excluded = "") {
  if (!state.organization?.rows) return [];
  const availableCodes = new Set(
    (state.data?.outletOptions || []).map(row => normalizeOutletCode(row.OutletCode)).filter(Boolean)
  );
  const base = availableCodes.size
    ? state.organization.rows.filter(row => availableCodes.has(row.OutletCode))
    : state.organization.rows;

  return base.filter(row => {
    if (excluded !== "region" && state.filters.region !== "all" && row.Division !== state.filters.region) return false;
    if (excluded !== "rho" && state.filters.rho !== "all" && row.RHO !== state.filters.rho) return false;
    if (excluded !== "zonal" && state.filters.zonal !== "all" && row.Zonal !== state.filters.zonal) return false;
    if (excluded !== "outlet" && state.filters.outletCode !== "all" && row.OutletCode !== state.filters.outletCode) return false;
    return true;
  });
}

function organizationScopeCodes() {
  if (state.filters.outletCode !== "all") {
    return [normalizeOutletCode(state.filters.outletCode)];
  }
  if (!state.organization?.rows) return null;
  const active = ["region", "rho", "zonal"].some(key => state.filters[key] !== "all");
  if (!active) return null;
  return currentOrganizationRows().map(row => row.OutletCode);
}

function buildPowerBIFilters() {
  const result = {
    days: state.filters.days,
    dateFrom: state.filters.dateFrom,
    dateTo: state.filters.dateTo,
    masterCategory: state.filters.masterCategory,
    category: state.filters.category,
    articleNo: state.filters.articleNo,
    userCode: state.filters.userCode,
    poNumber: state.filters.poNumber,
    movementCode: state.filters.movementCode,
    region: "all",
  };

  if (state.organization?.rows) {
    const onlyStandardRegion = state.filters.region !== "all"
      && state.filters.region !== "DhakaGBUD"
      && state.filters.rho === "all"
      && state.filters.zonal === "all"
      && state.filters.outletCode === "all";
    if (onlyStandardRegion) result.region = state.filters.region;
    else {
      const codes = organizationScopeCodes();
      if (codes) result.outletCodes = codes;
    }
  } else if (state.filters.region !== "all") {
    result.region = state.filters.region;
  }
  return result;
}

function canonicalOutletLabel(code) {
  const organization = organizationForCode(code);
  const powerBi = (state.data?.outletOptions || []).find(row => normalizeOutletCode(row.OutletCode) === normalizeOutletCode(code));
  const name = organization?.OutletName || powerBi?.Outlet || "Outlet";
  return `${normalizeOutletCode(code)} — ${name}`;
}

function canonicalArticleLabel(articleNo) {
  const row = (state.data?.articleOptions || []).find(item => String(item.ArticleNo) === String(articleNo));
  return row ? `${row.ArticleNo} — ${row.ArticleName || "Unnamed article"}` : String(articleNo);
}

function updateCascadingOptions() {
  if (!state.data) return;
  const categoryOptions = state.data.categoryOptions || [];
  const outletOptions = state.data.outletOptions || [];
  const articleOptions = state.data.articleOptions || [];
  const userOptions = state.data.userOptions || [];
  const movementOptions = state.data.movementOptions || [];
  const masterCategories = state.data.scope?.masterCategories?.length
    ? state.data.scope.masterCategories
    : ["COMPANY GOODS", "FRESH PRODUCE", "GENERAL MERCHANDISE", "LIFESTYLE", "LOOSE COMMODITY", "PACKED COMMODITY"];
  addSelectOptions(dom.masterCategoryFilter, masterCategories, "Select all", state.filters.masterCategory);
  addSelectOptions(dom.categoryFilter, categoryOptions.map(row => row.Category), "All categories", state.filters.category);

  if (state.organization?.rows) {
    addSelectOptions(dom.regionFilter, currentOrganizationRows("region").map(row => row.Division), "All divisions", state.filters.region);
    state.rhoSuggestions = uniqueSorted(currentOrganizationRows("rho").map(row => row.RHO));
    state.zonalSuggestions = uniqueSorted(currentOrganizationRows("zonal").map(row => row.Zonal));
    fillDatalist(dom.rhoOptions, state.rhoSuggestions);
    fillDatalist(dom.zonalOptions, state.zonalSuggestions);
    dom.rhoFilter.value = state.filters.rho === "all" ? "" : state.filters.rho;
    dom.zonalFilter.value = state.filters.zonal === "all" ? "" : state.filters.zonal;
    dom.rhoFilter.disabled = false;
    dom.zonalFilter.disabled = false;

    const organizationFilterActive = ["region", "rho", "zonal"].some(key => state.filters[key] !== "all");
    if (organizationFilterActive) {
      state.outletSuggestions = currentOrganizationRows("outlet")
        .sort((a, b) => a.OutletCode.localeCompare(b.OutletCode, undefined, { numeric: true }))
        .map(row => `${row.OutletCode} — ${row.OutletName || "Unnamed outlet"}`);
    } else {
      state.outletSuggestions = outletOptions
        .filter(row => row.OutletCode)
        .sort((a, b) => normalizeOutletCode(a.OutletCode).localeCompare(normalizeOutletCode(b.OutletCode), undefined, { numeric: true }))
        .map(row => {
          const code = normalizeOutletCode(row.OutletCode);
          const organization = organizationForCode(code);
          return `${code} — ${organization?.OutletName || row.Outlet || "Unnamed outlet"}`;
        });
    }
  } else {
    addSelectOptions(dom.regionFilter, outletOptions.map(row => row.Region), "All divisions", state.filters.region);
    state.rhoSuggestions = [];
    state.zonalSuggestions = [];
    fillDatalist(dom.rhoOptions, []);
    fillDatalist(dom.zonalOptions, []);
    dom.rhoFilter.value = "";
    dom.zonalFilter.value = "";
    dom.rhoFilter.disabled = true;
    dom.zonalFilter.disabled = true;
    state.outletSuggestions = outletOptions
      .filter(row => row.OutletCode)
      .map(row => `${row.OutletCode} — ${row.Outlet || "Unnamed outlet"}`);
  }
  fillDatalist(dom.outletOptions, state.outletSuggestions);

  const seenArticles = new Set();
  state.articleSuggestions = articleOptions
    .filter(row => row.ArticleNo && !seenArticles.has(String(row.ArticleNo)) && seenArticles.add(String(row.ArticleNo)))
    .map(row => `${row.ArticleNo} — ${row.ArticleName || "Unnamed article"}`);
  fillDatalist(dom.articleOptions, state.articleSuggestions);
  state.userSuggestions = uniqueSorted(userOptions.map(row => row.UserCode));
  state.movementSuggestions = uniqueSorted(movementOptions.map(row => row.MovementCode));
  fillDatalist(dom.userOptions, state.userSuggestions);
  fillDatalist(dom.movementOptions, state.movementSuggestions);

  dom.outletFilter.value = state.filters.outletCode === "all" ? "" : canonicalOutletLabel(state.filters.outletCode);
  dom.articleFilter.value = state.filters.articleNo === "all" ? "" : canonicalArticleLabel(state.filters.articleNo);
  dom.userFilter.value = state.filters.userCode === "all" ? "" : state.filters.userCode;
  dom.poFilter.value = state.filters.poNumber === "all" ? "" : state.filters.poNumber;
  dom.movementFilter.value = state.filters.movementCode === "all" ? "" : state.filters.movementCode;
  dom.incidentFilter.value = state.incidentMode;
  updateActiveFilterSummary();
}

function activeFilterLabels() {
  const labels = [state.filters.dateFrom && state.filters.dateTo
    ? `Date: ${dateRangeLabel({ start: state.filters.dateFrom, endExclusive: nextIsoDate(state.filters.dateTo) })}`
    : `${state.filters.days} days`];
  if (state.filters.masterCategory !== "all") labels.push(state.filters.masterCategory);
  if (state.filters.category !== "all") labels.push(state.filters.category);
  if (state.filters.region !== "all") labels.push(state.filters.region);
  if (state.filters.rho !== "all") labels.push(`RHO: ${state.filters.rho}`);
  if (state.filters.zonal !== "all") labels.push(`Zonal: ${state.filters.zonal}`);
  if (state.filters.outletCode !== "all") labels.push(`Outlet: ${state.filters.outletCode}`);
  if (state.filters.articleNo !== "all") labels.push(`Article: ${state.filters.articleNo}`);
  if (state.filters.userCode !== "all") labels.push(`User: ${state.filters.userCode}`);
  if (state.filters.poNumber !== "all") labels.push(`PO: ${state.filters.poNumber}`);
  if (state.filters.movementCode !== "all") labels.push(`Movement: ${state.filters.movementCode}`);
  return labels;
}

function updateActiveFilterSummary() {
  const labels = activeFilterLabels();
  dom.activeFilterCount.textContent = String(labels.length);
  dom.activeFilterSummary.textContent = labels.length ? labels.join(" · ") : "All data";
}

function nextIsoDate(value) {
  const date = new Date(`${value}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + 1);
  return date.toISOString().slice(0, 10);
}

function setLoading(loading) {
  dom.main.setAttribute("aria-busy", String(loading));
  dom.loadingBar.classList.toggle("is-active", loading);
  dom.refreshButton.disabled = loading;
  dom.refreshButton.classList.toggle("is-refreshing", loading);
  [dom.periodFilter, dom.fromDateFilter, dom.toDateFilter, dom.masterCategoryFilter, dom.categoryFilter, dom.regionFilter, dom.rhoFilter, dom.zonalFilter, dom.outletFilter, dom.articleFilter, dom.userFilter, dom.poFilter, dom.movementFilter, dom.incidentFilter, dom.resetButton]
    .forEach(node => { node.disabled = loading; });

  if (loading) {
    dom.statusDot.className = "status-dot is-loading";
    dom.connectionStatus.textContent = state.data ? "Applying filters" : "Loading saved snapshot";
    dom.sourceFreshness.textContent = state.data ? "Keeping the current view visible…" : "Opening the latest shared view…";
  } else if (!state.organization) {
    dom.rhoFilter.disabled = true;
    dom.zonalFilter.disabled = true;
  }
}

function showError(error) {
  dom.errorPanel.hidden = false;
  dom.errorMessage.textContent = `${error?.message || "The live source did not return data."} Check the connection, then try again.`;
  dom.statusDot.className = "status-dot is-error";
  dom.connectionStatus.textContent = navigator.onLine ? "Live data unavailable" : "Device is offline";
  dom.sourceFreshness.textContent = state.data ? "Showing the last successfully loaded view" : "No dashboard data loaded";
}

function clearError() {
  dom.errorPanel.hidden = true;
}

async function refreshOrganization() {
  try {
    const organization = await loadOrganizationSnapshot();
    state.organization = organization;
    dom.organizationStatus.classList.remove("is-warning");
    dom.organizationStatus.textContent = `${organization.rows.length.toLocaleString()} outlet hierarchy mappings loaded`;
    dom.organizationStatus.title = `Separate mapping file updated ${dhakaDateTime(organization.updatedAt)}`;
    if (state.data) renderAll(state.data);
    return organization;
  } catch (error) {
    console.warn("Organization mapping unavailable", error);
    if (!state.organization) {
      dom.organizationStatus.classList.add("is-warning");
      dom.organizationStatus.textContent = "Zonal/RHO mapping unavailable · Power BI data remains active";
      dom.organizationStatus.removeAttribute("title");
    }
    return state.organization;
  }
}

function isUnderMode() {
  return state.incidentMode === "under";
}

function modePrefix() {
  return isUnderMode() ? "Under" : "Over";
}

function modeValueKey() {
  return `${modePrefix()}Value`;
}

function modeIncidentKey() {
  return `${modePrefix()}Incidents`;
}

function underValueFor(grain, name) {
  if (state.underValue.status !== "ready") return null;
  const map = state.underValue[grain];
  return map ? finite(map.get(String(name ?? ""))) : null;
}

function underIsQuantity() {
  return isUnderMode() && state.underValue.kind === "quantity";
}

// The Under side of the published model stops at a quantity, so the exposure
// column and card are labelled and formatted for whichever measure exists.
function exposureLabel({ short = false } = {}) {
  if (!isUnderMode()) return short ? "Over value" : "Over-receiving value";
  if (underIsQuantity()) return short ? "Under units" : "Under-receiving units";
  return short ? "Under value" : "Under-receiving value";
}

function exposureDisplay(value) {
  return underIsQuantity() ? decimalUnits(value) : bdt(value);
}

function decimalUnits(value) {
  const number = finite(value);
  return number == null ? "—" : number.toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function applyIncidentModeVisibility() {
  const under = isUnderMode();
  if (dom.overValueCard) dom.overValueCard.hidden = under;
  if (dom.overIncidentCard) dom.overIncidentCard.hidden = under;
  if (dom.underValueCard) dom.underValueCard.hidden = !under;
  if (dom.underIncidentCard) dom.underIncidentCard.hidden = !under;
}

function applyModeToRegionHeader() {
  const button = document.querySelector('[data-sort-table="region"][data-mode-column="value"]');
  if (!button) return;
  const previousKey = button.dataset.sortKey;
  const key = modeValueKey();
  const label = exposureLabel({ short: true });
  if (previousKey === key && button.dataset.modeLabel === label) return;
  button.dataset.modeLabel = label;
  if (state.sorts.region.key === previousKey) state.sorts.region = { ...state.sorts.region, key };
  if (state.sorts.region.secondary) {
    state.sorts.region.secondary = state.sorts.region.secondary.map(item => (item.key === previousKey ? { ...item, key } : item));
  }
  button.dataset.sortKey = key;
  button.innerHTML = `${escapeHtml(label)} <span></span>`;
}

function underValueRequestKey() {
  return cacheHash(JSON.stringify({
    filters: buildPowerBIFilters(),
    range: state.data?.range || null,
    sourceTimestamp: state.data?.sourceTimestamp || "unknown",
  }));
}

async function ensureUnderValues({ force = false } = {}) {
  if (!state.data) return;
  const key = underValueRequestKey();
  if (!force && state.underValue.key === key && state.underValue.status !== "idle") return;

  state.underValue = { key, status: "loading", kind: "value", kpi: null, categories: new Map(), regions: new Map() };
  renderKpis(state.data.kpis?.[0] || {}, state.data);

  try {
    const result = await state.client.loadUnderValueContext(buildPowerBIFilters(), {
      range: state.data.range,
      scope: state.data.scope,
    });
    if (state.underValue.key !== key) return;
    state.underValue = {
      key,
      status: result.supported ? "ready" : "missing",
      kind: result.kind === "quantity" ? "quantity" : "value",
      kpi: result.kpi,
      categories: result.categories || new Map(),
      regions: result.regions || new Map(),
    };
  } catch (error) {
    console.warn("Under-receiving value could not be loaded", error);
    if (state.underValue.key !== key) return;
    state.underValue = { key, status: "error", kind: "value", kpi: null, categories: new Map(), regions: new Map() };
  }
  if (state.data) renderAll(state.data);
}

function renderPulse(kpi, data) {
  const received = finite(kpi.Receiving) ?? 0;
  const sales = finite(kpi.Sales) ?? 0;
  const gap = received - sales;
  const gapPct = sales ? (gap / sales) * 100 : null;
  const label = el("balance-label");

  label.className = "signal-badge";
  if (gap > 0) {
    label.classList.add("positive");
    label.textContent = "Inventory build";
    setText("balance-headline", `${compact(Math.abs(gap))} more units received than sold`, `${exact(Math.abs(gap))} units`);
    setText("balance-detail", sales ? `Receipts ran ${percentage(Math.abs(gapPct))} above sales in this data window.` : "Receipts were recorded while invoiced sales were zero in this data window.");
  } else if (gap < 0) {
    label.classList.add("negative");
    label.textContent = "Inventory drawdown";
    setText("balance-headline", `${compact(Math.abs(gap))} more units sold than received`, `${exact(Math.abs(gap))} units`);
    setText("balance-detail", `Sales ran ${percentage(Math.abs(gapPct))} above receipts in this data window.`);
  } else {
    label.classList.add("balanced");
    label.textContent = "Flow balanced";
    setText("balance-headline", "Received and sold units are equal");
    setText("balance-detail", "No net receipt balance is present in the selected data window.");
  }

  const scope = activeFilterLabels().slice(1, 4);
  if (activeFilterLabels().length > 4) scope.push(`+${activeFilterLabels().length - 4} more`);
  setText("report-context", scope.join(" · ") || "All business divisions");
  setText("data-window", dateRangeLabel(data.range));
  setText("query-time", `Queried ${dhakaDateTime(data.queryTimestamp)}`);
}

function setKpi(id, value, formatter = compact) {
  setText(id, formatter(value), finite(value) == null ? "" : nf.format(value));
}

function resolvedOverValue(kpi, data) {
  const sourceValue = finite(kpi?.OverValue);
  const groupedValues = (data?.categories || []).map(row => finite(row.OverValue)).filter(value => value != null);
  const groupedValue = groupedValues.length ? groupedValues.reduce((total, value) => total + value, 0) : null;
  const useGroupedValue = groupedValue != null && groupedValue !== 0 && (sourceValue == null || sourceValue === 0);
  return { value: useGroupedValue ? groupedValue : sourceValue, grouped: useGroupedValue };
}

function resolvedUnderValue() {
  if (state.underValue.status !== "ready") return { value: null, grouped: false };
  const sourceValue = finite(state.underValue.kpi);
  const groupedValues = [...state.underValue.categories.values()].map(value => finite(value)).filter(value => value != null);
  const groupedValue = groupedValues.length ? groupedValues.reduce((total, value) => total + value, 0) : null;
  const useGroupedValue = groupedValue != null && groupedValue !== 0 && (sourceValue == null || sourceValue === 0);
  return { value: useGroupedValue ? groupedValue : sourceValue, grouped: useGroupedValue };
}

function underValueNote(kpi) {
  if (state.underValue.status === "loading") return "Reading the source measure…";
  if (state.underValue.status === "missing") return "No under-receiving measure is published in the source model";
  if (state.underValue.status === "error") return "Source value unavailable · switch back to reload";
  if (underIsQuantity()) return `Source-calculated shortfall · ${exact(kpi?.UnderIncidents)} incidents · click for detail`;
  return `${exact(kpi?.UnderIncidents)} incidents · click for detail`;
}

function renderKpis(kpi, data) {
  const received = finite(kpi.Receiving);
  const sales = finite(kpi.Sales);
  const gap = received != null && sales != null ? received - sales : null;
  const ratio = sales ? (received / sales) * 100 : null;
  const overValue = resolvedOverValue(kpi, data);
  const underValue = resolvedUnderValue();

  data.resolvedOverValue = overValue.value;
  data.overValueUsesGroupedContext = overValue.grouped;
  data.resolvedUnderValue = underValue.value;
  data.underValueUsesGroupedContext = underValue.grouped;

  setKpi("kpi-receiving", received, exact);
  setKpi("kpi-sales", sales, exact);
  setKpi("kpi-gap", gap, signedExact);
  setKpi("kpi-inventory", kpi.Inventory, exact);
  setKpi("kpi-over-value", overValue.value, bdtExact);
  setKpi("kpi-outlets", kpi.ActiveOutlets, exact);
  setKpi("kpi-over-incidents", kpi.OverIncidents, exact);
  setKpi("kpi-under-incidents", kpi.UnderIncidents, exact);
  const underValueTitle = el("kpi-card-under-value")?.querySelector("h3");
  if (underValueTitle) underValueTitle.textContent = exposureLabel();
  setKpi("kpi-under-value", underValue.value, underIsQuantity() ? decimalUnits : bdtExact);
  const overValueNode = el("kpi-over-value");
  const underValueNode = el("kpi-under-value");
  if (overValueNode) overValueNode.dataset.drillValue = overValue.value == null ? "" : String(overValue.value);
  if (underValueNode) underValueNode.dataset.drillValue = underValue.value == null ? "" : String(underValue.value);
  setText("kpi-under-value-note", underValueNote(kpi));
  applyIncidentModeVisibility();
  setText("kpi-receiving-note", `${data.range.days}-day live total · click for detail`);
  setText("kpi-sales-note", `Invoiced sales · click for detail`);
  setText("kpi-gap-note", ratio == null ? "Received minus sold" : `Receipts equal ${percentage(ratio)} of sales`);
  setText("kpi-inventory-note", "Source measure · click for detail");
  setText("kpi-over-value-note", `${exact(kpi.OverIncidents)} incidents · click for detail`);
  setText("kpi-outlets-note", "Click to open the outlet list");
  setText("kpi-over-incidents-note", `${percentage(kpi.OverIncidentPct)} · click for user detail`);
  setText("kpi-under-incidents-note", `${percentage(kpi.UnderIncidentPct)} · click for user detail`);

  const gapNode = el("kpi-gap");
  if (gapNode) {
    gapNode.classList.toggle("is-positive", (gap ?? 0) > 0);
    gapNode.classList.toggle("is-negative", (gap ?? 0) < 0);
  }

  setText("stock-days", finite(kpi.StockDay) == null ? "—" : Number(kpi.StockDay).toFixed(1));
  setText("over-rate", percentage(kpi.OverIncidentPct));
  setText("under-rate", percentage(kpi.UnderIncidentPct));
  setText("over-incidents", `${exact(kpi.OverIncidents)} incidents`);
  setText("under-incidents", `${exact(kpi.UnderIncidents)} incidents`);
  setText("latest-stock", compact(kpi.LatestStock), exact(kpi.LatestStock));
}

function niceCeiling(value) {
  if (!Number.isFinite(value) || value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const nice = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return nice * magnitude;
}

function renderTrend(rows) {
  const values = rows
    .filter(row => toDate(row.Date))
    .map(row => ({ ...row, _date: toDate(row.Date), Sales: finite(row.Sales) ?? 0, Receiving: finite(row.Receiving) ?? 0 }))
    .sort((a, b) => a._date - b._date);

  dom.trendChart.classList.remove("chart-skeleton");
  if (!values.length) {
    dom.trendChart.innerHTML = '<div class="empty-chart">No daily trend data is available for this selection.</div>';
    return;
  }

  const width = 860;
  const height = 330;
  const margin = { top: 18, right: 18, bottom: 45, left: 62 };
  const plotW = width - margin.left - margin.right;
  const plotH = height - margin.top - margin.bottom;
  const yMax = niceCeiling(Math.max(...values.flatMap(row => [row.Sales, row.Receiving])) * 1.06);
  const x = index => margin.left + (values.length === 1 ? plotW / 2 : (index / (values.length - 1)) * plotW);
  const y = value => margin.top + (1 - value / yMax) * plotH;
  const path = key => values.map((row, index) => `${index ? "L" : "M"}${x(index).toFixed(2)},${y(row[key]).toFixed(2)}`).join(" ");
  const tickCount = 4;
  const yTicks = Array.from({ length: tickCount + 1 }, (_, index) => (yMax / tickCount) * index);
  const xTickCount = Math.min(5, values.length);
  const xIndexes = [...new Set(Array.from({ length: xTickCount }, (_, index) => Math.round(index * (values.length - 1) / Math.max(1, xTickCount - 1))))];
  const grid = yTicks.map(value => `<line class="chart-gridline" x1="${margin.left}" y1="${y(value)}" x2="${width - margin.right}" y2="${y(value)}"/><text class="axis-label" x="${margin.left - 10}" y="${y(value) + 4}" text-anchor="end">${escapeHtml(compact(value))}</text>`).join("");
  const xLabels = xIndexes.map(index => `<text class="axis-label" x="${x(index)}" y="${height - 14}" text-anchor="middle">${escapeHtml(dateTick(values[index]._date))}</text>`).join("");
  const pointTargets = values.map((row, index) => `<circle cx="${x(index)}" cy="${y(row.Receiving)}" r="9" fill="transparent"><title>${escapeHtml(`${dateTick(row._date)} · Receiving ${exact(row.Receiving)}`)}</title></circle><circle cx="${x(index)}" cy="${y(row.Sales)}" r="9" fill="transparent"><title>${escapeHtml(`${dateTick(row._date)} · Sales ${exact(row.Sales)}`)}</title></circle>`).join("");

  dom.trendChart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Daily line chart comparing received units with sold units"><defs><linearGradient id="receiving-area" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="var(--petrol)" stop-opacity=".18"/><stop offset="1" stop-color="var(--petrol)" stop-opacity="0"/></linearGradient></defs>${grid}<line class="chart-axis" x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}"/>${xLabels}<path d="${path("Receiving")} L${x(values.length - 1)},${height - margin.bottom} L${x(0)},${height - margin.bottom} Z" fill="url(#receiving-area)"/><path d="${path("Receiving")}" fill="none" stroke="var(--petrol)" stroke-width="3.5" stroke-linejoin="round" stroke-linecap="round"/><path d="${path("Sales")}" fill="none" stroke="var(--muted)" stroke-width="3" stroke-dasharray="7 6" stroke-linejoin="round" stroke-linecap="round"/>${pointTargets}</svg>`;
}

function renderCategoryChart(rows) {
  const values = rows
    .filter(row => row.Category)
    .map(row => ({ ...row, Gap: (finite(row.Receiving) ?? 0) - (finite(row.Sales) ?? 0) }))
    .sort((a, b) => Math.abs(b.Gap) - Math.abs(a.Gap));
  dom.categoryChart.classList.remove("chart-skeleton");
  if (!values.length) {
    dom.categoryChart.innerHTML = '<div class="empty-chart">No category data is available for this selection.</div>';
    return;
  }

  const width = 610;
  const rowHeight = 43;
  const height = Math.max(300, values.length * rowHeight + 64);
  const chartStart = 164;
  const chartEnd = 485;
  const centre = (chartStart + chartEnd) / 2;
  const half = (chartEnd - chartStart) / 2;
  const maxGap = Math.max(1, ...values.map(row => Math.abs(row.Gap)));
  const bars = values.map((row, index) => {
    const y = 42 + index * rowHeight;
    const barWidth = (Math.abs(row.Gap) / maxGap) * (half - 5);
    const positive = row.Gap >= 0;
    const x = positive ? centre : centre - barWidth;
    const title = `${row.Category}: received ${exact(row.Receiving)}, sold ${exact(row.Sales)}, balance ${signedCompact(row.Gap)}`;
    const drillAttributes = `class="chart-link" role="button" tabindex="0" data-drill-metric="Gap" data-context-type="category" data-context-value="${escapeHtml(row.Category)}" data-context-label="${escapeHtml(row.Category)}"`;
    return `<g ${drillAttributes} aria-label="Open ${escapeHtml(row.Category)} details"><text x="8" y="${y + 5}" fill="var(--ink)" font-size="12.5" font-weight="750">${escapeHtml(row.Category)}</text><rect x="${x}" y="${y - 10}" width="${Math.max(1.5, barWidth)}" height="20" rx="4" fill="${positive ? "var(--coral)" : "var(--petrol)"}"><title>${escapeHtml(title)}</title></rect><text x="${width - 8}" y="${y + 5}" text-anchor="end" fill="${positive ? "var(--coral)" : "var(--petrol)"}" font-size="12.5" font-weight="800">${escapeHtml(signedCompact(row.Gap))}</text></g>`;
  }).join("");
  dom.categoryChart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Diverging bar chart of received units minus sold units by category"><text x="${chartStart}" y="17" fill="var(--muted)" font-size="11.5" font-weight="700">Sales-led</text><text x="${chartEnd}" y="17" text-anchor="end" fill="var(--muted)" font-size="11.5" font-weight="700">Receipt-heavy</text><line class="chart-zero" x1="${centre}" y1="27" x2="${centre}" y2="${height - 12}"/>${bars}</svg>`;
}

function positionForGap(gap) {
  if (gap > 0) return { label: "Receipt-heavy", className: "receipt-heavy" };
  if (gap < 0) return { label: "Sales-led", className: "sales-led" };
  return { label: "Balanced", className: "balanced" };
}

const sortLabels = {
  Region: "division", Outlet: "outlet", OutletCode: "outlet code", RHO: "RHO", Zonal: "Zonal", ArticleNo: "article code", ArticleName: "article name",
  MasterCategory: "business division", Category: "category", Receiving: "received", Sales: "sold", Gap: "balance",
  Inventory: "inventory", StockDay: "stock days", OverValue: "over value", OverIncidents: "over incidents",
  OverIncidentPct: "over-receiving rate", UnderIncidents: "under incidents", UnderIncidentPct: "under-receiving rate",
  Incidents: "incidents", Position: "position", LatestStock: "closing stock", ActiveOutlets: "active outlets", OutletName: "outlet",
  ArticleCombo: "article group", OpeningStock: "opening stock", TotalInventory: "total inventory", TotalSales: "total sales",
  StdStockDays: "standard stock days", CurrentStockDay: "current stock day", CurrentStockSystem: "current stock in system",
  ClosingStockReceiving: "closing stock on receiving", OverReceiving: "over receiving", OverScore: "over receiving score",
  UnderValue: "under value", UnderReceiving: "under receiving", UnderScore: "under receiving score",
  StatusIcon: "status", Sorting: "sorting", EstimatedClosingStock: "estimated closing stock", OverIncidentPct: "incident percent",
  PONumber: "PO number", MovementType: "movement type", CreatedBy: "user ID", PODate: "PO date", ReceivingDate: "receiving date",
};

function incidentManagementColumns(prefix) {
  return [
    { key: `${prefix}Receiving`, label: `${prefix} Receiving`, numeric: true, ...(prefix === "Under" ? { decimals: 2 } : {}) },
    { key: `${prefix}Value`, label: `${prefix} Receiving Value`, numeric: true },
    { key: `${prefix}Score`, label: `${prefix} Receiving Score`, numeric: true, decimals: 2 },
    { key: "StatusIcon", label: ".", status: true },
  ];
}

function operationalManagementColumns(prefix) {
  return [
    { key: "OpeningStock", label: "Opening Stock", numeric: true },
    { key: "Receiving", label: "Receiving Qty", numeric: true },
    { key: "TotalInventory", label: "Total Inventory", numeric: true },
    { key: "TotalSales", label: "Total Sales Qty", numeric: true },
    { key: "StdStockDays", label: "STD. Stock Days", numeric: true, decimals: 2 },
    { key: "CurrentStockDay", label: "Current Stock Day", numeric: true, decimals: 2 },
    { key: "CurrentStockSystem", label: "Current Stock in System", numeric: true },
    { key: "ClosingStockReceiving", label: "Closing Stock On Receiving", numeric: true },
    ...incidentManagementColumns(prefix),
  ];
}

function managementTableDefinitions(prefix) {
  return {
    1: {
      title: `Table 1 – ${prefix} Receiving Incidents By Article`,
      defaultSort: `${prefix}Value`,
      columns: [
        { key: "OutletName", label: "OutletName" },
        { key: "ArticleNo", label: "Article No" },
        { key: "ArticleName", label: "ArticleName" },
        { key: "Category", label: "Category" },
        ...operationalManagementColumns(prefix),
        { key: "Sorting", label: "sorting", numeric: true },
      ],
    },
    2: {
      title: `Table 2 – ${prefix} Receiving Incidents By Article-Group (For Loose Commodity & PnP)`,
      defaultSort: `${prefix}Value`,
      columns: [
        { key: "OutletName", label: "OutletName" },
        { key: "ArticleCombo", label: "Article Combo" },
        ...operationalManagementColumns(prefix),
        { key: "Sorting", label: "sorting", numeric: true },
      ],
    },
    3: {
      title: `Table 3 – ${prefix} Receiving Incidents By Outlet`,
      defaultSort: `${prefix}Value`,
      columns: [
        { key: "OutletName", label: "OutletName" },
        ...operationalManagementColumns(prefix),
        { key: "Sorting", label: "sorting", numeric: true },
      ],
    },
    4: {
      title: `Table 4 – ${prefix} Receiving By Category`,
      defaultSort: `${prefix}Value`,
      columns: [
        { key: "Category", label: "Category3" },
        { key: "OpeningStock", label: "Opening Stock", numeric: true },
        { key: "Receiving", label: "Receiving Qty", numeric: true },
        { key: "TotalInventory", label: "Total Inventory", numeric: true },
        { key: "TotalSales", label: "Total Sales Qty", numeric: true },
        { key: "StdStockDays", label: "STD. Stock Days", numeric: true, decimals: 2 },
        { key: "EstimatedClosingStock", label: "Est. Closing Stock", numeric: true },
        ...incidentManagementColumns(prefix),
        { key: `${prefix}Incidents`, label: `${prefix} Receiving Incidents`, numeric: true },
        { key: `${prefix}IncidentPct`, label: `${prefix} Receiving Incident%`, numeric: true, decimals: 2 },
      ],
    },
    5: {
      title: "Table 5 – PO Detail Table",
      defaultSort: "ReceivingDate",
      defaultDirection: "desc",
      columns: [
        { key: "ArticleNo", label: "ArticleNo" },
        { key: "PONumber", label: "PO Number" },
        { key: "MovementType", label: "Movement Type" },
        { key: "CreatedBy", label: "Created By (User ID)" },
        { key: "PODate", label: "PO Date", date: true },
        { key: "ReceivingDate", label: "Receiving Date", date: true },
        { key: "Receiving", label: "Receiving Qty", numeric: true },
      ],
    },
    6: {
      title: `${prefix} Receiving Incidents by User`,
      defaultSort: `${prefix}Incidents`,
      columns: [
        { key: "OutletCode", label: "Outlet Code" },
        { key: "OutletName", label: "Outlet Name" },
        { key: "RHO", label: "RHO" },
        { key: "Zonal", label: "Zonal" },
        { key: "CreatedBy", label: "User Code" },
        { key: `${prefix}Incidents`, label: `${prefix} Receiving Incidents`, numeric: true },
        { key: `${prefix}IncidentPct`, label: `${prefix} Receiving Incident%`, numeric: true, decimals: 2 },
      ],
    },
  };
}

function incidentTableType() {
  if (["UnderIncidents", "UnderIncidentPct", "UnderValue"].includes(state.detailMetric)) return "under";
  if (["OverIncidents", "OverIncidentPct"].includes(state.detailMetric)) return "over";
  return state.incidentMode;
}

function managementTableDefinition(tableNumber) {
  const number = Number(tableNumber) || 1;
  const under = incidentTableType() === "under";
  const prefix = under ? "Under" : "Over";
  const definitions = managementTableDefinitions(prefix);
  const definition = definitions[number] || definitions[1];
  // A published model may not mirror every "Over" measure on the "Under" side;
  // those columns are dropped rather than shown as empty ones.
  const missing = new Set(under ? state.detailMissingColumns : []);
  const columns = definition.columns.filter(column => !missing.has(column.key));
  // The published page sorts the Over tables by Over Receiving Value and the
  // Under tables by Under Receiving Score; fall back down that chain to
  // whichever measure this model actually exposes.
  const preferred = under
    ? [`${prefix}Score`, `${prefix}Value`, `${prefix}Incidents`, `${prefix}Receiving`]
    : [definition.defaultSort, `${prefix}Value`, `${prefix}Score`, `${prefix}Incidents`];
  const defaultSort = [...preferred, definition.defaultSort]
    .find(key => columns.some(column => column.key === key))
    || columns.find(column => column.numeric)?.key
    || columns[0].key;
  return { ...definition, columns, defaultSort };
}

function csvColumns(table) {
  if (table === "region") return [
    { key: "Region", label: "Division", value: row => plainRegion(row) },
    { key: "Receiving", label: "Received" }, { key: "Sales", label: "Sold" },
    { key: "Gap", label: "Balance" }, { key: modeValueKey(), label: exposureLabel({ short: true }) },
    { key: "Incidents", label: "Incidents" }, { key: "Position", label: "Position" },
  ];
  if (table === "outlet") return [
    { key: "Rank", label: "Rank", value: row => row.__rank },
    { key: "OutletCode", label: "Outlet code" }, { key: "Outlet", label: "Outlet name", value: row => plainOutlet(row) },
    { key: "Region", label: "Division", value: row => plainRegion(row) },
    { key: "RHO", label: "RHO" }, { key: "Zonal", label: "Zonal" }, { key: "Area", label: "Area" },
    { key: "Receiving", label: "Received" }, { key: "Sales", label: "Sold" }, { key: "Gap", label: "Balance" },
    { key: "StockDay", label: "Stock days" }, { key: "OverValue", label: "Over value" },
    { key: "OverIncidents", label: "Over incidents" }, { key: "UnderIncidents", label: "Under incidents" },
  ];
  return state.detailColumns.map(column => ({ key: column.key, label: column.label }));
}

function csvCell(value) {
  if (value == null) return "";
  let rendered = String(value);
  if (typeof value === "string" && /^[=+\-@]/.test(rendered)) rendered = `'${rendered}`;
  return `"${rendered.replaceAll('"', '""')}"`;
}

function exportVisibleCsv(table) {
  const rows = state.visibleRows[table] || [];
  if (!rows.length) return;
  const columns = csvColumns(table);
  const lines = [
    columns.map(column => csvCell(column.label)).join(","),
    ...rows.map(row => columns.map(column => csvCell(column.value ? column.value(row) : row[column.key])).join(",")),
  ];
  const blob = new Blob(["\uFEFF", lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const scope = state.filters.masterCategory === "all" ? "all-business-divisions" : state.filters.masterCategory.toLocaleLowerCase().replace(/[^a-z0-9]+/g, "-");
  link.href = url;
  link.download = `receiving-${table}-visible-${scope}-${state.data?.range?.endExclusive || "current"}.csv`;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function setExportAvailability(table, hasRows) {
  const button = document.querySelector(`[data-export-table="${table}"]`);
  if (button) button.disabled = state.detailExportBusy && table === "detail" ? true : !hasRows;
}

function exportScopeSlug() {
  return state.filters.masterCategory === "all"
    ? "all-business-divisions"
    : state.filters.masterCategory.toLocaleLowerCase().replace(/[^a-z0-9]+/g, "-");
}

function workbookCell(row, column) {
  const value = row[column.key];
  if (column.date) return longDate(value);
  if (column.status) return value == null || value === "" ? "" : String(value);
  if (column.numeric) {
    const number = finite(value);
    return number == null ? null : (column.decimals == null ? number : Number(number.toFixed(column.decimals)));
  }
  return value == null || value === "" ? "" : String(value);
}

function incidentBreakdownColumns(prefix, missing = []) {
  const skip = new Set(missing);
  return [
    { key: "OutletCode", label: "Outlet Code" },
    { key: "OutletName", label: "Outlet Name" },
    { key: "RHO", label: "RHO" },
    { key: "Zonal", label: "Zonal" },
    { key: "CreatedBy", label: "User Code" },
    { key: "Category", label: "Category3" },
    { key: "OpeningStock", label: "Opening Stock", numeric: true },
    { key: "Receiving", label: "Receiving Qty", numeric: true },
    { key: "TotalInventory", label: "Total Inventory", numeric: true },
    { key: "TotalSales", label: "Total Sales Qty", numeric: true },
    { key: "StdStockDays", label: "STD. Stock Days", numeric: true, decimals: 2 },
    { key: "EstimatedClosingStock", label: "Est. Closing Stock", numeric: true },
    ...incidentManagementColumns(prefix),
    { key: `${prefix}Incidents`, label: `${prefix} Receiving Incidents`, numeric: true },
    { key: `${prefix}IncidentPct`, label: `${prefix} Receiving Incident%`, numeric: true, decimals: 2 },
  ].filter(column => !skip.has(column.key));
}

/**
 * The user-incident table exports as a two-sheet workbook: the visible,
 * filtered and sorted user rows, plus the category-level incident rows behind
 * exactly those outlet/user pairs. A CSV cannot hold two sheets, so this is a
 * real .xlsx built in the browser.
 */
async function exportIncidentWorkbook(button) {
  const rows = state.visibleRows.detail || [];
  if (!rows.length || state.detailExportBusy) return;

  const under = incidentTableType() === "under";
  const prefix = under ? "Under" : "Over";
  const originalLabel = button?.textContent;
  state.detailExportBusy = true;
  if (button) {
    button.disabled = true;
    button.textContent = "Preparing…";
  }

  try {
    const userColumns = state.detailColumns;
    const userSheet = [
      userColumns.map(column => column.label),
      ...rows.map(row => userColumns.map(column => workbookCell(row, column))),
    ];

    const order = new Map();
    rows.forEach((row, index) => order.set(`${row.OutletCode ?? ""}\u001f${row.CreatedBy ?? ""}`, index));
    const outletCodes = [...new Set(rows.map(row => row.OutletCode).filter(Boolean))];

    let breakdownSheet;
    let breakdownName = `${prefix} incidents by category`;
    try {
      const filters = { ...currentManagementFilters(), region: "all", outletCodes };
      const result = await state.client.loadIncidentCategoryBreakdown(filters, {
        range: state.data.range,
        scope: state.data.scope,
        mode: under ? "under" : "over",
      });
      const columns = incidentBreakdownColumns(prefix, result.missing || []);
      const detailRows = normalizeManagementRows(result.rows)
        .filter(row => order.has(`${row.OutletCode ?? ""}\u001f${row.CreatedBy ?? ""}`))
        .filter(row => (finite(row[`${prefix}Incidents`]) ?? 0) > 0 || (finite(row[`${prefix}Value`]) ?? 0) !== 0)
        .sort((left, right) => {
          const rank = order.get(`${left.OutletCode ?? ""}\u001f${left.CreatedBy ?? ""}`) - order.get(`${right.OutletCode ?? ""}\u001f${right.CreatedBy ?? ""}`);
          if (rank) return rank;
          return (finite(right[`${prefix}Incidents`]) ?? 0) - (finite(left[`${prefix}Incidents`]) ?? 0);
        });
      breakdownSheet = [columns.map(column => column.label), ...detailRows.map(row => columns.map(column => workbookCell(row, column)))];
      if (detailRows.length === 0) breakdownSheet.push([`No ${prefix.toLocaleLowerCase()}-receiving category rows were returned for these outlet and user codes.`]);
    } catch (error) {
      console.warn("Category incident breakdown could not be loaded", error);
      breakdownSheet = [
        ["Category detail unavailable"],
        [error?.message || "The source did not return the category-level incident rows."],
        ["Reopen the user table and export again once the source responds."],
      ];
    }

    await downloadWorkbook(
      [
        { name: `${prefix} incidents by user`, rows: userSheet },
        { name: breakdownName, rows: breakdownSheet },
      ],
      `receiving-${prefix.toLocaleLowerCase()}-incidents-${exportScopeSlug()}-${state.data?.range?.endExclusive || "current"}.xlsx`
    );
  } catch (error) {
    console.error("Incident workbook export failed", error);
    dom.detailError.hidden = false;
    dom.detailError.textContent = `${error?.message || "The workbook could not be created."} Try the export again.`;
  } finally {
    state.detailExportBusy = false;
    if (button) {
      button.textContent = originalLabel || "Visible CSV";
      button.disabled = !(state.visibleRows.detail || []).length;
    }
  }
}

function sortRows(rows, table) {
  const sort = state.sorts[table];
  const criteria = [{ key: sort.key, direction: sort.direction }, ...(sort.secondary || [])];
  return [...rows].sort((left, right) => {
    for (const { key, direction } of criteria) {
      const a = left[key];
      const b = right[key];
      const an = finite(a);
      const bn = finite(b);
      const result = an != null && bn != null
        ? an - bn
        : String(a ?? "").localeCompare(String(b ?? ""), undefined, { numeric: true, sensitivity: "base" });
      if (result) return result * (direction === "asc" ? 1 : -1);
    }
    return 0;
  });
}

function updateSortIndicators(table) {
  const sort = state.sorts[table];
  const criteria = [{ key: sort.key, direction: sort.direction }, ...(sort.secondary || [])];
  document.querySelectorAll(`[data-sort-table="${table}"]`).forEach(button => {
    const rank = criteria.findIndex(item => item.key === button.dataset.sortKey);
    const active = rank >= 0;
    const criterion = criteria[rank];
    button.classList.toggle("is-sorted", active);
    button.classList.toggle("is-ascending", active && criterion.direction === "asc");
    button.dataset.sortRank = active ? String(rank + 1) : "";
    button.setAttribute("aria-sort", active ? (criterion.direction === "asc" ? "ascending" : "descending") : "none");
  });
  const status = el(`${table}-sort-status`);
  if (status) status.textContent = `Sorted by ${criteria.map(item => `${sortLabelFor(item.key)} ${item.direction === "asc" ? "↑" : "↓"}`).join(", then ")}`;
}

function sortLabelFor(key) {
  // The under exposure column changes name with the measure behind it.
  if (key === "UnderValue") return exposureLabel({ short: true }).toLocaleLowerCase();
  return sortLabels[key] || key;
}

function drillText(display, metric, contextType, contextValue, contextLabel, extraClass = "") {
  return `<button class="text-link ${extraClass}" type="button" data-drill-metric="${escapeHtml(metric)}" data-context-type="${escapeHtml(contextType)}" data-context-value="${escapeHtml(contextValue)}" data-context-label="${escapeHtml(contextLabel)}" title="Open details for ${escapeHtml(contextLabel)}">${escapeHtml(display)}</button>`;
}

function drillNumber(display, rawValue, metric, contextType, contextValue, contextLabel, extraClass = "") {
  if (finite(rawValue) == null) return "—";
  return `<button class="number-link ${extraClass}" type="button" data-drill-metric="${escapeHtml(metric)}" data-drill-value="${escapeHtml(rawValue)}" data-context-type="${escapeHtml(contextType)}" data-context-value="${escapeHtml(contextValue)}" data-context-label="${escapeHtml(contextLabel)}" title="Open ${escapeHtml(sortLabelFor(metric))} details">${escapeHtml(display)}</button>`;
}

function renderRegions(rows) {
  const valueKey = modeValueKey();
  const values = rows.map(row => {
    const receiving = finite(row.Receiving) ?? 0;
    const sales = finite(row.Sales) ?? 0;
    const Gap = receiving - sales;
    const Incidents = (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0);
    const Position = positionForGap(Gap).label;
    const UnderValue = underValueFor("regions", row.Region ?? "");
    return { ...row, UnderValue, Receiving: receiving, Sales: sales, Gap, Incidents, Position };
  });
  const sorted = sortRows(values, "region");
  state.visibleRows.region = sorted;
  setExportAvailability("region", sorted.length > 0);
  updateSortIndicators("region");
  if (!sorted.length) {
    dom.regionTable.innerHTML = '<tr><td colspan="7" class="empty-cell">No division data is available for this selection.</td></tr>';
    return;
  }

  dom.regionTable.innerHTML = sorted.map(row => {
    const label = plainRegion(row);
    const contextValue = row.Region == null ? "__UNASSIGNED__" : row.Region;
    const position = positionForGap(row.Gap);
    return `<tr><td>${drillText(label, "Gap", "region", contextValue, label)}</td>
      <td class="numeric">${drillNumber(compact(row.Receiving), row.Receiving, "Receiving", "region", contextValue, label)}</td>
      <td class="numeric">${drillNumber(compact(row.Sales), row.Sales, "Sales", "region", contextValue, label)}</td>
      <td class="numeric">${drillNumber(signedCompact(row.Gap), row.Gap, "Gap", "region", contextValue, label, row.Gap > 0 ? "is-positive" : row.Gap < 0 ? "is-negative" : "")}</td>
      <td class="numeric">${drillNumber(exposureDisplay(row[valueKey]), row[valueKey], valueKey, "region", contextValue, label)}</td>
      <td class="numeric">${drillNumber(exact(row.Incidents), row.Incidents, "Incidents", "region", contextValue, label)}</td>
      <td>${drillText(position.label, "Gap", "region", contextValue, label, `position-pill ${position.className}`)}</td></tr>`;
  }).join("");
}

function largestBy(rows, field) {
  return [...rows].filter(row => row && finite(row[field]) != null).sort((a, b) => finite(b[field]) - finite(a[field]))[0] || null;
}

function renderSignals(data) {
  const under = isUnderMode();
  const word = under ? "under-receiving" : "over-receiving";
  const valueKey = modeValueKey();
  const incidentKey = modeIncidentKey();
  const categories = (data.categories || []).map(row => ({ ...row, UnderValue: underValueFor("categories", row.Category) }));
  const topCategoryValue = largestBy(categories, valueKey);
  const topCategoryIncidents = largestBy(categories, incidentKey);
  const regions = data.regions.map(row => ({ ...row, TotalIncidents: (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0) }));
  const topRegion = largestBy(regions, "TotalIncidents");
  const outlets = data.enrichedOutlets.filter(row => row.OutletCode);
  const topOutlet = under ? largestBy(outlets, "UnderIncidents") : largestBy(outlets, "OverValue");
  const categorySignal = topCategoryValue
    ? { title: `${topCategoryValue.Category} is the largest ${word} exposure`, detail: `${exposureDisplay(topCategoryValue[valueKey])}${underIsQuantity() ? " units" : ""} in the category context; this Power BI measure is non-additive across rows.` }
    : topCategoryIncidents && { title: `${topCategoryIncidents.Category} carries the most ${word} incidents`, detail: `${exact(topCategoryIncidents[incidentKey])} incidents in the category context for the selected window.` };
  const signals = [
    categorySignal,
    topRegion && { title: `${plainRegion(topRegion)} has the most incidents`, detail: `${exact(topRegion.TotalIncidents)} combined over- and under-receiving incidents in the selected window.` },
    topOutlet && {
      title: `${plainOutlet(topOutlet)} leads the outlet action queue`,
      detail: under
        ? `${exact(topOutlet.UnderIncidents)} under-receiving incidents under ${topOutlet.RHO} / ${topOutlet.Zonal}.`
        : `${bdt(topOutlet.OverValue)} over-receiving value under ${topOutlet.RHO} / ${topOutlet.Zonal}.`,
      link: true,
    },
  ].filter(Boolean);

  dom.managementSignals.innerHTML = signals.length
    ? signals.map(signal => `<li><strong>${escapeHtml(signal.title)}</strong><span>${escapeHtml(signal.detail)}${signal.link ? ' <a href="#exceptions-view" data-open-exceptions>Open outlet queue →</a>' : ""}</span></li>`).join("")
    : '<li><strong>No management signals available</strong><span>Try a broader filter selection.</span></li>';
}

const focusConfig = {
  over: { field: "OverValue", title: "Top over-receiving outlets", description: "Outlets ranked by over-receiving value for the selected scope.", value: row => bdt(row.OverValue) },
  under: { field: "UnderIncidents", title: "Top under-receiving outlets", description: "Outlets ranked by under-receiving incident count for the selected scope.", value: row => `${exact(row.UnderIncidents)} incidents` },
  stock: { field: "StockDay", title: "Highest stock-cover outlets", description: "Outlets ranked by the source-calculated stock-day measure.", value: row => finite(row.StockDay) == null ? "—" : `${Number(row.StockDay).toFixed(1)} days` },
};

function filteredRankedOutlets() {
  const config = focusConfig[state.exceptionFocus];
  const query = state.outletSearch.trim().toLocaleLowerCase();
  const filtered = state.data.enrichedOutlets
    .filter(row => row.OutletCode && (finite(row[config.field]) ?? 0) > 0)
    .filter(row => !query || [plainOutlet(row), row.OutletCode, plainRegion(row), row.RHO, row.Zonal, row.Area, row.Format]
      .some(value => String(value ?? "").toLocaleLowerCase().includes(query)));
  return sortRows(filtered, "outlet");
}

function renderExceptionSummary(rows) {
  const config = focusConfig[state.exceptionFocus];
  const top = rows.slice(0, 3);
  dom.exceptionSummary.innerHTML = top.length
    ? top.map((row, index) => `<article class="exception-card"><span class="exception-rank">0${index + 1}</span><span>${escapeHtml(`${row.RHO} · ${row.Zonal}`)}</span><strong title="${escapeHtml(`${row.OutletCode} — ${plainOutlet(row)}`)}">${escapeHtml(`${row.OutletCode} — ${plainOutlet(row)}`)}</strong><button class="exception-value number-link" type="button" data-drill-metric="${escapeHtml(config.field)}" data-context-type="outlet" data-context-value="${escapeHtml(row.OutletCode)}" data-context-label="${escapeHtml(`${row.OutletCode} — ${plainOutlet(row)}`)}">${escapeHtml(config.value(row))}</button></article>`).join("")
    : '<div class="empty-chart">No outlets match the selected exception type and filters.</div>';
}

function renderOutlets() {
  if (!state.data) return;
  const config = focusConfig[state.exceptionFocus];
  const rows = filteredRankedOutlets();
  const visible = rows.slice(0, 50).map((row, index) => ({ ...row, __rank: index + 1 }));
  state.visibleRows.outlet = visible;
  setExportAvailability("outlet", visible.length > 0);
  setText("exceptions-description", config.description);
  setText("outlet-table-title", config.title);
  setText("outlet-table-note", `Showing the first ${Math.min(50, rows.length)} ranked outlets. Click any number for article details.`);
  setText("outlet-result-count", `${exact(rows.length)} matching outlets`);
  setText("exception-count", compact(rows.length));
  updateSortIndicators("outlet");
  renderExceptionSummary(rows);

  document.querySelectorAll(".focus-button").forEach(button => {
    const active = button.dataset.focus === state.exceptionFocus;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });

  if (!visible.length) {
    dom.outletTable.innerHTML = '<tr><td colspan="12" class="empty-cell">No outlets match the current search and filters.</td></tr>';
    return;
  }

  dom.outletTable.innerHTML = visible.map((row, index) => {
    const context = ["outlet", row.OutletCode, `${row.OutletCode} — ${plainOutlet(row)}`];
    const drill = (display, raw, metric, cls = "") => drillNumber(display, raw, metric, ...context, cls);
    return `<tr><td>${index + 1}</td>
      <td><span class="outlet-name">${escapeHtml(plainOutlet(row))}<small>Code ${escapeHtml(row.OutletCode)}${row.Area ? ` · ${escapeHtml(row.Area)}` : ""}</small></span></td>
      <td>${escapeHtml(plainRegion(row))}</td><td>${escapeHtml(row.RHO)}</td><td>${escapeHtml(row.Zonal)}</td>
      <td class="numeric">${drill(compact(row.Receiving), row.Receiving, "Receiving")}</td>
      <td class="numeric">${drill(compact(row.Sales), row.Sales, "Sales")}</td>
      <td class="numeric">${drill(signedCompact(row.Gap), row.Gap, "Gap", row.Gap > 0 ? "is-positive" : row.Gap < 0 ? "is-negative" : "")}</td>
      <td class="numeric">${drill(finite(row.StockDay) == null ? "—" : Number(row.StockDay).toFixed(1), row.StockDay, "StockDay")}</td>
      <td class="numeric">${drill(bdt(row.OverValue), row.OverValue, "OverValue")}</td>
      <td class="numeric">${drill(exact(row.OverIncidents), row.OverIncidents, "OverIncidents")}</td>
      <td class="numeric">${drill(exact(row.UnderIncidents), row.UnderIncidents, "UnderIncidents")}</td></tr>`;
  }).join("");
}

function renderAll(data) {
  data.enrichedOutlets = (data.outlets || []).map(enrichOutlet);
  const kpi = data.kpis?.[0] || {};
  syncDateInputs(data);
  updateCascadingOptions();
  renderPulse(kpi, data);
  renderKpis(kpi, data);
  renderTrend(data.trend || []);
  renderCategoryChart(data.categories || []);
  applyModeToRegionHeader();
  renderRegions(data.regions || []);
  renderSignals(data);
  renderOutlets();
  if (isUnderMode()) ensureUnderValues();
}

function switchView(view) {
  state.activeView = view;
  const overview = view === "overview";
  dom.overviewView.hidden = !overview;
  dom.exceptionsView.hidden = overview;
  document.querySelectorAll(".view-tab").forEach(tab => {
    const active = tab.dataset.view === view;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", String(active));
    tab.tabIndex = active ? 0 : -1;
  });
}

function metricDefinition(metric) {
  const definitions = {
    Receiving: { title: "Received units", field: "Receiving", formatter: compact },
    Sales: { title: "Sold units", field: "Sales", formatter: compact },
    Gap: { title: "Receipt balance", field: "Gap", formatter: signedCompact },
    Inventory: { title: "Inventory", field: "Inventory", formatter: compact },
    StockDay: { title: "Stock cover", field: "StockDay", formatter: value => finite(value) == null ? "—" : `${Number(value).toFixed(1)} days` },
    OverValue: { title: "Over-receiving value", field: "OverValue", formatter: bdt },
    UnderValue: { title: "Under-receiving value", field: "UnderValue", formatter: bdt },
    OverIncidents: { title: "Over-receiving incidents", field: "OverIncidents", formatter: exact },
    OverIncidentPct: { title: "Over-receiving rate", field: "OverIncidentPct", formatter: percentage },
    UnderIncidents: { title: "Under-receiving incidents", field: "UnderIncidents", formatter: exact },
    UnderIncidentPct: { title: "Under-receiving rate", field: "UnderIncidentPct", formatter: percentage },
    Incidents: { title: "Receiving incidents", field: "Incidents", formatter: exact },
    ActiveOutlets: { title: "Active outlets", field: "ActiveOutlets", formatter: exact },
    LatestStock: { title: "Closing stock", field: "LatestStock", formatter: exact },
  };
  return definitions[metric] || definitions.OverValue;
}

function detailFiltersForContext(context) {
  const filters = buildPowerBIFilters();
  if (!context) return filters;
  if (context.type === "outlet") {
    filters.region = "all";
    filters.outletCodes = [context.value];
  } else if (context.type === "region") {
    if (context.value === "__UNASSIGNED__") {
      filters.region = "all";
      filters.outletCodes = state.data.enrichedOutlets.filter(row => !row.Region && row.OutletCode).map(row => row.OutletCode);
    } else {
      filters.region = context.value;
    }
  } else if (context.type === "category") {
    filters.category = context.value;
    filters.articleNo = "all";
  }
  return filters;
}

function openDialog() {
  if (!dom.detailDialog.open) dom.detailDialog.showModal();
}

function detailCacheKey(filters, tableNumber) {
  return JSON.stringify({ tableNumber, mode: incidentTableType(), filters });
}

function selectedMetricValue(metric, context) {
  const valueFor = row => {
    if (!row) return null;
    if (metric === "Gap") {
      const receiving = finite(row.Receiving);
      const sales = finite(row.Sales);
      return receiving == null && sales == null ? null : (receiving ?? 0) - (sales ?? 0);
    }
    if (metric === "Incidents") {
      const over = finite(row.OverIncidents);
      const under = finite(row.UnderIncidents);
      return over == null && under == null ? null : (over ?? 0) + (under ?? 0);
    }
    return finite(row[metric]);
  };
  if (metric === "UnderValue") {
    if (context?.type === "region") return underValueFor("regions", context.value === "__UNASSIGNED__" ? "" : context.value);
    if (context?.type === "category") return underValueFor("categories", context.value);
    if (!context) return finite(state.data.resolvedUnderValue);
  }
  if (context?.type === "outlet") return valueFor((state.data.enrichedOutlets || []).find(row => normalizeOutletCode(row.OutletCode) === normalizeOutletCode(context.value)));
  if (context?.type === "region") return valueFor((state.data.regions || []).find(row => String(row.Region ?? "__UNASSIGNED__") === String(context.value)));
  if (context?.type === "category") return valueFor((state.data.categories || []).find(row => String(row.Category) === String(context.value)));
  if (metric === "OverValue" && finite(state.data.resolvedOverValue) != null) return finite(state.data.resolvedOverValue);
  return valueFor(state.data.kpis?.[0]);
}

function initialManagementTable(metric, context) {
  if (["OverIncidents", "OverIncidentPct", "UnderIncidents", "UnderIncidentPct", "Incidents"].includes(metric)) return 6;
  if (context?.type === "category") return 4;
  if (context?.type === "outlet") return 1;
  if (context?.type === "region" || metric === "ActiveOutlets" || metric === "OverValue" || metric === "UnderValue") return 3;
  return 1;
}

function detailScopeLabel() {
  const labels = [];
  if (state.detailContext?.label) labels.push(state.detailContext.label);
  else if (state.filters.outletCode !== "all") labels.push(canonicalOutletLabel(state.filters.outletCode));
  else if (state.filters.zonal !== "all") labels.push(state.filters.zonal);
  else if (state.filters.rho !== "all") labels.push(state.filters.rho);
  else if (state.filters.region !== "all") labels.push(state.filters.region);
  else labels.push(state.filters.masterCategory === "all" ? "All business divisions" : "Current dashboard scope");

  if (state.filters.masterCategory !== "all") labels.push(`Business division ${state.filters.masterCategory}`);
  if (state.filters.category !== "all" && state.filters.category !== state.detailContext?.value) labels.push(state.filters.category);
  if (state.filters.articleNo !== "all") labels.push(`Article ${state.filters.articleNo}`);
  if (state.filters.userCode !== "all") labels.push(`User ${state.filters.userCode}`);
  if (state.filters.poNumber !== "all") labels.push(`PO ${state.filters.poNumber}`);
  if (state.filters.movementCode !== "all") labels.push(`Movement ${state.filters.movementCode}`);
  if (state.detailRowFilters.outletCodes?.length) labels.push(`Outlet ${state.detailRowFilters.outletCodes.join(", ")}`);
  if (state.detailRowFilters.category) labels.push(state.detailRowFilters.category);
  if (state.detailRowFilters.articleNo) labels.push(`Article ${state.detailRowFilters.articleNo}`);
  if (state.detailRowFilters.userCode) labels.push(`User ${state.detailRowFilters.userCode}`);
  if (state.detailRowFilters.movementCode) labels.push(`Movement ${state.detailRowFilters.movementCode}`);
  if (state.detailSourceSearch?.articleNo) labels.push(`Article ${state.detailSourceSearch.articleNo}`);
  if (state.detailWidened) labels.push("All business divisions");
  if (state.detailRelaxMovement) labels.push("All movement types");
  return [...new Set(labels)].join(" · ");
}

function currentManagementFilters() {
  const filters = detailFiltersForContext(state.detailContext);
  // The first five tables reproduce the saved Power BI "Over Receiving" page.
  // If the overview combines divisions, keep the page's saved division there.
  state.detailPinnedDivision = null;
  if (state.detailRelaxMovement) filters.allMovementTypes = true;
  if (state.detailTableNumber !== 6 && !state.detailFollowsIncidentScope && !state.detailWidened
    && filters.masterCategory === "all" && state.data?.scope?.masterCategory) {
    filters.masterCategory = state.data.scope.masterCategory;
    state.detailPinnedDivision = filters.masterCategory;
  }
  const additions = state.detailRowFilters;
  if (additions.outletCodes?.length) {
    filters.region = "all";
    filters.outletCodes = additions.outletCodes;
  }
  for (const key of ["category", "articleNo", "userCode", "poNumber", "movementCode"]) {
    if (additions[key]) filters[key] = additions[key];
  }
  if (state.detailSourceSearch?.articleNo) {
    filters.articleNo = state.detailSourceSearch.articleNo;
    // An exact search follows the dashboard's active business-division scope.
    // When the dashboard is on All, query the article across all divisions.
    if (state.filters.masterCategory === "all") filters.masterCategory = "all";
  }
  return filters;
}

function configureManagementTable(tableNumber) {
  const table = managementTableDefinition(tableNumber);
  state.detailTableNumber = Number(tableNumber) || 1;
  state.detailColumns = table.columns;
  state.sorts.detail = { key: table.defaultSort, direction: table.defaultDirection || "desc" };
  state.detailColumnFilters = {};
  // Force the header, and with it the per-column search inputs, to rebuild.
  state.detailHeaderSignature = "";
  dom.detailTableGrid.classList.toggle("incident-user-table", state.detailTableNumber === 6);
  setText("detail-title", table.title);
  setText("detail-context", `${detailScopeLabel()} · ${dateRangeLabel(state.data.range)}`);
  dom.detailSearch.placeholder = state.detailTableNumber === 5
    ? "Search article, PO, movement, user or date"
    : state.detailTableNumber === 6
      ? "Search outlet, RHO, Zonal or user"
      : "Search any value in this management table";
  dom.detailTabs.querySelectorAll("[data-management-table]").forEach(button => {
    const active = Number(button.dataset.managementTable) === state.detailTableNumber;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-current", active ? "page" : "false");
  });
}

function normalizeManagementRows(rows) {
  return (rows || []).map(row => {
    const inferredCode = String(row.OutletName || "").match(/^([A-Za-z]\d{3,})\b/)?.[1] || "";
    const outletCode = normalizeOutletCode(row.OutletCode || inferredCode);
    const organization = organizationForCode(outletCode);
    return {
      ...row,
      OutletCode: outletCode || null,
      OutletName: organization?.OutletName || row.OutletName || (outletCode ? outletCode : "Unassigned / HO"),
      RHO: organization?.RHO || "Not mapped",
      Zonal: organization?.Zonal || "Not mapped",
    };
  });
}

async function loadManagementTable(tableNumber, { preserveRowFilters = true } = {}) {
  // Row filters are cleared when moving between tabs, but the cross-division
  // scope of an incident drill is a property of the drill, not of one tab.
  if (!preserveRowFilters) state.detailRowFilters = {};
  state.detailMissingColumns = [];
  state.detailWidened = false;
  state.detailRelaxMovement = false;
  configureManagementTable(tableNumber);
  const sequence = ++state.detailSequence;
  dom.detailLoading.hidden = false;
  dom.detailError.hidden = true;
  dom.detailContent.hidden = true;

  try {
    const fetchRows = async () => {
      const filters = currentManagementFilters();
      const pinned = state.detailPinnedDivision;
      const cacheKey = detailCacheKey(filters, state.detailTableNumber);
      const cached = state.detailCache.get(cacheKey);
      if (cached && Date.now() - cached.savedAt < DETAIL_CACHE_MS) {
        return { rows: cached.rows, missing: cached.missing || [], pinned };
      }
      const result = await state.client.loadManagementTable(filters, state.detailTableNumber, {
        range: state.data.range,
        scope: state.data.scope,
        mode: incidentTableType(),
      });
      const rows = normalizeManagementRows(result.rows);
      const missing = result.missing || [];
      state.detailCache.set(cacheKey, { savedAt: Date.now(), rows, missing });
      return { rows, missing, pinned };
    };

    let outcome = await fetchRows();
    if (sequence !== state.detailSequence) return;
    // An empty table is usually a filter carried over from the saved Over
    // Receiving page rather than missing data. Relax those filters one at a
    // time, and report whichever relaxation returned rows.
    if (!outcome.rows.length && outcome.pinned) {
      state.detailWidened = true;
      outcome = await fetchRows();
      if (sequence !== state.detailSequence) return;
      if (!outcome.rows.length) state.detailWidened = false;
    }
    if (!outcome.rows.length) {
      state.detailRelaxMovement = true;
      outcome = await fetchRows();
      if (sequence !== state.detailSequence) return;
      if (!outcome.rows.length) {
        // Last attempt: no division pin and no movement-type restriction.
        state.detailWidened = true;
        outcome = await fetchRows();
        if (sequence !== state.detailSequence) return;
      }
      if (!outcome.rows.length) {
        state.detailRelaxMovement = false;
        state.detailWidened = false;
      }
    }
    state.detailRows = outcome.rows;
    state.detailMissingColumns = outcome.missing;
    if (sequence !== state.detailSequence) return;
    // Re-apply the definition now that unpublished measures are known.
    configureManagementTable(state.detailTableNumber);
    dom.detailLoading.hidden = true;
    dom.detailContent.hidden = false;
    renderDetail();
  } catch (error) {
    if (sequence !== state.detailSequence) return;
    dom.detailLoading.hidden = true;
    dom.detailError.hidden = false;
    dom.detailError.textContent = `${error?.message || "The management table could not be loaded."} Close this panel and try again.`;
  }
}

async function openDrill(metric, context = null, rawValue = null) {
  if (!state.data) return;
  const implied = String(metric).startsWith("Under")
    ? "under"
    : (String(metric).startsWith("Over") && metric !== "OverValue" ? "over" : null);
  if (implied && implied !== state.incidentMode) selectIncidentMode(implied, { refreshDialog: false });
  dom.incidentFilter.value = state.incidentMode;
  state.detailMetric = metric;
  state.detailContext = context;
  state.detailDrillValue = finite(rawValue) ?? selectedMetricValue(metric, context);
  state.detailRowFilters = {};
  state.detailFollowsIncidentScope = false;
  state.detailWidened = false;
  state.detailRelaxMovement = false;
  state.detailSourceSearch = null;
  window.clearTimeout(state.detailSearchTimer);
  state.detailSearch = "";
  dom.detailSearch.value = "";
  openDialog();
  await loadManagementTable(initialManagementTable(metric, context));
}

function renderDetailSummary(totalRows) {
  const definition = metricDefinition(state.detailMetric);
  const value = state.detailDrillValue == null ? "—" : definition.formatter(state.detailDrillValue);
  const isValueMetric = ["OverValue", "UnderValue"].includes(state.detailMetric);
  const overValueGrouped = isValueMetric && !state.detailContext
    && (state.detailMetric === "UnderValue" ? state.data?.underValueUsesGroupedContext : state.data?.overValueUsesGroupedContext);
  const cards = [
    [overValueGrouped ? "Grouped source value" : isValueMetric ? "Source DAX total" : "Clicked source value", value],
    ["Returned rows", exact(totalRows)],
    ["Data window", dateRangeLabel(state.data.range)],
    ["Selected scope", detailScopeLabel()],
  ];
  dom.detailSummary.innerHTML = cards.map(([label, cardValue]) => `<div class="detail-summary-card"><span>${escapeHtml(label)}</span><strong title="${escapeHtml(cardValue)}">${escapeHtml(cardValue)}</strong></div>`).join("");
  const valueName = state.detailMetric === "UnderValue" ? "Under Receiving Value" : "Over Receiving Value";
  dom.detailMeasureNote.innerHTML = overValueGrouped
    ? `<strong>Calculation rule:</strong> Power BI returns a blank grand total for ${valueName}, so the headline uses the complete category-level source results for the selected scope; visible drill-down rows do not recalculate it.`
    : isValueMetric
      ? `<strong>Why totals can differ:</strong> Power BI’s ${valueName} is a context-sensitive DAX measure. This view preserves the selected source total exactly; outlet and article rows are drill-down values and are never added to replace it.`
    : "<strong>Calculation rule:</strong> the clicked Power BI value and every management row are queried in the same date and filter context; displayed rows are not used to recalculate the source headline.";
}

function applySort(table, key, additive = false) {
  const current = state.sorts[table];
  const defaultDirection = ["Region", "Outlet", "RHO", "Zonal", "ArticleNo", "ArticleName", "MasterCategory", "Category", "Position"].includes(key) ? "asc" : "desc";
  if (!additive) {
    state.sorts[table] = current.key === key
      ? { key, direction: current.direction === "asc" ? "desc" : "asc" }
      : { key, direction: defaultDirection };
  } else {
    const criteria = [{ key: current.key, direction: current.direction }, ...(current.secondary || [])];
    const existing = criteria.find(item => item.key === key);
    if (existing) existing.direction = existing.direction === "asc" ? "desc" : "asc";
    else criteria.push({ key, direction: defaultDirection });
    state.sorts[table] = { ...criteria[0], secondary: criteria.slice(1) };
  }
  if (table === "region") renderRegions(state.data?.regions || []);
  if (table === "outlet") renderOutlets();
  if (table === "detail") renderDetail();
}

function columnSearchLabel(column) {
  return column.status ? "status" : column.label;
}

function positionColumnFilters() {
  const row = dom.detailHead.parentElement?.querySelector(".column-filter-row");
  if (!row) return;
  const offset = dom.detailHead.offsetHeight;
  row.querySelectorAll("th").forEach(cell => { cell.style.top = `${offset}px`; });
}

function renderColumnFilterRow() {
  const head = dom.detailHead.parentElement;
  if (!head) return;
  let row = head.querySelector(".column-filter-row");
  if (!row) {
    row = document.createElement("tr");
    row.className = "column-filter-row";
    head.append(row);
  }
  row.innerHTML = state.detailColumns.map(column => {
    const label = columnSearchLabel(column);
    const classes = column.numeric ? ' class="numeric"' : column.status ? ' class="status-column"' : "";
    return `<th${classes}><input class="column-search" type="search" autocomplete="off" data-column-filter="${escapeHtml(column.key)}" value="${escapeHtml(state.detailColumnFilters[column.key] || "")}" placeholder="Search" aria-label="Search ${escapeHtml(label)}" title="Search ${escapeHtml(label)}"></th>`;
  }).join("");
  window.requestAnimationFrame(positionColumnFilters);
}

function renderDetailHeader() {
  const signature = `${state.detailTableNumber}:${state.detailColumns.map(column => column.key).join("|")}`;
  // Rebuilding the header on every keystroke would drop focus out of the
  // per-column search inputs, so it is only rebuilt when the columns change.
  if (signature && signature === state.detailHeaderSignature) {
    positionColumnFilters();
    return;
  }
  state.detailHeaderSignature = signature;
  dom.detailHead.innerHTML = state.detailColumns.map(column => `<th scope="col" data-column-key="${escapeHtml(column.key)}"${column.numeric ? ' class="numeric"' : column.status ? ' class="status-column"' : ""}><button class="sort-button" type="button" data-sort-table="detail" data-sort-key="${escapeHtml(column.key)}">${escapeHtml(column.label)} <span></span></button></th>`).join("");
  dom.detailHead.querySelectorAll(".sort-button").forEach(button => {
    button.addEventListener("click", event => applySort("detail", button.dataset.sortKey, event.shiftKey));
  });
  renderColumnFilterRow();
}

function managementDisplay(row, column) {
  const value = row[column.key];
  if (column.date) return longDate(value);
  if (column.status) return value || "—";
  if (column.numeric) {
    const number = finite(value);
    if (number == null) return "—";
    return column.decimals == null ? exact(number) : number.toLocaleString("en-GB", { minimumFractionDigits: column.decimals, maximumFractionDigits: column.decimals });
  }
  return value == null || value === "" ? "—" : String(value);
}

function managementAction(row, column) {
  const table = state.detailTableNumber;
  if (table === 1) {
    if (column.key === "OutletName" && row.OutletCode) return { type: "outlet", value: row.OutletCode, label: row.OutletName };
    if (column.key === "Category" && row.Category) return { type: "category", value: row.Category, label: row.Category };
    if (row.ArticleNo) return { type: "article", value: row.ArticleNo, label: `${row.ArticleNo} — ${row.ArticleName || "Article"}`, outlet: row.OutletCode };
  }
  if ((table === 2 || table === 3) && row.OutletCode) return { type: "outlet", value: row.OutletCode, label: row.OutletName };
  if (table === 4 && row.Category) return { type: "category", value: row.Category, label: row.Category };
  if (table === 5) {
    if (column.key === "ArticleNo" && row.ArticleNo) return { type: "article", value: row.ArticleNo, label: `Article ${row.ArticleNo}` };
    if (column.key === "MovementType" && row.MovementType) return { type: "movement", value: row.MovementType, label: `Movement ${row.MovementType}` };
    if (column.key === "CreatedBy" && row.CreatedBy) return { type: "user", value: row.CreatedBy, label: `User ${row.CreatedBy}` };
  }
  if (table === 6) {
    if (column.key === "CreatedBy" && row.CreatedBy) return { type: "user", value: row.CreatedBy, label: `User ${row.CreatedBy}` };
    // Every incident count and rate opens the category-level incident list for
    // that outlet and user.
    if (["OverIncidents", "OverIncidentPct", "UnderIncidents", "UnderIncidentPct"].includes(column.key) && row.OutletCode) {
      return {
        type: "incidents",
        value: row.OutletCode,
        label: `${row.OutletName || row.OutletCode}${row.CreatedBy ? ` · ${row.CreatedBy}` : ""}`,
        user: row.CreatedBy || "",
      };
    }
  }
  return null;
}

function managementCell(row, column) {
  const display = managementDisplay(row, column);
  if (column.status) return `<span class="management-status-icon" title="Over-receiving score status">${escapeHtml(display)}</span>`;
  const action = managementAction(row, column);
  if (!action || display === "—") return escapeHtml(display);
  const extra = `${action.outlet ? ` data-management-outlet="${escapeHtml(action.outlet)}"` : ""}${action.user ? ` data-management-user="${escapeHtml(action.user)}"` : ""}`;
  const className = column.numeric ? "number-link" : "text-link";
  return `<button class="${className} management-cell-link" type="button" data-management-action="${escapeHtml(action.type)}" data-management-value="${escapeHtml(action.value)}" data-management-label="${escapeHtml(action.label)}"${extra} title="Open linked management detail">${escapeHtml(display)}</button>`;
}

function renderDetail() {
  const query = state.detailSearch.trim().toLocaleLowerCase();
  const incidentKey = incidentTableType() === "under" ? "UnderIncidents" : "OverIncidents";
  const tableRows = state.detailTableNumber === 6
    ? state.detailRows.filter(row => (finite(row[incidentKey]) ?? 0) > 0)
    : state.detailRows;
  const columnQueries = state.detailColumns
    .map(column => ({ column, needle: String(state.detailColumnFilters[column.key] || "").trim().toLocaleLowerCase() }))
    .filter(item => item.needle);
  const matchesColumn = (row, column, needle) => {
    const display = String(managementDisplay(row, column) ?? "").toLocaleLowerCase();
    if (display.includes(needle)) return true;
    // Let "2480" find a value shown as "2,480".
    return display.replaceAll(",", "").includes(needle.replaceAll(",", ""));
  };
  const filtered = tableRows.filter(row => {
    if (query && ![row.OutletCode, ...state.detailColumns.map(column => managementDisplay(row, column))]
      .some(value => String(value ?? "").toLocaleLowerCase().includes(query))) return false;
    return columnQueries.every(({ column, needle }) => matchesColumn(row, column, needle));
  });
  const sorted = sortRows(filtered, "detail");
  const visible = sorted.slice(0, DETAIL_ROW_LIMIT);
  state.visibleRows.detail = visible;
  setExportAvailability("detail", visible.length > 0);
  renderDetailHeader();
  updateSortIndicators("detail");
  renderDetailSummary(tableRows.length);
  setText("detail-result-count", visible.length < sorted.length ? `Showing ${exact(visible.length)} of ${exact(sorted.length)} matches` : `${exact(sorted.length)} rows`);

  if (!visible.length) {
    const searching = query || Object.values(state.detailColumnFilters).some(value => String(value || "").trim());
    dom.detailTable.innerHTML = `<tr><td colspan="${state.detailColumns.length}" class="empty-cell">${searching
      ? "No rows match the current search. Clear the column boxes to see the full table."
      : `The source returned no ${incidentTableType()}-receiving rows for ${escapeHtml(detailScopeLabel())}.`}</td></tr>`;
    return;
  }

  dom.detailTable.innerHTML = visible.map(row => `<tr>${state.detailColumns.map(column => {
    const display = managementDisplay(row, column);
    return `<td data-column-key="${escapeHtml(column.key)}" title="${escapeHtml(display)}"${column.numeric ? ' class="numeric"' : column.status ? ' class="status-column"' : ""}>${managementCell(row, column)}</td>`;
  }).join("")}</tr>`).join("");
}

function detailArticleCandidate(value) {
  const candidate = String(value || "").trim();
  return /^\d{5,}$/.test(candidate) ? candidate : null;
}

function runDetailSourceSearch(articleNo) {
  if (![1, 5].includes(state.detailTableNumber) || !articleNo) return;
  if (state.detailSourceSearch?.articleNo === articleNo) return;
  state.detailSourceSearch = { articleNo };
  loadManagementTable(state.detailTableNumber);
}

function handleDetailSearchInput(value, immediate = false) {
  window.clearTimeout(state.detailSearchTimer);
  state.detailSearch = value;
  const candidate = detailArticleCandidate(value);

  if (!candidate || ![1, 5].includes(state.detailTableNumber)) {
    if (state.detailSourceSearch) {
      state.detailSourceSearch = null;
      loadManagementTable(state.detailTableNumber);
    } else {
      renderDetail();
    }
    return;
  }

  renderDetail();
  if (state.detailSourceSearch?.articleNo === candidate) return;
  if (immediate) runDetailSourceSearch(candidate);
  else state.detailSearchTimer = window.setTimeout(() => runDetailSourceSearch(candidate), 450);
}

async function loadDashboard({ refreshMetadata = false } = {}) {
  const sequence = ++state.loadSequence;
  if (state.loadController) state.loadController.abort();
  state.loadController = null;
  if (state.backgroundStatusTimer) window.clearTimeout(state.backgroundStatusTimer);
  const blockingLoad = !state.data;
  if (blockingLoad) setLoading(true);
  clearError();
  if (!navigator.onLine) {
    if (blockingLoad) {
      setLoading(false);
      showError(new Error("This device is offline."));
    } else {
      dom.statusDot.className = "status-dot is-error";
      dom.connectionStatus.textContent = "Device is offline";
      dom.sourceFreshness.textContent = "Showing the last successfully loaded snapshot";
    }
    return;
  }

  try {
    const organizationPromise = refreshMetadata || !state.organizationPromise
      ? (state.organizationPromise = refreshOrganization())
      : state.organizationPromise;
    const snapshotPromise = !state.sharedSnapshot || refreshMetadata
      ? fetchSharedSnapshot()
      : Promise.resolve(state.sharedSnapshot);
    const [snapshot] = await Promise.all([snapshotPromise, organizationPromise]);
    if (sequence !== state.loadSequence) return;
    state.sharedSnapshot = snapshot;
    state.nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    saveDashboardCache(snapshot, { force: true });

    if (snapshotMatchesCurrentFilters(snapshot)) {
      state.data = snapshot;
      renderAll(snapshot);
      dom.statusDot.className = "status-dot";
      dom.connectionStatus.textContent = "Shared Power BI snapshot";
      dom.sourceFreshness.textContent = `Power BI data refreshed ${dhakaDateTime(snapshot.sourceTimestamp || snapshot.snapshotGeneratedAt || snapshot.queryTimestamp)}`;
      return;
    }

    const embeddedSelection = embeddedRangeSnapshot(snapshot);
    if (embeddedSelection) {
      state.data = embeddedSelection;
      renderAll(embeddedSelection);
      dom.statusDot.className = "status-dot";
      dom.connectionStatus.textContent = "Shared Power BI snapshot";
      dom.sourceFreshness.textContent = `Selected date range loaded instantly · Power BI data refreshed ${dhakaDateTime(snapshot.sourceTimestamp || snapshot.snapshotGeneratedAt || snapshot.queryTimestamp)}`;
      return;
    }

    if (refreshMetadata) {
      state.client = new PowerBIDataClient();
      state.detailCache.clear();
    }
    const requestFilters = buildPowerBIFilters();
    const cachedSelection = await readFilteredSnapshot(requestFilters, snapshot.sourceTimestamp);
    if (sequence !== state.loadSequence) return;
    if (cachedSelection) {
      state.data = cachedSelection;
      renderAll(cachedSelection);
      dom.statusDot.className = "status-dot";
      dom.connectionStatus.textContent = "Cached selected snapshot";
      dom.sourceFreshness.textContent = `Loaded instantly · Power BI data refreshed ${dhakaDateTime(snapshot.sourceTimestamp || snapshot.snapshotGeneratedAt || snapshot.queryTimestamp)}`;
      if (!refreshMetadata) return;
    } else {
      state.data = snapshot;
      renderAll(snapshot);
      dom.statusDot.className = "status-dot is-loading";
      dom.connectionStatus.textContent = "Latest snapshot displayed";
      dom.sourceFreshness.textContent = "Preparing the exact selected range in the background · controls remain available";
    }
    setLoading(false);

    const controller = new AbortController();
    state.loadController = controller;
    state.backgroundStatusTimer = window.setTimeout(() => {
      if (sequence !== state.loadSequence) return;
      dom.connectionStatus.textContent = "Latest snapshot displayed";
      dom.sourceFreshness.textContent = "Power BI is still processing the exact selection in the background";
    }, 8000);

    const data = await state.client.load(requestFilters, { section: "core", signal: controller.signal });
    if (sequence !== state.loadSequence) return;
    if (state.backgroundStatusTimer) window.clearTimeout(state.backgroundStatusTimer);
    state.backgroundStatusTimer = 0;
    state.data = coreWithSnapshotOptions(data, snapshot);
    state.nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    renderAll(state.data);
    dom.statusDot.className = "status-dot";
    dom.connectionStatus.textContent = "Live Power BI data";
    dom.sourceFreshness.textContent = `Overview ready · loading outlet and search data…`;

    try {
      const supporting = await state.client.load(requestFilters, { section: "supporting", signal: controller.signal });
      if (sequence !== state.loadSequence) return;
      state.data = { ...state.data, ...supporting };
      renderAll(state.data);
      await writeFilteredSnapshot(requestFilters, snapshot.sourceTimestamp, state.data);
      dom.sourceFreshness.textContent = `Model refreshed ${dhakaDateTime(state.data.sourceTimestamp)}`;
    } catch (error) {
      if (sequence !== state.loadSequence || error?.name === "AbortError") return;
      console.warn("Supporting dashboard data could not be loaded", error);
      dom.sourceFreshness.textContent = `Overview ready · outlet/search data stayed on the last snapshot`;
    }
  } catch (error) {
    if (sequence !== state.loadSequence || error?.name === "AbortError") return;
    console.error("Dashboard refresh failed", error);
    if (state.data) {
      clearError();
      dom.statusDot.className = "status-dot is-error";
      dom.connectionStatus.textContent = "Last saved snapshot";
      dom.sourceFreshness.textContent = "Exact live selection unavailable · the latest working snapshot remains visible";
    } else {
      showError(error);
    }
  } finally {
    if (sequence === state.loadSequence) {
      if (state.backgroundStatusTimer) window.clearTimeout(state.backgroundStatusTimer);
      state.backgroundStatusTimer = 0;
      state.loadController = null;
      setLoading(false);
    }
  }
}

function applySearchSelection(type) {
  const input = type === "outlet" ? dom.outletFilter : dom.articleFilter;
  const value = input.value.trim();
  const suggestions = type === "outlet" ? state.outletSuggestions : state.articleSuggestions;
  const key = type === "outlet" ? "outletCode" : "articleNo";
  if (!value) {
    if (state.filters[key] !== "all") {
      state.filters[key] = "all";
      loadDashboard();
    }
    return;
  }

  const exactLabel = suggestions.find(label => label.toLocaleLowerCase() === value.toLocaleLowerCase());
  const codeCandidate = value.split("—")[0].trim();
  const direct = suggestions.find(label => label.split("—")[0].trim().toLocaleLowerCase() === codeCandidate.toLocaleLowerCase());
  const nameMatches = suggestions.filter(label => label.split("—").slice(1).join("—").trim().toLocaleLowerCase() === value.toLocaleLowerCase());
  const match = exactLabel || direct || (nameMatches.length === 1 ? nameMatches[0] : null);
  const directArticleCode = type === "article" && /^[A-Za-z0-9][A-Za-z0-9._/-]{1,31}$/.test(codeCandidate)
    ? codeCandidate
    : null;
  if (!match && !directArticleCode) {
    dom.cascadeNote.textContent = `No exact ${type} match. Select a suggestion or enter an exact code.`;
    input.setAttribute("aria-invalid", "true");
    return;
  }

  input.removeAttribute("aria-invalid");
  state.filters[key] = match ? match.split("—")[0].trim() : directArticleCode;
  input.value = match || directArticleCode;
  dom.cascadeNote.textContent = match
    ? "Applying selection; all related filter options will shorten automatically."
    : `Checking exact article code ${directArticleCode} in the selected outlet and date range.`;
  loadDashboard();
}

let dateApplyTimer = 0;

function applyCustomDateRange() {
  const dateFrom = isoDate(dom.fromDateFilter.value);
  const dateTo = isoDate(dom.toDateFilter.value);
  if (!dateFrom || !dateTo) {
    dom.cascadeNote.textContent = "Select both From date and To date to apply a custom range.";
    return;
  }

  const latestTo = inclusiveEndIso({ endExclusive: state.data?.scope?.endExclusive || state.data?.range?.endExclusive });
  const invalidOrder = dateFrom > dateTo;
  const beyondLatest = latestTo && dateTo > latestTo;
  if (invalidOrder) dom.fromDateFilter.setAttribute("aria-invalid", "true");
  else dom.fromDateFilter.removeAttribute("aria-invalid");
  if (invalidOrder || beyondLatest) dom.toDateFilter.setAttribute("aria-invalid", "true");
  else dom.toDateFilter.removeAttribute("aria-invalid");
  if (invalidOrder) {
    dom.cascadeNote.textContent = "From date cannot be later than To date.";
    return;
  }
  if (beyondLatest) {
    dom.cascadeNote.textContent = `The latest available dashboard date is ${latestTo}.`;
    return;
  }

  const changed = state.filters.dateFrom !== dateFrom || state.filters.dateTo !== dateTo;
  state.filters.dateFrom = dateFrom;
  state.filters.dateTo = dateTo;
  dom.periodFilter.value = "custom";
  dom.cascadeNote.textContent = `Custom date range ${dateFrom} to ${dateTo} selected · loading the latest snapshot first.`;
  if (changed) loadDashboard();
}

function scheduleCustomDateRange() {
  if (dateApplyTimer) window.clearTimeout(dateApplyTimer);
  dom.cascadeNote.textContent = "Date selection detected · waiting briefly for both From and To dates.";
  dateApplyTimer = window.setTimeout(() => {
    dateApplyTimer = 0;
    applyCustomDateRange();
  }, 900);
}

function applyOrganizationSearch(type) {
  const isRho = type === "rho";
  const input = isRho ? dom.rhoFilter : dom.zonalFilter;
  const suggestions = isRho ? state.rhoSuggestions : state.zonalSuggestions;
  const value = input.value.trim();
  if (!value) {
    if (state.filters[type] !== "all") {
      state.filters[type] = "all";
      if (isRho) state.filters.zonal = "all";
      state.filters.outletCode = "all";
      dom.outletFilter.value = "";
      updateCascadingOptions();
      loadDashboard();
    }
    return;
  }

  const exactMatch = suggestions.find(item => item.toLocaleLowerCase() === value.toLocaleLowerCase());
  const partialMatches = suggestions.filter(item => item.toLocaleLowerCase().includes(value.toLocaleLowerCase()));
  const match = exactMatch || (partialMatches.length === 1 ? partialMatches[0] : null);
  if (!match) {
    input.setAttribute("aria-invalid", "true");
    dom.cascadeNote.textContent = `Select one ${type === "rho" ? "RHO" : "Zonal"} from the shortened suggestions.`;
    return;
  }

  input.removeAttribute("aria-invalid");
  input.value = match;
  state.filters[type] = match;
  if (isRho) {
    state.filters.zonal = "all";
    dom.zonalFilter.value = "";
  }
  state.filters.outletCode = "all";
  dom.outletFilter.value = "";
  updateCascadingOptions();
  loadDashboard();
}

function applyCodeSearch(type) {
  const isUser = type === "userCode";
  const input = isUser ? dom.userFilter : dom.movementFilter;
  const suggestions = isUser ? state.userSuggestions : state.movementSuggestions;
  const value = input.value.trim();
  if (!value) {
    if (state.filters[type] !== "all") {
      state.filters[type] = "all";
      loadDashboard();
    }
    return;
  }
  const exact = suggestions.find(item => item.toLocaleLowerCase() === value.toLocaleLowerCase());
  const partial = suggestions.filter(item => item.toLocaleLowerCase().includes(value.toLocaleLowerCase()));
  const match = exact || (partial.length === 1 ? partial[0] : null);
  if (!match) {
    input.setAttribute("aria-invalid", "true");
    dom.cascadeNote.textContent = `Select an exact ${isUser ? "user" : "movement"} code from the suggestions.`;
    return;
  }
  input.removeAttribute("aria-invalid");
  input.value = match;
  state.filters[type] = match;
  dom.cascadeNote.textContent = `Applying ${isUser ? "user" : "movement"} code ${match}.`;
  loadDashboard();
}

function applyPoSearch() {
  const value = dom.poFilter.value.trim();
  if (!value) {
    dom.poFilter.removeAttribute("aria-invalid");
    if (state.filters.poNumber !== "all") {
      state.filters.poNumber = "all";
      loadDashboard();
    }
    return;
  }
  if (!/^[A-Za-z0-9][A-Za-z0-9._/-]{0,63}$/.test(value)) {
    dom.poFilter.setAttribute("aria-invalid", "true");
    dom.cascadeNote.textContent = "Enter one exact PO number without spaces, then press Enter.";
    return;
  }
  dom.poFilter.removeAttribute("aria-invalid");
  state.filters.poNumber = value;
  dom.cascadeNote.textContent = `Checking exact PO number ${value} in the selected date and outlet scope.`;
  loadDashboard();
}

function selectIncidentMode(mode, { refreshDialog = true } = {}) {
  const next = mode === "under" ? "under" : "over";
  const changed = state.incidentMode !== next;
  state.incidentMode = next;
  state.exceptionFocus = next;
  dom.incidentFilter.value = next;
  state.sorts.outlet = { key: focusConfig[state.exceptionFocus].field, direction: "desc" };
  if (["OverValue", "UnderValue"].includes(state.sorts.region.key)) {
    state.sorts.region = { ...state.sorts.region, key: modeValueKey() };
  }

  if (state.data) renderAll(state.data);
  else applyIncidentModeVisibility();

  // A previous attempt that failed on the network is worth retrying when the
  // user deliberately switches back to Under.
  if (next === "under" && state.underValue.status === "error") ensureUnderValues({ force: true });

  if (!refreshDialog || !dom.detailDialog.open || !changed) return;
  if (state.detailTableNumber === 6) {
    state.detailMetric = next === "under" ? "UnderIncidents" : "OverIncidents";
    state.detailDrillValue = selectedMetricValue(state.detailMetric, state.detailContext);
    configureManagementTable(6);
    renderDetail();
    return;
  }
  if ([1, 2, 3, 4].includes(state.detailTableNumber)) {
    if (["OverValue", "UnderValue"].includes(state.detailMetric)) {
      state.detailMetric = modeValueKey();
      state.detailDrillValue = selectedMetricValue(state.detailMetric, state.detailContext);
    }
    loadManagementTable(state.detailTableNumber);
  }
}

function setTheme(theme) {
  const mode = theme === "dark" ? "dark" : "light";
  document.documentElement.dataset.theme = mode;
  dom.themeLabel.textContent = mode === "dark" ? "Light mode" : "Dark mode";
  document.querySelector('meta[name="theme-color"]').content = mode === "dark" ? "#081522" : "#102b4e";
  try { localStorage.setItem("receiving-dashboard-theme", mode); } catch {}
  if (state.data) {
    renderTrend(state.data.trend || []);
    renderCategoryChart(state.data.categories || []);
  }
}

document.querySelectorAll(".view-tab").forEach(tab => {
  tab.addEventListener("click", () => switchView(tab.dataset.view));
  tab.addEventListener("keydown", event => {
    if (!["ArrowLeft", "ArrowRight"].includes(event.key)) return;
    event.preventDefault();
    const view = event.key === "ArrowRight" ? "exceptions" : "overview";
    switchView(view);
    el(`${view}-tab`).focus();
  });
});

document.querySelectorAll(".focus-button").forEach(button => {
  button.addEventListener("click", () => {
    if (["over", "under"].includes(button.dataset.focus)) {
      selectIncidentMode(button.dataset.focus);
      return;
    }
    state.exceptionFocus = button.dataset.focus;
    state.sorts.outlet = { key: focusConfig[state.exceptionFocus].field, direction: "desc" };
    renderOutlets();
  });
});

document.querySelectorAll(".sort-button").forEach(button => {
  button.addEventListener("click", event => applySort(button.dataset.sortTable, button.dataset.sortKey, event.shiftKey));
});

document.querySelectorAll("[data-export-table]").forEach(button => {
  button.addEventListener("click", () => {
    // The user-incident table exports both its own rows and the category rows
    // behind them, which needs a two-sheet workbook rather than a CSV.
    if (button.dataset.exportTable === "detail" && state.detailTableNumber === 6) {
      exportIncidentWorkbook(button);
      return;
    }
    exportVisibleCsv(button.dataset.exportTable);
  });
});

dom.detailTableGrid.addEventListener("input", event => {
  const input = event.target.closest("[data-column-filter]");
  if (!input) return;
  state.detailColumnFilters[input.dataset.columnFilter] = input.value;
  window.clearTimeout(state.detailColumnFilterTimer);
  state.detailColumnFilterTimer = window.setTimeout(renderDetail, 140);
});

dom.detailTableGrid.addEventListener("keydown", event => {
  const input = event.target.closest("[data-column-filter]");
  if (!input || !["Enter", "Escape"].includes(event.key)) return;
  event.preventDefault();
  if (event.key === "Escape") {
    input.value = "";
    state.detailColumnFilters[input.dataset.columnFilter] = "";
  }
  window.clearTimeout(state.detailColumnFilterTimer);
  renderDetail();
});

window.addEventListener("resize", () => {
  if (dom.detailDialog.open) positionColumnFilters();
});

dom.main.addEventListener("click", event => {
  const drill = event.target.closest("[data-drill-metric]");
  if (!drill) return;
  openDrill(drill.dataset.drillMetric, drill.dataset.contextType ? { type: drill.dataset.contextType, value: drill.dataset.contextValue, label: drill.dataset.contextLabel } : null, drill.dataset.drillValue);
});

dom.main.addEventListener("keydown", event => {
  if (!["Enter", " "].includes(event.key)) return;
  const drill = event.target.closest("[data-drill-metric]");
  if (!drill) return;
  event.preventDefault();
  openDrill(drill.dataset.drillMetric, drill.dataset.contextType ? { type: drill.dataset.contextType, value: drill.dataset.contextValue, label: drill.dataset.contextLabel } : null, drill.dataset.drillValue);
});

dom.detailDialog.addEventListener("click", event => {
  const tab = event.target.closest("[data-management-table]");
  if (tab) {
    window.clearTimeout(state.detailSearchTimer);
    state.detailSourceSearch = null;
    state.detailSearch = "";
    dom.detailSearch.value = "";
    loadManagementTable(Number(tab.dataset.managementTable), { preserveRowFilters: false });
    return;
  }

  const link = event.target.closest("[data-management-action]");
  if (!link) return;
  const value = link.dataset.managementValue;
  window.clearTimeout(state.detailSearchTimer);
  state.detailSourceSearch = null;
  state.detailSearch = "";
  dom.detailSearch.value = "";
  if (link.dataset.managementAction === "incidents") {
    state.detailRowFilters = {
      outletCodes: [normalizeOutletCode(value)],
      ...(link.dataset.managementUser ? { userCode: link.dataset.managementUser } : {}),
    };
    // Table 6 spans every business division, so the category list it opens
    // must not fall back to the saved page's single division.
    state.detailFollowsIncidentScope = true;
    state.detailMetric = incidentTableType() === "under" ? "UnderIncidents" : "OverIncidents";
    state.detailContext = { type: "outlet", value: normalizeOutletCode(value), label: link.dataset.managementLabel || value };
    state.detailDrillValue = finite(link.textContent.replaceAll(",", "")) ?? state.detailDrillValue;
    loadManagementTable(4);
  } else if (link.dataset.managementAction === "outlet") {
    state.detailRowFilters = { outletCodes: [normalizeOutletCode(value)] };
    loadManagementTable(1);
  } else if (link.dataset.managementAction === "category") {
    state.detailRowFilters = { category: value, articleNo: "all" };
    loadManagementTable(1);
  } else if (link.dataset.managementAction === "article") {
    state.detailRowFilters = { articleNo: value, ...(link.dataset.managementOutlet ? { outletCodes: [normalizeOutletCode(link.dataset.managementOutlet)] } : {}) };
    loadManagementTable(5);
  } else if (link.dataset.managementAction === "user") {
    state.detailRowFilters = { ...state.detailRowFilters, userCode: value };
    loadManagementTable(5);
  } else if (link.dataset.managementAction === "movement") {
    state.detailRowFilters = { ...state.detailRowFilters, movementCode: value };
    loadManagementTable(5);
  }
});

dom.managementSignals.addEventListener("click", event => {
  const link = event.target.closest("[data-open-exceptions]");
  if (!link) return;
  event.preventDefault();
  switchView("exceptions");
  dom.exceptionsView.scrollIntoView({ behavior: "smooth", block: "start" });
});

dom.outletSearch.addEventListener("input", event => { state.outletSearch = event.target.value; renderOutlets(); });
dom.detailSearch.addEventListener("input", event => handleDetailSearchInput(event.target.value));
dom.detailSearch.addEventListener("change", event => handleDetailSearchInput(event.target.value, true));
dom.detailSearch.addEventListener("keydown", event => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  handleDetailSearchInput(event.currentTarget.value, true);
});

dom.periodFilter.addEventListener("change", event => {
  if (dateApplyTimer) {
    window.clearTimeout(dateApplyTimer);
    dateApplyTimer = 0;
  }
  if (event.target.value === "custom") {
    applyCustomDateRange();
    return;
  }
  state.filters.days = Number(event.target.value);
  state.filters.dateFrom = null;
  state.filters.dateTo = null;
  dom.fromDateFilter.removeAttribute("aria-invalid");
  dom.toDateFilter.removeAttribute("aria-invalid");
  loadDashboard();
});
[dom.fromDateFilter, dom.toDateFilter].forEach(input => input.addEventListener("change", scheduleCustomDateRange));
dom.masterCategoryFilter.addEventListener("change", event => {
  state.filters.masterCategory = event.target.value;
  state.filters.category = "all";
  state.filters.articleNo = "all";
  dom.articleFilter.value = "";
  loadDashboard();
});
dom.categoryFilter.addEventListener("change", event => {
  state.filters.category = event.target.value;
  state.filters.articleNo = "all";
  dom.articleFilter.value = "";
  loadDashboard();
});
dom.regionFilter.addEventListener("change", event => {
  state.filters.region = event.target.value;
  state.filters.rho = "all";
  state.filters.zonal = "all";
  state.filters.outletCode = "all";
  dom.rhoFilter.value = "";
  dom.zonalFilter.value = "";
  dom.outletFilter.value = "";
  updateCascadingOptions();
  loadDashboard();
});
[dom.rhoFilter, dom.zonalFilter].forEach(input => {
  const type = input === dom.rhoFilter ? "rho" : "zonal";
  input.addEventListener("change", () => applyOrganizationSearch(type));
  input.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      event.preventDefault();
      applyOrganizationSearch(type);
    }
  });
  input.addEventListener("input", () => {
    if (!input.value && state.filters[type] !== "all") applyOrganizationSearch(type);
    else dom.cascadeNote.textContent = `Type a ${type === "rho" ? "RHO" : "Zonal"} name, then press Enter or choose a suggestion.`;
  });
});

[dom.outletFilter, dom.articleFilter].forEach(input => {
  const type = input === dom.outletFilter ? "outlet" : "article";
  input.addEventListener("change", () => applySearchSelection(type));
  input.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      event.preventDefault();
      applySearchSelection(type);
    }
  });
  input.addEventListener("input", () => {
    refreshSearchDatalist(type === "outlet" ? dom.outletOptions : dom.articleOptions, type === "outlet" ? state.outletSuggestions : state.articleSuggestions, input.value);
    if (!input.value && state.filters[type === "outlet" ? "outletCode" : "articleNo"] !== "all") applySearchSelection(type);
    else dom.cascadeNote.textContent = `Type an exact ${type} code/name, then press Enter or choose a suggestion.`;
  });
  input.addEventListener("focus", () => refreshSearchDatalist(type === "outlet" ? dom.outletOptions : dom.articleOptions, type === "outlet" ? state.outletSuggestions : state.articleSuggestions, input.value));
});

[[dom.userFilter, "userCode"], [dom.movementFilter, "movementCode"]].forEach(([input, type]) => {
  input.addEventListener("change", () => applyCodeSearch(type));
  input.addEventListener("keydown", event => {
    if (event.key === "Enter") {
      event.preventDefault();
      applyCodeSearch(type);
    }
  });
  input.addEventListener("input", () => {
    refreshSearchDatalist(type === "userCode" ? dom.userOptions : dom.movementOptions, type === "userCode" ? state.userSuggestions : state.movementSuggestions, input.value);
    if (!input.value && state.filters[type] !== "all") applyCodeSearch(type);
    else dom.cascadeNote.textContent = `Type a ${type === "userCode" ? "user" : "movement"} code, then press Enter or choose a suggestion.`;
  });
  input.addEventListener("focus", () => refreshSearchDatalist(type === "userCode" ? dom.userOptions : dom.movementOptions, type === "userCode" ? state.userSuggestions : state.movementSuggestions, input.value));
});

dom.poFilter.addEventListener("change", applyPoSearch);
dom.poFilter.addEventListener("keydown", event => {
  if (event.key !== "Enter") return;
  event.preventDefault();
  applyPoSearch();
});
dom.poFilter.addEventListener("input", () => {
  if (!dom.poFilter.value && state.filters.poNumber !== "all") applyPoSearch();
  else dom.cascadeNote.textContent = "Enter an exact PO number, then press Enter to search.";
});

dom.incidentFilter.addEventListener("change", event => selectIncidentMode(event.target.value));

dom.resetButton.addEventListener("click", () => {
  if (dateApplyTimer) {
    window.clearTimeout(dateApplyTimer);
    dateApplyTimer = 0;
  }
  state.filters = { ...DEFAULT_FILTERS };
  state.incidentMode = "over";
  state.exceptionFocus = "over";
  state.underValue = { key: "", status: "idle", kind: "value", kpi: null, categories: new Map(), regions: new Map() };
  if (["OverValue", "UnderValue"].includes(state.sorts.region.key)) state.sorts.region = { key: "OverValue", direction: "desc" };
  applyIncidentModeVisibility();
  applyModeToRegionHeader();
  dom.periodFilter.value = "30";
  dom.fromDateFilter.removeAttribute("aria-invalid");
  dom.toDateFilter.removeAttribute("aria-invalid");
  dom.masterCategoryFilter.value = "all";
  dom.categoryFilter.value = "all";
  dom.regionFilter.value = "all";
  dom.rhoFilter.value = "";
  dom.zonalFilter.value = "";
  dom.outletFilter.value = "";
  dom.articleFilter.value = "";
  dom.userFilter.value = "";
  dom.poFilter.value = "";
  dom.movementFilter.value = "";
  dom.incidentFilter.value = "over";
  dom.outletSearch.value = "";
  state.outletSearch = "";
  loadDashboard();
});

dom.filterToggle.addEventListener("click", () => {
  const open = dom.filtersPanel.classList.toggle("is-open");
  dom.filterToggle.classList.toggle("is-active", open);
  dom.filterToggle.setAttribute("aria-expanded", String(open));
});

dom.themeButton.addEventListener("click", () => setTheme(document.documentElement.dataset.theme === "dark" ? "light" : "dark"));
dom.refreshButton.addEventListener("click", () => loadDashboard({ refreshMetadata: true }));
dom.retryButton.addEventListener("click", () => loadDashboard({ refreshMetadata: true }));
el("detail-close").addEventListener("click", () => {
  window.clearTimeout(state.detailSearchTimer);
  dom.detailDialog.close();
});

window.addEventListener("offline", () => {
  dom.statusDot.className = "status-dot is-error";
  dom.connectionStatus.textContent = "Device is offline";
  dom.sourceFreshness.textContent = state.data ? "Showing the last successfully loaded view" : "Waiting for a connection";
});
window.addEventListener("online", () => loadDashboard({ refreshMetadata: true }));
document.addEventListener("visibilitychange", () => {
  if (!document.hidden && state.nextRefreshAt && Date.now() >= state.nextRefreshAt) loadDashboard({ refreshMetadata: true });
});
window.setInterval(() => {
  if (!document.hidden && navigator.onLine && state.nextRefreshAt && Date.now() >= state.nextRefreshAt) loadDashboard({ refreshMetadata: true });
}, 60_000);

setTheme(document.documentElement.dataset.theme);
restoreDashboardCache();
export const dashboardReady = loadDashboard({ refreshMetadata: true });
