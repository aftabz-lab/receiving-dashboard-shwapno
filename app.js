import { PowerBIDataClient, POWER_BI_URL } from "./powerbi.js";
import { loadOrganizationSnapshot, normalizeOutletCode } from "./organization.js";

const AUTO_REFRESH_MS = 15 * 60 * 1000;
const DETAIL_CACHE_MS = 5 * 60 * 1000;
const DASHBOARD_CACHE_MS = 24 * 60 * 60 * 1000;
const DASHBOARD_CACHE_KEY = "receiving-dashboard-default-v8";
const DETAIL_ROW_LIMIT = 400;
const DEFAULT_FILTERS = Object.freeze({
  days: 30,
  masterCategory: "all",
  category: "all",
  region: "all",
  rho: "all",
  zonal: "all",
  outletCode: "all",
  articleNo: "all",
});

const nf = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });
const compactNf = new Intl.NumberFormat("en-GB", { notation: "compact", maximumFractionDigits: 1 });
const percentNf = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 1, maximumFractionDigits: 1 });

const state = {
  client: new PowerBIDataClient(),
  organization: null,
  organizationPromise: null,
  data: null,
  filters: { ...DEFAULT_FILTERS },
  activeView: "overview",
  exceptionFocus: "over",
  outletSearch: "",
  rhoSuggestions: [],
  zonalSuggestions: [],
  outletSuggestions: [],
  articleSuggestions: [],
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
  detailSearch: "",
  loadSequence: 0,
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
  masterCategoryFilter: el("master-category-filter"),
  categoryFilter: el("category-filter"),
  regionFilter: el("region-filter"),
  rhoFilter: el("rho-filter"),
  zonalFilter: el("zonal-filter"),
  rhoOptions: el("rho-options"),
  zonalOptions: el("zonal-options"),
  outletFilter: el("outlet-filter"),
  articleFilter: el("article-filter"),
  outletOptions: el("outlet-options"),
  articleOptions: el("article-options"),
  cascadeNote: el("cascade-note"),
  organizationStatus: el("organization-status"),
  activeFilterSummary: el("active-filter-summary"),
  activeFilterCount: el("active-filter-count"),
  statusDot: el("status-dot"),
  connectionStatus: el("connection-status"),
  sourceFreshness: el("source-freshness"),
  errorPanel: el("error-panel"),
  errorMessage: el("error-message"),
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
  detailHead: el("detail-table-head"),
  detailTable: el("detail-table-body"),
  selectedArticle: el("selected-article-detail"),
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

function toDate(value) {
  if (value instanceof Date) return value;
  const numeric = typeof value === "string" && /^\d+$/.test(value) ? Number(value) : value;
  const date = new Date(numeric);
  return Number.isNaN(date.getTime()) ? null : date;
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

function saveDashboardCache(data) {
  if (!filtersAreDefault()) return;
  try {
    localStorage.setItem(DASHBOARD_CACHE_KEY, JSON.stringify({ savedAt: Date.now(), data }));
  } catch {}
}

function restoreDashboardCache() {
  try {
    const cached = JSON.parse(localStorage.getItem(DASHBOARD_CACHE_KEY) || "null");
    if (!cached?.data || !cached.savedAt || Date.now() - cached.savedAt > DASHBOARD_CACHE_MS) return false;
    state.data = cached.data;
    renderAll(state.data);
    dom.statusDot.className = "status-dot is-loading";
    dom.connectionStatus.textContent = "Refreshing live data";
    dom.sourceFreshness.textContent = `Showing saved view from ${dhakaDateTime(cached.savedAt)}`;
    return true;
  } catch {
    return false;
  }
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

function fillDatalist(datalist, values) {
  const fragment = document.createDocumentFragment();
  values.forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    fragment.append(option);
  });
  datalist.replaceChildren(fragment);
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
    masterCategory: state.filters.masterCategory,
    category: state.filters.category,
    articleNo: state.filters.articleNo,
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
  const masterCategories = state.data.scope?.masterCategories?.length
    ? state.data.scope.masterCategories
    : ["COMPANY GOODS", "FRESH PRODUCE", "GENERAL MERCHANDISE", "LIFESTYLE", "LOOSE COMMODITY", "PACKED COMMODITY"];
  addSelectOptions(dom.masterCategoryFilter, masterCategories, "Select all", state.filters.masterCategory);
  addSelectOptions(dom.categoryFilter, state.data.categoryOptions.map(row => row.Category), "All categories", state.filters.category);

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
      state.outletSuggestions = state.data.outletOptions
        .filter(row => row.OutletCode)
        .sort((a, b) => normalizeOutletCode(a.OutletCode).localeCompare(normalizeOutletCode(b.OutletCode), undefined, { numeric: true }))
        .map(row => {
          const code = normalizeOutletCode(row.OutletCode);
          const organization = organizationForCode(code);
          return `${code} — ${organization?.OutletName || row.Outlet || "Unnamed outlet"}`;
        });
    }
  } else {
    addSelectOptions(dom.regionFilter, state.data.outletOptions.map(row => row.Region), "All divisions", state.filters.region);
    state.rhoSuggestions = [];
    state.zonalSuggestions = [];
    fillDatalist(dom.rhoOptions, []);
    fillDatalist(dom.zonalOptions, []);
    dom.rhoFilter.value = "";
    dom.zonalFilter.value = "";
    dom.rhoFilter.disabled = true;
    dom.zonalFilter.disabled = true;
    state.outletSuggestions = state.data.outletOptions
      .filter(row => row.OutletCode)
      .map(row => `${row.OutletCode} — ${row.Outlet || "Unnamed outlet"}`);
  }
  fillDatalist(dom.outletOptions, state.outletSuggestions);

  const seenArticles = new Set();
  state.articleSuggestions = state.data.articleOptions
    .filter(row => row.ArticleNo && !seenArticles.has(String(row.ArticleNo)) && seenArticles.add(String(row.ArticleNo)))
    .map(row => `${row.ArticleNo} — ${row.ArticleName || "Unnamed article"}`);
  fillDatalist(dom.articleOptions, state.articleSuggestions);

  dom.outletFilter.value = state.filters.outletCode === "all" ? "" : canonicalOutletLabel(state.filters.outletCode);
  dom.articleFilter.value = state.filters.articleNo === "all" ? "" : canonicalArticleLabel(state.filters.articleNo);
  updateActiveFilterSummary();
}

function activeFilterLabels() {
  const labels = [`${state.filters.days} days`];
  if (state.filters.masterCategory !== "all") labels.push(state.filters.masterCategory);
  if (state.filters.category !== "all") labels.push(state.filters.category);
  if (state.filters.region !== "all") labels.push(state.filters.region);
  if (state.filters.rho !== "all") labels.push(`RHO: ${state.filters.rho}`);
  if (state.filters.zonal !== "all") labels.push(`Zonal: ${state.filters.zonal}`);
  if (state.filters.outletCode !== "all") labels.push(`Outlet: ${state.filters.outletCode}`);
  if (state.filters.articleNo !== "all") labels.push(`Article: ${state.filters.articleNo}`);
  return labels;
}

function updateActiveFilterSummary() {
  const labels = activeFilterLabels();
  dom.activeFilterCount.textContent = String(labels.length);
  dom.activeFilterSummary.textContent = labels.length ? labels.join(" · ") : "All data";
}

function setLoading(loading) {
  dom.main.setAttribute("aria-busy", String(loading));
  dom.loadingBar.classList.toggle("is-active", loading);
  dom.refreshButton.disabled = loading;
  dom.refreshButton.classList.toggle("is-refreshing", loading);
  [dom.periodFilter, dom.masterCategoryFilter, dom.categoryFilter, dom.regionFilter, dom.rhoFilter, dom.zonalFilter, dom.outletFilter, dom.articleFilter, dom.resetButton]
    .forEach(node => { node.disabled = loading; });

  if (loading) {
    dom.statusDot.className = "status-dot is-loading";
    dom.connectionStatus.textContent = state.data ? "Applying filters" : "Connecting to live data";
    dom.sourceFreshness.textContent = state.data ? "Keeping the current view visible…" : "Reading Power BI and organization mapping…";
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
    dom.organizationStatus.textContent = `${organization.rows.length.toLocaleString()} outlet mappings · updated ${dhakaDateTime(organization.updatedAt)}`;
    if (state.data) renderAll(state.data);
    return organization;
  } catch (error) {
    console.warn("Organization mapping unavailable", error);
    if (!state.organization) {
      dom.organizationStatus.classList.add("is-warning");
      dom.organizationStatus.textContent = "Zonal/RHO mapping unavailable · Power BI data remains active";
    }
    return state.organization;
  }
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

function renderKpis(kpi, data) {
  const received = finite(kpi.Receiving);
  const sales = finite(kpi.Sales);
  const gap = received != null && sales != null ? received - sales : null;
  const ratio = sales ? (received / sales) * 100 : null;

  setKpi("kpi-receiving", received);
  setKpi("kpi-sales", sales);
  setKpi("kpi-gap", gap, signedCompact);
  setKpi("kpi-inventory", kpi.Inventory);
  setKpi("kpi-over-value", kpi.OverValue, bdt);
  setKpi("kpi-outlets", kpi.ActiveOutlets, exact);
  setText("kpi-receiving-note", `${data.range.days}-day live total · click for detail`);
  setText("kpi-sales-note", `Invoiced sales · click for detail`);
  setText("kpi-gap-note", ratio == null ? "Received minus sold" : `Receipts equal ${percentage(ratio)} of sales`);
  setText("kpi-inventory-note", "Source measure · click for detail");
  setText("kpi-over-value-note", `${exact(kpi.OverIncidents)} incidents · click for detail`);
  setText("kpi-outlets-note", "Click to open the outlet list");

  const gapNode = el("kpi-gap");
  gapNode.classList.toggle("is-positive", (gap ?? 0) > 0);
  gapNode.classList.toggle("is-negative", (gap ?? 0) < 0);

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
  Region: "division", Outlet: "outlet", RHO: "RHO", Zonal: "Zonal", ArticleNo: "article code", ArticleName: "article name",
  MasterCategory: "business division", Category: "category", Receiving: "received", Sales: "sold", Gap: "balance",
  Inventory: "inventory", StockDay: "stock days", OverValue: "over value", OverIncidents: "over incidents",
  UnderIncidents: "under incidents", Incidents: "incidents", Position: "position",
};

const detailFieldDefinitions = {
  ArticleNo: { label: "Article code", numeric: false, format: value => value ?? "" },
  ArticleName: { label: "Article name", numeric: false, format: value => value ?? "" },
  MasterCategory: { label: "Business division", numeric: false, format: value => value ?? "" },
  Category: { label: "Category", numeric: false, format: value => value ?? "" },
  Receiving: { label: "Received", numeric: true, format: exact },
  Sales: { label: "Sold", numeric: true, format: exact },
  Gap: { label: "Balance", numeric: true, format: exact },
  Inventory: { label: "Inventory", numeric: true, format: exact },
  StockDay: { label: "Stock days", numeric: true, format: value => finite(value) == null ? "" : Number(value).toFixed(1) },
  OverValue: { label: "Over value", numeric: true, format: value => finite(value) == null ? "" : Number(value).toFixed(2) },
  OverIncidents: { label: "Over incidents", numeric: true, format: exact },
  UnderIncidents: { label: "Under incidents", numeric: true, format: exact },
  Incidents: { label: "Total incidents", numeric: true, format: exact },
};

function columnsForDetailMetric(metric) {
  const metricFields = metric === "Gap"
    ? ["Receiving", "Sales", "Gap"]
    : metric === "Incidents"
      ? ["OverIncidents", "UnderIncidents", "Incidents"]
      : [metric];
  return ["ArticleNo", "ArticleName", "MasterCategory", "Category", ...metricFields]
    .map(key => ({ key, ...detailFieldDefinitions[key] }))
    .filter(column => column.label);
}

function csvColumns(table) {
  if (table === "region") return [
    { key: "Region", label: "Division", value: row => plainRegion(row) },
    { key: "Receiving", label: "Received" }, { key: "Sales", label: "Sold" },
    { key: "Gap", label: "Balance" }, { key: "OverValue", label: "Over value" },
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
  if (button) button.disabled = !hasRows;
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
  if (status) status.textContent = `Sorted by ${criteria.map(item => `${sortLabels[item.key] || item.key} ${item.direction === "asc" ? "↑" : "↓"}`).join(", then ")}`;
}

function drillText(display, metric, contextType, contextValue, contextLabel, extraClass = "") {
  return `<button class="text-link ${extraClass}" type="button" data-drill-metric="${escapeHtml(metric)}" data-context-type="${escapeHtml(contextType)}" data-context-value="${escapeHtml(contextValue)}" data-context-label="${escapeHtml(contextLabel)}" title="Open details for ${escapeHtml(contextLabel)}">${escapeHtml(display)}</button>`;
}

function drillNumber(display, rawValue, metric, contextType, contextValue, contextLabel, extraClass = "") {
  if (finite(rawValue) == null) return "—";
  return `<button class="number-link ${extraClass}" type="button" data-drill-metric="${escapeHtml(metric)}" data-context-type="${escapeHtml(contextType)}" data-context-value="${escapeHtml(contextValue)}" data-context-label="${escapeHtml(contextLabel)}" title="Open ${escapeHtml(sortLabels[metric] || metric)} details">${escapeHtml(display)}</button>`;
}

function renderRegions(rows) {
  const values = rows.map(row => {
    const receiving = finite(row.Receiving) ?? 0;
    const sales = finite(row.Sales) ?? 0;
    const Gap = receiving - sales;
    const Incidents = (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0);
    const Position = positionForGap(Gap).label;
    return { ...row, Receiving: receiving, Sales: sales, Gap, Incidents, Position };
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
      <td class="numeric">${drillNumber(bdt(row.OverValue), row.OverValue, "OverValue", "region", contextValue, label)}</td>
      <td class="numeric">${drillNumber(exact(row.Incidents), row.Incidents, "Incidents", "region", contextValue, label)}</td>
      <td>${drillText(position.label, "Gap", "region", contextValue, label, `position-pill ${position.className}`)}</td></tr>`;
  }).join("");
}

function largestBy(rows, field) {
  return [...rows].filter(row => row && finite(row[field]) != null).sort((a, b) => finite(b[field]) - finite(a[field]))[0] || null;
}

function renderSignals(data) {
  const topCategory = largestBy(data.categories, "OverValue");
  const regions = data.regions.map(row => ({ ...row, TotalIncidents: (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0) }));
  const topRegion = largestBy(regions, "TotalIncidents");
  const topOutlet = largestBy(data.enrichedOutlets.filter(row => row.OutletCode), "OverValue");
  const signals = [
    topCategory && { title: `${topCategory.Category} is the largest value exposure`, detail: `${bdt(topCategory.OverValue)} in the category context; this Power BI value measure is non-additive across rows.` },
    topRegion && { title: `${plainRegion(topRegion)} has the most incidents`, detail: `${exact(topRegion.TotalIncidents)} combined over- and under-receiving incidents in the selected window.` },
    topOutlet && { title: `${plainOutlet(topOutlet)} leads the outlet action queue`, detail: `${bdt(topOutlet.OverValue)} over-receiving value under ${topOutlet.RHO} / ${topOutlet.Zonal}.`, link: true },
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
  updateCascadingOptions();
  renderPulse(kpi, data);
  renderKpis(kpi, data);
  renderTrend(data.trend || []);
  renderCategoryChart(data.categories || []);
  renderRegions(data.regions || []);
  renderSignals(data);
  renderOutlets();
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
    OverIncidents: { title: "Over-receiving incidents", field: "OverIncidents", formatter: exact },
    UnderIncidents: { title: "Under-receiving incidents", field: "UnderIncidents", formatter: exact },
    Incidents: { title: "Receiving incidents", field: "Incidents", formatter: exact },
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

function detailCacheKey(filters, metric) {
  return JSON.stringify({ metric, filters });
}

function prepareDetailRows(rows) {
  return rows
    .filter(row => row.ArticleNo)
    .map(row => ({
      ...row,
      Gap: finite(row.Receiving) != null || finite(row.Sales) != null
        ? (finite(row.Receiving) ?? 0) - (finite(row.Sales) ?? 0)
        : null,
      Incidents: finite(row.OverIncidents) != null || finite(row.UnderIncidents) != null
        ? (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0)
        : null,
    }));
}

async function openDrill(metric, context = null) {
  if (!state.data) return;
  if (metric === "ActiveOutlets") {
    state.exceptionFocus = "stock";
    state.sorts.outlet = { key: "StockDay", direction: "desc" };
    switchView("exceptions");
    renderOutlets();
    dom.exceptionsView.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  const sequence = ++state.detailSequence;
  const definition = metricDefinition(metric);
  state.detailMetric = definition.field;
  state.detailColumns = columnsForDetailMetric(definition.field);
  state.detailContext = context;
  state.detailSearch = "";
  dom.detailSearch.value = "";
  state.sorts.detail = { key: definition.field, direction: "desc" };
  setText("detail-title", `${definition.title} details`);
  const contextText = context?.label ? context.label : activeFilterLabels().join(" · ");
  setText("detail-context", `${contextText} · ${dateRangeLabel(state.data.range)}`);
  dom.detailLoading.hidden = false;
  dom.detailError.hidden = true;
  dom.detailContent.hidden = true;
  dom.selectedArticle.hidden = true;
  openDialog();

  try {
    const filters = detailFiltersForContext(context);
    const cacheKey = detailCacheKey(filters, definition.field);
    const cached = state.detailCache.get(cacheKey);
    if (cached && Date.now() - cached.savedAt < DETAIL_CACHE_MS) {
      state.detailRows = cached.rows;
      dom.detailLoading.hidden = true;
      dom.detailContent.hidden = false;
      renderDetail();
      return;
    }

    const result = await state.client.loadArticleDetails(filters, definition.field);
    if (sequence !== state.detailSequence) return;
    state.detailRows = prepareDetailRows(result.rows);
    state.detailCache.set(cacheKey, { savedAt: Date.now(), rows: state.detailRows });
    dom.detailLoading.hidden = true;
    dom.detailContent.hidden = false;
    renderDetail();
  } catch (error) {
    if (sequence !== state.detailSequence) return;
    dom.detailLoading.hidden = true;
    dom.detailError.hidden = false;
    dom.detailError.textContent = `${error?.message || "Article details could not be loaded."} Close this panel and try again.`;
  }
}

function detailScopeLabel() {
  if (state.detailContext?.label) return state.detailContext.label;
  if (state.filters.outletCode !== "all") return canonicalOutletLabel(state.filters.outletCode);
  if (state.filters.zonal !== "all") return state.filters.zonal;
  if (state.filters.rho !== "all") return state.filters.rho;
  if (state.filters.region !== "all") return state.filters.region;
  return "Current dashboard scope";
}

function renderDetailSummary(totalRows) {
  const definition = metricDefinition(state.detailMetric);
  const contextOrganization = state.detailContext?.type === "outlet" ? organizationForCode(state.detailContext.value) : null;
  const cards = [
    ["Article rows", exact(totalRows)],
    ["Data window", dateRangeLabel(state.data.range)],
    ["Selected scope", detailScopeLabel()],
    ["RHO / Zonal", contextOrganization ? `${contextOrganization.RHO} / ${contextOrganization.Zonal}` : `Sorted by ${definition.title}`],
  ];
  dom.detailSummary.innerHTML = cards.map(([label, value]) => `<div class="detail-summary-card"><span>${escapeHtml(label)}</span><strong title="${escapeHtml(value)}">${escapeHtml(value)}</strong></div>`).join("");
}

function detailNumber(row, metric, display) {
  if (finite(row[metric]) == null) return "—";
  return `<button class="number-link" type="button" data-detail-article="${escapeHtml(row.ArticleNo)}" title="Show this article record">${escapeHtml(display)}</button>`;
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

function renderDetailHeader() {
  dom.detailHead.innerHTML = state.detailColumns.map(column => `<th scope="col"${column.numeric ? ' class="numeric"' : ""}><button class="sort-button" type="button" data-sort-table="detail" data-sort-key="${escapeHtml(column.key)}">${escapeHtml(column.label)} <span></span></button></th>`).join("");
  dom.detailHead.querySelectorAll(".sort-button").forEach(button => {
    button.addEventListener("click", event => applySort("detail", button.dataset.sortKey, event.shiftKey));
  });
}

function detailDisplay(row, column) {
  const value = row[column.key];
  if (!column.numeric) return `<button class="text-link" type="button" data-detail-article="${escapeHtml(row.ArticleNo)}" title="Show this article record">${escapeHtml(value || (column.key === "ArticleName" ? "Unnamed article" : "—"))}</button>`;
  const display = column.key === "Gap" ? signedCompact(value)
    : column.key === "OverValue" ? bdt(value)
      : column.key === "StockDay" ? (finite(value) == null ? "—" : Number(value).toFixed(1))
        : compact(value);
  return detailNumber(row, column.key, display);
}

function renderDetail() {
  const query = state.detailSearch.trim().toLocaleLowerCase();
  const filtered = state.detailRows.filter(row => !query || [row.ArticleNo, row.ArticleName, row.MasterCategory, row.Category, detailScopeLabel()]
    .some(value => String(value ?? "").toLocaleLowerCase().includes(query)));
  const sorted = sortRows(filtered, "detail");
  const visible = sorted.slice(0, DETAIL_ROW_LIMIT);
  state.visibleRows.detail = visible;
  setExportAvailability("detail", visible.length > 0);
  renderDetailHeader();
  updateSortIndicators("detail");
  renderDetailSummary(state.detailRows.length);
  setText("detail-result-count", visible.length < sorted.length ? `Showing ${exact(visible.length)} of ${exact(sorted.length)} matches` : `${exact(sorted.length)} article rows`);

  if (!visible.length) {
    dom.detailTable.innerHTML = `<tr><td colspan="${state.detailColumns.length}" class="empty-cell">No article rows match the current detail search.</td></tr>`;
    return;
  }

  dom.detailTable.innerHTML = visible.map(row => `<tr>${state.detailColumns.map(column => `<td${column.numeric ? ' class="numeric"' : ""}>${detailDisplay(row, column)}</td>`).join("")}</tr>`).join("");
}

function showArticleRecord(articleNo) {
  const row = state.detailRows.find(item => String(item.ArticleNo) === String(articleNo));
  if (!row) return;
  const contextOrganization = state.detailContext?.type === "outlet" ? organizationForCode(state.detailContext.value) : null;
  const metricFields = state.detailColumns.filter(column => column.numeric).map(column => [column.label, column.key === "OverValue" ? bdt(row[column.key]) : column.key === "Gap" ? signedCompact(row[column.key]) : column.key === "StockDay" ? (finite(row[column.key]) == null ? "—" : Number(row[column.key]).toFixed(1)) : exact(row[column.key])]);
  const fields = [
    ["Article", `${row.ArticleNo} — ${row.ArticleName || "Unnamed"}`],
    ["Selected scope", detailScopeLabel()],
    ["RHO / Zonal", contextOrganization ? `${contextOrganization.RHO} / ${contextOrganization.Zonal}` : "Current filtered scope"],
    ...metricFields,
  ];
  dom.selectedArticle.innerHTML = fields.map(([label, value]) => `<div class="record-field"><span>${escapeHtml(label)}</span><strong title="${escapeHtml(value)}">${escapeHtml(value)}</strong></div>`).join("");
  dom.selectedArticle.hidden = false;
}

async function loadDashboard({ refreshMetadata = false } = {}) {
  const sequence = ++state.loadSequence;
  setLoading(true);
  clearError();
  if (!navigator.onLine) {
    setLoading(false);
    showError(new Error("This device is offline."));
    return;
  }

  try {
    if (refreshMetadata) {
      state.client = new PowerBIDataClient();
      state.detailCache.clear();
    }
    const organizationPromise = refreshMetadata || !state.organizationPromise
      ? (state.organizationPromise = refreshOrganization())
      : state.organizationPromise;
    const requestFilters = buildPowerBIFilters();
    const dataPromise = state.client.load(requestFilters, { section: "core" });
    const [data] = await Promise.all([dataPromise, organizationPromise]);
    if (sequence !== state.loadSequence) return;
    state.data = data;
    state.nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    renderAll(data);
    dom.statusDot.className = "status-dot";
    dom.connectionStatus.textContent = "Live Power BI data";
    dom.sourceFreshness.textContent = `Overview ready · loading outlet and search data…`;
    setLoading(false);

    state.client.load(requestFilters, { section: "supporting" }).then(supporting => {
      if (sequence !== state.loadSequence) return;
      state.data = { ...state.data, ...supporting };
      saveDashboardCache(state.data);
      renderAll(state.data);
      dom.sourceFreshness.textContent = `Model refreshed ${dhakaDateTime(state.data.sourceTimestamp)}`;
    }).catch(error => {
      if (sequence !== state.loadSequence) return;
      console.warn("Supporting dashboard data could not be loaded", error);
      dom.sourceFreshness.textContent = `Overview ready · outlet/search data unavailable; refresh to retry`;
    });
  } catch (error) {
    if (sequence !== state.loadSequence) return;
    console.error("Dashboard refresh failed", error);
    showError(error);
  } finally {
    if (sequence === state.loadSequence) setLoading(false);
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
  if (!match) {
    dom.cascadeNote.textContent = `No exact ${type} match. Select a suggestion or enter an exact code.`;
    input.setAttribute("aria-invalid", "true");
    return;
  }

  input.removeAttribute("aria-invalid");
  state.filters[key] = match.split("—")[0].trim();
  input.value = match;
  dom.cascadeNote.textContent = "Applying selection; all related filter options will shorten automatically.";
  loadDashboard();
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
    state.exceptionFocus = button.dataset.focus;
    const field = focusConfig[state.exceptionFocus].field;
    state.sorts.outlet = { key: field, direction: "desc" };
    renderOutlets();
  });
});

document.querySelectorAll(".sort-button").forEach(button => {
  button.addEventListener("click", event => applySort(button.dataset.sortTable, button.dataset.sortKey, event.shiftKey));
});

document.querySelectorAll("[data-export-table]").forEach(button => {
  button.addEventListener("click", () => exportVisibleCsv(button.dataset.exportTable));
});

dom.main.addEventListener("click", event => {
  const drill = event.target.closest("[data-drill-metric]");
  if (!drill) return;
  openDrill(drill.dataset.drillMetric, drill.dataset.contextType ? { type: drill.dataset.contextType, value: drill.dataset.contextValue, label: drill.dataset.contextLabel } : null);
});

dom.main.addEventListener("keydown", event => {
  if (!["Enter", " "].includes(event.key)) return;
  const drill = event.target.closest("[data-drill-metric]");
  if (!drill) return;
  event.preventDefault();
  openDrill(drill.dataset.drillMetric, drill.dataset.contextType ? { type: drill.dataset.contextType, value: drill.dataset.contextValue, label: drill.dataset.contextLabel } : null);
});

dom.detailDialog.addEventListener("click", event => {
  const article = event.target.closest("[data-detail-article]");
  if (article) showArticleRecord(article.dataset.detailArticle);
});

dom.managementSignals.addEventListener("click", event => {
  const link = event.target.closest("[data-open-exceptions]");
  if (!link) return;
  event.preventDefault();
  switchView("exceptions");
  dom.exceptionsView.scrollIntoView({ behavior: "smooth", block: "start" });
});

dom.outletSearch.addEventListener("input", event => { state.outletSearch = event.target.value; renderOutlets(); });
dom.detailSearch.addEventListener("input", event => { state.detailSearch = event.target.value; renderDetail(); });

dom.periodFilter.addEventListener("change", event => { state.filters.days = Number(event.target.value); loadDashboard(); });
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
    if (!input.value && state.filters[type === "outlet" ? "outletCode" : "articleNo"] !== "all") applySearchSelection(type);
    else dom.cascadeNote.textContent = `Type an exact ${type} code/name, then press Enter or choose a suggestion.`;
  });
});

dom.resetButton.addEventListener("click", () => {
  state.filters = { ...DEFAULT_FILTERS };
  dom.periodFilter.value = "30";
  dom.masterCategoryFilter.value = "all";
  dom.categoryFilter.value = "all";
  dom.regionFilter.value = "all";
  dom.rhoFilter.value = "";
  dom.zonalFilter.value = "";
  dom.outletFilter.value = "";
  dom.articleFilter.value = "";
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
el("detail-close").addEventListener("click", () => dom.detailDialog.close());

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
