import { PowerBIDataClient, POWER_BI_URL } from "./powerbi.js";

const AUTO_REFRESH_MS = 15 * 60 * 1000;
const nf = new Intl.NumberFormat("en-GB", { maximumFractionDigits: 0 });
const compactNf = new Intl.NumberFormat("en-GB", { notation: "compact", maximumFractionDigits: 1 });
const percentNf = new Intl.NumberFormat("en-GB", { minimumFractionDigits: 1, maximumFractionDigits: 1 });

const state = {
  client: new PowerBIDataClient(),
  data: null,
  filters: { days: 30, region: "all", category: "all" },
  filterOptions: { regions: [], categories: [] },
  activeView: "overview",
  exceptionFocus: "over",
  outletSearch: "",
  loadSequence: 0,
  nextRefreshAt: 0,
};

const el = id => document.getElementById(id);

const dom = {
  main: el("main-content"),
  loadingBar: el("loading-bar"),
  refreshButton: el("refresh-dashboard"),
  retryButton: el("retry-dashboard"),
  resetButton: el("reset-filters"),
  periodFilter: el("period-filter"),
  regionFilter: el("region-filter"),
  categoryFilter: el("category-filter"),
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

function signedBdt(value) {
  const number = finite(value);
  if (number == null || number === 0) return bdt(number);
  return `${number > 0 ? "+" : "−"}${bdt(Math.abs(number))}`;
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
  return date
    ? new Intl.DateTimeFormat("en-GB", { day: "numeric", month: "short", timeZone: "UTC" }).format(date)
    : "";
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

function plainOutlet(row) {
  return row?.Outlet || row?.OutletCode || "Unassigned / HO";
}

function plainRegion(row) {
  return row?.Region || "Unassigned / HO";
}

function setText(id, value, title = "") {
  const node = el(id);
  node.textContent = value;
  if (title) node.title = title;
  else node.removeAttribute("title");
}

function populateSelect(select, values, firstLabel) {
  const selected = select.value;
  select.replaceChildren(new Option(firstLabel, "all"));
  values.forEach(value => select.add(new Option(value, value)));
  select.value = values.includes(selected) || selected === "all" ? selected : "all";
}

function setLoading(loading) {
  dom.main.setAttribute("aria-busy", String(loading));
  dom.loadingBar.classList.toggle("is-active", loading);
  dom.refreshButton.disabled = loading;
  dom.refreshButton.classList.toggle("is-refreshing", loading);
  [dom.periodFilter, dom.regionFilter, dom.categoryFilter, dom.resetButton].forEach(node => { node.disabled = loading; });

  if (loading) {
    dom.statusDot.className = "status-dot is-loading";
    dom.connectionStatus.textContent = state.data ? "Refreshing live data" : "Connecting to live data";
    dom.sourceFreshness.textContent = state.data ? "Keeping the current view visible…" : "Reading the published Power BI model…";
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

function updateFilterOptions(data) {
  if (!state.filterOptions.categories.length) {
    state.filterOptions.categories = [...new Set(data.categories.map(row => row.Category).filter(Boolean))]
      .sort((a, b) => String(a).localeCompare(String(b)));
  }
  if (!state.filterOptions.regions.length) {
    state.filterOptions.regions = [...new Set(data.regions.map(row => row.Region).filter(Boolean))]
      .sort((a, b) => String(a).localeCompare(String(b)));
  }
  populateSelect(dom.categoryFilter, state.filterOptions.categories, "All categories");
  populateSelect(dom.regionFilter, state.filterOptions.regions, "All divisions");
  dom.categoryFilter.value = state.filters.category;
  dom.regionFilter.value = state.filters.region;
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
    setText("balance-detail", sales
      ? `Receipts ran ${percentage(Math.abs(gapPct))} above sales in this data window.`
      : "Receipts were recorded while invoiced sales were zero in this data window.");
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

  const scope = [state.filters.category === "all" ? data.scope.masterCategory : state.filters.category];
  if (state.filters.region !== "all") scope.push(state.filters.region);
  scope.push(`${data.range.days} complete days`);
  setText("report-context", scope.join(" · "));
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
  setText("kpi-receiving-note", `${data.range.days}-day live total`);
  setText("kpi-sales-note", `Invoiced sales · ${data.range.days} days`);
  setText("kpi-gap-note", ratio == null ? "Received minus sold" : `Receipts equal ${percentage(ratio)} of sales`);
  setText("kpi-inventory-note", "Source measure · current scope");
  setText("kpi-over-value-note", `${exact(kpi.OverIncidents)} over-receiving incidents`);
  setText("kpi-outlets-note", "Outlets with latest stock above zero");

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
  const xIndexes = [...new Set(Array.from(
    { length: xTickCount },
    (_, index) => Math.round(index * (values.length - 1) / Math.max(1, xTickCount - 1))
  ))];

  const grid = yTicks.map(value => `
    <line class="chart-gridline" x1="${margin.left}" y1="${y(value)}" x2="${width - margin.right}" y2="${y(value)}" />
    <text class="axis-label" x="${margin.left - 10}" y="${y(value) + 4}" text-anchor="end">${escapeHtml(compact(value))}</text>
  `).join("");

  const xLabels = xIndexes.map(index => `
    <text class="axis-label" x="${x(index)}" y="${height - 14}" text-anchor="middle">${escapeHtml(dateTick(values[index]._date))}</text>
  `).join("");

  const pointTargets = values.map((row, index) => `
    <circle cx="${x(index)}" cy="${y(row.Receiving)}" r="9" fill="transparent"><title>${escapeHtml(`${dateTick(row._date)} · Receiving ${exact(row.Receiving)}`)}</title></circle>
    <circle cx="${x(index)}" cy="${y(row.Sales)}" r="9" fill="transparent"><title>${escapeHtml(`${dateTick(row._date)} · Sales ${exact(row.Sales)}`)}</title></circle>
  `).join("");

  dom.trendChart.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Daily line chart comparing received units with sold units">
      <defs>
        <linearGradient id="receiving-area" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#0b7b86" stop-opacity=".18"/><stop offset="1" stop-color="#0b7b86" stop-opacity="0"/></linearGradient>
      </defs>
      ${grid}
      <line class="chart-axis" x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}" />
      ${xLabels}
      <path d="${path("Receiving")} L${x(values.length - 1)},${height - margin.bottom} L${x(0)},${height - margin.bottom} Z" fill="url(#receiving-area)" />
      <path d="${path("Receiving")}" fill="none" stroke="#0b7b86" stroke-width="3.5" stroke-linejoin="round" stroke-linecap="round" />
      <path d="${path("Sales")}" fill="none" stroke="#657b8e" stroke-width="3" stroke-dasharray="7 6" stroke-linejoin="round" stroke-linecap="round" />
      ${pointTargets}
    </svg>`;
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
    return `
      <text x="8" y="${y + 5}" fill="#35465a" font-size="12.5" font-weight="750">${escapeHtml(row.Category)}</text>
      <rect x="${x}" y="${y - 10}" width="${Math.max(1.5, barWidth)}" height="20" rx="4" fill="${positive ? "#e56b55" : "#198693"}"><title>${escapeHtml(title)}</title></rect>
      <text x="${width - 8}" y="${y + 5}" text-anchor="end" fill="${positive ? "#b64b39" : "#08717b"}" font-size="12.5" font-weight="800">${escapeHtml(signedCompact(row.Gap))}</text>`;
  }).join("");

  dom.categoryChart.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Diverging bar chart of received units minus sold units by category">
      <text x="${chartStart}" y="17" fill="#718196" font-size="11.5" font-weight="700">Sales-led</text>
      <text x="${chartEnd}" y="17" text-anchor="end" fill="#718196" font-size="11.5" font-weight="700">Receipt-heavy</text>
      <line class="chart-zero" x1="${centre}" y1="27" x2="${centre}" y2="${height - 12}" />
      ${bars}
    </svg>`;
}

function positionForGap(gap) {
  if (gap > 0) return { label: "Receipt-heavy", className: "receipt-heavy" };
  if (gap < 0) return { label: "Sales-led", className: "sales-led" };
  return { label: "Balanced", className: "balanced" };
}

function renderRegions(rows) {
  const values = [...rows].sort((a, b) => (finite(b.OverValue) ?? 0) - (finite(a.OverValue) ?? 0));
  if (!values.length) {
    dom.regionTable.innerHTML = '<tr><td colspan="7" class="empty-cell">No division data is available for this selection.</td></tr>';
    return;
  }

  dom.regionTable.innerHTML = values.map(row => {
    const received = finite(row.Receiving) ?? 0;
    const sales = finite(row.Sales) ?? 0;
    const gap = received - sales;
    const incidents = (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0);
    const position = positionForGap(gap);
    return `<tr>
      <td>${escapeHtml(plainRegion(row))}</td>
      <td class="numeric" title="${escapeHtml(exact(received))}">${escapeHtml(compact(received))}</td>
      <td class="numeric" title="${escapeHtml(exact(sales))}">${escapeHtml(compact(sales))}</td>
      <td class="numeric ${gap > 0 ? "positive-number" : gap < 0 ? "negative-number" : ""}" title="${escapeHtml(exact(gap))}">${escapeHtml(signedCompact(gap))}</td>
      <td class="numeric" title="৳${escapeHtml(exact(row.OverValue))}">${escapeHtml(bdt(row.OverValue))}</td>
      <td class="numeric">${escapeHtml(exact(incidents))}</td>
      <td><span class="position-pill ${position.className}">${position.label}</span></td>
    </tr>`;
  }).join("");
}

function largestBy(rows, field) {
  return [...rows]
    .filter(row => row && finite(row[field]) != null)
    .sort((a, b) => finite(b[field]) - finite(a[field]))[0] || null;
}

function renderSignals(data, kpi) {
  const topCategory = largestBy(data.categories, "OverValue");
  const regions = data.regions.map(row => ({ ...row, TotalIncidents: (finite(row.OverIncidents) ?? 0) + (finite(row.UnderIncidents) ?? 0) }));
  const topRegion = largestBy(regions, "TotalIncidents");
  const usableOutlets = data.outlets.filter(row => row.Outlet || row.OutletCode);
  const topOutlet = largestBy(usableOutlets, "OverValue");
  const signals = [
    topCategory && {
      title: `${topCategory.Category} is the largest value exposure`,
      detail: `${bdt(topCategory.OverValue)} in the category context; this Power BI value measure is non-additive across rows.`,
    },
    topRegion && {
      title: `${plainRegion(topRegion)} has the most incidents`,
      detail: `${exact(topRegion.TotalIncidents)} combined over- and under-receiving incidents in the selected window.`,
    },
    topOutlet && {
      title: `${plainOutlet(topOutlet)} leads the outlet action queue`,
      detail: `${bdt(topOutlet.OverValue)} over-receiving value. Open the outlet view to see the ranked list.`,
      link: true,
    },
  ].filter(Boolean);

  if (!signals.length) {
    dom.managementSignals.innerHTML = '<li><strong>No management signals available</strong><span>Try a broader filter selection.</span></li>';
    return;
  }

  dom.managementSignals.innerHTML = signals.map(signal => `<li><strong>${escapeHtml(signal.title)}</strong><span>${escapeHtml(signal.detail)}${signal.link ? ' <a href="#exceptions-view" data-open-exceptions>Open outlet queue →</a>' : ""}</span></li>`).join("");
}

const focusConfig = {
  over: {
    field: "OverValue",
    title: "Top over-receiving outlets",
    description: "Outlets ranked by over-receiving value for the selected scope.",
    value: row => bdt(row.OverValue),
  },
  under: {
    field: "UnderIncidents",
    title: "Top under-receiving outlets",
    description: "Outlets ranked by under-receiving incident count for the selected scope.",
    value: row => `${exact(row.UnderIncidents)} incidents`,
  },
  stock: {
    field: "StockDay",
    title: "Highest stock-cover outlets",
    description: "Outlets ranked by the source-calculated stock-day measure.",
    value: row => finite(row.StockDay) == null ? "—" : `${Number(row.StockDay).toFixed(1)} days`,
  },
};

function filteredRankedOutlets() {
  const config = focusConfig[state.exceptionFocus];
  const query = state.outletSearch.trim().toLocaleLowerCase();
  return state.data.outlets
    .filter(row => (finite(row[config.field]) ?? 0) > 0)
    .filter(row => !query || [plainOutlet(row), row.OutletCode, plainRegion(row)].some(value => String(value ?? "").toLocaleLowerCase().includes(query)))
    .sort((a, b) => (finite(b[config.field]) ?? -Infinity) - (finite(a[config.field]) ?? -Infinity));
}

function renderExceptionSummary(rows) {
  const config = focusConfig[state.exceptionFocus];
  const top = rows.slice(0, 3);
  if (!top.length) {
    dom.exceptionSummary.innerHTML = '<div class="empty-chart">No outlets match this search.</div>';
    return;
  }
  dom.exceptionSummary.innerHTML = top.map((row, index) => `<article class="exception-card">
    <span class="exception-rank">0${index + 1}</span>
    <span>${escapeHtml(plainRegion(row))}</span>
    <strong title="${escapeHtml(plainOutlet(row))}">${escapeHtml(plainOutlet(row))}</strong>
    <b>${escapeHtml(config.value(row))}</b>
  </article>`).join("");
}

function renderOutlets() {
  if (!state.data) return;
  const config = focusConfig[state.exceptionFocus];
  const rows = filteredRankedOutlets();
  const visible = rows.slice(0, 25);
  setText("exceptions-description", config.description);
  setText("outlet-table-title", config.title);
  setText("outlet-table-note", `Showing the first ${Math.min(25, rows.length)} ranked outlets for the selected scope.`);
  setText("outlet-result-count", `${exact(rows.length)} matching outlets`);
  setText("exception-count", compact(rows.length));
  renderExceptionSummary(rows);

  document.querySelectorAll(".focus-button").forEach(button => {
    const active = button.dataset.focus === state.exceptionFocus;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });

  if (!visible.length) {
    dom.outletTable.innerHTML = '<tr><td colspan="10" class="empty-cell">No outlets match the current search and filters.</td></tr>';
    return;
  }

  dom.outletTable.innerHTML = visible.map((row, index) => {
    const receiving = finite(row.Receiving) ?? 0;
    const sales = finite(row.Sales) ?? 0;
    const gap = receiving - sales;
    return `<tr>
      <td>${index + 1}</td>
      <td><span class="outlet-name">${escapeHtml(plainOutlet(row))}<small>${row.OutletCode ? `Code ${escapeHtml(row.OutletCode)}` : "No outlet code"}</small></span></td>
      <td>${escapeHtml(plainRegion(row))}</td>
      <td class="numeric" title="${escapeHtml(exact(receiving))}">${escapeHtml(compact(receiving))}</td>
      <td class="numeric" title="${escapeHtml(exact(sales))}">${escapeHtml(compact(sales))}</td>
      <td class="numeric ${gap > 0 ? "positive-number" : gap < 0 ? "negative-number" : ""}" title="${escapeHtml(exact(gap))}">${escapeHtml(signedCompact(gap))}</td>
      <td class="numeric">${finite(row.StockDay) == null ? "—" : Number(row.StockDay).toFixed(1)}</td>
      <td class="numeric" title="৳${escapeHtml(exact(row.OverValue))}">${escapeHtml(bdt(row.OverValue))}</td>
      <td class="numeric">${escapeHtml(exact(row.OverIncidents))}</td>
      <td class="numeric">${escapeHtml(exact(row.UnderIncidents))}</td>
    </tr>`;
  }).join("");
}

function renderAll(data) {
  const kpi = data.kpis?.[0] || {};
  renderPulse(kpi, data);
  renderKpis(kpi, data);
  renderTrend(data.trend || []);
  renderCategoryChart(data.categories || []);
  renderRegions(data.regions || []);
  renderSignals(data, kpi);
  renderOutlets();
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
    if (refreshMetadata) state.client = new PowerBIDataClient();
    const data = await state.client.load(state.filters);
    if (sequence !== state.loadSequence) return;
    state.data = data;
    state.nextRefreshAt = Date.now() + AUTO_REFRESH_MS;
    updateFilterOptions(data);
    renderAll(data);
    dom.statusDot.className = "status-dot";
    dom.connectionStatus.textContent = "Live Power BI data";
    dom.sourceFreshness.textContent = `Model refreshed ${dhakaDateTime(data.sourceTimestamp)}`;
  } catch (error) {
    if (sequence !== state.loadSequence) return;
    console.error("Dashboard refresh failed", error);
    showError(error);
  } finally {
    if (sequence === state.loadSequence) setLoading(false);
  }
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
    renderOutlets();
  });
});

dom.managementSignals.addEventListener("click", event => {
  const link = event.target.closest("[data-open-exceptions]");
  if (!link) return;
  event.preventDefault();
  switchView("exceptions");
  dom.exceptionsView.scrollIntoView({ behavior: "smooth", block: "start" });
});

dom.outletSearch.addEventListener("input", event => {
  state.outletSearch = event.target.value;
  renderOutlets();
});

dom.periodFilter.addEventListener("change", event => {
  state.filters.days = Number(event.target.value);
  loadDashboard();
});

dom.regionFilter.addEventListener("change", event => {
  state.filters.region = event.target.value;
  loadDashboard();
});

dom.categoryFilter.addEventListener("change", event => {
  state.filters.category = event.target.value;
  loadDashboard();
});

dom.resetButton.addEventListener("click", () => {
  state.filters = { days: 30, region: "all", category: "all" };
  dom.periodFilter.value = "30";
  dom.regionFilter.value = "all";
  dom.categoryFilter.value = "all";
  dom.outletSearch.value = "";
  state.outletSearch = "";
  loadDashboard();
});

dom.refreshButton.addEventListener("click", () => loadDashboard({ refreshMetadata: true }));
dom.retryButton.addEventListener("click", () => loadDashboard({ refreshMetadata: true }));

window.addEventListener("offline", () => {
  dom.statusDot.className = "status-dot is-error";
  dom.connectionStatus.textContent = "Device is offline";
  dom.sourceFreshness.textContent = state.data ? "Showing the last successfully loaded view" : "Waiting for a connection";
});

window.addEventListener("online", () => loadDashboard({ refreshMetadata: true }));

document.addEventListener("visibilitychange", () => {
  if (!document.hidden && state.nextRefreshAt && Date.now() >= state.nextRefreshAt) {
    loadDashboard({ refreshMetadata: true });
  }
});

window.setInterval(() => {
  if (!document.hidden && navigator.onLine && state.nextRefreshAt && Date.now() >= state.nextRefreshAt) {
    loadDashboard({ refreshMetadata: true });
  }
}, 60_000);

loadDashboard({ refreshMetadata: true });
