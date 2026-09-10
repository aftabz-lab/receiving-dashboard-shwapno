const API_ROOT = "https://wabi-east-asia-b-primary-api.analysis.windows.net";
const RESOURCE_KEY = "09822098-d3f2-462b-bf91-1703f0e5e9cc";

export const POWER_BI_URL = "https://app.powerbi.com/view?r=eyJrIjoiMDk4MjIwOTgtZDNmMi00NjJiLWJmOTEtMTcwM2YwZTVlOWNjIiwidCI6IjNjZDA3OTg4LTUyNjMtNDA2NC1hZDU1LWU5NTZhYjNkZDExNyIsImMiOjEwfQ%3D%3D";

const DEFAULT_MOVEMENT_TYPES = ["101", "102", "303", "304", "305", "551", "552", "Z04", "122", "161", "162", "Z03"];
const DEFAULT_MASTER_CATEGORIES = [
  "COMPANY GOODS",
  "FRESH PRODUCE",
  "GENERAL MERCHANDISE",
  "LIFESTYLE",
  "LOOSE COMMODITY",
  "PACKED COMMODITY",
];
const MAX_QUERY_ROWS = 30000;

function uuid() {
  if (typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const bytes = crypto.getRandomValues(new Uint8Array(16));
  bytes[6] = (bytes[6] & 0x0f) | 0x40;
  bytes[8] = (bytes[8] & 0x3f) | 0x80;
  const hex = [...bytes].map(value => value.toString(16).padStart(2, "0"));
  return `${hex.slice(0, 4).join("")}-${hex.slice(4, 6).join("")}-${hex.slice(6, 8).join("")}-${hex.slice(8, 10).join("")}-${hex.slice(10).join("")}`;
}

function requestHeaders(includeContentType = false) {
  const headers = {
    Accept: "application/json",
    ActivityId: uuid(),
    RequestId: uuid(),
    "X-PowerBI-ResourceKey": RESOURCE_KEY,
  };

  if (includeContentType) headers["Content-Type"] = "application/json";
  return headers;
}

function field(source, property) {
  return { Column: { Expression: { SourceRef: { Source: source } }, Property: property } };
}

function column(source, property, name) {
  return { ...field(source, property), Name: name };
}

function sum(source, property, name) {
  return {
    Aggregation: { Expression: field(source, property), Function: 0 },
    Name: name,
  };
}

function measure(source, property, name) {
  return {
    Measure: { Expression: { SourceRef: { Source: source } }, Property: property },
    Name: name,
  };
}

function literal(value) {
  return { Literal: { Value: value } };
}

function stringLiteral(value) {
  return `'${String(value).replaceAll("'", "''")}'`;
}

function parseDateLiteral(value) {
  const match = /^datetime'(\d{4}-\d{2}-\d{2})T/.exec(value || "");
  return match?.[1] || null;
}

function shiftIsoDate(date, days) {
  const value = new Date(`${date}T00:00:00Z`);
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

function dhakaToday() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Dhaka",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}-${values.month}-${values.day}`;
}

function collectLiterals(node, output = []) {
  if (!node || typeof node !== "object") return output;
  if (node.Literal && typeof node.Literal.Value === "string") output.push(node.Literal.Value);
  for (const value of Object.values(node)) collectLiterals(value, output);
  return output;
}

function collectProperties(node, output = []) {
  if (!node || typeof node !== "object") return output;
  if (typeof node.Property === "string") output.push(node.Property);
  for (const value of Object.values(node)) collectProperties(value, output);
  return output;
}

function savedReportScope(explorationDocument) {
  const fallbackEnd = dhakaToday();
  const fallback = {
    start: shiftIsoDate(fallbackEnd, -30),
    endExclusive: fallbackEnd,
    masterCategory: "PACKED COMMODITY",
    masterCategories: DEFAULT_MASTER_CATEGORIES,
    movementTypes: DEFAULT_MOVEMENT_TYPES,
  };

  try {
    const document = JSON.parse(explorationDocument);
    const page = document.pages.pages.find(item => item.content?.displayName === "Over Receiving");
    if (!page) return fallback;

    let start = null;
    let endExclusive = null;
    let masterCategory = null;
    const masterCategories = new Set();
    let movementTypes = [];

    for (const container of page.visualContainers || []) {
      const visual = container.content?.visual;
      const filters = (visual?.objects?.general || [])
        .map(item => item?.properties?.filter?.filter)
        .filter(Boolean);

      for (const filter of filters) {
        const entities = (filter.From || []).map(item => item.Entity);
        const properties = collectProperties(filter);
        const literals = collectLiterals(filter);

        if (entities.includes("DimDate") && properties.includes("Date")) {
          const dates = literals.map(parseDateLiteral).filter(Boolean).sort();
          if (dates.length >= 2) [start, endExclusive] = [dates[0], dates.at(-1)];
        }

        if (entities.includes("DimArticle") && properties.includes("MasterCategory")) {
          const categoryValues = literals.filter(value => /^'.*'$/.test(value)).map(value => value.slice(1, -1));
          categoryValues.forEach(value => masterCategories.add(value));
          if (categoryValues.length === 1) masterCategory = categoryValues[0];
        }

        if (entities.includes("Query3") && properties.includes("movement_type")) {
          movementTypes = literals.filter(value => /^'.*'$/.test(value)).map(value => value.slice(1, -1));
        }
      }
    }

    return {
      start: start || fallback.start,
      endExclusive: endExclusive || fallback.endExclusive,
      masterCategory: masterCategory || fallback.masterCategory,
      masterCategories: [...new Set([...fallback.masterCategories, ...masterCategories])].sort(),
      movementTypes: movementTypes.length ? movementTypes : fallback.movementTypes,
    };
  } catch {
    return fallback;
  }
}

function isMasked(mask, index) {
  if (!mask) return false;
  return Math.floor(Number(mask) / (2 ** index)) % 2 === 1;
}

function numericMeasure(value, descriptor) {
  if (value == null || descriptor.Kind !== 2) return value;
  const parsed = typeof value === "string" ? Number(value) : value;
  return Number.isFinite(parsed) ? parsed : value;
}

function decodeResult(resultEntry) {
  const data = resultEntry?.result?.data;
  if (!data) {
    const message = resultEntry?.result?.error?.message || "Power BI returned an empty result.";
    throw new Error(message);
  }

  const descriptors = data.descriptor?.Select || [];
  const dataSet = data.dsr?.DS?.[0];
  if (!dataSet) return [];

  const dictionaries = dataSet.ValueDicts || {};
  const encodedRows = (dataSet.PH || []).flatMap(partition =>
    Object.entries(partition)
      .filter(([key, value]) => /^DM\d+$/.test(key) && Array.isArray(value))
      .flatMap(([, value]) => value)
  );

  let schema = descriptors.map(item => ({ N: item.Value }));
  let previous = Array(descriptors.length).fill(null);

  return encodedRows.map(encoded => {
    if (Array.isArray(encoded.S)) schema = encoded.S;
    const compressed = encoded.C || [];
    let cursor = 0;
    const values = descriptors.map((descriptor, index) => {
      const key = descriptor.Value;
      let value;

      if (Object.prototype.hasOwnProperty.call(encoded, key)) value = encoded[key];
      else if (isMasked(encoded.R, index)) value = previous[index];
      else if (isMasked(encoded["Ø"], index)) value = null;
      else value = compressed[cursor++];

      const dictionaryName = schema[index]?.DN;
      const dictionary = dictionaryName ? dictionaries[dictionaryName] : null;
      if (dictionary && Number.isInteger(value) && value >= 0 && value < dictionary.length) {
        value = dictionary[value];
      }

      return numericMeasure(value, descriptor);
    });

    previous = values;
    return Object.fromEntries(descriptors.map((descriptor, index) => [descriptor.Name, values[index]]));
  });
}

function createQuery(select, from, where, count = 1200) {
  const semanticQuery = { Version: 2, From: from, Select: select, Where: where };
  return {
    Query: {
      Commands: [{
        SemanticQueryDataShapeCommand: {
          Query: semanticQuery,
          Binding: {
            DataReduction: { DataVolume: 6, Primary: { Window: { Count: count } } },
            Primary: { Groupings: [{ Projections: select.map((_, index) => index) }] },
            Version: 1,
          },
          ExecutionMetricsKind: 1,
        },
      }],
    },
    QueryId: "",
  };
}

function commonWhere(range, scope, filters = {}, excluded = new Set()) {
  const movementTypes = filters.movementCode && filters.movementCode !== "all"
    ? [filters.movementCode]
    : scope.movementTypes;
  const conditions = [
    {
      Condition: {
        And: {
          Left: {
            Comparison: {
              ComparisonKind: 2,
              Left: field("d", "Date"),
              Right: literal(`datetime'${range.start}T00:00:00'`),
            },
          },
          Right: {
            Comparison: {
              ComparisonKind: 3,
              Left: field("d", "Date"),
              Right: literal(`datetime'${range.endExclusive}T00:00:00'`),
            },
          },
        },
      },
    },
    {
      Condition: {
        In: {
          Expressions: [field("r", "movement_type")],
          Values: movementTypes.map(value => [literal(stringLiteral(value))]),
        },
      },
    },
  ];

  if (excluded.has("movement")) conditions.splice(1, 1);

  const masterCategory = filters.masterCategory || scope.masterCategory;
  if (!excluded.has("masterCategory") && masterCategory !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("a", "MasterCategory")],
          Values: [[literal(stringLiteral(masterCategory))]],
        },
      },
    });
  }

  if (!excluded.has("category") && filters.category && filters.category !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("a", "Category3")],
          Values: [[literal(stringLiteral(filters.category))]],
        },
      },
    });
  }

  if (!excluded.has("region") && filters.region && filters.region !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("o", "RegionName")],
          Values: [[literal(stringLiteral(filters.region))]],
        },
      },
    });
  }

  if (!excluded.has("outlet") && Array.isArray(filters.outletCodes)) {
    const codes = filters.outletCodes.length ? filters.outletCodes : ["__NO_MATCHING_OUTLET__"];
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("o", "OutletCode")],
          Values: codes.map(value => [literal(stringLiteral(value))]),
        },
      },
    });
  }

  if (!excluded.has("article") && filters.articleNo && filters.articleNo !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("a", "ArticleNo")],
          Values: [[literal(stringLiteral(filters.articleNo))]],
        },
      },
    });
  }


  if (!excluded.has("user") && filters.userCode && filters.userCode !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("u", "created_by")],
          Values: [[literal(stringLiteral(filters.userCode))]],
        },
      },
    });
  }

  if (!excluded.has("po") && filters.poNumber && filters.poNumber !== "all") {
    conditions.push({
      Condition: {
        In: {
          Expressions: [field("r", "PO Clean")],
          Values: [[literal(stringLiteral(filters.poNumber))]],
        },
      },
    });
  }

  return conditions;
}

const COMMON_FROM = [
  { Name: "s", Entity: "Query1", Type: 0 },
  { Name: "r", Entity: "Query3", Type: 0 },
  { Name: "m", Entity: "Over/Under Receiving", Type: 0 },
  { Name: "o", Entity: "DimOutlet", Type: 0 },
  { Name: "d", Entity: "DimDate", Type: 0 },
  { Name: "a", Entity: "DimArticle", Type: 0 },
];

// Query source alias -> model entity, used to look measures up in the schema
// the published report returns with modelsAndExploration.
const SOURCE_ENTITIES = {
  m: "Over/Under Receiving",
  rank: "Over Receiving Rank Measures",
  vis: "For Visula",
  o: "DimOutlet",
  q2: "Query2",
};

// The published page only ever used the "Over" half of the model. The "Under"
// half is resolved at runtime so a model without the mirrored value/score
// measures degrades to the incident measures instead of failing.
const INCIDENT_MEASURES = {
  over: {
    receiving: { key: "OverReceiving", source: "m", candidates: ["Over Receiving"], required: true },
    value: { key: "OverValue", source: "m", candidates: ["Over Receiving Value"], required: true },
    score: { key: "OverScore", source: "rank", candidates: ["Over Receiving Score"], required: true },
    icon: { key: "StatusIcon", source: "vis", candidates: ["Over Receiving Score Icon"], required: true },
    incidents: { key: "OverIncidents", source: "m", candidates: ["Over Receiving Incidents"], required: true },
    incidentPct: { key: "OverIncidentPct", source: "m", candidates: ["Over Receiving Incident%"], required: true },
  },
  under: {
    receiving: { key: "UnderReceiving", source: "m", candidates: ["Under Receiving", "Under Receiving Qty"] },
    value: { key: "UnderValue", source: "m", candidates: ["Under Receiving Value", "Under Receiving  Value"] },
    score: { key: "UnderScore", source: "rank", candidates: ["Under Receiving Score"] },
    icon: { key: "StatusIcon", source: "vis", candidates: ["Under Receiving Score Icon"] },
    incidents: { key: "UnderIncidents", source: "m", candidates: ["Under Receiving Incidents"], required: true },
    incidentPct: { key: "UnderIncidentPct", source: "m", candidates: ["Under Receiving Incident%"], required: true },
  },
};

const OPTIONAL_INCIDENT_KEYS = ["receiving", "value", "score", "icon"];

const USER_SOURCE = { Name: "u", Entity: "Query4", Type: 0 };
const MANAGEMENT_FROM = [
  ...COMMON_FROM,
  USER_SOURCE,
  { Name: "q2", Entity: "Query2", Type: 0 },
  { Name: "g", Entity: "Group Name", Type: 0 },
  { Name: "rank", Entity: "Over Receiving Rank Measures", Type: 0 },
  { Name: "vis", Entity: "For Visula", Type: 0 },
];

function sourceSet(filters, includeUser = false) {
  return includeUser || (filters.userCode && filters.userCode !== "all") ? [...COMMON_FROM, USER_SOURCE] : COMMON_FROM;
}

const KPI_SELECT = [
  sum("s", "ActualInvoicedQuantity", "Sales"),
  sum("r", "qty_in_unit_of_entry", "Receiving"),
  measure("m", "Total Inventory", "Inventory"),
  measure("m", "Stock Day", "StockDay"),
  measure("m", "Latest  Date Stock", "LatestStock"),
  measure("m", "Over Receiving Incidents", "OverIncidents"),
  measure("m", "Over Receiving Incident%", "OverIncidentPct"),
  measure("m", "Under Receiving Incidents", "UnderIncidents"),
  measure("m", "Under Receiving Incident%", "UnderIncidentPct"),
  measure("m", "Over Receiving Value", "OverValue"),
  measure("o", "Outlets with Latest Stock > 0", "ActiveOutlets"),
];

function composeKpiContexts(rows, fallback = {}) {
  const numeric = value => value == null || value === "" ? null : (Number.isFinite(Number(value)) ? Number(value) : null);
  const sumPresent = fieldName => {
    const values = rows.map(row => numeric(row[fieldName])).filter(value => value != null);
    return values.length ? values.reduce((total, value) => total + value, 0) : null;
  };
  const preferSource = (fieldName, computed) => numeric(fallback[fieldName]) ?? computed;
  const compositeSales = sumPresent("Sales");
  const weightedStockDays = rows.reduce((total, row) => {
    const rowSales = numeric(row.Sales);
    const stockDay = numeric(row.StockDay);
    return total + (rowSales != null && stockDay != null ? rowSales * stockDay : 0);
  }, 0);
  const incidentRate = (incidentField, percentField) => {
    const incidents = sumPresent(incidentField);
    const population = rows.reduce((total, row) => {
      const count = numeric(row[incidentField]);
      const percent = numeric(row[percentField]);
      return total + (count != null && percent > 0 ? count / (percent / 100) : 0);
    }, 0);
    return incidents != null && population ? (incidents / population) * 100 : null;
  };
  return {
    Sales: preferSource("Sales", compositeSales),
    Receiving: preferSource("Receiving", sumPresent("Receiving")),
    Inventory: preferSource("Inventory", sumPresent("Inventory")),
    StockDay: preferSource("StockDay", compositeSales ? weightedStockDays / compositeSales : null),
    LatestStock: preferSource("LatestStock", sumPresent("LatestStock")),
    OverIncidents: preferSource("OverIncidents", sumPresent("OverIncidents")),
    OverIncidentPct: preferSource("OverIncidentPct", incidentRate("OverIncidents", "OverIncidentPct")),
    UnderIncidents: preferSource("UnderIncidents", sumPresent("UnderIncidents")),
    UnderIncidentPct: preferSource("UnderIncidentPct", incidentRate("UnderIncidents", "UnderIncidentPct")),
    OverValue: preferSource("OverValue", sumPresent("OverValue")),
    ActiveOutlets: numeric(fallback.ActiveOutlets),
  };
}

function combineIncidentUserContexts(rows) {
  const numeric = value => value == null || value === "" ? null : (Number.isFinite(Number(value)) ? Number(value) : null);
  const groups = new Map();
  for (const row of rows) {
    const outletCode = String(row.OutletCode ?? "").trim();
    const createdBy = String(row.CreatedBy ?? "").trim();
    const key = `${outletCode}\u001f${createdBy}`;
    const current = groups.get(key) || {
      OutletCode: outletCode || null,
      OutletName: row.OutletName || null,
      CreatedBy: createdBy || null,
      OverIncidents: 0,
      UnderIncidents: 0,
      _overPopulation: 0,
      _underPopulation: 0,
      _overSeen: false,
      _underSeen: false,
    };
    if (!current.OutletName && row.OutletName) current.OutletName = row.OutletName;
    for (const prefix of ["Over", "Under"]) {
      const count = numeric(row[`${prefix}Incidents`]);
      const percent = numeric(row[`${prefix}IncidentPct`]);
      if (count != null) {
        current[`${prefix}Incidents`] += count;
        current[`_${prefix.toLowerCase()}Seen`] = true;
      }
      if (count != null && percent > 0) current[`_${prefix.toLowerCase()}Population`] += count / (percent / 100);
    }
    groups.set(key, current);
  }
  return [...groups.values()].map(row => {
    for (const prefix of ["Over", "Under"]) {
      const lower = prefix.toLowerCase();
      if (!row[`_${lower}Seen`]) row[`${prefix}Incidents`] = null;
      row[`${prefix}IncidentPct`] = row[`_${lower}Population`] > 0
        ? (row[`${prefix}Incidents`] / row[`_${lower}Population`]) * 100
        : null;
      delete row[`_${lower}Population`];
      delete row[`_${lower}Seen`];
    }
    return row;
  });
}

function combineCategoryContexts(rows) {
  const numeric = value => (value == null || value === "" ? null : (Number.isFinite(Number(value)) ? Number(value) : null));
  const additive = ["OpeningStock", "Receiving", "TotalInventory", "TotalSales", "EstimatedClosingStock", "OverReceiving", "UnderReceiving", "OverValue", "UnderValue", "OverIncidents", "UnderIncidents"];
  const groups = new Map();
  for (const row of rows) {
    const key = String(row.Category ?? "");
    const current = groups.get(key);
    if (!current) {
      groups.set(key, { ...row, _overPopulation: 0, _underPopulation: 0 });
      continue;
    }
    for (const field of additive) {
      const value = numeric(row[field]);
      if (value != null) current[field] = (numeric(current[field]) ?? 0) + value;
    }
    for (const field of ["StdStockDays", "OverScore", "UnderScore", "StatusIcon"]) {
      if (current[field] == null && row[field] != null) current[field] = row[field];
    }
  }
  return [...groups.values()].map(row => {
    for (const prefix of ["Over", "Under"]) {
      const count = numeric(row[`${prefix}Incidents`]);
      const percent = numeric(row[`${prefix}IncidentPct`]);
      if (count != null && percent > 0) row[`${prefix}IncidentPct`] = percent;
      delete row[`_${prefix.toLowerCase()}Population`];
    }
    return row;
  });
}

function rangeForDays(scope, days) {
  const requestedDays = Number(days) || 30;
  return {
    start: shiftIsoDate(scope.endExclusive, -requestedDays),
    endExclusive: scope.endExclusive,
    days: requestedDays,
  };
}

function rangeForFilters(scope, filters = {}) {
  const dateFrom = /^\d{4}-\d{2}-\d{2}$/.test(String(filters.dateFrom || "")) ? String(filters.dateFrom) : null;
  const dateTo = /^\d{4}-\d{2}-\d{2}$/.test(String(filters.dateTo || "")) ? String(filters.dateTo) : null;
  if (dateFrom && dateTo && dateFrom <= dateTo) {
    const endExclusive = shiftIsoDate(dateTo, 1);
    const days = Math.max(1, Math.round((Date.parse(`${endExclusive}T00:00:00Z`) - Date.parse(`${dateFrom}T00:00:00Z`)) / 86_400_000));
    return { start: dateFrom, endExclusive, days };
  }
  return rangeForDays(scope, filters.days);
}

function buildMeasureIndex(model) {
  const tables = model?.model?.tables;
  if (!Array.isArray(tables) || !tables.length) return null;
  const index = new Map();
  for (const table of tables) {
    const names = (table?.measures || []).map(item => item?.name).filter(Boolean);
    if (table?.name) index.set(table.name, new Set(names));
  }
  return index.size ? index : null;
}

function mergeNumeric(target, key, value) {
  const number = value == null || value === "" ? null : Number(value);
  if (!Number.isFinite(number)) return;
  target[key] = (target[key] ?? 0) + number;
}

export class PowerBIDataClient {
  constructor() {
    this.model = null;
    this.report = null;
    this.scope = null;
    this.sourceTimestamp = null;
    this.measureIndex = null;
    this.unsupportedMeasures = new Set();
    this.underProbePromise = null;
  }

  async connect({ signal } = {}) {
    const response = await fetch(
      `${API_ROOT}/public/reports/${RESOURCE_KEY}/modelsAndExploration?preferReadOnlySession=true`,
      { headers: requestHeaders(), cache: "no-store", signal }
    );
    if (!response.ok) throw new Error(`Power BI connection failed (${response.status}).`);

    const payload = await response.json();
    this.model = payload.models?.[0];
    this.report = payload.exploration?.report;
    if (!this.model || !this.report) throw new Error("The published Power BI model could not be identified.");

    this.scope = savedReportScope(payload.exploration?.explorationContent?.explorationDocument || "");
    this.measureIndex = buildMeasureIndex(this.model);
    const refresh = this.model.LastRefreshTime || payload.package?.LastRefreshTime || null;
    this.sourceTimestamp = refresh && !/[zZ]|[+-]\d{2}:?\d{2}$/.test(refresh) ? `${refresh}Z` : refresh;
    return this;
  }

  resolveMeasure(source, candidates) {
    const entity = SOURCE_ENTITIES[source];
    const known = entity && this.measureIndex ? this.measureIndex.get(entity) : null;
    for (const candidate of candidates) {
      if (this.unsupportedMeasures.has(`${source}:${candidate}`)) continue;
      if (known) {
        if (known.has(candidate)) return candidate;
        continue;
      }
      return candidate;
    }
    return null;
  }

  /**
   * Returns the select fragments for the requested over/under measures plus the
   * keys the published model does not expose, so the caller can drop the
   * matching columns instead of showing empty ones.
   */
  resolveIncidentMeasures(mode, names) {
    const set = INCIDENT_MEASURES[mode === "under" ? "under" : "over"];
    const resolved = {};
    const missing = [];
    for (const name of names) {
      const definition = set[name];
      if (!definition) continue;
      const measureName = definition.required
        ? definition.candidates[0]
        : this.resolveMeasure(definition.source, definition.candidates);
      if (measureName) resolved[name] = { ...definition, measureName };
      else missing.push(definition.key);
    }
    return { resolved, missing };
  }

  incidentSelect(resolved, names) {
    return names
      .filter(name => resolved[name])
      .map(name => measure(resolved[name].source, resolved[name].measureName, resolved[name].key));
  }

  /**
   * One tolerant batch that asks for each optional "Under" measure on its own.
   * Results are inspected per query, so a model that publishes only some of
   * them keeps the ones it has.
   */
  async ensureUnderMeasureSupport({ range, scope, signal } = {}) {
    if (this.measureIndex) return;
    if (this.underProbePromise) return this.underProbePromise;

    const effectiveScope = scope ? { ...this.scope, ...scope } : this.scope;
    const probeRange = range || rangeForDays(effectiveScope, 7);
    const probes = OPTIONAL_INCIDENT_KEYS
      .map(name => INCIDENT_MEASURES.under[name])
      .flatMap(definition => definition.candidates.map(candidate => ({ definition, candidate })));

    this.underProbePromise = (async () => {
      const where = commonWhere(probeRange, effectiveScope, {});
      const specs = probes.map((probe, index) => ({
        key: `probe${index}`,
        query: createQuery([measure(probe.definition.source, probe.candidate, "Probe")], MANAGEMENT_FROM, where, 1),
      }));
      let outcomes;
      try {
        outcomes = await this.runSpecsTolerant(specs, { signal });
      } catch {
        outcomes = specs.map(() => false);
      }
      probes.forEach((probe, index) => {
        if (!outcomes[index]) this.unsupportedMeasures.add(`${probe.definition.source}:${probe.candidate}`);
      });
    })();

    return this.underProbePromise;
  }

  async runSpecsTolerant(specs, { signal } = {}) {
    const queries = specs.map(spec => ({
      ...spec.query,
      ApplicationContext: {
        DatasetId: this.model.dbName,
        Sources: [{ ReportId: this.report.objectId }],
      },
    }));
    const response = await fetch(`${API_ROOT}/public/reports/querydata?synchronous=true`, {
      method: "POST",
      headers: requestHeaders(true),
      cache: "no-store",
      signal,
      body: JSON.stringify({ version: "1.0.0", queries, cancelQueries: [], modelId: this.model.id }),
    });
    if (!response.ok) throw new Error(`Power BI data request failed (${response.status}).`);
    const payload = await response.json();
    return specs.map((_, index) => {
      const entry = payload.results?.[index];
      return Boolean(entry?.result?.data) && !entry?.result?.error;
    });
  }

  async runSpecs(specs, { signal } = {}) {
    const queries = specs.map(spec => ({
      ...spec.query,
      ApplicationContext: {
        DatasetId: this.model.dbName,
        Sources: [{ ReportId: this.report.objectId }],
      },
    }));

    const response = await fetch(`${API_ROOT}/public/reports/querydata?synchronous=true`, {
      method: "POST",
      headers: requestHeaders(true),
      cache: "no-store",
      signal,
      body: JSON.stringify({ version: "1.0.0", queries, cancelQueries: [], modelId: this.model.id }),
    });
    if (!response.ok) throw new Error(`Power BI data request failed (${response.status}).`);

    const payload = await response.json();
    if (payload.error) throw new Error(payload.error.message || "Power BI returned an error.");
    if (!Array.isArray(payload.results) || payload.results.length < specs.length) {
      throw new Error("Power BI returned an incomplete dashboard response.");
    }

    const decoded = {};
    specs.forEach((spec, index) => {
      decoded[spec.key] = decodeResult(payload.results?.[index]);
    });

    return {
      decoded,
      queryTimestamp: payload.results[0]?.result?.data?.timestamp || new Date().toISOString(),
    };
  }

  async load(filters = {}, { section = "all", signal } = {}) {
    if (!this.model || !this.report || !this.scope) await this.connect({ signal });

    const range = rangeForFilters(this.scope, filters);
    const where = commonWhere(range, this.scope, filters);
    const categoryOptionWhere = commonWhere(range, this.scope, filters, new Set(["category", "article"]));
    const articleOptionWhere = commonWhere(range, this.scope, filters, new Set(["article"]));
    const outletOptionWhere = commonWhere(range, this.scope, filters, new Set(["region", "outlet"]));
    const userOptionWhere = commonWhere(range, this.scope, filters, new Set(["user"]));
    const movementOptionWhere = commonWhere(range, this.scope, filters, new Set(["movement"]));
    const from = sourceSet(filters);
    const userFrom = sourceSet(filters, true);

    const specs = [
      { key: "kpis", query: createQuery(KPI_SELECT, from, where, 50) },
      {
        key: "trend",
        query: createQuery([
          column("d", "Date", "Date"),
          sum("s", "ActualInvoicedQuantity", "Sales"),
          sum("r", "qty_in_unit_of_entry", "Receiving"),
        ], from, where, 500),
      },
      {
        key: "categories",
        query: createQuery([
          column("a", "Category3", "Category"),
          sum("s", "ActualInvoicedQuantity", "Sales"),
          sum("r", "qty_in_unit_of_entry", "Receiving"),
          measure("m", "Total Inventory", "Inventory"),
          measure("m", "Over Receiving Value", "OverValue"),
          measure("m", "Over Receiving Incidents", "OverIncidents"),
          measure("m", "Under Receiving Incidents", "UnderIncidents"),
        ], from, where, 500),
      },
      {
        key: "regions",
        query: createQuery([
          column("o", "RegionName", "Region"),
          sum("s", "ActualInvoicedQuantity", "Sales"),
          sum("r", "qty_in_unit_of_entry", "Receiving"),
          measure("m", "Total Inventory", "Inventory"),
          measure("m", "Over Receiving Value", "OverValue"),
          measure("m", "Over Receiving Incidents", "OverIncidents"),
          measure("m", "Under Receiving Incidents", "UnderIncidents"),
        ], from, where, 100),
      },
      {
        key: "outlets",
        query: createQuery([
          column("o", "OutletCode", "OutletCode"),
          column("o", "OutletName", "Outlet"),
          column("o", "RegionName", "Region"),
          sum("s", "ActualInvoicedQuantity", "Sales"),
          sum("r", "qty_in_unit_of_entry", "Receiving"),
          measure("m", "Total Inventory", "Inventory"),
          measure("m", "Stock Day", "StockDay"),
          measure("m", "Over Receiving Value", "OverValue"),
          measure("m", "Over Receiving Incidents", "OverIncidents"),
          measure("m", "Under Receiving Incidents", "UnderIncidents"),
        ], from, where, 5000),
      },
      {
        key: "categoryOptions",
        query: createQuery([column("a", "Category3", "Category")], from, categoryOptionWhere, 1000),
      },
      {
        key: "articleOptions",
        query: createQuery([
          column("a", "ArticleNo", "ArticleNo"),
          column("a", "ArticleName", "ArticleName"),
          column("a", "Category3", "Category"),
          sum("r", "qty_in_unit_of_entry", "Receiving"),
          sum("s", "ActualInvoicedQuantity", "Sales"),
        ], from, articleOptionWhere, MAX_QUERY_ROWS),
      },
      {
        key: "outletOptions",
        query: createQuery([
          column("o", "OutletCode", "OutletCode"),
          column("o", "OutletName", "Outlet"),
          column("o", "RegionName", "Region"),
        ], from, outletOptionWhere, MAX_QUERY_ROWS),
      },
      {
        key: "userOptions",
        query: createQuery([column("u", "created_by", "UserCode")], userFrom, userOptionWhere, MAX_QUERY_ROWS),
      },
      {
        key: "movementOptions",
        query: createQuery([column("r", "movement_type", "MovementCode")], from, movementOptionWhere, 100),
      },
    ];

    const snapshotOutletSpec = {
      key: "snapshotOutlets",
      query: createQuery([
        column("o", "OutletCode", "OutletCode"),
        column("o", "OutletName", "Outlet"),
        column("o", "RegionName", "Region"),
        sum("s", "ActualInvoicedQuantity", "Sales"),
        sum("r", "qty_in_unit_of_entry", "Receiving"),
      ], from, where, MAX_QUERY_ROWS),
    };
    const sections = {
      core: specs.slice(0, 4),
      supporting: specs.slice(4),
      kpis: [specs[0]],
      trend: [specs[1]],
      outlets: [specs[4]],
      articles: [specs[6]],
      snapshotOutlets: [snapshotOutletSpec],
      kpiOutlets: [specs[0], specs[4]],
      snapshotBreakdowns: [specs[2], specs[3]],
      options: specs.slice(5),
    };
    let requestedSpecs = sections[section] || specs;
    const needsCompositeKpi = filters.masterCategory === "all" && (section === "core" || section === "all" || section === "kpis");
    const compositeKeys = [];
    if (needsCompositeKpi) {
      const compositeSpecs = this.scope.masterCategories.map((masterCategory, index) => {
        const key = `compositeKpi${index}`;
        compositeKeys.push(key);
        return {
          key,
          query: createQuery(KPI_SELECT, from, commonWhere(range, this.scope, { ...filters, masterCategory }), 50),
        };
      });
      requestedSpecs = [...requestedSpecs, ...compositeSpecs];
    }
    const { decoded, queryTimestamp } = await this.runSpecs(requestedSpecs, { signal });
    if (needsCompositeKpi) {
      const rows = compositeKeys.flatMap(key => decoded[key] || []);
      decoded.kpis = [composeKpiContexts(rows, decoded.kpis?.[0] || {})];
      compositeKeys.forEach(key => delete decoded[key]);
    }

    return {
      ...decoded,
      range,
      sourceTimestamp: this.sourceTimestamp,
      queryTimestamp,
      scope: this.scope,
    };
  }

  async loadArticleDetails(filters = {}, metric = "OverValue", { signal } = {}) {
    if (!this.model || !this.report || !this.scope) await this.connect({ signal });
    const range = rangeForFilters(this.scope, filters);
    const where = commonWhere(range, this.scope, filters);
    const from = sourceSet(filters);
    const metricColumns = {
      Receiving: [sum("r", "qty_in_unit_of_entry", "Receiving")],
      Sales: [sum("s", "ActualInvoicedQuantity", "Sales")],
      Gap: [
        sum("r", "qty_in_unit_of_entry", "Receiving"),
        sum("s", "ActualInvoicedQuantity", "Sales"),
      ],
      Inventory: [measure("m", "Total Inventory", "Inventory")],
      StockDay: [measure("m", "Stock Day", "StockDay")],
      OverValue: [measure("m", "Over Receiving Value", "OverValue")],
      OverIncidents: [measure("m", "Over Receiving Incidents", "OverIncidents")],
      UnderIncidents: [measure("m", "Under Receiving Incidents", "UnderIncidents")],
      Incidents: [
        measure("m", "Over Receiving Incidents", "OverIncidents"),
        measure("m", "Under Receiving Incidents", "UnderIncidents"),
      ],
    };
    const selectedMetrics = metricColumns[metric] || metricColumns.OverValue;
    const specs = [{
      key: "articles",
      query: createQuery([
        column("a", "ArticleNo", "ArticleNo"),
        column("a", "ArticleName", "ArticleName"),
        column("a", "MasterCategory", "MasterCategory"),
        column("a", "Category3", "Category"),
        ...selectedMetrics,
      ], from, where, MAX_QUERY_ROWS),
    }];
    const { decoded, queryTimestamp } = await this.runSpecs(specs, { signal });
    return { rows: decoded.articles, range, queryTimestamp };
  }

  async loadManagementTable(filters = {}, tableNumber = 1, queryContext = null, { signal } = {}) {
    if (!this.model || !this.report || !this.scope) await this.connect({ signal });
    // Keep drill-downs in the exact snapshot context that produced the
    // clicked value, even if the published report refreshes afterward.
    const effectiveScope = queryContext?.scope
      ? { ...this.scope, ...queryContext.scope }
      : this.scope;
    const range = queryContext?.range
      ? { ...queryContext.range, days: Number(queryContext.range.days) || Number(filters.days) || 30 }
      : rangeForFilters(effectiveScope, filters);
    const where = commonWhere(range, effectiveScope, filters);
    const table = Number(tableNumber) || 1;
    const mode = queryContext?.mode === "under" ? "under" : "over";
    // Tables 1-4 mirror the saved "Over Receiving" page; the same shapes are
    // rebuilt against the Under measures when the incident mode is Under.
    const modeApplies = [1, 2, 3, 4].includes(table);
    if (modeApplies && mode === "under") await this.ensureUnderMeasureSupport({ range, scope: effectiveScope, signal });
    const { resolved, missing } = this.resolveIncidentMeasures(
      modeApplies ? mode : "over",
      table === 4 ? ["receiving", "value", "score", "icon", "incidents", "incidentPct"] : ["receiving", "value", "score", "icon"]
    );

    const common = [
      measure("m", "Opening Stock", "OpeningStock"),
      sum("r", "qty_in_unit_of_entry", "Receiving"),
      measure("m", "Total Inventory", "TotalInventory"),
      sum("s", "ActualInvoicedQuantity", "TotalSales"),
      measure("m", "STD. Stock Days", "StdStockDays"),
    ];
    const operational = [
      ...common,
      measure("m", "Stock Day", "CurrentStockDay"),
      measure("q2", "Latest Stock", "CurrentStockSystem"),
      measure("m", "Closing Stock On Receiving", "ClosingStockReceiving"),
      ...this.incidentSelect(resolved, ["receiving", "value", "score", "icon"]),
    ];
    const definitions = {
      1: [column("o", "OutletName", "OutletName"), column("o", "OutletCode", "OutletCode"), column("a", "ArticleNo", "ArticleNo"), column("a", "ArticleName", "ArticleName"), column("a", "Category3", "Category"), ...operational, measure("rank", "sorting", "Sorting")],
      2: [column("o", "OutletName", "OutletName"), column("o", "OutletCode", "OutletCode"), column("g", "Article Combo", "ArticleCombo"), ...operational, measure("rank", "sorting", "Sorting")],
      3: [column("o", "OutletName", "OutletName"), column("o", "OutletCode", "OutletCode"), ...operational, measure("rank", "sorting", "Sorting")],
      4: [column("a", "Category3", "Category"), ...common, measure("o", "Est. Closing Stock", "EstimatedClosingStock"), ...this.incidentSelect(resolved, ["receiving", "value", "score", "icon", "incidents", "incidentPct"])],
      5: [column("a", "ArticleNo", "ArticleNo"), column("r", "PO Clean", "PONumber"), column("r", "movement_type", "MovementType"), column("u", "created_by", "CreatedBy"), column("u", "document_date", "PODate"), column("r", "posting_date", "ReceivingDate"), sum("r", "qty_in_unit_of_entry", "Receiving")],
      6: [column("o", "OutletCode", "OutletCode"), column("o", "OutletName", "OutletName"), column("u", "created_by", "CreatedBy"), measure("m", "Over Receiving Incidents", "OverIncidents"), measure("m", "Over Receiving Incident%", "OverIncidentPct"), measure("m", "Under Receiving Incidents", "UnderIncidents"), measure("m", "Under Receiving Incident%", "UnderIncidentPct")],
    };
    const select = definitions[table] || definitions[1];
    // Tables 6 and 4 are reached from the cross-division incident view, so an
    // "All" scope is answered one business division at a time and merged.
    if ((table === 6 || table === 4) && filters.masterCategory === "all") {
      const specs = effectiveScope.masterCategories.map((masterCategory, index) => ({
        key: `rows${index}`,
        query: createQuery(select, MANAGEMENT_FROM, commonWhere(range, effectiveScope, { ...filters, masterCategory }), MAX_QUERY_ROWS),
      }));
      const { decoded, queryTimestamp } = await this.runSpecs(specs, { signal });
      const merged = specs.flatMap(spec => decoded[spec.key] || []);
      const rows = table === 6 ? combineIncidentUserContexts(merged) : combineCategoryContexts(merged);
      return { rows, range, queryTimestamp, missing, mode };
    }
    const specs = [{ key: "rows", query: createQuery(select, MANAGEMENT_FROM, where, MAX_QUERY_ROWS) }];
    const { decoded, queryTimestamp } = await this.runSpecs(specs, { signal });
    return { rows: table === 6 ? combineIncidentUserContexts(decoded.rows) : decoded.rows, range, queryTimestamp, missing, mode };
  }

  /**
   * Under-receiving value for the KPI card, the category signal and the
   * division table. Kept apart from the shared snapshot so the fast default
   * load never depends on a measure the published model may not expose.
   */
  async loadUnderValueContext(filters = {}, { range, scope, signal } = {}) {
    if (!this.model || !this.report || !this.scope) await this.connect({ signal });
    const effectiveScope = scope ? { ...this.scope, ...scope } : this.scope;
    const effectiveRange = range
      ? { ...range, days: Number(range.days) || Number(filters.days) || 30 }
      : rangeForFilters(effectiveScope, filters);
    await this.ensureUnderMeasureSupport({ range: effectiveRange, scope: effectiveScope, signal });

    // The published report exposes Under Receiving (a negative quantity) and
    // Under Receiving Score, but no Under Receiving Value. Where the value
    // measure is absent, the quantity is reported in its place and labelled as
    // such rather than left blank.
    const { resolved } = this.resolveIncidentMeasures("under", ["value", "receiving"]);
    const chosen = resolved.value || resolved.receiving;
    if (!chosen) return { supported: false, kind: "none", kpi: null, categories: new Map(), regions: new Map() };
    const kind = resolved.value ? "value" : "quantity";

    const valueSelect = measure(chosen.source, chosen.measureName, "UnderValue");
    const partitions = filters.masterCategory === "all"
      ? effectiveScope.masterCategories.map(masterCategory => ({ ...filters, masterCategory }))
      : [filters];

    const specs = partitions.flatMap((partitionFilters, index) => {
      const partitionWhere = commonWhere(effectiveRange, effectiveScope, partitionFilters);
      return [
        { key: `kpi${index}`, query: createQuery([valueSelect], sourceSet(partitionFilters), partitionWhere, 50) },
        { key: `categories${index}`, query: createQuery([column("a", "Category3", "Category"), valueSelect], sourceSet(partitionFilters), partitionWhere, 500) },
        { key: `regions${index}`, query: createQuery([column("o", "RegionName", "Region"), valueSelect], sourceSet(partitionFilters), partitionWhere, 100) },
      ];
    });

    const { decoded, queryTimestamp } = await this.runSpecs(specs, { signal });
    const categories = new Map();
    const regions = new Map();
    let kpi = null;

    partitions.forEach((_, index) => {
      for (const row of decoded[`kpi${index}`] || []) {
        const value = row.UnderValue == null || row.UnderValue === "" ? null : Number(row.UnderValue);
        if (Number.isFinite(value)) kpi = (kpi ?? 0) + value;
      }
      for (const [key, target, field] of [[`categories${index}`, categories, "Category"], [`regions${index}`, regions, "Region"]]) {
        for (const row of decoded[key] || []) {
          const name = String(row[field] ?? "");
          const value = row.UnderValue == null || row.UnderValue === "" ? null : Number(row.UnderValue);
          if (!Number.isFinite(value)) continue;
          target.set(name, (target.get(name) ?? 0) + value);
        }
      }
    });

    return { supported: true, kind, kpi, categories, regions, range: effectiveRange, queryTimestamp };
  }

  /**
   * Category-level incidents for every outlet/user pair on the user-incident
   * table. This is the second sheet of the incident workbook export.
   */
  async loadIncidentCategoryBreakdown(filters = {}, { range, scope, mode = "over", signal } = {}) {
    if (!this.model || !this.report || !this.scope) await this.connect({ signal });
    const effectiveScope = scope ? { ...this.scope, ...scope } : this.scope;
    const effectiveRange = range
      ? { ...range, days: Number(range.days) || Number(filters.days) || 30 }
      : rangeForFilters(effectiveScope, filters);
    const incidentMode = mode === "under" ? "under" : "over";
    if (incidentMode === "under") await this.ensureUnderMeasureSupport({ range: effectiveRange, scope: effectiveScope, signal });

    const { resolved, missing } = this.resolveIncidentMeasures(incidentMode, ["receiving", "value", "score", "icon", "incidents", "incidentPct"]);
    const select = [
      column("o", "OutletCode", "OutletCode"),
      column("o", "OutletName", "OutletName"),
      column("u", "created_by", "CreatedBy"),
      column("a", "Category3", "Category"),
      measure("m", "Opening Stock", "OpeningStock"),
      sum("r", "qty_in_unit_of_entry", "Receiving"),
      measure("m", "Total Inventory", "TotalInventory"),
      sum("s", "ActualInvoicedQuantity", "TotalSales"),
      measure("m", "STD. Stock Days", "StdStockDays"),
      measure("o", "Est. Closing Stock", "EstimatedClosingStock"),
      ...this.incidentSelect(resolved, ["receiving", "value", "score", "icon", "incidents", "incidentPct"]),
    ];

    const partitions = filters.masterCategory === "all"
      ? effectiveScope.masterCategories.map(masterCategory => ({ ...filters, masterCategory }))
      : [filters];
    const specs = partitions.map((partitionFilters, index) => ({
      key: `rows${index}`,
      query: createQuery(select, MANAGEMENT_FROM, commonWhere(effectiveRange, effectiveScope, partitionFilters), MAX_QUERY_ROWS),
    }));

    const { decoded, queryTimestamp } = await this.runSpecs(specs, { signal });
    const grouped = new Map();
    for (const spec of specs) {
      for (const row of decoded[spec.key] || []) {
        const key = `${row.OutletCode ?? ""}\u001f${row.CreatedBy ?? ""}\u001f${row.Category ?? ""}`;
        const current = grouped.get(key);
        if (!current) {
          grouped.set(key, { ...row });
          continue;
        }
        for (const field of ["OpeningStock", "Receiving", "TotalInventory", "TotalSales", "EstimatedClosingStock", "OverReceiving", "UnderReceiving", "OverValue", "UnderValue", "OverIncidents", "UnderIncidents"]) {
          if (row[field] != null) mergeNumeric(current, field, row[field]);
        }
      }
    }

    return { rows: [...grouped.values()], range: effectiveRange, queryTimestamp, missing, mode: incidentMode };
  }
}
