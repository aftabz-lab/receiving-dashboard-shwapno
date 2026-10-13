const ORGANIZATION_SNAPSHOT_URL = "https://wstxgbzmsbosinmhhjbl.supabase.co/rest/v1/dashboard_snapshots?select=snapshot_key,payload,updated_at&snapshot_key=eq.zone-distribution&limit=1";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_MCw_J7uorsKtmmokW1OpCg_Ej5DURhw";
const ZONE_DATA_URL = "https://aftabz-lab.github.io/zone-distribution-dashboard/data/dashboard_data.json";
const FALLBACK_URL = new URL("./organization-fallback.json", import.meta.url).href;
const DRIVE_FOLDER_ID = "16HTr8nfPz4P2PMr4QB0bjgwiD110Qd-0";
let fallbackPromise = null;

export function normalizeOutletCode(value) {
  return String(value ?? "").replace(/[\u200B-\u200D\uFEFF]/g, "").trim().toUpperCase();
}
function text(value) { return String(value ?? "").trim(); }
function key(value) { return text(value).toLowerCase().replace(/[_\s]+/g, ""); }

function normalizeRow(row) {
  const fields = new Map(Object.entries(row || {}).map(([name, value]) => [key(name), value]));
  const pick = (...names) => {
    for (const name of names) { const value = text(fields.get(key(name))); if (value) return value; }
    return "";
  };
  const rhoName = pick("Regional Head HR Name", "RHO", "RHO Name", "Regional Head Name");
  const zonalName = pick("Zonal HR Name", "Zonal Name");
  return {
    OutletCode: normalizeOutletCode(pick("CODE", "OutletCode", "Outlet Code")),
    OutletName: pick("Outlet Name", "OutletName", "Name"),
    Division: pick("Division"), District: pick("District"), Area: pick("Area"),
    RHO: rhoName || pick("Leader"), RHOShort: pick("RHOShort", "Leader"),
    Zonal: zonalName || pick("Zonal"), ZonalShort: pick("ZonalShort", "Zonal"),
    Format: pick("Format"), Status: pick("Status"),
    rhoNameProvided: Boolean(rhoName), zonalNameProvided: Boolean(zonalName),
  };
}

function decode(payload) {
  const record = Array.isArray(payload) ? payload[0] : payload;
  const snapshot = record?.payload?.snapshot || record?.snapshot || record;
  if (snapshot?.driveFolderId && snapshot.driveFolderId !== DRIVE_FOLDER_ID) throw new Error("Zone mapping belongs to a different Drive folder.");
  const rows = (Array.isArray(snapshot?.rows) ? snapshot.rows : []).map(normalizeRow).filter(row => row.OutletCode);
  if (!rows.length || !rows.some(row => row.RHO || row.Zonal)) throw new Error("The published Zone Distribution snapshot contains no valid outlet hierarchy.");
  const byOutlet = new Map();
  rows.forEach(row => { if (!byOutlet.has(row.OutletCode)) byOutlet.set(row.OutletCode, row); });
  return {
    rows: [...byOutlet.values()], byOutlet,
    sourceFile: snapshot.fileName || snapshot.sourceFile || snapshot.meta?.sourceWorkbook || "Zone Distribution snapshot",
    updatedAt: snapshot.savedAt || record?.updated_at || snapshot.updatedAt || snapshot.meta?.savedAt || snapshot.meta?.generatedAt || null,
  };
}

async function readJson(url, headers = {}, cache = "no-store") {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 8000);
  try {
    const response = await fetch(url, { headers, cache, signal: controller.signal });
    if (!response.ok) throw new Error("Organization mapping failed (" + response.status + ").");
    return decode(await response.json());
  } finally { clearTimeout(timer); }
}
function sameRole(a, b) { return Boolean(text(a)) && text(a).toLowerCase() === text(b).toLowerCase(); }
function supplementRow(row, reference) {
  if (!reference) return row;
  return {
    ...row,
    // Resolve full HR names only for the same role. Changed assignments must
    // not retain a previous manager just because their outlet code matches.
    RHO: !row.rhoNameProvided && sameRole(row.RHOShort, reference.RHOShort) ? reference.RHO : row.RHO,
    Zonal: !row.zonalNameProvided && sameRole(row.ZonalShort, reference.ZonalShort) ? reference.Zonal : row.Zonal,
  };
}

export async function loadOrganizationSnapshot() {
  if (!fallbackPromise) fallbackPromise = readJson(FALLBACK_URL, {}, "no-cache").catch(() => null);
  const [cloudResult, zoneResult, fallbackResult] = await Promise.allSettled([
    readJson(ORGANIZATION_SNAPSHOT_URL, { apikey: SUPABASE_PUBLISHABLE_KEY }),
    readJson(ZONE_DATA_URL, {}, "no-cache"),
    fallbackPromise,
  ]);
  const cloud = cloudResult.status === "fulfilled" ? cloudResult.value : null;
  const zone = zoneResult.status === "fulfilled" ? zoneResult.value : null;
  const fallback = fallbackResult.status === "fulfilled" ? fallbackResult.value : null;
  const timestamp = source => Date.parse(source?.updatedAt || "") || 0;
  const primary = cloud && zone && timestamp(zone) > timestamp(cloud) ? zone : cloud || zone || fallback;
  if (!primary) throw new Error("No valid Zone Distribution outlet hierarchy is available.");
  const references = [fallback, zone].filter(Boolean).sort((a, b) => timestamp(a) - timestamp(b));
  const referenceByOutlet = new Map();
  references.forEach(source => source.rows.forEach(row => referenceByOutlet.set(row.OutletCode, row)));
  const rows = primary.rows.map(row => {
    const result = supplementRow(row, referenceByOutlet.get(row.OutletCode));
    const { rhoNameProvided, zonalNameProvided, ...publicRow } = result;
    return publicRow;
  });
  const byOutlet = new Map(rows.map(row => [row.OutletCode, row]));
  return { rows, byOutlet, sourceFile: primary.sourceFile, updatedAt: primary.updatedAt, revision: JSON.stringify(rows) };
}

