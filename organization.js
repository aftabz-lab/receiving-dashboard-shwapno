const ORGANIZATION_SNAPSHOT_URL = "https://wstxgbzmsbosinmhhjbl.supabase.co/rest/v1/dashboard_snapshots?select=snapshot_key,payload,updated_at&snapshot_key=eq.zone-distribution&limit=1";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_MCw_J7uorsKtmmokW1OpCg_Ej5DURhw";

export function normalizeOutletCode(value) {
  return String(value ?? "").trim().toUpperCase();
}

function text(value) {
  return String(value ?? "").trim();
}

function normalizeRow(row) {
  return {
    OutletCode: normalizeOutletCode(row.CODE),
    OutletName: text(row["Outlet Name"]),
    Division: text(row.Division),
    District: text(row.District),
    Area: text(row.Area),
    RHO: text(row["Regional Head HR Name"]) || text(row.Leader),
    RHOShort: text(row.Leader),
    Zonal: text(row["Zonal HR Name"]) || text(row.Zonal),
    ZonalShort: text(row.Zonal),
    Format: text(row.Format),
    Status: text(row.Status),
  };
}

export async function loadOrganizationSnapshot() {
  const response = await fetch(ORGANIZATION_SNAPSHOT_URL, {
    headers: { apikey: SUPABASE_PUBLISHABLE_KEY },
    cache: "no-store",
  });
  if (!response.ok) throw new Error(`Organization mapping failed (${response.status}).`);

  const result = await response.json();
  const record = Array.isArray(result) ? result[0] : null;
  const snapshot = record?.payload?.snapshot;
  if (!Array.isArray(snapshot?.rows) || !snapshot.rows.length) {
    throw new Error("The published Zone Distribution snapshot contains no outlet mapping.");
  }

  const rows = snapshot.rows.map(normalizeRow).filter(row => row.OutletCode);
  const byOutlet = new Map();
  rows.forEach(row => {
    if (!byOutlet.has(row.OutletCode)) byOutlet.set(row.OutletCode, row);
  });

  return {
    rows: [...byOutlet.values()],
    byOutlet,
    sourceFile: snapshot.fileName || "Zone Distribution snapshot",
    updatedAt: snapshot.savedAt || record?.updated_at || null,
  };
}
