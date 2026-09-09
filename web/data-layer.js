/*
 * data-layer.js
 *
 * Replaces whatever currently refetches on every date change.
 *
 * Rules this enforces:
 *   - a date change NEVER refetches days already in memory
 *   - only genuinely new days are fetched, in parallel, with a cap
 *   - a new range aborts the previous in-flight batch (this alone stops the
 *     endless spinner caused by stacked requests)
 *   - shards are cached in IndexedDB keyed by snapshot_id, so a repeat visit
 *     paints instantly and only then checks for a newer snapshot
 *
 * Wire-up:
 *   const dl = new DataLayer('./data/snapshot');
 *   await dl.init();
 *   const rows = await dl.range('2026-08-09', '2026-09-08');
 *   const kpis = dl.aggregate(rows, { division: 'X' });
 */

const DB_NAME = 'receiving-snapshot';
const STORE = 'shards';
const MAX_PARALLEL = 8;

function idb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      if (!req.result.objectStoreNames.contains(STORE)) {
        req.result.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function cacheGet(key) {
  try {
    const db = await idb();
    return await new Promise((resolve) => {
      const req = db.transaction(STORE, 'readonly').objectStore(STORE).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch { return null; }
}

async function cachePut(key, value) {
  try {
    const db = await idb();
    db.transaction(STORE, 'readwrite').objectStore(STORE).put(value, key);
  } catch { /* cache is an optimisation, never a dependency */ }
}

async function cacheDropOtherSnapshots(snapshotId) {
  try {
    const db = await idb();
    const store = db.transaction(STORE, 'readwrite').objectStore(STORE);
    const req = store.getAllKeys();
    req.onsuccess = () => {
      for (const key of req.result) {
        if (!String(key).startsWith(snapshotId + '|')) store.delete(key);
      }
    };
  } catch { /* ignore */ }
}

function eachDate(from, to) {
  const out = [];
  const cursor = new Date(from + 'T00:00:00Z');
  const end = new Date(to + 'T00:00:00Z');
  while (cursor <= end) {
    out.push(cursor.toISOString().slice(0, 10));
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return out;
}

export class DataLayer {
  constructor(base = './data/snapshot') {
    this.base = base.replace(/\/$/, '');
    this.manifest = null;
    this.dims = null;
    this.days = new Map();       // 'YYYY-MM-DD' -> array of rows
    this.available = new Set();  // days the snapshot actually has
    this.inflight = null;        // AbortController for the current batch
  }

  async init() {
    const res = await fetch(`${this.base}/manifest.json`, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`manifest ${res.status}`);
    this.manifest = await res.json();
    this.available = new Set(this.manifest.shards.map(s => s.date));

    const dimsKey = `${this.manifest.snapshot_id}|dims`;
    let dims = await cacheGet(dimsKey);
    if (!dims) {
      dims = await (await fetch(`${this.base}/dims.json`)).json();
      cachePut(dimsKey, dims);
    }
    this.dims = dims;

    cacheDropOtherSnapshots(this.manifest.snapshot_id);
    return this.manifest;
  }

  label(kind, id) {
    const table = this.dims?.[kind];
    return table ? (table[id] ?? '(unknown)') : String(id);
  }

  /** Days in range that are neither in memory nor outside the snapshot. */
  missing(from, to) {
    return eachDate(from, to).filter(d => this.available.has(d) && !this.days.has(d));
  }

  async #loadShard(date, signal) {
    const key = `${this.manifest.snapshot_id}|${date}`;
    const cached = await cacheGet(key);
    if (cached) { this.days.set(date, cached); return; }

    const res = await fetch(`${this.base}/daily/${date}.json.gz`, { signal });
    if (!res.ok) { this.days.set(date, []); return; }
    const payload = await res.json();   // server sends Content-Encoding: gzip
    this.days.set(date, payload.data);
    cachePut(key, payload.data);
  }

  /**
   * Returns every pre-aggregated row in the range.
   * Fetches only the days it doesn't already hold; if all are held,
   * this resolves synchronously fast with no network at all.
   */
  async range(from, to, onProgress) {
    const needed = this.missing(from, to);

    if (needed.length) {
      if (this.inflight) this.inflight.abort();
      this.inflight = new AbortController();
      const { signal } = this.inflight;

      let done = 0;
      const queue = [...needed];
      const worker = async () => {
        while (queue.length) {
          const date = queue.shift();
          await this.#loadShard(date, signal);
          done += 1;
          onProgress?.(done, needed.length);
        }
      };
      try {
        await Promise.all(
          Array.from({ length: Math.min(MAX_PARALLEL, needed.length) }, worker)
        );
      } catch (err) {
        if (err.name === 'AbortError') return null;   // superseded by a newer range
        throw err;
      }
      this.inflight = null;
    }

    const cols = this.manifest.dimensions.concat(this.manifest.measures, ['rows']);
    const out = [];
    for (const date of eachDate(from, to)) {
      const rows = this.days.get(date);
      if (rows) for (const row of rows) out.push({ date, row, cols });
    }
    return out;
  }

  /** Sum measures over rows, with optional dimension filters. */
  aggregate(rows, filters = {}) {
    const dims = this.manifest.dimensions;
    const measures = this.manifest.measures;

    const wanted = Object.entries(filters)
      .filter(([, v]) => v !== null && v !== undefined && v !== 'all')
      .map(([k, v]) => [dims.indexOf(k), v]);

    const totals = Object.fromEntries(measures.map(m => [m, 0]));
    totals.rows = 0;

    outer:
    for (const { row } of rows) {
      for (const [idx, value] of wanted) {
        if (idx < 0) continue;
        const cell = this.label(dims[idx], row[idx]);
        if (Array.isArray(value) ? !value.includes(cell) : cell !== value) continue outer;
      }
      measures.forEach((m, i) => { totals[m] += row[dims.length + i]; });
      totals.rows += row[dims.length + measures.length];
    }
    return totals;
  }
}

/** Debounce helper for the date inputs. */
export function debounce(fn, wait = 250) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}
