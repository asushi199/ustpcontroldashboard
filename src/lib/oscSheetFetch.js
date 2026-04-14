import { parseOscSheetCsv } from "./oscSheetCsv.js";

/** @typedef {import("./oscSheetCsv.js").OscSheetCardRow} OscSheetCardRow */

let cachedRows = null;
let inflightPromise = null;

function sheetCsvRequestUrl() {
  const id = import.meta.env.VITE_GOOGLE_SHEET_ID;
  if (!id) return "";
  const gid = import.meta.env.VITE_GOOGLE_SHEET_GID ?? "0";
  const base = import.meta.env.BASE_URL || "/";
  const root = base.endsWith("/") ? base.slice(0, -1) : base;
  return `${root}/api/sheet-csv?id=${encodeURIComponent(id)}&gid=${encodeURIComponent(gid)}`;
}

/**
 * Muat semua baris kad OSC daripada Google Sheet (CSV).
 * @returns {Promise<OscSheetCardRow[]>}
 */
export function loadOscSheetRows() {
  if (cachedRows) return Promise.resolve(cachedRows);
  if (inflightPromise) return inflightPromise;

  const url = sheetCsvRequestUrl();
  if (!url) {
    return Promise.resolve([]);
  }

  inflightPromise = fetch(url)
    .then((r) => {
      if (!r.ok) {
        throw new Error(`Gagal muat CSV (${r.status})`);
      }
      return r.text();
    })
    .then((text) => parseOscSheetCsv(text))
    .then((rows) => {
      cachedRows = rows;
      return rows;
    })
    .finally(() => {
      inflightPromise = null;
    });

  return inflightPromise;
}

/** Untikan ujian / hot reload */
export function clearOscSheetCache() {
  cachedRows = null;
}

/**
 * Tapis mengikut nilai lajur `section` (padanan tepat, case-sensitive seperti dalam sheet).
 * @param {OscSheetCardRow[]} rows
 * @param {string} sectionKey
 */
export function filterOscRowsBySection(rows, sectionKey) {
  const k = String(sectionKey ?? "").trim();
  if (!k) return [];
  return rows
    .filter((r) => String(r.section ?? "").trim() === k)
    .sort((a, b) => a.sort - b.sort || a.title.localeCompare(b.title, "ms"));
}
