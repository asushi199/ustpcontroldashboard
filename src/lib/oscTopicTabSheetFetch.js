import { fetchAnalisisCsvText } from "./analisisSheetFetch.js";
import { parseOscSheetCsv } from "./oscSheetCsv.js";

/** @type {Map<string, import("./oscSheetCsv.js").OscSheetCardRow[]>} */
const cacheByKey = new Map();
/** @type {Map<string, Promise<import("./oscSheetCsv.js").OscSheetCardRow[]>>} */
const inflightByKey = new Map();

/**
 * Muat CSV satu tab topik OSC.
 * - Ada `VITE_GOOGLE_SHEET_ID` + `gid` → tab Google Sheet (tiada lajur `section`).
 * - **Tiada** `VITE_GOOGLE_SHEET_ID` + `demoPath` → fail CSV templat dalam `public/` (contoh pembangunan).
 *
 * @param {{ gid?: string, demoPath?: string }} opts
 * @returns {Promise<import("./oscSheetCsv.js").OscSheetCardRow[]>}
 */
export function loadOscTopicTabRows(opts = {}) {
  const id = import.meta.env.VITE_GOOGLE_SHEET_ID;
  const gid = String(opts.gid ?? "").trim();
  const demoPath = opts.demoPath
    ? String(opts.demoPath).trim().replace(/^\//, "")
    : "";

  const useSheet = Boolean(id && gid);
  const useDemo = Boolean(demoPath) && !id;
  if (!useSheet && !useDemo) return Promise.resolve([]);

  const cacheKey = useSheet ? `g:${gid}` : `d:${demoPath}`;
  const cached = cacheByKey.get(cacheKey);
  if (cached) return Promise.resolve(cached);

  const existing = inflightByKey.get(cacheKey);
  if (existing) return existing;

  const p = (useSheet
    ? fetchAnalisisCsvText({ gid })
    : fetchAnalisisCsvText({ demoPath })
  )
    .then((text) => parseOscSheetCsv(text))
    .then((rows) => {
      cacheByKey.set(cacheKey, rows);
      return rows;
    })
    .finally(() => {
      inflightByKey.delete(cacheKey);
    });

  inflightByKey.set(cacheKey, p);
  return p;
}

export function clearOscTopicTabSheetCache() {
  cacheByKey.clear();
}
