import { fetchAnalisisCsvText } from "./analisisSheetFetch.js";
import { parseOscSheetCsv } from "./oscSheetCsv.js";

let cachedRows = null;
let inflightPromise = null;

/** Muat CSV tab Bahan Sokongan (gid berasingan). Tanpa ID+gid → [] (guna UI statik). */
export function loadBahanSokonganSheetRows() {
  const id = import.meta.env.VITE_GOOGLE_SHEET_ID;
  const gid = String(import.meta.env.VITE_BAHAN_SOKONGAN_GID ?? "").trim();
  if (!id || !gid) {
    return Promise.resolve([]);
  }

  if (cachedRows) return Promise.resolve(cachedRows);
  if (inflightPromise) return inflightPromise;

  inflightPromise = fetchAnalisisCsvText({
    gid,
    demoPath: "data/bahan-sokongan-demo.csv",
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

export function clearBahanSokonganSheetCache() {
  cachedRows = null;
}
