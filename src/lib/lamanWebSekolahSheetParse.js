import { parseCsvMatrix } from "./oscSheetCsv.js";

function normHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

/**
 * CSV tab khusus laman web sekolah (bukan format OSC kad).
 * Lajur dibenarkan (mana-mana padanan): kod sekolah / code, nama sekolah / name, url / pautan / pautan laman web
 *
 * @param {string} csvText
 * @returns {{ code: string, name: string, website: string }[]}
 */
export function parseLamanWebSekolahSheetCsv(csvText) {
  const matrix = parseCsvMatrix(csvText);
  if (matrix.length < 2) return [];

  const headers = matrix[0].map(normHeader);
  const idx = (aliases) => {
    for (const a of aliases) {
      const i = headers.findIndex((h) => h === a || h.startsWith(a));
      if (i >= 0) return i;
    }
    return -1;
  };

  const iCode = idx(["kod sekolah", "code", "kod"]);
  const iName = idx(["nama sekolah", "name", "nama"]);
  const iUrl = idx(["pautan laman web", "url", "pautan", "link"]);

  const out = [];
  for (let ri = 1; ri < matrix.length; ri++) {
    const row = matrix[ri];
    const code = iCode >= 0 ? String(row[iCode] ?? "").trim() : "";
    const name = iName >= 0 ? String(row[iName] ?? "").trim() : "";
    const urlRaw = iUrl >= 0 ? String(row[iUrl] ?? "").trim() : "";
    if (!code && !name) continue;
    out.push({ code, name, website: urlRaw });
  }
  return out;
}
