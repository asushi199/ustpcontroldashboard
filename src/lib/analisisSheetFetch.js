/**
 * Muat CSV untuk blok Analisis Data (gid berbeza atau fail demo dalam public/data).
 */
export function analisisSheetCsvUrl(id, gid) {
  const base = import.meta.env.BASE_URL || "/";
  const root = base.endsWith("/") ? base.slice(0, -1) : base;
  const cb = Date.now();
  return `${root}/api/sheet-csv?id=${encodeURIComponent(id)}&gid=${encodeURIComponent(gid)}&_cb=${cb}`;
}

function assertCsvNotHtml(text) {
  const s = String(text ?? "").trimStart();
  if (s.startsWith("<") || s.toLowerCase().includes("<!doctype")) {
    throw new Error(
      "Google memulangkan HTML (bukan CSV). Kongsi spreadsheet: Sesiapa yang mempunyai pautan boleh melihat (Pelihat), kemudian refresh.",
    );
  }
}

/**
 * @param {{ gid?: string, demoPath?: string }} opts
 * @returns {Promise<string>}
 */
export async function fetchAnalisisCsvText(opts) {
  const id = String(import.meta.env.VITE_GOOGLE_SHEET_ID ?? "").trim();
  const gid = opts.gid != null ? String(opts.gid).trim() : "";

  if (id && gid !== "") {
    const r = await fetch(analisisSheetCsvUrl(id, gid));
    if (!r.ok) throw new Error(`Gagal muat CSV (${r.status})`);
    const text = await r.text();
    assertCsvNotHtml(text);
    return text;
  }

  if (opts.demoPath) {
    const base = import.meta.env.BASE_URL || "/";
    const root = base.endsWith("/") ? base : `${base}/`;
    const path = opts.demoPath.replace(/^\//, "");
    const r = await fetch(`${root}${path}`);
    if (!r.ok) throw new Error(`Fail demo ${path} (${r.status})`);
    return r.text();
  }

  throw new Error("Tiada VITE_GOOGLE_SHEET_ID+gid atau demoPath");
}
