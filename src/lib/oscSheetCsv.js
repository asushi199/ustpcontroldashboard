/**
 * Parse CSV dengan sokongan medan berpetik ("...") — satu baris satu rekod.
 * @returns {string[][]}
 */
export function parseCsvMatrix(csvText) {
  const text = String(csvText ?? "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let cur = "";
  let i = 0;
  let inQuotes = false;

  const pushCell = () => {
    row.push(cur);
    cur = "";
  };
  const pushRow = () => {
    if (row.length === 1 && row[0] === "" && rows.length === 0) return;
    if (row.some((c) => String(c).trim() !== "") || row.length > 1) rows.push(row);
    row = [];
  };

  while (i < text.length) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          cur += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i += 1;
        continue;
      }
      cur += c;
      i += 1;
      continue;
    }
    if (c === '"') {
      inQuotes = true;
      i += 1;
      continue;
    }
    if (c === ",") {
      pushCell();
      i += 1;
      continue;
    }
    if (c === "\n") {
      pushCell();
      pushRow();
      i += 1;
      continue;
    }
    if (c === "\r") {
      i += 1;
      if (text[i] === "\n") i += 1;
      pushCell();
      pushRow();
      continue;
    }
    cur += c;
    i += 1;
  }
  pushCell();
  if (row.some((c) => String(c).trim() !== "")) pushRow();

  return rows;
}

const HEADER_ALIASES = {
  page: ["page", "halaman", "laman"],
  section: ["section", "bahagian", "bidang"],
  subtopik_key: ["subtopik_key", "subtopik key", "bab", "chapter"],
  subtopik_title: ["subtopik_title", "subtopik title", "tajuk subtopik"],
  subtopik_sort: ["subtopik_sort", "subtopik sort", "susunan subtopik"],
  subtopik_blurb: ["subtopik_blurb", "subtopik catatan", "ringkasan subtopik"],
  subtopik_icon: ["subtopik_icon", "subtopik icon", "ikon subtopik", "icon", "ikon", "emoji"],
  sort: ["sort", "order", "susunan", "no"],
  title: ["title", "tajuk", "nama"],
  url: ["url", "link", "pautan"],
  type: ["type", "jenis", "format"],
  blurb: ["blurb", "catatan", "nota", "keterangan"],
  key: ["key", "id"],
  preview_url: [
    "preview_url",
    "preview url",
    "preview",
    "thumbnail",
    "imej pratonton",
    "pratonton",
    "imej",
  ],
};

function normHeader(h) {
  return String(h ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function pickColumnIndex(headers, field) {
  const aliases = HEADER_ALIASES[field] ?? [field];
  for (let ci = 0; ci < headers.length; ci++) {
    const h = normHeader(headers[ci]);
    if (aliases.includes(h)) return ci;
  }
  return -1;
}

/**
 * @typedef {{ key: string, section: string, sort: number, title: string, url: string, type: string, blurb?: string, previewUrl?: string, page?: string, subtopikKey?: string, subtopikTitle?: string, subtopikSort?: number, subtopikBlurb?: string, subtopikIcon?: string }} OscSheetCardRow
 */

/**
 * @param {string} csvText
 * @returns {OscSheetCardRow[]}
 */
export function parseOscSheetCsv(csvText) {
  const matrix = parseCsvMatrix(csvText);
  if (matrix.length < 2) return [];

  const idx = {
    page: pickColumnIndex(matrix[0], "page"),
    section: pickColumnIndex(matrix[0], "section"),
    subtopik_key: pickColumnIndex(matrix[0], "subtopik_key"),
    subtopik_title: pickColumnIndex(matrix[0], "subtopik_title"),
    subtopik_sort: pickColumnIndex(matrix[0], "subtopik_sort"),
    subtopik_blurb: pickColumnIndex(matrix[0], "subtopik_blurb"),
    subtopik_icon: pickColumnIndex(matrix[0], "subtopik_icon"),
    sort: pickColumnIndex(matrix[0], "sort"),
    title: pickColumnIndex(matrix[0], "title"),
    url: pickColumnIndex(matrix[0], "url"),
    type: pickColumnIndex(matrix[0], "type"),
    blurb: pickColumnIndex(matrix[0], "blurb"),
    key: pickColumnIndex(matrix[0], "key"),
    preview_url: pickColumnIndex(matrix[0], "preview_url"),
  };

  if (idx.title < 0 || idx.url < 0) return [];

  const out = [];
  for (let ri = 1; ri < matrix.length; ri++) {
    const cells = matrix[ri];
    const title = String(cells[idx.title] ?? "").trim();
    const url = String(cells[idx.url] ?? "").trim();
    if (!title || !url) continue;

    const page = idx.page >= 0 ? String(cells[idx.page] ?? "").trim() : "";
    const section =
      idx.section >= 0 ? String(cells[idx.section] ?? "").trim() : "";
    const subtopikKey =
      idx.subtopik_key >= 0
        ? String(cells[idx.subtopik_key] ?? "").trim()
        : "";
    const subtopikTitle =
      idx.subtopik_title >= 0
        ? String(cells[idx.subtopik_title] ?? "").trim()
        : "";
    let subtopikSort = 999;
    if (idx.subtopik_sort >= 0) {
      const sn = Number(String(cells[idx.subtopik_sort] ?? "").trim());
      if (Number.isFinite(sn)) subtopikSort = sn;
    }
    const subtopikBlurb =
      idx.subtopik_blurb >= 0
        ? String(cells[idx.subtopik_blurb] ?? "").trim()
        : "";
    const subtopikIconRaw =
      idx.subtopik_icon >= 0
        ? String(cells[idx.subtopik_icon] ?? "").trim()
        : "";
    let sort = 999;
    if (idx.sort >= 0) {
      const n = Number(String(cells[idx.sort] ?? "").trim());
      if (Number.isFinite(n)) sort = n;
    }
    const typeRaw =
      idx.type >= 0 ? String(cells[idx.type] ?? "").trim().toLowerCase() : "";
    const type =
      typeRaw === "gdoc" ||
      typeRaw === "docs" ||
      typeRaw === "google doc" ||
      typeRaw === "google_docs"
        ? "gdoc"
        : typeRaw === "canva"
          ? "canva"
          : typeRaw === "youtube" || typeRaw === "yt"
            ? "youtube"
            : typeRaw === "image" || typeRaw === "img" || typeRaw === "gambar"
              ? "image"
              : typeRaw === "embed" ||
                  typeRaw === "iframe" ||
                  typeRaw === "looker" ||
                  typeRaw === "looker studio"
                ? "embed"
                : "pdf";

    const blurb =
      idx.blurb >= 0
        ? String(cells[idx.blurb] ?? "").trim()
        : type === "pdf"
          ? "PDF (Google Drive)"
          : "";

    const keyCol = idx.key >= 0 ? String(cells[idx.key] ?? "").trim() : "";
    const key =
      keyCol ||
      `row-${ri}-${title.slice(0, 24).replace(/\s+/g, "-")}`;
    const previewRaw =
      idx.preview_url >= 0
        ? String(cells[idx.preview_url] ?? "").trim()
        : "";

    /** @type {OscSheetCardRow} */
    const row = {
      key,
      section,
      sort,
      title,
      url,
      type,
      blurb: blurb || undefined,
    };
    if (previewRaw) row.previewUrl = previewRaw;
    if (page) row.page = page;
    if (subtopikKey) row.subtopikKey = subtopikKey;
    if (subtopikTitle) row.subtopikTitle = subtopikTitle;
    if (subtopikSort !== 999 || idx.subtopik_sort >= 0) {
      row.subtopikSort = subtopikSort;
    }
    if (subtopikBlurb) row.subtopikBlurb = subtopikBlurb;
    if (subtopikIconRaw) row.subtopikIcon = subtopikIconRaw;
    out.push(row);
  }

  return out;
}
