import { parseCsvMatrix } from "./oscSheetCsv.js";

function normCell(v) {
  return String(v ?? "").trim().toLowerCase();
}

function normHeaderKey(v) {
  return String(v ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

/**
 * Cari baris tajuk jadual pegawai: lajur 1 = nama, lajur 2 = jawatan.
 * @param {string[][]} matrix
 * @returns {number}
 */
function findPegawaiHeaderRowIndex(matrix) {
  for (let i = 0; i < matrix.length; i++) {
    const row = matrix[i];
    if (!row?.length) continue;
    if (normCell(row[0]) === "nama" && normCell(row[1]) === "jawatan") {
      return i;
    }
  }
  return -1;
}

/**
 * @param {string} csvText
 * @returns {{ config: Record<string, string>, pegawaiRows: Array<{ nama: string, jawatan: string, telefon: string, detailUrl: string, detailImage: string }> | null }}
 */
export function parseMaklumatAsasCsv(csvText) {
  const matrix = parseCsvMatrix(csvText);
  const config = {};

  const pegawaiHeaderIdx = findPegawaiHeaderRowIndex(matrix);
  const configEnd = pegawaiHeaderIdx >= 0 ? pegawaiHeaderIdx : matrix.length;

  for (let i = 0; i < configEnd; i++) {
    const row = matrix[i];
    if (!row?.length) continue;
    const k = String(row[0] ?? "").trim();
    const v = row.slice(1).join(",").trim();
    if (!k) continue;
    if (normCell(k) === "key" && normCell(row[1] ?? "") === "value") {
      continue;
    }
    if (v) config[k] = v;
  }

  if (pegawaiHeaderIdx < 0) {
    return { config, pegawaiRows: null };
  }

  const header = matrix[pegawaiHeaderIdx];
  /** @type {Record<string, number>} */
  const col = {};
  for (let c = 0; c < header.length; c++) {
    const h = normHeaderKey(header[c]);
    if (h === "nama") col.nama = c;
    else if (h === "jawatan") col.jawatan = c;
    else if (h === "telefon") col.telefon = c;
    else if (h === "detail_url") col.detailUrl = c;
    else if (h === "detail_image") col.detailImage = c;
  }

  const pegawaiRows = [];
  for (let i = pegawaiHeaderIdx + 1; i < matrix.length; i++) {
    const row = matrix[i];
    if (!row?.length) continue;
    const nama =
      col.nama !== undefined
        ? String(row[col.nama] ?? "").trim()
        : String(row[0] ?? "").trim();
    const jawatan =
      col.jawatan !== undefined
        ? String(row[col.jawatan] ?? "").trim()
        : String(row[1] ?? "").trim();
    if (!nama && !jawatan) continue;
    pegawaiRows.push({
      nama: nama || "—",
      jawatan,
      telefon:
        col.telefon !== undefined
          ? String(row[col.telefon] ?? "").trim()
          : "",
      detailUrl:
        col.detailUrl !== undefined
          ? String(row[col.detailUrl] ?? "").trim()
          : "",
      detailImage:
        col.detailImage !== undefined
          ? String(row[col.detailImage] ?? "").trim()
          : "",
    });
  }

  return { config, pegawaiRows };
}
