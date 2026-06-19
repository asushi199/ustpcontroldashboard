import { parseCsvMatrix } from "./oscSheetCsv.js";

function num(v) {
  const n = Number(String(v ?? "").replace(",", ".").trim());
  return Number.isFinite(n) ? n : null;
}

/**
 * Cantum lajur nilai selepas lajur kunci. Buang lajur kosong di hujung dulu
 * supaya jadual berbilang lajur (cth ada lajur C/D kosong) tidak hasilkan koma
 * berganda di hujung nilai (cth "Buka sumber,,").
 * @param {string[]} row
 * @returns {string}
 */
function joinValueCells(row) {
  const cells = row.slice(1);
  while (cells.length && String(cells[cells.length - 1] ?? "").trim() === "") {
    cells.pop();
  }
  return cells.join(",").trim();
}

export function parseKeyValueCsv(csvText) {
  const m = parseCsvMatrix(csvText);
  const out = {};
  for (let i = 0; i < m.length; i++) {
    const row = m[i];
    if (!row?.length) continue;
    const k = String(row[0] ?? "").trim();
    if (!k || k.toLowerCase() === "key") continue;
    const v = joinValueCells(row);
    if (v) out[k] = v;
  }
  return out;
}

/** Kunci pilihan (key,value) untuk paparan ringkas seperti dashboard Excel — isi dari jadual/pivot. */
const DELIMA_INSIGHT_CONFIG_KEYS = new Set([
  "bil_sekolah",
  "schools",
  "jumlah_sekolah",
  "khidmat_bantu_kali",
  "bantu_sessions_total",
  "bantu_kali",
  "khidmat_bantu_sekolah",
  "bantu_sekolah_ada_rekod",
  "bantu_schools",
  "dis_capai_guru_bil",
  "dis_capai_guru",
  "schools_guru_gte_tov",
  "capai_guru_bil",
  "dis_capai_murid_bil",
  "dis_capai_murid",
  "schools_murid_gte_tov",
  "capai_murid_bil",
  "avg_tov_guru",
  "avg_tov_murid",
  "avg_nov_guru",
  "avg_nov_murid",
  "avg_dis_guru",
  "avg_dis_murid",
  "avg_guru_aktif_pct",
  "avg_guru_aktif",
  "guru_aktif_pct",
  "purata_guru_aktif",
  "kpi_guru_capai",
  "kpi_murid_capai",
]);

function configKeyLower(config, keys) {
  const lowerMap = {};
  for (const [k, v] of Object.entries(config)) {
    lowerMap[String(k).trim().toLowerCase()] = v;
  }
  for (const key of keys) {
    const raw = lowerMap[key.toLowerCase()];
    if (raw != null && String(raw).trim() !== "") return raw;
  }
  return undefined;
}

function firstNumFromConfig(config, keys) {
  for (const key of keys) {
    const raw = configKeyLower(config, [key]);
    if (raw === undefined) continue;
    const n = num(raw);
    if (n != null) return n;
  }
  return null;
}

function boolishFromConfig(config, keys) {
  for (const key of keys) {
    const raw = configKeyLower(config, [key]);
    if (raw === undefined) continue;
    const s = String(raw).trim().toLowerCase();
    if (s === "1" || s === "yes" || s === "true" || s === "y" || s === "capai") return true;
    if (s === "0" || s === "no" || s === "false" || s === "n" || s === "belum") return false;
  }
  return null;
}

/** Nilai untuk baris key,value dari Sheet — guna lajur B sahaja (elak koma dalam intro pecah CSV). */
function delimaConfigValueFromRow(row) {
  if (!row?.length) return "";
  for (let j = 1; j < row.length; j++) {
    const t = String(row[j] ?? "").trim();
    if (t) return t;
  }
  return "";
}

function delimaInsightFromConfig(config, kpiGuru, kpiMurid) {
  const hasAny = Object.keys(config).some((k) =>
    DELIMA_INSIGHT_CONFIG_KEYS.has(String(k).trim().toLowerCase()),
  );
  if (!hasAny) return null;

  const schools = firstNumFromConfig(config, [
    "bil_sekolah",
    "schools",
    "jumlah_sekolah",
  ]);
  const bantuSessionsTotal = firstNumFromConfig(config, [
    "khidmat_bantu_kali",
    "bantu_sessions_total",
    "bantu_kali",
  ]);
  const bantuSchoolsWithRekod = firstNumFromConfig(config, [
    "khidmat_bantu_sekolah",
    "bantu_sekolah_ada_rekod",
    "bantu_schools",
  ]);
  const schoolsGuruGteTov = firstNumFromConfig(config, [
    "dis_capai_guru_bil",
    "dis_capai_guru",
    "schools_guru_gte_tov",
    "capai_guru_bil",
  ]);
  const schoolsMuridGteTov = firstNumFromConfig(config, [
    "dis_capai_murid_bil",
    "dis_capai_murid",
    "schools_murid_gte_tov",
    "capai_murid_bil",
  ]);
  const avgTovGuru = firstNumFromConfig(config, ["avg_tov_guru"]);
  const avgTovMurid = firstNumFromConfig(config, ["avg_tov_murid"]);
  const avgNovGuru = firstNumFromConfig(config, ["avg_nov_guru"]);
  const avgNovMurid = firstNumFromConfig(config, ["avg_nov_murid"]);
  const avgDisGuru = firstNumFromConfig(config, ["avg_dis_guru"]);
  const avgDisMurid = firstNumFromConfig(config, ["avg_dis_murid"]);
  const avgAktiviti = firstNumFromConfig(config, [
    "avg_guru_aktif_pct",
    "avg_guru_aktif",
    "guru_aktif_pct",
    "purata_guru_aktif",
  ]);

  const kpiGuruExplicit = boolishFromConfig(config, ["kpi_guru_capai"]);
  const kpiMuridExplicit = boolishFromConfig(config, ["kpi_murid_capai"]);
  let kpiGuruOk = false;
  let kpiMuridOk = false;
  if (kpiGuruExplicit != null) kpiGuruOk = kpiGuruExplicit;
  else if (avgDisGuru != null) kpiGuruOk = avgDisGuru + 1e-6 >= kpiGuru;
  if (kpiMuridExplicit != null) kpiMuridOk = kpiMuridExplicit;
  else if (avgDisMurid != null) kpiMuridOk = avgDisMurid + 1e-6 >= kpiMurid;

  return {
    schools,
    bantuSessionsTotal,
    bantuSchoolsWithRekod,
    schoolsGuruGteTov,
    schoolsMuridGteTov,
    avgTovGuru,
    avgTovMurid,
    avgNovGuru,
    avgNovMurid,
    avgDisGuru,
    avgDisMurid,
    avgAktiviti,
    kpiGuruOk,
    kpiMuridOk,
    showKpiGuruPill: kpiGuruExplicit != null || avgDisGuru != null,
    showKpiMuridPill: kpiMuridExplicit != null || avgDisMurid != null,
  };
}

/** Jika tiada baris key insight dalam CSV, isi sebahagian daripada baris Nov / Dis (label tepat). */
function delimaInsightFromMonthRows(months, kpiGuru, kpiMurid) {
  if (!months?.length) return null;
  const pick = (lab) =>
    months.find((m) => String(m.label ?? "").trim().toLowerCase() === lab);
  const nov = pick("nov");
  const dis = pick("dis");
  if (!nov && !dis) return null;
  const avgNovGuru = nov?.guru ?? null;
  const avgNovMurid = nov?.murid ?? null;
  const avgDisGuru = dis?.guru ?? null;
  const avgDisMurid = dis?.murid ?? null;
  if (
    avgNovGuru == null &&
    avgNovMurid == null &&
    avgDisGuru == null &&
    avgDisMurid == null
  ) {
    return null;
  }
  const kpiGuruOk = avgDisGuru != null && avgDisGuru + 1e-6 >= kpiGuru;
  const kpiMuridOk = avgDisMurid != null && avgDisMurid + 1e-6 >= kpiMurid;
  return {
    schools: null,
    bantuSessionsTotal: null,
    bantuSchoolsWithRekod: null,
    schoolsGuruGteTov: null,
    schoolsMuridGteTov: null,
    avgTovGuru: null,
    avgTovMurid: null,
    avgNovGuru,
    avgNovMurid,
    avgDisGuru,
    avgDisMurid,
    avgAktiviti: null,
    kpiGuruOk,
    kpiMuridOk,
    showKpiGuruPill: avgDisGuru != null,
    showKpiMuridPill: avgDisMurid != null,
  };
}

export function parseDelimaAnalisisCsv(csvText) {
  const m = parseCsvMatrix(csvText);
  const config = {};
  const months = [];

  let mi = -1;
  let khidmatBantuSeq = 0;
  for (let i = 0; i < m.length; i++) {
    const row = m[i];
    if (!row?.length) continue;
    const h0 = String(row[0] ?? "").trim().toLowerCase();
    if (h0 === "month_label" && row.length >= 4) {
      mi = i + 1;
      break;
    }
    if (h0 === "key" && String(row[1] ?? "").toLowerCase() === "value") continue;
    const keyRaw = String(row[0] ?? "")
      .trim()
      .replace(/^\uFEFF/, "");
    const v = delimaConfigValueFromRow(row);
    if (!keyRaw || !v) continue;
    /** Dua baris `khidmat_bantu` dalam Sheet (kali, kemudian bil. sekolah) — elak overwrite satu nilai. */
    if (keyRaw.toLowerCase() === "khidmat_bantu") {
      khidmatBantuSeq += 1;
      if (khidmatBantuSeq === 1) config.khidmat_bantu_kali = v;
      else if (khidmatBantuSeq === 2) config.khidmat_bantu_sekolah = v;
      continue;
    }
    config[keyRaw] = v;
  }

  if (mi >= 0) {
    for (let i = mi; i < m.length; i++) {
      const row = m[i];
      if (!row?.length) continue;
      const label = String(row[0] ?? "").trim();
      if (!label || label.toLowerCase() === "month_label") continue;
      const guru = num(row[1]);
      const murid = num(row[2]);
      const inc = String(row[3] ?? "yes").trim().toLowerCase();
      const chartLabel = String(row[4] ?? "").trim();
      const includeChart = inc === "yes" || inc === "true" || inc === "1" || inc === "y";
      months.push({ label, guru, murid, includeChart, chartLabel: chartLabel || undefined });
    }
  }

  const seriesDisplay = months.map((mo) => ({
    label: mo.label,
    chartLabel: mo.chartLabel,
    guru: mo.guru,
    murid: mo.murid,
  }));
  const seriesForChart = months.filter((mo) => mo.includeChart).map((mo) => ({
    label: mo.label,
    chartLabel: mo.chartLabel,
    guru: mo.guru,
    murid: mo.murid,
  }));

  const kpiGuru = num(config.kpi_guru) ?? 78;
  const kpiMurid = num(config.kpi_murid) ?? 65;

  let insight = delimaInsightFromConfig(config, kpiGuru, kpiMurid);
  if (!insight) insight = delimaInsightFromMonthRows(months, kpiGuru, kpiMurid);

  return {
    sourceUrl: config.source_url || "",
    kpiGuru,
    kpiMurid,
    intro: config.intro || "",
    months,
    seriesDisplay,
    seriesForChart,
    insight,
  };
}

export function parseDcsAnalisisCsv(csvText) {
  const kv = parseKeyValueCsv(csvText);
  return {
    tov: num(kv.tov) ?? num(kv.tov_2024) ?? 65,
    kpi: num(kv.kpi) ?? num(kv.kpi_kebangsaan) ?? 78,
    capai: num(kv.capai) ?? num(kv.capai_2025) ?? 78.24,
    yMin: num(kv.y_min) ?? 58,
    yMax: num(kv.y_max) ?? 82,
    updatedText: kv.updated_text || kv.kemaskini || "",
    footer: (kv.footer || kv.footer_text || "").replace(/\\n/g, "\n"),
    sourceUrl: kv.source_url || "",
    sourceLabel: kv.source_label || "Buka sumber data",
    imageUrl: kv.image_url || "",
    imageLabel: kv.image_label || "Buka gambar penuh",
  };
}

export function parseAinsAnalisisCsv(csvText) {
  const kv = parseKeyValueCsv(csvText);
  return {
    approved: num(kv.approved) ?? 0,
    rejected: num(kv.rejected) ?? 0,
    skApproved: num(kv.sk_approved) ?? num(kv.sk) ?? 0,
    sjkcApproved: num(kv.sjkc_approved) ?? num(kv.sjkc) ?? 0,
    sjktApproved: num(kv.sjkt_approved) ?? num(kv.sjkt) ?? 0,
    sourceUrl: kv.source_url || "",
    sourceLabel: kv.source_label || "Buka Google Sheet (data penuh)",
  };
}

export function parsePensijilanAnalisisCsv(csvText) {
  const m = parseCsvMatrix(csvText);
  const meta = {};
  const locations = [];
  const schools = [];
  for (const row of m) {
    if (!row?.length) continue;
    const c0 = String(row[0] ?? "").trim().toLowerCase();
    if (c0 === "category") {
      const kind = String(row[1] ?? "").trim().toLowerCase();
      const label = String(row[2] ?? "").trim();
      const count = num(row[3]);
      if (!label || count == null) continue;
      if (kind === "location" || kind === "lokasi") locations.push([label, count]);
      else if (kind === "school" || kind === "sekolah" || kind === "jenis") schools.push([label, count]);
      continue;
    }
    if (c0 === "key" && String(row[1] ?? "").toLowerCase() === "value") continue;
    if (row.length >= 2) {
      const k = String(row[0] ?? "").trim();
      const v = joinValueCells(row);
      if (k && v) meta[k] = v;
    }
  }
  const kv = parseKeyValueCsv(csvText);
  for (const [k, v] of Object.entries(kv)) {
    if (!(k in meta)) meta[k] = v;
  }
  return {
    intro: meta.intro || meta.keterangan || "",
    sourceUrl: meta.source_url || "",
    sourceLabel: meta.source_label || "Buka sumber / Sheet",
    imageUrl: meta.image_url || "",
    imageModalLabel: meta.image_label || "Lihat infografik penuh",
    locations: locations.length ? locations : null,
    schools: schools.length ? schools : null,
  };
}

export function parseOptikAnalisisCsv(csvText) {
  const kv = parseKeyValueCsv(csvText);
  return {
    kpiKebangsaan: num(kv.kpi_kebangsaan) ?? 55,
    asAt: kv.as_at || kv.asat || "",
    selesaiPct: num(kv.selesai_pct) ?? 0,
    selesaiBil: num(kv.selesai_bil) ?? 0,
    belumPct: num(kv.belum_pct) ?? 0,
    belumBil: num(kv.belum_bil) ?? 0,
    tov2024: num(kv.tov2024) ?? num(kv.tov_2024) ?? 0,
    ar1Julai: num(kv.ar1_julai) ?? num(kv.ar1) ?? 0,
    ar2Okt: num(kv.ar2_okt) ?? num(kv.ar2) ?? 0,
    footerNote: (kv.footer_note || "").replace(/\\n/g, "\n"),
    sourceUrl: kv.source_url || "",
    sourceLabel: kv.source_label || "Buka sumber data",
  };
}
