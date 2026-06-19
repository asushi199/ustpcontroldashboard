import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { BahanSokonganPageSection } from "./components/BahanSokonganPageSection.jsx";
import { DetailsCollapseFooter } from "./components/DetailsCollapseFooter.jsx";
import { OscTopicSheetBody } from "./components/OscTopicSheetBody.jsx";
import {
  canvaViewEmbedUrl,
  driveGoogleFilePreviewUrl,
  driveGoogleImageUrl,
  googleDocEmbedPreviewUrl,
  maklumatAsasPreviewMode,
} from "./lib/embedUrls.js";
import { fetchAnalisisCsvText } from "./lib/analisisSheetFetch.js";
import {
  parseDelimaAnalisisCsv,
  parseDcsAnalisisCsv,
  parseAinsAnalisisCsv,
  parsePensijilanAnalisisCsv,
  parseOptikAnalisisCsv,
} from "./lib/analisisSheetParse.js";
import { parseMaklumatAsasCsv } from "./lib/maklumatAsasSheetParse.js";
import { parseLamanWebSekolahSheetCsv } from "./lib/lamanWebSekolahSheetParse.js";
import { ItmLamanWebSekolahSection } from "./components/ItmLamanWebSekolahSection.jsx";

/** Logo rasmi PPD (PNG telus) — watermark latar tiga mod reka */
const USTP_WATERMARK_SRC = "/assets/ustp-ppd-manjung-watermark.png";
/** Carta Organisasi & Maklumat PKG / COE — infografik rasmi (Buku Pengurusan USTP) */
const CARTA_ORGANISASI_IMG = "/assets/carta-organisasi-ustp-ppd-manjung.png";
const MAKLUMAT_PKG_COE_IMG = "/assets/maklumat-pkg-coe-daerah-manjung.png";

/** Kalendar kumpulan USTP (zon Asia/Kuala Lumpur) */
const USTP_CALENDAR_EMBED =
  "https://calendar.google.com/calendar/embed?src=c_07ea831973519ec6379185af0a2fd2053aeec6d5c15fab56dc24461b74e5c2e2%40group.calendar.google.com&ctz=Asia%2FKuala_Lumpur";

const LAMAN_WEB_SEKOLAH_XLSX_URL = "/data/laman-web-sekolah-bengkel-responses.xlsx";

/** CSV templat topik OSC (public/data/) — dipapar hanya jika `VITE_GOOGLE_SHEET_ID` kosong; salin struktur ke tab Sheet sebenar. */
const OSC_TOPIK_DEMO_CSV = {
  integrasi: "data/osc-topik-integrasi-demo.csv",
  hebahan: "data/osc-topik-hebahan-demo.csv",
  itm: "data/osc-topik-itm-demo.csv",
  pembudayaanMembaca: "data/osc-topik-pembudayaan-membaca-demo.csv",
  pemerkasaan: "data/osc-topik-pemerkasaan-demo.csv",
};
const DELIMA_GOOGLE_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1pGLkWr8Vt4kPW7haQ-_AZ-trpptimI-5r038vWbdblk/edit?gid=0#gid=0";

/** Paparan pill sasaran guru/murid DELIMa — matikan kerana garis panduan KPM semasa. Tukar kepada true untuk papar semula. */
const DELIMA_SHOW_SASARAN_KPI = false;

/** Modal pegawai: `detailImage` boleh pautan Drive (imej terus) atau imej terus — papar sebagai `<img>` bersih. */
function PegawaiProfilePreview({ pegawai }) {
  if (pegawai.detailUrl) {
    return (
      <iframe
        title={`${pegawai.nama} - personal profile`}
        src={pegawai.detailUrl}
        className="h-[70vh] w-full"
        style={{ border: 0, background: "#0b1220" }}
        allowFullScreen
      />
    );
  }
  if (pegawai.detailImage) {
    // Gambar pegawai sememangnya imej: pautan Drive → URL imej terus (lh3), bukan bingkai Drive.
    const isDrive = /drive\.google\.com/i.test(pegawai.detailImage);
    if (isDrive) {
      return (
        <img
          alt={`${pegawai.nama} - personal profile`}
          src={driveGoogleImageUrl(pegawai.detailImage)}
          className="h-auto w-full"
        />
      );
    }
    const prev = maklumatAsasPreviewMode(pegawai.detailImage);
    if (prev.kind === "iframe") {
      return (
        <iframe
          title={`${pegawai.nama} - personal profile`}
          src={prev.src}
          className="h-[70vh] w-full"
          style={{ border: 0, background: "#0b1220" }}
          allowFullScreen
        />
      );
    }
    return (
      <img
        alt={`${pegawai.nama} - personal profile`}
        src={prev.src}
        className="h-auto w-full"
      />
    );
  }
  return (
    <div className="flex h-[70vh] items-center justify-center p-6 text-center text-sm text-slate-300">
      Belum diset: beri saya pautan embed atau gambar untuk {pegawai.nama}
    </div>
  );
}

function normalizeSchoolWebsiteUrl(raw) {
  const s = String(raw ?? "").trim();
  if (!s) return "";
  if (/^https?:\/\//i.test(s)) return s;
  return `https://${s}`;
}

/** Padankan kod dari Excel (mungkin ada ruang: "AEE 1026"). */
function normalizeLamanWebSchoolCode(code) {
  return String(code ?? "").replace(/\s+/g, "").toUpperCase();
}

/**
 * 8 sekolah paparan utama — susunan SMK → SK → SJKC → SJKT.
 * Jika Excel tiada baris / kod tak sepadan, masih papar 8 kad (nama sandaran + pautan tetap).
 * Tambah varian kod dalam `matchCodes` jika borang guna ejaan lain.
 */
const LAMAN_WEB_FEATURED = [
  {
    code: "AEE1026",
    matchCodes: ["AEE1026"],
    url: "https://sites.google.com/moe-dl.edu.my/smkmethodistacssitiawan/home?authuser=0",
    fallbackName: "SMK METHODIST (ACS) SITIAWAN",
  },
  {
    code: "AEE1032",
    matchCodes: ["AEE1032"],
    url: "https://sites.google.com/moe-dl.edu.my/smkrajashahriman",
    fallbackName: "SMK RAJA SHAHRIMAN",
  },
  {
    code: "ABA1022",
    matchCodes: ["ABA1022"],
    url: "https://sites.google.com/moe-dl.edu.my/sklumut",
    fallbackName: "SK LUMUT",
  },
  {
    code: "ABA1035",
    matchCodes: ["ABA1035"],
    url: "https://sites.google.com/moe-dl.edu.my/skserisamudera",
    fallbackName: "SK SERI SAMUDERA",
  },
  {
    code: "ABC1062",
    matchCodes: ["ABC1062"],
    url: "https://sites.google.com/moe-dl.edu.my/sjkcminte/home",
    fallbackName: "SJKC MIN TE",
  },
  {
    code: "ABC1071",
    matchCodes: ["ABC1071"],
    url: "https://sites.google.com/moe-dl.edu.my/sjkcyenmin?usp=sharing",
    fallbackName: "SJKC YEN MIN",
  },
  {
    code: "ABD1075",
    matchCodes: ["ABD1075"],
    url: "https://sites.google.com/moe-dl.edu.my/sjktpangkor",
    fallbackName: "SJKT PANGKOR",
  },
  {
    code: "ABD1073",
    matchCodes: ["ABD1073"],
    url: "https://sites.google.com/moe-dl.edu.my/sjktmahaganesaviddyasalai/utama",
    fallbackName: "SJKT MAHA GANESA VIDDYASALAI",
  },
];

const LAMAN_WEB_FEATURED_URL_BY_CODE = Object.fromEntries(
  LAMAN_WEB_FEATURED.map((d) => [d.code, d.url]),
);

function findLamanWebRowForFeatured(byNorm, def) {
  for (const c of def.matchCodes) {
    const row = byNorm.get(normalizeLamanWebSchoolCode(c));
    if (row) return row;
  }
  return null;
}

function mergeLamanWebFeaturedRow(def, excelRow) {
  const base = excelRow ?? {
    code: def.code,
    name: def.fallbackName,
    website: "",
  };
  const fromData = normalizeSchoolWebsiteUrl(base.website || "");
  const fallbackUrl = normalizeSchoolWebsiteUrl(def.url);
  return {
    ...base,
    code: base.code?.trim() ? base.code : def.code,
    name: (base.name && String(base.name).trim()) || def.fallbackName,
    website: fromData || fallbackUrl,
  };
}

function lamanWebRowWithFeaturedWebsiteIfAny(r) {
  const k = normalizeLamanWebSchoolCode(r.code);
  if (!k) return r;
  if (String(r.website ?? "").trim()) return r;
  const o = LAMAN_WEB_FEATURED_URL_BY_CODE[k];
  return o ? { ...r, website: normalizeSchoolWebsiteUrl(o) } : r;
}

/** Borang Google: KOD SEKOLAH, NAMA SEKOLAH, PAUTAN LAMAN WEB SEKOLAH (mungkin ada ruang akhir). */
function parseLamanWebSekolahXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName =
    wb.SheetNames.find((n) => /response/i.test(n)) ?? wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const out = [];
  for (const r of rows) {
    const code = String(r["KOD SEKOLAH"] ?? "").trim();
    const name = String(r["NAMA SEKOLAH"] ?? "").trim();
    let urlRaw = "";
    for (const k of Object.keys(r)) {
      if (String(k).trim().toUpperCase().startsWith("PAUTAN LAMAN WEB")) {
        urlRaw = String(r[k] ?? "").trim();
        break;
      }
    }
    if (!code && !name) continue;
    out.push({
      code,
      name,
      website: normalizeSchoolWebsiteUrl(urlRaw),
    });
  }
  out.sort((a, b) => a.code.localeCompare(b.code, "ms", { sensitivity: "base" }));
  return out;
}

function fmtPct1(n) {
  if (n == null || Number.isNaN(n)) return "—";
  return `${n.toFixed(1)}%`;
}

function DelimaDeltaPp({ before, after, title }) {
  if (before == null || after == null) {
    return <span className="text-slate-500">—</span>;
  }
  const d = after - before;
  const up = d > 0.49;
  const down = d < -0.49;
  const icon = up ? "▲" : down ? "▼" : "▬";
  const cls = up ? "text-emerald-400" : down ? "text-rose-400" : "text-slate-400";
  return (
    <span
      className={`inline-flex items-center gap-1 text-xs font-semibold ${cls}`}
      title={title}
    >
      <span aria-hidden>{icon}</span>
      <span>
        {d > 0 ? "+" : ""}
        {d.toFixed(1)} pp
      </span>
    </span>
  );
}

/** Ringkasan angka daerah (isi dari Google Sheet key,value) — panel berlapis seperti dashboard Excel. */
function DelimaSheetInsightBlock({ insight, kpiGuru, kpiMurid }) {
  if (!insight) return null;
  const n = insight.schools;
  const nStr = n != null ? String(Math.round(n)) : "—";

  const hasTiles =
    insight.schools != null ||
    insight.bantuSessionsTotal != null ||
    insight.bantuSchoolsWithRekod != null ||
    insight.schoolsGuruGteTov != null ||
    insight.schoolsMuridGteTov != null;
  const hasDeltas =
    insight.avgTovGuru != null ||
    insight.avgDisGuru != null ||
    insight.avgTovMurid != null ||
    insight.avgDisMurid != null ||
    insight.avgNovGuru != null;
  const hasPills =
    insight.showKpiGuruPill ||
    insight.showKpiMuridPill ||
    insight.avgAktiviti != null;

  return (
    <div className="space-y-3">
      {hasTiles ? (
        <div className="rounded-xl border border-cyan-400/30 bg-slate-950/50 p-3 shadow-[inset_0_1px_0_rgba(34,211,238,0.08)]">
          <div className="grid grid-cols-2 gap-2 text-xs sm:grid-cols-3">
            {insight.schools != null ? (
              <div className="rounded-lg border border-cyan-400/20 bg-slate-900/60 px-2 py-1.5">
                <p className="text-slate-500">Sekolah</p>
                <p className="text-sm font-semibold text-white">{nStr}</p>
              </div>
            ) : null}
            {insight.bantuSessionsTotal != null || insight.bantuSchoolsWithRekod != null ? (
              <div className="rounded-lg border border-cyan-400/20 bg-slate-900/60 px-2 py-1.5">
                <p className="text-slate-500">Khidmat bantu</p>
                <p className="text-sm font-semibold text-cyan-200">
                  {insight.bantuSessionsTotal != null
                    ? `${Math.round(insight.bantuSessionsTotal)} `
                    : "— "}
                  <span className="font-normal text-slate-500">kali</span>
                </p>
                {insight.bantuSchoolsWithRekod != null ? (
                  <p className="mt-0.5 text-[10px] text-slate-500">
                    {Math.round(insight.bantuSchoolsWithRekod)} sekolah ada rekod
                  </p>
                ) : null}
              </div>
            ) : null}
            {insight.schoolsGuruGteTov != null || insight.schoolsMuridGteTov != null ? (
              <div className="col-span-2 rounded-lg border border-cyan-400/20 bg-slate-900/60 px-2 py-1.5 sm:col-span-1">
                <p className="text-slate-500">Capai ≥ TOV (Dis)</p>
                <p className="text-sm font-semibold text-white">
                  G{" "}
                  {insight.schoolsGuruGteTov != null ? Math.round(insight.schoolsGuruGteTov) : "—"}/
                  {nStr}{" "}
                  <span className="text-slate-500">·</span> M{" "}
                  {insight.schoolsMuridGteTov != null ? Math.round(insight.schoolsMuridGteTov) : "—"}/
                  {nStr}
                </p>
              </div>
            ) : null}
          </div>
        </div>
      ) : null}

      {hasDeltas ? (
        <div className="rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2.5 text-xs">
          <div className="space-y-1.5">
            {insight.avgTovGuru != null || insight.avgDisGuru != null ? (
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-slate-400">
                  Guru: TOV {fmtPct1(insight.avgTovGuru)} → Dis {fmtPct1(insight.avgDisGuru)}
                </span>
                <DelimaDeltaPp
                  before={insight.avgTovGuru}
                  after={insight.avgDisGuru}
                  title="Purata daerah: TOV → Disember"
                />
              </div>
            ) : null}
            {insight.avgTovMurid != null || insight.avgDisMurid != null ? (
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-slate-400">
                  Murid: TOV {fmtPct1(insight.avgTovMurid)} → Dis {fmtPct1(insight.avgDisMurid)}
                </span>
                <DelimaDeltaPp
                  before={insight.avgTovMurid}
                  after={insight.avgDisMurid}
                  title="Purata daerah: TOV → Disember"
                />
              </div>
            ) : null}
            {insight.avgNovGuru != null || insight.avgDisGuru != null ? (
              <div className="flex flex-wrap items-center justify-between gap-2">
                <span className="text-slate-400">
                  Guru: Nov {fmtPct1(insight.avgNovGuru)} → Dis {fmtPct1(insight.avgDisGuru)}
                </span>
                <DelimaDeltaPp
                  before={insight.avgNovGuru}
                  after={insight.avgDisGuru}
                  title="Purata daerah: November → Disember"
                />
              </div>
            ) : null}
          </div>
        </div>
      ) : null}

      {hasPills ? (
        <div className="flex flex-wrap gap-2">
          {insight.showKpiGuruPill ? (
            <StatusPill
              label={`Sasaran Dis guru (${kpiGuru}%): ${insight.kpiGuruOk ? "Capai" : "Belum"}`}
              tone={insight.kpiGuruOk ? "good" : "warn"}
            />
          ) : null}
          {insight.showKpiMuridPill ? (
            <StatusPill
              label={`Sasaran Dis murid (${kpiMurid}%): ${insight.kpiMuridOk ? "Capai" : "Belum"}`}
              tone={insight.kpiMuridOk ? "good" : "warn"}
            />
          ) : null}
          {insight.avgAktiviti != null ? (
            <StatusPill
              label={`Guru aktif (bil): ${fmtPct1(insight.avgAktiviti)}`}
              tone="good"
            />
          ) : null}
        </div>
      ) : null}
    </div>
  );
}

/** Carta garis purata daerah: % guru & % murid (label pakai chartLabel jika ada). */
function DelimaTrendChart({ series, height = 200 }) {
  const n = series.length;
  if (n < 1) return null;

  const w = n > 6 ? 620 : 520;
  const h = height;
  const pad = { t: 14, r: 12, b: n > 7 ? 32 : 28, l: 36 };
  const innerW = w - pad.l - pad.r;
  const innerH = h - pad.t - pad.b;

  const xAt = (i) =>
    n === 1 ? pad.l + innerW / 2 : pad.l + (i / (n - 1)) * innerW;

  const xLabelSize = n > 7 ? 7 : 8;

  const linePath = (key) => {
    const pts = [];
    for (let i = 0; i < n; i++) {
      const v = series[i][key];
      if (v == null || !Number.isFinite(v)) continue;
      const x = xAt(i);
      const y = pad.t + (1 - Math.min(100, Math.max(0, v)) / 100) * innerH;
      pts.push({ x, y, v });
    }
    if (pts.length < 2) return null;
    return pts
      .map((p, idx) => `${idx === 0 ? "M" : "L"} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`)
      .join(" ");
  };

  const pathG = linePath("guru");
  const pathM = linePath("murid");

  const dotsForSinglePoint = (key, color) => {
    if (n !== 1) return null;
    const v = series[0][key];
    if (v == null || !Number.isFinite(v)) return null;
    const x = xAt(0);
    const y = pad.t + (1 - Math.min(100, Math.max(0, v)) / 100) * innerH;
    return (
      <circle cx={x} cy={y} r={5} fill={color} fillOpacity={0.9} />
    );
  };

  return (
    <svg
      viewBox={`0 0 ${w} ${h}`}
      className="h-auto w-full max-w-full text-slate-500"
      role="img"
      aria-label="Purata peratus aktif guru dan murid mengikut potongan masa"
    >
      {[0, 25, 50, 75, 100].map((pct) => {
        const y = pad.t + (1 - pct / 100) * innerH;
        return (
          <g key={pct}>
            <line
              x1={pad.l}
              y1={y}
              x2={w - pad.r}
              y2={y}
              stroke="currentColor"
              strokeOpacity={0.12}
              strokeWidth={1}
            />
            <text x={4} y={y + 4} className="fill-slate-500 text-[9px]">
              {pct}%
            </text>
          </g>
        );
      })}
      {pathG ? (
        <path
          d={pathG}
          fill="none"
          stroke="#22d3ee"
          strokeWidth={2.2}
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      ) : null}
      {pathM ? (
        <path
          d={pathM}
          fill="none"
          stroke="#a78bfa"
          strokeWidth={2.2}
          strokeLinecap="round"
          strokeLinejoin="round"
        />
      ) : null}
      {!pathG && n === 1 ? dotsForSinglePoint("guru", "#22d3ee") : null}
      {!pathM && n === 1 ? dotsForSinglePoint("murid", "#a78bfa") : null}
      {series.map((p, i) => (
        <text
          key={p.label}
          x={xAt(i)}
          y={h - 6}
          textAnchor="middle"
          className="fill-slate-400"
          style={{ fontSize: xLabelSize }}
        >
          {p.chartLabel ?? p.label}
        </text>
      ))}
      <text x={pad.l} y={12} className="fill-cyan-300 text-[10px] font-medium">
        Guru (cyan)
      </text>
      <text x={pad.l + 78} y={12} className="fill-violet-300 text-[10px] font-medium">
        Murid (ungu)
      </text>
    </svg>
  );
}

/** Jadual purata daerah — langkau Ogs¹ & Okt¹ (potongan pertama). */
function DelimaMonthAvgTable({ seriesDisplay, maxHeightClass = "max-h-[168px]", showCaption = true }) {
  if (!seriesDisplay.length) {
    return (
      <div className="border-t border-cyan-400/10 pt-2 text-[10px] text-slate-500">
        Tiada data.
      </div>
    );
  }

  return (
    <div className={showCaption ? "border-t border-cyan-400/10 pt-2" : ""}>
      {showCaption ? (
        <p className="mb-1 text-[10px] text-slate-400">
          Purata daerah — <span className="text-cyan-200/90">guru</span> /{" "}
          <span className="text-violet-200/80">murid</span> (% aktif)
        </p>
      ) : null}
      <div
        className={`overflow-y-auto rounded-lg border border-cyan-400/10 bg-slate-950/30 ${maxHeightClass}`}
      >
        <table className="w-full text-left text-[10px]">
          <thead className="sticky top-0 z-[1] bg-slate-950/95 backdrop-blur-sm">
            <tr className="text-slate-500">
              <th className="px-2 py-1 font-medium">Potongan</th>
              <th className="px-2 py-1 font-medium text-cyan-200/85">Guru</th>
              <th className="px-2 py-1 font-medium text-violet-200/75">Murid</th>
            </tr>
          </thead>
          <tbody>
            {seriesDisplay.map((p) => (
              <tr
                key={p.label}
                className="border-t border-cyan-400/10 text-slate-200"
              >
                <td className="px-2 py-0.5">{p.chartLabel ?? p.label}</td>
                <td className="px-2 py-0.5 tabular-nums text-cyan-100/90">
                  {fmtPct1(p.guru)}
                </td>
                <td className="px-2 py-0.5 tabular-nums text-violet-100/85">
                  {fmtPct1(p.murid)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/** Program Pemerkasaan Bacaan Murid — Bicara Buku (YouTube Live / rakaman) */
const BICARA_BUKU_YOUTUBE_WATCH_URL =
  "https://www.youtube.com/live/XmP3d3XwdC4";
const BICARA_BUKU_YOUTUBE_EMBED_URL =
  "https://www.youtube.com/embed/XmP3d3XwdC4";
/** Program Pemerkasaan Bacaan Murid — Video kreatif (Google Docs) */
const PEMERKASAAN_VIDEO_KREATIF_LENSA_EDU_DOC_URL =
  "https://docs.google.com/document/d/1GArnxqpVufIKIzxzjgkYHu5TU_yw9nic7NKsFUeMKeQ/edit?tab=t.0";
/** Program Pemerkasaan Bacaan Murid — Video kreatif IMPAK TVPSS (PDF, Google Drive) */
const PEMERKASAAN_VIDEO_KREATIF_IMPAK_TVPSS_PDF_URL =
  "https://drive.google.com/file/d/1DTDFDW46uPyiTsEvNEUJjJLVMOE2pBuc/view?usp=drive_link";
/** Program Pemerkasaan Bacaan Murid — Pembaca Bestari (Google Docs) */
const PEMERKASAAN_PEMBACA_BESTARI_PENILAIAN_DOC_URL =
  "https://docs.google.com/document/d/1FDPdQjRbPcv7dA-drIpWw0gghPu3OG7xSp2DRwiH8TU/edit?tab=t.0";
const KERTAS_KERJA_EDUSPARK_COE_ROADSHOW_URL =
  "https://www.canva.com/design/DAG0hWYthb4/bzP48XBfMSZMxXUqsL_g3g/view";
/** Kertas kerja program — Bengkel Digital 2025 (PDF, Google Drive) */
const KERTAS_KERJA_BENGKEL_DIGITAL_2025_PDF_URL =
  "https://drive.google.com/file/d/1mLYYq1JLCGeOJFczmSYc_unJLCY3iMdZ/view?usp=drive_link";
/** Integrasi — JNJ */
const JNJ_PANDAI_EDUCATION_DOC_URL =
  "https://docs.google.com/document/d/1Fri298_f1FPov0K3YLfFZdVaauCSSkKoe7f2BKLt4bQ/edit?tab=t.0";
const JNJ_GOOLEE_MUSTREAD_PDF_URL =
  "https://drive.google.com/file/d/1FMmqWJNRU0qHmvV84MpZ5poQRAuCPnUe/view?usp=drive_link";
/** Integrasi — JNJ Minecraft (PDF, Google Drive) */
const JNJ_MINECRAFT_PDF_URL =
  "https://drive.google.com/file/d/1ZzDIs8yhHwzhKktP5CpdkGGmqEkFW1zp/view?usp=drive_link";
/** Integrasi — JNJ Mikrobotik & Edu Kit (PDF, Google Drive) */
const JNJ_MIKROBOTIK_EDUKIT_PDF_URL =
  "https://drive.google.com/file/d/1rzOHBhXG81YtpdwQzvg6i-bXhf47DK9E/view?usp=drive_link";

/** Integrasi Teknologi Pendidikan — Laporan (PDF Google Drive) */
const INTEGRASI_LAPORAN_CARDS = [
  {
    title: "Canva Poster",
    viewUrl:
      "https://drive.google.com/file/d/1MZILesbmM0KUCFwqV3bPLLplECamUCGk/view?usp=drive_link",
  },
  {
    title: "Dron",
    viewUrl:
      "https://drive.google.com/file/d/15nLg2S1yoXCnCpQ6REDrM_qsij4282fP/view?usp=drive_link",
  },
  {
    title: "Khidmat bantu dron",
    viewUrl:
      "https://docs.google.com/document/d/1euIyXy0yb2wqE3_QwLRwAlOTnuVdAz2Vbo1HbK7Bw9A/edit?tab=t.0",
    kind: "gdoc",
  },
  {
    title: "Pencetak 3D",
    viewUrl:
      "https://drive.google.com/file/d/11w7hSfjDNVr2v8NtRbXYfTc-QNiJ6PTK/view?usp=drive_link",
  },
  {
    title: "Reka Bentuk 3D",
    viewUrl:
      "https://drive.google.com/file/d/1sXYW-8zGpgqJz4-Y7eGTGgPHtKljC-fy/view?usp=drive_link",
  },
  {
    title: "OPR Bengkel MYRC",
    viewUrl:
      "https://drive.google.com/file/d/1fxIy4U61VSvmDBkLuUhMf57pJfPA8RVQ/view?usp=drive_link",
  },
  {
    title: "OPR Bengkel Laman Web",
    viewUrl:
      "https://drive.google.com/file/d/18HWD-H5jS_o8klAr6ospW7-pottj3QJI/view?usp=drive_link",
  },
  {
    title:
      "OPR Bengkel Penggunaan Mikrobotik & Edu Kit (COE USTP & JU RBT Daerah)",
    viewUrl:
      "https://drive.google.com/file/d/1HLCjQTSwIJMD2-q6nrRRPhfwFFn-AWsk/view?usp=drive_link",
  },
];

/** Integrasi — JNJ: laporan PDF tambahan (kolaborasi / Ana Muslim, dll.) */
const INTEGRASI_JNJ_DRIVE_PDF_CARDS = [
  {
    title: "JNJ Manjung & Padang Terap",
    blurb: "Bersama USTP PPD Padang Terap · SK Bandar Baru Jitra — PDF (Google Drive)",
    viewUrl:
      "https://drive.google.com/file/d/1Kig1xzkFNG-8HVGVrTQoxuhJdYPu-3Zy/view?usp=drive_link",
  },
  {
    title: "JNJ Ana Muslim Seri Selamat",
    blurb: "PDF (Google Drive)",
    viewUrl:
      "https://drive.google.com/file/d/14S77Zsupr7CayN4rPzSyHAklE6LEwd8i/view?usp=drive_link",
  },
  {
    title: "JNJ Ana Muslim SG Tiram",
    blurb: "PDF (Google Drive)",
    viewUrl:
      "https://drive.google.com/file/d/11fojvC7hJDlNSlPC1isFmStYlR2ZlOjw/view?usp=drive_link",
  },
];

/** Integrasi — JNJ: JNJ YTY (Canva Grafik per fasa, PDF Google Drive) */
const INTEGRASI_JNJ_YTY_PDF_CARDS = [
  {
    viewUrl:
      "https://drive.google.com/file/d/1cHQpv3ddBx-RLfvOCDn71_VGaLy15aDW/view?usp=drive_link",
    blurb: "Canva Grafik Fasa 1 — PDF (Google Drive)",
  },
  {
    viewUrl:
      "https://drive.google.com/file/d/1vnq-bV0QSfGzdCyOxB2ZbahycHeV5mKK/view?usp=drive_link",
    blurb: "Canva Grafik Fasa 2 — PDF (Google Drive)",
  },
];

/** Integrasi — Impak JNJ (imej ringkasan dalam public/) */
const INTEGRASI_IMPAK_JNJ_MYRC_IMG = `${import.meta.env.BASE_URL}itm-impak-jnj-myrc.jpg`;
const INTEGRASI_IMPAK_JNJ_MINECRAFT_IMG = `${import.meta.env.BASE_URL}itm-impak-jnj-minecraft.jpg`;
const INTEGRASI_IMPAK_JNJ_PANDAI_APP_IMG = `${import.meta.env.BASE_URL}itm-impak-jnj-pandai-app.jpg`;
const INTEGRASI_IMPAK_JNJ_REKA_EDUKIT_IMG = `${import.meta.env.BASE_URL}itm-impak-jnj-reka-edukit.jpg`;
/** Impak JNJ — Mikrobotik & Reka Edukit (PDF, Google Drive) */
const INTEGRASI_IMPAK_JNJ_MIKROBOTIK_REKA_EDUKIT_PDF_URL =
  "https://drive.google.com/file/d/15kvLB4q5Z8-zw-YnPz5wbS9qqMqZOUr5/view?usp=drive_link";

/** Hebahan Pendidikan Digital — COE kepada komuniti (imej dalam public/) */
const HEBAHAN_HARI_TERBUKA_PPD_MANJUNG_IMG = `${import.meta.env.BASE_URL}hebahan-hari-terbuka-ppd-manjung.png`;
const HEBAHAN_HARI_TERBUKA_PAMERAN_USTP_IMG = `${import.meta.env.BASE_URL}hebahan-hari-terbuka-pameran-ustp.png`;
const HEBAHAN_KARNIVAL_DIGITAL_MANJUNG_IMG = `${import.meta.env.BASE_URL}hebahan-karnival-digital-ppd-manjung.png`;
const HEBAHAN_KARNIVAL_DIGITAL_PANGKOR_IMG = `${import.meta.env.BASE_URL}hebahan-karnival-digital-pulau-pangkor.png`;
const HEBAHAN_PAMERAN_COE_NEGERI_IMG = `${import.meta.env.BASE_URL}hebahan-pameran-coe-negeri-perak.png`;

/** Inisiatif Teknologi Maklumat — Google Classroom & JNJ (Docs) */
const GOOGLE_CLASSROOM_USTP_MANJUNG_URL = "https://classroom.google.com/u/0/h";
const ITM_GOOGLE_CLASSROOM_PREVIEW_IMAGE = `${import.meta.env.BASE_URL}itm-card-google-classroom-ustp-manjung.png`;
const ITM_JNJ_MAXIS_DOC_URL =
  "https://docs.google.com/document/d/1WWFVYr3diOegC380eH8_-eN953vb66Jzc4a3lFHp9GI/edit?tab=t.0";
const ITM_JNJ_IAB_DOC_URL =
  "https://docs.google.com/document/d/1qIA5pEZcCfWgFo0aJcOrr5hzoFvkskSox_-kuZ6N2tM/edit?tab=t.0";
/** ITM — JNJ NADI (PDF, Google Drive) */
const ITM_JNJ_NADI_PDF_URL =
  "https://drive.google.com/file/d/1OIh9uBmlUZN5AvfVV25HcUjsSAfEX-Y3/view?usp=drive_link";

/** Pembudayaan Amalan Membaca — JNJ Blink Book (RAK MAYA) */
const PEMBUDAYAAN_JNJ_BLINK_BOOK_PSS_MAYA_URL = "https://sesdamaya.rakmaya.com/";
const PEMBUDAYAAN_JNJ_BLINK_BOOK_SK_SERI_SAMUDERA_URL =
  "https://skeses.rakmaya.com/bio/";
const PEMBUDAYAAN_JNJ_BLINK_BOOK_PSS_MAYA_PREVIEW_IMAGE = `${import.meta.env.BASE_URL}pembudayaan-jnj-blink-book-pss-maya-sesda.png`;
const PEMBUDAYAAN_JNJ_BLINK_BOOK_SK_SERI_SAMUDERA_PREVIEW_IMAGE = `${import.meta.env.BASE_URL}pembudayaan-jnj-blink-book-sk-seri-samudera.png`;

/** Pembudayaan Amalan Membaca — Program Inovasi */
const PEMBUDAYAAN_PROGRAM_INOVASI_VIDEO_DRIVE_URL =
  "https://drive.google.com/file/d/1xAuQEgy4nokBvt1ezd6J_nvxXpdl7b-P/view?usp=drive_link";
const PEMBUDAYAAN_PROGRAM_INOVASI_BAHAN_ARTSTEPS_URL =
  "https://www.artsteps.com/view/68dbfbda95f4a3708b75bd04";
const PEMBUDAYAAN_PROGRAM_INOVASI_ARTSTEPS_PREVIEW_IMAGE = `${import.meta.env.BASE_URL}pembudayaan-program-inovasi-artsteps.png`;

/** Status DCS — peratus (laras bersama teks kad & carta) */
const DCS_TOV_2024_DAERAH = 65;
const DCS_KPI_2025_KEBANGSAAN = 78;
const DCS_CAPAI_2025_DAERAH = 78.24;

const pegawaiData = [
  {
    nama: "Syed Muhammad Rujhan Syed Alwi",
    jawatan: "Ketua Unit (Penolong PPD Sumber & Teknologi Pendidikan PPD Manjung)",
    telefon: "019-9669812 / 05-692 7814",
    // Masukkan salah satu:
    // 1) detailUrl: link (boleh iframe) ke halaman Canva/Looker/etc
    // 2) detailImage: path imej (/pegawai/...) atau pautan Google Drive (view) — pratontoh dalam modal
    detailUrl: "",
    detailImage: "/pegawai/SydMuhammadRujhan.png",
  },
  {
    nama: "Ong Chong Xiao",
    jawatan: "Penolong PPD Sumber & Teknologi Pendidikan PPD Manjung",
    telefon: "017-550 3909 / 05-692 7814",
    detailUrl: "",
    detailImage: "/pegawai/OngChongXiao.png",
  },
  {
    nama: "Parimilah a/p Vadiveloo",
    jawatan: "Penolong PPD Sumber & Teknologi Pendidikan PPD Manjung",
    telefon: "016-5401338 / 05-692 7814",
    detailUrl: "",
    detailImage: "/pegawai/Parimilah.png",
  },
  {
    nama: "Zaliyatunnur binti Awang Kechik",
    jawatan: "Penolong PPD Sumber & Teknologi Pendidikan PPD Manjung",
    telefon: "019-5742887 / 05-692 7814",
    detailUrl: "",
    detailImage: "/pegawai/Zaliyatunnur.png",
  },
  {
    nama: "Mohd Hairul Anuar bin Hashim",
    jawatan: "Penolong PPD Sumber & Teknologi Pendidikan PPD Manjung",
    telefon: "010-6525569 / 05-692 7814",
    detailUrl: "",
    detailImage: "/pegawai/MohdHairulAnuar.png",
  },
];

/** Infografik: /assets/pensijilan-digital.png — data utama untuk kad ringkasan */
const PENSIJILAN_DIGITAL_IMAGE = "/assets/pensijilan-digital.png";

const PENSIJILAN_BY_SCHOOL = [
  ["SK", 746],
  ["SMK", 436],
  ["SJK(C)", 116],
  ["SJK(T)", 195],
  ["SRK", 127],
  ["SM Agama (SABK)", 21],
];

const PENSIJILAN_BY_LOCATION = [
  ["Bandar", 1109],
  ["Luar bandar", 674],
];

/** Ringkasan statistik pada kad (selari data dalam infografik) */
function PensijilanDigitalSummary({ data } = {}) {
  const locPairs = data?.locations ?? PENSIJILAN_BY_LOCATION;
  const schoolPairs = data?.schools ?? PENSIJILAN_BY_SCHOOL;
  const locSum = locPairs.reduce((a, [, n]) => a + n, 0);
  return (
    <div className="mt-4 space-y-3 rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3 text-xs text-slate-300">
      <div>
        <p className="mb-1.5 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          Jumlah mengikut lokasi
        </p>
        <div className="flex flex-wrap gap-2">
          {locPairs.map(([label, n]) => (
            <span
              key={label}
              className="rounded-lg border border-cyan-400/20 bg-slate-950/60 px-2 py-1 tabular-nums text-slate-200"
            >
              <span className="text-slate-400">{label}:</span>{" "}
              <span className="font-semibold text-cyan-200">{n.toLocaleString("en-MY")}</span>
            </span>
          ))}
          <span className="rounded-lg border border-cyan-400/10 px-2 py-1 text-[10px] text-slate-500">
            Jumlah: {locSum.toLocaleString("en-MY")}
          </span>
        </div>
      </div>

      <div>
        <p className="mb-1.5 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          Jenis sekolah (terpilih)
        </p>
        <div className="max-h-[140px] overflow-y-auto rounded-lg border border-cyan-400/10 bg-slate-950/30">
          <table className="w-full text-left text-[10px]">
            <tbody>
              {schoolPairs.map(([label, n]) => (
                <tr key={label} className="border-t border-cyan-400/10">
                  <td className="px-2 py-1 text-slate-400">{label}</td>
                  <td className="px-2 py-1 text-right tabular-nums font-medium text-slate-100">
                    {n.toLocaleString("en-MY")}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div>
        <p className="mb-1.5 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          Pensijilan penyedia
        </p>
        <div className="flex flex-wrap gap-1.5">
          <span className="rounded-full border border-cyan-400/25 bg-cyan-500/10 px-2 py-0.5 text-[10px] text-cyan-100">
            Google GCE1: 996
          </span>
          <span className="rounded-full border border-cyan-400/25 bg-cyan-500/10 px-2 py-0.5 text-[10px] text-cyan-100">
            Gemini: 730
          </span>
          <span className="rounded-full border border-violet-400/25 bg-violet-500/10 px-2 py-0.5 text-[10px] text-violet-100">
            GCE2: 77
          </span>
          <span className="rounded-full border border-slate-500/30 px-2 py-0.5 text-[10px] text-slate-300">
            GCT: 9
          </span>
          <span className="rounded-full border border-amber-400/25 bg-amber-500/10 px-2 py-0.5 text-[10px] text-amber-100">
            Apple: 11
          </span>
          <span className="rounded-full border border-sky-400/25 bg-sky-500/10 px-2 py-0.5 text-[10px] text-sky-100">
            MIE / MIEE: 1 / 1
          </span>
        </div>
      </div>
    </div>
  );
}


/** Garisan KPI kebangsaan & snapshot OPTIK 2 / AI Tools dalam DELIMa (data slaid) */
const AI_TOOLS_KPI_KEBANGSAAN = 55;
const AI_TOOLS_SNAPSHOT_OPTIK = {
  asAt: "27 Nov 2025",
  selesaiPct: 86.43,
  selesaiBil: 3179,
  belumPct: 13.57,
  belumBil: 499,
};
const AI_TOOLS_PERGERAKAN = {
  tov2024: 63.43,
  ar1Julai: 45.29,
  ar2Okt: 60.36,
};

/** Carta kawasan OPTIK — gaya DCS: lorek cyan + isi kecerunan + garisan sasaran 55%. */
function AiToolsOptikAreaChart({ snapshot, pergerakan, kpiPct }) {
  const W = 400;
  const H = 214;
  const pad = { l: 42, r: 24, t: 20, b: 62 };
  const innerW = W - pad.l - pad.r;
  const innerH = H - pad.t - pad.b;
  const yMin = 40;
  const yMax = 92;
  const kpi = kpiPct;

  const pts = [
    { pct: pergerakan.tov2024, line1: "TOV 2024", line2: `${pergerakan.tov2024.toFixed(2)}%` },
    { pct: pergerakan.ar1Julai, line1: "AR1 (Julai)", line2: `${pergerakan.ar1Julai.toFixed(2)}%` },
    { pct: pergerakan.ar2Okt, line1: "AR2 (Okt)", line2: `${pergerakan.ar2Okt.toFixed(2)}%` },
    {
      pct: snapshot.selesaiPct,
      line1: "Selesai",
      line2: `${snapshot.selesaiPct.toFixed(2)}%`,
      line3: snapshot.asAt,
    },
  ];
  const n = pts.length;

  const yAt = (pct) =>
    pad.t + (1 - (pct - yMin) / (yMax - yMin)) * innerH;
  const yBase = pad.t + innerH;
  const xAt = (i) => pad.l + (n === 1 ? innerW / 2 : (i / (n - 1)) * innerW);

  const xs = pts.map((_, i) => xAt(i));
  const ys = pts.map((p) => yAt(p.pct));
  const yKpi = yAt(kpi);

  const lineD = xs.map((x, i) => `${i === 0 ? "M" : "L"} ${x} ${ys[i]}`).join(" ");
  const areaD = `M ${xs[0]} ${yBase} ${xs.map((x, i) => `L ${x} ${ys[i]}`).join(" ")} L ${xs[n - 1]} ${yBase} Z`;

  const yTicks = [40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90].filter(
    (v) => v >= yMin && v <= yMax
  );
  const x1 = pad.l + innerW;

  return (
    <svg
      viewBox={`0 0 ${W} ${H}`}
      className="h-auto w-full max-w-full text-slate-500"
      role="img"
      aria-label={`Trend OPTIK daripada TOV ${pergerakan.tov2024}% ke selesai ${snapshot.selesaiPct}% pada ${snapshot.asAt}, sasaran ${kpi}%`}
    >
      <defs>
        <linearGradient id="aiToolsOptikAreaFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#22d3ee" stopOpacity="0.35" />
          <stop offset="100%" stopColor="#22d3ee" stopOpacity="0" />
        </linearGradient>
      </defs>

      {yTicks.map((v) => {
        const y = yAt(v);
        return (
          <g key={v}>
            <line
              x1={pad.l}
              y1={y}
              x2={x1}
              y2={y}
              stroke="currentColor"
              strokeOpacity={0.1}
              strokeWidth={1}
            />
            <text
              x={pad.l - 6}
              y={y + 3}
              textAnchor="end"
              className="fill-slate-500"
              style={{ fontSize: 9 }}
            >
              {v}%
            </text>
          </g>
        );
      })}

      <line
        x1={pad.l}
        y1={yKpi}
        x2={x1}
        y2={yKpi}
        stroke="#fbbf24"
        strokeWidth={1.75}
        strokeDasharray="7 5"
        strokeOpacity={0.95}
      />
      <text
        x={x1}
        y={yKpi - 5}
        textAnchor="end"
        className="fill-amber-200"
        style={{ fontSize: 9 }}
      >
        Sasaran KPI kebangsaan ({kpi}%)
      </text>

      <path d={areaD} fill="url(#aiToolsOptikAreaFill)" />
      <path
        d={lineD}
        fill="none"
        stroke="#22d3ee"
        strokeWidth={2.5}
        strokeLinecap="round"
        strokeLinejoin="round"
      />

      {xs.map((x, i) => {
        const y = ys[i];
        const last = i === n - 1;
        const first = i === 0;
        if (first) {
          return (
            <circle
              key={i}
              cx={x}
              cy={y}
              r={4.5}
              fill="#0f172a"
              stroke="#22d3ee"
              strokeWidth={2}
            />
          );
        }
        if (last) {
          return (
            <circle
              key={i}
              cx={x}
              cy={y}
              r={6}
              fill="#34d399"
              stroke="#6ee7b7"
              strokeWidth={2}
            />
          );
        }
        return (
          <circle
            key={i}
            cx={x}
            cy={y}
            r={4}
            fill="#0f172a"
            stroke={pts[i].pct >= kpi ? "#22d3ee" : "#fbbf24"}
            strokeWidth={1.75}
          />
        );
      })}

      {pts.map((p, i) => {
        const x = xs[i];
        const bottomY = H - 8;
        return (
          <g key={p.line1}>
            <text
              x={x}
              y={bottomY - 22}
              textAnchor="middle"
              className="fill-slate-400"
              style={{ fontSize: 8.5 }}
            >
              {p.line1}
            </text>
            <text
              x={x}
              y={bottomY - 10}
              textAnchor="middle"
              className={
                i === n - 1
                  ? "fill-emerald-300/95"
                  : "fill-cyan-200/90"
              }
              style={{ fontSize: 9, fontWeight: i === n - 1 ? 600 : 500 }}
            >
              {p.line2}
            </text>
            {p.line3 ? (
              <text
                x={x}
                y={bottomY + 2}
                textAnchor="middle"
                className="fill-slate-500"
                style={{ fontSize: 7.5 }}
              >
                {p.line3}
              </text>
            ) : null}
          </g>
        );
      })}
    </svg>
  );
}

function AiToolsDelimaSummary({ sheet } = {}) {
  const k = sheet
    ? {
        asAt: sheet.asAt,
        selesaiPct: sheet.selesaiPct,
        selesaiBil: sheet.selesaiBil,
        belumPct: sheet.belumPct,
        belumBil: sheet.belumBil,
      }
    : AI_TOOLS_SNAPSHOT_OPTIK;
  const p = sheet
    ? {
        tov2024: sheet.tov2024,
        ar1Julai: sheet.ar1Julai,
        ar2Okt: sheet.ar2Okt,
      }
    : AI_TOOLS_PERGERAKAN;
  const line = sheet ? sheet.kpiKebangsaan : AI_TOOLS_KPI_KEBANGSAAN;

  const footerText =
    sheet?.footerNote ||
    `AR2 (Okt) melebihi sasaran ${line}% berbanding AR1; titik akhir ialah peratus selesai snapshot ${k.asAt} (bukan skala sama dengan TOV/AR).`;

  return (
    <div className="mt-3 flex flex-1 flex-col space-y-4 text-xs text-slate-300">
      <div className="flex flex-wrap items-center gap-3 rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3">
        <span
          className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg border border-amber-400/35 bg-amber-500/10 text-amber-200"
          title="Sasaran kebangsaan"
        >
          <svg className="h-5 w-5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.75} aria-hidden>
            <circle cx="12" cy="12" r="10" />
            <circle cx="12" cy="12" r="3" />
          </svg>
        </span>
        <div className="min-w-0 flex-1">
          <p className="text-[10px] font-semibold uppercase tracking-wide text-slate-500">
            Garisan KPI kebangsaan
          </p>
          <p className="text-lg font-bold tabular-nums text-amber-100">{line.toFixed(2)}%</p>
        </div>
        <span className="inline-flex items-center gap-1.5 rounded-full border border-emerald-400/35 bg-emerald-500/12 px-2.5 py-1 text-[11px] font-medium text-emerald-200">
          <svg className="h-3.5 w-3.5 shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5} aria-hidden>
            <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
          </svg>
          Melebihi garisan (snapshot daerah)
        </span>
      </div>

      <div className="rounded-xl border border-cyan-400/15 bg-slate-950/35 p-3">
        <p className="mb-2 flex flex-wrap items-center gap-2 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          <svg className="h-3.5 w-3.5 text-cyan-300" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden>
            <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
          </svg>
          Aktif / selesai — setakat {k.asAt}
        </p>
        <svg viewBox="0 0 100 26" className="h-14 w-full" preserveAspectRatio="none" aria-hidden>
          <rect x="0" y="6" width={k.selesaiPct} height="14" rx="1" fill="rgba(52,211,153,0.55)" />
          <rect x={k.selesaiPct} y="6" width={k.belumPct} height="14" rx="1" fill="rgba(251,113,133,0.45)" />
          <line
            x1={line}
            y1="2"
            x2={line}
            y2="24"
            stroke="#fcd34d"
            strokeWidth="0.9"
            strokeDasharray="2.2 2.2"
            vectorEffect="non-scaling-stroke"
          />
        </svg>
        <div className="mt-2 flex flex-wrap gap-2">
          <span className="inline-flex items-center gap-1 rounded-lg border border-emerald-400/25 bg-emerald-500/10 px-2 py-1 tabular-nums text-emerald-100">
            <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden>
              <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
            </svg>
            Selesai {k.selesaiPct}% ({k.selesaiBil.toLocaleString("en-MY")})
          </span>
          <span className="inline-flex items-center gap-1 rounded-lg border border-rose-400/25 bg-rose-500/10 px-2 py-1 tabular-nums text-rose-100">
            <svg className="h-3.5 w-3.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2} aria-hidden>
              <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
            </svg>
            Belum {k.belumPct}% ({k.belumBil.toLocaleString("en-MY")})
          </span>
        </div>
      </div>

      <div className="rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3">
        <p className="mb-2 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          Pergerakan & snapshot (TOV → AR → selesai)
        </p>
        <AiToolsOptikAreaChart snapshot={k} pergerakan={p} kpiPct={line} />
        <div className="mt-2 flex flex-wrap items-center justify-center gap-x-4 gap-y-1 text-[10px] text-slate-500">
          <span className="inline-flex items-center gap-1.5">
            <span
              className="inline-block h-0 w-8 border-t-2 border-dashed border-amber-300/90"
              aria-hidden
            />
            Garisan sasaran kebangsaan
          </span>
          <span className="inline-flex items-center gap-1.5">
            <span
              className="inline-block h-0.5 w-8 rounded-full bg-cyan-400"
              aria-hidden
            />
            Trend capaian
          </span>
        </div>
        <p className="mt-2 text-[10px] leading-relaxed text-slate-500">{footerText}</p>
      </div>
    </div>
  );
}


function StatusPill({ label, tone }) {
  const map = {
    good: "bg-emerald-500/15 text-emerald-300 border-emerald-400/30",
    warn: "bg-amber-500/15 text-amber-300 border-amber-400/30",
    bad: "bg-rose-500/15 text-rose-300 border-rose-400/30",
  };
  return (
    <span className={`rounded-full border px-3 py-1 text-xs ${map[tone]}`}>
      {label}
    </span>
  );
}

/** Carta garis: trend daerah TOV → capai 2025 + garisan sasaran KPI kebangsaan. */
function DcsKpiLineChart({
  tov = DCS_TOV_2024_DAERAH,
  kpi = DCS_KPI_2025_KEBANGSAAN,
  capai = DCS_CAPAI_2025_DAERAH,
  yMin = 58,
  yMax = 82,
} = {}) {
  const W = 400;
  const H = 168;
  const pad = { l: 40, r: 28, t: 18, b: 44 };
  const innerW = W - pad.l - pad.r;
  const innerH = H - pad.t - pad.b;
  const yAt = (pct) =>
    pad.t + (1 - (pct - yMin) / (yMax - yMin)) * innerH;

  const x0 = pad.l;
  const x1 = pad.l + innerW;
  const xCapaiLabel = x1 - 22;
  const y0 = yAt(tov);
  const y1 = yAt(capai);
  const yKpi = yAt(kpi);

  const yTicks = [60, 65, 70, 75, 80].filter((v) => v >= yMin && v <= yMax);

  return (
    <svg
      viewBox={`0 0 ${W} ${H}`}
      className="h-auto w-full max-w-full text-slate-500"
      role="img"
      aria-label={`Trend capaian DCS daerah daripada ${tov}% ke ${capai}%, sasaran kebangsaan ${kpi}%`}
    >
      <defs>
        <linearGradient id="dcsAreaFill" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor="#22d3ee" stopOpacity="0.35" />
          <stop offset="100%" stopColor="#22d3ee" stopOpacity="0" />
        </linearGradient>
      </defs>

      {yTicks.map((v) => {
        const y = yAt(v);
        return (
          <g key={v}>
            <line
              x1={pad.l}
              y1={y}
              x2={x1}
              y2={y}
              stroke="currentColor"
              strokeOpacity={0.1}
              strokeWidth={1}
            />
            <text
              x={pad.l - 6}
              y={y + 3}
              textAnchor="end"
              className="fill-slate-500"
              style={{ fontSize: 9 }}
            >
              {v}%
            </text>
          </g>
        );
      })}

      <line
        x1={x0}
        y1={yKpi}
        x2={x1}
        y2={yKpi}
        stroke="#fbbf24"
        strokeWidth={1.75}
        strokeDasharray="7 5"
        strokeOpacity={0.95}
      />
      <text
        x={x1}
        y={yKpi - 5}
        textAnchor="end"
        className="fill-amber-200"
        style={{ fontSize: 9 }}
      >
        Sasaran KPI kebangsaan ({kpi}%)
      </text>

      <path
        d={`M ${x0} ${pad.t + innerH} L ${x0} ${y0} L ${x1} ${y1} L ${x1} ${pad.t + innerH} Z`}
        fill="url(#dcsAreaFill)"
      />
      <path
        d={`M ${x0} ${y0} L ${x1} ${y1}`}
        fill="none"
        stroke="#22d3ee"
        strokeWidth={2.75}
        strokeLinecap="round"
        strokeLinejoin="round"
      />

      <circle cx={x0} cy={y0} r={4.5} fill="#0f172a" stroke="#22d3ee" strokeWidth={2} />
      <circle cx={x1} cy={y1} r={6} fill="#34d399" stroke="#6ee7b7" strokeWidth={2} />

      <text
        x={x0}
        y={H - 10}
        textAnchor="middle"
        className="fill-slate-400"
        style={{ fontSize: 10 }}
      >
        TOV 2024
      </text>
      <text
        x={x0}
        y={H - 22}
        textAnchor="middle"
        className="fill-cyan-200/90"
        style={{ fontSize: 9 }}
      >
        {tov}%
      </text>
      <text
        x={xCapaiLabel}
        y={H - 10}
        textAnchor="middle"
        className="fill-slate-400"
        style={{ fontSize: 10 }}
      >
        Capai 2025
      </text>
      <text
        x={xCapaiLabel}
        y={H - 22}
        textAnchor="middle"
        className="fill-emerald-300/95"
        style={{ fontSize: 9 }}
      >
        {capai}%
      </text>
    </svg>
  );
}


/** Carta ringkas Program Ains: nisbah Approved/Rejected + approved mengikut jenis sekolah. */
function ProgramAinsCharts({ stats, loading, error }) {
  if (loading) {
    return (
      <p className="py-5 text-center text-sm text-slate-400">Memuatkan carta…</p>
    );
  }
  if (error) {
    return (
      <p className="py-4 text-center text-sm text-rose-200/90">Carta tidak tersedia</p>
    );
  }

  const {
    approved,
    rejected,
    skApproved,
    sjkcApproved,
    sjktApproved,
  } = stats;

  const W = 340;
  const H = 212;
  const padX = 10;
  const barFullW = W - padX * 2;
  const yBar = 34;
  const hBar = 20;
  const sumAR = approved + rejected;
  const hasAR = sumAR > 0;
  let wAp = hasAR ? (approved / sumAR) * barFullW : 0;
  let wRe = hasAR ? (rejected / sumAR) * barFullW : 0;
  if (hasAR && approved > 0 && wAp < 6) wAp = 6;
  if (hasAR && rejected > 0 && wRe < 6) wRe = 6;
  if (hasAR && wAp + wRe > barFullW) {
    const s = wAp + wRe;
    wAp = (wAp / s) * barFullW;
    wRe = (wRe / s) * barFullW;
  }

  const types = [
    { key: "SK", v: skApproved, fill: "#22d3ee", x: 58 },
    { key: "SJKC", v: sjkcApproved, fill: "#a78bfa", x: 170 },
    { key: "SJKT", v: sjktApproved, fill: "#fbbf24", x: 282 },
  ];
  const maxT = Math.max(skApproved, sjkcApproved, sjktApproved, 1);
  /** Jarak antara label Approved/Rejected dengan carta palang — elak palang SK naik ke atas teks. */
  const yCountBaseline = yBar + hBar + 20;
  const ySec2Title = yCountBaseline + 26;
  const barPlotH = 48;
  /** plotBaseY = dasar palang; puncak palang = plotBaseY − bh. Mesti > teks tajuk (baseline + turun huruf). */
  const gapBelowSec2Title = 24;
  const plotBaseY = ySec2Title + gapBelowSec2Title + barPlotH;

  return (
    <svg
      viewBox={`0 0 ${W} ${H}`}
      className="h-auto w-full max-w-full text-slate-500"
      role="img"
      aria-label="Carta Program Ains: approved berbanding rejected dan approved mengikut jenis sekolah"
    >
      <text
        x={padX}
        y={16}
        className="fill-slate-400"
        style={{ fontSize: 10 }}
      >
        Approved berbanding Rejected (kumulatif)
      </text>
      <rect
        x={padX}
        y={yBar}
        width={barFullW}
        height={hBar}
        rx={4}
        className="fill-slate-800/80"
      />
      {hasAR ? (
        <>
          <rect
            x={padX}
            y={yBar}
            width={wAp}
            height={hBar}
            rx={wRe > 0 ? 0 : 4}
            className="fill-cyan-500/75"
          />
          {wRe > 0 ? (
            <rect
              x={padX + wAp}
              y={yBar}
              width={wRe}
              height={hBar}
              rx={4}
              className="fill-rose-500/65"
            />
          ) : null}
        </>
      ) : (
        <text
          x={W / 2}
          y={yBar + 14}
          textAnchor="middle"
          className="fill-slate-500"
          style={{ fontSize: 9 }}
        >
          Tiada approved / rejected
        </text>
      )}
      <text
        x={padX}
        y={yCountBaseline}
        className="fill-cyan-200/90"
        style={{ fontSize: 9 }}
      >
        Approved: {approved.toLocaleString("en-MY")}
      </text>
      <text
        x={padX + barFullW}
        y={yCountBaseline}
        textAnchor="end"
        className="fill-rose-200/85"
        style={{ fontSize: 9 }}
      >
        Rejected: {rejected.toLocaleString("en-MY")}
      </text>

      <text
        x={padX}
        y={ySec2Title}
        className="fill-slate-400"
        style={{ fontSize: 10 }}
      >
        Approved mengikut jenis sekolah
      </text>
      {types.map(({ key, v, fill, x }) => {
        const bh = Math.max(4, (v / maxT) * barPlotH);
        const bx = x - 18;
        const by = plotBaseY - bh;
        return (
          <g key={key}>
            <rect
              x={bx}
              y={by}
              width={36}
              height={bh}
              rx={3}
              fill={fill}
              fillOpacity={0.82}
            />
            <text
              x={x}
              y={plotBaseY + 14}
              textAnchor="middle"
              className="fill-slate-400"
              style={{ fontSize: 9 }}
            >
              {key}
            </text>
            <text
              x={x}
              y={plotBaseY + 26}
              textAnchor="middle"
              className="fill-slate-200"
              style={{ fontSize: 9 }}
            >
              {v.toLocaleString("en-MY")}
            </text>
          </g>
        );
      })}
    </svg>
  );
}

export default function App() {
  const [keyword, setKeyword] = useState("");
  const [selectedPegawai, setSelectedPegawai] = useState(null);
  const [isDcsImageOpen, setIsDcsImageOpen] = useState(false);
  const [isPensijilanDigitalImageOpen, setIsPensijilanDigitalImageOpen] =
    useState(false);

  const [delimaAn, setDelimaAn] = useState(null);
  const [delimaAnLoading, setDelimaAnLoading] = useState(true);
  const [delimaAnError, setDelimaAnError] = useState("");

  const [dcsAn, setDcsAn] = useState(null);
  const [_dcsAnLoading, setDcsAnLoading] = useState(true);
  const [dcsAnError, setDcsAnError] = useState("");

  const [ainsAn, setAinsAn] = useState(null);
  const [ainsLoading, setAinsLoading] = useState(true);
  const [ainsError, setAinsError] = useState("");

  const [pensijilanAn, setPensijilanAn] = useState(null);
  const [_pensijilanAnLoading, setPensijilanAnLoading] = useState(true);
  const [pensijilanAnError, setPensijilanAnError] = useState("");

  const [optikAn, setOptikAn] = useState(null);
  const [optikLoading, setOptikLoading] = useState(true);
  const [optikError, setOptikError] = useState("");

  const [lamanWebRows, setLamanWebRows] = useState([]);
  const [lamanWebLoading, setLamanWebLoading] = useState(true);
  const [lamanWebError, setLamanWebError] = useState("");
  const [lamanWebSearch, setLamanWebSearch] = useState("");
  /** Baris tambahan / kemas kini URL daripada tab Google Sheet (`VITE_OSC_GID_LAMAN_WEB_SEKOLAH`). */
  const [lamanWebSheetRows, setLamanWebSheetRows] = useState([]);

  // Design mode:
  // - "dark": current version (深色背景)
  // - "geminiA": 浅色背景图 + 深色卡片（方案A）
  // - "neonDark": 深色背景 + 霓虹网格（第三个选项）
  const [designMode, setDesignMode] = useState("dark");

  const [maklumatAsasConfig, setMaklumatAsasConfig] = useState({});
  /** null = guna `pegawaiData`; array = dari Sheet (boleh kosong) */
  const [maklumatAsasPegawaiOverride, setMaklumatAsasPegawaiOverride] =
    useState(null);
  const [maklumatAsasLoading, setMaklumatAsasLoading] = useState(true);
  const [maklumatAsasError, setMaklumatAsasError] = useState("");

  const maklumatAsasDisplay = useMemo(() => {
    const c = maklumatAsasConfig;
    const val = (key, fallback) => {
      const v = c[key];
      const t = v != null ? String(v).trim() : "";
      return t || fallback;
    };
    const cartaImage = val("carta_image_url", CARTA_ORGANISASI_IMG);
    const pkgImage = val("pkg_image_url", MAKLUMAT_PKG_COE_IMG);
    const takwimEmbed = val("takwim_embed_url", USTP_CALENDAR_EMBED);
    return {
      cartaImage,
      cartaPreview: maklumatAsasPreviewMode(cartaImage),
      cartaFull: val("carta_full_url", cartaImage),
      cartaTitle: val("carta_title", "Carta Organisasi"),
      cartaBlurb: val(
        "carta_blurb",
        "Organisasi USTP PPD Manjung — hierarki PPD, PKG dan COE.",
      ),
      pkgImage,
      pkgPreview: maklumatAsasPreviewMode(pkgImage),
      pkgFull: val("pkg_full_url", pkgImage),
      pkgTitle: val("pkg_title", "Maklumat PKG"),
      pkgBlurb: val(
        "pkg_blurb",
        "Maklumat COE Daerah Manjung (AQA1001–AQA1005) — kemudahan & hubungi.",
      ),
      takwimEmbed,
      takwimFull: val("takwim_full_url", takwimEmbed),
      takwimTitle: val("takwim_title", "Takwim"),
    };
  }, [maklumatAsasConfig]);

  const filteredPegawai = useMemo(() => {
    const pegawaiList =
      maklumatAsasPegawaiOverride !== null
        ? maklumatAsasPegawaiOverride
        : pegawaiData;
    const k = keyword.toLowerCase().trim();
    if (!k) return pegawaiList;
    return pegawaiList.filter(
      (p) =>
        p.nama.toLowerCase().includes(k) ||
        p.jawatan.toLowerCase().includes(k) ||
        p.telefon.includes(k),
    );
  }, [keyword, maklumatAsasPegawaiOverride]);

  useEffect(() => {
    let cancelled = false;
    const gid = (k) => String(import.meta.env[k] ?? "").trim();

    const loadOne = async ({
      demoPath,
      gidKey,
      parse,
      setData,
      setLoading,
      setError,
    }) => {
      try {
        setLoading(true);
        setError("");
        const text = await fetchAnalisisCsvText({
          gid: gid(gidKey),
          demoPath,
        });
        if (cancelled) return;
        setData(parse(text));
      } catch (e) {
        if (cancelled) return;
        setError(
          e instanceof Error ? e.message : "Gagal memuatkan data analisis.",
        );
        setData(null);
      } finally {
        if (!cancelled) setLoading(false);
      }
    };

    void Promise.all([
      loadOne({
        demoPath: "data/analisis-delima-demo.csv",
        gidKey: "VITE_ANALISIS_GID_DELIMA",
        parse: parseDelimaAnalisisCsv,
        setData: setDelimaAn,
        setLoading: setDelimaAnLoading,
        setError: setDelimaAnError,
      }),
      loadOne({
        demoPath: "data/analisis-dcs-demo.csv",
        gidKey: "VITE_ANALISIS_GID_DCS",
        parse: parseDcsAnalisisCsv,
        setData: setDcsAn,
        setLoading: setDcsAnLoading,
        setError: setDcsAnError,
      }),
      loadOne({
        demoPath: "data/analisis-ains-demo.csv",
        gidKey: "VITE_ANALISIS_GID_AINS",
        parse: parseAinsAnalisisCsv,
        setData: setAinsAn,
        setLoading: setAinsLoading,
        setError: setAinsError,
      }),
      loadOne({
        demoPath: "data/analisis-pensijilan-demo.csv",
        gidKey: "VITE_ANALISIS_GID_PENSIJILAN",
        parse: parsePensijilanAnalisisCsv,
        setData: setPensijilanAn,
        setLoading: setPensijilanAnLoading,
        setError: setPensijilanAnError,
      }),
      loadOne({
        demoPath: "data/analisis-optik-demo.csv",
        gidKey: "VITE_ANALISIS_GID_OPTIK",
        parse: parseOptikAnalisisCsv,
        setData: setOptikAn,
        setLoading: setOptikLoading,
        setError: setOptikError,
      }),
    ]);

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        setMaklumatAsasLoading(true);
        setMaklumatAsasError("");
        const text = await fetchAnalisisCsvText({
          gid: String(import.meta.env.VITE_MAKLUMAT_ASAS_GID ?? "").trim(),
          demoPath: "data/maklumat-asas-demo.csv",
        });
        if (cancelled) return;
        const { config, pegawaiRows } = parseMaklumatAsasCsv(text);
        setMaklumatAsasConfig(config);
        setMaklumatAsasPegawaiOverride(pegawaiRows);
      } catch (e) {
        if (cancelled) return;
        setMaklumatAsasError(
          e instanceof Error ? e.message : "Gagal memuatkan Maklumat Asas.",
        );
        setMaklumatAsasConfig({});
        setMaklumatAsasPegawaiOverride(null);
      } finally {
        if (!cancelled) setMaklumatAsasLoading(false);
      }
    };
    void load();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        setLamanWebLoading(true);
        setLamanWebError("");
        const res = await fetch(LAMAN_WEB_SEKOLAH_XLSX_URL);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const buf = await res.arrayBuffer();
        const parsed = parseLamanWebSekolahXlsx(buf);
        if (cancelled) return;
        setLamanWebRows(parsed);
      } catch {
        if (cancelled) return;
        setLamanWebError("Gagal memuatkan senarai laman web sekolah (Excel).");
        setLamanWebRows([]);
      } finally {
        if (!cancelled) setLamanWebLoading(false);
      }
    };
    load();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    const gid = String(import.meta.env.VITE_OSC_GID_LAMAN_WEB_SEKOLAH ?? "").trim();
    const id = import.meta.env.VITE_GOOGLE_SHEET_ID;

    const load = async () => {
      try {
        let text;
        if (id && gid) {
          text = await fetchAnalisisCsvText({ gid });
        } else if (!id) {
          text = await fetchAnalisisCsvText({
            demoPath: "data/laman-web-sekolah-sheet-demo.csv",
          });
        } else {
          if (!cancelled) setLamanWebSheetRows([]);
          return;
        }
        if (cancelled) return;
        setLamanWebSheetRows(parseLamanWebSekolahSheetCsv(text));
      } catch {
        if (!cancelled) setLamanWebSheetRows([]);
      }
    };
    void load();
    return () => {
      cancelled = true;
    };
  }, []);

  const ainsStats = useMemo(() => {
    if (!ainsAn) {
      return {
        approved: 0,
        rejected: 0,
        skApproved: 0,
        sjkcApproved: 0,
        sjktApproved: 0,
      };
    }
    return {
      approved: ainsAn.approved,
      rejected: ainsAn.rejected,
      skApproved: ainsAn.skApproved,
      sjkcApproved: ainsAn.sjkcApproved,
      sjktApproved: ainsAn.sjktApproved,
    };
  }, [ainsAn]);

  const dcsChart = useMemo(() => {
    if (!dcsAn) {
      return {
        tov: DCS_TOV_2024_DAERAH,
        kpi: DCS_KPI_2025_KEBANGSAAN,
        capai: DCS_CAPAI_2025_DAERAH,
        yMin: 58,
        yMax: 82,
      };
    }
    return {
      tov: dcsAn.tov,
      kpi: dcsAn.kpi,
      capai: dcsAn.capai,
      yMin: dcsAn.yMin,
      yMax: dcsAn.yMax,
    };
  }, [dcsAn]);

  const lamanWebEffectiveRows = useMemo(() => {
    const byCode = new Map();
    for (const r of lamanWebRows) {
      const k = normalizeLamanWebSchoolCode(r.code);
      if (k) {
        byCode.set(k, {
          code: r.code,
          name: r.name,
          website: r.website,
        });
      }
    }
    const orphans = lamanWebRows.filter((r) => !normalizeLamanWebSchoolCode(r.code));

    for (const r of lamanWebSheetRows) {
      const k = normalizeLamanWebSchoolCode(r.code);
      if (!k) continue;
      const website = normalizeSchoolWebsiteUrl(r.website || "");
      const nm = String(r.name ?? "").trim();
      const ex = byCode.get(k);
      if (ex) {
        byCode.set(k, {
          code: ex.code,
          name: nm || ex.name,
          website: website || ex.website,
        });
      } else {
        byCode.set(k, {
          code: String(r.code ?? "").trim() || k,
          name: nm || "—",
          website,
        });
      }
    }

    const merged = [...byCode.values()].sort((a, b) =>
      normalizeLamanWebSchoolCode(a.code).localeCompare(
        normalizeLamanWebSchoolCode(b.code),
        "ms",
        { sensitivity: "base" },
      ),
    );
    return [...merged, ...orphans];
  }, [lamanWebRows, lamanWebSheetRows]);

  const lamanWebVisibleRows = useMemo(() => {
    const q = lamanWebSearch.toLowerCase().trim();
    if (q) {
      return lamanWebEffectiveRows
        .filter((r) => {
          const hay = `${r.code} ${r.name} ${r.website}`.toLowerCase();
          return hay.includes(q);
        })
        .map(lamanWebRowWithFeaturedWebsiteIfAny);
    }
    const byNorm = new Map();
    for (const r of lamanWebEffectiveRows) {
      const k = normalizeLamanWebSchoolCode(r.code);
      if (k) byNorm.set(k, r);
    }
    return LAMAN_WEB_FEATURED.map((def) =>
      mergeLamanWebFeaturedRow(def, findLamanWebRowForFeatured(byNorm, def)),
    );
  }, [lamanWebEffectiveRows, lamanWebSearch]);

  const lamanWebStats = useMemo(() => {
    const total = lamanWebEffectiveRows.length;
    const withUrl = lamanWebEffectiveRows.filter((r) => r.website).length;
    return { total, withUrl };
  }, [lamanWebEffectiveRows]);

  return (
    <main className="relative min-h-screen px-4 py-8 md:px-10">
      {designMode === "geminiA" && (
        <>
          <div className="absolute inset-0 bg-[#eef2f7] bg-[url('/gemini-bg.png')] bg-cover bg-center bg-fixed" />
          <div className="absolute inset-0 bg-white/30" />
        </>
      )}
      {designMode === "neonDark" && (
        <>
          <div className="absolute inset-0 bg-[#060b14]" />
          <div className="absolute inset-0 bg-[radial-gradient(circle_at_20%_20%,rgba(0,229,255,0.22)_0%,transparent_55%)]" />
          <div className="absolute inset-0 bg-[linear-gradient(to_right,rgba(0,229,255,0.10)_1px,transparent_1px),linear-gradient(to_bottom,rgba(0,229,255,0.10)_1px,transparent_1px)] bg-[size:48px_48px]" />
        </>
      )}

      {/* Watermark USTP PPD Manjung — tiga mod; gelap + neon: glow sains-teknologi */}
      <div
        className="pointer-events-none fixed inset-0 z-0 flex items-center justify-center overflow-hidden"
        aria-hidden
      >
        <img
          src={USTP_WATERMARK_SRC}
          alt=""
          decoding="async"
          className={
            designMode === "geminiA"
              ? "h-auto w-[min(88vw,520px)] max-w-none select-none opacity-[0.13] ustp-watermark-day"
              : designMode === "neonDark"
                ? "h-auto w-[min(92vw,580px)] max-w-none select-none opacity-[0.11] ustp-watermark-glow-neon"
                : "h-auto w-[min(90vw,560px)] max-w-none select-none opacity-[0.085] ustp-watermark-glow-dark"
          }
        />
      </div>

      <div className="relative z-10 mx-auto max-w-7xl space-y-6">
        <header className="rounded-2xl border border-cyan-300/20 bg-slate-900/22 p-6 backdrop-blur-2xl">
          <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
            <div className="flex items-start gap-4">
              <img
                src="/logo.png"
                alt="USTP PPD Manjung Logo"
                className="mt-1 h-16 w-16 rounded-full border border-cyan-300/20 bg-white/5 p-1 shadow-[0_0_25px_rgba(0,229,255,0.12)] sm:h-20 sm:w-20 md:h-24 md:w-24"
              />
              <div className="min-w-0">
                <p className="text-xs tracking-[0.2em] text-cyan-300/80">
                  USTP DAERAH MANJUNG
                </p>
                <h1 className="mt-2 text-3xl font-bold text-white md:text-5xl">
                  ONE STOP CENTER
                </h1>
                <p className="mt-2 max-w-prose text-sm leading-relaxed text-slate-300 sm:text-base">
                  Pusat Paparan Laporan, Data dan Status.
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-2 sm:justify-end">
              <button
                type="button"
                onClick={() => setDesignMode("dark")}
                className={
                  designMode === "dark"
                    ? "rounded-lg border border-cyan-300/40 bg-slate-950/60 px-3 py-2 text-xs font-semibold text-cyan-200"
                    : "rounded-lg border border-cyan-300/20 bg-slate-950/30 px-3 py-2 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                }
              >
                Dark
              </button>

              <button
                type="button"
                onClick={() => setDesignMode("geminiA")}
                className={
                  designMode === "geminiA"
                    ? "rounded-lg border border-cyan-300/40 bg-slate-950/60 px-3 py-2 text-xs font-semibold text-cyan-200"
                    : "rounded-lg border border-cyan-300/20 bg-slate-950/30 px-3 py-2 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                }
              >
                Day
              </button>

              <button
                type="button"
                onClick={() => setDesignMode("neonDark")}
                className={
                  designMode === "neonDark"
                    ? "rounded-lg border border-cyan-300/40 bg-slate-950/60 px-3 py-2 text-xs font-semibold text-cyan-200"
                    : "rounded-lg border border-cyan-300/20 bg-slate-950/30 px-3 py-2 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                }
              >
                Neon Grid
              </button>
            </div>
          </div>
        </header>

        <section aria-label="OSC — Maklumat asas" className="space-y-3">
          <details
            name="osc-page-topik"
            defaultOpen
            className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
          >
            <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
              <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-slate-400/35 bg-slate-600/15 text-slate-200">
                <svg
                  className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                  fill="none"
                  viewBox="0 0 24 24"
                  stroke="currentColor"
                  strokeWidth={2}
                  aria-hidden
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    d="M2.25 21h19.5m-18-18v18m10.5-18v18m6-13.5V21M6.75 6.75h.75v.75h-.75v-.75zM6.75 12h.75v.75h-.75v-.75zM6.75 17.25h.75v.75h-.75v-.75zM12 6.75h.75v.75H12v-.75zM12 12h.75v.75H12v-.75zM12 17.25h.75v.75H12v-.75zM17.25 6.75h.75v.75h-.75v-.75zM17.25 12h.75v.75h-.75v-.75zM17.25 17.25h.75v.75h-.75v-.75zM3.75 6.75h.75v.75h-.75v-.75zM3.75 12h.75v.75h-.75v-.75z"
                  />
                </svg>
              </span>
              <div className="min-w-0 flex-1">
                <p className="text-sm font-semibold text-white">Maklumat Asas</p>
                <p className="text-xs text-slate-400">
                  Carta organisasi, maklumat pegawai, PKG, takwim
                </p>
              </div>
            </summary>
            <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/15 p-4">
              {maklumatAsasLoading ? (
                <p className="mb-3 text-sm text-slate-400">Memuatkan Maklumat Asas…</p>
              ) : null}
              {maklumatAsasError ? (
                <p className="mb-3 text-sm text-rose-300">
                  {maklumatAsasError} — memaparkan sandaran tempatan.
                </p>
              ) : null}
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
                <article className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-2 xl:col-span-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h2 className="text-lg font-semibold text-white">
                    {maklumatAsasDisplay.cartaTitle}
                  </h2>
                  <a
                    href={maklumatAsasDisplay.cartaFull}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                  >
                    Buka gambar penuh
                  </a>
                </div>
                <p className="mt-1 text-xs text-slate-500">
                  {maklumatAsasDisplay.cartaBlurb}
                </p>
                <div className="mt-3 max-h-[min(72vh,620px)] flex-1 overflow-auto rounded-xl border border-rose-900/25 bg-black/20">
                  {maklumatAsasDisplay.cartaPreview.kind === "iframe" ? (
                    <iframe
                      title={maklumatAsasDisplay.cartaTitle}
                      src={maklumatAsasDisplay.cartaPreview.src}
                      className="h-[min(72vh,620px)] w-full min-w-0"
                      style={{ border: 0, background: "#0b1220" }}
                      loading="lazy"
                      allowFullScreen
                    />
                  ) : (
                    <img
                      src={maklumatAsasDisplay.cartaPreview.src}
                      alt="Carta Organisasi USTP PPD Manjung"
                      className="w-full min-w-0 object-contain object-top"
                      loading="lazy"
                      decoding="async"
                    />
                  )}
                </div>
              </article>

              <article
                className="flex min-h-[min(52vh,520px)] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-2 xl:col-span-2"
              >
                <h2 className="text-lg font-semibold text-white">Maklumat Pegawai</h2>
                <input
                  value={keyword}
                  onChange={(e) => setKeyword(e.target.value)}
                  placeholder="Cari nama / jawatan / telefon"
                  className="mt-3 w-full rounded-lg border border-cyan-300/30 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none focus:border-cyan-300"
                />
                <div className="mt-4 flex-1 overflow-auto">
                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                    {filteredPegawai.map((p) => (
                      <button
                        key={p.nama}
                        type="button"
                        onClick={() => setSelectedPegawai(p)}
                        className="group text-left rounded-xl border border-cyan-400/20 bg-slate-950/40 p-3 transition hover:-translate-y-0.5 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                      >
                        <div className="flex items-start justify-between gap-3">
                          <div className="min-w-0">
                            {/* 不用 truncate，避免窄屏时出现省略号遮挡感 */}
                            <p className="break-words font-medium text-white leading-snug">
                              {p.nama}
                            </p>
                            <p className="mt-1 break-words text-xs text-slate-300">{p.jawatan}</p>
                            <p className="mt-2 break-words text-xs text-cyan-200">{p.telefon}</p>
                          </div>
                          <div className="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg border border-cyan-400/20 bg-cyan-400/5 text-cyan-200">
                            {p.nama
                              .split(" ")
                              .filter(Boolean)
                              .slice(0, 2)
                              .map((w) => w[0]?.toUpperCase())
                              .join("")}
                          </div>
                        </div>
                        <div className="mt-3 flex items-center justify-between">
                          <span className="text-xs text-slate-400">Klik untuk paparan penuh</span>
                          <span className="rounded-md border border-cyan-400/20 bg-cyan-400/10 px-2 py-1 text-xs text-cyan-200">
                            Open
                          </span>
                        </div>
                      </button>
                    ))}
                  </div>
                </div>
              </article>

              <article className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-1 xl:col-span-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h2 className="text-lg font-semibold text-white">
                    {maklumatAsasDisplay.pkgTitle}
                  </h2>
                  <a
                    href={maklumatAsasDisplay.pkgFull}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                  >
                    Buka gambar penuh
                  </a>
                </div>
                <p className="mt-1 text-xs text-slate-500">
                  {maklumatAsasDisplay.pkgBlurb}
                </p>
                <div className="mt-3 max-h-[min(72vh,620px)] flex-1 overflow-auto rounded-xl border border-rose-900/25 bg-black/20">
                  {maklumatAsasDisplay.pkgPreview.kind === "iframe" ? (
                    <iframe
                      title={maklumatAsasDisplay.pkgTitle}
                      src={maklumatAsasDisplay.pkgPreview.src}
                      className="h-[min(72vh,620px)] w-full min-w-0"
                      style={{ border: 0, background: "#0b1220" }}
                      loading="lazy"
                      allowFullScreen
                    />
                  ) : (
                    <img
                      src={maklumatAsasDisplay.pkgPreview.src}
                      alt="Maklumat PKG / COE Daerah Manjung"
                      className="w-full min-w-0 object-contain object-top"
                      loading="lazy"
                      decoding="async"
                    />
                  )}
                </div>
              </article>

              <article className="flex min-h-[380px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-1 xl:col-span-2">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                  <h2 className="text-lg font-semibold text-white">
                    {maklumatAsasDisplay.takwimTitle}
                  </h2>
                  <a
                    href={maklumatAsasDisplay.takwimFull}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                  >
                    Buka penuh
                  </a>
                </div>
                <div className="min-h-[min(52vh,420px)] flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
                  <iframe
                    loading="lazy"
                    title="Takwim USTP — pratontoh OSC"
                    src={maklumatAsasDisplay.takwimEmbed}
                    width="100%"
                    style={{ border: 0, background: "#0b1220" }}
                    className="h-full min-h-[min(52vh,420px)] w-full min-w-0"
                    allowFullScreen
                  />
                </div>
              </article>
              </div>
              <DetailsCollapseFooter />
            </div>
          </details>
        </section>

        <section
          aria-label="OSC — analisis, integrasi, hebahan digital COE, program pemerkasaan bacaan dan bahan sokongan"
          className="space-y-3"
        >
              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-indigo-400/35 bg-indigo-500/10 text-indigo-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">
                      Analisis Data
                    </p>
                    <p className="text-xs text-slate-400">
                      Subtopik: Data DELIMa, Status DCS, AINS, Pensijilan, OPTIK — data dari Google Sheet /
                      demo CSV; refresh selepas edit. Buka satu subtopik pada satu masa.
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <p className="mb-4 text-[11px] text-slate-500">
                    Klik subtopik di bawah untuk paparan penuh — tidak lagi digabungkan dalam satu skrin.
                  </p>

                  <details
                    name="osc-sub-analisis"
                    className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
                  >
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/30 bg-cyan-500/10 text-cyan-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Data DELIMa</p>
                        <p className="text-xs text-slate-400">
                          Carta purata bulanan, ringkasan daerah &amp; pautan Sheet
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <article className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
                        <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                          <h2 className="text-xl font-semibold text-white">Data DELIMa</h2>
                          {(delimaAn?.sourceUrl || DELIMA_GOOGLE_SHEET_URL) ? (
                            <a
                              href={delimaAn?.sourceUrl || DELIMA_GOOGLE_SHEET_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-cyan-400/30 bg-slate-950/40 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 hover:border-cyan-300/50"
                            >
                              Buka Google Sheet
                            </a>
                          ) : null}
                        </div>
                        {delimaAn?.intro ? (
                          <p className="mb-2 text-sm text-slate-400">{delimaAn.intro}</p>
                        ) : (
                          <p className="mb-2 text-sm text-slate-400">
                            Carta purata bulanan — data dari Google Sheet (refresh halaman selepas kemaskini).
                          </p>
                        )}
                        <div className="flex flex-1 flex-col gap-4">
                          {delimaAnLoading ? (
                            <p className="text-sm text-slate-300">Memuatkan data…</p>
                          ) : delimaAnError ? (
                            <p className="text-sm text-rose-200">{delimaAnError}</p>
                          ) : delimaAn &&
                            (delimaAn.seriesForChart.length > 0 ||
                              delimaAn.seriesDisplay.length > 0 ||
                              delimaAn.insight) ? (
                            <>
                              {delimaAn.insight ? (
                                <DelimaSheetInsightBlock
                                  insight={delimaAn.insight}
                                  kpiGuru={delimaAn.kpiGuru}
                                  kpiMurid={delimaAn.kpiMurid}
                                />
                              ) : null}
                              {delimaAn.seriesForChart.length > 0 || delimaAn.seriesDisplay.length > 0 ? (
                                <>
                                  <div className="rounded-xl border border-cyan-400/20 bg-slate-950/30 p-2 sm:p-3">
                                    <DelimaTrendChart
                                      series={
                                        delimaAn.seriesForChart.length > 0
                                          ? delimaAn.seriesForChart
                                          : delimaAn.seriesDisplay
                                      }
                                      height={220}
                                    />
                                  </div>
                                  <div className="rounded-xl border border-cyan-400/15 bg-slate-950/35 p-2 sm:p-3">
                                    <DelimaMonthAvgTable seriesDisplay={delimaAn.seriesDisplay} />
                                  </div>
                                </>
                              ) : (
                                <p className="text-sm text-slate-500">
                                  Tiada jadual purata bulan — tambah baris di bawah{" "}
                                  <span className="font-mono text-slate-400">month_label</span> dalam Sheet
                                  untuk carta &amp; jadual.
                                </p>
                              )}
                              {DELIMA_SHOW_SASARAN_KPI ? (
                                <div className="flex flex-wrap gap-2 pt-1">
                                  <StatusPill
                                    label={`Sasaran guru (${delimaAn.kpiGuru}%) — rujukan Sheet`}
                                    tone="good"
                                  />
                                  <StatusPill
                                    label={`Sasaran murid (${delimaAn.kpiMurid}%) — rujukan Sheet`}
                                    tone="good"
                                  />
                                </div>
                              ) : null}
                              {(delimaAn.sourceUrl || DELIMA_GOOGLE_SHEET_URL) ? (
                                <a
                                  href={delimaAn.sourceUrl || DELIMA_GOOGLE_SHEET_URL}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="block w-full rounded-xl border border-cyan-400/25 bg-slate-950/50 px-3 py-2.5 text-center text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                                >
                                  Paparan penuh
                                </a>
                              ) : null}
                            </>
                          ) : (
                            <p className="text-sm text-slate-400">
                              Tiada ringkasan daerah atau baris bulan dalam Sheet / CSV.
                            </p>
                          )}
                        </div>
                      </article>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details
                    name="osc-sub-analisis"
                    className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
                  >
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-amber-400/30 bg-amber-500/10 text-amber-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M7 12l3-3 3 3 4-4M8 21l4-4 4 4M3 4h18M4 4h16v12a1 1 0 01-1 1H5a1 1 0 01-1-1V4z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Status DCS</p>
                        <p className="text-xs text-slate-400">
                          Trend TOV → capai, KPI kebangsaan, gambar &amp; pautan Sheet
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <article className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
                        <h2 className="text-xl font-semibold text-white">Status DCS</h2>
                        <p className="mt-2 text-sm text-slate-300">
                          {dcsAn?.updatedText
                            ? `Kemaskini: ${dcsAn.updatedText}`
                            : "Carta & teks dikawal melalui Google Sheet — refresh halaman selepas edit."}
                        </p>
                        {dcsAnError ? (
                          <p className="mt-2 text-sm text-rose-200">{dcsAnError}</p>
                        ) : null}

                        <div className="mt-4 flex min-h-0 flex-1 flex-col rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3">
                          <DcsKpiLineChart
                            tov={dcsChart.tov}
                            kpi={dcsChart.kpi}
                            capai={dcsChart.capai}
                            yMin={dcsChart.yMin}
                            yMax={dcsChart.yMax}
                          />
                          <div className="mt-2 flex flex-wrap items-center justify-center gap-x-4 gap-y-1 text-[10px] text-slate-500">
                            <span className="inline-flex items-center gap-1.5">
                              <span
                                className="inline-block h-0 w-8 border-t-2 border-dashed border-amber-300/90"
                                aria-hidden
                              />
                              Garisan sasaran kebangsaan
                            </span>
                            <span className="inline-flex items-center gap-1.5">
                              <span
                                className="inline-block h-0.5 w-8 rounded-full bg-cyan-400"
                                aria-hidden
                              />
                              Capaian daerah
                            </span>
                          </div>
                          {dcsChart.capai >= dcsChart.kpi ? (
                            <p className="mt-2 text-center text-[11px] font-medium text-emerald-300/90">
                              Daerah melebihi sasaran kebangsaan
                            </p>
                          ) : null}
                        </div>

                        {dcsAn?.footer ? (
                          <p className="mt-4 whitespace-pre-line text-sm text-slate-400">{dcsAn.footer}</p>
                        ) : null}

                        <div className="mt-4 flex flex-wrap gap-2">
                          <StatusPill
                            label={`TOV 2024 (Daerah): ${dcsChart.tov}%`}
                            tone="warn"
                          />
                          <StatusPill
                            label={`KPI 2025 (Kebangsaan): ${dcsChart.kpi}%`}
                            tone="good"
                          />
                          <StatusPill
                            label={`Capai KPI 2025 (Daerah): ${dcsChart.capai}%`}
                            tone="good"
                          />
                        </div>

                        <div className="mt-4 flex flex-col gap-2 sm:flex-row">
                          {dcsAn?.sourceUrl ? (
                            <a
                              href={dcsAn.sourceUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="inline-flex w-full items-center justify-center rounded-xl border border-cyan-400/30 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 sm:flex-1"
                            >
                              {dcsAn.sourceLabel || "Buka sumber data"}
                            </a>
                          ) : null}
                          <button
                            type="button"
                            onClick={() => setIsDcsImageOpen(true)}
                            className="w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 sm:flex-1"
                          >
                            {dcsAn?.imageLabel || "Klik untuk lihat gambar penuh"}
                          </button>
                        </div>
                      </article>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-analisis" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-indigo-400/30 bg-indigo-500/10 text-indigo-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M11 3.055A9.001 9.001 0 1020.945 13H11V3.055z M20.488 9H15V3.512A9.025 9.025 0 0120.488 9z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Program Ains</p>
                        <p className="text-xs text-slate-400">
                          Statistik dari Sheet — senarai penuh di Sheet sumber
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                    <article
                      className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl"
                    >
                      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
                        <h2 className="text-xl font-semibold text-white">Program Ains</h2>
                        {ainsAn?.sourceUrl ? (
                          <a
                            href={ainsAn.sourceUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="shrink-0 rounded-lg border border-cyan-400/30 bg-slate-950/40 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 hover:border-cyan-300/50"
                          >
                            {ainsAn.sourceLabel || "Buka Sheet penuh"}
                          </a>
                        ) : null}
                      </div>
                      <p className="mb-3 text-sm text-slate-400">
                        Carta daripada nilai dalam Google Sheet — senarai pelajar dibuka di Sheet sumber.
                      </p>
                      <div className="flex flex-col gap-3">
                        <div className="rounded-xl border border-cyan-400/15 bg-slate-950/40 px-2 py-2">
                          <ProgramAinsCharts
                            stats={ainsStats}
                            loading={ainsLoading}
                            error={ainsError}
                          />
                        </div>
                        <div className="rounded-xl border border-cyan-400/20 bg-slate-950/20 p-3">
                          <div className="flex flex-wrap gap-2">
                            <StatusPill
                              label={`Approved: ${ainsStats.approved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`Rejected: ${ainsStats.rejected}`}
                              tone={ainsStats.rejected > 0 ? "bad" : "warn"}
                            />
                            <StatusPill
                              label={`SK Approved: ${ainsStats.skApproved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`SJKC Approved: ${ainsStats.sjkcApproved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`SJKT Approved: ${ainsStats.sjktApproved}`}
                              tone="good"
                            />
                          </div>
                        </div>
                      </div>
                    </article>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-analisis" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-emerald-400/30 bg-emerald-500/10 text-emerald-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Analisis Pensijilan Digital</p>
                        <p className="text-xs text-slate-400">
                          Ringkasan lokasi, jenis sekolah dan penyedia
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                    <article
                      className="flex min-h-[380px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl"
                    >
                      <div className="flex flex-wrap items-start justify-between gap-2">
                        <h2 className="text-xl font-semibold leading-snug text-white">
                          ANALISIS PENSIJILAN DIGITAL
                        </h2>
                        {pensijilanAn?.sourceUrl ? (
                          <a
                            href={pensijilanAn.sourceUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="shrink-0 rounded-lg border border-cyan-400/30 bg-slate-950/40 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 hover:border-cyan-300/50"
                          >
                            {pensijilanAn.sourceLabel || "Buka sumber"}
                          </a>
                        ) : null}
                      </div>
                      <p className="mt-2 text-sm text-slate-300">
                        {pensijilanAn?.intro ||
                          "Ringkasan data pensijilan mengikut lokasi, jenis sekolah dan penyedia."}
                      </p>
                      {pensijilanAnError ? (
                        <p className="mt-2 text-sm text-rose-200">{pensijilanAnError}</p>
                      ) : null}
                      <PensijilanDigitalSummary
                        data={
                          pensijilanAn?.locations && pensijilanAn?.schools
                            ? {
                                locations: pensijilanAn.locations,
                                schools: pensijilanAn.schools,
                              }
                            : undefined
                        }
                      />
                      <button
                        type="button"
                        onClick={() => setIsPensijilanDigitalImageOpen(true)}
                        className="mt-4 w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                      >
                        {pensijilanAn?.imageModalLabel || "Klik untuk lihat gambar penuh"}
                      </button>
                    </article>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-analisis" className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-sky-400/30 bg-sky-500/10 text-sky-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">AI Tools dalam DELIMa</p>
                        <p className="text-xs text-slate-400">
                          OPTIK 2 dari Sheet — berbanding KPI kebangsaan
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                    <article className="flex min-h-[380px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
                      <div className="flex flex-wrap items-start justify-between gap-2">
                        <h2 className="text-xl font-semibold text-white">AI Tools dalam DELIMa</h2>
                        {optikAn?.sourceUrl ? (
                          <a
                            href={optikAn.sourceUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="shrink-0 rounded-lg border border-cyan-400/30 bg-slate-950/40 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 hover:border-cyan-300/50"
                          >
                            {optikAn.sourceLabel || "Buka sumber"}
                          </a>
                        ) : null}
                      </div>
                      <p className="mt-2 text-sm text-slate-300">
                        Ringkasan OPTIK 2 berbanding garisan KPI kebangsaan — data dari Sheet.
                      </p>
                      {optikLoading ? (
                        <p className="mt-3 text-sm text-slate-400">Memuatkan…</p>
                      ) : optikError ? (
                        <p className="mt-3 text-sm text-rose-200">{optikError}</p>
                      ) : (
                        <AiToolsDelimaSummary sheet={optikAn} />
                      )}
                    </article>
                    <DetailsCollapseFooter />
                    </div>
                  </details>
                  <DetailsCollapseFooter />
                </div>

              </details>

              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-teal-400/35 bg-teal-500/10 text-teal-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">
                      Integrasi Teknologi Pendidikan
                    </p>
                    <p className="text-xs text-slate-400">
                      Subtopik: Kertas kerja · Laporan · JNJ (YTY, Pandai, Goolee, Minecraft,
                      Mikrobotik &amp; Edu Kit, kolaborasi &amp; Ana Muslim) · Impak JNJ
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <OscTopicSheetBody
                    sheetGid={import.meta.env.VITE_OSC_GID_INTEGRASI}
                    demoPath={OSC_TOPIK_DEMO_CSV.integrasi}
                    detailsName="osc-sub-integrasi"
                    loadingLabel="Memuatkan tab Integrasi daripada Google Sheet…"
                    fallback={
                      <>
                  <details name="osc-sub-integrasi" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-violet-400/25 bg-violet-500/10 text-violet-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Kertas kerja</p>
                        <p className="text-xs text-slate-400">
                          Bengkel Digital 2025 (PDF) · Eduspark COE Roadshow (Canva)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                Kertas Kerja Program
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Bengkel Digital 2025 — PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={KERTAS_KERJA_BENGKEL_DIGITAL_2025_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="Kertas Kerja Program — Bengkel Digital 2025"
                              src={driveGoogleFilePreviewUrl(KERTAS_KERJA_BENGKEL_DIGITAL_2025_PDF_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                Kertas Kerja Program Eduspark
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                COE Roadshow (Canva)
                              </p>
                            </div>
                            <a
                              href={KERTAS_KERJA_EDUSPARK_COE_ROADSHOW_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="Kertas Kerja Program Eduspark — COE Roadshow"
                              src={canvaViewEmbedUrl(KERTAS_KERJA_EDUSPARK_COE_ROADSHOW_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                              allow="fullscreen"
                            />
                          </div>
                        </article>
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-integrasi" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-amber-400/35 bg-amber-500/10 text-amber-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2h-1.5m-7.5 0H9"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Laporan</p>
                        <p className="text-xs text-slate-400">
                          Canva poster · Dron (PDF) · Khidmat bantu dron (Google Docs) ·
                          Pencetak 3D · Reka bentuk 3D · OPR Bengkel MYRC · OPR Bengkel Laman Web ·
                          OPR Mikrobotik &amp; Edu Kit (COE / JU RBT) (PDF Drive)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        {INTEGRASI_LAPORAN_CARDS.map(({ title, viewUrl, kind }) => (
                          <article
                            key={title}
                            className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                              <div className="min-w-0 pr-1">
                                <h3 className="text-sm font-semibold leading-snug text-white">
                                  {title}
                                </h3>
                                <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                  {kind === "gdoc"
                                    ? "LAPORAN SPb 8/6/2025 — Google Docs"
                                    : "PDF (Google Drive)"}
                                </p>
                              </div>
                              <a
                                href={viewUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                              <iframe
                                loading="lazy"
                                title={
                                  kind === "gdoc"
                                    ? `${title} — Google Docs`
                                    : `${title} — PDF`
                                }
                                src={
                                  kind === "gdoc"
                                    ? googleDocEmbedPreviewUrl(viewUrl)
                                    : driveGoogleFilePreviewUrl(viewUrl)
                                }
                                width="100%"
                                height="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="block h-full w-full"
                                allowFullScreen
                              />
                            </div>
                          </article>
                        ))}
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-integrasi" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-rose-400/30 bg-rose-500/10 text-rose-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">JNJ</p>
                        <p className="text-xs text-slate-400">
                          YTY · Pandai · Goolee · Minecraft · Mikrobotik &amp; Edu Kit · Manjung
                          &amp; Padang Terap · Ana Muslim (Seri Selamat, SG Tiram) — PDF / Docs
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        {INTEGRASI_JNJ_YTY_PDF_CARDS.map(({ viewUrl, blurb }) => (
                          <article
                            key={viewUrl}
                            className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                              <div className="min-w-0 pr-1">
                                <h3 className="text-sm font-semibold leading-snug text-white">
                                  YTY
                                </h3>
                                <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                  {blurb}
                                </p>
                              </div>
                              <a
                                href={viewUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                              <iframe
                                loading="lazy"
                                title={`JNJ YTY — ${blurb}`}
                                src={driveGoogleFilePreviewUrl(viewUrl)}
                                width="100%"
                                height="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="block h-full w-full"
                                allowFullScreen
                              />
                            </div>
                          </article>
                        ))}
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                Pandai Education
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Google Docs
                              </p>
                            </div>
                            <a
                              href={JNJ_PANDAI_EDUCATION_DOC_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="Pandai Education — Google Docs"
                              src={googleDocEmbedPreviewUrl(JNJ_PANDAI_EDUCATION_DOC_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                Goolee &amp; Must Read
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={JNJ_GOOLEE_MUSTREAD_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="Goolee & Must Read — PDF"
                              src={driveGoogleFilePreviewUrl(JNJ_GOOLEE_MUSTREAD_PDF_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                JNJ Minecraft
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                MINECRAFT VALE — PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={JNJ_MINECRAFT_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="JNJ Minecraft — MINECRAFT VALE PDF"
                              src={driveGoogleFilePreviewUrl(JNJ_MINECRAFT_PDF_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                JNJ Mikrobotik &amp; Edukit
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                jnj reka edukit — PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={JNJ_MIKROBOTIK_EDUKIT_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="JNJ Mikrobotik & Edukit — PDF"
                              src={driveGoogleFilePreviewUrl(JNJ_MIKROBOTIK_EDUKIT_PDF_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        {INTEGRASI_JNJ_DRIVE_PDF_CARDS.map(({ title, blurb, viewUrl }) => (
                          <article
                            key={viewUrl}
                            className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                              <div className="min-w-0 pr-1">
                                <h3 className="text-sm font-semibold leading-snug text-white">
                                  {title}
                                </h3>
                                <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                  {blurb}
                                </p>
                              </div>
                              <a
                                href={viewUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                              <iframe
                                loading="lazy"
                                title={`${title} — PDF`}
                                src={driveGoogleFilePreviewUrl(viewUrl)}
                                width="100%"
                                height="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="block h-full w-full"
                                allowFullScreen
                              />
                            </div>
                          </article>
                        ))}
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-integrasi" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-orange-400/35 bg-orange-500/10 text-orange-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Impak JNJ</p>
                        <p className="text-xs text-slate-400">
                          MYRC · Minecraft · Pandai Apps · Reka Edukit · Mikrobotik &amp; Reka
                          Edukit (PDF) — ringkasan pencapaian &amp; susulan
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                            <h3 className="text-sm font-semibold leading-snug text-white">
                              Impak JNJ MYRC
                            </h3>
                            <a
                              href={INTEGRASI_IMPAK_JNJ_MYRC_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] leading-relaxed text-slate-400">
                            Top 3 kategori sekolah rendah dan sekolah menengah pada peringkat
                            negeri; tempat kelima pada peringkat kebangsaan.
                          </p>
                          <div className="mt-auto max-h-[min(56vh,480px)] min-h-[180px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={INTEGRASI_IMPAK_JNJ_MYRC_IMG}
                              alt="Impak JNJ MYRC — ringkasan pencapaian"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                            <h3 className="text-sm font-semibold leading-snug text-white">
                              Impak JNJ Minecraft
                            </h3>
                            <a
                              href={INTEGRASI_IMPAK_JNJ_MINECRAFT_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] leading-relaxed text-slate-400">
                            Impak sekolah yang masuk bengkel: Raja Shariman, Methodist dan
                            Dindings. Tahun 2026: 6 SR terpilih untuk program yang sama —
                            jadual Mei 2026.
                          </p>
                          <div className="mt-auto max-h-[min(56vh,480px)] min-h-[180px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={INTEGRASI_IMPAK_JNJ_MINECRAFT_IMG}
                              alt="Impak JNJ Minecraft — ringkasan sekolah & susulan 2026"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                            <h3 className="text-sm font-semibold leading-snug text-white">
                              Impak JNJ Pandai Apps
                            </h3>
                            <a
                              href={INTEGRASI_IMPAK_JNJ_PANDAI_APP_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] leading-relaxed text-slate-400">
                            SK Seri Bayu Seri Manjung menang kategori murid peringkat
                            kebangsaan.
                          </p>
                          <div className="mt-auto max-h-[min(56vh,480px)] min-h-[180px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={INTEGRASI_IMPAK_JNJ_PANDAI_APP_IMG}
                              alt="Impak JNJ Pandai Apps — SK Seri Bayu Seri Manjung"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                            <h3 className="text-sm font-semibold leading-snug text-white">
                              IMPAK REKA EDUKIT
                            </h3>
                            <a
                              href={INTEGRASI_IMPAK_JNJ_REKA_EDUKIT_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] leading-relaxed text-slate-400">
                            Ringkasan impak JNJ Reka Edukit — rujuk gambar penuh untuk butiran.
                          </p>
                          <div className="mt-auto max-h-[min(56vh,480px)] min-h-[180px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={INTEGRASI_IMPAK_JNJ_REKA_EDUKIT_IMG}
                              alt="IMPAK REKA EDUKIT — JNJ"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                IMPAK MIKROBIT &amp; REKA EDUKIT
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Bengkel dan impak peringkat kebangsaan — PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={INTEGRASI_IMPAK_JNJ_MIKROBOTIK_REKA_EDUKIT_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="IMPAK Mikrobotik & Reka Edukit — PDF"
                              src={driveGoogleFilePreviewUrl(
                                INTEGRASI_IMPAK_JNJ_MIKROBOTIK_REKA_EDUKIT_PDF_URL,
                              )}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>
                      </>
                    }
                  />
                </div>
              </details>

              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-fuchsia-400/35 bg-fuchsia-500/10 text-fuchsia-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M11 5.882V19.24a1.76 1.76 0 01-3.417.592l-2.147-6.15M18 13a3 3 0 100-6M5.436 13.683A4.001 4.001 0 017 6h1c.256 0 .512.02.758.06C9.879 3.161 12.32 1 15 1c2.68 0 5.121 2.161 6.242 5.06.246-.04.502-.06.758-.06h1a4 4 0 013.424 6.868M15 13a3 3 0 11-6 0 3 3 0 016 0z"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-[13px] font-semibold leading-snug text-white">
                      Hebahan Pendidikan Digital melalui Peralatan COE kepada Komuniti Sekolah
                      dan Setempat
                    </p>
                    <p className="mt-0.5 text-[11px] font-medium leading-snug text-slate-300">
                      (Orang Awam)
                    </p>
                    <p className="mt-1.5 text-xs leading-relaxed text-slate-400">
                      Subtopik: Hari terbuka PPD Manjung · Karnival Digital (Manjung &amp; Pulau
                      Pangkor) · Pameran COE negeri &amp; kebangsaan
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <OscTopicSheetBody
                    sheetGid={import.meta.env.VITE_OSC_GID_HEBAHAN}
                    demoPath={OSC_TOPIK_DEMO_CSV.hebahan}
                    detailsName="osc-sub-hebahan"
                    loadingLabel="Memuatkan tab Hebahan daripada Google Sheet…"
                    fallback={
                      <>
                  <details name="osc-sub-hebahan" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/30 bg-cyan-500/10 text-cyan-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Hari terbuka PPD Manjung</p>
                        <p className="text-xs text-slate-400">Dua paparan gambar hebahan</p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-sm font-semibold text-white">
                              Hari Terbuka PPD Manjung
                            </h3>
                            <a
                              href={HEBAHAN_HARI_TERBUKA_PPD_MANJUNG_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <div className="mt-1 max-h-[min(72vh,620px)] min-h-[200px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={HEBAHAN_HARI_TERBUKA_PPD_MANJUNG_IMG}
                              alt="Hebahan — Hari Terbuka PPD Manjung"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-sm font-semibold text-white">
                              Pameran USTP (Hari Terbuka)
                            </h3>
                            <a
                              href={HEBAHAN_HARI_TERBUKA_PAMERAN_USTP_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <div className="mt-1 max-h-[min(72vh,620px)] min-h-[200px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={HEBAHAN_HARI_TERBUKA_PAMERAN_USTP_IMG}
                              alt="Hebahan — Hari Terbuka PPD Manjung, pameran USTP"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-hebahan" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-violet-400/30 bg-violet-500/10 text-violet-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M14.828 14.828a4 4 0 01-5.656 0M9 10h.01M15 10h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">
                          Karnival Digital Daerah Manjung dan Pulau Pangkor
                        </p>
                        <p className="text-xs text-slate-400">Manjung · Edisi Pulau Pangkor</p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-sm font-semibold text-white">
                              Karnival Digital COE PPD Manjung
                            </h3>
                            <a
                              href={HEBAHAN_KARNIVAL_DIGITAL_MANJUNG_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <div className="mt-1 max-h-[min(72vh,620px)] min-h-[200px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={HEBAHAN_KARNIVAL_DIGITAL_MANJUNG_IMG}
                              alt="Hebahan — Karnival Digital COE PPD Manjung"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-sm font-semibold text-white">
                              Karnival Digital — Edisi Pulau Pangkor
                            </h3>
                            <a
                              href={HEBAHAN_KARNIVAL_DIGITAL_PANGKOR_IMG}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                            >
                              Buka gambar penuh
                            </a>
                          </div>
                          <div className="mt-1 max-h-[min(72vh,620px)] min-h-[200px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                            <img
                              src={HEBAHAN_KARNIVAL_DIGITAL_PANGKOR_IMG}
                              alt="Hebahan — Karnival Digital COE PPD Manjung edisi Pulau Pangkor"
                              className="w-full min-w-0 object-contain object-top"
                              loading="lazy"
                              decoding="async"
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-hebahan" className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-amber-400/35 bg-amber-500/10 text-amber-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M19 21V5a2 2 0 00-2-2H7a2 2 0 00-2 2v16m14 0h2m-2 0h-5m-9 0H3m2 0h5M9 7h1m-1 4h1m4-4h1m-1 4h1m-5 10v-5a1 1 0 011-1h2a1 1 0 011 1v5m-4 0h4"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">
                          Pameran Peralatan COE Peringkat Negeri &amp; Kebangsaan
                        </p>
                        <p className="text-xs text-slate-400">Pameran peringkat negeri Perak</p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl lg:max-w-4xl">
                        <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                          <h3 className="text-sm font-semibold text-white">
                            Pameran Peralatan COE Negeri Perak
                          </h3>
                          <a
                            href={HEBAHAN_PAMERAN_COE_NEGERI_IMG}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                          >
                            Buka gambar penuh
                          </a>
                        </div>
                        <div className="mt-1 max-h-[min(72vh,620px)] min-h-[200px] flex-1 overflow-auto rounded-lg border border-cyan-400/15 bg-black/20">
                          <img
                            src={HEBAHAN_PAMERAN_COE_NEGERI_IMG}
                            alt="Hebahan — Pameran peralatan COE peringkat negeri Perak"
                            className="w-full min-w-0 object-contain object-top"
                            loading="lazy"
                            decoding="async"
                          />
                        </div>
                      </article>
                      <DetailsCollapseFooter />
                    </div>
                  </details>
                      </>
                    }
                  />
                </div>
              </details>

              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-blue-400/35 bg-blue-500/10 text-blue-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M21 12a9 9 0 01-9 9m9-9a9 9 0 00-9-9m9 9H3m9 9a9 9 0 01-9-9m9 9c1.657 0 3-4.03 3-9s-1.343-9-3-9m0 18c-1.657 0-3-4.03-3-9s1.343-9 3-9m-9 9a9 9 0 019-9"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">
                      Inisiatif Teknologi Maklumat
                    </p>
                    <p className="text-xs text-slate-400">
                      Subtopik: Google Classroom · JNJ (MAXIS, IAB, NADI) · Laman web sekolah
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <ItmLamanWebSekolahSection
                    loading={lamanWebLoading}
                    error={lamanWebError}
                    stats={lamanWebStats}
                    search={lamanWebSearch}
                    onSearchChange={(e) => setLamanWebSearch(e.target.value)}
                    visibleRows={lamanWebVisibleRows}
                    featuredCount={LAMAN_WEB_FEATURED.length}
                    mergedTotal={lamanWebEffectiveRows.length}
                  />
                  <OscTopicSheetBody
                    sheetGid={import.meta.env.VITE_OSC_GID_ITM}
                    demoPath={OSC_TOPIK_DEMO_CSV.itm}
                    detailsName="osc-sub-itm"
                    loadingLabel="Memuatkan tab ITM daripada Google Sheet…"
                    omitTrailingPageFooter
                    fallback={
                      <>
                  <details name="osc-sub-itm" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-emerald-400/35 bg-emerald-500/10 text-emerald-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M9.75 17L9 20l-1 1h8l-1-1-.75-3M3 13h18M5 17h14a2 2 0 002-2V5a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Google Classroom</p>
                        <p className="text-xs text-slate-400">
                          Papan kelas Google — log masuk akaun MOE / DELIMa
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        <a
                          href={GOOGLE_CLASSROOM_USTP_MANJUNG_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="group/card flex flex-col overflow-hidden rounded-xl border border-emerald-400/30 bg-slate-900/50 shadow-[0_12px_40px_rgba(0,0,0,0.35)] backdrop-blur-xl transition duration-200 hover:-translate-y-0.5 hover:border-emerald-400/55 hover:shadow-[0_16px_48px_rgba(16,185,129,0.12)]"
                        >
                          <div className="relative aspect-[16/10] overflow-hidden bg-slate-950">
                            <img
                              src={ITM_GOOGLE_CLASSROOM_PREVIEW_IMAGE}
                              alt="Pratontak Google Classroom — kelas USTP PPD Manjung"
                              loading="lazy"
                              decoding="async"
                              className="h-full w-full object-cover object-top transition duration-300 group-hover/card:scale-[1.02]"
                            />
                            <div className="pointer-events-none absolute inset-0 bg-gradient-to-t from-slate-950/80 via-transparent to-transparent" />
                            <span className="absolute left-2 top-2 rounded-md bg-emerald-600/95 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-white shadow-sm">
                              Classroom
                            </span>
                          </div>
                          <div className="flex flex-1 flex-col gap-1 p-3.5">
                            <h3 className="text-sm font-semibold text-white">
                              Google Classroom USTP Manjung
                            </h3>
                            <p className="text-[11px] leading-relaxed text-slate-400">
                              Pautan ke halaman utama Classroom (perlu log masuk).
                            </p>
                            <span className="mt-1 inline-flex items-center gap-1 text-xs font-semibold text-emerald-300 group-hover/card:text-emerald-200">
                              Buka Google Classroom
                              <span aria-hidden>→</span>
                            </span>
                          </div>
                        </a>
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-itm" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-rose-400/30 bg-rose-500/10 text-rose-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">JNJ</p>
                        <p className="text-xs text-slate-400">
                          MAXIS · IAB (Google Docs) · NADI (PDF Drive — pratontak &amp; pautan penuh)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">JNJ MAXIS</h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Google Docs
                              </p>
                            </div>
                            <a
                              href={ITM_JNJ_MAXIS_DOC_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="JNJ MAXIS — Google Docs"
                              src={googleDocEmbedPreviewUrl(ITM_JNJ_MAXIS_DOC_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">JNJ IAB</h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Google Docs
                              </p>
                            </div>
                            <a
                              href={ITM_JNJ_IAB_DOC_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="JNJ IAB — Google Docs"
                              src={googleDocEmbedPreviewUrl(ITM_JNJ_IAB_DOC_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">JNJ NADI</h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={ITM_JNJ_NADI_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="JNJ NADI — PDF"
                              src={driveGoogleFilePreviewUrl(ITM_JNJ_NADI_PDF_URL)}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>
                      </>
                    }
                  />
                  <DetailsCollapseFooter />
                </div>
              </details>

              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-orange-400/35 bg-orange-500/10 text-orange-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M12 6.042A8.967 8.967 0 006 3.75c-1.052 0-2.062.18-3 .512v14.25A8.987 8.987 0 016 18c2.305 0 4.408.867 6 2.292m0-14.25a8.966 8.966 0 016-2.292c1.052 0 2.062.18 3 .512v14.25A8.987 8.987 0 0018 18a8.967 8.967 0 00-6 2.292m0-14.25v14.25"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">
                      Pembudayaan Amalan Membaca
                    </p>
                    <p className="text-xs text-slate-400">
                      Subtopik: OPR 2025 (kad Sheet) · Program Inovasi · JNJ (Blink Book, Must Read)
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <OscTopicSheetBody
                    sheetGid={import.meta.env.VITE_OSC_GID_PEMBUDAYAAN_MEMBACA}
                    demoPath={OSC_TOPIK_DEMO_CSV.pembudayaanMembaca}
                    detailsName="osc-sub-membaca"
                    loadingLabel="Memuatkan tab Pembudayaan Membaca daripada Google Sheet…"
                    fallback={
                      <>
                  <details name="osc-sub-membaca" className="group mb-4 rounded-2xl border border-violet-400/25 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_20px_rgba(167,139,250,0.06)] transition-[border-color,box-shadow] duration-200 open:border-violet-400/45 open:shadow-[0_0_28px_rgba(167,139,250,0.1)]">
                      <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                        <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-violet-400/35 bg-violet-500/10 text-violet-200">
                          <svg
                            className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                            fill="none"
                            viewBox="0 0 24 24"
                            stroke="currentColor"
                            strokeWidth={2}
                            aria-hidden
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"
                            />
                          </svg>
                        </span>
                        <div className="min-w-0 flex-1">
                          <p className="text-sm font-semibold text-white">
                            Program Inovasi
                          </p>
                          <p className="text-xs text-slate-400">
                            Video (Google Drive) · Bahan Inovasi (Artsteps)
                          </p>
                        </div>
                      </summary>
                      <div className="overflow-hidden rounded-b-2xl border-t border-violet-400/15 bg-slate-950/30 px-4 py-4">
                        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                          <article className="flex flex-col rounded-xl border border-violet-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                            <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                              <div className="min-w-0 pr-1">
                                <h3 className="text-sm font-semibold leading-snug text-white">
                                  Video Inovasi
                                </h3>
                                <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                  Fail video — Google Drive (pratontak)
                                </p>
                              </div>
                              <a
                                href={PEMBUDAYAAN_PROGRAM_INOVASI_VIDEO_DRIVE_URL}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 self-start rounded-lg bg-gradient-to-r from-violet-400 to-fuchsia-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-violet-400/15 sm:h-[320px]">
                              <iframe
                                loading="lazy"
                                title="Video Inovasi — Google Drive"
                                src={driveGoogleFilePreviewUrl(
                                  PEMBUDAYAAN_PROGRAM_INOVASI_VIDEO_DRIVE_URL,
                                )}
                                width="100%"
                                height="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="block h-full w-full"
                                allowFullScreen
                              />
                            </div>
                          </article>
                          <a
                            href={PEMBUDAYAAN_PROGRAM_INOVASI_BAHAN_ARTSTEPS_URL}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="group/card flex flex-col overflow-hidden rounded-xl border border-violet-400/20 bg-slate-900/50 shadow-[0_12px_40px_rgba(0,0,0,0.35)] backdrop-blur-xl transition duration-200 hover:-translate-y-0.5 hover:border-violet-400/45 hover:shadow-[0_16px_48px_rgba(167,139,250,0.12)]"
                          >
                            <div className="relative aspect-[16/10] overflow-hidden border-b border-violet-400/15 bg-slate-950">
                              <img
                                src={PEMBUDAYAAN_PROGRAM_INOVASI_ARTSTEPS_PREVIEW_IMAGE}
                                alt="Pratontak Bahan Inovasi — Artsteps (V-BOOK 360°)"
                                loading="lazy"
                                decoding="async"
                                className="h-full w-full object-cover object-top transition duration-300 group-hover/card:scale-[1.02]"
                              />
                              <div className="pointer-events-none absolute inset-0 bg-gradient-to-t from-slate-950/80 via-transparent to-transparent" />
                              <span className="absolute left-2 top-2 rounded-md bg-violet-600/90 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-white shadow-sm">
                                Artsteps
                              </span>
                            </div>
                            <div className="flex flex-1 flex-col gap-1 p-3.5">
                              <h3 className="text-sm font-semibold text-white">
                                Bahan Inovasi
                              </h3>
                              <p className="text-[11px] leading-relaxed text-slate-400">
                                Paparan interaktif 3D — artsteps.com
                              </p>
                              <span className="mt-1 inline-flex items-center gap-1 text-xs font-semibold text-violet-300 group-hover/card:text-violet-200">
                                Buka di Artsteps
                                <span aria-hidden>→</span>
                              </span>
                            </div>
                          </a>
                        </div>
                        <DetailsCollapseFooter />
                      </div>
                    </details>

                    <details name="osc-sub-membaca" className="group rounded-2xl border border-cyan-400/25 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_20px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/45 open:shadow-[0_0_28px_rgba(0,229,255,0.1)]">
                      <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                        <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/35 bg-cyan-500/10 text-cyan-200">
                          <svg
                            className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                            fill="none"
                            viewBox="0 0 24 24"
                            stroke="currentColor"
                            strokeWidth={2}
                            aria-hidden
                          >
                            <path
                              strokeLinecap="round"
                              strokeLinejoin="round"
                              d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"
                            />
                          </svg>
                        </span>
                        <div className="min-w-0 flex-1">
                          <p className="text-sm font-semibold text-white">JNJ</p>
                          <p className="text-xs text-slate-400">
                            Blink Book (RAK MAYA) · JNJ Must Read (PDF Drive)
                          </p>
                        </div>
                      </summary>
                      <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/15 bg-slate-950/30 px-4 py-4">
                        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                          <a
                            href={PEMBUDAYAAN_JNJ_BLINK_BOOK_PSS_MAYA_URL}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="group/card flex flex-col overflow-hidden rounded-xl border border-cyan-400/20 bg-slate-900/50 shadow-[0_12px_40px_rgba(0,0,0,0.35)] backdrop-blur-xl transition duration-200 hover:-translate-y-0.5 hover:border-cyan-400/45 hover:shadow-[0_16px_48px_rgba(0,229,255,0.12)]"
                          >
                            <div className="relative aspect-[16/10] overflow-hidden border-b border-cyan-400/15 bg-slate-950">
                              <img
                                src={PEMBUDAYAAN_JNJ_BLINK_BOOK_PSS_MAYA_PREVIEW_IMAGE}
                                alt="Pratontak PSS Maya SESDA — rak digital Blink Book"
                                loading="lazy"
                                decoding="async"
                                className="h-full w-full object-cover object-top transition duration-300 group-hover/card:scale-[1.02]"
                              />
                              <div className="pointer-events-none absolute inset-0 bg-gradient-to-t from-slate-950/80 via-transparent to-transparent" />
                              <span className="absolute left-2 top-2 rounded-md bg-cyan-600/90 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-white shadow-sm">
                                JNJ Blink Book
                              </span>
                            </div>
                            <div className="flex flex-1 flex-col gap-1 p-3.5">
                              <h3 className="text-sm font-semibold text-white">
                                PSS Maya SESDA
                              </h3>
                              <p className="text-[11px] leading-relaxed text-slate-400">
                                Portal RAK MAYA — sesdamaya.rakmaya.com
                              </p>
                              <span className="mt-1 inline-flex items-center gap-1 text-xs font-semibold text-cyan-300 group-hover/card:text-cyan-200">
                                Buka laman
                                <span aria-hidden>→</span>
                              </span>
                            </div>
                          </a>
                          <a
                            href={PEMBUDAYAAN_JNJ_BLINK_BOOK_SK_SERI_SAMUDERA_URL}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="group/card flex flex-col overflow-hidden rounded-xl border border-cyan-400/20 bg-slate-900/50 shadow-[0_12px_40px_rgba(0,0,0,0.35)] backdrop-blur-xl transition duration-200 hover:-translate-y-0.5 hover:border-cyan-400/45 hover:shadow-[0_16px_48px_rgba(0,229,255,0.12)]"
                          >
                            <div className="relative aspect-[16/10] overflow-hidden border-b border-cyan-400/15 bg-slate-950">
                              <img
                                src={PEMBUDAYAAN_JNJ_BLINK_BOOK_SK_SERI_SAMUDERA_PREVIEW_IMAGE}
                                alt="Pratontak Rak Maya SK Seri Samudera — halaman bio"
                                loading="lazy"
                                decoding="async"
                                className="h-full w-full object-cover object-top transition duration-300 group-hover/card:scale-[1.02]"
                              />
                              <div className="pointer-events-none absolute inset-0 bg-gradient-to-t from-slate-950/80 via-transparent to-transparent" />
                              <span className="absolute left-2 top-2 rounded-md bg-cyan-600/90 px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide text-white shadow-sm">
                                JNJ Blink Book
                              </span>
                            </div>
                            <div className="flex flex-1 flex-col gap-1 p-3.5">
                              <h3 className="text-sm font-semibold text-white">
                                SK Seri Samudera — Bio
                              </h3>
                              <p className="text-[11px] leading-relaxed text-slate-400">
                                Halaman bio RAK MAYA — skeses.rakmaya.com/bio
                              </p>
                              <span className="mt-1 inline-flex items-center gap-1 text-xs font-semibold text-cyan-300 group-hover/card:text-cyan-200">
                                Buka laman
                                <span aria-hidden>→</span>
                              </span>
                            </div>
                          </a>
                          <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                            <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                              <div className="min-w-0 pr-1">
                                <h3 className="text-sm font-semibold leading-snug text-white">
                                  JNJ Must Read
                                </h3>
                                <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                  Goolee &amp; Must Read — PDF (Google Drive)
                                </p>
                              </div>
                              <a
                                href={JNJ_GOOLEE_MUSTREAD_PDF_URL}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                              <iframe
                                loading="lazy"
                                title="JNJ Must Read — Goolee & Must Read PDF"
                                src={driveGoogleFilePreviewUrl(JNJ_GOOLEE_MUSTREAD_PDF_URL)}
                                width="100%"
                                height="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="block h-full w-full"
                                allowFullScreen
                              />
                            </div>
                          </article>
                        </div>
                        <DetailsCollapseFooter />
                      </div>
                    </details>
                      </>
                    }
                  />
                </div>
              </details>

              <details
                name="osc-page-topik"
                className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
              >
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-emerald-400/35 bg-emerald-500/10 text-emerald-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M12 6.042A8.967 8.967 0 006 3.75c-1.052 0-2.062.18-3 .512v14.25A8.987 8.987 0 016 18c2.305 0 4.408.867 6 2.292m0-14.25a8.966 8.966 0 016-2.292c1.052 0 2.062.18 3 .512v14.25A8.987 8.987 0 0018 18a8.967 8.967 0 00-6 2.292m0-14.25v14.25"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">
                      Program Pemerkasaan Bacaan Murid
                    </p>
                    <p className="text-xs text-slate-400">
                      Subtopik: Video kreatif · Bicara Buku · Pembaca Bestari
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <OscTopicSheetBody
                    sheetGid={import.meta.env.VITE_OSC_GID_PEMERKASAAN}
                    demoPath={OSC_TOPIK_DEMO_CSV.pemerkasaan}
                    detailsName="osc-sub-pemerkasaan"
                    loadingLabel="Memuatkan tab Pemerkasaan daripada Google Sheet…"
                    fallback={
                      <>
                  <details name="osc-sub-pemerkasaan" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-rose-400/30 bg-rose-500/10 text-rose-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M15 10l4.553-2.276A1 1 0 0121 8.618v6.764a1 1 0 01-1.447.894L15 14M5 18h8a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Video kreatif</p>
                        <p className="text-xs text-slate-400">
                          Bahan video kreatif berkaitan pemerkasaan bacaan
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                BENGKEL PRODUKSI MEDIA DAN KANDUNGAN DIGITAL LENSA EDU PERAK 3.0
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                LAPORAN SPb 7/22/2025 — Google Docs
                              </p>
                            </div>
                            <a
                              href={PEMERKASAAN_VIDEO_KREATIF_LENSA_EDU_DOC_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="BENGKEL PRODUKSI MEDIA DAN KANDUNGAN DIGITAL LENSA EDU PERAK 3.0 — Google Docs"
                              src={googleDocEmbedPreviewUrl(
                                PEMERKASAAN_VIDEO_KREATIF_LENSA_EDU_DOC_URL,
                              )}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[280px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                IMPAK TVPSS
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                PDF (Google Drive)
                              </p>
                            </div>
                            <a
                              href={PEMERKASAAN_VIDEO_KREATIF_IMPAK_TVPSS_PDF_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="IMPAK TVPSS — PDF"
                              src={driveGoogleFilePreviewUrl(
                                PEMERKASAAN_VIDEO_KREATIF_IMPAK_TVPSS_PDF_URL,
                              )}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-pemerkasaan" className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-violet-400/25 bg-violet-500/10 text-violet-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M19 11a7 7 0 01-7 7m0 0a7 7 0 01-7-7m7 7v4m0 0H8m4 0h4m-4-8a3 3 0 01-3-3V5a3 3 0 116 0v6a3 3 0 01-3 3z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Bicara Buku</p>
                        <p className="text-xs text-slate-400">
                          Sesi bicara buku &amp; bahan sokongan
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch lg:max-w-4xl">
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl sm:col-span-2">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                Bicara Buku — siaran YouTube
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                Rakaman / live — buka penuh di YouTube jika embed tidak dimuatkan
                              </p>
                            </div>
                            <a
                              href={BICARA_BUKU_YOUTUBE_WATCH_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka di YouTube
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="Bicara Buku — Program Pemerkasaan Bacaan Murid (YouTube)"
                              src={BICARA_BUKU_YOUTUBE_EMBED_URL}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share"
                              allowFullScreen
                              referrerPolicy="strict-origin-when-cross-origin"
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details name="osc-sub-pemerkasaan" className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-amber-400/35 bg-amber-500/10 text-amber-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.887a1 1 0 00-1.176 0l-3.976 2.887c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.784-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Pembaca Bestari</p>
                        <p className="text-xs text-slate-400">
                          Program &amp; bahan pembaca bestari
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:items-stretch lg:max-w-4xl">
                        <article className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex min-h-[4.5rem] flex-wrap items-start justify-between gap-2">
                            <div className="min-w-0 pr-1">
                              <h3 className="text-sm font-semibold leading-snug text-white">
                                PENILAIAN PEMBACA BESTARI PERINGKAT NEGERI PERAK
                              </h3>
                              <p className="mt-0.5 text-[11px] leading-relaxed text-slate-500">
                                LAPORAN SPb 9/22/2025 — Google Docs
                              </p>
                            </div>
                            <a
                              href={PEMERKASAAN_PEMBACA_BESTARI_PENILAIAN_DOC_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 self-start rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="h-[280px] w-full shrink-0 overflow-hidden rounded-lg border border-cyan-400/15 sm:h-[320px]">
                            <iframe
                              loading="lazy"
                              title="PENILAIAN PEMBACA BESTARI PERINGKAT NEGERI PERAK — Google Docs"
                              src={googleDocEmbedPreviewUrl(
                                PEMERKASAAN_PEMBACA_BESTARI_PENILAIAN_DOC_URL,
                              )}
                              width="100%"
                              height="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="block h-full w-full"
                              allowFullScreen
                            />
                          </div>
                        </article>
                      </div>
                      <DetailsCollapseFooter />
                    </div>
                  </details>
                      </>
                    }
                  />
                </div>
              </details>

              <BahanSokonganPageSection />

            </section>

        {selectedPegawai && (
              <div
                className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
                role="dialog"
                aria-modal="true"
              >
                <div className="flex max-h-[90vh] w-full max-w-4xl flex-col overflow-hidden rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                  <div className="flex shrink-0 items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/40 p-4">
                    <div className="min-w-0">
                      <p className="truncate text-lg font-semibold text-white">{selectedPegawai.nama}</p>
                      <p className="truncate text-sm text-slate-300">{selectedPegawai.jawatan}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      {selectedPegawai.detailUrl || selectedPegawai.detailImage ? (
                        <a
                          href={
                            selectedPegawai.detailUrl || selectedPegawai.detailImage
                          }
                          target="_blank"
                          rel="noreferrer"
                          className="rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
                        >
                          Buka dalam tab baru
                        </a>
                      ) : null}
                      <button
                        type="button"
                        onClick={() => setSelectedPegawai(null)}
                        className="rounded-lg border border-cyan-400/20 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                      >
                        Tutup
                      </button>
                    </div>
                  </div>

                  <div className="relative flex-1 overflow-auto">
                    <PegawaiProfilePreview pegawai={selectedPegawai} />
                  </div>

                  <div className="shrink-0 border-t border-cyan-400/10 bg-slate-950/40 p-4">
                    <p className="text-sm text-slate-300">
                      Telefon: <span className="text-cyan-200">{selectedPegawai.telefon}</span>
                    </p>
                  </div>
                </div>
              </div>
        )}

        {isDcsImageOpen && (
            <div
              className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
              role="dialog"
              aria-modal="true"
            >
              <div className="w-full max-w-4xl overflow-hidden rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                <div className="flex items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/40 p-4">
                  <div className="min-w-0">
                    <p className="truncate text-lg font-semibold text-white">Status DCS - Gambar</p>
                    <p className="truncate text-sm text-slate-300">DCS Tahap 3 (keatas)</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setIsDcsImageOpen(false)}
                    className="rounded-lg border border-cyan-400/20 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                  >
                    Tutup
                  </button>
                </div>

                <div className="p-4">
                  <img
                    alt="Status DCS Tahap 3 (ke atas) infografik"
                    src={driveGoogleImageUrl(dcsAn?.imageUrl || "/assets/status-dcs.png")}
                    className="w-full h-auto rounded-xl border border-cyan-400/10 bg-black/20"
                  />
                </div>
              </div>
            </div>
          )}

          {isPensijilanDigitalImageOpen && (
            <div
              className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
              role="dialog"
              aria-modal="true"
            >
              <div className="max-h-[92vh] w-full max-w-6xl overflow-hidden overflow-y-auto rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                <div className="sticky top-0 z-10 flex items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/90 p-4 backdrop-blur-md">
                  <div className="min-w-0">
                    <p className="text-lg font-semibold text-white">
                      ANALISIS PENSIJILAN DIGITAL
                    </p>
                    <p className="truncate text-sm text-slate-300">Infografik penuh</p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setIsPensijilanDigitalImageOpen(false)}
                    className="shrink-0 rounded-lg border border-cyan-400/20 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                  >
                    Tutup
                  </button>
                </div>

                <div className="p-4">
                  <img
                    alt="Analisis pensijilan digital: jenis sekolah, lokasi dan pensijilan Google, Apple, Microsoft"
                    src={driveGoogleImageUrl(pensijilanAn?.imageUrl || PENSIJILAN_DIGITAL_IMAGE)}
                    className="w-full h-auto rounded-xl border border-cyan-400/10 bg-black/20"
                  />
                </div>
              </div>
            </div>
          )}

      </div>
    </main>
  );
}
