import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

/** Logo rasmi PPD (PNG telus) — watermark latar tiga mod reka */
const USTP_WATERMARK_SRC = "/assets/ustp-ppd-manjung-watermark.png";
/** Carta Organisasi & Maklumat PKG / COE — infografik rasmi (Buku Pengurusan USTP) */
const CARTA_ORGANISASI_IMG = "/assets/carta-organisasi-ustp-ppd-manjung.png";
const MAKLUMAT_PKG_COE_IMG = "/assets/maklumat-pkg-coe-daerah-manjung.png";

const TAKWIM_EMBED =
  "https://lookerstudio.google.com/embed/reporting/9a5abf82-8012-42b8-aef5-cc1cebebbfe6/page/p_o99byp19xc";
const PELAPORAN_DPD_EMBED =
  "https://lookerstudio.google.com/embed/reporting/97c54e64-01ea-495c-be82-300adf618bc6/page/JbWhE";

/** Kalendar kumpulan USTP (zon Asia/Kuala Lumpur) */
const USTP_CALENDAR_EMBED =
  "https://calendar.google.com/calendar/embed?src=c_07ea831973519ec6379185af0a2fd2053aeec6d5c15fab56dc24461b74e5c2e2%40group.calendar.google.com&ctz=Asia%2FKuala_Lumpur";

// Program Ains (Looker Studio)
const PROGRAM_AINS_CSV_URL = "/data/users_PERAK.csv";

// Data DELIMa (Excel — kemas kini dengan mengganti fail di public/data)
const DELIMA_XLSX_URL = "/data/delima-ppd-manjung.xlsx";
const LAMAN_WEB_SEKOLAH_XLSX_URL = "/data/laman-web-sekolah-bengkel-responses.xlsx";
/** OPR Amalan Membaca — Manjung sahaja (JSON daripada scripts/build-opr-amalan-membaca-manjung.mjs) */
const OPR_AMALAN_MEMBACA_JSON_URL = "/data/opr-amalan-membaca-manjung.json";
const DELIMA_GOOGLE_SHEET_URL =
  "https://docs.google.com/spreadsheets/d/1pGLkWr8Vt4kPW7haQ-_AZ-trpptimI-5r038vWbdblk/edit?gid=0#gid=0";
/** Sasaran rujukan: guru selaras KPI DCS; murid lebih rendah kerana corak penggunaan berbeza (laras jika perlu). */
const DELIMA_KPI_GURU_PCT = 78;
const DELIMA_KPI_MURID_PCT = 65;

/** Pasangan lajur % guru / % murid bagi setiap potongan masa (indeks 0 = baris data). */
const DELIMA_MONTH_DEF = [
  { label: "Apr", g: 6, m: 7 },
  { label: "Mei", g: 8, m: 9 },
  { label: "Jun", g: 10, m: 11 },
  { label: "Jul", g: 12, m: 13 },
  { label: "Ogs¹", g: 14, m: 15 },
  { label: "Ogs²", g: 16, m: 17 },
  { label: "Okt¹", g: 21, m: 22 },
  { label: "Okt²", g: 23, m: 24 },
  { label: "Nov", g: 25, m: 26 },
  { label: "Dis", g: 27, m: 28 },
];

/** Paparan carta / jadual: langkau Ogs¹ & Okt¹; kekalkan potongan kedua sahaja. */
const DELIMA_SKIP_CHART_LABELS = new Set(["Ogs¹", "Okt¹"]);

const toNumber = (v) => {
  const n = Number(String(v ?? "").trim());
  return Number.isFinite(n) ? n : 0;
};

const shortenText = (v, maxLen) => {
  const s = String(v ?? "");
  if (s.length <= maxLen) return s;
  if (maxLen <= 3) return s.slice(0, maxLen);
  return `${s.slice(0, maxLen - 3)}...`;
};

function DetailsCollapseFooter() {
  return (
    <div className="flex justify-center border-t border-cyan-400/10 bg-slate-950/35 px-4 py-3">
      <button
        type="button"
        className="text-xs font-semibold text-cyan-300/90 underline-offset-2 hover:text-cyan-200 hover:underline"
        onClick={(e) => {
          const d = e.currentTarget.closest("details");
          if (d) d.removeAttribute("open");
        }}
      >
        Tutup bahagian
      </button>
    </div>
  );
}

function parseCsvText(csvText) {
  // CSV is comma-separated, without quoted commas in this dataset.
  const lines = csvText.replace(/\r/g, "").split("\n").filter(Boolean);
  if (lines.length < 2) return [];

  const headers = lines[0].split(",").map((h) => h.trim().replace(/^"|"$/g, ""));

  const idx = (name) => headers.indexOf(name);

  const iState = idx("state");
  const iDistrict = idx("district");
  const iCode = idx("code");
  const iSchool = idx("school");
  const iLevel = idx("school level");
  const iType = idx("school type");
  const iName = idx("name");
  const iRejected = idx("rejected");
  const iApproved = idx("approved");

  const required = [
    iDistrict,
    iCode,
    iSchool,
    iLevel,
    iType,
    iName,
    iRejected,
    iApproved,
  ];
  if (required.some((x) => x < 0)) return [];

  const keep = [];
  for (let li = 1; li < lines.length; li++) {
    const cols = lines[li].split(",");
    if (cols.length < headers.length) continue;

    const district = cols[iDistrict] ?? "";
    // Daerah Manjung (district names like "PPD MANJUNG")
    if (!String(district).toLowerCase().includes("manjung")) continue;

    keep.push({
      state: iState >= 0 ? cols[iState] : "",
      district,
      code: cols[iCode],
      school: cols[iSchool],
      schoolLevel: cols[iLevel],
      schoolType: cols[iType],
      name: cols[iName],
      rejected: toNumber(cols[iRejected]),
      approved: toNumber(cols[iApproved]),
    });
  }

  return keep;
}

function delimaCellNum(v) {
  const n = Number(String(v ?? "").trim());
  return Number.isFinite(n) ? n : null;
}

/**
 * Bilangan kali khidmat/bengkel: setiap segmen dipisah dengan "/" dikira 1.
 * Contoh: "3.3.2025/11.3.25" = 2. Teks satu blok (cth. Ladap … 30.10.2025) = 1.
 */
function countDelimaBantuSegments(tarikhBantu) {
  const raw = String(tarikhBantu ?? "").replace(/\r\n/g, " ").trim();
  if (!raw) return 0;
  return raw
    .split("/")
    .map((p) => p.trim())
    .filter((p) => p.length > 0).length;
}

/** Baris data bermula indeks 2 (selepas 2 baris tajuk). Lajur sepadan fail PPD Manjung. */
function parseDelimaXlsx(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = wb.SheetNames.includes("DATA DELIMA 2025")
    ? "DATA DELIMA 2025"
    : wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const out = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    if (!Array.isArray(r)) continue;
    const kod = String(r[1] ?? "").trim();
    const nama = String(r[2] ?? "").trim();
    if (!kod || !nama) continue;
    if (!/^[A-Z]{2,4}\d{3,5}$/i.test(kod)) continue;

    const tovGuru = delimaCellNum(r[3]);
    const tovMurid = delimaCellNum(r[4]);
    const tarikhBantu = String(r[5] ?? "")
      .replace(/\r\n/g, " ")
      .trim();
    const bantuKali = countDelimaBantuSegments(tarikhBantu);
    const bilGuru = delimaCellNum(r[18]);
    const bilDashboard = delimaCellNum(r[19]);
    const bilGuruAktif = delimaCellNum(r[20]);
    const novGuru = delimaCellNum(r[25]);
    const novMurid = delimaCellNum(r[26]);
    const disGuru = delimaCellNum(r[27]);
    const disMurid = delimaCellNum(r[28]);

    let aktivitiPct = null;
    if (bilGuru != null && bilGuru > 0 && bilGuruAktif != null) {
      aktivitiPct = (bilGuruAktif / bilGuru) * 100;
    }

    const monthly = DELIMA_MONTH_DEF.map(({ label, g, m }) => ({
      label,
      guru: delimaCellNum(r[g]),
      murid: delimaCellNum(r[m]),
    }));

    out.push({
      kod,
      nama,
      tovGuru,
      tovMurid,
      tarikhBantu,
      bantuKali,
      bilGuru,
      bilDashboard,
      bilGuruAktif,
      novGuru,
      novMurid,
      disGuru,
      disMurid,
      aktivitiPct,
      monthly,
    });
  }
  return out;
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
  return {
    ...base,
    code: base.code?.trim() ? base.code : def.code,
    name: (base.name && String(base.name).trim()) || def.fallbackName,
    website: normalizeSchoolWebsiteUrl(def.url),
  };
}

function lamanWebRowWithFeaturedWebsiteIfAny(r) {
  const k = normalizeLamanWebSchoolCode(r.code);
  const o = k ? LAMAN_WEB_FEATURED_URL_BY_CODE[k] : "";
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

function meanFinite(nums) {
  const v = nums.filter((x) => x != null && Number.isFinite(x));
  if (!v.length) return null;
  return v.reduce((a, b) => a + b, 0) / v.length;
}

/** Anak panah perbandingan dalam mata peratus (pp). */
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

// Epelaporan (Canva) — `/view` + `?embed` untuk iframe
/** Bahan Rujukan — Epelaporan 2022 */
const EPelaporan_2022_LINK =
  "https://www.canva.com/design/DAE5EDf_XPY/y4cvQai_KnNMwM2ModqGRA/view";
/** Bahan Rujukan — Epelaporan 2023 */
const EPelaporan_2023_LINK =
  "https://www.canva.com/design/DAHFBKPyUho/vDbVidnAHom0TJcNXIl5Sg/view";
/** Kad embed — Epelaporan 2024 */
const EPelaporan_2024_LINK =
  "https://www.canva.com/design/DAHFAwfv5qI/kDTtol21C9NoGT8X7TL8dw/view";
const EPelaporan_2025_LINK =
  "https://www.canva.com/design/DAGgurDIiVU/tQr9dAKI1fjuF7iyS73LMQ/view";

/** Epelaporan Canva mengikut tahun + pelaporan tambahan (Takwim Looker) — untuk Bahan Sokongan */
const EPELAPORAN_BY_YEAR = [
  { year: 2022, viewUrl: EPelaporan_2022_LINK },
  { year: 2023, viewUrl: EPelaporan_2023_LINK },
  { year: 2024, viewUrl: EPelaporan_2024_LINK },
  { year: 2025, viewUrl: EPelaporan_2025_LINK },
];

/** Buku Pengurusan USTP — Canva mengikut tahun (tertib: 2023 → 2026) */
const BOOK_PENGURUSAN_AUTHOR = "Rujhan Alwi";
const BOOK_PENGURUSAN_BY_YEAR = [
  {
    year: 2023,
    viewUrl:
      "https://www.canva.com/design/DAFWv3QoB9g/SiacIkkrSFJG7FuO65g-Jg/view",
  },
  {
    year: 2024,
    viewUrl:
      "https://www.canva.com/design/DAF28MMwqT4/dV65v3f2G1G9V6A_KXM5Bg/view",
  },
  {
    year: 2025,
    viewUrl:
      "https://www.canva.com/design/DAGZuhOHzIM/kiiF86Fawa-ChTVvuH7DaQ/view",
  },
  {
    year: 2026,
    viewUrl:
      "https://www.canva.com/design/DAHFBMI56W8/rdm4z_hoorUd10XE3HgQfA/view",
  },
];

/** Penyebaran dasar — MJ3PD & JKPA (Canva), susunan paparan */
const PENYEBARAN_DASAR_CANVA_CARDS = [
  {
    key: "mj3pd-2025",
    title: "MJ3PD 2025",
    blurb: "Bahan penyebaran dasar (Canva).",
    viewUrl:
      "https://www.canva.com/design/DAGjdXwaoXw/eudwsuvBJEx7J7D-5zwGrQ/view",
  },
  {
    key: "jkpa-2026",
    title: "JKPA 2026",
    blurb: "Bahan dasar JKPA (Canva).",
    viewUrl:
      "https://www.canva.com/design/DAG-vuv7xWw/ZSoPv8yyr_7gMStIqeAgjA/view",
  },
  {
    key: "jkpa-2025",
    title: "JKPA 2025",
    blurb: "Bahan dasar JKPA (Canva).",
    viewUrl:
      "https://www.canva.com/design/DAGf7B_YivM/FDYNAE-lIPOuUKLrg9qH1g/view",
  },
  {
    key: "jkpa-2024",
    title: "JKPA 2024",
    blurb: "Bahan dasar JKPA (Canva).",
    viewUrl:
      "https://www.canva.com/design/DAFrG17dkI0/s_JU059YoGnI6UMs5q56Vw/view",
  },
];

function canvaViewEmbedUrl(viewUrl) {
  return viewUrl.includes("?") ? `${viewUrl}&embed` : `${viewUrl}?embed`;
}

/** Google Drive — PDF surat pemerkasaan DELIMa */
const SURAT_PEMERKASAAN_DELIMA_2023_2024_URL =
  "https://drive.google.com/file/d/1lyQd8IOz4lmgAFo9Lz2ZK9bKDZQNPy4F/view?usp=sharing";
const SURAT_PEMERKASAAN_DELIMA_2025_2026_URL =
  "https://drive.google.com/file/d/196YTk8wLq0qvdfSRVIFLxeQ_wLhTjjC-/view?usp=sharing";
/** Surat arahan khidmat sokongan — ARAHAN TUGAS USTP 2026 (PDF Drive) */
const SURAT_ARAHAN_KHIDMAT_SOKONGAN_USTP_2026_URL =
  "https://drive.google.com/file/d/1Ycx8VFxQryk0Gt-XR2mKhRLEugAdpQkC/view?usp=sharing";

function driveGoogleFilePreviewUrl(viewUrl) {
  const s = String(viewUrl ?? "");
  const fileD = s.match(/\/file\/d\/([^/?]+)/);
  if (fileD) return `https://drive.google.com/file/d/${fileD[1]}/preview`;
  if (s.includes("drive.google.com")) {
    const openId = s.match(/[?&]id=([^&]+)/);
    if (openId) return `https://drive.google.com/file/d/${openId[1]}/preview`;
  }
  return s;
}

function googleDocEmbedPreviewUrl(docUrl) {
  const m = docUrl.match(/\/document\/d\/([^/?]+)/);
  return m ? `https://docs.google.com/document/d/${m[1]}/preview` : docUrl;
}

/** Kad surat dalam Bahan Sokongan — Surat punca kuasa */
const SURAT_PUNCA_KUASA_CARDS = [
  {
    key: "delima-2324",
    title: "Surat Pemerkasaan DELIMa 2023–2024",
    blurb: "Punca kuasa pemerkasaan DELIMa bagi tempoh 2023–2024.",
    viewUrl: SURAT_PEMERKASAAN_DELIMA_2023_2024_URL,
  },
  {
    key: "delima-2526",
    title: "Surat Pemerkasaan DELIMa 2025–2026",
    blurb: "Punca kuasa pemerkasaan DELIMa bagi tempoh 2025–2026.",
    viewUrl: SURAT_PEMERKASAAN_DELIMA_2025_2026_URL,
  },
  {
    key: "arahan-ustp-2026",
    title: "Arahan Tugas USTP 2026",
    blurb: "Surat arahan khidmat sokongan (PDF).",
    viewUrl: SURAT_ARAHAN_KHIDMAT_SOKONGAN_USTP_2026_URL,
  },
];

const ANALISIS_PENARAFAN_KENDIRI_URL =
  "https://sites.google.com/moe-dl.edu.my/perakdrift/kluster/pendigitalan/instrumen-pendigitalan/analisis-penarafan-kendiri";
const OPPR_PENDIDIKAN_PERAK_URL =
  "https://sites.google.com/moe-dl.edu.my/perakdrift/kluster/pendigitalan/instrumen-pendigitalan/oppr-pendigitalan-perak?authuser=0";
const PENCAPAIAN_USTP_2025_URL =
  "https://www.canva.com/design/DAGxVjtQuhg/FTA-gZELFpsmrLn3HyhRnQ/view";
const PENCAPAIAN_USTP_2025_EMBED = `${PENCAPAIAN_USTP_2025_URL}?embed`;
const TIKTOK_USTP_MANJUNG_URL =
  "https://www.tiktok.com/@ustpmanjung1?is_from_webapp=1&sender_device=pc";
const YOUTUBE_USTP_MANJUNG_URL =
  "https://www.youtube.com/channel/UC00YHEDSN_X5xVGqV9b4rmw";
/** Pratontak skrin — fail dalam `public/` */
const VIDEO_CARD_TIKTOK_IMAGE = `${import.meta.env.BASE_URL}video-card-tiktok.png`;
const VIDEO_CARD_YOUTUBE_IMAGE = `${import.meta.env.BASE_URL}video-card-youtube.png`;
const PDP_DIGITAL_CARD_RUANG_ILMU_IMAGE = `${import.meta.env.BASE_URL}pdp-digital-card-ruang-ilmu.png`;
const RUANG_ILMU_RESOURCE_URL = "https://ruangilmu.moe-dl.edu.my/resource";

/** Kad dalam Bahan Sokongan — Bahan pdp digital */
const BAHAN_PDP_DIGITAL_CARDS = [
  {
    key: "tiktok",
    href: TIKTOK_USTP_MANJUNG_URL,
    platform: "TikTok",
    title: "USTP PPD Manjung",
    subtitle: "@ustpmanjung1 · Centre Of Excellent",
    cta: "Buka TikTok",
    previewSrc: VIDEO_CARD_TIKTOK_IMAGE,
    previewAlt: "Pratontak TikTok USTP PPD Manjung",
    ringClass: "border-rose-400/25 hover:border-rose-400/50",
    badgeClass: "bg-rose-500/90 text-white",
  },
  {
    key: "youtube",
    href: YOUTUBE_USTP_MANJUNG_URL,
    platform: "YouTube",
    title: "USTP PPD Manjung",
    subtitle: "Saluran rasmi · video program & ceramah",
    cta: "Tonton di YouTube",
    previewSrc: VIDEO_CARD_YOUTUBE_IMAGE,
    previewAlt: "Pratontak YouTube USTP PPD Manjung",
    ringClass: "border-red-500/30 hover:border-red-400/55",
    badgeClass: "bg-red-600/95 text-white",
  },
  {
    key: "ruang-ilmu",
    href: RUANG_ILMU_RESOURCE_URL,
    platform: "Ruang Ilmu",
    title: "Bank sumber DELIMa",
    subtitle: "Panduan, bank soalan & bahan mengikut subjek · log masuk ID DELIMa",
    cta: "Buka Ruang Ilmu",
    previewSrc: PDP_DIGITAL_CARD_RUANG_ILMU_IMAGE,
    previewAlt: "Pratontak laman Ruang Ilmu DELIMa",
    ringClass: "border-sky-400/35 hover:border-sky-400/60",
    badgeClass: "bg-sky-600/95 text-white",
  },
];
const KAD_PENGHARGAAN_PPD_MANJUNG_URL =
  "https://www.canva.com/design/DAGa60sQAjg/bmiRnetB9hWhfU20s6eOkQ/watch";
/** Untuk iframe — `/view?embed` (pautan penuh kekal `/watch`) */
const KAD_PENGHARGAAN_EMBED = `${KAD_PENGHARGAAN_PPD_MANJUNG_URL.replace(/\/watch\/?$/, "/view")}?embed`;
const SUCCESS_STORY_USTP_URL =
  "https://www.canva.com/design/DAGl_Rzw6so/PWBQvkwQprFOIQ12LZ1TZQ/view";
const SUCCESS_STORY_USTP_EMBED = `${SUCCESS_STORY_USTP_URL}?embed`;
const MAJLIS_APRESIASI_DIGITAL_URL =
  "https://www.canva.com/design/DAG3I0mucKY/ehUSnXhpz06UxXWmTN8-wA/edit?utm_content=DAG3I0mucKY&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton";
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

/** Integrasi Teknologi Pendidikan — Laporan (PDF Google Drive) */
const INTEGRASI_LAPORAN_CARDS = [
  {
    title: "Canva Grafik Fasa 1",
    viewUrl:
      "https://drive.google.com/file/d/1cHQpv3ddBx-RLfvOCDn71_VGaLy15aDW/view?usp=drive_link",
  },
  {
    title: "Canva Grafik Fasa 2",
    viewUrl:
      "https://drive.google.com/file/d/1vnq-bV0QSfGzdCyOxB2ZbahycHeV5mKK/view?usp=drive_link",
  },
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
    title: "Pencetak 3D",
    viewUrl:
      "https://drive.google.com/file/d/11w7hSfjDNVr2v8NtRbXYfTc-QNiJ6PTK/view?usp=drive_link",
  },
  {
    title: "Reka Bentuk 3D",
    viewUrl:
      "https://drive.google.com/file/d/1sXYW-8zGpgqJz4-Y7eGTGgPHtKljC-fy/view?usp=drive_link",
  },
];

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
    // 2) detailImage: link atau path gambar (PNG/JPG) untuk dipaparkan dalam modal
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
function PensijilanDigitalSummary() {
  const locSum = PENSIJILAN_BY_LOCATION.reduce((a, [, n]) => a + n, 0);
  return (
    <div className="mt-4 space-y-3 rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3 text-xs text-slate-300">
      <div>
        <p className="mb-1.5 text-[10px] font-semibold uppercase tracking-wide text-slate-500">
          Jumlah mengikut lokasi
        </p>
        <div className="flex flex-wrap gap-2">
          {PENSIJILAN_BY_LOCATION.map(([label, n]) => (
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
              {PENSIJILAN_BY_SCHOOL.map(([label, n]) => (
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

function AiToolsDelimaSummary() {
  const k = AI_TOOLS_SNAPSHOT_OPTIK;
  const p = AI_TOOLS_PERGERAKAN;
  const line = AI_TOOLS_KPI_KEBANGSAAN;

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
        <p className="mt-2 text-[10px] leading-relaxed text-slate-500">
          AR2 (Okt) melebihi sasaran {line}% berbanding AR1; titik akhir ialah peratus selesai snapshot {k.asAt}
          (bukan skala sama dengan TOV/AR).
        </p>
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
function DcsKpiLineChart() {
  const W = 400;
  const H = 168;
  const pad = { l: 40, r: 28, t: 18, b: 44 };
  const innerW = W - pad.l - pad.r;
  const innerH = H - pad.t - pad.b;
  const yMin = 58;
  const yMax = 82;
  const tov = DCS_TOV_2024_DAERAH;
  const kpi = DCS_KPI_2025_KEBANGSAAN;
  const capai = DCS_CAPAI_2025_DAERAH;

  const yAt = (pct) =>
    pad.t + (1 - (pct - yMin) / (yMax - yMin)) * innerH;

  const x0 = pad.l;
  const x1 = pad.l + innerW;
  /** Teks "Capai 2025" di bawah — ancar tengah sedikit ke kiri supaya tidak terpotong di tepi SVG */
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

  // Program Ains: CSV-based stats (since Looker iframe may be blocked)
  const [programRows, setProgramRows] = useState([]);
  const [programLoading, setProgramLoading] = useState(true);
  const [programError, setProgramError] = useState("");
  const [programSearch, setProgramSearch] = useState("");
  const [programModalOpen, setProgramModalOpen] = useState(false);
  const [programModalType, setProgramModalType] = useState("ALL"); // ALL | SK | SJKC | SJKT
  const [programModalStatus, setProgramModalStatus] = useState("ALL"); // ALL | APPROVED | REJECTED

  const [delimaRows, setDelimaRows] = useState([]);
  const [delimaLoading, setDelimaLoading] = useState(true);
  const [delimaError, setDelimaError] = useState("");
  const [delimaModalOpen, setDelimaModalOpen] = useState(false);
  const [delimaSearch, setDelimaSearch] = useState("");

  const [lamanWebRows, setLamanWebRows] = useState([]);
  const [lamanWebLoading, setLamanWebLoading] = useState(true);
  const [lamanWebError, setLamanWebError] = useState("");
  const [lamanWebSearch, setLamanWebSearch] = useState("");

  const [oprMembaca, setOprMembaca] = useState(null);
  const [oprMembacaLoading, setOprMembacaLoading] = useState(true);
  const [oprMembacaError, setOprMembacaError] = useState("");
  const [oprMembacaSearch, setOprMembacaSearch] = useState("");
  const [oprMembacaDimensi, setOprMembacaDimensi] = useState("ALL");
  const [oprMembacaStatus, setOprMembacaStatus] = useState("ALL");

  // Design mode:
  // - "dark": current version (深色背景)
  // - "geminiA": 浅色背景图 + 深色卡片（方案A）
  // - "neonDark": 深色背景 + 霓虹网格（第三个选项）
  const [designMode, setDesignMode] = useState("dark");

  const filteredPegawai = useMemo(() => {
    const k = keyword.toLowerCase().trim();
    if (!k) return pegawaiData;
    return pegawaiData.filter(
      (p) =>
        p.nama.toLowerCase().includes(k) ||
        p.jawatan.toLowerCase().includes(k) ||
        p.telefon.includes(k)
    );
  }, [keyword]);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        setProgramLoading(true);
        setProgramError("");
        const res = await fetch(PROGRAM_AINS_CSV_URL);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const text = await res.text();
        const parsed = parseCsvText(text);
        if (cancelled) return;
        setProgramRows(parsed);
      } catch {
        if (cancelled) return;
        setProgramError("Gagal memuatkan data CSV Program Ains.");
        setProgramRows([]);
      } finally {
        if (!cancelled) setProgramLoading(false);
      }
    };
    load();
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    const load = async () => {
      try {
        setDelimaLoading(true);
        setDelimaError("");
        const res = await fetch(DELIMA_XLSX_URL);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const buf = await res.arrayBuffer();
        const parsed = parseDelimaXlsx(buf);
        if (cancelled) return;
        setDelimaRows(parsed);
      } catch {
        if (cancelled) return;
        setDelimaError("Gagal memuatkan fail Excel DELIMa.");
        setDelimaRows([]);
      } finally {
        if (!cancelled) setDelimaLoading(false);
      }
    };
    load();
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
    const load = async () => {
      try {
        setOprMembacaLoading(true);
        setOprMembacaError("");
        const res = await fetch(OPR_AMALAN_MEMBACA_JSON_URL);
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const data = await res.json();
        if (!data || !Array.isArray(data.items)) {
          throw new Error("Format JSON tidak sah.");
        }
        if (cancelled) return;
        setOprMembaca(data);
      } catch {
        if (cancelled) return;
        setOprMembacaError("Gagal memuatkan rekod OPR Amalan Membaca (Manjung).");
        setOprMembaca(null);
      } finally {
        if (!cancelled) setOprMembacaLoading(false);
      }
    };
    load();
    return () => {
      cancelled = true;
    };
  }, []);

  const programVisibleRows = useMemo(() => {
    const q = programSearch.toLowerCase().trim();
    if (!q) return programRows;
    return programRows.filter((r) => {
      const hay = `${r.name} ${r.school} ${r.code}`.toLowerCase();
      return hay.includes(q);
    });
  }, [programRows, programSearch]);

  const programModalRows = useMemo(() => {
    let rows = programVisibleRows;

    if (programModalType !== "ALL") {
      const t = programModalType.toUpperCase().trim();
      rows = rows.filter(
        (r) => String(r.schoolType ?? "").toUpperCase().trim() === t
      );
    }

    if (programModalStatus === "APPROVED") {
      rows = rows.filter((r) => toNumber(r.approved) > 0);
    } else if (programModalStatus === "REJECTED") {
      rows = rows.filter((r) => toNumber(r.rejected) > 0);
    }

    return rows;
  }, [programVisibleRows, programModalStatus, programModalType]);

  const programStats = useMemo(() => {
    const rows = programRows;
    const approved = rows.reduce((acc, r) => acc + toNumber(r.approved), 0);
    const rejected = rows.reduce((acc, r) => acc + toNumber(r.rejected), 0);

    const sumApprovedByType = (t) =>
      rows.reduce((acc, r) => {
        const st = String(r.schoolType ?? "").toUpperCase().trim();
        return st === t ? acc + toNumber(r.approved) : acc;
      }, 0);
    const skApproved = sumApprovedByType("SK");
    const sjkcApproved = sumApprovedByType("SJKC");
    const sjktApproved = sumApprovedByType("SJKT");

    return {
      approved,
      rejected,
      skApproved,
      sjkcApproved,
      sjktApproved,
    };
  }, [programRows]);

  const lamanWebVisibleRows = useMemo(() => {
    const q = lamanWebSearch.toLowerCase().trim();
    if (q) {
      return lamanWebRows
        .filter((r) => {
          const hay = `${r.code} ${r.name} ${r.website}`.toLowerCase();
          return hay.includes(q);
        })
        .map(lamanWebRowWithFeaturedWebsiteIfAny);
    }
    const byNorm = new Map();
    for (const r of lamanWebRows) {
      const k = normalizeLamanWebSchoolCode(r.code);
      if (k) byNorm.set(k, r);
    }
    return LAMAN_WEB_FEATURED.map((def) =>
      mergeLamanWebFeaturedRow(def, findLamanWebRowForFeatured(byNorm, def)),
    );
  }, [lamanWebRows, lamanWebSearch]);

  const lamanWebStats = useMemo(() => {
    const total = lamanWebRows.length;
    const withUrl = lamanWebRows.filter((r) => r.website).length;
    return { total, withUrl };
  }, [lamanWebRows]);

  const oprMembacaItems = useMemo(
    () => oprMembaca?.items ?? [],
    [oprMembaca],
  );

  const oprMembacaFacetDimensi = useMemo(() => {
    const s = new Set();
    for (const r of oprMembacaItems) {
      const d = String(r.dimensi ?? "").trim();
      if (d) s.add(d);
    }
    return [...s].sort((a, b) => a.localeCompare(b, "ms"));
  }, [oprMembacaItems]);

  const oprMembacaFacetStatus = useMemo(() => {
    const s = new Set();
    for (const r of oprMembacaItems) {
      const t = String(r.statusOpr ?? "").trim();
      if (t) s.add(t);
    }
    return [...s].sort((a, b) => a.localeCompare(b, "ms"));
  }, [oprMembacaItems]);

  const oprMembacaVisible = useMemo(() => {
    const q = oprMembacaSearch.toLowerCase().trim();
    return oprMembacaItems.filter((r) => {
      if (oprMembacaDimensi !== "ALL") {
        if (
          String(r.dimensi ?? "").trim().toUpperCase() !==
          oprMembacaDimensi.toUpperCase()
        ) {
          return false;
        }
      }
      if (oprMembacaStatus !== "ALL") {
        if (
          String(r.statusOpr ?? "").trim().toUpperCase() !==
          oprMembacaStatus.toUpperCase()
        ) {
          return false;
        }
      }
      if (!q) return true;
      const hay = `${r.namaProgram} ${r.sekolah} ${r.pegawaiPelapor} ${r.nosiriOpr} ${r.tarikhMula} ${r.tarikhTamat}`.toLowerCase();
      return hay.includes(q);
    });
  }, [
    oprMembacaItems,
    oprMembacaSearch,
    oprMembacaDimensi,
    oprMembacaStatus,
  ]);

  const oprMembacaStats = useMemo(() => {
    const total = oprMembacaItems.length;
    const denganPautan = oprMembacaItems.filter((r) =>
      String(r.oprLink ?? "").startsWith("http"),
    ).length;
    const byStatus = {};
    for (const r of oprMembacaItems) {
      const k = String(r.statusOpr ?? "").trim() || "—";
      byStatus[k] = (byStatus[k] ?? 0) + 1;
    }
    return { total, denganPautan, byStatus };
  }, [oprMembacaItems]);

  const delimaInsight = useMemo(() => {
    const n = delimaRows.length;
    if (!n) {
      return {
        schools: 0,
        avgTovGuru: null,
        avgTovMurid: null,
        avgAktiviti: null,
        avgDisGuru: null,
        avgDisMurid: null,
        avgNovGuru: null,
        avgNovMurid: null,
        bantuSessionsTotal: 0,
        bantuSchoolsWithRekod: 0,
        series: [],
        seriesDisplay: [],
        schoolsGuruGteTov: 0,
        schoolsMuridGteTov: 0,
        kpiGuruOk: false,
        kpiMuridOk: false,
      };
    }
    const sum = (arr) => arr.reduce((a, b) => a + b, 0);
    const tovG = delimaRows.map((r) => r.tovGuru).filter((x) => x != null);
    const tovM = delimaRows.map((r) => r.tovMurid).filter((x) => x != null);
    const akt = delimaRows.map((r) => r.aktivitiPct).filter((x) => x != null);
    const dG = delimaRows.map((r) => r.disGuru).filter((x) => x != null);
    const dM = delimaRows.map((r) => r.disMurid).filter((x) => x != null);
    const novG = delimaRows.map((r) => r.novGuru).filter((x) => x != null);
    const novM = delimaRows.map((r) => r.novMurid).filter((x) => x != null);

    const avgTovGuru = tovG.length ? sum(tovG) / tovG.length : null;
    const avgTovMurid = tovM.length ? sum(tovM) / tovM.length : null;
    const avgDisGuru = dG.length ? sum(dG) / dG.length : null;
    const avgDisMurid = dM.length ? sum(dM) / dM.length : null;
    const avgNovGuru = novG.length ? sum(novG) / novG.length : null;
    const avgNovMurid = novM.length ? sum(novM) / novM.length : null;

    const bantuSessionsTotal = delimaRows.reduce(
      (acc, r) => acc + (r.bantuKali ?? 0),
      0
    );
    const bantuSchoolsWithRekod = delimaRows.filter((r) => r.bantuKali > 0)
      .length;

    let schoolsGuruGteTov = 0;
    let schoolsMuridGteTov = 0;
    for (const r of delimaRows) {
      if (r.tovGuru != null && r.disGuru != null && r.disGuru >= r.tovGuru) {
        schoolsGuruGteTov += 1;
      }
      if (r.tovMurid != null && r.disMurid != null && r.disMurid >= r.tovMurid) {
        schoolsMuridGteTov += 1;
      }
    }

    const series = DELIMA_MONTH_DEF.map(({ label }, mi) => ({
      label,
      guru: meanFinite(delimaRows.map((r) => r.monthly[mi]?.guru)),
      murid: meanFinite(delimaRows.map((r) => r.monthly[mi]?.murid)),
    }));

    const seriesDisplay = DELIMA_MONTH_DEF.map((m, mi) => {
      let chartLabel = m.label;
      if (m.label === "Okt²") chartLabel = "Okt";
      else if (m.label === "Ogs²") chartLabel = "Ogs";
      return {
        label: m.label,
        chartLabel,
        guru: series[mi]?.guru ?? null,
        murid: series[mi]?.murid ?? null,
      };
    }).filter((p) => !DELIMA_SKIP_CHART_LABELS.has(p.label));

    return {
      schools: n,
      avgTovGuru,
      avgTovMurid,
      avgAktiviti: akt.length ? sum(akt) / akt.length : null,
      avgDisGuru,
      avgDisMurid,
      avgNovGuru,
      avgNovMurid,
      bantuSessionsTotal,
      bantuSchoolsWithRekod,
      series,
      seriesDisplay,
      schoolsGuruGteTov,
      schoolsMuridGteTov,
      kpiGuruOk:
        avgDisGuru != null && avgDisGuru + 1e-6 >= DELIMA_KPI_GURU_PCT,
      kpiMuridOk:
        avgDisMurid != null && avgDisMurid + 1e-6 >= DELIMA_KPI_MURID_PCT,
    };
  }, [delimaRows]);

  const filteredDelimaRows = useMemo(() => {
    const q = delimaSearch.toLowerCase().trim();
    if (!q) return delimaRows;
    return delimaRows.filter((r) => {
      const hay = `${r.kod} ${r.nama}`.toLowerCase();
      return hay.includes(q);
    });
  }, [delimaRows, delimaSearch]);

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
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2 xl:grid-cols-4">
                <article className="flex min-h-[320px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-2 xl:col-span-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h2 className="text-lg font-semibold text-white">
                    Carta Organisasi
                  </h2>
                  <a
                    href={CARTA_ORGANISASI_IMG}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                  >
                    Buka gambar penuh
                  </a>
                </div>
                <p className="mt-1 text-xs text-slate-500">
                  Organisasi USTP PPD Manjung — hierarki PPD, PKG dan COE.
                </p>
                <div className="mt-3 max-h-[min(72vh,620px)] flex-1 overflow-auto rounded-xl border border-rose-900/25 bg-black/20">
                  <img
                    src={CARTA_ORGANISASI_IMG}
                    alt="Carta Organisasi USTP PPD Manjung"
                    className="w-full min-w-0 object-contain object-top"
                    loading="lazy"
                    decoding="async"
                  />
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
                  <h2 className="text-lg font-semibold text-white">Maklumat PKG</h2>
                  <a
                    href={MAKLUMAT_PKG_COE_IMG}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg border border-rose-800/50 bg-rose-950/40 px-2.5 py-1 text-[11px] font-semibold text-rose-100 hover:border-rose-600/60"
                  >
                    Buka gambar penuh
                  </a>
                </div>
                <p className="mt-1 text-xs text-slate-500">
                  Maklumat COE Daerah Manjung (AQA1001–AQA1005) — kemudahan &
                  hubungi.
                </p>
                <div className="mt-3 max-h-[min(72vh,620px)] flex-1 overflow-auto rounded-xl border border-rose-900/25 bg-black/20">
                  <img
                    src={MAKLUMAT_PKG_COE_IMG}
                    alt="Maklumat PKG / COE Daerah Manjung"
                    className="w-full min-w-0 object-contain object-top"
                    loading="lazy"
                    decoding="async"
                  />
                </div>
              </article>

              <article className="flex min-h-[380px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl md:col-span-1 xl:col-span-2">
                <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                  <h2 className="text-lg font-semibold text-white">Takwim</h2>
                  <a
                    href={USTP_CALENDAR_EMBED}
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
                    src={USTP_CALENDAR_EMBED}
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

        <section aria-label="OSC — analisis hingga bahan sokongan" className="space-y-3">
              <details className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40">
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
                      DCS · AINS · DELIMa · Pensijilan · AI Tools (OPTIK 2)
                    </p>
                  </div>
                </summary>
                <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <p className="mb-4 text-xs font-medium uppercase tracking-wide text-slate-500">
                    Data & analisis utama
                  </p>
                  <div className="grid grid-cols-1 gap-6 xl:grid-cols-2">
                    <article
                      role="button"
                      tabIndex={0}
                      onClick={(e) => {
                        if (e.target.closest("a,button")) return;
                        if (!delimaLoading && !delimaError && delimaRows.length > 0) {
                          setDelimaModalOpen(true);
                        }
                      }}
                      onKeyDown={(e) => {
                        if (e.key !== "Enter" && e.key !== " ") return;
                        e.preventDefault();
                        if (!delimaLoading && !delimaError && delimaRows.length > 0) {
                          setDelimaModalOpen(true);
                        }
                      }}
                      className="flex min-h-[520px] cursor-pointer flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl transition hover:border-cyan-400/45 hover:shadow-[0_0_28px_rgba(0,229,255,0.1)] focus:outline-none focus-visible:ring-2 focus-visible:ring-cyan-400/50"
                    >
                      <div className="mb-2 flex flex-wrap items-start justify-between gap-2">
                        <div>
                          <h2 className="text-xl font-semibold text-white">Data DELIMa</h2>
                        </div>
                        <a
                          href={DELIMA_GOOGLE_SHEET_URL}
                          target="_blank"
                          rel="noreferrer"
                          onClick={(e) => e.stopPropagation()}
                          className="shrink-0 rounded-lg border border-cyan-400/30 bg-slate-950/40 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 hover:border-cyan-300/50"
                        >
                          Google Sheet
                        </a>
                      </div>

                      <div className="flex flex-1 flex-col gap-2 rounded-xl border border-cyan-400/20 bg-slate-950/20 p-3">
                        {delimaLoading ? (
                          <p className="text-sm text-slate-300">Memuatkan data...</p>
                        ) : delimaError ? (
                          <p className="text-sm text-rose-200">{delimaError}</p>
                        ) : (
                          <>
                            <div className="grid grid-cols-2 gap-2 text-xs sm:grid-cols-3">
                              <div className="rounded-lg border border-cyan-400/15 bg-slate-950/40 px-2 py-1.5">
                                <p className="text-slate-500">Sekolah</p>
                                <p className="text-sm font-semibold text-white">
                                  {delimaInsight.schools}
                                </p>
                              </div>
                              <div className="rounded-lg border border-cyan-400/15 bg-slate-950/40 px-2 py-1.5">
                                <p className="text-slate-500">Khidmat bantu</p>
                                <p className="text-sm font-semibold text-cyan-200">
                                  {delimaInsight.bantuSessionsTotal}{" "}
                                  <span className="font-normal text-slate-500">kali</span>
                                </p>
                                <p className="mt-0.5 text-[10px] text-slate-500">
                                  {delimaInsight.bantuSchoolsWithRekod} sekolah ada rekod
                                </p>
                              </div>
                              <div className="col-span-2 rounded-lg border border-cyan-400/15 bg-slate-950/40 px-2 py-1.5 sm:col-span-1">
                                <p className="text-slate-500">Capai ≥ TOV (Dis)</p>
                                <p className="text-sm font-semibold text-white">
                                  G {delimaInsight.schoolsGuruGteTov}/{delimaInsight.schools}{" "}
                                  <span className="text-slate-500">·</span> M{" "}
                                  {delimaInsight.schoolsMuridGteTov}/{delimaInsight.schools}
                                </p>
                              </div>
                            </div>

                            <div className="space-y-1.5 border-t border-cyan-400/10 pt-2 text-xs">
                              <div className="flex flex-wrap items-center justify-between gap-2">
                                <span className="text-slate-400">
                                  Guru: TOV {fmtPct1(delimaInsight.avgTovGuru)} → Dis{" "}
                                  {fmtPct1(delimaInsight.avgDisGuru)}
                                </span>
                                <DelimaDeltaPp
                                  before={delimaInsight.avgTovGuru}
                                  after={delimaInsight.avgDisGuru}
                                  title="Purata daerah: TOV → Disember"
                                />
                              </div>
                              <div className="flex flex-wrap items-center justify-between gap-2">
                                <span className="text-slate-400">
                                  Murid: TOV {fmtPct1(delimaInsight.avgTovMurid)} → Dis{" "}
                                  {fmtPct1(delimaInsight.avgDisMurid)}
                                </span>
                                <DelimaDeltaPp
                                  before={delimaInsight.avgTovMurid}
                                  after={delimaInsight.avgDisMurid}
                                  title="Purata daerah: TOV → Disember"
                                />
                              </div>
                              <div className="flex flex-wrap items-center justify-between gap-2">
                                <span className="text-slate-400">
                                  Guru: Nov {fmtPct1(delimaInsight.avgNovGuru)} → Dis{" "}
                                  {fmtPct1(delimaInsight.avgDisGuru)}
                                </span>
                                <DelimaDeltaPp
                                  before={delimaInsight.avgNovGuru}
                                  after={delimaInsight.avgDisGuru}
                                  title="Purata daerah: November → Disember"
                                />
                              </div>
                            </div>

                            <div className="flex flex-wrap gap-2 pt-1">
                              <StatusPill
                                label={`Sasaran Dis guru (${DELIMA_KPI_GURU_PCT}%): ${delimaInsight.kpiGuruOk ? "Capai" : "Belum"}`}
                                tone={delimaInsight.kpiGuruOk ? "good" : "warn"}
                              />
                              <StatusPill
                                label={`Sasaran Dis murid (${DELIMA_KPI_MURID_PCT}%): ${delimaInsight.kpiMuridOk ? "Capai" : "Belum"}`}
                                tone={delimaInsight.kpiMuridOk ? "good" : "warn"}
                              />
                              <StatusPill
                                label={`Guru aktif (bil): ${fmtPct1(delimaInsight.avgAktiviti)}`}
                                tone="good"
                              />
                            </div>

                            <DelimaMonthAvgTable seriesDisplay={delimaInsight.seriesDisplay} />
                          </>
                        )}
                      </div>

                      <button
                        type="button"
                        onClick={(e) => {
                          e.stopPropagation();
                          if (!delimaLoading && !delimaError && delimaRows.length > 0) {
                            setDelimaModalOpen(true);
                          }
                        }}
                        disabled={delimaLoading || !!delimaError || delimaRows.length === 0}
                        className="mt-3 w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)] disabled:cursor-not-allowed disabled:opacity-40"
                      >
                        Paparan penuh
                      </button>
                    </article>
                    <article
                      className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl"
                    >
                      <h2 className="text-xl font-semibold text-white">Status DCS</h2>
                      <p className="mt-2 text-sm text-slate-300">Kemaskini terakhir: 25/03/2026, 14:30</p>

                      <div className="mt-4 flex min-h-0 flex-1 flex-col rounded-xl border border-cyan-400/15 bg-slate-950/40 p-3">
                        <DcsKpiLineChart />
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
                        {DCS_CAPAI_2025_DAERAH >= DCS_KPI_2025_KEBANGSAAN ? (
                          <p className="mt-2 text-center text-[11px] font-medium text-emerald-300/90">
                            Daerah melebihi sasaran kebangsaan
                          </p>
                        ) : null}
                      </div>

                      <div className="mt-4 flex flex-wrap gap-2">
                        <StatusPill
                          label={`TOV 2024 (Daerah): ${DCS_TOV_2024_DAERAH}%`}
                          tone="warn"
                        />
                        <StatusPill
                          label={`KPI 2025 (Kebangsaan): ${DCS_KPI_2025_KEBANGSAAN}%`}
                          tone="good"
                        />
                        <StatusPill
                          label={`Capai KPI 2025 (Daerah): ${DCS_CAPAI_2025_DAERAH}%`}
                          tone="good"
                        />
                      </div>

                      <button
                        type="button"
                        onClick={() => setIsDcsImageOpen(true)}
                        className="mt-4 w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                      >
                        Klik untuk lihat gambar penuh
                      </button>
                    </article>
                    <article
                      className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl"
                    >
                      <div className="mb-3 flex items-center justify-between">
                        <h2 className="text-xl font-semibold text-white">Program Ains</h2>
                      </div>
                      <div className="flex flex-col gap-3">
                        <div className="rounded-xl border border-cyan-400/15 bg-slate-950/40 px-2 py-2">
                          <ProgramAinsCharts
                            stats={programStats}
                            loading={programLoading}
                            error={programError}
                          />
                        </div>
                        <div className="rounded-xl border border-cyan-400/20 bg-slate-950/20 p-3">
                          <div className="flex flex-wrap gap-2">
                            <StatusPill
                              label={`Approved: ${programStats.approved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`Rejected: ${programStats.rejected}`}
                              tone={programStats.rejected > 0 ? "bad" : "warn"}
                            />
                            <StatusPill
                              label={`SK Approved: ${programStats.skApproved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`SJKC Approved: ${programStats.sjkcApproved}`}
                              tone="good"
                            />
                            <StatusPill
                              label={`SJKT Approved: ${programStats.sjktApproved}`}
                              tone="good"
                            />
                          </div>
                        </div>

                        <button
                          type="button"
                          onClick={() => setProgramModalOpen(true)}
                          className="mt-auto w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                        >
                          Lihat Senarai Pelajar (Manjung)
                        </button>
                      </div>
                    </article>
                    <article
                      className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl"
                    >
                      <h2 className="text-xl font-semibold leading-snug text-white">
                        ANALISIS PENSIJILAN DIGITAL
                      </h2>
                      <p className="mt-2 text-sm text-slate-300">
                        Ringkasan data pensijilan mengikut lokasi, jenis sekolah dan penyedia.
                      </p>
                      <PensijilanDigitalSummary />
                      <button
                        type="button"
                        onClick={() => setIsPensijilanDigitalImageOpen(true)}
                        className="mt-4 w-full rounded-xl border border-cyan-400/20 bg-slate-950/40 px-3 py-2 text-sm font-semibold text-cyan-200 hover:border-cyan-300/50 hover:shadow-[0_0_25px_rgba(0,229,255,0.12)]"
                      >
                        Klik untuk lihat gambar penuh
                      </button>
                    </article>

                    <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
                      <h2 className="text-xl font-semibold text-white">AI Tools dalam DELIMa</h2>
                      <p className="mt-2 text-sm text-slate-300">
                        Ringkasan OPTIK 2 berbanding garisan KPI kebangsaan (55%).
                      </p>
                      <AiToolsDelimaSummary />
                    </article>
                  </div>
                  <DetailsCollapseFooter />
                </div>

              </details>

              <details className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40">
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
                      Subtopik: Kertas kerja · Laporan · JNJ (Pandai Education, Goolee
                      &amp; Must Read)
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <details className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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

                  <details className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                          Canva grafik &amp; poster · Dron · Pencetak 3D · Reka bentuk 3D
                          (PDF Drive)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                        {INTEGRASI_LAPORAN_CARDS.map(({ title, viewUrl }) => (
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
                                  PDF (Google Drive)
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

                  <details className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                          Pandai Education (Google Docs) · Goolee &amp; Must Read (PDF)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
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
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>
                  <DetailsCollapseFooter />
                </div>
              </details>

              <details className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40">
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
                  <details className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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

                  <details className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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

                  <details className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-sky-400/35 bg-sky-500/10 text-sky-200">
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
                        <p className="text-sm font-semibold text-white">Laman web sekolah</p>
                        <p className="text-xs text-slate-400">
                          8 sekolah rujukan (SMK → SK → SJKC → SJKT); cari untuk sekolah lain
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      {lamanWebLoading ? (
                        <p className="text-sm text-slate-400">Memuatkan senarai sekolah...</p>
                      ) : lamanWebError ? (
                        <p className="text-sm text-rose-300">{lamanWebError}</p>
                      ) : (
                        <>
                          <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:flex-wrap sm:items-center sm:justify-between">
                            <p className="text-[11px] text-slate-500">
                              Jumlah dalam data:{" "}
                              <span className="font-semibold text-cyan-200/90">
                                {lamanWebStats.total}
                              </span>
                              {" · "}
                              ada pautan:{" "}
                              <span className="font-semibold text-cyan-200/90">
                                {lamanWebStats.withUrl}
                              </span>
                              {!lamanWebSearch.trim() ? (
                                <>
                                  {" "}
                                  · paparan utama:{" "}
                                  <span className="text-slate-400">
                                    {lamanWebVisibleRows.length}
                                  </span>
                                  /{LAMAN_WEB_FEATURED.length}
                                </>
                              ) : lamanWebVisibleRows.length !== lamanWebRows.length ? (
                                <>
                                  {" "}
                                  · hasil carian:{" "}
                                  <span className="text-slate-400">
                                    {lamanWebVisibleRows.length}
                                  </span>
                                </>
                              ) : null}
                            </p>
                            <input
                              type="search"
                              value={lamanWebSearch}
                              onChange={(e) => setLamanWebSearch(e.target.value)}
                              placeholder="Cari sekolah lain — kod, nama atau URL…"
                              className="w-full rounded-lg border border-cyan-400/25 bg-slate-950/60 px-3 py-2 text-xs text-slate-100 outline-none placeholder:text-slate-500 focus:border-cyan-400/50 sm:max-w-xs"
                              aria-label="Tapis laman web sekolah"
                            />
                          </div>
                          {lamanWebVisibleRows.length === 0 ? (
                            <p className="text-sm text-slate-400">Tiada rekod sepadan.</p>
                          ) : (
                            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                              {lamanWebVisibleRows.map((r) => (
                                <article
                                  key={`${r.code}-${r.name}`}
                                  className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                                >
                                  <p className="font-mono text-[10px] font-semibold tracking-wide text-cyan-300/90">
                                    {r.code || "—"}
                                  </p>
                                  <h3 className="mt-1 line-clamp-2 text-sm font-semibold leading-snug text-white">
                                    {r.name || "—"}
                                  </h3>
                                  {r.website ? (
                                    <a
                                      href={r.website}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="mt-2 inline-flex w-fit items-center gap-1 rounded-lg border border-cyan-400/30 bg-slate-950/50 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 transition hover:border-cyan-400/55 hover:bg-slate-900/60"
                                    >
                                      Buka laman web
                                      <span aria-hidden>→</span>
                                    </a>
                                  ) : (
                                    <p className="mt-2 text-[11px] text-slate-500">Tiada pautan</p>
                                  )}
                                </article>
                              ))}
                            </div>
                          )}
                        </>
                      )}
                    <DetailsCollapseFooter />
                    </div>
                  </details>
                  <DetailsCollapseFooter />
                </div>
              </details>

              <details className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40">
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
                      Subtopik: OPR · Program Inovasi (Video, Artsteps) · JNJ (Blink Book, Must Read)
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <details className="group mb-4 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                            d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">OPR</p>
                        <p className="text-xs text-slate-400">
                          OPR Amalan Membaca · (subtopik OPR lain akan ditambah)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      {oprMembacaLoading ? (
                        <p className="text-sm text-slate-400">
                          Memuatkan rekod OPR…
                        </p>
                      ) : oprMembacaError ? (
                        <p className="text-sm text-rose-300">{oprMembacaError}</p>
                      ) : (
                        <>
                          <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                            <article className="flex min-h-[220px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                              <p className="text-xs font-semibold uppercase tracking-wide text-slate-500">
                                OPR AMALAN MEMBACA
                              </p>
                              <p className="mt-2 text-sm font-semibold text-white">
                                OPR Amalan Membaca Manjung
                              </p>
                              <p className="mt-1 text-[11px] leading-relaxed text-slate-500">
                                Skop Manjung (Perak Drift) · zon Asia/Kuala Lumpur — sumber{" "}
                                <code className="text-cyan-300/80">
                                  opr-amalan-membaca-manjung.json
                                </code>
                              </p>
                              <div className="mt-auto space-y-2 border-t border-cyan-400/15 pt-3 text-[11px] text-slate-500">
                                <p>
                                  Jumlah rekod:{" "}
                                  <span className="font-semibold text-cyan-200/90">
                                    {oprMembacaStats.total}
                                  </span>
                                  {" · "}
                                  ada pautan:{" "}
                                  <span className="font-semibold text-cyan-200/90">
                                    {oprMembacaStats.denganPautan}
                                  </span>
                                </p>
                                {Object.keys(oprMembacaStats.byStatus).length >
                                0 ? (
                                  <p>
                                    Status:{" "}
                                    {Object.entries(
                                      oprMembacaStats.byStatus,
                                    ).map(([k, n]) => (
                                      <span key={k} className="mr-2 inline">
                                        <span className="text-slate-400">
                                          {k}
                                        </span>
                                        <span className="font-semibold text-slate-300">
                                          ({n})
                                        </span>
                                      </span>
                                    ))}
                                  </p>
                                ) : null}
                              </div>
                            </article>
                          </div>
                          <div className="mt-3 flex flex-col gap-2 sm:flex-row sm:flex-wrap sm:items-center">
                            <input
                              type="search"
                              value={oprMembacaSearch}
                              onChange={(e) => setOprMembacaSearch(e.target.value)}
                              placeholder="Cari nama program, sekolah, pelapor…"
                              className="w-full rounded-lg border border-cyan-400/25 bg-slate-950/60 px-3 py-2 text-xs text-slate-100 outline-none placeholder:text-slate-500 focus:border-cyan-400/45 sm:max-w-xs"
                              aria-label="Tapis rekod OPR Amalan Membaca"
                            />
                            <label className="flex items-center gap-2 text-[11px] text-slate-400">
                              <span className="shrink-0">Dimensi</span>
                              <select
                                value={oprMembacaDimensi}
                                onChange={(e) =>
                                  setOprMembacaDimensi(e.target.value)
                                }
                                className="rounded-lg border border-cyan-400/25 bg-slate-950/60 px-2 py-2 text-xs text-slate-100 outline-none focus:border-cyan-400/45"
                              >
                                <option value="ALL">Semua</option>
                                {oprMembacaFacetDimensi.map((d) => (
                                  <option key={d} value={d}>
                                    {d}
                                  </option>
                                ))}
                              </select>
                            </label>
                            <label className="flex items-center gap-2 text-[11px] text-slate-400">
                              <span className="shrink-0">Status OPR</span>
                              <select
                                value={oprMembacaStatus}
                                onChange={(e) =>
                                  setOprMembacaStatus(e.target.value)
                                }
                                className="rounded-lg border border-cyan-400/25 bg-slate-950/60 px-2 py-2 text-xs text-slate-100 outline-none focus:border-cyan-400/45"
                              >
                                <option value="ALL">Semua</option>
                                {oprMembacaFacetStatus.map((d) => (
                                  <option key={d} value={d}>
                                    {d}
                                  </option>
                                ))}
                              </select>
                            </label>
                          </div>
                          {oprMembacaVisible.length === 0 ? (
                            <p className="mt-3 text-sm text-slate-400">
                              Tiada rekod sepadan.
                            </p>
                          ) : (
                            <div className="mt-3 grid grid-cols-1 gap-3 sm:grid-cols-2 sm:items-stretch">
                              {oprMembacaVisible.map((r) => (
                                <article
                                  key={`opr-membaca-${r.id}-${r.oprLink}`}
                                  className="flex h-full min-h-[220px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                                >
                                  <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                                    <p className="font-mono text-[10px] font-semibold tracking-wide text-cyan-300/90">
                                      {r.nosiriOpr || "—"}
                                    </p>
                                    {r.oprLink ? (
                                      <a
                                        href={r.oprLink}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                                      >
                                        Buka OPR
                                      </a>
                                    ) : (
                                      <span className="shrink-0 rounded-lg border border-slate-600/50 px-2.5 py-1 text-[11px] text-slate-500">
                                        Tiada pautan
                                      </span>
                                    )}
                                  </div>
                                  <h4 className="line-clamp-3 text-sm font-semibold leading-snug text-white">
                                    {r.namaProgram || "—"}
                                  </h4>
                                  {r.sekolah ? (
                                    <p className="mt-1 line-clamp-2 text-[11px] leading-relaxed text-slate-400">
                                      {r.sekolah}
                                    </p>
                                  ) : null}
                                  <div className="mt-2 flex flex-wrap gap-1.5">
                                    {r.dimensi ? (
                                      <span className="rounded-md border border-slate-600/50 bg-slate-950/50 px-2 py-0.5 text-[10px] font-medium text-slate-300">
                                        {r.dimensi}
                                      </span>
                                    ) : null}
                                    {r.statusOpr ? (
                                      <span className="rounded-md border border-emerald-500/30 bg-emerald-500/10 px-2 py-0.5 text-[10px] font-medium text-emerald-200/90">
                                        {r.statusOpr}
                                      </span>
                                    ) : null}
                                  </div>
                                  <div className="mt-auto space-y-1 border-t border-cyan-400/15 pt-2">
                                    <p className="text-[11px] text-slate-500">
                                      {r.tarikhMula || "—"}
                                      {r.tarikhTamat &&
                                      r.tarikhTamat !== r.tarikhMula
                                        ? ` → ${r.tarikhTamat}`
                                        : null}
                                      {r.tahun ? ` · ${r.tahun}` : null}
                                    </p>
                                    {(r.bilGuru || r.bilMurid) && (
                                      <p className="text-[11px] text-slate-500">
                                        Guru/pegawai: {r.bilGuru || "—"} · Murid:{" "}
                                        {r.bilMurid || "—"}
                                      </p>
                                    )}
                                    {r.pegawaiPelapor ? (
                                      <p className="text-[11px] text-slate-500">
                                        Pelapor: {r.pegawaiPelapor}
                                        {r.jawatan ? ` (${r.jawatan})` : ""}
                                      </p>
                                    ) : null}
                                  </div>
                                </article>
                              ))}
                            </div>
                          )}
                        </>
                      )}
                      <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-4 rounded-2xl border border-violet-400/25 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_20px_rgba(167,139,250,0.06)] transition-[border-color,box-shadow] duration-200 open:border-violet-400/45 open:shadow-[0_0_28px_rgba(167,139,250,0.1)]">
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

                    <details className="group rounded-2xl border border-cyan-400/25 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_20px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/45 open:shadow-[0_0_28px_rgba(0,229,255,0.1)]">
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
                  <DetailsCollapseFooter />
                </div>
              </details>

              <details className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40">
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/35 bg-cyan-500/12 text-cyan-200">
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
                        d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">Bahan Sokongan</p>
                    <p className="text-xs text-slate-400">
                      Buku pengurusan, pelaporan, pencapaian, surat, penyebaran dasar
                      (MJ3PD/JKPA), bahan PDP digital, bahan rujukan
                    </p>
                  </div>
                </summary>
                <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                            d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Buku Pengurusan</p>
                        <p className="text-xs text-slate-400">
                          USTP {BOOK_PENGURUSAN_BY_YEAR[0].year}–{BOOK_PENGURUSAN_BY_YEAR[BOOK_PENGURUSAN_BY_YEAR.length - 1].year} ·{" "}
                          <span className="text-cyan-300/90">{BOOK_PENGURUSAN_AUTHOR}</span>
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        {BOOK_PENGURUSAN_BY_YEAR.map(({ year, viewUrl }) => (
                          <article
                            key={year}
                            className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                              <h3 className="text-base font-semibold text-white">
                                Buku Pengurusan USTP {year}
                              </h3>
                              <a
                                href={viewUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                              <iframe
                                loading="lazy"
                                title={`Buku Pengurusan USTP ${year}`}
                                src={`${viewUrl}?embed`}
                                width="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                                allowFullScreen
                                allow="fullscreen"
                              />
                            </div>
                          </article>
                        ))}
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                            d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Pelaporan</p>
                        <p className="text-xs text-slate-400">
                          Epelaporan {EPELAPORAN_BY_YEAR[0].year}–{EPELAPORAN_BY_YEAR[EPELAPORAN_BY_YEAR.length - 1].year} (Canva) · Takwim · Pelaporan DPD (Looker)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        {EPELAPORAN_BY_YEAR.map(({ year, viewUrl }) => (
                          <article
                            key={year}
                            className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                              <h3 className="text-base font-semibold text-white">Epelaporan {year}</h3>
                              <a
                                href={viewUrl}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                              >
                                Buka Penuh
                              </a>
                            </div>
                            <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                              <iframe
                                loading="lazy"
                                title={`Epelaporan ${year}`}
                                src={`${viewUrl}?embed`}
                                width="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                                allowFullScreen
                                allow="fullscreen"
                              />
                            </div>
                          </article>
                        ))}
                        <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-base font-semibold text-white">Pelaporan Tambahan</h3>
                            <a
                              href={TAKWIM_EMBED.replace("/embed", "")}
                              target="_blank"
                              rel="noreferrer"
                              className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] text-slate-500">Takwim / Looker Studio</p>
                          <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                            <iframe
                              loading="lazy"
                              title="Pelaporan Tambahan — Takwim Looker"
                              src={TAKWIM_EMBED}
                              width="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                              allowFullScreen
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-base font-semibold text-white">Pelaporan DPD</h3>
                            <a
                              href={PELAPORAN_DPD_EMBED.replace("/embed", "")}
                              target="_blank"
                              rel="noreferrer"
                              className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <p className="mb-2 text-[11px] text-slate-500">Looker Studio</p>
                          <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                            <iframe
                              loading="lazy"
                              title="Pelaporan DPD Looker"
                              src={PELAPORAN_DPD_EMBED}
                              width="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                              allowFullScreen
                            />
                          </div>
                        </article>
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-amber-400/30 bg-amber-500/10 text-amber-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110"
                          fill="currentColor"
                          viewBox="0 0 24 24"
                          aria-hidden
                        >
                          <path
                            fillRule="evenodd"
                            d="M10.788 3.21c.448-1.077 1.976-1.077 2.424 0l2.082 5.007 5.404.433c1.164.093 1.636 1.545.749 2.305l-4.117 3.527 1.257 5.273c.271 1.136-.964 2.033-1.96 1.425L12 18.354 7.373 21.18c-.996.608-2.231-.29-1.96-1.425l1.257-5.273-4.117-3.527c-.887-.76-.415-2.212.749-2.305l5.404-.433 2.082-5.006z"
                            clipRule="evenodd"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Pencapaian</p>
                        <p className="text-xs text-slate-400">
                          Pencapaian USTP 2025 · kad penghargaan · success story (Canva)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-base font-semibold text-white">Pencapaian USTP 2025</h3>
                            <a
                              href={PENCAPAIAN_USTP_2025_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                            <iframe
                              loading="lazy"
                              title="Pencapaian USTP 2025"
                              src={PENCAPAIAN_USTP_2025_EMBED}
                              width="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                              allowFullScreen
                              allow="fullscreen"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-base font-semibold text-white">Kad Penghargaan PPD Manjung</h3>
                            <a
                              href={KAD_PENGHARGAAN_PPD_MANJUNG_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                            <iframe
                              loading="lazy"
                              title="Kad Penghargaan PPD Manjung"
                              src={KAD_PENGHARGAAN_EMBED}
                              width="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                              allowFullScreen
                              allow="fullscreen"
                            />
                          </div>
                        </article>
                        <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
                          <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                            <h3 className="text-base font-semibold text-white">Success Story</h3>
                            <a
                              href={SUCCESS_STORY_USTP_URL}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                            >
                              Buka Penuh
                            </a>
                          </div>
                          <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                            <iframe
                              loading="lazy"
                              title="Success Story USTP"
                              src={SUCCESS_STORY_USTP_EMBED}
                              width="100%"
                              style={{ border: 0, background: "#0b1220" }}
                              className="h-[min(42vh,320px)] w-full sm:h-[min(48vh,360px)]"
                              allowFullScreen
                              allow="fullscreen"
                            />
                          </div>
                        </article>
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
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
                            d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Surat punca kuasa</p>
                        <p className="text-xs text-slate-400">
                          Pemerkasaan DELIMa (2 surat) · arahan khidmat sokongan —{" "}
                          <span className="text-amber-200/90">baharu</span>
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                        {SURAT_PUNCA_KUASA_CARDS.map(({ key, title, blurb, viewUrl }) => (
                          <article
                            key={key}
                            className="flex min-h-[360px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex flex-col gap-1">
                              <div className="flex flex-wrap items-start justify-between gap-2">
                                <h3 className="text-sm font-semibold leading-snug text-white">{title}</h3>
                                <a
                                  href={viewUrl}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                                >
                                  Buka Penuh
                                </a>
                              </div>
                              <p className="text-[11px] leading-relaxed text-slate-500">{blurb}</p>
                            </div>
                            <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                              <iframe
                                loading="lazy"
                                title={title}
                                src={driveGoogleFilePreviewUrl(viewUrl)}
                                width="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="h-[min(38vh,280px)] w-full sm:h-[min(44vh,320px)]"
                                allowFullScreen
                              />
                            </div>
                          </article>
                        ))}
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-sky-400/35 bg-sky-500/10 text-sky-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-105"
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
                        <p className="text-sm font-semibold text-white">Bahan pdp digital</p>
                        <p className="text-xs text-slate-400">
                          TikTok, YouTube USTP Manjung & Ruang Ilmu DELIMa — pratontak
                          skrin (klik kad untuk buka)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                        {BAHAN_PDP_DIGITAL_CARDS.map(
                          ({
                            key,
                            href,
                            platform,
                            title,
                            subtitle,
                            cta,
                            previewSrc,
                            previewAlt,
                            ringClass,
                            badgeClass,
                          }) => (
                            <a
                              key={key}
                              href={href}
                              target="_blank"
                              rel="noopener noreferrer"
                              className={`group/card flex flex-col overflow-hidden rounded-xl border bg-slate-900/50 shadow-[0_12px_40px_rgba(0,0,0,0.35)] backdrop-blur-xl transition duration-200 hover:-translate-y-0.5 hover:shadow-[0_16px_48px_rgba(0,229,255,0.12)] ${ringClass}`}
                            >
                              <div className="relative aspect-[16/10] overflow-hidden bg-slate-950">
                                <img
                                  src={previewSrc}
                                  alt={previewAlt}
                                  loading="lazy"
                                  decoding="async"
                                  className="h-full w-full object-cover object-top transition duration-300 group-hover/card:scale-[1.02]"
                                />
                                <div className="pointer-events-none absolute inset-0 bg-gradient-to-t from-slate-950/80 via-transparent to-transparent" />
                                <span
                                  className={`absolute left-2 top-2 rounded-md px-2 py-0.5 text-[10px] font-bold uppercase tracking-wide shadow-sm ${badgeClass}`}
                                >
                                  {platform}
                                </span>
                              </div>
                              <div className="flex flex-1 flex-col gap-1 p-3.5">
                                <h3 className="text-sm font-semibold text-white">{title}</h3>
                                <p className="text-[11px] leading-relaxed text-slate-400">{subtitle}</p>
                                <span className="mt-1 inline-flex items-center gap-1 text-xs font-semibold text-cyan-300 group-hover/card:text-cyan-200">
                                  {cta}
                                  <span aria-hidden>→</span>
                                </span>
                              </div>
                            </a>
                          ),
                        )}
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <details className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                    <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                      <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-emerald-400/30 bg-emerald-500/10 text-emerald-200">
                        <svg
                          className="h-4 w-4 transition-transform duration-200 ease-out group-open:scale-105"
                          fill="none"
                          viewBox="0 0 24 24"
                          stroke="currentColor"
                          strokeWidth={2}
                          aria-hidden
                        >
                          <path
                            strokeLinecap="round"
                            strokeLinejoin="round"
                            d="M19 20H5a2 2 0 01-2-2V6a2 2 0 012-2h10a2 2 0 012 2v1m2 13a2 2 0 01-2-2V7m2 13a2 2 0 002-2V9a2 2 0 00-2-2h-2m-4-3H9M7 16h6M7 8h6v4H7V8z"
                          />
                        </svg>
                      </span>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-white">Penyebaran dasar</p>
                        <p className="text-xs text-slate-400">
                          MJ3PD 2025 · JKPA 2026 · JKPA 2025 · JKPA 2024 (Canva — pratontak
                          & pautan penuh)
                        </p>
                      </div>
                    </summary>
                    <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
                      <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
                        {PENYEBARAN_DASAR_CANVA_CARDS.map(({ key, title, blurb, viewUrl }) => (
                          <article
                            key={key}
                            className="flex min-h-[320px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                          >
                            <div className="mb-2 flex flex-col gap-1">
                              <div className="flex flex-wrap items-start justify-between gap-2">
                                <h3 className="text-sm font-semibold leading-snug text-white">{title}</h3>
                                <a
                                  href={viewUrl}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                                >
                                  Buka Penuh
                                </a>
                              </div>
                              <p className="text-[11px] leading-relaxed text-slate-500">{blurb}</p>
                            </div>
                            <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                              <iframe
                                loading="lazy"
                                title={title}
                                src={canvaViewEmbedUrl(viewUrl)}
                                width="100%"
                                style={{ border: 0, background: "#0b1220" }}
                                className="h-[min(38vh,280px)] w-full sm:h-[min(44vh,320px)]"
                                allowFullScreen
                                allow="fullscreen"
                              />
                            </div>
                          </article>
                        ))}
                      </div>
                    <DetailsCollapseFooter />
                    </div>
                  </details>

                  <div className="mt-6 space-y-3 border-t border-cyan-400/15 pt-4">
                    <details className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                      <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                        <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-purple-400/30 bg-purple-500/10 text-purple-200">
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
                              d="M5 5a2 2 0 012-2h10a2 2 0 012 2v16l-7-3.5L5 21V5z"
                            />
                          </svg>
                        </span>
                        <div className="min-w-0 flex-1">
                          <p className="text-sm font-semibold text-white">Bahan Rujukan</p>
                          <p className="text-xs text-slate-400">
                            Kertas kerja, OPPR & lain-lain (MJ3PD/JKPA: Penyebaran dasar;
                            surat rasmi: Surat punca kuasa)
                          </p>
                        </div>
                      </summary>
                      <div className="space-y-2 overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-3">
                        <a
                          href={MAJLIS_APRESIASI_DIGITAL_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                        >
                          Kertas Kerja Majlis Apresiasi Digital
                        </a>
                        <a
                          href={KERTAS_KERJA_EDUSPARK_COE_ROADSHOW_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                        >
                          Kertas Kerja Program Eduspark COE Roadshow
                        </a>
                        <a
                          href={OPPR_PENDIDIKAN_PERAK_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                        >
                          OPPR Pendigitalan Perak
                        </a>
                        <a
                          href={ANALISIS_PENARAFAN_KENDIRI_URL}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                        >
                          Analisis Penarafan Kendiri
                        </a>
                      <DetailsCollapseFooter />
                      </div>
                    </details>
                  </div>
                  <DetailsCollapseFooter />
                </div>
              </details>
            </section>

        {selectedPegawai && (
              <div
                className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
                role="dialog"
                aria-modal="true"
              >
                <div className="w-full max-w-4xl overflow-hidden rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                  <div className="flex items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/40 p-4">
                    <div className="min-w-0">
                      <p className="truncate text-lg font-semibold text-white">{selectedPegawai.nama}</p>
                      <p className="truncate text-sm text-slate-300">{selectedPegawai.jawatan}</p>
                    </div>
                    <div className="flex items-center gap-2">
                      {selectedPegawai.detailUrl ? (
                        <a
                          href={selectedPegawai.detailUrl}
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

                  <div className="relative">
                    {selectedPegawai.detailUrl ? (
                      <iframe
                        title={`${selectedPegawai.nama} - personal profile`}
                        src={selectedPegawai.detailUrl}
                        className="h-[70vh] w-full"
                        style={{ border: 0, background: "#0b1220" }}
                        allowFullScreen
                      />
                    ) : selectedPegawai.detailImage ? (
                      <img
                        alt={`${selectedPegawai.nama} - personal profile`}
                        src={selectedPegawai.detailImage}
                        className="h-auto w-full"
                      />
                    ) : (
                      <div className="flex h-[70vh] items-center justify-center p-6 text-center text-sm text-slate-300">
                        Belum diset: beri saya pautan embed atau gambar untuk {selectedPegawai.nama}
                      </div>
                    )}
                  </div>

                  <div className="border-t border-cyan-400/10 bg-slate-950/40 p-4">
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
                    src="/assets/status-dcs.png"
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
                    src={PENSIJILAN_DIGITAL_IMAGE}
                    className="w-full h-auto rounded-xl border border-cyan-400/10 bg-black/20"
                  />
                </div>
              </div>
            </div>
          )}

          {programModalOpen && (
            <div
              className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
              role="dialog"
              aria-modal="true"
            >
              <div className="w-full max-w-5xl overflow-hidden rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                <div className="flex items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/40 p-4">
                  <div className="min-w-0">
                    <p className="truncate text-lg font-semibold text-white">
                      Program Ains - Senarai Pelajar
                    </p>
                    <p className="truncate text-sm text-slate-300">
                      Penapisan: jenis sekolah & status (tanpa email)
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setProgramModalOpen(false)}
                    className="rounded-lg border border-cyan-400/20 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                  >
                    Tutup
                  </button>
                </div>

                <div className="p-4">
                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
                    <input
                      value={programSearch}
                      onChange={(e) => setProgramSearch(e.target.value)}
                      placeholder="Cari nama / sekolah / kod..."
                      className="w-full rounded-lg border border-cyan-300/30 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none focus:border-cyan-300 sm:col-span-1"
                    />

                    <select
                      value={programModalType}
                      onChange={(e) => setProgramModalType(e.target.value)}
                      className="w-full rounded-lg border border-cyan-300/30 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none focus:border-cyan-300"
                    >
                      <option value="ALL">Semua Jenis</option>
                      <option value="SK">SK</option>
                      <option value="SJKC">SJKC</option>
                      <option value="SJKT">SJKT</option>
                    </select>

                    <select
                      value={programModalStatus}
                      onChange={(e) => setProgramModalStatus(e.target.value)}
                      className="w-full rounded-lg border border-cyan-300/30 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none focus:border-cyan-300"
                    >
                      <option value="ALL">Semua Status</option>
                      <option value="APPROVED">Approved &gt; 0</option>
                      <option value="REJECTED">Rejected &gt; 0</option>
                    </select>
                  </div>

                  <div className="mt-3 text-xs text-slate-400">
                    Jumlah rekod ditunjukkan:{" "}
                    <span className="text-cyan-200">{programModalRows.length}</span>
                  </div>
                </div>

                <div className="max-h-[55vh] overflow-auto border-t border-cyan-400/10">
                  {programLoading ? (
                    <div className="p-4 text-sm text-slate-300">
                      Memuatkan data...
                    </div>
                  ) : programError ? (
                    <div className="p-4 text-sm text-rose-200">
                      {programError}
                    </div>
                  ) : programModalRows.length === 0 ? (
                    <div className="p-4 text-sm text-slate-300">
                      Tiada data untuk penapisan ini.
                    </div>
                  ) : (
                    <table className="min-w-full text-left text-sm">
                      <thead className="sticky top-0 bg-slate-950/60">
                        <tr className="text-xs text-slate-400">
                          <th className="px-3 py-2">Nama</th>
                          <th className="px-3 py-2">Sekolah</th>
                          <th className="px-3 py-2">Kod</th>
                          <th className="px-3 py-2">Approved</th>
                        </tr>
                      </thead>
                      <tbody>
                        {programModalRows.slice(0, 600).map((r) => (
                          <tr
                            key={`${r.code}-${r.name}`}
                            className="border-t border-cyan-400/10 hover:bg-white/5"
                          >
                            <td className="px-3 py-2 align-top">
                              <div className="break-words font-medium text-white">
                                {shortenText(r.name, 28)}
                              </div>
                              <div className="mt-1 break-words text-xs text-slate-300">
                                {r.schoolLevel} / {r.schoolType}
                              </div>
                            </td>
                            <td className="px-3 py-2 align-top">
                              <div className="break-words text-slate-200">
                                {shortenText(r.school, 26)}
                              </div>
                            </td>
                            <td className="px-3 py-2 align-top text-slate-200">
                              {r.code}
                            </td>
                            <td className="px-3 py-2 align-top text-slate-200">
                              {r.approved}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>
          )}

          {delimaModalOpen && (
            <div
              className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-4"
              role="dialog"
              aria-modal="true"
            >
              <div className="max-h-[92vh] w-full max-w-6xl overflow-hidden overflow-y-auto rounded-2xl border border-cyan-400/20 bg-slate-900/85 backdrop-blur-xl">
                <div className="sticky top-0 z-10 flex flex-wrap items-center justify-between gap-3 border-b border-cyan-400/10 bg-slate-950/90 p-4 backdrop-blur-md">
                  <div className="min-w-0">
                    <p className="text-lg font-semibold text-white">
                      DELIMa — Paparan penuh (Manjung)
                    </p>
                    <p className="mt-1 text-xs text-slate-400">
                      Sumber:{" "}
                      <span className="text-slate-300">delima-ppd-manjung.xlsx</span>
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2">
                    <a
                      href={DELIMA_GOOGLE_SHEET_URL}
                      target="_blank"
                      rel="noreferrer"
                      className="rounded-lg border border-cyan-400/30 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                    >
                      Buka Sheet
                    </a>
                    <button
                      type="button"
                      onClick={() => setDelimaModalOpen(false)}
                      className="rounded-lg border border-cyan-400/20 bg-slate-950/40 px-3 py-1.5 text-xs font-semibold text-cyan-200 hover:border-cyan-300/50"
                    >
                      Tutup
                    </button>
                  </div>
                </div>

                <div className="space-y-4 p-4">
                  {!delimaLoading && !delimaError && delimaRows.length > 0 ? (
                    <>
                      <div className="grid gap-3 sm:grid-cols-2">
                        <div className="rounded-xl border border-cyan-400/15 bg-slate-950/35 p-3 text-sm">
                          <p className="text-xs text-slate-500">Kunjungan khidmat / bengkel</p>
                          <p className="mt-1 text-lg font-semibold text-white">
                            {delimaInsight.bantuSessionsTotal}{" "}
                            <span className="text-sm font-normal text-slate-400">kali</span>
                          </p>
                          <p className="mt-1 text-xs text-slate-500">
                            {delimaInsight.bantuSchoolsWithRekod} sekolah ada rekod
                          </p>
                        </div>
                        <div className="rounded-xl border border-cyan-400/15 bg-slate-950/35 p-3 text-sm">
                          <p className="text-xs text-slate-500">
                            Perbandingan purata daerah (TOV → Dis)
                          </p>
                          <div className="mt-2 flex flex-wrap gap-3 text-xs">
                            <span className="text-slate-300">
                              Guru{" "}
                              <DelimaDeltaPp
                                before={delimaInsight.avgTovGuru}
                                after={delimaInsight.avgDisGuru}
                                title="Guru"
                              />
                            </span>
                            <span className="text-slate-300">
                              Murid{" "}
                              <DelimaDeltaPp
                                before={delimaInsight.avgTovMurid}
                                after={delimaInsight.avgDisMurid}
                                title="Murid"
                              />
                            </span>
                          </div>
                        </div>
                        <div className="rounded-xl border border-cyan-400/15 bg-slate-950/35 p-3 text-sm sm:col-span-2 lg:col-span-1">
                          <p className="text-xs text-slate-500">
                            Sasaran purata Disember (guru {DELIMA_KPI_GURU_PCT}% · murid{" "}
                            {DELIMA_KPI_MURID_PCT}%)
                          </p>
                          <p className="mt-1 text-slate-200">
                            Guru:{" "}
                            <span
                              className={
                                delimaInsight.kpiGuruOk
                                  ? "font-semibold text-emerald-300"
                                  : "font-semibold text-amber-200"
                              }
                            >
                              {delimaInsight.kpiGuruOk ? "Capai" : "Belum capai"}
                            </span>
                            {" · "}
                            Murid:{" "}
                            <span
                              className={
                                delimaInsight.kpiMuridOk
                                  ? "font-semibold text-emerald-300"
                                  : "font-semibold text-amber-200"
                              }
                            >
                              {delimaInsight.kpiMuridOk ? "Capai" : "Belum capai"}
                            </span>
                          </p>
                          <p className="mt-1 text-xs text-slate-500">
                            Dis ≥ TOV: G {delimaInsight.schoolsGuruGteTov}/
                            {delimaInsight.schools} · M {delimaInsight.schoolsMuridGteTov}/
                            {delimaInsight.schools}
                          </p>
                        </div>
                      </div>

                      <div className="rounded-xl border border-cyan-400/15 bg-slate-950/25 p-3">
                        <p className="mb-2 text-xs font-medium text-slate-400">
                          Purata daerah — % aktif guru &amp; murid mengikut potongan masa
                        </p>
                        <DelimaTrendChart series={delimaInsight.seriesDisplay} height={240} />
                        <div className="mt-3 border-t border-cyan-400/10 pt-3">
                          <DelimaMonthAvgTable
                            seriesDisplay={delimaInsight.seriesDisplay}
                            maxHeightClass="max-h-[220px]"
                            showCaption={false}
                          />
                        </div>
                      </div>
                    </>
                  ) : null}

                  <div>
                    <p className="mb-2 text-sm font-medium text-slate-300">
                      Jadual data per sekolah
                    </p>
                    <input
                      value={delimaSearch}
                      onChange={(e) => setDelimaSearch(e.target.value)}
                      placeholder="Cari kod / nama sekolah..."
                      className="w-full rounded-lg border border-cyan-300/30 bg-slate-950/60 px-3 py-2 text-sm text-slate-100 outline-none focus:border-cyan-300"
                    />
                    <div className="mt-2 text-xs text-slate-400">
                      Jumlah baris:{" "}
                      <span className="text-cyan-200">{filteredDelimaRows.length}</span>
                      {filteredDelimaRows.length > 600 ? (
                        <span className="text-slate-500"> (paparan pertama 600)</span>
                      ) : null}
                    </div>
                  </div>
                </div>

                <div className="max-h-[min(50vh,480px)] overflow-auto border-t border-cyan-400/10 px-4 pb-4">
                  {delimaLoading ? (
                    <div className="p-4 text-sm text-slate-300">Memuatkan data...</div>
                  ) : delimaError ? (
                    <div className="p-4 text-sm text-rose-200">{delimaError}</div>
                  ) : filteredDelimaRows.length === 0 ? (
                    <div className="p-4 text-sm text-slate-300">Tiada rekod sepadan.</div>
                  ) : (
                    <table className="min-w-full text-left text-sm">
                      <thead className="sticky top-0 bg-slate-950/95">
                        <tr className="text-xs text-slate-400">
                          <th className="px-2 py-2">Kod</th>
                          <th className="px-2 py-2">Sekolah</th>
                          <th className="px-2 py-2">Bantu (kali)</th>
                          <th className="px-2 py-2">TOV G</th>
                          <th className="px-2 py-2">TOV M</th>
                          <th className="px-2 py-2">Nov G</th>
                          <th className="px-2 py-2">Nov M</th>
                          <th className="px-2 py-2">Dis G</th>
                          <th className="px-2 py-2">Dis M</th>
                          <th className="px-2 py-2">Guru</th>
                          <th className="px-2 py-2">Dash</th>
                          <th className="px-2 py-2">Aktif</th>
                          <th className="px-2 py-2">%</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredDelimaRows.slice(0, 600).map((r, idx) => (
                          <tr
                            key={`${r.kod}-${idx}`}
                            className="border-t border-cyan-400/10 hover:bg-white/5"
                          >
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {r.kod}
                            </td>
                            <td className="max-w-[180px] px-2 py-2 align-top text-slate-200">
                              <span className="break-words">{shortenText(r.nama, 36)}</span>
                            </td>
                            <td
                              className="max-w-[120px] px-2 py-2 align-top text-xs text-slate-300"
                              title={r.tarikhBantu || ""}
                            >
                              {r.bantuKali > 0 ? (
                                <>
                                  <span className="font-semibold tabular-nums text-cyan-200">
                                    {r.bantuKali}×
                                  </span>
                                  {r.tarikhBantu ? (
                                    <span className="mt-0.5 block break-words text-[10px] leading-snug text-slate-500">
                                      {shortenText(r.tarikhBantu, 22)}
                                    </span>
                                  ) : null}
                                </>
                              ) : (
                                "—"
                              )}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.tovGuru)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.tovMurid)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.novGuru)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.novMurid)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.disGuru)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.disMurid)}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {r.bilGuru == null ? "—" : String(Math.round(r.bilGuru))}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {r.bilDashboard == null
                                ? "—"
                                : String(Math.round(r.bilDashboard))}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {r.bilGuruAktif == null
                                ? "—"
                                : String(Math.round(r.bilGuruAktif))}
                            </td>
                            <td className="whitespace-nowrap px-2 py-2 align-top text-slate-200">
                              {fmtPct1(r.aktivitiPct)}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>
        )}
      </div>
    </main>
  );
}