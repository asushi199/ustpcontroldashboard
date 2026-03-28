import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

/** Logo rasmi PPD (PNG telus) — watermark latar tiga mod reka */
const USTP_WATERMARK_SRC = "/assets/ustp-ppd-manjung-watermark.png";

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

const EPelaporan_2024_EMBED = `${EPelaporan_2024_LINK}?embed`;
const EPelaporan_2025_EMBED = `${EPelaporan_2025_LINK}?embed`;

// Buku pengurusan (Canva embed)
const BOOK_PENGURUSAN_EMBED =
  "https://www.canva.com/design/DAGZuhOHzIM/kiiF86Fawa-ChTVvuH7DaQ/view?embed";
const BOOK_PENGURUSAN_LINK =
  "https://www.canva.com/design/DAGZuhOHzIM/kiiF86Fawa-ChTVvuH7DaQ/view?utm_content=DAGZuhOHzIM&utm_campaign=designshare&utm_medium=embeds&utm_source=link";
const BOOK_PENGURUSAN_AUTHOR = "Rujhan Alwi";

// Bahan sokongan (Canva /view 一般用于“打开全文”，避免 iframe 预览被拦截)
const JKPA_PGB_2025_URL =
  "https://www.canva.com/design/DAG-vuv7xWw/ZSoPv8yyr_7gMStIqeAgjA/view";

/** Google Drive — PDF surat pemerkasaan DELIMa */
const SURAT_PEMERKASAAN_DELIMA_2023_2024_URL =
  "https://drive.google.com/file/d/1lyQd8IOz4lmgAFo9Lz2ZK9bKDZQNPy4F/view?usp=sharing";
const SURAT_PEMERKASAAN_DELIMA_2025_2026_URL =
  "https://drive.google.com/file/d/196YTk8wLq0qvdfSRVIFLxeQ_wLhTjjC-/view?usp=sharing";

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
const KAD_PENGHARGAAN_PPD_MANJUNG_URL =
  "https://www.canva.com/design/DAGa60sQAjg/bmiRnetB9hWhfU20s6eOkQ/watch";
const MAJLIS_APRESIASI_DIGITAL_URL =
  "https://www.canva.com/design/DAG3I0mucKY/ehUSnXhpz06UxXWmTN8-wA/edit?utm_content=DAG3I0mucKY&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton";
const KERTAS_KERJA_EDUSPARK_COE_ROADSHOW_URL =
  "https://www.canva.com/design/DAG0hWYthb4/edr15Oq3_X1l5QgqseiQXg/edit?utm_content=DAG0hWYthb4&utm_campaign=designshare&utm_medium=link2&utm_source=sharebutton";

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
      } catch (e) {
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
      } catch (e) {
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
                  Command Dashboard
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

        {/* Baris 1: Takwim USTP · Data DELIMa · Status DCS */}
        <section className="grid grid-cols-1 gap-6 lg:grid-cols-2 xl:grid-cols-3">
          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
              <h2 className="text-xl font-semibold text-white">Takwim USTP</h2>
              <a
                href={USTP_CALENDAR_EMBED}
                target="_blank"
                rel="noopener noreferrer"
                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                loading="lazy"
                title="Takwim USTP — Google Calendar"
                src={USTP_CALENDAR_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
              />
            </div>
          </article>

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

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
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
        </section>

        {/* Baris 2: Program Ains · Epelaporan 2024 · Epelaporan 2025 */}
        <section className="grid grid-cols-1 gap-6 lg:grid-cols-2 xl:grid-cols-3">
          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
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

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between gap-3">
              <h2 className="text-xl font-semibold text-white">Epelaporan 2024</h2>
              <a
                href={EPelaporan_2024_LINK}
                target="_blank"
                rel="noreferrer"
                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                title="Epelaporan 2024"
                src={EPelaporan_2024_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
              />
            </div>
          </article>

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between gap-3">
              <h2 className="text-xl font-semibold text-white">Epelaporan 2025</h2>
              <a
                href={EPelaporan_2025_LINK}
                target="_blank"
                rel="noreferrer"
                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                title="Epelaporan 2025"
                src={EPelaporan_2025_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
              />
            </div>
          </article>
        </section>

        {/* Baris 3: Pelaporan Tambahan · Pelaporan DPD · Pencapaian USTP 2025 */}
        <section className="grid grid-cols-1 gap-6 lg:grid-cols-2 xl:grid-cols-3">
          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-xl font-semibold text-white">Pelaporan Tambahan</h2>
              <a
                href={TAKWIM_EMBED.replace("/embed", "")}
                target="_blank"
                rel="noreferrer"
                className="rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                title="Takwim Looker"
                src={TAKWIM_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
              />
            </div>
          </article>

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between">
              <h2 className="text-xl font-semibold text-white">Pelaporan DPD</h2>
              <a
                href={PELAPORAN_DPD_EMBED.replace("/embed", "")}
                target="_blank"
                rel="noreferrer"
                className="rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                title="Pelaporan DPD Looker"
                src={PELAPORAN_DPD_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
              />
            </div>
          </article>

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between gap-3">
              <h2 className="text-xl font-semibold text-white">Pencapaian USTP 2025</h2>
              <a
                href={PENCAPAIAN_USTP_2025_URL}
                target="_blank"
                rel="noopener noreferrer"
                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                loading="lazy"
                title="Pencapaian USTP 2025"
                src={PENCAPAIAN_USTP_2025_EMBED}
                width="100%"
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
                allow="fullscreen"
              />
            </div>
          </article>
        </section>

        {/* Baris akhir: Analisis Pensijilan Digital · Buku Pengurusan · Maklumat Pegawai */}
        <section className="grid grid-cols-1 gap-6 lg:grid-cols-2 xl:grid-cols-3">
          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
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

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-4 backdrop-blur-2xl">
            <div className="mb-3 flex items-center justify-between gap-3">
              <div className="min-w-0">
                <h2 className="text-xl font-semibold text-white">
                  Buku Pengurusan USTP 2025
                </h2>
                <p className="mt-1 text-sm text-slate-300">
                  Penulis:{" "}
                  <span className="text-cyan-200">{BOOK_PENGURUSAN_AUTHOR}</span>
                </p>
              </div>
              <a
                href={BOOK_PENGURUSAN_LINK}
                target="_blank"
                rel="noopener noreferrer"
                className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-3 py-1.5 text-xs font-semibold text-slate-950"
              >
                Buka Penuh
              </a>
            </div>
            <div className="flex-1 overflow-hidden rounded-xl border border-cyan-400/20">
              <iframe
                loading="lazy"
                title="Buku Pengurusan USTP 2025 - Canva embed"
                src={BOOK_PENGURUSAN_EMBED}
                style={{ border: 0, background: "#0b1220" }}
                className="h-full w-full"
                allowFullScreen
                allow="fullscreen"
              />
            </div>
          </article>

          <article className="flex min-h-[520px] flex-col rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
            <h2 className="text-xl font-semibold text-white">Maklumat Pegawai</h2>
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

            {/* Modal: tampilkan halaman / poster peribadi */}
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
          </article>

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
                      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
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
        </section>

        {/* Bahan sokongan - always at the bottom */}
        <section className="mt-6">
          <article className="rounded-2xl border border-cyan-400/20 bg-slate-900/28 p-5 backdrop-blur-2xl">
            <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
              <div className="min-w-0">
                <h3 className="text-xl font-semibold text-white">Bahan Sokongan</h3>
                <p className="mt-2 text-sm text-slate-300">
                  Video di bawah; semua bahan lain dalam Bahan Rujukan. Klik baris untuk buka senarai.
                </p>
              </div>
            </div>

            <div className="mt-4 space-y-3">
              <details className="group overflow-hidden rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/25 bg-cyan-400/10 text-cyan-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:rotate-90"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9 5l7 7-7 7"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">Video</p>
                    <p className="text-xs text-slate-400">TikTok & YouTube USTP Manjung</p>
                  </div>
                </summary>
                <div className="space-y-2 border-t border-cyan-400/10 bg-slate-950/30 px-4 py-3">
                  <a
                    href={TIKTOK_USTP_MANJUNG_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    TikTok USTP Manjung
                  </a>
                  <a
                    href={YOUTUBE_USTP_MANJUNG_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    YouTube USTP Manjung
                  </a>
                </div>
              </details>

              <details className="group overflow-hidden rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]">
                <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
                  <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-cyan-400/25 bg-cyan-400/10 text-cyan-200">
                    <svg
                      className="h-4 w-4 transition-transform duration-200 ease-out group-open:rotate-90"
                      fill="none"
                      viewBox="0 0 24 24"
                      stroke="currentColor"
                      strokeWidth={2}
                      aria-hidden
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        d="M9 5l7 7-7 7"
                      />
                    </svg>
                  </span>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-white">Bahan Rujukan</p>
                    <p className="text-xs text-slate-400">
                      Kertas kerja, surat DELIMa, JKPA, OPPR & lain-lain
                    </p>
                  </div>
                </summary>
                <div className="space-y-2 border-t border-cyan-400/10 bg-slate-950/30 px-4 py-3">
                  <a
                    href={EPelaporan_2022_LINK}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Epelaporan 2022
                  </a>
                  <a
                    href={EPelaporan_2023_LINK}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Epelaporan 2023
                  </a>
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
                    href={JKPA_PGB_2025_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    JKPA PGB 2025
                  </a>
                  <a
                    href={SURAT_PEMERKASAAN_DELIMA_2023_2024_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Surat Pemerkasaan DELIMa 2023–2024
                  </a>
                  <a
                    href={SURAT_PEMERKASAAN_DELIMA_2025_2026_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Surat Pemerkasaan DELIMa 2025–2026
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
                    href={KAD_PENGHARGAAN_PPD_MANJUNG_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Kad Penghargaan PPD Manjung
                  </a>
                  <a
                    href={ANALISIS_PENARAFAN_KENDIRI_URL}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block rounded-xl border border-cyan-400/15 bg-slate-950/45 px-3 py-2.5 text-xs font-medium text-cyan-100 transition hover:border-cyan-400/45 hover:bg-slate-900/55 hover:shadow-[0_0_18px_rgba(0,229,255,0.08)]"
                  >
                    Analisis Penarafan Kendiri
                  </a>
                </div>
              </details>
            </div>
          </article>
        </section>
      </div>
    </main>
  );
}