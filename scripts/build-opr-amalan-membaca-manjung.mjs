/**
 * Menapis rekod OPR Amalan Membaca: hanya DAERAH = MANJUNG,
 * eksport ke public/data/opr-amalan-membaca-manjung.json (tanpa AX/AY).
 *
 * Penggunaan:
 *   node scripts/build-opr-amalan-membaca-manjung.mjs [laluan-ke-fail.xlsx]
 */
import XLSX from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "public", "data", "opr-amalan-membaca-manjung.json");

const LINK_COL = "[Document Studio] File Link #ls76z04g";

function pickSekolah(row) {
  return Object.keys(row)
    .filter((k) => k.startsWith("SILA PILIH SEKOLAH"))
    .map((k) => String(row[k] ?? "").trim())
    .filter(Boolean)
    .join(" · ");
}

function main() {
  const src =
    process.argv[2] ||
    path.join(
      process.env.USERPROFILE || "",
      "Desktop",
      "OPR AMALAN MEMBACA MANJUNG DARI PERAK DRIFT.xlsx",
    );

  if (!fs.existsSync(src)) {
    console.error("Fail tidak dijumpai:", src);
    process.exit(1);
  }

  const wb = XLSX.readFile(src);
  const sh = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sh, { defval: "", raw: false });

  const manjung = rows.filter(
    (r) => String(r.DAERAH ?? "").trim().toUpperCase() === "MANJUNG",
  );

  const items = manjung.map((r, i) => {
    const oprLink = String(r[LINK_COL] ?? "").trim();
    const tarikhTamat =
      String(r["TARIKH PROGRAM TAMAT"] ?? "").trim() ||
      String(r["TARIKH PROGRAM TAMAT_1"] ?? "").trim();

    return {
      id: i,
      timestamp: String(r.Timestamp ?? "").trim(),
      tahun: String(r.TAHUN ?? "").trim(),
      namaProgram: String(r["NAMA PROGRAM"] ?? "").trim(),
      sekolah: pickSekolah(r),
      dimensi: String(r["SILA PILIH DIMENSI"] ?? "").trim(),
      tarikhMula: String(r["TARIKH PROGRAM BERMULA"] ?? "").trim(),
      tarikhTamat,
      tempoh: String(r.TEMPOH ?? "").trim(),
      statusOpr: String(r["STATUS OPR"] ?? "").trim(),
      nosiriOpr: String(r.NOSIRIOPR ?? "").trim(),
      bilGuru: String(r["BILANGAN PEGAWAI/GURU TERLIBAT"] ?? "").trim(),
      bilMurid: String(r["BILANGAN MURID TERLIBAT"] ?? "").trim(),
      pegawaiPelapor: String(r["PEGAWAI PELAPOR"] ?? "").trim(),
      jawatan: String(r.JAWATAN ?? "").trim(),
      oprLink: oprLink.startsWith("http") ? oprLink : "",
    };
  });

  const payload = {
    meta: {
      daerah: "MANJUNG",
      sumber:
        "Borang Perak Drift — ditapis lajur DAERAH; pautan OPR daripada lajur Document Studio (setara AW).",
      bilangan: items.length,
      dijana: new Date().toISOString().slice(0, 10),
    },
    items,
  };

  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
  console.log("Tulis", items.length, "rekod →", OUT);
}

main();
