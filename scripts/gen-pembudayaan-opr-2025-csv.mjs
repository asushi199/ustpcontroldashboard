/**
 * Jana CSV kad OPR 2025 daripada public/data/opr-amalan-membaca-manjung.json
 *   node scripts/gen-pembudayaan-opr-2025-csv.mjs
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, "..");
const JSON_PATH = path.join(ROOT, "public", "data", "opr-amalan-membaca-manjung.json");
const OUT = path.join(ROOT, "public", "data", "pembudayaan-opr-2025-salin-ke-sheet.csv");

function esc(s) {
  const t = String(s ?? "");
  if (/[",\n\r]/.test(t)) return `"${t.replace(/"/g, '""')}"`;
  return t;
}

const j = JSON.parse(fs.readFileSync(JSON_PATH, "utf8"));
const header =
  "subtopik_key,subtopik_title,subtopik_sort,subtopik_blurb,subtopik_icon,sort,key,title,url,type,blurb,preview_url";
const lines = [header];

for (const it of j.items) {
  const blurb = [
    it.nosiriOpr,
    it.statusOpr,
    it.dimensi,
    it.tarikhMula === it.tarikhTamat
      ? it.tarikhMula
      : `${it.tarikhMula} → ${it.tarikhTamat}`,
    it.sekolah,
  ]
    .filter(Boolean)
    .join(" · ");
  lines.push(
    [
      "opr-2025",
      "OPR 2025",
      1,
      "Kad OPR Amalan Membaca Manjung (2025) — satu program satu pautan Drive",
      "doc",
      it.id + 1,
      `opr-manjung-2025-${String(it.id).padStart(2, "0")}`,
      it.namaProgram,
      it.oprLink || "",
      "pdf",
      blurb,
      "",
    ]
      .map(esc)
      .join(","),
  );
}

fs.writeFileSync(OUT, lines.join("\n"), "utf8");
console.log(`Wrote ${j.items.length} rows → ${path.relative(ROOT, OUT)}`);
