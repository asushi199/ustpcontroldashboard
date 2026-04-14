import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  /** Hanya `VITE_*` — elak baca keseluruhan process.env */
  const env = loadEnv(mode, process.cwd(), "VITE_");
  const defaultSheetId = (env.VITE_GOOGLE_SHEET_ID || "").trim();
  const defaultGid = (env.VITE_GOOGLE_SHEET_GID || "0").trim();

  return {
    plugins: [react()],
    /**
     * Sentiasa daftar proksi — jika `server` kosong kerana .env tidak dibaca semasa config,
     * `/api/sheet-csv` akan 404 dan aplikasi senyap jatuh ke demo CSV.
     * ID/gid utama tetap dari query (pelayar); lalai hanya sandaran.
     */
    server: {
      proxy: {
        "/api/sheet-csv": {
          target: "https://docs.google.com",
          changeOrigin: true,
          rewrite: (path) => {
            const q = path.includes("?") ? path.slice(path.indexOf("?")) : "";
            const sp = new URLSearchParams(q.startsWith("?") ? q.slice(1) : q);
            const id = (sp.get("id") || defaultSheetId).trim();
            const gid = (sp.get("gid") ?? defaultGid).trim() || "0";
            if (!id) return path;
            return `/spreadsheets/d/${id}/export?format=csv&gid=${encodeURIComponent(gid)}`;
          },
        },
      },
    },
  };
});
