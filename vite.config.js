import { defineConfig, loadEnv } from "vite";
import react from "@vitejs/plugin-react";

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), "");
  const sheetId = env.VITE_GOOGLE_SHEET_ID;
  const gid = env.VITE_GOOGLE_SHEET_GID || "0";

  return {
    plugins: [react()],
    server: sheetId
      ? {
          proxy: {
            "/api/sheet-csv": {
              target: `https://docs.google.com/spreadsheets/d/${sheetId}`,
              changeOrigin: true,
              rewrite: () => `/export?format=csv&gid=${encodeURIComponent(gid)}`,
            },
          },
        }
      : {},
  };
});
