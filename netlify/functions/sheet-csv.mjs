/**
 * Proksi CSV Google Sheets (elakkan CORS di pelayar).
 * Format `handler` — serasi runtime Node Netlify (bukan `export default` Request/Response sahaja).
 * Query: id = SPREADSHEET_ID, gid = tab
 */
export const handler = async (event) => {
  const qs = event.queryStringParameters || {};
  const sheetId = String(qs.id ?? "").trim();
  const gid = String(qs.gid ?? "0").trim() || "0";

  if (!sheetId || !/^[a-zA-Z0-9_-]+$/.test(sheetId)) {
    return {
      statusCode: 400,
      headers: { "Content-Type": "text/plain; charset=utf-8" },
      body: "Missing or invalid id",
    };
  }

  const upstream = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${encodeURIComponent(gid)}`;

  const r = await fetch(upstream, {
    headers: { "User-Agent": "USTP-OSC-Netlify-Function/1.0" },
  });

  if (!r.ok) {
    const t = await r.text();
    return {
      statusCode: 502,
      headers: { "Content-Type": "text/plain; charset=utf-8" },
      body: t || "Upstream error",
    };
  }

  const text = await r.text();
  return {
    statusCode: 200,
    headers: {
      "Content-Type": "text/csv; charset=utf-8",
      "Access-Control-Allow-Origin": "*",
      "Cache-Control": "public, max-age=60",
    },
    body: text,
  };
};
