/** Canva — `/view` + `?embed` untuk iframe (`/watch` → `/view`). Hash (#page) diletak selepas query. */
export function canvaViewEmbedUrl(viewUrl) {
  let u = String(viewUrl ?? "").trim();
  if (!u) return u;
  u = u.replace(/\/watch\/?(\?|#|$)/i, "/view$1");
  const hashIdx = u.indexOf("#");
  const base = hashIdx >= 0 ? u.slice(0, hashIdx) : u;
  const hash = hashIdx >= 0 ? u.slice(hashIdx) : "";
  const withEmbed = base.includes("?") ? `${base}&embed` : `${base}?embed`;
  return `${withEmbed}${hash}`;
}

/**
 * Carta / PKG Maklumat Asas: URL dari Google Sheet.
 * - Pautan Drive (view atau /file/d/…) → iframe `/preview` (boleh pratontoh tanpa URL imej terus).
 * - Google Docs → iframe preview.
 * - Lain: anggap imej terus (laluan `/assets/...` atau sambungan .png dll).
 *
 * @param {string} url
 * @returns {{ kind: "iframe" | "img", src: string }}
 */
export function maklumatAsasPreviewMode(url) {
  const raw = String(url ?? "").trim();
  if (!raw) return { kind: "img", src: raw };
  if (/drive\.google\.com/i.test(raw)) {
    return { kind: "iframe", src: driveGoogleFilePreviewUrl(raw) };
  }
  if (/docs\.google\.com\/document\//i.test(raw)) {
    return { kind: "iframe", src: googleDocEmbedPreviewUrl(raw) };
  }
  if (
    raw.startsWith("/") ||
    /\.(png|jpe?g|gif|webp|svg)(\?|#|$)/i.test(raw)
  ) {
    return { kind: "img", src: raw };
  }
  return { kind: "img", src: raw };
}

/** Google Drive fail — URL paparan ke `/preview` untuk iframe */
export function driveGoogleFilePreviewUrl(viewUrl) {
  const s = String(viewUrl ?? "");
  const fileD = s.match(/\/file\/d\/([^/?]+)/);
  if (fileD) return `https://drive.google.com/file/d/${fileD[1]}/preview`;
  if (s.includes("drive.google.com")) {
    const openId = s.match(/[?&]id=([^&]+)/);
    if (openId) return `https://drive.google.com/file/d/${openId[1]}/preview`;
  }
  return s;
}

/**
 * Google Drive — pautan kongsi fail → URL imej terus (lh3) untuk dipapar sebagai `<img>`.
 * Untuk gambar pegawai yang sememangnya imej, supaya papar bersih tanpa bingkai Drive.
 * Pulangkan URL asal jika bukan pautan Drive / gagal cari ID.
 * @param {string} url
 * @returns {string}
 */
export function driveGoogleImageUrl(url) {
  const s = String(url ?? "").trim();
  if (!/drive\.google\.com/i.test(s)) return s;
  const fileD = s.match(/\/file\/d\/([^/?]+)/);
  if (fileD) return `https://lh3.googleusercontent.com/d/${fileD[1]}`;
  const idParam = s.match(/[?&]id=([^&]+)/);
  if (idParam) return `https://lh3.googleusercontent.com/d/${idParam[1]}`;
  return s;
}

/** Google Docs — `/preview` untuk iframe */
export function googleDocEmbedPreviewUrl(docUrl) {
  const m = docUrl.match(/\/document\/d\/([^/?]+)/);
  return m ? `https://docs.google.com/document/d/${m[1]}/preview` : docUrl;
}

/**
 * Looker Studio — URL `/reporting/` tidak boleh dalam iframe; perlu `/embed/reporting/`.
 * @param {string} url
 */
export function lookerStudioEmbedUrl(url) {
  const s = String(url ?? "").trim();
  if (!s) return s;
  if (
    /lookerstudio\.google\.com\/reporting\//i.test(s) &&
    !/\/embed\/reporting\//i.test(s)
  ) {
    return s.replace(
      /lookerstudio\.google\.com\/reporting\//i,
      "lookerstudio.google.com/embed/reporting/",
    );
  }
  return s;
}

/** YouTube — URL tontonan / live → embed */
export function youtubeWatchToEmbedUrl(url) {
  const s = String(url ?? "");
  const m = s.match(
    /(?:youtube\.com\/(?:watch\?v=|live\/|embed\/)|youtu\.be\/)([a-zA-Z0-9_-]{11})/,
  );
  return m ? `https://www.youtube.com/embed/${m[1]}` : s;
}
