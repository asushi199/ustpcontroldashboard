/** Canva — `/view` + `?embed` untuk iframe */
export function canvaViewEmbedUrl(viewUrl) {
  return viewUrl.includes("?") ? `${viewUrl}&embed` : `${viewUrl}?embed`;
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

/** Google Docs — `/preview` untuk iframe */
export function googleDocEmbedPreviewUrl(docUrl) {
  const m = docUrl.match(/\/document\/d\/([^/?]+)/);
  return m ? `https://docs.google.com/document/d/${m[1]}/preview` : docUrl;
}

/** YouTube — URL tontonan / live → embed */
export function youtubeWatchToEmbedUrl(url) {
  const s = String(url ?? "");
  const m = s.match(
    /(?:youtube\.com\/(?:watch\?v=|live\/|embed\/)|youtu\.be\/)([a-zA-Z0-9_-]{11})/,
  );
  return m ? `https://www.youtube.com/embed/${m[1]}` : s;
}
