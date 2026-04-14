import { useEffect, useMemo, useState } from "react";
import {
  canvaViewEmbedUrl,
  driveGoogleFilePreviewUrl,
  googleDocEmbedPreviewUrl,
  lookerStudioEmbedUrl,
  youtubeWatchToEmbedUrl,
} from "../lib/embedUrls.js";
import { filterOscRowsBySection, loadOscSheetRows } from "../lib/oscSheetFetch.js";

/**
 * @typedef {{ key: string, title: string, blurb?: string, viewUrl: string, type?: string, previewUrl?: string }} FallbackCard
 * @typedef {{ key: string, title: string, blurb?: string, url: string, type?: string, previewUrl?: string }} OscCardGridRow
 */

function isLikelyDirectImageUrl(raw) {
  const u = String(raw ?? "").trim();
  if (!u) return false;
  if (u.startsWith("/")) return true;
  return /\.(png|jpe?g|gif|webp|svg|avif|bmp)(\?|#|$)/i.test(u);
}

/**
 * Pratonton kad: guna `preview_url` jika ada, jika tidak guna `url`. Pautan "Buka Penuh" kekal `url`.
 *
 * @param {OscCardGridRow} row
 * @returns {{ kind: "iframe" | "img" | "placeholder", src?: string, allow?: string, referrerPolicy?: string, message?: string }}
 */
export function resolveOscCardEmbed(row) {
  const t = row.type ?? "pdf";
  const linkUrl = row.url;
  const previewBase = String(row.previewUrl ?? "").trim();
  const u = previewBase || linkUrl;

  const driveReferrer = "strict-origin-when-cross-origin";

  if (t === "gdoc") {
    return {
      kind: "iframe",
      src: googleDocEmbedPreviewUrl(u),
      referrerPolicy: driveReferrer,
    };
  }
  if (t === "canva") {
    return { kind: "iframe", src: canvaViewEmbedUrl(u) };
  }
  if (t === "embed") {
    const src = /lookerstudio\.google\.com/i.test(u) ? lookerStudioEmbedUrl(u) : u;
    return { kind: "iframe", src };
  }
  if (t === "youtube") {
    return {
      kind: "iframe",
      src: youtubeWatchToEmbedUrl(u),
      allow:
        "accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture; web-share",
    };
  }
  if (t === "image") {
    if (/drive\.google\.com/i.test(u)) {
      return {
        kind: "iframe",
        src: driveGoogleFilePreviewUrl(u),
        referrerPolicy: driveReferrer,
      };
    }
    if (!isLikelyDirectImageUrl(u)) {
      return {
        kind: "placeholder",
        message: previewBase
          ? "Nilai preview_url bukan imej terus (.png, .jpg, …) atau pautan Drive — semak URL."
          : "URL utama ialah halaman web (contoh Artsteps), bukan fail imej. Tambah lajur preview_url dengan pautan imej / Drive, atau tukar type kepada embed / canva.",
      };
    }
    const base = import.meta.env.BASE_URL || "/";
    const src = u.startsWith("/")
      ? `${base.endsWith("/") ? base : `${base}/`}${u.slice(1)}`
      : u;
    return { kind: "img", src };
  }
  return {
    kind: "iframe",
    src: driveGoogleFilePreviewUrl(u),
    referrerPolicy: driveReferrer,
  };
}

const OSC_CARD_IMG_PREVIEW_CLASS =
  "h-[min(38vh,280px)] w-full object-contain object-top sm:h-[min(44vh,320px)]";

const OSC_CARD_IFRAME_BOX_CLASS =
  "h-[min(38vh,280px)] w-full sm:h-[min(44vh,320px)]";

/**
 * Pratonton kad: imej boleh klik → `row.url`. Pratonton iframe Google Drive pula dilitupi pautan ke `row.url`.
 *
 * @param {{ row: OscCardGridRow, emb: { kind: string, src?: string, allow?: string, referrerPolicy?: string, message?: string } }} props
 */
function OscCardMediaPreview({ row, emb }) {
  const linkHref = String(row.url ?? "").trim();

  if (emb.kind === "placeholder") {
    return (
      <div className="flex h-[min(38vh,280px)] flex-col items-center justify-center gap-2 bg-slate-950/80 px-3 text-center sm:h-[min(44vh,320px)]">
        <p className="text-[11px] leading-relaxed text-slate-400">{emb.message}</p>
        <p className="text-[10px] text-slate-600">
          Video Drive: kongsi &quot;Sesiapa yang mempunyai pautan&quot; boleh lihat.
        </p>
      </div>
    );
  }

  if (emb.kind === "img") {
    const img = (
      <img
        src={emb.src}
        alt={row.title}
        loading="lazy"
        decoding="async"
        className={OSC_CARD_IMG_PREVIEW_CLASS}
      />
    );
    if (!linkHref) return img;
    return (
      <a
        href={linkHref}
        target="_blank"
        rel="noopener noreferrer"
        className="block cursor-pointer focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-cyan-400/60"
      >
        {img}
      </a>
    );
  }

  const iframeEl = (
    <iframe
      loading="lazy"
      title={row.title}
      src={emb.src}
      width="100%"
      style={{ border: 0, background: "#0b1220" }}
      className={OSC_CARD_IFRAME_BOX_CLASS}
      allowFullScreen
      allow={emb.allow}
      referrerPolicy={
        emb.referrerPolicy ??
        (row.type === "youtube" ? "strict-origin-when-cross-origin" : undefined)
      }
    />
  );

  const isDrivePreview =
    /drive\.google\.com/i.test(String(emb.src ?? ""));

  if (linkHref && isDrivePreview) {
    return (
      <div className={`relative ${OSC_CARD_IFRAME_BOX_CLASS}`}>
        <iframe
          loading="lazy"
          title={row.title}
          src={emb.src}
          width="100%"
          style={{ border: 0, background: "#0b1220" }}
          className="pointer-events-none absolute inset-0 h-full w-full"
          allowFullScreen
          allow={emb.allow}
          referrerPolicy={
            emb.referrerPolicy ??
            (row.type === "youtube"
              ? "strict-origin-when-cross-origin"
              : undefined)
          }
        />
        <a
          href={linkHref}
          target="_blank"
          rel="noopener noreferrer"
          className="absolute inset-0 z-10 block cursor-pointer focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-cyan-400/60"
          aria-label={`Buka pautan utama — ${row.title}`}
        />
      </div>
    );
  }

  return iframeEl;
}

/**
 * @param {{ rows: OscCardGridRow[], className?: string }} props
 */
export function OscCardGridFromRows({ rows, className = "" }) {
  return (
    <div className={className}>
      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        {rows.map((row) => {
          const emb = resolveOscCardEmbed(row);
          return (
            <article
              key={row.key}
              className="flex min-h-[360px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
            >
              <div className="mb-2 flex flex-col gap-1">
                <div className="flex flex-wrap items-start justify-between gap-2">
                  <h3 className="text-sm font-semibold leading-snug text-white">{row.title}</h3>
                  <a
                    href={row.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                  >
                    Buka Penuh
                  </a>
                </div>
                {row.blurb ? (
                  <p className="text-[11px] leading-relaxed text-slate-500">{row.blurb}</p>
                ) : null}
              </div>
              <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                <OscCardMediaPreview row={row} emb={emb} />
              </div>
            </article>
          );
        })}
      </div>
    </div>
  );
}

/**
 * Kad grid — data daripada Google Sheet (lajur section) atau sandaran statik.
 *
 * @param {{ section: string, fallbackCards?: FallbackCard[], className?: string }} props
 */
export function OscSheetCardGrid({ section, fallbackCards = [], className = "" }) {
  const sheetEnabled = Boolean(import.meta.env.VITE_GOOGLE_SHEET_ID);
  const [loading, setLoading] = useState(sheetEnabled);
  const [error, setError] = useState(null);
  const [sheetRows, setSheetRows] = useState([]);

  useEffect(() => {
    if (!sheetEnabled) return;
    let cancelled = false;
    loadOscSheetRows()
      .then((all) => {
        if (cancelled) return;
        setSheetRows(filterOscRowsBySection(all, section));
      })
      .catch((e) => {
        if (cancelled) return;
        setError(e instanceof Error ? e.message : "Ralat muat data");
        setSheetRows([]);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [section, sheetEnabled]);

  const rows = useMemo(() => {
    if (sheetRows.length > 0) {
      return sheetRows.map((r) => ({
        key: r.key,
        title: r.title,
        blurb: r.blurb,
        url: r.url,
        type: r.type,
        previewUrl: r.previewUrl,
      }));
    }
    return fallbackCards.map((c) => ({
      key: c.key,
      title: c.title,
      blurb: c.blurb,
      url: c.viewUrl,
      type: c.type ?? "pdf",
      previewUrl: c.previewUrl,
    }));
  }, [sheetRows, fallbackCards]);

  const showSheetEmptyHint =
    sheetEnabled && !loading && !error && sheetRows.length === 0 && fallbackCards.length > 0;

  return (
    <div className={className}>
      {loading ? (
        <p className="mb-3 text-sm text-slate-400">Memuatkan data Google Sheet…</p>
      ) : null}
      {error ? (
        <p className="mb-3 text-sm text-rose-300">
          {error} — memaparkan sandaran jika ada.
        </p>
      ) : null}
      {showSheetEmptyHint ? (
        <p className="mb-3 text-[11px] text-amber-200/90">
          Tiada baris untuk bahagian <span className="font-mono">{section}</span> dalam sheet —
          memaparkan sandaran.
        </p>
      ) : null}

      <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
        {rows.map((row) => {
          const emb = resolveOscCardEmbed(row);
          return (
            <article
              key={row.key}
              className="flex min-h-[360px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
            >
              <div className="mb-2 flex flex-col gap-1">
                <div className="flex flex-wrap items-start justify-between gap-2">
                  <h3 className="text-sm font-semibold leading-snug text-white">{row.title}</h3>
                  <a
                    href={row.url}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="shrink-0 rounded-lg bg-gradient-to-r from-cyan-400 to-indigo-500 px-2.5 py-1 text-[11px] font-semibold text-slate-950"
                  >
                    Buka Penuh
                  </a>
                </div>
                {row.blurb ? (
                  <p className="text-[11px] leading-relaxed text-slate-500">{row.blurb}</p>
                ) : null}
              </div>
              <div className="min-h-0 flex-1 overflow-hidden rounded-lg border border-cyan-400/15">
                <OscCardMediaPreview row={row} emb={emb} />
              </div>
            </article>
          );
        })}
      </div>
    </div>
  );
}
