import { useEffect, useState } from "react";
import { loadBahanSokonganSheetRows } from "../lib/bahanSokonganSheetFetch.js";
import { groupBahanSokonganSheetRows } from "../lib/oscSheetSubtopik.js";
import {
  BOOK_PENGURUSAN_AUTHOR,
  BOOK_PENGURUSAN_BY_YEAR,
  CONTOH_BAHAN_DELIMA_FALLBACK,
  EPELAPORAN_BY_YEAR,
  PENYEBARAN_DASAR_CANVA_CARDS,
  PENYEBARAN_DASAR_DRIVE_PDF_CARDS,
  PELAPORAN_DPD_EMBED,
  PENCAPAIAN_USTP_2025_EMBED,
  PENCAPAIAN_USTP_2025_URL,
  TAKWIM_EMBED,
  BAHAN_PDP_DIGITAL_CARDS,
  SURAT_PUNCA_KUASA_CARDS,
  KAD_PENGHARGAAN_EMBED,
  KAD_PENGHARGAAN_PPD_MANJUNG_URL,
  SUCCESS_STORY_USTP_EMBED,
  SUCCESS_STORY_USTP_URL,
  MAJLIS_APRESIASI_DIGITAL_EMBED,
  MAJLIS_APRESIASI_DIGITAL_URL,
  SLAID_MAJLIS_APRESIASI_EMBED,
  SLAID_MAJLIS_APRESIASI_URL,
} from "../lib/bahanSokonganConstants.js";
import { canvaViewEmbedUrl, driveGoogleFilePreviewUrl } from "../lib/embedUrls.js";
import { DetailsCollapseFooter } from "./DetailsCollapseFooter.jsx";
import { OscCardGridFromRows, OscSheetCardGrid } from "./OscSheetCardGrid.jsx";
import { OscSubtopikSummaryIcon } from "./OscSubtopikSummaryIcon.jsx";

function BahanSokonganStaticInner() {
  return (
    <>
      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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
              USTP {BOOK_PENGURUSAN_BY_YEAR[0].year}–
              {BOOK_PENGURUSAN_BY_YEAR[BOOK_PENGURUSAN_BY_YEAR.length - 1].year} ·{" "}
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
                  <h3 className="text-base font-semibold text-white">Buku Pengurusan USTP {year}</h3>
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

      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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
              Epelaporan {EPELAPORAN_BY_YEAR[0].year}–{EPELAPORAN_BY_YEAR[EPELAPORAN_BY_YEAR.length - 1].year}{" "}
              (Canva) · Takwim · Pelaporan DPD (Looker)
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

      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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
              Kertas kerja Majlis Apresiasi Digital · Slaid Majlis Apresiasi (Canva) · Pencapaian USTP
              2025 · kad penghargaan · success story (Canva)
            </p>
          </div>
        </summary>
        <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
          <div className="grid grid-cols-1 gap-4 lg:grid-cols-2">
            <article className="flex min-h-[380px] flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl">
              <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
                <h3 className="text-base font-semibold text-white">Kertas Kerja Majlis Apresiasi Digital</h3>
                <a
                  href={MAJLIS_APRESIASI_DIGITAL_URL}
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
                  title="Kertas Kerja Majlis Apresiasi Digital"
                  src={MAJLIS_APRESIASI_DIGITAL_EMBED}
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
                <h3 className="text-base font-semibold text-white">Slaid Majlis Apresiasi</h3>
                <a
                  href={SLAID_MAJLIS_APRESIASI_URL}
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
                  title="Slaid Majlis Apresiasi"
                  src={SLAID_MAJLIS_APRESIASI_EMBED}
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

      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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

      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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
              TikTok, YouTube USTP Manjung & Ruang Ilmu DELIMa — pratontak skrin (klik kad untuk buka)
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

      <details
        name="osc-sub-bahan"
        className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
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
              MJ3PD 2025 · JKPA 2026 · JKPA 2025 · JKPA 2024 (Canva) · contoh KPM (hebahan COE) ·
              pameran (Hari Guru 2025, R.E.A.Digital, Townhall Ipoh 2025 — PDF Drive)
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
            {PENYEBARAN_DASAR_DRIVE_PDF_CARDS.map(({ key, title, blurb, viewUrl }) => (
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

      <details
        name="osc-sub-bahan"
        className="group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
      >
        <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
          <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-indigo-400/35 bg-indigo-500/10 text-indigo-200">
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
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
          </span>
          <div className="min-w-0 flex-1">
            <p className="text-sm font-semibold text-white">Contoh Bahan Delima</p>
            <p className="text-xs text-slate-400">
              Jadual penggunaan DELIMA / DELIMa, contoh GC, OPR — PDF sekolah rujukan
            </p>
          </div>
        </summary>
        <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
          <OscSheetCardGrid
            key="contoh-bahan-delima"
            section="contoh-bahan-delima"
            fallbackCards={CONTOH_BAHAN_DELIMA_FALLBACK}
          />
          <DetailsCollapseFooter />
        </div>
      </details>
    </>
  );
}

/** @param {{ groups: { subtopikKey: string, subtopikTitle: string, subtopikBlurb?: string, subtopikIcon?: string, cards: import("../lib/oscSheetCsv.js").OscSheetCardRow[] }[] }} props */
function BahanSokonganDynamicInner({ groups }) {
  const n = groups.length;
  return (
    <>
      {groups.map((g, i) => {
        const isLast = i === n - 1;
        const wrapClass = isLast
          ? "group rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
          : "group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]";
        const rows = g.cards.map((r) => ({
          key: r.key,
          title: r.title,
          blurb: r.blurb,
          url: r.url,
          type: r.type,
          previewUrl: r.previewUrl,
        }));
        return (
          <details key={g.subtopikKey} name="osc-sub-bahan" className={wrapClass}>
            <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
              <OscSubtopikSummaryIcon icon={g.subtopikIcon} />
              <div className="min-w-0 flex-1">
                <p className="text-sm font-semibold text-white">{g.subtopikTitle}</p>
                {g.subtopikBlurb ? (
                  <p className="text-xs text-slate-400">{g.subtopikBlurb}</p>
                ) : null}
              </div>
            </summary>
            <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
              <OscCardGridFromRows rows={rows} />
              <DetailsCollapseFooter />
            </div>
          </details>
        );
      })}
    </>
  );
}

export function BahanSokonganPageSection() {
  const sheetEnabled = Boolean(
    import.meta.env.VITE_GOOGLE_SHEET_ID &&
      String(import.meta.env.VITE_BAHAN_SOKONGAN_GID ?? "").trim() !== "",
  );
  const [dynamicGroups, setDynamicGroups] = useState(null);

  useEffect(() => {
    if (!sheetEnabled) return undefined;
    let cancelled = false;
    loadBahanSokonganSheetRows()
      .then((rows) => {
        if (cancelled) return;
        const g = groupBahanSokonganSheetRows(rows).filter((x) => x.cards.length > 0);
        setDynamicGroups(g);
      })
      .catch(() => {
        if (cancelled) return;
        setDynamicGroups([]);
      });
    return () => {
      cancelled = true;
    };
  }, [sheetEnabled]);

  const loading = sheetEnabled && dynamicGroups === null;
  const useDynamic = sheetEnabled && Array.isArray(dynamicGroups) && dynamicGroups.length > 0;

  return (
    <details
      name="osc-page-topik"
      className="group rounded-2xl border border-cyan-400/20 bg-slate-900/28 backdrop-blur-2xl open:border-cyan-400/40"
    >
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
            Buku pengurusan, pelaporan, pencapaian, surat, penyebaran dasar (MJ3PD/JKPA), bahan PDP
            digital · Contoh Bahan Delima
          </p>
        </div>
      </summary>
      <div className="rounded-b-2xl border-t border-cyan-400/15 p-4">
        {loading ? (
          <p className="mb-4 text-sm text-slate-400">Memuatkan Bahan Sokongan daripada Google Sheet…</p>
        ) : null}
        {useDynamic ? <BahanSokonganDynamicInner groups={dynamicGroups} /> : null}
        {!loading && !useDynamic ? <BahanSokonganStaticInner /> : null}
        <DetailsCollapseFooter />
      </div>
    </details>
  );
}
