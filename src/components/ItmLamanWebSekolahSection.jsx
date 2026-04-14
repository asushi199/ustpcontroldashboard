/**
 * @param {{
 *   loading: boolean,
 *   error: string,
 *   stats: { total: number, withUrl: number },
 *   search: string,
 *   onSearchChange: (e: import("react").ChangeEvent<HTMLInputElement>) => void,
 *   visibleRows: { code: string, name: string, website: string }[],
 *   featuredCount: number,
 *   mergedTotal: number,
 * }} props
 */
export function ItmLamanWebSekolahSection({
  loading,
  error,
  stats,
  search,
  onSearchChange,
  visibleRows,
  featuredCount,
  mergedTotal,
}) {
  return (
    <details
      name="osc-sub-itm-laman-web"
      className="group mb-6 rounded-2xl border border-cyan-400/20 bg-gradient-to-br from-slate-950/55 via-slate-950/35 to-slate-900/25 shadow-[0_0_24px_rgba(0,229,255,0.06)] transition-[border-color,box-shadow] duration-200 open:border-cyan-400/40 open:shadow-[0_0_32px_rgba(0,229,255,0.12)]"
    >
      <summary className="flex cursor-pointer list-none items-center gap-3 px-4 py-3.5 transition-colors hover:bg-white/[0.04] [&::-webkit-details-marker]:hidden">
        <span className="flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-sky-400/35 bg-sky-500/10 text-sky-200">
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
              d="M21 12a9 9 0 01-9 9m9-9a9 9 0 00-9-9m9 9H3m9 9a9 9 0 01-9-9m9 9c1.657 0 3-4.03 3-9s-1.343-9-3-9m0 18c-1.657 0-3-4.03-3-9s1.343-9 3-9m-9 9a9 9 0 019-9"
            />
          </svg>
        </span>
        <div className="min-w-0 flex-1">
          <p className="text-sm font-semibold text-white">Laman web sekolah</p>
          <p className="text-xs text-slate-400">
            8 sekolah rujukan (SMK → SK → SJKC → SJKT); cari untuk sekolah lain · pautan boleh dikemas kini
            melalui tab Google Sheet (jika ditetapkan)
          </p>
        </div>
      </summary>
      <div className="overflow-hidden rounded-b-2xl border-t border-cyan-400/10 bg-slate-950/30 px-4 py-4">
        {loading ? (
          <p className="text-sm text-slate-400">Memuatkan senarai sekolah...</p>
        ) : error ? (
          <p className="text-sm text-rose-300">{error}</p>
        ) : (
          <>
            <div className="mb-3 flex flex-col gap-2 sm:flex-row sm:flex-wrap sm:items-center sm:justify-between">
              <p className="text-[11px] text-slate-500">
                Jumlah dalam data:{" "}
                <span className="font-semibold text-cyan-200/90">{stats.total}</span>
                {" · "}
                ada pautan:{" "}
                <span className="font-semibold text-cyan-200/90">{stats.withUrl}</span>
                {!search.trim() ? (
                  <>
                    {" "}
                    · paparan utama:{" "}
                    <span className="text-slate-400">{visibleRows.length}</span>/{featuredCount}
                  </>
                ) : visibleRows.length !== mergedTotal ? (
                  <>
                    {" "}
                    · hasil carian:{" "}
                    <span className="text-slate-400">{visibleRows.length}</span>
                  </>
                ) : null}
              </p>
              <input
                type="search"
                value={search}
                onChange={onSearchChange}
                placeholder="Cari sekolah lain — kod, nama atau URL…"
                className="w-full rounded-lg border border-cyan-400/25 bg-slate-950/60 px-3 py-2 text-xs text-slate-100 outline-none placeholder:text-slate-500 focus:border-cyan-400/50 sm:max-w-xs"
                aria-label="Tapis laman web sekolah"
              />
            </div>
            {visibleRows.length === 0 ? (
              <p className="text-sm text-slate-400">Tiada rekod sepadan.</p>
            ) : (
              <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                {visibleRows.map((r) => (
                  <article
                    key={`${r.code}-${r.name}`}
                    className="flex flex-col rounded-xl border border-cyan-400/20 bg-slate-900/40 p-3 backdrop-blur-xl"
                  >
                    <p className="font-mono text-[10px] font-semibold tracking-wide text-cyan-300/90">
                      {r.code || "—"}
                    </p>
                    <h3 className="mt-1 line-clamp-2 text-sm font-semibold leading-snug text-white">
                      {r.name || "—"}
                    </h3>
                    {r.website ? (
                      <a
                        href={r.website}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="mt-2 inline-flex w-fit items-center gap-1 rounded-lg border border-cyan-400/30 bg-slate-950/50 px-2.5 py-1.5 text-[11px] font-semibold text-cyan-200 transition hover:border-cyan-400/55 hover:bg-slate-900/60"
                      >
                        Buka laman web
                        <span aria-hidden>→</span>
                      </a>
                    ) : (
                      <p className="mt-2 text-[11px] text-slate-500">Tiada pautan</p>
                    )}
                  </article>
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </details>
  );
}
