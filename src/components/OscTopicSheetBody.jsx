import { useEffect, useState } from "react";
import { DetailsCollapseFooter } from "./DetailsCollapseFooter.jsx";
import { OscCardGridFromRows } from "./OscSheetCardGrid.jsx";
import { OscSubtopikSummaryIcon } from "./OscSubtopikSummaryIcon.jsx";
import { loadOscTopicTabRows } from "../lib/oscTopicTabSheetFetch.js";
import { groupOscRowsBySubtopik } from "../lib/oscSheetSubtopik.js";

/**
 * @param {{ groups: { subtopikKey: string, subtopikTitle: string, subtopikBlurb?: string, subtopikIcon?: string, cards: import("../lib/oscSheetCsv.js").OscSheetCardRow[] }[] }} props
 */
function OscSheetSubtopikDetailsList({ groups, detailsName }) {
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
          <details key={g.subtopikKey} name={detailsName} className={wrapClass}>
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

/**
 * OSC topik — satu tab Sheet (`gid`) atau CSV templat (`demoPath` bila tiada VITE_GOOGLE_SHEET_ID).
 * Lajur: subtopik_* (pilihan: subtopik_icon = emoji atau kod seperti video, doc), sort, key, title, url, type, blurb (tiada `section`).
 *
 * @param {{ sheetGid: string | undefined, demoPath?: string, detailsName: string, fallback: import("react").ReactNode, loadingLabel?: string, omitTrailingPageFooter?: boolean }} props
 */
export function OscTopicSheetBody({
  sheetGid,
  demoPath,
  detailsName,
  fallback,
  loadingLabel,
  omitTrailingPageFooter = false,
}) {
  const id = import.meta.env.VITE_GOOGLE_SHEET_ID;
  const gid = String(sheetGid ?? "").trim();
  const demo = String(demoPath ?? "").trim().replace(/^\//, "");
  const loadFromSheet = Boolean(id && gid);
  const loadFromDemo = Boolean(demo) && !id;
  const shouldFetch = loadFromSheet || loadFromDemo;
  const [groups, setGroups] = useState(null);

  useEffect(() => {
    if (!shouldFetch) return undefined;
    let cancelled = false;
    loadOscTopicTabRows({
      gid: loadFromSheet ? gid : "",
      demoPath: loadFromDemo ? demo : undefined,
    })
      .then((rows) => {
        if (cancelled) return;
        const g = groupOscRowsBySubtopik(rows, {}).filter((x) => x.cards.length > 0);
        setGroups(g);
      })
      .catch(() => {
        if (!cancelled) setGroups([]);
      });
    return () => {
      cancelled = true;
    };
  }, [id, gid, demo, shouldFetch, loadFromSheet, loadFromDemo]);

  const loadMsg =
    loadingLabel ?? "Memuatkan daripada Google Sheet…";

  const pageFooter = omitTrailingPageFooter ? null : <DetailsCollapseFooter />;

  if (!shouldFetch) {
    return (
      <>
        {fallback}
        {pageFooter}
      </>
    );
  }

  if (groups === null) {
    return (
      <>
        <p className="mb-4 text-sm text-slate-400">{loadMsg}</p>
        {pageFooter}
      </>
    );
  }

  if (groups.length > 0) {
    return (
      <>
        <OscSheetSubtopikDetailsList groups={groups} detailsName={detailsName} />
        {pageFooter}
      </>
    );
  }

  return (
    <>
      {fallback}
      {pageFooter}
    </>
  );
}
