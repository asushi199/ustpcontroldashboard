/**
 * @param {Array<{ page?: string, subtopikKey?: string, subtopikTitle?: string, subtopikSort?: number, subtopikBlurb?: string, subtopikIcon?: string, sort: number, title: string }>} rows
 * @param {{ pageFilter?: string | null }} opts
 */
export function groupOscRowsBySubtopik(rows, opts = {}) {
  const pageFilter =
    opts.pageFilter != null && String(opts.pageFilter).trim() !== ""
      ? String(opts.pageFilter).trim()
      : null;

  const filtered = pageFilter
    ? rows.filter((r) => String(r.page ?? "").trim() === pageFilter)
    : rows;

  const byKey = new Map();
  for (const r of filtered) {
    const key = String(r.subtopikKey ?? "").trim() || "__default";
    const title = String(r.subtopikTitle ?? "").trim() || key;
    let ss = 999;
    const rawSs = r.subtopikSort;
    if (rawSs != null && Number.isFinite(Number(rawSs))) ss = Number(rawSs);
    const blurb = r.subtopikBlurb != null ? String(r.subtopikBlurb).trim() : "";

    if (!byKey.has(key)) {
      byKey.set(key, {
        subtopikKey: key,
        subtopikTitle: title,
        subtopikSort: ss,
        subtopikBlurb: blurb || undefined,
        subtopikIcon: undefined,
        cards: [],
      });
    }
    const g = byKey.get(key);
    const iconCell =
      r.subtopikIcon != null ? String(r.subtopikIcon).trim() : "";
    if (iconCell && !g.subtopikIcon) g.subtopikIcon = iconCell;
    g.cards.push(r);
  }

  for (const g of byKey.values()) {
    g.cards.sort(
      (a, b) => a.sort - b.sort || a.title.localeCompare(b.title, "ms"),
    );
  }

  return [...byKey.values()].sort(
    (a, b) =>
      a.subtopikSort - b.subtopikSort ||
      a.subtopikTitle.localeCompare(b.subtopikTitle, "ms"),
  );
}

/**
 * Jika mana-mana baris ada `page` berisi teks, tapis kepada `bahan` (master sheet).
 * Tab khusus tanpa lajur `page` — jangan tapis.
 *
 * @param {Array<{ page?: string }>} rows
 */
export function groupBahanSokonganSheetRows(rows) {
  const anyPage = rows.some((r) => String(r.page ?? "").trim() !== "");
  const pageFilter = anyPage ? "bahan" : null;
  return groupOscRowsBySubtopik(rows, { pageFilter });
}
