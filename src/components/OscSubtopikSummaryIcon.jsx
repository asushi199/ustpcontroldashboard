/**
 * Ikon ringkas subtopik (Sheet: lajur subtopik_icon).
 * - Kosong → buku lalai
 * - Kod ASCII huruf kecil (contoh: video, doc, chart) → SVG pratetap
 * - Lain → papar sebagai emoji/teks (disyorkan: satu emoji dalam sel)
 */
const WRAP =
  "flex h-9 w-9 shrink-0 items-center justify-center rounded-xl border border-violet-400/25 bg-violet-500/10 text-violet-200";

const SVG_CLS =
  "h-4 w-4 transition-transform duration-200 ease-out group-open:scale-110";

function Svg({ children }) {
  return (
    <svg
      className={SVG_CLS}
      fill="none"
      viewBox="0 0 24 24"
      stroke="currentColor"
      strokeWidth={2}
      aria-hidden
    >
      {children}
    </svg>
  );
}

/** @type {Record<string, () => import("react").ReactNode>} */
const PRESETS = {
  book: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M12 6.253v13m0-13C10.832 5.477 9.246 5 7.5 5S4.168 5.477 3 6.253v13C4.168 18.477 5.754 18 7.5 18s3.332.477 4.5 1.253m0-13C13.168 5.477 14.754 5 16.5 5c1.747 0 3.332.477 4.5 1.253v13C19.832 18.477 18.247 18 16.5 18c-1.746 0-3.332.477-4.5 1.253"
      />
    </Svg>
  ),
  doc: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
      />
    </Svg>
  ),
  document: () => PRESETS.doc(),
  pdf: () => PRESETS.doc(),
  video: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M15 10l4.553-2.276A1 1 0 0121 8.618v6.764a1 1 0 01-1.447.894L15 14M5 18h8a2 2 0 002-2V8a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"
      />
    </Svg>
  ),
  chart: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M9 19v-6m0 0V5m0 12h6m-6 0H9m6 0v-6m0 6V5m0 12h6m0-6v6"
      />
    </Svg>
  ),
  megaphone: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M11 5.882V19.24a1.76 1.76 0 01-3.417.592l-2.147-6.15M18 13a3 3 0 100-6M5.436 13.683A4.001 4.001 0 017 6h1.018c.26.604.817 1.06 1.491 1.234M5.436 13.683A4.001 4.001 0 0117 6h1.018c.26.604.817 1.06 1.491 1.234M5.436 13.683A4.001 4.001 0 0012 17h.01"
      />
    </Svg>
  ),
  photo: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"
      />
    </Svg>
  ),
  image: () => PRESETS.photo(),
  link: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"
      />
    </Svg>
  ),
  folder: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M3 7v10a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-6l-2-2H5a2 2 0 00-2 2z"
      />
    </Svg>
  ),
  calendar: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"
      />
    </Svg>
  ),
  users: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M12 4.354a4 4 0 110 5.292M15 21H3v-1a6 6 0 0112 0v1zm0 0h6v-1a6 6 0 00-9-5.197M13 7a4 4 0 11-8 0 4 4 0 018 0z"
      />
    </Svg>
  ),
  star: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M11.049 2.927c.3-.921 1.603-.921 1.902 0l1.519 4.674a1 1 0 00.95.69h4.915c.969 0 1.371 1.24.588 1.81l-3.976 2.888a1 1 0 00-.363 1.118l1.518 4.674c.3.922-.755 1.688-1.538 1.118l-3.976-2.888a1 1 0 00-1.176 0l-3.976 2.888c-.783.57-1.838-.197-1.538-1.118l1.518-4.674a1 1 0 00-.363-1.118l-3.976-2.888c-.784-.57-.38-1.81.588-1.81h4.914a1 1 0 00.951-.69l1.519-4.674z"
      />
    </Svg>
  ),
  sparkles: () => PRESETS.star(),
  music: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M9 19V6l12-3v13M9 19c0 1.105-1.343 2-3 2s-3-.895-3-2 1.343-2 3-2 3 .895 3 2zm12-3c0 1.105-1.343 2-3 2s-3-.895-3-2 1.343-2 3-2 3 .895 3 2zM9 10l12-3"
      />
    </Svg>
  ),
  globe: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M3.055 11H5a2 2 0 012 2v1a2 2 0 002 2 2 2 0 012 2v2.945M8 3.935V5.5A2.5 2.5 0 0010.5 8h.5a2 2 0 012 2 2 2 0 104 0 2 2 0 012-2h1.064M15 20.488V18a2 2 0 012-2h3.064M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
      />
    </Svg>
  ),
  clipboard: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-3 7h3m-3 4h3m-6-4h.01M9 16h.01"
      />
    </Svg>
  ),
  lightbulb: () => (
    <Svg>
      <path
        strokeLinecap="round"
        strokeLinejoin="round"
        d="M9.663 17h4.673M12 3v1m6.364 1.636l-.707.707M21 12h-1M4 12H3m3.343-5.657l-.707-.707m2.828 9.9a5 5 0 117.072 0l-.548.547A3.374 3.374 0 0014 18.469V19a2 2 0 11-4 0v-.531c0-.895-.356-1.754-.988-2.386l-.548-.547z"
      />
    </Svg>
  ),
};

const SLUG = /^[a-z][a-z0-9_-]{0,31}$/;

/**
 * @param {{ icon?: string | null }} props
 */
export function OscSubtopikSummaryIcon({ icon }) {
  const raw = String(icon ?? "").trim();
  if (!raw) {
    return <span className={WRAP}>{PRESETS.book()}</span>;
  }
  const slug = raw.toLowerCase();
  if (SLUG.test(slug) && PRESETS[slug]) {
    return <span className={WRAP}>{PRESETS[slug]()}</span>;
  }
  return (
    <span className={WRAP} title={raw}>
      <span
        className="max-w-[2.25rem] truncate text-center text-[1.2rem] leading-none"
        aria-hidden
      >
        {raw}
      </span>
    </span>
  );
}
