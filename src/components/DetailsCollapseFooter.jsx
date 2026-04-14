export function DetailsCollapseFooter() {
  return (
    <div className="flex justify-center border-t border-cyan-400/10 bg-slate-950/35 px-4 py-3">
      <button
        type="button"
        className="text-xs font-semibold text-cyan-300/90 underline-offset-2 hover:text-cyan-200 hover:underline"
        onClick={(e) => {
          const d = e.currentTarget.closest("details");
          if (d) d.removeAttribute("open");
        }}
      >
        Tutup bahagian
      </button>
    </div>
  );
}
