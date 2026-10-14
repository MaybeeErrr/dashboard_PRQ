export type CompactRankingInsightItem = {
  achievement: number | null | undefined;
  id: string;
  label: string;
};

export function CompactRankingInsight({
  formatPercent,
  isAllScope,
  periodLabel,
  rows,
  scopeLabel,
  selectedId,
  valueClassName = "text-emerald-700",
}: {
  formatPercent: (value: number | null | undefined) => string;
  isAllScope: boolean;
  periodLabel: string;
  rows: CompactRankingInsightItem[];
  scopeLabel: string;
  selectedId: string;
  valueClassName?: string;
}) {
  const rankedRows = [...rows].sort((a, b) => {
    const achievementDelta = (b.achievement ?? -1) - (a.achievement ?? -1);
    if (achievementDelta !== 0) return achievementDelta;
    return a.label.localeCompare(b.label);
  });
  const selectedIndex = rankedRows.findIndex((row) => row.id === selectedId);
  const selectedRow = selectedIndex >= 0 ? rankedRows[selectedIndex] : null;
  const topRow = rankedRows[0] ?? null;
  const insightRow = isAllScope ? topRow : selectedRow;
  const title = isAllScope ? `TOP ${periodLabel}` : selectedIndex >= 0 ? `#${selectedIndex + 1} RANKING ${periodLabel}` : `RANKING ${periodLabel}`;
  const detail = isAllScope ? insightRow?.label : `dari ${rankedRows.length} ${scopeLabel}`;

  if (!insightRow) {
    return (
      <div className="mt-2.5 border-t border-slate-200/80 pt-2.5">
        <p className="text-[10px] font-bold uppercase tracking-[0.08em] text-slate-400">{title}</p>
        <p className="mt-1 text-xs font-semibold uppercase text-slate-500">Data belum tersedia</p>
      </div>
    );
  }

  return (
    <div className="mt-2.5 border-t border-slate-200/80 pt-2.5">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0">
          <p className="text-[10px] font-bold uppercase tracking-[0.08em] text-slate-400">{title}</p>
          <p className="mt-1 truncate text-xs font-semibold uppercase text-slate-800">{detail}</p>
        </div>
        <span className={`shrink-0 text-xs font-bold ${valueClassName}`}>{formatPercent(insightRow.achievement)}</span>
      </div>
    </div>
  );
}
