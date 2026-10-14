import type React from "react";

export type DashboardMetricTone = "neutral" | "good" | "warn";

export type DashboardMetricCardTheme = {
  accentLineClass: string;
  detailToneClass: Record<DashboardMetricTone, string>;
  hoverClass: string;
  iconToneClass: Record<DashboardMetricTone, string>;
  labelClass: string;
  panelClass: string;
  toneClass: Record<DashboardMetricTone, string>;
  valueClass: string;
};

export function DashboardMetricCard({
  detail,
  icon: Icon,
  label,
  theme,
  tone = "neutral",
  value,
}: {
  detail: string;
  icon?: React.ComponentType<{ className?: string }>;
  label: string;
  theme: DashboardMetricCardTheme;
  tone?: DashboardMetricTone;
  value: string;
}) {
  return (
    <section
      className={`${theme.panelClass} group relative min-h-[92px] overflow-hidden bg-gradient-to-br ${theme.toneClass[tone]} p-3 transition ${theme.hoverClass}`}
    >
      <span className={`absolute inset-x-0 top-0 h-0.5 ${theme.accentLineClass}`} />
      <div className="flex items-start justify-between gap-2">
        <p className={`max-w-[10rem] text-[10px] font-semibold uppercase leading-3.5 ${theme.labelClass}`}>{label}</p>
        {Icon && (
          <span className={`flex h-6 w-6 shrink-0 items-center justify-center rounded-md border ${theme.iconToneClass[tone]}`}>
            <Icon className="h-3 w-3" />
          </span>
        )}
      </div>
      <p className={`mt-2 truncate text-xl font-semibold leading-none tracking-normal ${theme.valueClass}`}>{value}</p>
      <p className={`mt-1.5 truncate text-[10px] uppercase leading-4 ${theme.detailToneClass[tone]}`}>{detail}</p>
    </section>
  );
}
