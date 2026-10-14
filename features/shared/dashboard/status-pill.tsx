import type React from "react";

export type DashboardStatusTone = "neutral" | "good" | "warn" | "bad";

export function DashboardStatusPill({
  children,
  tone = "neutral",
}: {
  children: React.ReactNode;
  tone?: DashboardStatusTone;
}) {
  const toneClass =
    tone === "good"
      ? "border-emerald-200/80 bg-emerald-50/90 text-emerald-700"
      : tone === "warn"
        ? "border-amber-200/80 bg-amber-50/90 text-amber-800"
        : tone === "bad"
          ? "border-rose-200/80 bg-rose-50/90 text-rose-700"
          : "border-slate-200/80 bg-white/80 text-slate-600";

  return (
    <span className={`inline-flex items-center rounded-full border px-3 py-1 text-xs font-semibold shadow-sm backdrop-blur ${toneClass}`}>
      {children}
    </span>
  );
}
