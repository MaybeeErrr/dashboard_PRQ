"use client";

export type DashboardPeriodFilterMode = "current-month" | "monthly" | "date-range" | "ytd";

export type DashboardPeriodFilterState = {
  endDate: string;
  endMonth: string;
  mode: DashboardPeriodFilterMode;
  startDate: string;
  startMonth: string;
};

export type DashboardPeriodOption = {
  label: string;
  periodKey: string;
};

type DashboardPeriodFilterProps = {
  className?: string;
  onChange: (value: DashboardPeriodFilterState) => void;
  options: DashboardPeriodOption[];
  theme?: "hsi" | "wms" | "lts";
  validationMessage?: string | null;
  value: DashboardPeriodFilterState;
};

const themeClasses = {
  hsi: {
    focus: "focus:border-blue-400 focus:ring-blue-100 hover:border-blue-300",
    text: "text-[#071426]",
  },
  lts: {
    focus: "focus:border-violet-400 focus:ring-violet-100 hover:border-violet-300",
    text: "text-[#24103f]",
  },
  wms: {
    focus: "focus:border-amber-400 focus:ring-amber-100 hover:border-amber-300",
    text: "text-[#2a1a10]",
  },
} as const;

function inputClass(theme: keyof typeof themeClasses) {
  const selected = themeClasses[theme];
  return `h-9 rounded-md border border-[#d8e1ec] bg-white px-3 text-xs font-semibold ${selected.text} shadow-sm outline-none transition focus:ring-4 ${selected.focus}`;
}

export function defaultPeriodFilter(periodKey: string): DashboardPeriodFilterState {
  const year = periodKey.slice(0, 4);
  const month = periodKey.slice(4, 6);
  const startDate = `${year}-${month}-01`;
  const endDate = `${year}-${month}-${String(new Date(Number(year), Number(month), 0).getDate()).padStart(2, "0")}`;
  return {
    endDate,
    endMonth: periodKey,
    mode: "current-month",
    startDate,
    startMonth: periodKey,
  };
}

export function DashboardPeriodFilter({
  className = "",
  onChange,
  options,
  theme = "hsi",
  validationMessage,
  value,
}: DashboardPeriodFilterProps) {
  const fieldClass = inputClass(theme);
  const sortedOptions = options.length > 0 ? options : [{ label: value.startMonth, periodKey: value.startMonth }];
  const setValue = (patch: Partial<DashboardPeriodFilterState>) => onChange({ ...value, ...patch });

  return (
    <div className={`grid gap-1 text-[11px] font-semibold uppercase text-slate-500 ${className}`}>
      <span>PERIODE</span>
      <div className="flex flex-wrap items-center gap-2">
        <select className={`${fieldClass} w-[145px]`} value={value.mode} onChange={(event) => setValue({ mode: event.target.value as DashboardPeriodFilterMode })}>
          <option value="current-month">Current Month</option>
          <option value="monthly">Month Range</option>
          <option value="date-range">Date Range</option>
          <option value="ytd">YTD</option>
        </select>
        {value.mode === "date-range" ? (
          <>
            <input
              aria-label="Start date"
              className={`${fieldClass} w-[145px]`}
              type="date"
              value={value.startDate}
              onChange={(event) => setValue({ startDate: event.target.value })}
            />
            <input
              aria-label="End date"
              className={`${fieldClass} w-[145px]`}
              type="date"
              value={value.endDate}
              onChange={(event) => setValue({ endDate: event.target.value })}
            />
          </>
        ) : (
          <>
            <select className={`${fieldClass} w-[170px]`} value={value.startMonth} onChange={(event) => setValue({ startMonth: event.target.value })}>
              {sortedOptions.map((option) => (
                <option key={option.periodKey} value={option.periodKey}>
                  {option.label}
                </option>
              ))}
            </select>
            {value.mode === "monthly" && (
              <select className={`${fieldClass} w-[170px]`} value={value.endMonth} onChange={(event) => setValue({ endMonth: event.target.value })}>
                {sortedOptions.map((option) => (
                  <option key={option.periodKey} value={option.periodKey}>
                    {option.label}
                  </option>
                ))}
              </select>
            )}
          </>
        )}
      </div>
      {validationMessage && <span className="max-w-[460px] text-[10px] font-semibold normal-case text-rose-600">{validationMessage}</span>}
    </div>
  );
}
