export const INDONESIAN_MONTHS = [
  "JANUARI",
  "FEBRUARI",
  "MARET",
  "APRIL",
  "MEI",
  "JUNI",
  "JULI",
  "AGUSTUS",
  "SEPTEMBER",
  "OKTOBER",
  "NOVEMBER",
  "DESEMBER",
];

export function periodKeyFromDate(date: Date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}`;
}

export function monthYearFromPeriodKey(periodKey: string) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  return {
    monthName: INDONESIAN_MONTHS[month - 1] ?? periodKey,
    year: Number.isFinite(year) ? year : new Date().getFullYear(),
  };
}

export function dateInputForPeriodKey(periodKey: string, day: "first" | "last") {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  const date = day === "first" ? 1 : new Date(year, month, 0).getDate();
  return `${year}-${String(month).padStart(2, "0")}-${String(date).padStart(2, "0")}`;
}

export function monthRangeKeys(startPeriodKey: string, endPeriodKey: string) {
  const startYear = Number(startPeriodKey.slice(0, 4));
  const startMonth = Number(startPeriodKey.slice(4, 6));
  const endYear = Number(endPeriodKey.slice(0, 4));
  const endMonth = Number(endPeriodKey.slice(4, 6));
  if (!Number.isFinite(startYear) || !Number.isFinite(startMonth) || !Number.isFinite(endYear) || !Number.isFinite(endMonth)) return [];
  const startIndex = startYear * 12 + startMonth - 1;
  const endIndex = endYear * 12 + endMonth - 1;
  if (startIndex > endIndex) return [];
  return Array.from({ length: endIndex - startIndex + 1 }, (_, index) => {
    const current = startIndex + index;
    const year = Math.floor(current / 12);
    const month = (current % 12) + 1;
    return `${year}${String(month).padStart(2, "0")}`;
  });
}

export type DashboardPeriodPart = {
  endDay: number;
  periodKey: string;
  startDay: number;
  weight: number;
};

export type DashboardPeriodSemanticMode =
  | "MONTH"
  | "MONTH_RANGE"
  | "DATE_RANGE"
  | "YTD"
  | "current-month"
  | "monthly"
  | "date-range"
  | "ytd";

export function getPeriodSemanticLabel(mode: DashboardPeriodSemanticMode) {
  if (mode === "YTD" || mode === "ytd") return "YTD";
  if (mode === "MONTH_RANGE" || mode === "DATE_RANGE" || mode === "monthly" || mode === "date-range") return "PERIODE";
  return "MTD";
}

export function dateRangeToMonthParts(startDate: string, endDate: string): DashboardPeriodPart[] {
  const start = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) return [];
  const startKey = periodKeyFromDate(start);
  const endKey = periodKeyFromDate(end);
  return monthRangeKeys(startKey, endKey).map((periodKey) => {
    const year = Number(periodKey.slice(0, 4));
    const month = Number(periodKey.slice(4, 6));
    const daysInMonth = new Date(year, month, 0).getDate();
    const startDay = periodKey === startKey ? start.getDate() : 1;
    const endDay = periodKey === endKey ? end.getDate() : daysInMonth;
    return {
      endDay,
      periodKey,
      startDay,
      weight: Math.max(0, endDay - startDay + 1) / Math.max(daysInMonth, 1),
    };
  });
}

export function periodRangeLabel(startPeriodKey: string, endPeriodKey: string) {
  if (startPeriodKey === endPeriodKey) {
    const { monthName, year } = monthYearFromPeriodKey(startPeriodKey);
    return `${monthName} ${year}`;
  }
  const startYear = startPeriodKey.slice(0, 4);
  const endYear = endPeriodKey.slice(0, 4);
  const startMonth = INDONESIAN_MONTHS[Number(startPeriodKey.slice(4, 6)) - 1] ?? startPeriodKey;
  const endMonth = INDONESIAN_MONTHS[Number(endPeriodKey.slice(4, 6)) - 1] ?? endPeriodKey;
  return startYear === endYear ? `${startMonth} - ${endMonth} ${endYear}` : `${startMonth} ${startYear} - ${endMonth} ${endYear}`;
}

export function dateRangeLabel(startDate: string, endDate: string) {
  const start = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return "";
  const startDay = String(start.getDate()).padStart(2, "0");
  const endDay = String(end.getDate()).padStart(2, "0");
  const startMonth = INDONESIAN_MONTHS[start.getMonth()];
  const endMonth = INDONESIAN_MONTHS[end.getMonth()];

  if (start.getTime() === end.getTime()) return `${endDay} ${endMonth} ${end.getFullYear()}`;
  if (start.getFullYear() !== end.getFullYear()) return `${startDay} ${startMonth} ${start.getFullYear()} - ${endDay} ${endMonth} ${end.getFullYear()}`;
  if (start.getMonth() !== end.getMonth()) return `${startDay} ${startMonth} - ${endDay} ${endMonth} ${end.getFullYear()}`;
  return `${startDay} - ${endDay} ${endMonth} ${end.getFullYear()}`;
}

export function formatPeriodDate(date: Date) {
  return `${String(date.getDate()).padStart(2, "0")} ${INDONESIAN_MONTHS[date.getMonth()] ?? ""} ${date.getFullYear()}`.trim();
}

export function formatDataCutoffDate(date: Date) {
  return `DATA S.D. ${formatPeriodDate(date)}`;
}

export function formatYtdLabel(endDate: string | Date) {
  const date = typeof endDate === "string" ? new Date(`${endDate}T00:00:00`) : endDate;
  if (Number.isNaN(date.getTime())) return "JANUARI";
  return `JANUARI - ${formatPeriodDate(date).replace(/^01 /, "")}`;
}

export function cutoffDateFromPeriodPart(part: Pick<DashboardPeriodPart, "endDay" | "periodKey">) {
  const year = Number(part.periodKey.slice(0, 4));
  const month = Number(part.periodKey.slice(4, 6));
  const daysInMonth = new Date(year, month, 0).getDate();
  const day = Math.max(1, Math.min(daysInMonth, Math.trunc(part.endDay)));
  return new Date(year, month - 1, day);
}

export function formatDataCutoff({ day, periodKey }: { day: number; periodKey: string }) {
  return formatDataCutoffDate(cutoffDateFromPeriodPart({ endDay: day, periodKey }));
}
