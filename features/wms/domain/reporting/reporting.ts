import { wmsSeedData } from "./seed-data";

export type MasterCell = string | number | boolean | null;
export type MasterRow = MasterCell[];
export type ReportScope = "witel" | "telda";

export type WmsReportRow = {
  ach: number | null;
  achYtd: number | null;
  days: (number | null)[];
  growthFm: number | null;
  growthMtd: number | null;
  growthYtd: number | null;
  label: string;
  mtd: number;
  rank: number | null;
  rankYtd: number | null;
  realFmPrev: number;
  realMtdPrev: number;
  realYtdPrev: number;
  realYtd: number;
  shortageFm: number;
  shortageYtd: number;
  targetFm: number;
  targetYtd: number;
};

export type WmsComputedReport = {
  cutoffDay: number;
  cutoffLabel: string;
  latestUpdatedAt: string | null;
  periodKey: string;
  periodLabel: string;
  previousMonthLabel: string;
  telda: {
    rows: WmsReportRow[];
    total: WmsReportRow;
  };
  validation: {
    masterRows: number;
    witelMtd: number;
  };
  witel: {
    rows: WmsReportRow[];
    total: WmsReportRow;
  };
};

const seed = wmsSeedData;

export const masterHeaders = [...seed.masterHeaders];
export const initialMasterRows = seed.masterRows.map((row) => [...row]) as MasterRow[];
export const defaultWmsPeriodKey: string = seed.periodKey;
export const wmsDataVersion: string = String(seed.dataVersion);

type TargetEntry = {
  monthly: readonly number[];
  total: number;
  ytdSep: number;
};

type PreviousEntry = {
  currentYtd: number;
  previousMonthDays: readonly number[];
  previousMonthMtd: number;
  previousMonthTotal: number;
  previousYtd: number;
};

const monthNames = [
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
const shortMonthNames = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AUG", "SEPT", "OKT", "NOV", "DES"];

const indexByHeader = Object.fromEntries(masterHeaders.map((header, index) => [header, index]));

function toNumber(value: unknown) {
  if (value === null || value === undefined || value === "") return 0;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function toLabel(value: unknown) {
  return String(value ?? "").trim().toUpperCase();
}

function periodFromRow(row: MasterRow) {
  const tanggal = toNumber(row[indexByHeader.TANGGAL]);
  if (tanggal) return String(Math.trunc(tanggal / 100));
  const nper = row[indexByHeader.NPER];
  if (typeof nper === "number") return String(nper);
  return String(nper ?? "").replace(/\D/g, "").slice(0, 6);
}

export function periodKeyFromMasterRow(row: MasterRow) {
  const periodKey = periodFromRow(row);
  return /^\d{6}$/.test(periodKey) ? periodKey : null;
}

export function periodLabelFromKey(periodKey: string) {
  const monthIndex = Number(periodKey.slice(4, 6)) - 1;
  return `${monthNames[monthIndex] ?? periodKey} ${periodKey.slice(0, 4)}`;
}

function previousMonthLabelFromKey(periodKey: string) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  const previousMonth = month <= 1 ? 12 : month - 1;
  const previousYear = month <= 1 ? year - 1 : year;
  return `${shortMonthNames[previousMonth - 1] ?? seed.previousMonthLabel} ${Number.isFinite(previousYear) ? previousYear : periodKey.slice(0, 4)}`;
}

function previousPeriodKey(periodKey: string) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  if (!Number.isFinite(year) || !Number.isFinite(month)) return null;
  const previousMonth = month <= 1 ? 12 : month - 1;
  const previousYear = month <= 1 ? year - 1 : year;
  return `${previousYear}${String(previousMonth).padStart(2, "0")}`;
}

function dayFromRow(row: MasterRow) {
  const day = toNumber(row[indexByHeader["Day of LAST_UPDATED_DATE"]]);
  if (day >= 1 && day <= 31) return day;
  const tanggal = toNumber(row[indexByHeader.TANGGAL]);
  return tanggal ? tanggal % 100 : 0;
}

function currentJakartaPeriod() {
  const parts = new Intl.DateTimeFormat("en-CA", {
    day: "2-digit",
    month: "2-digit",
    timeZone: "Asia/Jakarta",
    year: "numeric",
  })
    .formatToParts(new Date())
    .reduce<Record<string, string>>((acc, part) => {
      if (part.type !== "literal") acc[part.type] = part.value;
      return acc;
    }, {});
  return {
    day: Number(parts.day),
    periodKey: `${parts.year}${parts.month}`,
  };
}

function latestUpdatedAt(rows: MasterRow[]) {
  const column = indexByHeader.LAST_UPDATE_DATE;
  const values = rows.map((row) => row[column]).filter(Boolean).map(String).sort();
  return values.at(-1) ?? null;
}

export function periodOptionsFromRows(rows: MasterRow[]) {
  const keys = Array.from(new Set(rows.map(periodFromRow).filter((period) => /^\d{6}$/.test(period)))).sort();
  const options = keys.map((periodKey) => {
    return {
      label: periodLabelFromKey(periodKey),
      periodKey,
    };
  });
  return options.length > 0 ? options : [{ label: seed.periodLabel, periodKey: seed.periodKey }];
}

function targetFor(scope: ReportScope, label: string, periodKey: string) {
  const monthIndex = Number(periodKey.slice(4, 6)) - 1;
  const table = (scope === "witel" ? seed.targets.witel : seed.targets.telda) as Record<string, TargetEntry>;
  const target = table[label];
  const monthly = target?.monthly?.[monthIndex] ?? 0;
  const targetYtd = target?.monthly?.slice(0, monthIndex + 1).reduce((sum: number, value: number) => sum + toNumber(value), 0) ?? 0;
  return { monthly: toNumber(monthly), targetYtd };
}

function previousFor(scope: ReportScope, label: string) {
  const table = (scope === "witel" ? seed.previousReal.witel : seed.previousReal.telda) as Record<string, PreviousEntry>;
  return table[label] ?? null;
}

function previousStatsFor(
  rows: MasterRow[],
  periodKey: string,
  scopeColumn: number,
  label: string,
  cutoffDay: number,
) {
  const previousPeriod = previousPeriodKey(periodKey);
  const activeYear = periodKey.slice(0, 4);
  const activeMonth = Number(periodKey.slice(4, 6));
  const normalized = toLabel(label);
  let realMtdPrev = 0;
  let realFmPrev = 0;
  let realYtdPrev = 0;

  for (const row of rows) {
    if (toLabel(row[scopeColumn]) !== normalized) continue;
    const period = periodFromRow(row);
    const sales = toNumber(row[indexByHeader.Jml_Sales]);
    if (!sales) continue;

    if (period === previousPeriod) {
      realFmPrev += sales;
      const day = dayFromRow(row);
      if (day >= 1 && day <= cutoffDay) realMtdPrev += sales;
    }

    if (period.slice(0, 4) === activeYear && Number(period.slice(4, 6)) < activeMonth) {
      realYtdPrev += sales;
    }
  }

  return { realFmPrev, realMtdPrev, realYtdPrev };
}

function labelsForScope(scope: ReportScope) {
  return (scope === "witel" ? seed.reportRows.witel : seed.reportRows.telda)
    .map((label) => String(label).trim())
    .filter(Boolean);
}

function makeRows(rows: MasterRow[], periodKey: string, scope: ReportScope) {
  const month = Number(periodKey.slice(4, 6));
  const year = Number(periodKey.slice(0, 4));
  const daysInMonth = new Date(year, month, 0).getDate();
  const currentPeriod = currentJakartaPeriod();
  const cutoffDay =
    periodKey === currentPeriod.periodKey
      ? Math.max(1, Math.min(daysInMonth, currentPeriod.day - 1))
      : daysInMonth;
  const scopeColumn = scope === "witel" ? indexByHeader.WITEL : indexByHeader.TELKOM_DAERAH;
  const salesColumn = indexByHeader.Jml_Sales;
  const labels = labelsForScope(scope);

  const reportRows: WmsReportRow[] = labels.map((rawLabel) => {
    const label = String(rawLabel).trim();
    const normalized = toLabel(label);
    const days = Array.from({ length: daysInMonth }, () => 0);
    let realYtd = 0;

    for (const row of rows) {
      if (toLabel(row[scopeColumn]) !== normalized) continue;
      const period = periodFromRow(row);
      const sales = toNumber(row[salesColumn]);
      if (!sales) continue;
      if (period === periodKey) {
        const day = dayFromRow(row);
        if (day >= 1 && day <= daysInMonth) days[day - 1] += sales;
      }
      if (period.slice(0, 4) === String(year) && Number(period.slice(4, 6)) <= month) {
        realYtd += sales;
      }
    }

    const cleanDays = days.map((value) => (value === 0 ? null : value));
    const mtd = days.reduce((sum, value) => sum + value, 0);
    const target = targetFor(scope, label, periodKey);
    const previous = previousFor(scope, label);
    const dynamicPrevious = previousStatsFor(rows, periodKey, scopeColumn, label, cutoffDay);
    const realMtdPrev = dynamicPrevious.realMtdPrev || toNumber(previous?.previousMonthMtd);
    const realFmPrev = dynamicPrevious.realFmPrev || toNumber(previous?.previousMonthTotal);
    const realYtdPrev = dynamicPrevious.realYtdPrev || toNumber(previous?.previousYtd);
    return {
      ach: target.monthly ? mtd / target.monthly : 0,
      achYtd: target.targetYtd ? realYtd / target.targetYtd : 0,
      days: cleanDays,
      growthFm: realFmPrev ? mtd / realFmPrev - 1 : null,
      growthMtd: realMtdPrev ? mtd / realMtdPrev - 1 : null,
      growthYtd: realYtdPrev ? realYtd / realYtdPrev - 1 : null,
      label,
      mtd,
      rank: null,
      rankYtd: null,
      realFmPrev,
      realMtdPrev,
      realYtdPrev,
      realYtd,
      shortageFm: mtd - target.monthly,
      shortageYtd: realYtd - target.targetYtd,
      targetFm: target.monthly,
      targetYtd: target.targetYtd,
    };
  });

  assignRanks(reportRows, "ach", "rank");
  assignRanks(reportRows, "achYtd", "rankYtd");
  return reportRows;
}

function assignRanks(rows: WmsReportRow[], source: "ach" | "achYtd", target: "rank" | "rankYtd") {
  const ranked = rows
    .map((row, index) => ({ index, value: row[source] ?? -Infinity }))
    .sort((a, b) => b.value - a.value);
  ranked.forEach((item, index) => {
    rows[item.index][target] = Number.isFinite(item.value) ? index + 1 : null;
  });
}

function totalRow(rows: WmsReportRow[], label = "Total"): WmsReportRow {
  const dayCount = Math.max(...rows.map((row) => row.days.length), 31);
  const days = Array.from({ length: dayCount }, (_, index) => {
    const value = rows.reduce((sum, row) => sum + toNumber(row.days[index]), 0);
    return value === 0 ? null : value;
  });
  const mtd = rows.reduce((sum, row) => sum + row.mtd, 0);
  const targetFm = rows.reduce((sum, row) => sum + row.targetFm, 0);
  const targetYtd = rows.reduce((sum, row) => sum + row.targetYtd, 0);
  const realMtdPrev = rows.reduce((sum, row) => sum + row.realMtdPrev, 0);
  const realFmPrev = rows.reduce((sum, row) => sum + row.realFmPrev, 0);
  const realYtdPrev = rows.reduce((sum, row) => sum + row.realYtdPrev, 0);
  const realYtd = rows.reduce((sum, row) => sum + row.realYtd, 0);
  return {
    ach: targetFm ? mtd / targetFm : 0,
    achYtd: targetYtd ? realYtd / targetYtd : 0,
    days,
    growthFm: realFmPrev ? mtd / realFmPrev - 1 : null,
    growthMtd: realMtdPrev ? mtd / realMtdPrev - 1 : null,
    growthYtd: realYtdPrev ? realYtd / realYtdPrev - 1 : null,
    label,
    mtd,
    rank: null,
    rankYtd: null,
    realFmPrev,
    realMtdPrev,
    realYtdPrev,
    realYtd,
    shortageFm: mtd - targetFm,
    shortageYtd: realYtd - targetYtd,
    targetFm,
    targetYtd,
  };
}

export function computeWmsReport(rows: MasterRow[], periodKey: string = seed.periodKey): WmsComputedReport {
  const options = periodOptionsFromRows(rows);
  const activePeriod = options.some((option) => option.periodKey === periodKey) ? periodKey : options.at(-1)?.periodKey ?? seed.periodKey;
  const monthIndex = Number(activePeriod.slice(4, 6)) - 1;
  const witelRows = makeRows(rows, activePeriod, "witel");
  const teldaRows = makeRows(rows, activePeriod, "telda");
  const currentPeriod = currentJakartaPeriod();
  const daysInMonth = new Date(Number(activePeriod.slice(0, 4)), Number(activePeriod.slice(4, 6)), 0).getDate();
  const latestDataDay = Math.max(1, ...rows.filter((row) => periodFromRow(row) === activePeriod).map(dayFromRow));
  const latestDay =
    activePeriod === currentPeriod.periodKey
      ? Math.max(latestDataDay, Math.max(1, Math.min(daysInMonth, currentPeriod.day - 1)))
      : latestDataDay;
  const periodRows = rows.filter((row) => periodFromRow(row) === activePeriod);
  return {
    cutoffDay: latestDay,
    cutoffLabel: `DATA S.D. TANGGAL ${latestDay}`,
    latestUpdatedAt: latestUpdatedAt(periodRows),
    periodKey: activePeriod,
    periodLabel: `${monthNames[monthIndex] ?? seed.periodLabel} ${activePeriod.slice(0, 4)}`,
    previousMonthLabel: previousMonthLabelFromKey(activePeriod),
    telda: {
      rows: teldaRows,
      total: totalRow(teldaRows, "YK JATENG SELATAN"),
    },
    validation: {
      masterRows: rows.length,
      witelMtd: witelRows.reduce((sum, row) => sum + row.mtd, 0),
    },
    witel: {
      rows: witelRows,
      total: totalRow(witelRows),
    },
  };
}

