import { wmsBranchMembers, wmsHistory, wmsTeldaMasterNames, type WmsHistoryEntry } from "./history-data";
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

// Hasil export Tableau memakai header "TABLE_NAME", sedangkan seed lama menulis "Table Name".
// Tanpa penyesuaian ini upload CSV/Excel ditolak ("kurang kolom wajib: Table Name").
// Lebar dan urutan kolom tidak berubah, jadi data seed/localStorage lama tetap sejajar.
export const masterHeaders: string[] = seed.masterHeaders.map((header) => (header === "Table Name" ? "TABLE_NAME" : header));
export const initialMasterRows = seed.masterRows.map((row) => [...row]) as MasterRow[];
export const defaultWmsPeriodKey: string = seed.periodKey;
export const wmsDataVersion: string = String(seed.dataVersion);

type TargetEntry = {
  monthly: readonly number[];
  total: number;
  ytdSep: number;
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

const indexByHeader: Record<string, number> = Object.fromEntries(masterHeaders.map((header, index) => [header, index]));

function toNumber(value: unknown) {
  if (value === null || value === undefined || value === "") return 0;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function toLabel(value: unknown) {
  return String(value ?? "").trim().toUpperCase();
}

function pad2(value: number) {
  return String(value).padStart(2, "0");
}

function daysInMonth(year: number, month: number) {
  return new Date(year, month, 0).getDate();
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
  return `${previousYear}${pad2(previousMonth)}`;
}

// Kolom "Day of LAST_UPDATED_DATE" (MASTER!B) = dasar tanggal harian di workbook.
function dayFromRow(row: MasterRow) {
  const day = toNumber(row[indexByHeader["Day of LAST_UPDATED_DATE"]]);
  if (day >= 1 && day <= 31) return day;
  const tanggal = toNumber(row[indexByHeader.TANGGAL]);
  return tanggal ? tanggal % 100 : 0;
}

type JakartaNow = { day: number; hour: number; month: number; periodKey: string; year: number };

function jakartaNow(date: Date = new Date()): JakartaNow {
  const parts = new Intl.DateTimeFormat("en-CA", {
    day: "2-digit",
    hour: "2-digit",
    hourCycle: "h23",
    month: "2-digit",
    timeZone: "Asia/Jakarta",
    year: "numeric",
  })
    .formatToParts(date)
    .reduce<Record<string, string>>((acc, part) => {
      if (part.type !== "literal") acc[part.type] = part.value;
      return acc;
    }, {});
  return {
    day: Number(parts.day),
    hour: Number(parts.hour),
    month: Number(parts.month),
    periodKey: `${parts.year}${parts.month}`,
    year: Number(parts.year),
  };
}

// Padanan REAL!AH1 pada workbook: batas hari untuk "REAL MTD bulan lalu".
// - Sebelum jam 12:00 data hari ini belum lengkap, jadi dipakai tanggal kemarin.
// - Jika tanggal acuan = akhir bulan, dipakai jumlah hari bulan sebelumnya.
// - Untuk laporan bulan yang sudah lewat, seluruh bulan sebelumnya dihitung.
function previousMonthCutoffDay(periodKey: string, now: JakartaNow) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  const previousYear = month <= 1 ? year - 1 : year;
  const previousMonth = month <= 1 ? 12 : month - 1;
  const previousMonthDays = daysInMonth(previousYear, previousMonth);
  if (periodKey < now.periodKey) return previousMonthDays;

  let base = Date.UTC(now.year, now.month - 1, now.day);
  if (now.hour < 12) base -= 24 * 60 * 60 * 1000;
  const baseDate = new Date(base);
  const baseDay = baseDate.getUTCDate();
  const baseMonthDays = new Date(Date.UTC(baseDate.getUTCFullYear(), baseDate.getUTCMonth() + 1, 0)).getUTCDate();
  if (baseDay === baseMonthDays) {
    return new Date(Date.UTC(baseDate.getUTCFullYear(), baseDate.getUTCMonth(), 0)).getUTCDate();
  }
  return baseDay;
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

// TARGET FM = INDEX(TARGET!bulan, bulan laporan); TARGET YTD = SUM(Jan .. bulan laporan).
function targetFor(scope: ReportScope, label: string, periodKey: string) {
  const monthIndex = Number(periodKey.slice(4, 6)) - 1;
  const table = (scope === "witel" ? seed.targets.witel : seed.targets.telda) as Record<string, TargetEntry>;
  const target = table[label];
  const monthly = target?.monthly?.[monthIndex] ?? 0;
  const targetYtd = target?.monthly?.slice(0, monthIndex + 1).reduce((sum: number, value: number) => sum + toNumber(value), 0) ?? 0;
  return { monthly: toNumber(monthly), targetYtd };
}

function labelsForScope(scope: ReportScope) {
  return (scope === "witel" ? seed.reportRows.witel : seed.reportRows.telda)
    .map((label) => String(label).trim())
    .filter(Boolean);
}

// ---------------------------------------------------------------------------
// Indeks penjualan: nama (WITEL / TELKOM_DAERAH) -> periode -> penjualan per hari (index 1..31).
// Setara SUMPRODUCT((MASTER!M=label)*(MASTER!B=hari)*MASTER!BB) pada Report Sementara.
// ---------------------------------------------------------------------------
type SalesIndex = Map<string, Map<string, number[]>>;

function buildSalesIndex(rows: MasterRow[], scopeColumn: number): SalesIndex {
  const index: SalesIndex = new Map();
  const salesColumn = indexByHeader.Jml_Sales;
  for (const row of rows) {
    const sales = toNumber(row[salesColumn]);
    if (!sales) continue;
    const day = dayFromRow(row);
    if (day < 1 || day > 31) continue;
    const name = toLabel(row[scopeColumn]);
    if (!name) continue;
    const period = periodFromRow(row);
    let byPeriod = index.get(name);
    if (!byPeriod) {
      byPeriod = new Map();
      index.set(name, byPeriod);
    }
    let days = byPeriod.get(period);
    if (!days) {
      days = Array.from({ length: 32 }, () => 0);
      byPeriod.set(period, days);
    }
    days[day] += sales;
  }
  return index;
}

function indexedDays(index: SalesIndex, names: readonly string[], period: string) {
  const days = Array.from({ length: 32 }, () => 0);
  for (const name of names) {
    const source = index.get(name)?.get(period);
    if (!source) continue;
    for (let day = 1; day <= 31; day += 1) days[day] += source[day];
  }
  return days;
}

function indexedMonthTotal(index: SalesIndex, names: readonly string[], period: string) {
  return indexedDays(index, names, period).reduce((sum, value) => sum + value, 0);
}

function combineHistory(entries: WmsHistoryEntry[]): WmsHistoryEntry | null {
  if (entries.length === 0) return null;
  const add = (selector: (entry: WmsHistoryEntry) => readonly number[], length: number) =>
    Array.from({ length }, (_, index) => entries.reduce((sum, entry) => sum + toNumber(selector(entry)[index]), 0));
  return {
    prevDays: add((entry) => entry.prevDays, 31),
    prevTotal: entries.reduce((sum, entry) => sum + entry.prevTotal, 0),
    y2025: add((entry) => entry.y2025, 12),
    y2026: add((entry) => entry.y2026, 12),
  };
}

function historyFor(scope: ReportScope, label: string) {
  if (scope === "witel") {
    const entry = wmsHistory.witel[toLabel(label)];
    return entry ? combineHistory([entry]) : null;
  }
  const members = wmsBranchMembers[toLabel(label)] ?? [toLabel(label)];
  return combineHistory(members.map((member) => wmsHistory.telda[member]).filter(Boolean));
}

function masterNamesFor(scope: ReportScope, label: string) {
  if (scope === "witel") return [toLabel(label)];
  const members = wmsBranchMembers[toLabel(label)] ?? [toLabel(label)];
  return members.flatMap((member) => wmsTeldaMasterNames[member] ?? [member]).map(toLabel);
}

function makeRows(rows: MasterRow[], periodKey: string, scope: ReportScope, now: JakartaNow) {
  const month = Number(periodKey.slice(4, 6));
  const year = Number(periodKey.slice(0, 4));
  const dayCount = daysInMonth(year, month);
  const scopeColumn = scope === "witel" ? indexByHeader.WITEL : indexByHeader.TELKOM_DAERAH;
  const salesIndex = buildSalesIndex(rows, scopeColumn);
  const cutoffPrevious = previousMonthCutoffDay(periodKey, now);
  const previousPeriod = previousPeriodKey(periodKey) ?? "";
  const previousYear = Number(previousPeriod.slice(0, 4));
  const previousMonth = Number(previousPeriod.slice(4, 6));
  const historyApplies = year === 2026;

  const reportRows: WmsReportRow[] = labelsForScope(scope).map((label) => {
    const names = masterNamesFor(scope, label);
    const history = historyApplies ? historyFor(scope, label) : null;

    // Realisasi harian bulan berjalan dan MTD.
    const monthDays = indexedDays(salesIndex, names, periodKey);
    const days = Array.from({ length: dayCount }, (_, index) => monthDays[index + 1]);
    const mtd = days.reduce((sum, value) => sum + value, 0);
    const target = targetFor(scope, label, periodKey);

    // REAL MTD bulan lalu (sampai batas hari REAL!AH1) dan REAL FM bulan lalu (REAL!AG).
    let previousDays: readonly number[];
    let realFmPrev: number;
    if (history && wmsHistory.reportPeriodKey === periodKey) {
      previousDays = history.prevDays;
      realFmPrev = history.prevTotal;
    } else {
      const fromRows = indexedDays(salesIndex, names, previousPeriod).slice(1);
      previousDays = fromRows;
      realFmPrev =
        history && previousYear === 2026 && previousMonth >= 1 && previousMonth <= wmsHistory.lastClosedMonth2026
          ? toNumber(history.y2026[previousMonth - 1])
          : fromRows.reduce((sum, value) => sum + value, 0);
    }
    const realMtdPrev = previousDays.slice(0, cutoffPrevious).reduce((sum, value) => sum + toNumber(value), 0);

    // REAL YTD = total bulan Jan..(bulan-1) + MTD bulan berjalan (REAL!AX = SUM(AL:AW)).
    let realYtd = mtd;
    for (let m = 1; m < month; m += 1) {
      realYtd +=
        history && m <= wmsHistory.lastClosedMonth2026
          ? toNumber(history.y2026[m - 1])
          : indexedMonthTotal(salesIndex, names, `${year}${pad2(m)}`);
    }

    // YTD tahun lalu pada bulan yang sama (REAL!BN) = pembanding GR YTD.
    let realYtdPrev = 0;
    for (let m = 1; m <= month; m += 1) {
      realYtdPrev += history ? toNumber(history.y2025[m - 1]) : indexedMonthTotal(salesIndex, names, `${year - 1}${pad2(m)}`);
    }

    return {
      ach: target.monthly ? mtd / target.monthly : 0,
      achYtd: target.targetYtd ? realYtd / target.targetYtd : 0,
      days: days.map((value) => (value === 0 ? null : value)),
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

// Sama dengan Excel RANK(): nilai yang sama mendapat peringkat yang sama
// (peringkat = 1 + jumlah baris dengan nilai lebih besar).
function assignRanks(rows: WmsReportRow[], source: "ach" | "achYtd", target: "rank" | "rankYtd") {
  for (const row of rows) {
    const current = row[source];
    row[target] = current === null || current === undefined ? null : 1 + rows.filter((other) => (other[source] ?? -Infinity) > current).length;
  }
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

export function computeWmsReport(
  rows: MasterRow[],
  periodKey: string = seed.periodKey,
  options: { now?: Date } = {},
): WmsComputedReport {
  const now = jakartaNow(options.now);
  const periodOptions = periodOptionsFromRows(rows);
  const activePeriod = periodOptions.some((option) => option.periodKey === periodKey) ? periodKey : periodOptions.at(-1)?.periodKey ?? seed.periodKey;
  const monthIndex = Number(activePeriod.slice(4, 6)) - 1;
  const witelRows = makeRows(rows, activePeriod, "witel", now);
  const teldaRows = makeRows(rows, activePeriod, "telda", now);
  const daysInActiveMonth = daysInMonth(Number(activePeriod.slice(0, 4)), Number(activePeriod.slice(4, 6)));
  const periodRows = rows.filter((row) => periodFromRow(row) === activePeriod);
  const latestDataDay = Math.max(1, ...periodRows.map(dayFromRow));
  const latestDay =
    activePeriod === now.periodKey
      ? Math.max(latestDataDay, Math.max(1, Math.min(daysInActiveMonth, now.day - 1)))
      : latestDataDay;
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
