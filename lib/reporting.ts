import { seedData } from "./seed-data";

export type CellValue = string | number | boolean | null;
export type MasterRow = CellValue[];

type TargetConfig = {
  label: string;
  rawDatel?: string;
  monthly: readonly CellValue[];
  targetYtd: CellValue;
};

type PreviousConfig = {
  mtd: Record<string, CellValue>;
  fm: Record<string, CellValue>;
  ytd: Record<string, CellValue>;
  ytdPrev: Record<string, CellValue>;
};
type ReportingPeriod = {
  days: number;
  month: number;
  monthIndex: number;
  monthName: string;
  previousMonthName: string;
  year: number;
};

export type ReportRow = {
  label: string;
  targetFm: number;
  targetYtd: number;
  days: (number | null)[];
  mtd: number;
  targetAch: number | null;
  rankAch: number | null;
  shortageFm: number;
  realMtdPrev: number;
  growthMtd: number | null;
  realFmPrev: number;
  growthFm: number | null;
  realYtd: number;
  achYtd: number | null;
  growthYtd: number | null;
  rankYtd: number | null;
  shortageYtd: number;
};

export type ReportTable = {
  title: string;
  subtitle: string;
  rows: ReportRow[];
  total: ReportRow;
};

export type ComputedReport = {
  periodLabel: string;
  previousMonthLabel: string;
  cutoffLabel: string;
  source: string;
  updatedAt: string | null;
  witel: ReportTable;
  branch: ReportTable;
  targetStf: {
    witel: ReportTable;
    branch: ReportTable;
  };
  dailyTrend: { day: number; realisasi: number; cumulative: number }[];
  validation: { checked: number; mismatches: number; notes: string[] };
};

const DEFAULT_PERIOD = seedData.period;
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
const MASTER_INDEX = {
  witel: seedData.masterHeaders.indexOf("WITEL"),
  datel: seedData.masterHeaders.indexOf("DATEL"),
  ndem: seedData.masterHeaders.indexOf("NDEM"),
  lastUpdated: seedData.masterHeaders.indexOf("LAST UPDATED DATE"),
};

const excelEpoch = new Date(1899, 11, 30);

function toNumber(value: CellValue): number {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "string") {
    const parsed = Number(value.replace(/,/g, "").trim());
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function asText(value: CellValue): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function validNdem(value: CellValue): boolean {
  return asText(value).replace(/'/g, "").trim().length > 0;
}

function parsePeriodKey(periodKey: string | null | undefined) {
  if (!periodKey || !/^\d{6}$/.test(periodKey)) return null;
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  if (!Number.isFinite(year) || month < 1 || month > 12) return null;
  return { month, year };
}

function periodFromKey(periodKey: string | null | undefined): ReportingPeriod {
  const parsed = parsePeriodKey(periodKey);
  const year = parsed?.year ?? DEFAULT_PERIOD.year;
  const month = parsed?.month ?? DEFAULT_PERIOD.month;
  const previousMonth = month <= 1 ? 12 : month - 1;
  return {
    days: new Date(year, month, 0).getDate(),
    month,
    monthIndex: month - 1,
    monthName: monthNames[month - 1] ?? DEFAULT_PERIOD.monthName,
    previousMonthName: shortMonthNames[previousMonth - 1] ?? DEFAULT_PERIOD.previousMonthName,
    year,
  };
}

function excelSerialToDate(value: number): Date {
  return new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
}

function normalizeWorkbookDate(value: CellValue, period: ReportingPeriod): Date | null {
  let parsed: Date | null = null;

  if (typeof value === "number") {
    parsed = excelSerialToDate(value);
  } else if (typeof value === "string") {
    const text = value.trim();
    const dmy = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    const ymd = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (dmy) {
      parsed = new Date(
        Number(dmy[3]),
        Number(dmy[2]) - 1,
        Number(dmy[1]),
        Number(dmy[4] ?? 0),
        Number(dmy[5] ?? 0),
        Number(dmy[6] ?? 0),
      );
    } else if (ymd) {
      parsed = new Date(
        Number(ymd[1]),
        Number(ymd[2]) - 1,
        Number(ymd[3]),
        Number(ymd[4] ?? 0),
        Number(ymd[5] ?? 0),
        Number(ymd[6] ?? 0),
      );
    }
  }

  if (!parsed || Number.isNaN(parsed.getTime())) return null;

  if (parsed.getMonth() + 1 === period.month) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  // Mirrors the workbook helper formula that swaps day/month when Excel parses
  // imported dates with the wrong locale.
  return new Date(parsed.getFullYear(), parsed.getDate() - 1, parsed.getMonth() + 1);
}

function buildCounts(masterRows: MasterRow[], key: "witel" | "datel", period: ReportingPeriod) {
  const index = key === "witel" ? MASTER_INDEX.witel : MASTER_INDEX.datel;
  const counts = new Map<string, number[]>();

  for (const [rowIndex, row] of masterRows.entries()) {
    // The workbook's FILTER spills the first MASTER record into HELPER!A1, while
    // the COUNTIFS ranges start from HELPER!A2. Keeping this offset preserves the
    // published report totals exactly.
    if (rowIndex === 0) continue;
    if (!validNdem(row[MASTER_INDEX.ndem])) continue;
    const label = asText(row[index]);
    if (!label) continue;
    const normalizedDate = normalizeWorkbookDate(row[MASTER_INDEX.lastUpdated], period);
    if (
      !normalizedDate ||
      normalizedDate.getFullYear() !== period.year ||
      normalizedDate.getMonth() !== period.monthIndex
    ) {
      continue;
    }
    const day = normalizedDate.getDate();
    if (day < 1 || day > period.days) continue;
    const existing = counts.get(label) ?? Array(period.days).fill(0);
    existing[day - 1] += 1;
    counts.set(label, existing);
  }

  return counts;
}

function sumDayArrays(arrays: number[][], dayCount: number): number[] {
  return Array.from({ length: dayCount }, (_, day) =>
    arrays.reduce((sum, current) => sum + (current[day] ?? 0), 0),
  );
}

function rankDescending(rows: ReportRow[], metric: "targetAch" | "achYtd") {
  for (const row of rows) {
    const current = row[metric];
    row[metric === "targetAch" ? "rankAch" : "rankYtd"] =
      current === null ? null : rows.filter((candidate) => (candidate[metric] ?? -Infinity) > current).length + 1;
  }
}

function safeRatio(numerator: number, denominator: number): number | null {
  return denominator === 0 ? null : numerator / denominator;
}

function safeGrowth(current: number, previous: number): number | null {
  return previous === 0 ? null : current / previous - 1;
}

function makeRow(
  label: string,
  target: TargetConfig,
  daysRaw: number[],
  previous: PreviousConfig,
  referenceMtd: number,
  monthIndex: number,
): ReportRow {
  const days = daysRaw.map((value) => (value === 0 ? null : value));
  const mtd = daysRaw.reduce((sum, value) => sum + value, 0);
  const targetFm = toNumber(target.monthly[monthIndex]);
  const targetYtd = toNumber(target.targetYtd);
  const realMtdPrev = toNumber(previous.mtd[label]);
  const realFmPrev = toNumber(previous.fm[label]);
  const excelYtd = toNumber(previous.ytd[label]);
  const ytdBaseBeforeCurrentMonth = Math.max(0, excelYtd - referenceMtd);
  const realYtd = ytdBaseBeforeCurrentMonth + mtd;
  const ytdPrev = toNumber(previous.ytdPrev[label]);

  return {
    label,
    targetFm,
    targetYtd,
    days,
    mtd,
    targetAch: safeRatio(mtd, targetFm),
    rankAch: null,
    shortageFm: mtd - targetFm,
    realMtdPrev,
    growthMtd: safeGrowth(mtd, realMtdPrev),
    realFmPrev,
    growthFm: safeGrowth(mtd, realFmPrev),
    realYtd,
    achYtd: safeRatio(realYtd, targetYtd),
    growthYtd: safeGrowth(realYtd, ytdPrev),
    rankYtd: null,
    shortageYtd: realYtd - targetYtd,
  };
}

function makeTotal(label: string, rows: ReportRow[], dayCount: number): ReportRow {
  const daysRaw = Array.from({ length: dayCount }, (_, day) =>
    rows.reduce((sum, row) => sum + (row.days[day] ?? 0), 0),
  );
  const total: ReportRow = {
    label,
    targetFm: rows.reduce((sum, row) => sum + row.targetFm, 0),
    targetYtd: rows.reduce((sum, row) => sum + row.targetYtd, 0),
    days: daysRaw,
    mtd: rows.reduce((sum, row) => sum + row.mtd, 0),
    targetAch: null,
    rankAch: null,
    shortageFm: 0,
    realMtdPrev: rows.reduce((sum, row) => sum + row.realMtdPrev, 0),
    growthMtd: null,
    realFmPrev: rows.reduce((sum, row) => sum + row.realFmPrev, 0),
    growthFm: null,
    realYtd: rows.reduce((sum, row) => sum + row.realYtd, 0),
    achYtd: null,
    growthYtd: null,
    rankYtd: null,
    shortageYtd: 0,
  };
  total.targetAch = safeRatio(total.mtd, total.targetFm);
  total.shortageFm = total.mtd - total.targetFm;
  total.growthMtd = safeGrowth(total.mtd, total.realMtdPrev);
  total.growthFm = safeGrowth(total.mtd, total.realFmPrev);
  total.achYtd = safeRatio(total.realYtd, total.targetYtd);
  total.shortageYtd = total.realYtd - total.targetYtd;
  const previousYtd = rows.reduce((sum, row) => {
    const ytdPrevious = toNumber(
      (seedData.previousByWitel.ytdPrev as Record<string, CellValue>)[row.label] ??
        (seedData.previousByBranch.ytdPrev as Record<string, CellValue>)[row.label],
    );
    return sum + ytdPrevious;
  }, 0);
  total.growthYtd = safeGrowth(total.realYtd, previousYtd);
  return total;
}

type ReferenceKind = "witel" | "branch" | "stfWitel" | "stfBranch";

function referenceMtd(label: string, kind: ReferenceKind) {
  const source = seedData.excelReference[kind];
  return toNumber((source.find((row) => row.label === label)?.mtd ?? 0) as CellValue);
}

function validation(rows: ReportRow[], kind: ReferenceKind) {
  const reference = seedData.excelReference[kind];
  const notes: string[] = [];
  let checked = 0;
  let mismatches = 0;
  for (const row of rows) {
    const expected = reference.find((item) => item.label === row.label);
    if (!expected) continue;
    checked += 1;
    const fields: Array<keyof ReportRow> = ["targetFm", "targetYtd", "mtd", "shortageFm"];
    for (const field of fields) {
      const actualValue = toNumber(row[field] as CellValue);
      const expectedValue = toNumber(expected[field] as CellValue);
      if (Math.abs(actualValue - expectedValue) > 0.0001) {
        mismatches += 1;
        notes.push(`${kind}:${row.label}:${field} expected ${expectedValue}, got ${actualValue}`);
      }
    }
  }
  return { checked, mismatches, notes };
}

function makeBranchRows(
  targets: readonly TargetConfig[],
  datelCounts: Map<string, number[]>,
  referenceKind: "branch" | "stfBranch",
  period: ReportingPeriod,
) {
  const teldaRows = new Map(
    targets.map((target) => {
      const days = datelCounts.get(target.rawDatel ?? target.label) ?? Array(period.days).fill(0);
      const row = makeRow(
        target.label,
        target,
        days,
        seedData.previousByBranch as PreviousConfig,
        referenceMtd(target.label, referenceKind),
        period.monthIndex,
      );
      return [target.label, row] as const;
    }),
  );

  const rows = seedData.branchGroups.map((group) => {
    const memberRows = group.members.map((member) => teldaRows.get(member)).filter(Boolean) as ReportRow[];
    const daysRaw = sumDayArrays(memberRows.map((row) => row.days.map((day) => day ?? 0)), period.days);
    const target: TargetConfig = {
      label: group.label,
      monthly: Array.from({ length: 12 }, (_, month) =>
        memberRows.reduce((sum, row) => sum + (month === period.monthIndex ? row.targetFm : 0), 0),
      ),
      targetYtd: memberRows.reduce((sum, row) => sum + row.targetYtd, 0),
    };
    const previous: PreviousConfig = {
      mtd: { [group.label]: memberRows.reduce((sum, row) => sum + row.realMtdPrev, 0) },
      fm: { [group.label]: memberRows.reduce((sum, row) => sum + row.realFmPrev, 0) },
      ytd: { [group.label]: memberRows.reduce((sum, row) => sum + row.realYtd, 0) },
      ytdPrev: {
        [group.label]: group.members.reduce(
          (sum, member) => sum + toNumber((seedData.previousByBranch.ytdPrev as Record<string, CellValue>)[member]),
          0,
        ),
      },
    };
    const row = makeRow(group.label, target, daysRaw, previous, referenceMtd(group.label, referenceKind), period.monthIndex);
    row.realYtd = memberRows.reduce((sum, member) => sum + member.realYtd, 0);
    row.achYtd = safeRatio(row.realYtd, row.targetYtd);
    row.growthYtd = safeGrowth(row.realYtd, previous.ytdPrev[group.label] as number);
    row.shortageYtd = row.realYtd - row.targetYtd;
    return row;
  });
  rankDescending(rows, "targetAch");
  rankDescending(rows, "achYtd");
  return rows;
}

export function computeReport(masterRows: MasterRow[], savedAt: string | null = null, periodKey?: string): ComputedReport {
  const period = periodFromKey(periodKey);
  const witelCounts = buildCounts(masterRows, "witel", period);
  const datelCounts = buildCounts(masterRows, "datel", period);

  const witelRows = seedData.witelTargets.map((target) =>
    makeRow(
      target.label,
      target,
      witelCounts.get(target.label) ?? Array(period.days).fill(0),
      seedData.previousByWitel as PreviousConfig,
      referenceMtd(target.label, "witel"),
      period.monthIndex,
    ),
  );
  rankDescending(witelRows, "targetAch");
  rankDescending(witelRows, "achYtd");

  const teldaRows = new Map(
    seedData.teldaTargets.map((target) => {
      const days = datelCounts.get(target.rawDatel) ?? Array(period.days).fill(0);
      const row = makeRow(
        target.label,
        target,
        days,
        seedData.previousByBranch as PreviousConfig,
        referenceMtd(target.label, "branch"),
        period.monthIndex,
      );
      return [target.label, row] as const;
    }),
  );

  const branchRows = seedData.branchGroups.map((group) => {
    const memberRows = group.members.map((member) => teldaRows.get(member)).filter(Boolean) as ReportRow[];
    const daysRaw = sumDayArrays(memberRows.map((row) => row.days.map((day) => day ?? 0)), period.days);
    const target: TargetConfig = {
      label: group.label,
      monthly: Array.from({ length: 12 }, (_, month) =>
        memberRows.reduce((sum, row) => sum + (month === period.monthIndex ? row.targetFm : 0), 0),
      ),
      targetYtd: memberRows.reduce((sum, row) => sum + row.targetYtd, 0),
    };
    const previous: PreviousConfig = {
      mtd: { [group.label]: memberRows.reduce((sum, row) => sum + row.realMtdPrev, 0) },
      fm: { [group.label]: memberRows.reduce((sum, row) => sum + row.realFmPrev, 0) },
      ytd: { [group.label]: memberRows.reduce((sum, row) => sum + row.realYtd, 0) },
      ytdPrev: {
        [group.label]: group.members.reduce(
          (sum, member) => sum + toNumber((seedData.previousByBranch.ytdPrev as Record<string, CellValue>)[member]),
          0,
        ),
      },
    };
    const row = makeRow(group.label, target, daysRaw, previous, referenceMtd(group.label, "branch"), period.monthIndex);
    row.realYtd = memberRows.reduce((sum, member) => sum + member.realYtd, 0);
    row.achYtd = safeRatio(row.realYtd, row.targetYtd);
    row.growthYtd = safeGrowth(row.realYtd, previous.ytdPrev[group.label] as number);
    row.shortageYtd = row.realYtd - row.targetYtd;
    return row;
  });
  rankDescending(branchRows, "targetAch");
  rankDescending(branchRows, "achYtd");

  const stfWitelRows = seedData.stfWitelTargets.map((target) =>
    makeRow(
      target.label,
      target,
      witelCounts.get(target.label) ?? Array(period.days).fill(0),
      seedData.previousByWitel as PreviousConfig,
      referenceMtd(target.label, "stfWitel"),
      period.monthIndex,
    ),
  );
  rankDescending(stfWitelRows, "targetAch");
  rankDescending(stfWitelRows, "achYtd");

  const stfBranchRows = makeBranchRows(seedData.stfTeldaTargets, datelCounts, "stfBranch", period);

  const witelTotal = makeTotal("Total", witelRows, period.days);
  const branchTotal = makeTotal("YK JATENG SELATAN", branchRows, period.days);
  const stfWitelTotal = makeTotal("Total", stfWitelRows, period.days);
  const stfBranchTotal = makeTotal("YK JATENG SELATAN", stfBranchRows, period.days);
  const trend = Array.from({ length: period.days }, (_, index) => {
    const day = index + 1;
    const realisasi = witelTotal.days[index] ?? 0;
    const cumulative = witelTotal.days.slice(0, index + 1).reduce<number>((sum, value) => sum + (value ?? 0), 0);
    return { day, realisasi, cumulative };
  });
  const shouldValidate = period.year === DEFAULT_PERIOD.year && period.month === DEFAULT_PERIOD.month;
  const v1 = shouldValidate ? validation([...witelRows, witelTotal], "witel") : { checked: 0, mismatches: 0, notes: [] };
  const v2 = shouldValidate ? validation([...branchRows, branchTotal], "branch") : { checked: 0, mismatches: 0, notes: [] };
  const v3 = shouldValidate ? validation([...stfWitelRows, stfWitelTotal], "stfWitel") : { checked: 0, mismatches: 0, notes: [] };
  const v4 = shouldValidate ? validation([...stfBranchRows, stfBranchTotal], "stfBranch") : { checked: 0, mismatches: 0, notes: [] };

  return {
    periodLabel: `${period.monthName} ${period.year}`,
    previousMonthLabel: `${period.previousMonthName} ${period.month <= 1 ? period.year - 1 : period.year}`,
    cutoffLabel: savedAt ? new Date(savedAt).toLocaleString("id-ID", { dateStyle: "medium", timeStyle: "short" }) : seedData.initialCutoff,
    updatedAt: savedAt,
    source: seedData.source,
    witel: {
      title: "PS HARIAN HSI",
      subtitle: "WITEL YOGYA JATENG SELATAN",
      rows: witelRows,
      total: witelTotal,
    },
    branch: {
      title: "PS HARIAN HSI",
      subtitle: "WITEL YOGYA JATENG SELATAN",
      rows: branchRows,
      total: branchTotal,
    },
    targetStf: {
      witel: {
        title: "PS HARIAN HSI",
        subtitle: "WITEL YOGYA JATENG SELATAN",
        rows: stfWitelRows,
        total: stfWitelTotal,
      },
      branch: {
        title: "PS HARIAN HSI",
        subtitle: "WITEL YOGYA JATENG SELATAN",
        rows: stfBranchRows,
        total: stfBranchTotal,
      },
    },
    dailyTrend: trend,
    validation: {
      checked: v1.checked + v2.checked + v3.checked + v4.checked,
      mismatches: v1.mismatches + v2.mismatches + v3.mismatches + v4.mismatches,
      notes: [...v1.notes, ...v2.notes, ...v3.notes, ...v4.notes].slice(0, 8),
    },
  };
}

export function initialMasterRows(): MasterRow[] {
  return seedData.masterRows.map((row) => [...row]) as MasterRow[];
}

export const masterHeaders = [...seedData.masterHeaders] as string[];
