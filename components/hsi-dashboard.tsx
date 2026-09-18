"use client";

import {
  Bar,
  BarChart,
  CartesianGrid,
  LabelList,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import {
  Activity,
  Database,
  Download,
  FileSpreadsheet,
  Home,
  LogOut,
  Save,
  ShieldCheck,
  Target,
  Trash2,
  Undo2,
  Upload,
} from "lucide-react";
import { useRouter } from "next/navigation";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import { seedData } from "@/lib/seed-data";
import {
  computeReport,
  initialMasterRows,
  masterHeaders,
  type ComputedReport,
  type MasterRow,
  type ReportRow,
  type ReportTable,
} from "@/lib/reporting";

const STORAGE_KEY = "hsi-master-rows-v1";
const SAVED_AT_KEY = "hsi-master-saved-at-v1";
const REPORT_HISTORY_KEY = "hsi-report-history-v1";
const PERIOD_ARCHIVE_KEY = "hsi-period-archive-v1";
const SELECTED_PERIOD_KEY = "hsi-selected-period-v1";
const DATA_VERSION_KEY = "hsi-seed-data-version-v1";
const SAVE_SNAPSHOTS_KEY = "hsi-save-snapshots-v1";
const MASTER_HEADERS_CLEARED_KEY = "hsi-master-headers-cleared-v1";
const AUTO_SAVE_DELAY_MS = 3 * 60 * 1000;

type View = "home" | "master" | "report" | "stf";

type ReportHistoryTable = Record<string, (number | null)[]>;
type ReportHistory = {
  branch: ReportHistoryTable;
  stfBranch: ReportHistoryTable;
  stfWitel: ReportHistoryTable;
  updatedAt: string | null;
  witel: ReportHistoryTable;
};
type PeriodArchiveEntry = {
  label: string;
  periodKey: string;
  reportSnapshot: ComputedReport | null;
  reportHistory: ReportHistory | null;
  rows: MasterRow[];
  savedAt: string | null;
  updatedAt: string | null;
};
type PeriodArchive = Record<string, PeriodArchiveEntry>;
type SavedVersion = {
  periodKey: string;
  reportHistory: ReportHistory | null;
  rows: MasterRow[];
  savedAt: string | null;
};
type SaveSnapshots = Record<string, SavedVersion[]>;

type SheetJsWorkbook = {
  SheetNames: string[];
  Sheets: Record<string, unknown>;
};
type SheetJsApi = {
  read: (data: ArrayBuffer, options?: { cellDates?: boolean; type?: "array" }) => SheetJsWorkbook;
  utils: {
    sheet_to_json: (sheet: unknown, options?: { blankrows?: boolean; defval?: null; header?: 1; raw?: boolean }) => unknown[][];
  };
};

declare global {
  interface Window {
    XLSX?: SheetJsApi;
  }
}

const navItems: { id: View; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
  { id: "home", label: "Home", icon: Home },
  { id: "master", label: "MASTER", icon: Database },
  { id: "report", label: "Report Sementara SEPT 2026", icon: FileSpreadsheet },
  { id: "stf", label: "Report HSI Target STF SEPT", icon: Target },
];

const panelClass =
  "rounded-lg border border-white/70 bg-white/90 shadow-[0_18px_50px_-36px_rgba(15,23,42,0.55)] ring-1 ring-slate-950/[0.03] backdrop-blur";
const primaryButtonClass =
  "inline-flex h-10 items-center gap-2 rounded-md bg-slate-950 px-4 text-sm font-semibold text-white shadow-[0_14px_28px_-18px_rgba(15,23,42,0.9)] transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300";
const secondaryButtonClass =
  "inline-flex h-10 items-center gap-2 rounded-md border border-slate-200/80 bg-white/85 px-3 text-sm font-medium text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-white disabled:cursor-not-allowed disabled:border-slate-100 disabled:bg-slate-100 disabled:text-slate-400 disabled:shadow-none";
const dangerButtonClass =
  "inline-flex h-10 items-center gap-2 rounded-md border border-rose-200 bg-white/90 px-3 text-sm font-medium text-rose-700 shadow-sm transition hover:bg-rose-50";
const chartTooltipStyle = {
  background: "rgba(255,255,255,0.96)",
  border: "1px solid rgba(203,213,225,0.75)",
  borderRadius: 8,
  boxShadow: "0 18px 42px -26px rgba(15,23,42,0.45)",
} as const;
const reportCanvasColumns = {
  day: 28,
  label: 168,
  summary: 78,
  target: 72,
} as const;
const nperColumnIndex = masterHeaders.indexOf("NPER");
const fallbackPeriodKey = "202609";
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

function formatNumber(value: number | null | undefined) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return new Intl.NumberFormat("id-ID", { maximumFractionDigits: 0 }).format(value);
}

function formatPercent(value: number | null | undefined) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return new Intl.NumberFormat("id-ID", {
    style: "percent",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(value);
}

function formatCell(value: unknown) {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return String(value);
  return String(value);
}

function tooltipNumber(value: unknown) {
  return typeof value === "number" ? formatNumber(value) : formatNumber(Number(value ?? 0));
}

function tooltipPercent(value: unknown) {
  const numeric = typeof value === "number" ? value : Number(value ?? 0);
  return `${numeric.toFixed(2)}%`;
}

function parseClipboard(text: string) {
  return text
    .replace(/\r/g, "")
    .split("\n")
    .filter((line, index, lines) => !(index === lines.length - 1 && line === ""))
    .map((line) => line.split("\t"));
}

function ensureRowWidth(row: unknown[]): MasterRow {
  return Array.from({ length: masterHeaders.length }, (_, index) => {
    const value = row[index];
    return value === undefined || value === "" ? null : (value as MasterRow[number]);
  });
}

function formatWorkbookDate(value: Date) {
  const pad = (input: number) => String(input).padStart(2, "0");
  return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(
    value.getMinutes(),
  )}:${pad(value.getSeconds())}`;
}

function cleanWorkbookCell(value: unknown): MasterRow[number] {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) return formatWorkbookDate(value);
  if (typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
}

function rowsIncludeMasterHeader(rows: unknown[][]) {
  const first = rows[0];
  if (!first) return false;
  const normalizedFirst = first.slice(0, masterHeaders.length).map((item) => String(item ?? "").trim().toUpperCase());
  const normalizedHeaders = masterHeaders.map((item) => item.toUpperCase());
  return normalizedHeaders.every((header, index) => normalizedFirst[index] === header);
}

function parseDelimitedRows(text: string, delimiter: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((item) => item.trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell);
  if (row.some((item) => item.trim() !== "")) rows.push(row);
  return rows;
}

function readDelimitedMasterRows(buffer: ArrayBuffer) {
  const text = new TextDecoder("utf-8").decode(buffer).replace(/^\uFEFF/, "");
  if (!text.slice(0, 2000).includes("#")) return null;
  const parsed = parseDelimitedRows(text, "#");
  if (!rowsIncludeMasterHeader(parsed) && parsed[0]?.length !== masterHeaders.length) return null;
  const dataRows = rowsIncludeMasterHeader(parsed) ? parsed.slice(1) : parsed;
  const rows = dataRows
    .map((row) => ensureRowWidth(row.map(cleanWorkbookCell)))
    .filter((row) => row.some((cell) => cell !== null && String(cell).trim() !== ""));
  return rows.length > 0 ? rows : null;
}

async function loadSheetJs() {
  if (window.XLSX) return window.XLSX;
  await new Promise<void>((resolve, reject) => {
    const existing = document.querySelector<HTMLScriptElement>('script[data-sheetjs="true"]');
    if (existing) {
      existing.addEventListener("load", () => resolve(), { once: true });
      existing.addEventListener("error", () => reject(new Error("Gagal memuat parser Excel.")), { once: true });
      return;
    }
    const script = document.createElement("script");
    script.src = "/vendor/xlsx.full.min.js";
    script.async = true;
    script.dataset.sheetjs = "true";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Gagal memuat parser Excel."));
    document.head.appendChild(script);
  });
  if (!window.XLSX) throw new Error("Parser Excel belum siap.");
  return window.XLSX;
}

async function readMasterRowsFromWorkbook(file: File) {
  const buffer = await file.arrayBuffer();
  const delimitedRows = readDelimitedMasterRows(buffer);
  if (delimitedRows) return delimitedRows;

  const XLSX = await loadSheetJs();
  const workbook = XLSX.read(buffer, { cellDates: true, type: "array" });
  const sheetName =
    workbook.SheetNames.find((name) => name.trim().toUpperCase() === "MASTER") ?? workbook.SheetNames[0];
  if (!sheetName) throw new Error("Workbook tidak memiliki sheet yang bisa dibaca.");
  const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    blankrows: false,
    defval: null,
    header: 1,
    raw: true,
  });
  const dataRows = rowsIncludeMasterHeader(sheetRows) ? sheetRows.slice(1) : sheetRows;
  const rows = dataRows
    .map((row) => ensureRowWidth(row.map(cleanWorkbookCell)))
    .filter((row) => row.some((cell) => cell !== null && String(cell).trim() !== ""));
  if (rows.length === 0) throw new Error(`Sheet ${sheetName} tidak berisi data MASTER.`);
  return rows;
}

function StatusPill({ children, tone = "neutral" }: { children: React.ReactNode; tone?: "neutral" | "good" | "warn" }) {
  const toneClass =
    tone === "good"
      ? "border-emerald-200/80 bg-emerald-50/90 text-emerald-700"
      : tone === "warn"
        ? "border-amber-200/80 bg-amber-50/90 text-amber-800"
        : "border-slate-200/80 bg-white/80 text-slate-600";
  return (
    <span className={`inline-flex items-center rounded-full border px-3 py-1 text-xs font-semibold shadow-sm backdrop-blur ${toneClass}`}>
      {children}
    </span>
  );
}

function formatJakartaDateTime(date: Date, seconds = true) {
  const dateText = new Intl.DateTimeFormat("id-ID", {
    day: "numeric",
    month: "long",
    timeZone: "Asia/Jakarta",
    year: "numeric",
  }).format(date);
  const timeText = new Intl.DateTimeFormat("id-ID", {
    hour: "2-digit",
    hour12: false,
    minute: "2-digit",
    second: seconds ? "2-digit" : undefined,
    timeZone: "Asia/Jakarta",
  })
    .format(date)
    .replace(/\./g, ":");
  return `${dateText}${seconds ? "," : " •"} ${timeText} WIB`;
}

function parsePeriodKey(value: unknown) {
  const digits = String(value ?? "").replace(/\D/g, "");
  if (!/^\d{6}$/.test(digits)) return null;
  const month = Number(digits.slice(4, 6));
  return month >= 1 && month <= 12 ? digits : null;
}

function currentJakartaPeriodKey(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-CA", {
    month: "2-digit",
    timeZone: "Asia/Jakarta",
    year: "numeric",
  })
    .formatToParts(date)
    .reduce<Record<string, string>>((acc, part) => {
      if (part.type !== "literal") acc[part.type] = part.value;
      return acc;
    }, {});
  return `${parts.year}${parts.month}`;
}

function workbookSeedPeriodKey() {
  return `${seedData.period.year}${String(seedData.period.month).padStart(2, "0")}`;
}

function blankMasterRows() {
  return Array.from({ length: 25 }, () => Array(masterHeaders.length).fill(null) as MasterRow);
}

function detectPeriodKey(rows: MasterRow[], fallback = fallbackPeriodKey) {
  const counts = new Map<string, number>();
  if (nperColumnIndex < 0) return fallback;
  for (const row of rows) {
    const key = parsePeriodKey(row[nperColumnIndex]);
    if (key) counts.set(key, (counts.get(key) ?? 0) + 1);
  }
  return [...counts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? fallback;
}

function periodLabelFromKey(periodKey: string) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  return `${monthNames[month - 1] ?? "SEPTEMBER"} ${Number.isFinite(year) ? year : 2026}`;
}

function previousMonthLabelFromKey(periodKey: string) {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  const previousMonth = month <= 1 ? 12 : month - 1;
  const previousYear = month <= 1 ? year - 1 : year;
  return `${shortMonthNames[previousMonth - 1] ?? "AUG"} ${Number.isFinite(previousYear) ? previousYear : 2026}`;
}

function formatSavedAt(value: string | null) {
  return value ? formatJakartaDateTime(new Date(value), true) : "Belum pernah disimpan";
}

function reportFileName(date: Date, variant: "sementara" | "stf") {
  const parts = new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    hour: "2-digit",
    hour12: false,
    minute: "2-digit",
    month: "short",
    timeZone: "Asia/Jakarta",
    year: "numeric",
  })
    .formatToParts(date)
    .reduce<Record<string, string>>((acc, part) => {
      if (part.type !== "literal") acc[part.type] = part.value;
      return acc;
    }, {});
  const prefix = variant === "stf" ? "HSI-Target-STF-Report" : "HSI-Report";
  return `${prefix}-${parts.day}-${parts.month}-${parts.year}-${parts.hour}${parts.minute}-WIB.png`;
}

function useJakartaClock() {
  const [now, setNow] = useState(() => new Date());
  useEffect(() => {
    const interval = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(interval);
  }, []);
  return now;
}

function latestPopulatedDay(tables: ReportTable[]) {
  let latest = 0;
  for (const table of tables) {
    for (const row of [...table.rows, table.total]) {
      row.days.forEach((value, index) => {
        if (typeof value === "number" && value > 0) latest = Math.max(latest, index + 1);
      });
    }
  }
  return Math.max(1, latest);
}

function emptyReportHistory(): ReportHistory {
  return { branch: {}, stfBranch: {}, stfWitel: {}, updatedAt: null, witel: {} };
}

function hasActualValue(value: number | null | undefined) {
  return typeof value === "number" && Number.isFinite(value) && value > 0;
}

function normalizeHistoryDays(value: unknown, fallbackLength: number) {
  const source = Array.isArray(value) ? value : [];
  return Array.from({ length: fallbackLength }, (_, index) => {
    const numeric = Number(source[index]);
    return Number.isFinite(numeric) && numeric > 0 ? numeric : null;
  });
}

function parseReportHistory(value: string | null): ReportHistory | null {
  if (!value) return null;
  try {
    const parsed = JSON.parse(value) as Partial<ReportHistory>;
    return {
      branch: parsed.branch && typeof parsed.branch === "object" ? parsed.branch : {},
      stfBranch: parsed.stfBranch && typeof parsed.stfBranch === "object" ? parsed.stfBranch : {},
      stfWitel: parsed.stfWitel && typeof parsed.stfWitel === "object" ? parsed.stfWitel : {},
      updatedAt: typeof parsed.updatedAt === "string" ? parsed.updatedAt : null,
      witel: parsed.witel && typeof parsed.witel === "object" ? parsed.witel : {},
    };
  } catch {
    return null;
  }
}

function parsePeriodArchive(value: string | null): PeriodArchive {
  if (!value) return {};
  try {
    const parsed = JSON.parse(value) as Record<string, Partial<PeriodArchiveEntry>>;
    return Object.fromEntries(
      Object.entries(parsed).flatMap(([key, entry]) => {
        const periodKey = parsePeriodKey(entry.periodKey ?? key);
        if (!periodKey || !Array.isArray(entry.rows)) return [];
        const reportHistory =
          entry.reportHistory && typeof entry.reportHistory === "object"
            ? ({
                branch: entry.reportHistory.branch ?? {},
                stfBranch: entry.reportHistory.stfBranch ?? {},
                stfWitel: entry.reportHistory.stfWitel ?? {},
                updatedAt: typeof entry.reportHistory.updatedAt === "string" ? entry.reportHistory.updatedAt : null,
                witel: entry.reportHistory.witel ?? {},
              } satisfies ReportHistory)
            : null;
        return [
          [
            periodKey,
            {
              label: typeof entry.label === "string" ? entry.label : periodLabelFromKey(periodKey),
              periodKey,
              reportSnapshot: entry.reportSnapshot ? (entry.reportSnapshot as ComputedReport) : null,
              reportHistory,
              rows: entry.rows.map((row) => ensureRowWidth(Array.isArray(row) ? row : [])),
              savedAt: typeof entry.savedAt === "string" ? entry.savedAt : null,
              updatedAt: typeof entry.updatedAt === "string" ? entry.updatedAt : null,
            },
          ],
        ];
      }),
    );
  } catch {
    return {};
  }
}

function parseSaveSnapshots(value: string | null): SaveSnapshots {
  if (!value) return {};
  try {
    const parsed = JSON.parse(value) as Record<string, Partial<SavedVersion>[]>;
    return Object.fromEntries(
      Object.entries(parsed).map(([periodKey, versions]) => [
        periodKey,
        Array.isArray(versions)
          ? versions
              .filter((version) => Array.isArray(version.rows))
              .map((version) => ({
                periodKey: parsePeriodKey(version.periodKey) ?? periodKey,
                reportHistory:
                  version.reportHistory && typeof version.reportHistory === "object"
                    ? (version.reportHistory as ReportHistory)
                    : null,
                rows: (version.rows ?? []).map((row) => ensureRowWidth(Array.isArray(row) ? row : [])),
                savedAt: typeof version.savedAt === "string" ? version.savedAt : null,
              }))
          : [],
      ]),
    );
  } catch {
    return {};
  }
}

function createPeriodEntry(
  periodKey: string,
  rows: MasterRow[],
  savedAt: string | null,
  reportHistory: ReportHistory | null,
  reportSnapshot: ComputedReport | null = null,
): PeriodArchiveEntry {
  return {
    label: periodLabelFromKey(periodKey),
    periodKey,
    reportSnapshot,
    reportHistory,
    rows,
    savedAt,
    updatedAt: savedAt,
  };
}

function initialSeedPeriodArchive(): PeriodArchive {
  return Object.fromEntries(
    Object.entries(seedData.historicalReports ?? {}).map(([periodKey, snapshot]) => [
      periodKey,
      createPeriodEntry(
        periodKey,
        blankMasterRows(),
        typeof snapshot.updatedAt === "string" ? snapshot.updatedAt : null,
        null,
        snapshot as unknown as ComputedReport,
      ),
    ]),
  );
}

function snapshotTable(table: ReportTable): ReportHistoryTable {
  return Object.fromEntries(
    [...table.rows, table.total].map((row) => [
      row.label,
      row.days.map((value) => (hasActualValue(value) ? value : null)),
    ]),
  );
}

function createReportHistory(report: ComputedReport, updatedAt: string | null): ReportHistory {
  return {
    branch: snapshotTable(report.branch),
    stfBranch: snapshotTable(report.targetStf.branch),
    stfWitel: snapshotTable(report.targetStf.witel),
    updatedAt,
    witel: snapshotTable(report.witel),
  };
}

function mergeHistoryTable(previous: ReportHistoryTable, table: ReportTable) {
  const next: ReportHistoryTable = { ...previous };
  for (const row of [...table.rows, table.total]) {
    const stored = normalizeHistoryDays(next[row.label], row.days.length);
    next[row.label] = row.days.map((value, index) => (hasActualValue(value) ? value : stored[index]));
  }
  return next;
}

function mergeReportHistory(previous: ReportHistory | null, report: ComputedReport, updatedAt: string | null): ReportHistory {
  const base = previous ?? emptyReportHistory();
  return {
    branch: mergeHistoryTable(base.branch, report.branch),
    stfBranch: mergeHistoryTable(base.stfBranch, report.targetStf.branch),
    stfWitel: mergeHistoryTable(base.stfWitel, report.targetStf.witel),
    updatedAt,
    witel: mergeHistoryTable(base.witel, report.witel),
  };
}

function safeRatioValue(numerator: number, denominator: number) {
  return denominator === 0 ? null : numerator / denominator;
}

function safeGrowthValue(current: number, previous: number) {
  return previous === 0 ? null : current / previous - 1;
}

function previousFromGrowth(current: number, growth: number | null) {
  if (growth === null || growth <= -1) return 0;
  return current / (growth + 1);
}

function recalcRowFromDays(row: ReportRow, days: (number | null)[]): ReportRow {
  const mtd = days.reduce<number>((sum, value) => sum + (value ?? 0), 0);
  const currentMonthYtdBase = row.realYtd - row.mtd;
  const ytdPrevious = previousFromGrowth(row.realYtd, row.growthYtd);
  const realYtd = currentMonthYtdBase + mtd;

  return {
    ...row,
    achYtd: safeRatioValue(realYtd, row.targetYtd),
    days,
    growthFm: safeGrowthValue(mtd, row.realFmPrev),
    growthMtd: safeGrowthValue(mtd, row.realMtdPrev),
    growthYtd: safeGrowthValue(realYtd, ytdPrevious),
    mtd,
    realYtd,
    shortageFm: mtd - row.targetFm,
    shortageYtd: realYtd - row.targetYtd,
    targetAch: safeRatioValue(mtd, row.targetFm),
  };
}

function rankReportRows(rows: ReportRow[], metric: "targetAch" | "achYtd"): ReportRow[] {
  return rows.map((row) => {
    const value = row[metric];
    const next = { ...row };
    if (metric === "targetAch") {
      next.rankAch = value === null ? null : rows.filter((candidate) => (candidate.targetAch ?? -Infinity) > value).length + 1;
    } else {
      next.rankYtd = value === null ? null : rows.filter((candidate) => (candidate.achYtd ?? -Infinity) > value).length + 1;
    }
    return next;
  });
}

function recalcTotalFromRows(label: string, rows: ReportRow[], originalTotal: ReportRow): ReportRow {
  const days = Array.from({ length: originalTotal.days.length }, (_, day) =>
    rows.reduce<number>((sum, row) => sum + (row.days[day] ?? 0), 0),
  );
  const mtd = rows.reduce<number>((sum, row) => sum + row.mtd, 0);
  const realYtd = rows.reduce<number>((sum, row) => sum + row.realYtd, 0);
  const ytdPrevious = previousFromGrowth(originalTotal.realYtd, originalTotal.growthYtd);

  return {
    ...originalTotal,
    achYtd: safeRatioValue(realYtd, originalTotal.targetYtd),
    days,
    growthFm: safeGrowthValue(mtd, originalTotal.realFmPrev),
    growthMtd: safeGrowthValue(mtd, originalTotal.realMtdPrev),
    growthYtd: safeGrowthValue(realYtd, ytdPrevious),
    label,
    mtd,
    realYtd,
    shortageFm: mtd - originalTotal.targetFm,
    shortageYtd: realYtd - originalTotal.targetYtd,
    targetAch: safeRatioValue(mtd, originalTotal.targetFm),
  };
}

function applyHistoryTable(table: ReportTable, history: ReportHistoryTable): ReportTable {
  let rows = table.rows.map((row) => {
    const stored = normalizeHistoryDays(history[row.label], row.days.length);
    const days = row.days.map((value, index) => (hasActualValue(value) ? value : stored[index]));
    return recalcRowFromDays(row, days);
  });
  rows = rankReportRows(rankReportRows(rows, "targetAch"), "achYtd");
  return {
    ...table,
    rows,
    total: recalcTotalFromRows(table.total.label, rows, table.total),
  };
}

function applyReportHistory(report: ComputedReport, history: ReportHistory | null): ComputedReport {
  if (!history) return report;
  const witel = applyHistoryTable(report.witel, history.witel);
  const branch = applyHistoryTable(report.branch, history.branch);
  const stfWitel = applyHistoryTable(report.targetStf.witel, history.stfWitel);
  const stfBranch = applyHistoryTable(report.targetStf.branch, history.stfBranch);
  const dailyTrend = witel.total.days.map((value, index) => ({
    cumulative: witel.total.days.slice(0, index + 1).reduce<number>((sum, dayValue) => sum + (dayValue ?? 0), 0),
    day: index + 1,
    realisasi: value ?? 0,
  }));

  return {
    ...report,
    branch,
    dailyTrend,
    targetStf: {
      branch: stfBranch,
      witel: stfWitel,
    },
    witel,
  };
}

function achievementTone(value: number | null) {
  if (value === null) return "neutral";
  if (value >= 1) return "good";
  if (value >= 0.8) return "near";
  return "low";
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function MetricCard({
  label,
  value,
  detail,
  tone = "neutral",
  icon: Icon,
}: {
  label: string;
  value: string;
  detail: string;
  tone?: "neutral" | "good" | "warn";
  icon?: React.ComponentType<{ className?: string }>;
}) {
  const toneClass =
    tone === "good"
      ? "from-emerald-50/90 to-white text-emerald-700"
      : tone === "warn"
        ? "from-amber-50/90 to-white text-amber-700"
        : "from-slate-50/90 to-white text-slate-600";
  const iconClass =
    tone === "good"
      ? "border-emerald-200 bg-emerald-50 text-emerald-700"
      : tone === "warn"
        ? "border-amber-200 bg-amber-50 text-amber-700"
        : "border-slate-200 bg-white text-slate-600";
  return (
    <section className={`${panelClass} overflow-hidden bg-gradient-to-br ${toneClass} p-4 transition hover:-translate-y-0.5 hover:shadow-[0_24px_60px_-38px_rgba(15,23,42,0.75)]`}>
      <div className="flex items-start justify-between gap-3">
        <p className="text-xs font-semibold uppercase text-slate-500">{label}</p>
        {Icon && (
          <span className={`flex h-8 w-8 items-center justify-center rounded-md border ${iconClass}`}>
            <Icon className="h-4 w-4" />
          </span>
        )}
      </div>
      <p className="mt-3 text-2xl font-semibold tracking-normal text-slate-950">{value}</p>
      <p className={`mt-2 text-sm ${tone === "good" ? "text-emerald-700" : tone === "warn" ? "text-amber-700" : "text-slate-500"}`}>
        {detail}
      </p>
    </section>
  );
}

function shortPeriodLabel(periodLabel: string) {
  const [name, year] = periodLabel.split(" ");
  const idx = monthNames.indexOf(name);
  const short = idx >= 0 ? shortMonthNames[idx] : (name ?? "").slice(0, 4);
  return `${short} ${year ?? ""}`.trim();
}

function shortageValue(row: { targetFm: number; mtd: number }) {
  return Math.max(0, row.targetFm - row.mtd);
}

function safeAchievement(realized: number, target: number): number | null {
  return target === 0 ? null : realized / target;
}

const OWN_WITEL_LABEL = "YOGYA JATENG SELATAN";

function InlineStat({ label, value, tone = "neutral" }: { label: string; value: string; tone?: "neutral" | "good" | "warn" }) {
  const valueClass = tone === "good" ? "text-emerald-700" : tone === "warn" ? "text-amber-700" : "text-slate-950";
  return (
    <div className="min-w-[7.5rem] flex-1">
      <p className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">{label}</p>
      <p className={`mt-1 text-base font-semibold ${valueClass}`}>{value}</p>
    </div>
  );
}

function TeldaTable({ rows, teldaFilter }: { rows: ReportRow[]; teldaFilter: string }) {
  const visibleRows = teldaFilter === "ALL" ? rows : rows.filter((row) => row.label === teldaFilter);
  return (
    <div className="overflow-x-auto">
      <table className="w-full min-w-[720px] border-collapse text-sm">
        <thead>
          <tr className="border-b border-slate-200 text-left text-[11px] font-semibold uppercase tracking-wide text-slate-500">
            <th className="py-2 pr-3">TELDA</th>
            <th className="py-2 pr-3 text-right">TARGET</th>
            <th className="py-2 pr-3 text-right">REALISASI</th>
            <th className="py-2 pr-3 text-right">ACHIEVEMENT</th>
            <th className="py-2 pr-3 text-right">SHORTAGE</th>
            <th className="py-2 pr-3 text-right">GROWTH</th>
            <th className="py-2 pr-3 text-right">RANKING</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-slate-100">
          {visibleRows.map((row) => (
            <tr key={row.label} className={teldaFilter === row.label ? "bg-rose-50/70" : undefined}>
              <td className="py-2 pr-3 font-medium text-slate-800">{row.label}</td>
              <td className="py-2 pr-3 text-right text-slate-700">{formatNumber(row.targetFm)}</td>
              <td className="py-2 pr-3 text-right text-slate-700">{formatNumber(row.mtd)}</td>
              <td className="py-2 pr-3 text-right font-semibold text-rose-700">{formatPercent(row.targetAch)}</td>
              <td className="py-2 pr-3 text-right text-amber-700">{formatNumber(shortageValue(row))}</td>
              <td className="py-2 pr-3 text-right text-slate-700">{row.growthMtd === null ? "-" : formatPercent(row.growthMtd)}</td>
              <td className="py-2 pr-3 text-right font-semibold text-slate-950">{row.rankAch ?? "-"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function DashboardHome({ report }: { report: ComputedReport }) {
  const [rkapStf, setRkapStf] = useState<"RKAP" | "STF">("RKAP");
  const [teldaFilter, setTeldaFilter] = useState<string>("ALL");

  const regionalRkap = report.witel;
  const regionalStf = report.targetStf.witel;
  const regionalSource = rkapStf === "RKAP" ? regionalRkap : regionalStf;

  const teldaRkap = report.branch;
  const teldaStf = report.targetStf.branch;
  const teldaSource = rkapStf === "RKAP" ? teldaRkap : teldaStf;

  const witelOwnRkap = regionalRkap.rows.find((row) => row.label === OWN_WITEL_LABEL) ?? regionalRkap.total;
  const witelOwnStf = regionalStf.rows.find((row) => row.label === OWN_WITEL_LABEL) ?? regionalStf.total;

  const latestDay = latestPopulatedDay([report.witel, report.branch]);
  const periodShort = shortPeriodLabel(report.periodLabel);

  const targetChart = regionalRkap.rows.map((row) => {
    const stfRow = regionalStf.rows.find((item) => item.label === row.label);
    const activeTarget = rkapStf === "RKAP" ? row.targetFm : stfRow?.targetFm ?? 0;
    const activeMtd = rkapStf === "RKAP" ? row.mtd : stfRow?.mtd ?? 0;
    const activeAch = rkapStf === "RKAP" ? row.targetAch : stfRow?.targetAch ?? null;
    return {
      name: row.label.replace(" JATENG ", " "),
      target: activeTarget,
      realisasi: activeMtd,
      shortage: Math.max(0, activeTarget - activeMtd),
      achievement: activeAch === null ? 0 : Number((activeAch * 100).toFixed(1)),
    };
  });

  const rkapBarChart = teldaRkap.rows.map((row) => ({
    name: row.label,
    target: row.targetFm,
    realisasi: row.mtd,
    achievement: row.targetAch === null ? 0 : Number((row.targetAch * 100).toFixed(1)),
    rank: row.rankAch,
  }));
  const stfBarChart = teldaStf.rows.map((row) => ({
    name: row.label,
    target: row.targetFm,
    realisasi: row.mtd,
    achievement: row.targetAch === null ? 0 : Number((row.targetAch * 100).toFixed(1)),
    rank: row.rankAch,
  }));

  const trendTotal = rkapStf === "RKAP" ? report.witel.total : report.targetStf.witel.total;
  const trendDayCount = trendTotal.days.length;
  const dailyTargetValue = trendDayCount > 0 ? trendTotal.targetFm / trendDayCount : 0;
  let runningCumulative = 0;
  const trendData = trendTotal.days.map((value, index) => {
    const realisasiHarian = value ?? 0;
    runningCumulative += realisasiHarian;
    return {
      day: index + 1,
      realisasiHarian,
      targetHarian: dailyTargetValue,
      cumulative: runningCumulative,
    };
  });

  const cardWitelRow = teldaFilter === "ALL" ? teldaRkap.total : teldaRkap.rows.find((row) => row.label === teldaFilter) ?? teldaRkap.total;
  const cardStfRow = teldaFilter === "ALL" ? teldaStf.total : teldaStf.rows.find((row) => row.label === teldaFilter) ?? teldaStf.total;
  const cardTitle =
    teldaFilter === "ALL" ? "HSI PERFORMA TELKOM JOGJA JATENG SELATAN" : `HSI PERFORMA TELDA ${teldaFilter}`;

  return (
    <div className="space-y-6">
      {/* HEADER */}
      <section className={`${panelClass} overflow-hidden`}>
        <div className="relative p-5 lg:p-6">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <p className="inline-flex items-center gap-2 text-xs font-semibold uppercase text-rose-700">
                <Activity className="h-4 w-4" />
                PS HARIAN HSI
              </p>
              <h1 className="mt-2 text-2xl font-bold uppercase tracking-tight text-slate-950 sm:text-3xl">
                HSI PERFORMA MTD {periodShort}
              </h1>
              <p className="mt-1 text-sm text-slate-500">Cut-off {report.cutoffLabel} · Source: {report.source}</p>
            </div>
            <StatusPill tone={report.validation.mismatches === 0 ? "good" : "warn"}>
              Validasi: {report.validation.mismatches === 0 ? "match" : `${report.validation.mismatches} selisih`}
            </StatusPill>
          </div>

          {/* FILTER BAR */}
          <div className="mt-5 flex flex-wrap items-center gap-3 rounded-lg border border-slate-200/70 bg-white/70 p-3">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-500">HSI PERFORMANCE CURRENT MONTH {report.periodLabel.split(" ")[1]}</span>
            <label className="flex items-center gap-2 text-xs font-semibold text-slate-600">
              TELDA
              <select
                className="rounded-md border border-slate-200 bg-white px-2 py-1.5 text-sm font-medium text-slate-950 outline-none"
                value={teldaFilter}
                onChange={(event) => setTeldaFilter(event.target.value)}
              >
                <option value="ALL">SEMUA TELDA</option>
                {teldaRkap.rows.map((row) => (
                  <option key={row.label} value={row.label}>
                    {row.label}
                  </option>
                ))}
              </select>
            </label>
            <div className="inline-flex overflow-hidden rounded-md border border-slate-200">
              <button
                onClick={() => setRkapStf("RKAP")}
                className={`px-3 py-1.5 text-xs font-semibold uppercase transition ${rkapStf === "RKAP" ? "bg-slate-950 text-white" : "bg-white text-slate-600 hover:bg-slate-50"}`}
              >
                RKAP
              </button>
              <button
                onClick={() => setRkapStf("STF")}
                className={`px-3 py-1.5 text-xs font-semibold uppercase transition ${rkapStf === "STF" ? "bg-slate-950 text-white" : "bg-white text-slate-600 hover:bg-slate-50"}`}
              >
                STF
              </button>
            </div>
            <StatusPill>CURRENT MONTH: {report.periodLabel}</StatusPill>
          </div>
        </div>
      </section>

      {/* MTD SECTION */}
      <section className={`${panelClass} p-5`}>
        <div className="mb-3 flex items-center justify-between gap-3">
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">HSI PERFORMA MTD {periodShort}</h2>
          <StatusPill>Tanggal data {latestDay}</StatusPill>
        </div>
        <div className="flex flex-wrap gap-6">
          <InlineStat label="Target" value={formatNumber(regionalSource.total.targetFm)} />
          <InlineStat label="Realisasi MTD" value={formatNumber(regionalSource.total.mtd)} tone="good" />
          <InlineStat label="Shortage" value={formatNumber(shortageValue(regionalSource.total))} tone="warn" />
          <InlineStat label="Achievement" value={formatPercent(regionalSource.total.targetAch)} tone="good" />
          <InlineStat label="Growth vs Bulan Lalu" value={regionalSource.total.growthMtd === null ? "-" : formatPercent(regionalSource.total.growthMtd)} />
        </div>
      </section>

      {/* YTD SECTION */}
      <section className={`${panelClass} p-5`}>
        <div className="mb-3 flex items-center justify-between gap-3">
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">HSI PERFORMA YTD {periodShort}</h2>
          <StatusPill>Kumulatif Januari - {periodShort}</StatusPill>
        </div>
        <div className="flex flex-wrap gap-6">
          <InlineStat label="Target YTD" value={formatNumber(regionalSource.total.targetYtd)} />
          <InlineStat label="Realisasi YTD" value={formatNumber(regionalSource.total.realYtd)} tone="good" />
          <InlineStat label="Shortage YTD" value={formatNumber(Math.max(0, regionalSource.total.targetYtd - regionalSource.total.realYtd))} tone="warn" />
          <InlineStat label="Achievement YTD" value={formatPercent(regionalSource.total.achYtd)} tone="good" />
          <InlineStat label="Growth YTD" value={regionalSource.total.growthYtd === null ? "-" : formatPercent(regionalSource.total.growthYtd)} />
        </div>
      </section>

      {/* TARGET VS REALISASI */}
      <section className={`${panelClass} p-5`}>
        <div className="mb-4 flex items-center justify-between gap-3">
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">Target vs Realisasi ({rkapStf})</h2>
          <StatusPill>{report.periodLabel}</StatusPill>
        </div>
        <div className="h-[360px]">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={targetChart} margin={{ left: 0, right: 16, top: 24, bottom: 36 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="name" angle={-28} textAnchor="end" interval={0} height={80} tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip contentStyle={chartTooltipStyle} formatter={(value) => tooltipNumber(value)} />
              <Legend verticalAlign="top" height={28} />
              <Bar dataKey="target" name="Target" fill="#94a3b8" radius={[4, 4, 0, 0]}>
                <LabelList dataKey="target" position="top" formatter={(value: number) => formatNumber(value)} style={{ fontSize: 10, fill: "#475569" }} />
              </Bar>
              <Bar dataKey="realisasi" name="Realisasi" fill="#f97316" radius={[4, 4, 0, 0]}>
                <LabelList dataKey="realisasi" position="top" formatter={(value: number) => formatNumber(value)} style={{ fontSize: 10, fill: "#c2410c" }} />
              </Bar>
              <Bar dataKey="shortage" name="Shortage" fill="#f43f5e" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
        <div className="mt-3 flex flex-wrap items-center gap-4 text-xs text-slate-600">
          <span className="inline-flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-sm bg-slate-400" /> Abu-abu = Target</span>
          <span className="inline-flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-sm bg-orange-500" /> Oranye = Realisasi</span>
          <span className="inline-flex items-center gap-1.5"><span className="h-2.5 w-2.5 rounded-sm bg-rose-500" /> Merah = Shortage</span>
          <span className="text-slate-400">Regional: BALI · JTT · JTB · NUSRA · SCU · SJT · SURAMADU · JJS</span>
        </div>
        <div className="mt-4 flex flex-wrap gap-4 border-t border-slate-100 pt-3 text-sm">
          {targetChart.map((item) => (
            <span key={item.name} className="text-slate-600">
              <span className="font-semibold text-slate-900">{item.name}</span>: ACH {item.achievement.toFixed(1)}%
            </span>
          ))}
        </div>
      </section>

      {/* CARD KANAN - HSI PERFORMANCE */}
      <section className={`${panelClass} p-5`}>
        <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">{cardTitle}</h2>
        <div className="mt-4 flex flex-wrap items-center gap-x-10 gap-y-4">
          <InlineStat label="Achievement RKAP" value={formatPercent(cardWitelRow.targetAch)} tone="good" />
          <InlineStat label="Shortage STF" value={formatNumber(shortageValue(cardStfRow))} tone="warn" />
        </div>
        <p className="mt-4 text-xs text-slate-400">Source: {report.source}</p>
      </section>

      {/* PERFORMANCE BY RKAP 2026 (card kecil bawah) */}
      <section className={`${panelClass} p-5`}>
        <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">PERFORMANCE BY RKAP {report.periodLabel.split(" ")[1]}</h2>
        <div className="mt-4 flex flex-wrap gap-x-10 gap-y-4">
          <InlineStat label="Target RKAP" value={formatNumber(teldaRkap.total.targetFm)} />
          <InlineStat label="Realisasi MTD" value={formatNumber(teldaRkap.total.mtd)} tone="good" />
          <InlineStat label="Achievement RKAP" value={formatPercent(teldaRkap.total.targetAch)} tone="good" />
          <InlineStat label="Current Month" value={report.periodLabel} />
          <InlineStat label="Growth vs Last Month" value={teldaRkap.total.growthMtd === null ? "-" : formatPercent(teldaRkap.total.growthMtd)} />
          <InlineStat label="Ranking" value={String(witelOwnRkap.rankAch ?? "-")} />
        </div>
      </section>

      {/* DUA CHART UTAMA */}
      <div className="grid gap-5 xl:grid-cols-2">
        <section className={`${panelClass} p-5`}>
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">PERFORMANCE BY RKAP {report.periodLabel.split(" ")[1]}</h2>
          <div className="mt-4 h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={rkapBarChart} margin={{ left: 0, right: 16, top: 20, bottom: 36 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" angle={-28} textAnchor="end" interval={0} height={70} tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip contentStyle={chartTooltipStyle} formatter={(value) => tooltipNumber(value)} />
                <Legend verticalAlign="top" height={24} />
                <Bar dataKey="target" name="Target RKAP" fill="#94a3b8" radius={[4, 4, 0, 0]} />
                <Bar dataKey="realisasi" name="Realisasi" fill="#0f766e" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="achievement" position="top" formatter={(value: number) => `${value.toFixed(0)}%`} style={{ fontSize: 10, fill: "#0f766e" }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="mt-3 flex flex-wrap gap-x-4 gap-y-1 border-t border-slate-100 pt-3 text-xs text-slate-600">
            {[...rkapBarChart].sort((a, b) => (a.rank ?? 999) - (b.rank ?? 999)).map((item) => (
              <span key={item.name}>#{item.rank} {item.name}</span>
            ))}
          </div>
        </section>

        <section className={`${panelClass} p-5`}>
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">PERFORMANCE BY STF {report.periodLabel.split(" ")[1]}</h2>
          <div className="mt-4 h-[300px]">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={stfBarChart} margin={{ left: 0, right: 16, top: 20, bottom: 36 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="name" angle={-28} textAnchor="end" interval={0} height={70} tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 12 }} />
                <Tooltip contentStyle={chartTooltipStyle} formatter={(value) => tooltipNumber(value)} />
                <Legend verticalAlign="top" height={24} />
                <Bar dataKey="target" name="Target STF" fill="#94a3b8" radius={[4, 4, 0, 0]} />
                <Bar dataKey="realisasi" name="Realisasi" fill="#f59e0b" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="achievement" position="top" formatter={(value: number) => `${value.toFixed(0)}%`} style={{ fontSize: 10, fill: "#b45309" }} />
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
          <p className="mt-3 border-t border-slate-100 pt-3 text-xs text-slate-400">Filter mengikuti RKAP apabila data STF per TELDA tidak tersedia.</p>
        </section>
      </div>

      {/* TREND HARIAN */}
      <section className={`${panelClass} p-5`}>
        <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">Trend Target Harian vs Realisasi Harian ({rkapStf})</h2>
          <StatusPill>Achievement: {formatPercent(safeAchievement(runningCumulative, trendTotal.targetFm))}</StatusPill>
        </div>
        <div className="mb-3 flex flex-wrap gap-6 text-sm">
          <InlineStat label="Target Harian" value={formatNumber(dailyTargetValue)} />
          <InlineStat label="Realisasi MTD" value={formatNumber(runningCumulative)} tone="good" />
        </div>
        <div className="h-[300px]">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={trendData} margin={{ left: 0, right: 16, top: 8, bottom: 8 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="day" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip contentStyle={chartTooltipStyle} formatter={(value) => tooltipNumber(value)} labelFormatter={(day) => `Tanggal ${day}`} />
              <Legend verticalAlign="top" height={24} />
              <Line type="monotone" dataKey="targetHarian" name="Target Harian" stroke="#64748b" strokeWidth={2} strokeDasharray="5 4" dot={false} />
              <Line type="monotone" dataKey="realisasiHarian" name="Realisasi Harian" stroke="#be123c" strokeWidth={2} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </section>

      {/* WITEL PERFORMANCE */}
      <section className={`${panelClass} p-5`}>
        <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">WITEL Performance - {OWN_WITEL_LABEL}</h2>
        <div className="mt-4 flex flex-wrap gap-x-8 gap-y-4">
          <InlineStat label="Target" value={formatNumber(rkapStf === "RKAP" ? witelOwnRkap.targetFm : witelOwnStf.targetFm)} />
          <InlineStat label="Realisasi" value={formatNumber(rkapStf === "RKAP" ? witelOwnRkap.mtd : witelOwnStf.mtd)} tone="good" />
          <InlineStat label="Achievement" value={formatPercent(rkapStf === "RKAP" ? witelOwnRkap.targetAch : witelOwnStf.targetAch)} tone="good" />
          <InlineStat label="Shortage" value={formatNumber(shortageValue(rkapStf === "RKAP" ? witelOwnRkap : witelOwnStf))} tone="warn" />
          <InlineStat label="Growth" value={(rkapStf === "RKAP" ? witelOwnRkap.growthMtd : witelOwnStf.growthMtd) === null ? "-" : formatPercent(rkapStf === "RKAP" ? witelOwnRkap.growthMtd : witelOwnStf.growthMtd)} />
          <InlineStat label="Ranking" value={String((rkapStf === "RKAP" ? witelOwnRkap.rankAch : witelOwnStf.rankAch) ?? "-")} />
          <InlineStat label="MTD" value={formatNumber(rkapStf === "RKAP" ? witelOwnRkap.mtd : witelOwnStf.mtd)} />
          <InlineStat label="YTD" value={formatNumber(rkapStf === "RKAP" ? witelOwnRkap.realYtd : witelOwnStf.realYtd)} />
        </div>
      </section>

      {/* TELDA PERFORMANCE */}
      <section className={`${panelClass} p-5`}>
        <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
          <h2 className="text-base font-bold uppercase tracking-tight text-slate-950">TELDA Performance ({rkapStf})</h2>
          <StatusPill>{teldaFilter === "ALL" ? "Semua TELDA" : teldaFilter}</StatusPill>
        </div>
        <TeldaTable rows={teldaSource.rows} teldaFilter={teldaFilter} />
      </section>
    </div>
  );
}

function MasterPage({
  autoSaveStatus,
  canGoBack,
  headersCleared,
  periodLabel,
  rows,
  savedAt,
  setHeadersCleared,
  setRows,
  onBack,
  onSave,
  onClear,
}: {
  autoSaveStatus: string;
  canGoBack: boolean;
  headersCleared: boolean;
  periodLabel: string;
  rows: MasterRow[];
  savedAt: string | null;
  setHeadersCleared: React.Dispatch<React.SetStateAction<boolean>>;
  setRows: React.Dispatch<React.SetStateAction<MasterRow[]>>;
  onBack: () => void;
  onSave: (options?: { automated?: boolean; pushSnapshot?: boolean }) => void;
  onClear: () => void;
}) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [importError, setImportError] = useState<string | null>(null);
  const displayHeaders = headersCleared ? masterHeaders.map(() => "") : masterHeaders;

  function updateCell(rowIndex: number, colIndex: number, value: string) {
    setRows((current) => {
      const next = current.map((row) => [...row]);
      next[rowIndex][colIndex] = value === "" ? null : value;
      return next;
    });
  }

  function pasteAt(event: React.ClipboardEvent<HTMLTableSectionElement>) {
    const target = event.target as HTMLElement;
    const cell = target.closest<HTMLTableCellElement>("[data-row][data-col]");
    if (!cell) return;
    const text = event.clipboardData.getData("text/plain");
    const pasted = parseClipboard(text);
    if (pasted.length <= 1 && pasted[0]?.length <= 1) return;
    event.preventDefault();
    const startRow = Number(cell.dataset.row);
    const startCol = Number(cell.dataset.col);
    const pastedRows = startRow === 0 && startCol === 0 && rowsIncludeMasterHeader(pasted) ? pasted.slice(1) : pasted;
    setHeadersCleared(false);
    setRows((current) => {
      const next =
        startRow === 0 && startCol === 0
          ? pastedRows.map((row) => ensureRowWidth(row))
          : current.map((row) => [...row]);
      while (next.length < startRow + pastedRows.length) next.push(Array(masterHeaders.length).fill(null));
      pastedRows.forEach((pastedRow, rowOffset) => {
        pastedRow.forEach((value, colOffset) => {
          const col = startCol + colOffset;
          if (col < masterHeaders.length) next[startRow + rowOffset][col] = value === "" ? null : value;
        });
      });
      return next.length > 0 ? next : blankMasterRows();
    });
  }

  async function importExcel(file: File | undefined) {
    if (!file) return;
    try {
      const importedRows = await readMasterRowsFromWorkbook(file);
      setImportError(null);
      setHeadersCleared(false);
      setRows(importedRows);
      if (fileInputRef.current) fileInputRef.current.value = "";
    } catch (error) {
      setImportError(error instanceof Error ? error.message : "File Excel tidak bisa dibaca.");
    }
  }

  return (
    <div className="space-y-5">
      <div className={`${panelClass} flex flex-wrap items-end justify-between gap-4 bg-[linear-gradient(135deg,#ffffff,#f8fafc)] p-5`}>
        <div>
          <p className="text-sm font-semibold uppercase text-rose-700">MASTER</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-normal text-slate-950">Input Data Utama</h1>
          <p className="mt-2 max-w-3xl text-sm text-slate-500">
            Struktur kolom mengikuti sheet MASTER. Paste multi-row dan multi-column dari Excel langsung ke sel tabel. Saat disimpan,
            angka tanggal yang sudah pernah masuk report tetap dipertahankan.
          </p>
          <div className="mt-3 flex flex-wrap gap-2">
            <StatusPill>Bulan aktif: {periodLabel}</StatusPill>
            <StatusPill>Terakhir disimpan: {formatSavedAt(savedAt)}</StatusPill>
            <StatusPill tone={autoSaveStatus.includes("tersimpan") ? "good" : "neutral"}>{autoSaveStatus}</StatusPill>
          </div>
          {importError && <p className="mt-3 text-sm font-medium text-rose-700">{importError}</p>}
        </div>
        <div className="flex flex-wrap gap-2">
          <button className={secondaryButtonClass} onClick={() => fileInputRef.current?.click()}>
            <Upload className="h-4 w-4" /> Upload Excel
          </button>
          <button className={secondaryButtonClass} disabled={!canGoBack} onClick={onBack}>
            <Undo2 className="h-4 w-4" /> Back
          </button>
          <button className={dangerButtonClass} onClick={onClear}>
            <Trash2 className="h-4 w-4" /> Clear Data
          </button>
          <button className={primaryButtonClass} onClick={() => onSave()}>
            <Save className="h-4 w-4" /> Save Data
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.xlsm"
            className="hidden"
            onChange={(event) => {
              void importExcel(event.target.files?.[0]);
            }}
          />
        </div>
      </div>

      <section className={`${panelClass} overflow-hidden`}>
        <div className="flex items-center justify-between border-b border-slate-200/70 bg-white/75 px-4 py-3">
          <div>
            <p className="text-sm font-semibold text-slate-700">{formatNumber(rows.length)} baris MASTER</p>
            <p className="mt-0.5 text-xs text-slate-500">Last saved: {formatSavedAt(savedAt)}</p>
          </div>
          <StatusPill>55 kolom MASTER</StatusPill>
        </div>
        <div className="max-h-[68vh] overflow-auto">
          <table className="min-w-[4200px] border-separate border-spacing-0 text-left text-sm">
            <thead>
              <tr>
                <th className="sticky left-0 top-0 z-30 w-14 border-b border-r border-slate-700 bg-slate-950 px-2 py-2 text-xs font-semibold text-white">
                  {headersCleared ? "" : "#"}
                </th>
                {displayHeaders.map((header, index) => (
                  <th key={`${masterHeaders[index]}-${index}`} className="sticky top-0 z-20 min-w-36 border-b border-r border-slate-700 bg-slate-900 px-2 py-2 text-xs font-semibold text-slate-100">
                    {header}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody onPaste={pasteAt}>
              {rows.map((row, rowIndex) => (
                <tr key={rowIndex} className="odd:bg-white even:bg-slate-50/60 hover:bg-rose-50/40">
                  <th className="sticky left-0 z-10 border-b border-r border-slate-200 bg-inherit px-2 py-1 text-xs font-semibold text-slate-500">
                    {rowIndex + 1}
                  </th>
                  {masterHeaders.map((header, colIndex) => (
                    <td
                      key={`${rowIndex}-${header}`}
                      data-row={rowIndex}
                      data-col={colIndex}
                      contentEditable
                      suppressContentEditableWarning
                      className="min-w-36 max-w-52 border-b border-r border-slate-100 px-2 py-1 align-top text-slate-800 outline-none transition focus:bg-amber-50 focus:ring-2 focus:ring-inset focus:ring-amber-300"
                      onBlur={(event) => updateCell(rowIndex, colIndex, event.currentTarget.innerText)}
                    >
                      {formatCell(row[colIndex])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
}

function summaryHeadersFor(previousMonthLabel: string) {
  return [
    "Total",
    "TARGET ACH",
    "RANK ACH",
    "SHORTAGE TO TARGET FM",
    `REAL MTD ${previousMonthLabel}`,
    "GR MTD",
    `REAL FM ${previousMonthLabel}`,
    "GR FM",
    "REAL YTD 2026",
    "ACH YTD",
    "GR YTD",
    "RANK YTD",
    "SHORTAGE TO TARGET YTD",
  ];
}

function ReportDataRow({
  latestDay,
  previousMonthLabel,
  row,
  total = false,
}: {
  latestDay: number;
  previousMonthLabel: string;
  row: ReportRow;
  total?: boolean;
}) {
  const summaryHeaders = summaryHeadersFor(previousMonthLabel);
  const baseClass = total
    ? "bg-slate-900 font-semibold text-white"
    : "odd:bg-white even:bg-slate-50/70 text-slate-800 hover:bg-rose-50/40";
  const summaryValues = [
    formatNumber(row.mtd),
    formatPercent(row.targetAch),
    formatNumber(row.rankAch),
    formatNumber(row.shortageFm),
    formatNumber(row.realMtdPrev),
    formatPercent(row.growthMtd),
    formatNumber(row.realFmPrev),
    formatPercent(row.growthFm),
    formatNumber(row.realYtd),
    formatPercent(row.achYtd),
    formatPercent(row.growthYtd),
    formatNumber(row.rankYtd),
    formatNumber(row.shortageYtd),
  ];
  const summaryMeta = [
    "actual",
    `achievement-${achievementTone(row.targetAch)}`,
    "rank",
    row.shortageFm < 0 ? "shortage-low" : "shortage-good",
    "previous",
    "growth",
    "previous",
    "growth",
    "actual",
    `achievement-${achievementTone(row.achYtd)}`,
    "growth",
    "rank",
    row.shortageYtd < 0 ? "shortage-low" : "shortage-good",
  ];

  return (
    <tr className={baseClass}>
      <th className={`sticky left-0 z-10 w-[170px] min-w-[170px] max-w-[170px] border-b border-r border-slate-200 px-2 py-1.5 text-left ${total ? "bg-slate-900 text-white" : "bg-inherit"}`}>
        {row.label}
      </th>
      <td className={`w-[76px] min-w-[76px] max-w-[76px] border-b border-r border-slate-200 px-1.5 py-1.5 text-right font-medium ${total ? "bg-slate-800 text-white" : "bg-indigo-50/70 text-indigo-950"}`}>{formatNumber(row.targetFm)}</td>
      <td className={`w-[76px] min-w-[76px] max-w-[76px] border-b border-r border-slate-200 px-1.5 py-1.5 text-right font-medium ${total ? "bg-slate-800 text-white" : "bg-indigo-50/70 text-indigo-950"}`}>{formatNumber(row.targetYtd)}</td>
      {row.days.map((value, index) => (
        <td
          key={index}
          className={`w-8 min-w-8 max-w-8 border-b border-r border-slate-100 px-1 py-1.5 text-right ${
            total
              ? "bg-slate-900 text-white"
              : index + 1 > latestDay
                ? "bg-slate-50 text-slate-300"
                : "bg-white text-slate-800"
          }`}
        >
          {formatNumber(value)}
        </td>
      ))}
      {summaryValues.map((value, index) => (
        <td key={summaryHeaders[index]} className={`w-[80px] min-w-[80px] max-w-[80px] border-b border-r border-slate-200 px-1.5 py-1.5 text-right ${summaryCellClass(summaryMeta[index])}`}>
          {summaryMeta[index] === "rank" && value ? <span className="rounded-full bg-slate-900 px-2 py-0.5 text-[11px] font-semibold text-white">{value}</span> : value}
        </td>
      ))}
    </tr>
  );
}

function summaryCellClass(kind: string) {
  if (kind === "actual") return "bg-teal-50/80 font-semibold text-teal-950";
  if (kind === "previous") return "bg-slate-50 text-slate-700";
  if (kind === "rank") return "bg-slate-50";
  if (kind === "shortage-low") return "bg-rose-50/80 font-semibold text-rose-800";
  if (kind === "shortage-good") return "bg-emerald-50/80 font-semibold text-emerald-800";
  if (kind === "achievement-good") return "bg-emerald-50/80 font-semibold text-emerald-800";
  if (kind === "achievement-near") return "bg-amber-50/80 font-semibold text-amber-800";
  if (kind === "achievement-low") return "bg-rose-50/80 font-semibold text-rose-800";
  return "bg-white text-slate-800";
}

function ReportTableView({
  latestDay,
  periodLabel,
  previousMonthLabel,
  table,
}: {
  latestDay: number;
  periodLabel: string;
  previousMonthLabel: string;
  table: ReportTable;
}) {
  const dayCount = table.rows[0]?.days.length ?? 30;
  const summaryHeaders = summaryHeadersFor(previousMonthLabel);
  return (
    <section className={`${panelClass} overflow-hidden`}>
      <div className="border-b border-slate-800 bg-[linear-gradient(135deg,#020617,#0f172a_58%,#155e75)] px-5 py-4 text-center text-white">
        <div className="mx-auto flex max-w-3xl flex-col items-center gap-2">
          <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-cyan-200">Main Table</p>
          <h2 className="text-xl font-semibold tracking-normal text-white">{table.title}</h2>
          <p className="text-sm font-medium text-slate-300">{table.subtitle}</p>
          <span className="rounded-full border border-white/15 bg-white/10 px-3 py-1 text-xs font-semibold text-cyan-50">
            {periodLabel} | Data s.d. Tgl {latestDay}
          </span>
        </div>
      </div>
      <div className="max-h-[72vh] overflow-auto">
        <table className="min-w-[2350px] border-separate border-spacing-0 text-[11px]">
          <thead>
            <tr>
              <th rowSpan={2} className="sticky left-0 top-0 z-30 w-[170px] min-w-[170px] max-w-[170px] border-b border-r border-slate-700 bg-slate-950 px-2 py-2 text-center font-semibold text-white shadow-[8px_0_18px_-18px_rgba(15,23,42,0.8)]">
                {table.total.label === "Total" ? "WITEL" : "BRANCH"}
              </th>
              <th rowSpan={2} className="sticky top-0 z-20 w-[76px] min-w-[76px] max-w-[76px] border-b border-r border-indigo-700 bg-indigo-800 px-1.5 py-2 text-center font-semibold leading-tight text-white">
                TARGET FM
              </th>
              <th rowSpan={2} className="sticky top-0 z-20 w-[76px] min-w-[76px] max-w-[76px] border-b border-r border-indigo-700 bg-indigo-800 px-1.5 py-2 text-center font-semibold leading-tight text-white">
                TARGET YTD
              </th>
              <th colSpan={dayCount} className="sticky top-0 z-20 border-b border-r border-teal-700 bg-teal-800 px-2 py-2 text-center font-semibold text-white">
                TANGGAL
              </th>
              {summaryHeaders.map((header) => (
                <th key={header} rowSpan={2} className="sticky top-0 z-20 w-[80px] min-w-[80px] max-w-[80px] whitespace-normal border-b border-r border-slate-700 bg-slate-900 px-1.5 py-2 text-center text-[9px] font-semibold leading-[1.12] text-white">
                  {header}
                </th>
              ))}
            </tr>
            <tr>
              {Array.from({ length: dayCount }, (_, index) => (
                <th
                  key={index}
                  className={`sticky top-[34px] z-20 w-8 min-w-8 max-w-8 border-b border-r border-slate-200 px-1 py-1.5 text-center text-[10px] font-semibold ${
                    index + 1 > latestDay ? "bg-slate-100 text-slate-400" : "bg-teal-50 text-teal-900"
                  }`}
                >
                  {index + 1}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {table.rows.map((row) => (
              <ReportDataRow key={row.label} row={row} latestDay={latestDay} previousMonthLabel={previousMonthLabel} />
            ))}
            <ReportDataRow row={table.total} latestDay={latestDay} previousMonthLabel={previousMonthLabel} total />
          </tbody>
        </table>
      </div>
    </section>
  );
}

function wrapCanvasText(context: CanvasRenderingContext2D, text: string, maxWidth: number) {
  const words = text.split(" ");
  const lines: string[] = [];
  let current = "";
  for (const word of words) {
    const next = current ? `${current} ${word}` : word;
    if (context.measureText(next).width <= maxWidth || !current) {
      current = next;
    } else {
      lines.push(current);
      current = word;
    }
  }
  if (current) lines.push(current);
  return lines;
}

function drawCell(
  context: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  width: number,
  height: number,
  options: {
    align?: CanvasTextAlign;
    border?: string;
    fill?: string;
    font?: string;
    padding?: number;
    textColor?: string;
    wrap?: boolean;
  } = {},
) {
  context.fillStyle = options.fill ?? "#ffffff";
  context.fillRect(x, y, width, height);
  context.strokeStyle = options.border ?? "#e2e8f0";
  context.lineWidth = 1;
  context.strokeRect(x, y, width, height);
  context.fillStyle = options.textColor ?? "#0f172a";
  context.font = options.font ?? "12px Arial";
  context.textAlign = options.align ?? "right";
  context.textBaseline = "middle";
  const padding = options.padding ?? 8;
  const textX = options.align === "left" ? x + padding : options.align === "center" ? x + width / 2 : x + width - padding;
  if (options.wrap) {
    const lines = wrapCanvasText(context, text, width - padding * 2).slice(0, 3);
    const lineHeight = 11;
    const startY = y + height / 2 - ((lines.length - 1) * lineHeight) / 2;
    lines.forEach((line, index) => context.fillText(line, textX, startY + index * lineHeight));
  } else {
    context.fillText(text, textX, y + height / 2);
  }
}

function drawReportTableCanvas(
  context: CanvasRenderingContext2D,
  table: ReportTable,
  latestDay: number,
  periodLabel: string,
  previousMonthLabel: string,
  x: number,
  y: number,
  totalWidth: number,
) {
  const summaryHeaders = summaryHeadersFor(previousMonthLabel);
  const labelWidth = reportCanvasColumns.label;
  const targetWidth = reportCanvasColumns.target;
  const dayWidth = reportCanvasColumns.day;
  const summaryWidth = reportCanvasColumns.summary;
  const titleHeight = 58;
  const groupHeaderHeight = 36;
  const dayHeaderHeight = 26;
  const rowHeight = 28;
  const rows = [...table.rows, table.total];
  const tableHeight = titleHeight + groupHeaderHeight + dayHeaderHeight + rows.length * rowHeight;

  context.fillStyle = "#ffffff";
  context.fillRect(x, y, totalWidth, tableHeight);
  context.strokeStyle = "#cbd5e1";
  context.strokeRect(x, y, totalWidth, tableHeight);

  context.fillStyle = "#020617";
  context.fillRect(x, y, totalWidth, titleHeight);
  const accent = context.createLinearGradient(x, y, x + totalWidth, y);
  accent.addColorStop(0, "#22d3ee");
  accent.addColorStop(0.45, "#e11d48");
  accent.addColorStop(1, "#f59e0b");
  context.fillStyle = accent;
  context.fillRect(x, y, totalWidth, 4);
  context.fillStyle = "#ffffff";
  context.font = "700 18px Arial";
  context.textAlign = "center";
  context.textBaseline = "top";
  context.fillText(table.title, x + totalWidth / 2, y + 12);
  context.fillStyle = "#cbd5e1";
  context.font = "600 12px Arial";
  context.fillText(table.subtitle, x + totalWidth / 2, y + 34);
  context.font = "700 11px Arial";
  context.fillText(`${periodLabel} | Data s.d. Tgl ${latestDay}`, x + totalWidth / 2, y + 48);
  context.textAlign = "left";

  let cursorX = x;
  const headerY = y + titleHeight;
  const labelHeader = table.total.label === "Total" ? "WITEL" : "BRANCH";
  drawCell(context, labelHeader, cursorX, headerY, labelWidth, groupHeaderHeight + dayHeaderHeight, {
    align: "center",
    border: "#334155",
    fill: "#020617",
    font: "700 11px Arial",
    textColor: "#ffffff",
  });
  cursorX += labelWidth;
  drawCell(context, "TARGET FM", cursorX, headerY, targetWidth, groupHeaderHeight + dayHeaderHeight, {
    align: "center",
    border: "#4338ca",
    fill: "#3730a3",
    font: "700 10px Arial",
    textColor: "#ffffff",
    wrap: true,
  });
  cursorX += targetWidth;
  drawCell(context, "TARGET YTD", cursorX, headerY, targetWidth, groupHeaderHeight + dayHeaderHeight, {
    align: "center",
    border: "#4338ca",
    fill: "#3730a3",
    font: "700 10px Arial",
    textColor: "#ffffff",
    wrap: true,
  });
  cursorX += targetWidth;
  drawCell(context, "TANGGAL", cursorX, headerY, latestDay * dayWidth, groupHeaderHeight, {
    align: "center",
    border: "#0f766e",
    fill: "#0f766e",
    font: "700 11px Arial",
    textColor: "#ffffff",
  });
  for (let index = 0; index < latestDay; index += 1) {
    drawCell(context, String(index + 1), cursorX + index * dayWidth, headerY + groupHeaderHeight, dayWidth, dayHeaderHeight, {
      align: "center",
      border: "#99f6e4",
      fill: "#ccfbf1",
      font: "700 10px Arial",
      textColor: "#134e4a",
    });
  }
  cursorX += latestDay * dayWidth;
  for (const header of summaryHeaders) {
    drawCell(context, header, cursorX, headerY, summaryWidth, groupHeaderHeight + dayHeaderHeight, {
      align: "center",
      border: "#334155",
      fill: "#1e293b",
      font: "700 8px Arial",
      padding: 4,
      textColor: "#ffffff",
      wrap: true,
    });
    cursorX += summaryWidth;
  }

  rows.forEach((row, rowIndex) => {
    const total = rowIndex === rows.length - 1;
    const rowY = headerY + groupHeaderHeight + dayHeaderHeight + rowIndex * rowHeight;
    const baseFill = total ? "#e2e8f0" : rowIndex % 2 === 0 ? "#ffffff" : "#f8fafc";
    let cellX = x;
    const bodyFont = total ? "700 11px Arial" : "11px Arial";
    drawCell(context, row.label, cellX, rowY, labelWidth, rowHeight, {
      align: "left",
      fill: baseFill,
      font: bodyFont,
    });
    cellX += labelWidth;
    drawCell(context, formatNumber(row.targetFm), cellX, rowY, targetWidth, rowHeight, {
      fill: "#eef2ff",
      font: "700 11px Arial",
      textColor: "#312e81",
    });
    cellX += targetWidth;
    drawCell(context, formatNumber(row.targetYtd), cellX, rowY, targetWidth, rowHeight, {
      fill: "#eef2ff",
      font: "700 11px Arial",
      textColor: "#312e81",
    });
    cellX += targetWidth;
    row.days.slice(0, latestDay).forEach((value) => {
      drawCell(context, formatNumber(value), cellX, rowY, dayWidth, rowHeight, {
        fill: baseFill,
        font: bodyFont,
      });
      cellX += dayWidth;
    });

    const summaryValues = [
      { value: formatNumber(row.mtd), fill: "#ccfbf1", color: "#134e4a", bold: true },
      { value: formatPercent(row.targetAch), tone: achievementTone(row.targetAch) },
      { value: formatNumber(row.rankAch), fill: "#f1f5f9", color: "#0f172a", bold: true },
      { value: formatNumber(row.shortageFm), fill: row.shortageFm < 0 ? "#ffe4e6" : "#dcfce7", color: row.shortageFm < 0 ? "#9f1239" : "#166534", bold: true },
      { value: formatNumber(row.realMtdPrev), fill: "#f8fafc", color: "#334155" },
      { value: formatPercent(row.growthMtd), fill: baseFill, color: "#0f172a" },
      { value: formatNumber(row.realFmPrev), fill: "#f8fafc", color: "#334155" },
      { value: formatPercent(row.growthFm), fill: baseFill, color: "#0f172a" },
      { value: formatNumber(row.realYtd), fill: "#ccfbf1", color: "#134e4a", bold: true },
      { value: formatPercent(row.achYtd), tone: achievementTone(row.achYtd) },
      { value: formatPercent(row.growthYtd), fill: baseFill, color: "#0f172a" },
      { value: formatNumber(row.rankYtd), fill: "#f1f5f9", color: "#0f172a", bold: true },
      { value: formatNumber(row.shortageYtd), fill: row.shortageYtd < 0 ? "#ffe4e6" : "#dcfce7", color: row.shortageYtd < 0 ? "#9f1239" : "#166534", bold: true },
    ];
    for (const item of summaryValues) {
      let fill = item.fill ?? baseFill;
      let color = item.color ?? "#0f172a";
      if (item.tone === "good") {
        fill = "#dcfce7";
        color = "#166534";
      } else if (item.tone === "near") {
        fill = "#fef3c7";
        color = "#92400e";
      } else if (item.tone === "low") {
        fill = "#ffe4e6";
        color = "#9f1239";
      }
      drawCell(context, item.value, cellX, rowY, summaryWidth, rowHeight, {
        fill,
        font: item.bold || item.tone ? "700 11px Arial" : bodyFont,
        textColor: color,
      });
      cellX += summaryWidth;
    }
  });

  return tableHeight;
}

function drawConfidentialWatermark(context: CanvasRenderingContext2D, width: number, height: number) {
  context.save();
  context.globalAlpha = 0.055;
  context.fillStyle = "#0f172a";
  context.font = "700 58px Arial";
  context.textAlign = "center";
  context.textBaseline = "middle";
  context.translate(width / 2, height / 2);
  context.rotate((-22 * Math.PI) / 180);
  context.fillText("CONFIDENTIAL", 0, 0);
  context.restore();
}

async function createReportPng(
  witelTable: ReportTable,
  branchTable: ReportTable,
  latestDay: number,
  timestamp: Date,
  variant: "sementara" | "stf",
  periodLabel: string,
  previousMonthLabel: string,
) {
  await document.fonts?.ready;
  const scale = 2;
  const padding = 28;
  const labelWidth = reportCanvasColumns.label;
  const targetWidth = reportCanvasColumns.target;
  const dayWidth = reportCanvasColumns.day;
  const summaryWidth = reportCanvasColumns.summary;
  const summaryHeaders = summaryHeadersFor(previousMonthLabel);
  const tableWidth = labelWidth + targetWidth * 2 + latestDay * dayWidth + summaryHeaders.length * summaryWidth;
  const tableHeight = 58 + 36 + 26 + 9 * 28;
  const logicalWidth = tableWidth + padding * 2;
  const logicalHeight = padding + 68 + tableHeight + 22 + tableHeight + padding;
  const canvas = document.createElement("canvas");
  canvas.width = logicalWidth * scale;
  canvas.height = logicalHeight * scale;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas browser tidak tersedia.");
  context.scale(scale, scale);
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, logicalWidth, logicalHeight);
  context.fillStyle = "#0f172a";
  context.font = "700 28px Arial";
  context.textAlign = "left";
  context.textBaseline = "top";
  context.fillText(variant === "stf" ? "Report HSI Target STF SEPT" : "Report Sementara SEPT 2026", padding, padding);
  context.fillStyle = "#475569";
  context.font = "15px Arial";
  context.fillText(`Downloaded: ${formatJakartaDateTime(timestamp, false)}`, padding, padding + 36);
  context.fillText(`Periode: ${periodLabel} | Cut-off tanggal aktual terakhir: ${latestDay}`, padding + 360, padding + 36);
  context.fillStyle = "#881337";
  context.font = "700 12px Arial";
  context.textAlign = "right";
  context.fillText("CONFIDENTIAL - INTERNAL USE ONLY", logicalWidth - padding, padding + 8);
  context.textAlign = "left";
  const firstTableHeight = drawReportTableCanvas(context, witelTable, latestDay, periodLabel, previousMonthLabel, padding, padding + 68, tableWidth);
  drawReportTableCanvas(context, branchTable, latestDay, periodLabel, previousMonthLabel, padding, padding + 68 + firstTableHeight + 22, tableWidth);
  drawConfidentialWatermark(context, logicalWidth, logicalHeight);
  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((result) => (result ? resolve(result) : reject(new Error("Gagal menyimpan PNG."))), "image/png", 1);
  });
  downloadBlob(blob, reportFileName(timestamp, variant));
}

function ReportPage({ report, variant }: { report: ComputedReport; variant: "sementara" | "stf" }) {
  const isStf = variant === "stf";
  const witelTable = isStf ? report.targetStf.witel : report.witel;
  const branchTable = isStf ? report.targetStf.branch : report.branch;
  const latestDay = latestPopulatedDay([witelTable, branchTable]);
  const now = useJakartaClock();
  const [downloadState, setDownloadState] = useState<"idle" | "working" | "error">("idle");

  async function downloadReport() {
    const timestamp = new Date();
    setDownloadState("working");
    try {
      await createReportPng(witelTable, branchTable, latestDay, timestamp, variant, report.periodLabel, report.previousMonthLabel);
      setDownloadState("idle");
    } catch {
      setDownloadState("error");
    }
  }

  return (
    <div className="space-y-6">
      <div className={`${panelClass} flex flex-wrap items-end justify-between gap-4 bg-[linear-gradient(135deg,#ffffff,#f8fafc_55%,#fff1f2)] p-5`}>
        <div>
          <p className="text-sm font-semibold uppercase text-rose-700">
            {isStf ? "Report HSI Target STF SEPT" : "Report Sementara SEPT 2026"}
          </p>
          <h1 className="mt-1 text-3xl font-semibold tracking-normal text-slate-950">
            {isStf ? "Dua Tabel Utama Target STF" : "Dua Tabel Utama Report"}
          </h1>
          <p className="mt-2 max-w-3xl text-sm text-slate-500">
            Area helper dan tabel internal tidak ditampilkan. Perhitungan tetap dijalankan di JavaScript dari MASTER.
          </p>
          <div className="mt-3 flex flex-wrap items-center gap-2">
            <p className="text-sm font-medium text-slate-600">
              Last rendered: {formatJakartaDateTime(now, true)}
            </p>
            <span className="inline-flex items-center gap-1.5 rounded-full border border-rose-200 bg-rose-50 px-3 py-1 text-xs font-semibold uppercase text-rose-800">
              <ShieldCheck className="h-3.5 w-3.5" />
              Confidential
            </span>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <StatusPill>{report.periodLabel}</StatusPill>
          <StatusPill>Bulan report: {report.periodLabel}</StatusPill>
          <StatusPill>Cut-off {report.cutoffLabel}</StatusPill>
          <StatusPill>Data s.d. tanggal {latestDay}</StatusPill>
          <StatusPill>{isStf ? "Target STF RSME" : report.source}</StatusPill>
          <button
            className={primaryButtonClass}
            disabled={downloadState === "working"}
            onClick={downloadReport}
          >
            <Download className="h-4 w-4" />
            {downloadState === "working" ? "Preparing PNG..." : "Download Report"}
          </button>
        </div>
      </div>
      {downloadState === "error" && (
        <div className="rounded-md border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-800">
          PNG belum berhasil dibuat. Coba tekan Download Report sekali lagi.
        </div>
      )}
      <ReportTableView table={witelTable} latestDay={latestDay} periodLabel={report.periodLabel} previousMonthLabel={report.previousMonthLabel} />
      <ReportTableView table={branchTable} latestDay={latestDay} periodLabel={report.periodLabel} previousMonthLabel={report.previousMonthLabel} />
    </div>
  );
}

export function HsiDashboard() {
  const router = useRouter();
  const [view, setView] = useState<View>("home");
  const [rows, setRows] = useState<MasterRow[]>(() => initialMasterRows());
  const [periodArchive, setPeriodArchive] = useState<PeriodArchive>({});
  const [selectedPeriodKey, setSelectedPeriodKey] = useState(() => currentJakartaPeriodKey());
  const [reportHistory, setReportHistory] = useState<ReportHistory | null>(null);
  const [savedAt, setSavedAt] = useState<string | null>(null);
  const [saveSnapshots, setSaveSnapshots] = useState<SaveSnapshots>({});
  const [masterHeadersCleared, setMasterHeadersCleared] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [autoSaveStatus, setAutoSaveStatus] = useState("Auto-save standby");
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    const timeout = window.setTimeout(() => {
      const shouldRefreshSeed = window.localStorage.getItem(DATA_VERSION_KEY) !== seedData.dataVersion;
      const storedRows = window.localStorage.getItem(STORAGE_KEY);
      const storedAt = window.localStorage.getItem(SAVED_AT_KEY);
      const storedHistory = parseReportHistory(window.localStorage.getItem(REPORT_HISTORY_KEY));
      const storedArchive = parsePeriodArchive(window.localStorage.getItem(PERIOD_ARCHIVE_KEY));
      const storedSnapshots = parseSaveSnapshots(window.localStorage.getItem(SAVE_SNAPSHOTS_KEY));
      const storedSelectedPeriod = parsePeriodKey(window.localStorage.getItem(SELECTED_PERIOD_KEY));
      const storedHeadersCleared = window.localStorage.getItem(MASTER_HEADERS_CLEARED_KEY) === "true";
      const calendarPeriodKey = currentJakartaPeriodKey();
      const seedPeriodKey = workbookSeedPeriodKey();
      const seedRows = initialMasterRows();
      const seedRowsPeriodKey = detectPeriodKey(seedRows, seedPeriodKey);
      const hasStoredArchive = Object.keys(storedArchive).length > 0;
      const currentSavedAt = shouldRefreshSeed ? null : storedAt;
      let nextRows = seedRows;
      if (!shouldRefreshSeed && !hasStoredArchive && storedRows) {
        try {
          const parsed = JSON.parse(storedRows) as unknown[];
          nextRows = parsed.map((row) => ensureRowWidth(Array.isArray(row) ? row : []));
        } catch {
          nextRows = seedRows;
        }
      }
      const currentRowsPeriodKey = detectPeriodKey(nextRows, seedRowsPeriodKey);
      let nextHistory = shouldRefreshSeed ? null : storedHistory;
      if (nextHistory) {
        nextHistory = storedHistory;
      } else {
        const initialHistory = createReportHistory(computeReport(seedRows, null, seedRowsPeriodKey), null);
        nextHistory = mergeReportHistory(initialHistory, computeReport(nextRows, currentSavedAt, currentRowsPeriodKey), currentSavedAt);
        window.localStorage.setItem(REPORT_HISTORY_KEY, JSON.stringify(nextHistory));
      }
      const seedArchive = initialSeedPeriodArchive();
      const archiveWithCurrent: PeriodArchive = shouldRefreshSeed
        ? { ...storedArchive, ...seedArchive }
        : { ...seedArchive, ...storedArchive };
      const seedHistory = createReportHistory(computeReport(seedRows, null, seedRowsPeriodKey), null);
      if (shouldRefreshSeed || !archiveWithCurrent[seedRowsPeriodKey]) {
        archiveWithCurrent[seedRowsPeriodKey] = createPeriodEntry(seedRowsPeriodKey, seedRows, null, seedHistory);
      }
      if (!shouldRefreshSeed && !hasStoredArchive) {
        archiveWithCurrent[currentRowsPeriodKey] = createPeriodEntry(currentRowsPeriodKey, nextRows, currentSavedAt, nextHistory);
      }
      const hadCalendarPeriod = Boolean(archiveWithCurrent[calendarPeriodKey]);
      if (!hadCalendarPeriod) {
        archiveWithCurrent[calendarPeriodKey] = createPeriodEntry(calendarPeriodKey, blankMasterRows(), null, null);
      }
      const activePeriodKey =
        shouldRefreshSeed && archiveWithCurrent[seedRowsPeriodKey]
          ? seedRowsPeriodKey
          : !hadCalendarPeriod
          ? calendarPeriodKey
          : storedSelectedPeriod && archiveWithCurrent[storedSelectedPeriod]
            ? storedSelectedPeriod
            : calendarPeriodKey;
      const activePeriod = archiveWithCurrent[activePeriodKey];
      setPeriodArchive(archiveWithCurrent);
      setSelectedPeriodKey(activePeriodKey);
      setRows(activePeriod.rows);
      setSavedAt(activePeriod.savedAt);
      setReportHistory(activePeriod.reportHistory);
      setSaveSnapshots(storedSnapshots);
      setMasterHeadersCleared(shouldRefreshSeed ? false : storedHeadersCleared);
      setDirty(false);
      window.localStorage.setItem(PERIOD_ARCHIVE_KEY, JSON.stringify(archiveWithCurrent));
      window.localStorage.setItem(SELECTED_PERIOD_KEY, activePeriodKey);
      window.localStorage.setItem(DATA_VERSION_KEY, seedData.dataVersion);
      window.localStorage.setItem(SAVE_SNAPSHOTS_KEY, JSON.stringify(storedSnapshots));
      window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, shouldRefreshSeed ? "false" : String(storedHeadersCleared));
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(activePeriod.rows));
      if (activePeriod.savedAt) {
        window.localStorage.setItem(SAVED_AT_KEY, activePeriod.savedAt);
      } else {
        window.localStorage.removeItem(SAVED_AT_KEY);
      }
      if (activePeriod.reportHistory) {
        window.localStorage.setItem(REPORT_HISTORY_KEY, JSON.stringify(activePeriod.reportHistory));
      } else {
        window.localStorage.removeItem(REPORT_HISTORY_KEY);
      }
      setHydrated(true);
    }, 0);
    return () => window.clearTimeout(timeout);
  }, []);

  const report = useMemo(() => {
    const activePeriodKey = selectedPeriodKey || detectPeriodKey(rows);
    const archivedSnapshot = periodArchive[activePeriodKey]?.reportSnapshot;
    if (archivedSnapshot) return archivedSnapshot;
    const computed = applyReportHistory(computeReport(rows, savedAt, activePeriodKey), reportHistory);
    return {
      ...computed,
      periodLabel: periodLabelFromKey(activePeriodKey),
      previousMonthLabel: previousMonthLabelFromKey(activePeriodKey),
    };
  }, [periodArchive, reportHistory, rows, savedAt, selectedPeriodKey]);

  const periodOptions = useMemo(
    () => {
      const options = Object.values(periodArchive)
        .sort((a, b) => b.periodKey.localeCompare(a.periodKey))
        .map((entry) => ({ label: entry.label, periodKey: entry.periodKey, savedAt: entry.savedAt }));
      if (!options.some((option) => option.periodKey === selectedPeriodKey)) {
        options.push({ label: periodLabelFromKey(selectedPeriodKey), periodKey: selectedPeriodKey, savedAt });
      }
      return options.sort((a, b) => b.periodKey.localeCompare(a.periodKey));
    },
    [periodArchive, savedAt, selectedPeriodKey],
  );

  const saveRows = useCallback((options: { automated?: boolean; pushSnapshot?: boolean } = {}) => {
    const timestamp = new Date().toISOString();
    const periodKey = detectPeriodKey(rows, selectedPeriodKey || currentJakartaPeriodKey());
    const currentReport = computeReport(rows, timestamp, periodKey);
    const nextHistory = mergeReportHistory(reportHistory, currentReport, timestamp);
    const previousEntry = periodArchive[periodKey];
    const shouldPushSnapshot = options.pushSnapshot !== false && Boolean(previousEntry?.savedAt);
    const nextSnapshots = shouldPushSnapshot
      ? {
          ...saveSnapshots,
          [periodKey]: [
            ...(saveSnapshots[periodKey] ?? []),
            {
              periodKey,
              reportHistory: previousEntry?.reportHistory ?? null,
              rows: previousEntry?.rows ?? [],
              savedAt: previousEntry?.savedAt ?? null,
            },
          ].slice(-10),
        }
      : saveSnapshots;
    const nextArchive = {
      ...periodArchive,
      [periodKey]: createPeriodEntry(periodKey, rows, timestamp, nextHistory),
    };
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(rows));
    window.localStorage.setItem(SAVED_AT_KEY, timestamp);
    window.localStorage.setItem(REPORT_HISTORY_KEY, JSON.stringify(nextHistory));
    window.localStorage.setItem(PERIOD_ARCHIVE_KEY, JSON.stringify(nextArchive));
    window.localStorage.setItem(SELECTED_PERIOD_KEY, periodKey);
    window.localStorage.setItem(DATA_VERSION_KEY, seedData.dataVersion);
    window.localStorage.setItem(SAVE_SNAPSHOTS_KEY, JSON.stringify(nextSnapshots));
    window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, String(masterHeadersCleared));
    setPeriodArchive(nextArchive);
    setReportHistory(nextHistory);
    setSelectedPeriodKey(periodKey);
    setSavedAt(timestamp);
    setSaveSnapshots(nextSnapshots);
    setDirty(false);
    setAutoSaveStatus(options.automated ? `Auto-save tersimpan: ${formatJakartaDateTime(new Date(timestamp), true)}` : "Data tersimpan manual");
  }, [masterHeadersCleared, periodArchive, reportHistory, rows, saveSnapshots, selectedPeriodKey]);

  useEffect(() => {
    if (!hydrated) return undefined;
    if (!dirty) return undefined;
    const timeout = window.setTimeout(() => {
      saveRows({ automated: true });
    }, AUTO_SAVE_DELAY_MS);
    return () => window.clearTimeout(timeout);
  }, [dirty, hydrated, saveRows]);

  function clearRows() {
    setRows(blankMasterRows());
    setMasterHeadersCleared(true);
    setDirty(true);
    setAutoSaveStatus("Auto-save dalam 3 menit jika tidak ada perubahan");
    window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, "true");
  }

  function restorePreviousSave() {
    const versions = saveSnapshots[selectedPeriodKey] ?? [];
    const previous = versions[versions.length - 1];
    if (!previous) return;
    const nextSnapshots = {
      ...saveSnapshots,
      [selectedPeriodKey]: versions.slice(0, -1),
    };
    const nextArchive = {
      ...periodArchive,
      [selectedPeriodKey]: createPeriodEntry(selectedPeriodKey, previous.rows, previous.savedAt, previous.reportHistory),
    };
    setRows(previous.rows);
    setSavedAt(previous.savedAt);
    setReportHistory(previous.reportHistory);
    setPeriodArchive(nextArchive);
    setSaveSnapshots(nextSnapshots);
    setMasterHeadersCleared(false);
    setDirty(false);
    setAutoSaveStatus("Kembali ke save sebelumnya");
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(previous.rows));
    if (previous.savedAt) {
      window.localStorage.setItem(SAVED_AT_KEY, previous.savedAt);
    } else {
      window.localStorage.removeItem(SAVED_AT_KEY);
    }
    if (previous.reportHistory) {
      window.localStorage.setItem(REPORT_HISTORY_KEY, JSON.stringify(previous.reportHistory));
    } else {
      window.localStorage.removeItem(REPORT_HISTORY_KEY);
    }
    window.localStorage.setItem(PERIOD_ARCHIVE_KEY, JSON.stringify(nextArchive));
    window.localStorage.setItem(SAVE_SNAPSHOTS_KEY, JSON.stringify(nextSnapshots));
    window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, "false");
  }

  function selectPeriod(periodKey: string) {
    const entry = periodArchive[periodKey];
    if (!entry) return;
    setSelectedPeriodKey(periodKey);
    setRows(entry.rows);
    setSavedAt(entry.savedAt);
    setReportHistory(entry.reportHistory);
    setDirty(false);
    setAutoSaveStatus("Auto-save standby");
    setMasterHeadersCleared(false);
    window.localStorage.setItem(SELECTED_PERIOD_KEY, periodKey);
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(entry.rows));
    window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, "false");
    if (entry.savedAt) {
      window.localStorage.setItem(SAVED_AT_KEY, entry.savedAt);
    } else {
      window.localStorage.removeItem(SAVED_AT_KEY);
    }
    if (entry.reportHistory) {
      window.localStorage.setItem(REPORT_HISTORY_KEY, JSON.stringify(entry.reportHistory));
    } else {
      window.localStorage.removeItem(REPORT_HISTORY_KEY);
    }
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" });
    router.replace("/login");
    router.refresh();
  }

  return (
    <main className="min-h-screen bg-[linear-gradient(180deg,#f8fafc_0%,#eef2f7_48%,#fff7f7_100%)] text-slate-950">
      <header className="sticky top-0 z-40 border-b border-white/70 bg-white/82 shadow-[0_18px_40px_-34px_rgba(15,23,42,0.7)] backdrop-blur-xl">
        <div className="mx-auto flex max-w-[1920px] flex-col gap-3 px-4 py-3 sm:px-6 xl:flex-row xl:items-center xl:justify-between">
          <div className="flex items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-slate-950 text-sm font-bold text-white shadow-[0_12px_26px_-16px_rgba(15,23,42,0.9)]">
                HSI
              </div>
              <div>
                <p className="text-sm font-semibold text-slate-950">Reporting HSI 2026</p>
                <p className="text-xs text-slate-500">{hydrated ? `${formatNumber(rows.length)} baris MASTER` : "Memuat data"}</p>
              </div>
            </div>
            <button
              className={`${secondaryButtonClass} xl:hidden`}
              onClick={logout}
            >
              <LogOut className="h-4 w-4" />
              Logout
            </button>
          </div>
          <div className="flex min-w-0 flex-1 items-center gap-3 xl:justify-end">
            <label className="hidden shrink-0 items-center gap-2 rounded-md border border-slate-200/70 bg-white/70 px-3 py-2 text-xs font-semibold text-slate-600 shadow-sm sm:flex">
              Periode
              <select
                className="bg-transparent text-sm font-semibold text-slate-950 outline-none"
                value={selectedPeriodKey}
                onChange={(event) => selectPeriod(event.target.value)}
              >
                {periodOptions.map((option) => (
                  <option key={option.periodKey} value={option.periodKey}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
            <nav className="scrollbar-thin flex min-w-0 flex-1 gap-2 overflow-x-auto xl:flex-none">
            {navItems.map((item) => {
              const Icon = item.icon;
              const active = view === item.id;
              return (
                <button
                  key={item.id}
                  onClick={() => setView(item.id)}
                  className={`inline-flex h-10 shrink-0 items-center gap-2 rounded-md px-3 text-sm font-semibold transition ${
                    active
                      ? "bg-slate-950 text-white shadow-[0_14px_28px_-20px_rgba(15,23,42,0.9)]"
                      : "border border-slate-200/60 bg-white/65 text-slate-700 hover:border-slate-300 hover:bg-white"
                  }`}
                >
                  <Icon className="h-4 w-4" />
                  {item.label}
                </button>
              );
            })}
            </nav>
            <div className="hidden shrink-0 items-center gap-2 xl:flex">
              <StatusPill>Cut-off {report.cutoffLabel}</StatusPill>
              <button
                className={secondaryButtonClass}
                onClick={logout}
              >
                <LogOut className="h-4 w-4" />
                Logout
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1920px] px-3 py-5 sm:px-5 lg:px-6">
          {view === "home" && <DashboardHome report={report} />}
          {view === "master" && (
            <MasterPage
              autoSaveStatus={autoSaveStatus}
              canGoBack={(saveSnapshots[selectedPeriodKey] ?? []).length > 0}
              headersCleared={masterHeadersCleared}
              periodLabel={report.periodLabel}
              rows={rows}
              savedAt={savedAt}
              setHeadersCleared={setMasterHeadersCleared}
              setRows={(updater) => {
                setRows(updater);
                setDirty(true);
                setAutoSaveStatus("Auto-save dalam 3 menit jika tidak ada perubahan");
              }}
              onBack={restorePreviousSave}
              onSave={saveRows}
              onClear={clearRows}
            />
          )}
          {view === "report" && <ReportPage report={report} variant="sementara" />}
          {view === "stf" && <ReportPage report={report} variant="stf" />}
      </div>
    </main>
  );
}
