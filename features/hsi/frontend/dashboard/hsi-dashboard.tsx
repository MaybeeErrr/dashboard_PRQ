"use client";

import {
  Bar,
  CartesianGrid,
  ComposedChart,
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
  ArrowLeft,
  ArrowUpRight,
  Database,
  Download,
  FileSpreadsheet,
  Gauge,
  Home,
  LockKeyhole,
  Save,
  ShieldCheck,
  Target,
  Trash2,
  Undo2,
  TrendingUp,
  Upload,
} from "lucide-react";
import Image from "next/image";
import { useRouter } from "next/navigation";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import { seedData } from "@/features/hsi/domain/reporting/seed-data";
import type { DashboardPermissions } from "@/features/hsi/backend/auth/session";
import { CompactRankingInsight } from "@/features/shared/dashboard/compact-ranking-insight";
import { cutoffDateFromPeriodPart, dateRangeLabel as formatSharedDateRangeLabel, formatDataCutoff, formatDataCutoffDate, formatYtdLabel, getPeriodSemanticLabel } from "@/features/shared/dashboard/date-label";
import { DashboardMetricCard, type DashboardMetricTone } from "@/features/shared/dashboard/metric-card";
import { MasterUnlockModal } from "@/features/shared/dashboard/master-unlock-modal";
import { DashboardPeriodFilter, defaultPeriodFilter, type DashboardPeriodFilterState } from "@/features/shared/dashboard/period-filter";
import { DashboardStatusPill } from "@/features/shared/dashboard/status-pill";
import {
  computeReport,
  initialMasterRows,
  masterHeaders,
  type ComputedReport,
  type MasterRow,
  type ReportRow,
  type ReportTable,
} from "@/features/hsi/domain/reporting/reporting";

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
type DashboardPersistedState = {
  dataVersion: string;
  masterHeadersCleared: boolean;
  periodArchive: PeriodArchive;
  saveSnapshots: SaveSnapshots;
  selectedPeriodKey: string;
  updatedAt: string | null;
};

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
  { id: "report", label: "Report HSI Target RKAP", icon: FileSpreadsheet },
  { id: "stf", label: "Report HSI Target STF", icon: Target },
];

const panelClass =
  "rounded-xl border border-white/15 bg-white shadow-[0_18px_44px_-30px_rgba(0,0,0,0.58)] ring-1 ring-slate-950/[0.04] transition-shadow duration-200";
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
  return <DashboardStatusPill tone={tone}>{children}</DashboardStatusPill>;
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

function dateInputForPeriodKey(periodKey: string, day: "first" | "last") {
  const year = Number(periodKey.slice(0, 4));
  const month = Number(periodKey.slice(4, 6));
  const date = day === "first" ? 1 : new Date(year, month, 0).getDate();
  return `${year}-${String(month).padStart(2, "0")}-${String(date).padStart(2, "0")}`;
}

function monthRangeKeys(startPeriodKey: string, endPeriodKey: string) {
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

function dateRangeToMonthParts(startDate: string, endDate: string) {
  const start = new Date(`${startDate}T00:00:00`);
  const end = new Date(`${endDate}T00:00:00`);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime()) || start > end) return [];
  const startKey = `${start.getFullYear()}${String(start.getMonth() + 1).padStart(2, "0")}`;
  const endKey = `${end.getFullYear()}${String(end.getMonth() + 1).padStart(2, "0")}`;
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

function periodRangeLabel(startPeriodKey: string, endPeriodKey: string) {
  if (startPeriodKey === endPeriodKey) return periodLabelFromKey(startPeriodKey);
  const startYear = startPeriodKey.slice(0, 4);
  const endYear = endPeriodKey.slice(0, 4);
  const startMonth = monthNames[Number(startPeriodKey.slice(4, 6)) - 1] ?? startPeriodKey;
  const endMonth = monthNames[Number(endPeriodKey.slice(4, 6)) - 1] ?? endPeriodKey;
  return startYear === endYear ? `${startMonth} - ${endMonth} ${endYear}` : `${startMonth} ${startYear} - ${endMonth} ${endYear}`;
}

function dateRangeLabel(startDate: string, endDate: string) {
  return formatSharedDateRangeLabel(startDate, endDate);
}

function normalizePeriodFilter(value: DashboardPeriodFilterState) {
  if (value.mode === "current-month") {
    return {
      ...value,
      endDate: dateInputForPeriodKey(value.startMonth, "last"),
      endMonth: value.startMonth,
      startDate: dateInputForPeriodKey(value.startMonth, "first"),
    };
  }
  if (value.mode === "ytd") {
    const year = value.endMonth.slice(0, 4);
    return {
      ...value,
      endDate: dateInputForPeriodKey(value.endMonth, "last"),
      startDate: `${year}-01-01`,
      startMonth: `${year}01`,
    };
  }
  if (value.mode === "monthly") {
    return {
      ...value,
      endDate: dateInputForPeriodKey(value.endMonth, "last"),
      startDate: dateInputForPeriodKey(value.startMonth, "first"),
    };
  }
  const startMonth = value.startDate ? value.startDate.slice(0, 7).replace("-", "") : value.startMonth;
  const endMonth = value.endDate ? value.endDate.slice(0, 7).replace("-", "") : value.endMonth;
  return { ...value, endMonth, startMonth };
}

function validatePeriodFilter(value: DashboardPeriodFilterState) {
  const normalized = normalizePeriodFilter(value);
  if (normalized.mode === "date-range") {
    if (!normalized.startDate || !normalized.endDate) return "Start date dan end date harus diisi.";
    if (new Date(`${normalized.startDate}T00:00:00`) > new Date(`${normalized.endDate}T00:00:00`)) {
      return "Start date must be before or equal to end date.";
    }
  } else if (normalized.startMonth > normalized.endMonth) {
    return "Start period must be before or equal to end period.";
  }
  return null;
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

function parsePersistedDashboardState(value: unknown): DashboardPersistedState | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Partial<DashboardPersistedState>;
  const selectedPeriodKey = parsePeriodKey(candidate.selectedPeriodKey);
  if (!selectedPeriodKey) return null;
  const periodArchive = parsePeriodArchive(JSON.stringify(candidate.periodArchive ?? {}));
  if (!periodArchive[selectedPeriodKey]) return null;
  return {
    dataVersion: typeof candidate.dataVersion === "string" ? candidate.dataVersion : seedData.dataVersion,
    masterHeadersCleared: Boolean(candidate.masterHeadersCleared),
    periodArchive,
    saveSnapshots: parseSaveSnapshots(JSON.stringify(candidate.saveSnapshots ?? {})),
    selectedPeriodKey,
    updatedAt: typeof candidate.updatedAt === "string" ? candidate.updatedAt : null,
  };
}

function persistDashboardStateLocally(state: DashboardPersistedState) {
  const activePeriod = state.periodArchive[state.selectedPeriodKey];
  window.localStorage.setItem(PERIOD_ARCHIVE_KEY, JSON.stringify(state.periodArchive));
  window.localStorage.setItem(SELECTED_PERIOD_KEY, state.selectedPeriodKey);
  window.localStorage.setItem(DATA_VERSION_KEY, state.dataVersion);
  window.localStorage.setItem(SAVE_SNAPSHOTS_KEY, JSON.stringify(state.saveSnapshots));
  window.localStorage.setItem(MASTER_HEADERS_CLEARED_KEY, String(state.masterHeadersCleared));
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
}

async function loadDashboardStateFromServer(): Promise<{ configured: boolean; state: DashboardPersistedState | null }> {
  const response = await fetch("/api/dashboard-state", { cache: "no-store" });
  if (!response.ok) throw new Error("Gagal memuat data dari server.");
  const payload = (await response.json()) as { configured?: boolean; state?: unknown };
  return {
    configured: Boolean(payload.configured),
    state: parsePersistedDashboardState(payload.state),
  };
}

async function saveDashboardStateToServer(state: DashboardPersistedState, masterUnlockToken: string | null): Promise<{ configured: boolean }> {
  const response = await fetch("/api/dashboard-state", {
    body: JSON.stringify(state),
    headers: {
      "Content-Type": "application/json",
      ...(masterUnlockToken ? { "x-master-unlock-token": masterUnlockToken } : {}),
    },
    method: "PUT",
  });
  if (response.status === 503) return { configured: false };
  if (!response.ok) throw new Error("Gagal menyimpan data ke server.");
  const payload = (await response.json()) as { configured?: boolean };
  return { configured: Boolean(payload.configured) };
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

type PeriodReportPart = {
  endDay: number;
  periodKey: string;
  report: ComputedReport;
  startDay: number;
  weight: number;
};

function latestActualDateFromHsiParts(parts: PeriodReportPart[]) {
  for (const part of [...parts].reverse()) {
    const days = part.report.witel.total.days;
    for (let day = part.endDay; day >= part.startDay; day -= 1) {
      if (hasActualValue(days[day - 1])) {
        return cutoffDateFromPeriodPart({ endDay: day, periodKey: part.periodKey });
      }
    }
  }
  const fallback = parts.at(-1);
  return fallback ? cutoffDateFromPeriodPart(fallback) : new Date();
}

function rowSegment(row: ReportRow, startDay: number, endDay: number) {
  const days = row.days.slice(startDay - 1, endDay);
  const mtd = days.reduce<number>((sum, value) => sum + (value ?? 0), 0);
  return { days, mtd };
}

function aggregateRows(
  label: string,
  rows: { row: ReportRow; startDay: number; endDay: number; weight: number }[],
  ytdRows = rows,
): ReportRow {
  const base = rows[0]?.row;
  const days = rows.flatMap((part) => rowSegment(part.row, part.startDay, part.endDay).days);
  const mtd = rows.reduce((sum, part) => sum + rowSegment(part.row, part.startDay, part.endDay).mtd, 0);
  const targetFm = rows.reduce((sum, part) => sum + part.row.targetFm * part.weight, 0);
  const realYtd = ytdRows.reduce((sum, part) => sum + rowSegment(part.row, part.startDay, part.endDay).mtd, 0);
  const targetYtd = ytdRows.reduce((sum, part) => sum + part.row.targetFm * part.weight, 0);
  const realMtdPrev = rows.reduce((sum, part) => sum + part.row.realMtdPrev * part.weight, 0);
  const realFmPrev = rows.reduce((sum, part) => sum + part.row.realFmPrev * part.weight, 0);
  const previousYtd = ytdRows.reduce((sum, part) => {
    const previous = previousFromGrowth(part.row.realYtd, part.row.growthYtd);
    return sum + previous * part.weight;
  }, 0);
  return {
    ...(base ?? {
      achYtd: null,
      days: [],
      growthFm: null,
      growthMtd: null,
      growthYtd: null,
      label,
      mtd: 0,
      rankAch: null,
      rankYtd: null,
      realFmPrev: 0,
      realMtdPrev: 0,
      realYtd: 0,
      shortageFm: 0,
      shortageYtd: 0,
      targetAch: null,
      targetFm: 0,
      targetYtd: 0,
    }),
    achYtd: safeRatioValue(realYtd, targetYtd),
    days,
    growthFm: safeGrowthValue(mtd, realFmPrev),
    growthMtd: safeGrowthValue(mtd, realMtdPrev),
    growthYtd: safeGrowthValue(realYtd, previousYtd),
    label,
    mtd,
    rankAch: null,
    rankYtd: null,
    realFmPrev,
    realMtdPrev,
    realYtd,
    shortageFm: mtd - targetFm,
    shortageYtd: realYtd - targetYtd,
    targetAch: safeRatioValue(mtd, targetFm),
    targetFm,
    targetYtd,
  };
}

function aggregateTable(
  parts: PeriodReportPart[],
  selectTable: (report: ComputedReport) => ReportTable,
  ytdParts = parts,
): ReportTable {
  const firstTable = selectTable(parts[0].report);
  const rows = firstTable.rows.map((row) =>
    aggregateRows(
      row.label,
      parts.map((part) => ({
        endDay: part.endDay,
        row: selectTable(part.report).rows.find((candidate) => candidate.label === row.label) ?? row,
        startDay: part.startDay,
        weight: part.weight,
      })),
      ytdParts.map((part) => ({
        endDay: part.endDay,
        row: selectTable(part.report).rows.find((candidate) => candidate.label === row.label) ?? row,
        startDay: part.startDay,
        weight: part.weight,
      })),
    ),
  );
  const rankedRows = rankReportRows(rankReportRows(rows, "targetAch"), "achYtd");
  return {
    ...firstTable,
    rows: rankedRows,
    total: aggregateRows(
      firstTable.total.label,
      parts.map((part) => ({
        endDay: part.endDay,
        row: selectTable(part.report).total,
        startDay: part.startDay,
        weight: part.weight,
      })),
      ytdParts.map((part) => ({
        endDay: part.endDay,
        row: selectTable(part.report).total,
        startDay: part.startDay,
        weight: part.weight,
      })),
    ),
  };
}

function reportForArchiveEntry(entry: PeriodArchiveEntry) {
  return entry.reportSnapshot ?? applyReportHistory(computeReport(entry.rows, entry.savedAt, entry.periodKey), entry.reportHistory);
}

function reportForPeriodFilter(
  baseReport: ComputedReport,
  periodArchive: PeriodArchive,
  selectedPeriodKey: string,
  filter: DashboardPeriodFilterState,
) {
  const normalized = normalizePeriodFilter(filter);
  const validationMessage = validatePeriodFilter(normalized);
  if (validationMessage) return { report: baseReport, validationMessage };

  const parts =
    normalized.mode === "date-range"
      ? dateRangeToMonthParts(normalized.startDate, normalized.endDate)
      : monthRangeKeys(normalized.startMonth, normalized.endMonth).map((periodKey) => ({
          endDay: new Date(Number(periodKey.slice(0, 4)), Number(periodKey.slice(4, 6)), 0).getDate(),
          periodKey,
          startDay: 1,
          weight: 1,
        }));

  if (parts.length <= 1 && parts[0]?.periodKey === selectedPeriodKey && parts[0]?.startDay === 1 && parts[0]?.weight === 1) {
    return { report: baseReport, validationMessage: null };
  }

  const reportParts = parts
    .map((part) => {
      const entry = periodArchive[part.periodKey];
      if (!entry) return null;
      return { ...part, report: reportForArchiveEntry(entry) };
    })
    .filter(Boolean) as PeriodReportPart[];

  if (reportParts.length === 0) return { report: baseReport, validationMessage: "Data periode yang dipilih belum tersedia." };

  const lastPart = reportParts[reportParts.length - 1];
  const ytdStartDate = `${lastPart.periodKey.slice(0, 4)}-01-01`;
  const ytdEndDate = `${lastPart.periodKey.slice(0, 4)}-${lastPart.periodKey.slice(4, 6)}-${String(lastPart.endDay).padStart(2, "0")}`;
  const ytdParts = dateRangeToMonthParts(ytdStartDate, ytdEndDate)
    .map((part) => {
      const entry = periodArchive[part.periodKey];
      if (!entry) return null;
      return { ...part, report: reportForArchiveEntry(entry) };
    })
    .filter(Boolean) as PeriodReportPart[];
  const effectiveYtdParts = ytdParts.length > 0 ? ytdParts : reportParts;

  const witel = aggregateTable(reportParts, (report) => report.witel, effectiveYtdParts);
  const branch = aggregateTable(reportParts, (report) => report.branch, effectiveYtdParts);
  const stfWitel = aggregateTable(reportParts, (report) => report.targetStf.witel, effectiveYtdParts);
  const stfBranch = aggregateTable(reportParts, (report) => report.targetStf.branch, effectiveYtdParts);
  const dailyTrend = witel.total.days.map((value, index) => ({
    cumulative: witel.total.days.slice(0, index + 1).reduce<number>((sum, dayValue) => sum + (dayValue ?? 0), 0),
    day: index + 1,
    realisasi: value ?? 0,
  }));
  const periodLabel =
    normalized.mode === "date-range"
      ? dateRangeLabel(normalized.startDate, normalized.endDate)
      : periodRangeLabel(reportParts[0].periodKey, reportParts[reportParts.length - 1].periodKey);
  const cutoffLabel = formatDataCutoffDate(latestActualDateFromHsiParts(reportParts));

  return {
    report: {
      ...baseReport,
      branch,
      cutoffLabel,
      dailyTrend,
      periodLabel,
      previousMonthLabel: baseReport.previousMonthLabel,
      targetStf: {
        branch: stfBranch,
        witel: stfWitel,
      },
      validation: { checked: 0, mismatches: 0, notes: [] },
      witel,
    },
    validationMessage: reportParts.length < parts.length ? "Sebagian data periode belum tersedia, hanya periode yang tersedia yang dihitung." : null,
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
  icon,
}: {
  label: string;
  value: string;
  detail: string;
  tone?: DashboardMetricTone;
  icon?: React.ComponentType<{ className?: string }>;
}) {
  return (
    <DashboardMetricCard
      detail={detail}
      icon={icon}
      label={label}
      theme={{
        accentLineClass: "bg-gradient-to-r from-slate-900 via-rose-500 to-emerald-500 opacity-75",
        detailToneClass: {
          good: "text-emerald-700",
          neutral: "text-slate-500",
          warn: "text-amber-700",
        },
        hoverClass: "hover:-translate-y-0.5 hover:border-white/30 hover:shadow-[0_24px_46px_-28px_rgba(0,0,0,0.7)]",
        iconToneClass: {
          good: "border-emerald-200 bg-emerald-50 text-emerald-700",
          neutral: "border-slate-200 bg-white text-slate-600",
          warn: "border-amber-200 bg-amber-50 text-amber-700",
        },
        labelClass: "text-slate-500",
        panelClass,
        toneClass: {
          good: "from-emerald-50/90 to-white text-emerald-700",
          neutral: "from-slate-50/90 to-white text-slate-600",
          warn: "from-amber-50/90 to-white text-amber-700",
        },
        valueClass: "text-slate-950",
      }}
      tone={tone}
      value={value}
    />
  );
}

type DashboardTargetMode = "rkap" | "stf";
type DashboardPeriodMode = "mtd" | "ytd";
type DashboardPeriodOption = { label: string; periodKey: string; savedAt: string | null };
type WilayahFilterKind = "all-telda" | "all-witel" | "telda" | "witel";
type WilayahFilterOption = {
  kind: WilayahFilterKind;
  label: string;
  rowLabel?: string;
  value: string;
};

const wilayahFilterOptions: WilayahFilterOption[] = [
  { kind: "all-telda", label: "ALL TELDA", value: "all-telda" },
  { kind: "all-witel", label: "ALL WITEL", value: "all-witel" },
  { kind: "telda", label: "WONOSOBO", value: "telda:WONOSOBO" },
  { kind: "telda", label: "PURBALINGGA", value: "telda:PURBALINGGA" },
  { kind: "telda", label: "PURWOKERTO", value: "telda:PURWOKERTO" },
  { kind: "telda", label: "CILACAP", value: "telda:CILACAP" },
  { kind: "telda", label: "MAGELANG", value: "telda:MAGELANG" },
  { kind: "telda", label: "YOGYAKARTA", rowLabel: "MEA YOGYAKARTA", value: "telda:MEA YOGYAKARTA" },
  { kind: "telda", label: "BANTUL", value: "telda:BANTUL" },
  { kind: "telda", label: "KEBUMEN", value: "telda:KEBUMEN" },
  { kind: "witel", label: "BALI", rowLabel: "BALI", value: "witel:BALI" },
  { kind: "witel", label: "JTB", rowLabel: "JATIM BARAT", value: "witel:JATIM BARAT" },
  { kind: "witel", label: "JTT", rowLabel: "JATIM TIMUR", value: "witel:JATIM TIMUR" },
  { kind: "witel", label: "NUSRA", rowLabel: "NUSA TENGGARA", value: "witel:NUSA TENGGARA" },
  { kind: "witel", label: "SCU", rowLabel: "SEMARANG JATENG UTARA", value: "witel:SEMARANG JATENG UTARA" },
  { kind: "witel", label: "SJT", rowLabel: "SOLO JATENG TIMUR", value: "witel:SOLO JATENG TIMUR" },
  { kind: "witel", label: "SURAMADU", rowLabel: "SURAMADU", value: "witel:SURAMADU" },
  { kind: "witel", label: "JJS", rowLabel: "YOGYA JATENG SELATAN", value: "witel:YOGYA JATENG SELATAN" },
];

const witelShortLabels: Record<string, string> = {
  BALI: "BALI",
  "JATIM TIMUR": "JTT",
  "JATIM BARAT": "JTB",
  "NUSA TENGGARA": "NUSRA",
  "SEMARANG JATENG UTARA": "SCU",
  "SOLO JATENG TIMUR": "SJT",
  "SOLO JATIMUR": "SJT",
  SURAMADU: "SURAMADU",
  "YOGYA JATENG SELATAN": "JJS",
};

function rowForWilayah(table: ReportTable, option: WilayahFilterOption) {
  if (option.kind === "all-telda" || option.kind === "all-witel") return table.total;
  const rowLabel = option.rowLabel ?? option.label;
  return table.rows.find((row) => row.label === rowLabel) ?? table.total;
}

function displayLabelForRow(row: ReportRow, option?: WilayahFilterOption) {
  if (option?.kind === "witel") return option.label;
  if (option?.kind === "telda" && option.label === "YOGYAKARTA") return "YOGYAKARTA";
  return witelShortLabels[row.label] ?? row.label;
}

function latestPopulatedDayForRows(rows: ReportRow[]) {
  let latest = 0;
  rows.forEach((row) => {
    row.days.forEach((value, index) => {
      if (hasActualValue(value)) latest = Math.max(latest, index + 1);
    });
  });
  return Math.max(1, latest);
}

function comparableGrowth(current: number, previous: number | null | undefined) {
  if (!previous || previous <= 0) return null;
  return (current - previous) / previous;
}

function rankByAchievement(rows: ReportRow[], label: string) {
  const ranked = [...rows].sort((a, b) => (b.targetAch ?? -1) - (a.targetAch ?? -1));
  const index = ranked.findIndex((row) => row.label === label);
  return index >= 0 ? index + 1 : null;
}

function shortageValue(row: ReportRow) {
  return row.shortageFm;
}

function chartLabelNumber(value: unknown) {
  return formatNumber(Number(value ?? 0));
}

function chartLabelPercent(value: unknown) {
  return `${Number(value ?? 0).toFixed(1)}%`;
}

function DashboardHome({
  onPeriodChange,
  onPeriodFilterChange,
  periodOptions,
  periodFilter,
  periodValidationMessage,
  report,
  selectedPeriodKey,
}: {
  onPeriodChange: (periodKey: string) => void;
  onPeriodFilterChange: (value: DashboardPeriodFilterState) => void;
  periodOptions: DashboardPeriodOption[];
  periodFilter: DashboardPeriodFilterState;
  periodValidationMessage: string | null;
  report: ComputedReport;
  selectedPeriodKey: string;
}) {
  const [wilayahFilter, setWilayahFilter] = useState("all-telda");
  const [analyticsFilter, setAnalyticsFilter] = useState<{ period: DashboardPeriodMode; targetType: DashboardTargetMode }>({
    period: "mtd",
    targetType: "rkap",
  });
  const [trendMode, setTrendMode] = useState<DashboardTargetMode>("rkap");
  const normalizedPeriodFilter = normalizePeriodFilter(periodFilter);
  const isSingleMonthPeriod =
    normalizedPeriodFilter.startMonth === normalizedPeriodFilter.endMonth &&
    normalizedPeriodFilter.startDate === dateInputForPeriodKey(normalizedPeriodFilter.startMonth, "first") &&
    normalizedPeriodFilter.endDate === dateInputForPeriodKey(normalizedPeriodFilter.endMonth, "last");
  const periodMetricLabel = getPeriodSemanticLabel(normalizedPeriodFilter.mode);

  const selectedWilayah = wilayahFilterOptions.find((option) => option.value === wilayahFilter) ?? wilayahFilterOptions[0];
  const isWitelScope = selectedWilayah.kind === "all-witel" || selectedWilayah.kind === "witel";
  const selectedRkapTable = isWitelScope ? report.witel : report.branch;
  const selectedStfTable = isWitelScope ? report.targetStf.witel : report.targetStf.branch;
  const selectedRkap = rowForWilayah(selectedRkapTable, selectedWilayah);
  const selectedStf = rowForWilayah(selectedStfTable, selectedWilayah);
  const latestDay = latestPopulatedDayForRows([selectedRkap, selectedStf]);
  const periodProgress = Math.round((latestDay / Math.max(selectedRkap.days.length, 1)) * 100);
  const isAllWilayah = selectedWilayah.kind === "all-telda" || selectedWilayah.kind === "all-witel";
  const activeScopeLabel = selectedWilayah.label;
  const activeScopeTypeLabel = isWitelScope ? "WITEL" : "TELDA";
  const shortageRkap = shortageValue(selectedRkap);
  const shortageStf = selectedStf.shortageFm;
  const analyticsTable = analyticsFilter.targetType === "rkap" ? selectedRkapTable : selectedStfTable;
  const analyticsSelectedRow = rowForWilayah(analyticsTable, selectedWilayah);
  const analyticsAchievement = (row: ReportRow) => (analyticsFilter.period === "mtd" ? row.targetAch : row.achYtd) ?? 0;
  const ranking = [...analyticsTable.rows].sort((a, b) => analyticsAchievement(b) - analyticsAchievement(a));
  const selectedRank = isAllWilayah ? null : rankByAchievement(selectedRkapTable.rows, selectedRkap.label);
  const growth = comparableGrowth(selectedRkap.mtd, selectedRkap.realMtdPrev);
  const topWilayah = ranking[0];
  const allTeldaRegionLabel = selectedWilayah.kind === "all-telda" ? " (YOGYA JATENG SELATAN)" : "";
  const darkSummaryTitle = isAllWilayah
    ? `HSI PERFORMA ${activeScopeLabel}${allTeldaRegionLabel}`
    : `HSI PERFORMA ${activeScopeTypeLabel} ${activeScopeLabel}`;
  const darkSummaryScope = `CAKUPAN: ${activeScopeLabel}`;
  const targetChartPeriodLabel = analyticsFilter.period === "mtd" ? periodMetricLabel : "YTD";
  const analyticsTargetLabel = analyticsFilter.targetType.toUpperCase();
  const dataCutoffLabel = report.cutoffLabel.startsWith("DATA S.D.")
    ? report.cutoffLabel
    : formatDataCutoff({ day: latestDay, periodKey: normalizedPeriodFilter.endMonth });
  const ytdPeriodLabel = formatYtdLabel(normalizedPeriodFilter.endDate);
  const summaryMtdRankingRows = selectedRkapTable.rows.map((row) => ({
    achievement: row.targetAch,
    id: row.label,
    label: displayLabelForRow(row),
  }));
  const summaryYtdRankingRows = selectedRkapTable.rows.map((row) => ({
    achievement: row.achYtd,
    id: row.label,
    label: displayLabelForRow(row),
  }));
  const targetChartValues = (row: ReportRow) => {
    const target = analyticsFilter.period === "mtd" ? row.targetFm : row.targetYtd;
    const realisasi = analyticsFilter.period === "mtd" ? row.mtd : row.realYtd;
    const achievement = target > 0 ? realisasi / target : 0;
    return {
      achievement: Number((achievement * 100).toFixed(2)),
      realisasi,
      shortage: Math.max(target - realisasi, 0),
      target,
    };
  };

  const targetChart =
    isAllWilayah
      ? analyticsTable.rows.map((row) => ({
          ...targetChartValues(row),
          name: displayLabelForRow(row),
        }))
      : [
          {
            ...targetChartValues(analyticsSelectedRow),
            name: activeScopeLabel,
          },
        ];

  const trendTargetRow = trendMode === "rkap" ? selectedRkap : selectedStf;
  const targetPerDay = trendTargetRow.targetFm / Math.max(trendTargetRow.days.length, 1);
  const trendData = trendTargetRow.days.slice(0, latestDay).map((value, index) => {
    const day = index + 1;
    const realisasiKumulatif = selectedRkap.days
      .slice(0, day)
      .reduce<number>((sum, item) => sum + (typeof item === "number" ? item : 0), 0);
    return {
      day,
      realisasi: typeof value === "number" ? value : null,
      realisasiKumulatif,
      targetKumulatif: Number((targetPerDay * day).toFixed(2)),
    };
  });

  return (
    <div className="space-y-5">
      <section className={`${panelClass} relative overflow-hidden`}>
        <div className="absolute inset-0 bg-[radial-gradient(circle_at_88%_10%,rgba(56,189,248,0.16),transparent_28%),radial-gradient(circle_at_8%_18%,rgba(37,99,235,0.08),transparent_24%),linear-gradient(135deg,#ffffff_0%,#f8fbff_52%,#eef6ff_100%)]" />
        <div className="absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-[#071426] via-[#2563eb] to-[#38bdf8]" />
        <div className="relative flex flex-wrap items-center justify-between gap-3 px-4 py-3">
          <div className="min-w-[280px] flex-1">
            <div>
              <p className="inline-flex items-center gap-2 text-[11px] font-semibold uppercase text-blue-700">
                <Activity className="h-4 w-4" />
                DASHBOARD REPORTING HSI 2026
              </p>
              <h1 className="mt-1 text-xl font-semibold uppercase tracking-normal text-[#071426] sm:text-2xl">
                HSI PERFORMANCE
              </h1>
              <p className="mt-1 max-w-2xl text-xs font-medium uppercase leading-5 text-slate-500">
                {report.periodLabel} | {activeScopeLabel} | {dataCutoffLabel}
              </p>
            </div>
          </div>
          <div className="flex flex-wrap items-end gap-2">
            <DashboardPeriodFilter
              onChange={(value) => {
                const normalized = normalizePeriodFilter(value);
                onPeriodFilterChange(normalized);
                const anchorPeriod = normalized.mode === "date-range" ? normalized.endMonth : normalized.endMonth;
                if (anchorPeriod !== selectedPeriodKey && periodOptions.some((option) => option.periodKey === anchorPeriod)) {
                  onPeriodChange(anchorPeriod);
                }
              }}
              options={periodOptions}
              theme="hsi"
              validationMessage={periodValidationMessage}
              value={periodFilter}
            />
            <label className="grid gap-1 text-[11px] font-semibold uppercase text-slate-500">
              WILAYAH
              <select
                className="h-9 w-[210px] rounded-md border border-[#d8e1ec] bg-white px-3 text-xs font-semibold text-[#071426] shadow-sm outline-none transition hover:border-blue-300 focus:border-blue-400 focus:ring-4 focus:ring-blue-100"
                value={wilayahFilter}
                onChange={(event) => setWilayahFilter(event.target.value)}
              >
                <option value="all-telda">ALL TELDA</option>
                <option value="all-witel">ALL WITEL</option>
                <optgroup label="TELDA">
                  {wilayahFilterOptions
                    .filter((option) => option.kind === "telda")
                    .map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                </optgroup>
                <optgroup label="WITEL">
                  {wilayahFilterOptions
                    .filter((option) => option.kind === "witel")
                    .map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                </optgroup>
              </select>
            </label>
            <StatusPill tone={report.validation.mismatches === 0 ? "good" : "warn"}>
              {report.validation.mismatches === 0 ? "VALIDASI MATCH" : `${report.validation.mismatches} SELISIH`}
            </StatusPill>
          </div>
        </div>
      </section>

      <section className={`${panelClass} overflow-hidden`}>
        <div className="relative grid items-start gap-4 bg-[linear-gradient(135deg,#ffffff_0%,#f8fbff_48%,#f4f8ff_100%)] p-4 xl:grid-cols-[1.35fr_0.65fr]">
          <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(37,99,235,0.035)_1px,transparent_1px),linear-gradient(90deg,rgba(37,99,235,0.03)_1px,transparent_1px)] bg-[size:44px_44px] opacity-50" />
          <div className="relative overflow-hidden rounded-xl border border-slate-200/90 bg-white shadow-[0_16px_36px_-28px_rgba(15,23,42,0.65)]">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-slate-200 bg-[#071426] px-3 py-2.5 text-white">
              <div>
                <p className="text-[11px] font-semibold uppercase text-rose-200">HSI PERFORMANCE SUMMARY</p>
                <h2 className="mt-0.5 text-sm font-semibold tracking-normal">PERIODE {report.periodLabel} | {activeScopeLabel}</h2>
              </div>
              <StatusPill>{dataCutoffLabel}</StatusPill>
            </div>
            <div className="grid gap-2.5 p-2.5 xl:grid-cols-2">
              <div className="overflow-hidden rounded-md border border-slate-200 bg-white shadow-sm">
                <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-2.5 py-2">
                  <div>
                    <p className="text-[10px] font-semibold uppercase text-rose-700">HSI PERFORMA {periodMetricLabel} ({report.periodLabel})</p>
                    <p className="mt-0.5 text-xs font-semibold uppercase text-slate-950">{report.periodLabel} | {activeScopeLabel}</p>
                  </div>
                  <span className="rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-semibold uppercase text-slate-600">
                    RKAP
                  </span>
                </div>
                <div className="grid gap-2 p-2.5">
                  <div className="grid gap-2 sm:grid-cols-2">
                    <div className="rounded-md border border-slate-200 bg-white p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-slate-500">TARGET RKAP {periodMetricLabel}</p>
                      <p className="mt-0.5 text-base font-semibold text-slate-950">{formatNumber(selectedRkap.targetFm)}</p>
                    </div>
                    <div className="rounded-md border border-emerald-100 bg-emerald-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-emerald-700">REALISASI {periodMetricLabel}</p>
                      <p className="mt-0.5 text-base font-semibold text-emerald-950">{formatNumber(selectedRkap.mtd)}</p>
                    </div>
                    <div className="rounded-md border border-amber-100 bg-amber-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-amber-700">ACH RKAP {periodMetricLabel}</p>
                      <p className="mt-0.5 text-base font-semibold text-amber-950">{formatPercent(selectedRkap.targetAch)}</p>
                    </div>
                    <div className="rounded-md border border-rose-100 bg-rose-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-rose-700">SHORTAGE RKAP {periodMetricLabel}</p>
                      <p className="mt-0.5 text-base font-semibold text-rose-950">{formatNumber(shortageRkap)}</p>
                    </div>
                  </div>
                  <div className="rounded-md border border-slate-200 bg-slate-50/80 p-2">
                    <div className="flex items-center justify-between text-[11px] font-semibold uppercase">
                      <span className="text-slate-500">ACHIEVEMENT {periodMetricLabel}</span>
                      <span className="text-emerald-700">{formatPercent(selectedRkap.targetAch)}</span>
                    </div>
                    <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-slate-200">
                      <div
                        className="h-full rounded-full bg-gradient-to-r from-rose-500 via-amber-400 to-emerald-500"
                        style={{ width: `${Math.min((selectedRkap.targetAch ?? 0) * 100, 100)}%` }}
                      />
                    </div>
                    <CompactRankingInsight
                      formatPercent={formatPercent}
                      isAllScope={isAllWilayah}
                      periodLabel={periodMetricLabel}
                      rows={summaryMtdRankingRows}
                      scopeLabel={activeScopeTypeLabel}
                      selectedId={selectedRkap.label}
                    />
                  </div>
                </div>
              </div>

              <div className="overflow-hidden rounded-md border border-slate-200 bg-white shadow-sm">
                <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-2.5 py-2">
                  <div>
                    <p className="text-[10px] font-semibold uppercase text-rose-700">HSI PERFORMA YTD (TAHUN BERJALAN)</p>
                    <p className="mt-0.5 text-xs font-semibold uppercase text-slate-950">{ytdPeriodLabel} | {activeScopeLabel}</p>
                  </div>
                  <span className="rounded-md border border-slate-200 bg-slate-50 px-2.5 py-1 text-[11px] font-semibold uppercase text-slate-600">
                    RKAP
                  </span>
                </div>
                <div className="grid gap-2 p-2.5">
                  <div className="grid gap-2 sm:grid-cols-2">
                    <div className="rounded-md border border-slate-200 bg-white p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-slate-500">TARGET RKAP YTD</p>
                      <p className="mt-0.5 text-base font-semibold text-slate-950">{formatNumber(selectedRkap.targetYtd)}</p>
                    </div>
                    <div className="rounded-md border border-emerald-100 bg-emerald-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-emerald-700">REALISASI YTD</p>
                      <p className="mt-0.5 text-base font-semibold text-emerald-950">{formatNumber(selectedRkap.realYtd)}</p>
                    </div>
                    <div className="rounded-md border border-amber-100 bg-amber-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-amber-700">ACH RKAP YTD</p>
                      <p className="mt-0.5 text-base font-semibold text-amber-950">{formatPercent(selectedRkap.achYtd)}</p>
                    </div>
                    <div className="rounded-md border border-rose-100 bg-rose-50/70 p-2 shadow-sm">
                      <p className="text-[10px] font-semibold uppercase text-rose-700">SHORTAGE RKAP YTD</p>
                      <p className="mt-0.5 text-base font-semibold text-rose-950">{formatNumber(selectedRkap.shortageYtd)}</p>
                    </div>
                  </div>
                  <div className="rounded-md border border-slate-200 bg-slate-50/80 p-2">
                    <div className="flex items-center justify-between text-[11px] font-semibold uppercase">
                      <span className="text-slate-500">ACHIEVEMENT YTD</span>
                      <span className="text-emerald-700">{formatPercent(selectedRkap.achYtd)}</span>
                    </div>
                    <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-slate-200">
                      <div
                        className="h-full rounded-full bg-gradient-to-r from-rose-500 via-amber-400 to-emerald-500"
                        style={{ width: `${Math.min((selectedRkap.achYtd ?? 0) * 100, 100)}%` }}
                      />
                    </div>
                    <CompactRankingInsight
                      formatPercent={formatPercent}
                      isAllScope={isAllWilayah}
                      periodLabel="YTD"
                      rows={summaryYtdRankingRows}
                      scopeLabel={activeScopeTypeLabel}
                      selectedId={selectedRkap.label}
                    />
                  </div>
                </div>
              </div>
            </div>
          </div>
          <div className="relative self-start overflow-hidden rounded-xl border border-cyan-300/20 bg-[linear-gradient(145deg,#020617_0%,#071426_48%,#0a1f3b_100%)] p-5 text-white shadow-[0_30px_78px_-42px_rgba(56,189,248,0.55)]">
            <div className="absolute inset-0 bg-[radial-gradient(circle_at_86%_4%,rgba(56,189,248,0.18),transparent_30%)]" />
            <div className="absolute inset-x-0 top-0 h-px bg-gradient-to-r from-transparent via-cyan-300 to-transparent" />
            <div className="relative">
              <div className="flex items-center justify-between gap-4">
                <p className="max-w-[250px] text-sm font-semibold uppercase leading-5 text-slate-100">
                  {darkSummaryTitle}
                </p>
                <span className="rounded-md border border-emerald-400/30 bg-emerald-400/10 px-2 py-1 text-[11px] font-semibold text-emerald-200">
                  SOURCE: XPRO
                </span>
              </div>
              <p className="mt-3 text-[11px] font-semibold uppercase text-slate-500">{darkSummaryScope}</p>
              <p className="mt-3 text-xs font-semibold uppercase text-slate-400">TOTAL REALISASI {periodMetricLabel}</p>
              <p className="mt-1.5 text-4xl font-semibold tracking-normal text-white">{formatNumber(selectedRkap.mtd)}</p>
              <div className="mt-4 h-2 overflow-hidden rounded-full bg-white/10">
                <div className="h-full rounded-full bg-gradient-to-r from-rose-500 via-amber-400 to-emerald-400" style={{ width: `${Math.min(periodProgress, 100)}%` }} />
              </div>
              <div className="mt-4 grid grid-cols-2 gap-2.5 text-sm">
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-slate-400">ACH RKAP</p>
                  <p className="mt-1 text-base font-semibold text-white">{formatPercent(selectedRkap.targetAch)}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-slate-400">ACH STF</p>
                  <p className="mt-1 text-base font-semibold text-white">{formatPercent(selectedStf.targetAch)}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-slate-400">SHORTAGE RKAP</p>
                  <p className="mt-1 text-base font-semibold text-amber-200">{formatNumber(shortageRkap)}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-slate-400">SHORTAGE STF</p>
                  <p className="mt-1 text-base font-semibold text-amber-200">{formatNumber(shortageStf)}</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <div className="grid gap-2.5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
        <MetricCard label="PERFORMANCE BY RKAP 2026" value={formatNumber(selectedRkap.targetYtd)} detail="TARGET RKAP YTD 2026" icon={Target} />
        <MetricCard label={`REALISASI ${periodMetricLabel}`} value={formatNumber(selectedRkap.mtd)} detail={`UPDATE ${report.cutoffLabel}`} tone="good" icon={TrendingUp} />
        <MetricCard label="ACH RKAP" value={formatPercent(selectedRkap.targetAch)} detail="REALISASI / TARGET RKAP" icon={ArrowUpRight} />
        <MetricCard label={isSingleMonthPeriod ? "SHORTAGE CURRENT MONTH RKAP" : "SHORTAGE SELECTED PERIOD RKAP"} value={formatNumber(shortageRkap)} detail={`REALISASI ${periodMetricLabel} - TARGET RKAP ${periodMetricLabel}`} tone="warn" icon={FileSpreadsheet} />
        <MetricCard label={isSingleMonthPeriod ? "GROWTH VS LAST MONTH" : "GROWTH VS COMPARABLE PERIOD"} value={growth === null ? "N/A" : formatPercent(growth)} detail={isSingleMonthPeriod ? "DIBANDINGKAN MTD BULAN SEBELUMNYA" : "DIBANDINGKAN PERIODE PEMBANDING"} icon={Gauge} />
        <MetricCard
          label="RANKING WILAYAH"
          value={isAllWilayah ? (topWilayah ? displayLabelForRow(topWilayah) : "N/A") : selectedRank ? `#${selectedRank}` : "N/A"}
          detail={isAllWilayah ? `TOP ${activeScopeTypeLabel} ${formatPercent(topWilayah?.targetAch)}` : `DARI ${selectedRkapTable.rows.length} ${activeScopeTypeLabel}`}
          icon={ShieldCheck}
        />
      </div>

      <section className={`${panelClass} overflow-hidden bg-white`}>
        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-[#eaeff5] px-6 py-5">
          <div>
            <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-blue-600">Performance Analytics</p>
            <h2 className="mt-1 text-[28px] font-bold uppercase leading-tight tracking-[-0.02em] text-slate-950">Regional Performance</h2>
            <p className="mt-1 text-sm font-medium text-slate-500">Ringkasan performa dan ranking achievement per TELDA</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <select
              className="h-11 rounded-xl border border-slate-200 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-blue-400"
              value={analyticsFilter.period}
              onChange={(event) => setAnalyticsFilter((current) => ({ ...current, period: event.target.value as DashboardPeriodMode }))}
            >
              <option value="mtd">{periodMetricLabel}</option>
              <option value="ytd">YTD</option>
            </select>
            <select
              className="h-11 rounded-xl border border-slate-200 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-blue-400"
              value={analyticsFilter.targetType}
              onChange={(event) => setAnalyticsFilter((current) => ({ ...current, targetType: event.target.value as DashboardTargetMode }))}
            >
              <option value="rkap">RKAP</option>
              <option value="stf">STF</option>
            </select>
            <StatusPill>{report.periodLabel}</StatusPill>
          </div>
        </div>
        <div className="grid xl:grid-cols-[1.45fr_0.95fr]">
          <div className="p-4">
            <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-xl font-bold uppercase text-slate-950">TARGET VS REALISASI PER WILAYAH</h2>
              <p className="mt-1 text-[13px] font-medium uppercase text-slate-500">
                INDIKATOR {targetChartPeriodLabel} {analyticsTargetLabel}: TARGET, REALISASI, SHORTAGE, DAN ACHIEVEMENT PER {activeScopeTypeLabel}
              </p>
            </div>
          </div>
          <div className="h-[360px]">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={targetChart} margin={{ left: 0, right: 22, top: 24, bottom: isAllWilayah ? 84 : 42 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis
                  dataKey="name"
                  interval={0}
                  tick={{ angle: isAllWilayah ? -42 : 0, dx: isAllWilayah ? -10 : 0, dy: isAllWilayah ? 20 : 0, fontSize: isAllWilayah ? 8 : 12, fontWeight: 700, textAnchor: isAllWilayah ? "end" : "middle" }}
                  tickMargin={isAllWilayah ? 18 : 6}
                />
                <YAxis yAxisId="left" tick={{ fontSize: 12 }} />
                <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 12 }} unit="%" />
                <Tooltip
                  contentStyle={chartTooltipStyle}
                  formatter={(value, name) => (name === "ACHIEVEMENT" ? tooltipPercent(value) : tooltipNumber(value))}
                />
                <Bar yAxisId="left" dataKey="target" name="TARGET" fill="#64748b" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="target" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Bar yAxisId="left" dataKey="realisasi" name="REALISASI" fill="#f97316" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="realisasi" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Bar yAxisId="left" dataKey="shortage" name="SHORTAGE" fill="#e11d48" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="shortage" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Line yAxisId="right" type="monotone" dataKey="achievement" name="ACHIEVEMENT" stroke="#059669" strokeWidth={3} dot={{ r: 3 }}>
                  <LabelList dataKey="achievement" position="top" formatter={chartLabelPercent} fontSize={10} fill="#047857" />
                </Line>
              </ComposedChart>
            </ResponsiveContainer>
          </div>
          <div className="mt-2 flex flex-wrap items-center justify-center gap-x-4 gap-y-2 text-xs font-bold uppercase">
            <span className="inline-flex items-center gap-1.5 text-emerald-600">
              <span className="h-0.5 w-5 rounded-full bg-emerald-600" />
              Achievement
            </span>
            <span className="inline-flex items-center gap-1.5 text-orange-600">
              <span className="h-3 w-3 rounded-sm bg-orange-500" />
              Realisasi
            </span>
            <span className="inline-flex items-center gap-1.5 text-rose-600">
              <span className="h-3 w-3 rounded-sm bg-rose-600" />
              Shortage
            </span>
            <span className="inline-flex items-center gap-1.5 text-slate-600">
              <span className="h-3 w-3 rounded-sm bg-slate-500" />
              Target
            </span>
          </div>
          </div>

        <div className="border-t border-slate-100 p-4 xl:border-l xl:border-t-0">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-xl font-bold uppercase text-slate-950">RANKING ACHIEVEMENT</h2>
              <p className="mt-1 text-[13px] font-medium uppercase text-slate-500">RANKING BERDASARKAN ACH {analyticsTargetLabel} {targetChartPeriodLabel}</p>
            </div>
            <ShieldCheck className="h-5 w-5 text-emerald-600" />
          </div>
          <div className="mt-3 divide-y divide-slate-100">
            {ranking.map((row) => (
              <div
                key={row.label}
                className={`grid grid-cols-[34px_1fr_auto] items-center gap-3 py-2.5 text-sm ${row.label === analyticsSelectedRow.label ? "rounded-md bg-rose-50/70 px-2" : ""}`}
              >
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-slate-950 text-xs font-semibold text-white shadow-sm">
                  {ranking.findIndex((item) => item.label === row.label) + 1}
                </span>
                <div className="min-w-0">
                  <span className="font-medium uppercase text-slate-800">{displayLabelForRow(row)}</span>
                  <div className="mt-1.5 h-1.5 overflow-hidden rounded-full bg-slate-100">
                    <div className="h-full rounded-full bg-gradient-to-r from-rose-500 to-emerald-500" style={{ width: `${Math.min(analyticsAchievement(row) * 100, 100)}%` }} />
                  </div>
                </div>
                <span className="font-semibold text-rose-700">{formatPercent(analyticsAchievement(row))}</span>
              </div>
            ))}
          </div>
        </div>
        </div>
      </section>

      <section className={`${panelClass} p-4`}>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-lg font-semibold uppercase text-slate-950">TREN TARGET HARIAN VS REALISASI HARIAN</h2>
            <p className="mt-1 text-xs font-medium uppercase text-slate-500">TARGET KUMULATIF HARIAN VS REALISASI KUMULATIF HARIAN</p>
          </div>
          <select
            className="h-9 rounded-md border border-slate-200 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-blue-400 focus:ring-4 focus:ring-blue-100"
            value={trendMode}
            onChange={(event) => setTrendMode(event.target.value as DashboardTargetMode)}
          >
            <option value="rkap">RKAP</option>
            <option value="stf">STF</option>
          </select>
        </div>
        <div className="mt-3 h-[350px]">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={trendData} margin={{ left: 0, right: 24, top: 16, bottom: 8 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="day" tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip contentStyle={chartTooltipStyle} formatter={(value) => tooltipNumber(value)} labelFormatter={(day) => `TANGGAL ${day}`} />
              <Legend wrapperStyle={{ fontSize: 12, fontWeight: 700 }} />
              <Line type="monotone" dataKey="targetKumulatif" name="TARGET KUMULATIF HARIAN" stroke="#64748b" strokeWidth={3} dot={false} />
              <Line type="monotone" dataKey="realisasiKumulatif" name="REALISASI KUMULATIF HARIAN" stroke="#0f766e" strokeWidth={3} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </div>
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
  exportStartDay: number,
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
  const exportDayCount = latestDay - exportStartDay + 1;
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
  drawCell(context, "TANGGAL", cursorX, headerY, exportDayCount * dayWidth, groupHeaderHeight, {
    align: "center",
    border: "#0f766e",
    fill: "#0f766e",
    font: "700 11px Arial",
    textColor: "#ffffff",
  });
  for (let day = exportStartDay; day <= latestDay; day += 1) {
    drawCell(context, String(day), cursorX + (day - exportStartDay) * dayWidth, headerY + groupHeaderHeight, dayWidth, dayHeaderHeight, {
      align: "center",
      border: "#99f6e4",
      fill: "#ccfbf1",
      font: "700 10px Arial",
      textColor: "#134e4a",
    });
  }
  cursorX += exportDayCount * dayWidth;
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
    row.days.slice(exportStartDay - 1, latestDay).forEach((value) => {
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
  const exportStartDay = Math.max(1, latestDay - 6);
  const exportDayCount = latestDay - exportStartDay + 1;
  const tableWidth = labelWidth + targetWidth * 2 + exportDayCount * dayWidth + summaryHeaders.length * summaryWidth;
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
  context.fillText(variant === "stf" ? "Report HSI Target STF" : "Report HSI Target RKAP", padding, padding);
  context.fillStyle = "#475569";
  context.font = "15px Arial";
  context.fillText(`Downloaded: ${formatJakartaDateTime(timestamp, false)}`, padding, padding + 36);
  context.fillText(`Periode: ${periodLabel} | Tanggal ${exportStartDay}-${latestDay}`, padding + 360, padding + 36);
  context.fillStyle = "#881337";
  context.font = "700 12px Arial";
  context.textAlign = "right";
  context.fillText("CONFIDENTIAL - INTERNAL USE ONLY", logicalWidth - padding, padding + 8);
  context.textAlign = "left";
  const firstTableHeight = drawReportTableCanvas(context, witelTable, exportStartDay, latestDay, periodLabel, previousMonthLabel, padding, padding + 68, tableWidth);
  drawReportTableCanvas(context, branchTable, exportStartDay, latestDay, periodLabel, previousMonthLabel, padding, padding + 68 + firstTableHeight + 22, tableWidth);
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
            {isStf ? "Report HSI Target STF" : "Report HSI Target RKAP"}
          </p>
          <h1 className="mt-1 text-3xl font-semibold tracking-normal text-slate-950">
            {isStf ? "Dua Tabel Utama Target STF" : "Dua Tabel Utama Target RKAP"}
          </h1>
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

const fullAccessPermissions: DashboardPermissions = {
  canAccessMaster: true,
  canEditMaster: true,
  canViewDashboard: true,
};

export function HsiDashboard({ permissions = fullAccessPermissions }: { permissions?: DashboardPermissions }) {
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
  const [masterUnlockToken, setMasterUnlockToken] = useState<string | null>(null);
  const [masterUnlockOpen, setMasterUnlockOpen] = useState(false);
  const [periodFilter, setPeriodFilter] = useState<DashboardPeriodFilterState>(() => defaultPeriodFilter(currentJakartaPeriodKey()));
  const visibleNavItems = useMemo(
    () => navItems.filter((item) => item.id !== "master" || permissions.canAccessMaster),
    [permissions.canAccessMaster],
  );
  const renderView: View = view === "master" && (!permissions.canAccessMaster || !masterUnlockToken) ? "home" : view;

  const applyPersistedState = useCallback((state: DashboardPersistedState, status: string) => {
    const activePeriod = state.periodArchive[state.selectedPeriodKey];
    if (!activePeriod) return;
    setPeriodArchive(state.periodArchive);
    setSelectedPeriodKey(state.selectedPeriodKey);
    setRows(activePeriod.rows);
    setSavedAt(activePeriod.savedAt);
    setReportHistory(activePeriod.reportHistory);
    setSaveSnapshots(state.saveSnapshots);
    setMasterHeadersCleared(state.masterHeadersCleared);
    setDirty(false);
    setAutoSaveStatus(status);
    persistDashboardStateLocally(state);
  }, []);

  const syncDashboardState = useCallback((state: DashboardPersistedState, successStatus: string) => {
    void saveDashboardStateToServer(state, masterUnlockToken)
      .then((result) => {
        setAutoSaveStatus(result.configured ? successStatus : "Data tersimpan di browser; database server belum dikonfigurasi");
      })
      .catch(() => {
        setAutoSaveStatus("Data tersimpan di browser; sync server gagal");
      });
  }, [masterUnlockToken]);

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
      void loadDashboardStateFromServer()
        .then(({ configured, state }) => {
          if (state) {
            applyPersistedState(state, "Data server dimuat");
          } else if (!configured) {
            setAutoSaveStatus("Database server belum dikonfigurasi; data tersimpan di browser ini");
          }
        })
        .catch(() => {
          setAutoSaveStatus("Data lokal dimuat; server belum bisa diakses");
        });
      setHydrated(true);
    }, 0);
    return () => window.clearTimeout(timeout);
  }, [applyPersistedState]);

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

  const filteredReportState = useMemo(
    () => reportForPeriodFilter(report, periodArchive, selectedPeriodKey, periodFilter),
    [periodArchive, periodFilter, report, selectedPeriodKey],
  );
  const displayReport = filteredReportState.report;

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
    const persistedState: DashboardPersistedState = {
      dataVersion: seedData.dataVersion,
      masterHeadersCleared,
      periodArchive: nextArchive,
      saveSnapshots: nextSnapshots,
      selectedPeriodKey: periodKey,
      updatedAt: timestamp,
    };
    persistDashboardStateLocally(persistedState);
    setPeriodArchive(nextArchive);
    setReportHistory(nextHistory);
    setSelectedPeriodKey(periodKey);
    setSavedAt(timestamp);
    setSaveSnapshots(nextSnapshots);
    setDirty(false);
    const status = options.automated ? `Auto-save tersimpan: ${formatJakartaDateTime(new Date(timestamp), true)}` : "Data tersimpan manual";
    setAutoSaveStatus(status);
    syncDashboardState(persistedState, `${status} dan tersinkron ke server`);
  }, [masterHeadersCleared, periodArchive, reportHistory, rows, saveSnapshots, selectedPeriodKey, syncDashboardState]);

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
    const persistedState: DashboardPersistedState = {
      dataVersion: seedData.dataVersion,
      masterHeadersCleared: false,
      periodArchive: nextArchive,
      saveSnapshots: nextSnapshots,
      selectedPeriodKey,
      updatedAt: previous.savedAt,
    };
    setRows(previous.rows);
    setSavedAt(previous.savedAt);
    setReportHistory(previous.reportHistory);
    setPeriodArchive(nextArchive);
    setSaveSnapshots(nextSnapshots);
    setMasterHeadersCleared(false);
    setDirty(false);
    setAutoSaveStatus("Kembali ke save sebelumnya");
    persistDashboardStateLocally(persistedState);
    syncDashboardState(persistedState, "Save sebelumnya tersinkron ke server");
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

  function updatePeriodFilter(value: DashboardPeriodFilterState) {
    const normalized = normalizePeriodFilter(value);
    setPeriodFilter(normalized);
  }

  function openMasterView() {
    if (!permissions.canAccessMaster) return;
    if (!masterUnlockToken) {
      setMasterUnlockOpen(true);
      return;
    }
    setView("master");
  }

  return (
    <main className="min-h-screen bg-[#071426] bg-[radial-gradient(circle_at_90%_8%,rgba(56,189,248,0.16),transparent_32%),radial-gradient(circle_at_8%_24%,rgba(37,99,235,0.12),transparent_28%),linear-gradient(135deg,#061426_0%,#0a1f3b_50%,#07182d_100%)] text-slate-950">
      <header className="sticky top-0 z-40 border-b border-cyan-300/20 bg-[#071426]/95 bg-[radial-gradient(circle_at_12%_0%,rgba(56,189,248,0.20),transparent_28%),linear-gradient(135deg,#061426_0%,#0a1f3b_54%,#07182d_100%)] shadow-[0_22px_48px_-34px_rgba(0,0,0,0.92)] backdrop-blur-xl">
        <div className="mx-auto flex max-w-[1920px] flex-col gap-3 px-4 py-3 sm:px-6 xl:flex-row xl:items-center xl:justify-between">
          <div className="flex items-center justify-between gap-4">
            <div className="flex items-center gap-3 rounded-xl border border-white/10 bg-white/[0.07] px-2.5 py-2 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)]">
              <div className="flex h-11 w-11 items-center justify-center overflow-hidden rounded-lg bg-white/[0.08] text-sm font-bold text-white shadow-[0_14px_30px_-16px_rgba(56,189,248,0.8)] ring-1 ring-cyan-300/25">
                <Image
                  alt="HSI"
                  className="h-full w-full rounded-lg object-cover"
                  height={80}
                  priority
                  src="/module-covers/hsi-design-card.png"
                  width={80}
                />
              </div>
              <div>
                <p className="text-sm font-bold uppercase tracking-normal text-white">Reporting HSI 2026</p>
                <p className="text-xs font-semibold uppercase text-cyan-100/72">{hydrated ? `${formatNumber(rows.length)} BARIS MASTER` : "MEMUAT DATA"}</p>
              </div>
            </div>
          </div>
          <div className="flex min-w-0 flex-1 items-center gap-3 xl:justify-end">
            <nav className="scrollbar-thin flex min-w-0 flex-1 gap-1.5 overflow-x-auto rounded-xl border border-white/10 bg-white/[0.08] p-1.5 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)] backdrop-blur xl:flex-none">
            <button
              onClick={() => router.push("/portal")}
              className="inline-flex h-9 shrink-0 items-center gap-2 rounded-md px-3 text-sm font-semibold text-slate-100/85 transition hover:bg-white/12 hover:text-white"
            >
              <ArrowLeft className="h-4 w-4" />
              Portal
            </button>
            {visibleNavItems.map((item) => {
              const Icon = item.icon;
              const active = renderView === item.id;
              const masterLocked = item.id === "master" && !masterUnlockToken;
              return (
                <button
                  key={item.id}
                  onClick={() => {
                    if (item.id === "master") {
                      openMasterView();
                    } else {
                      setMasterUnlockToken(null);
                      setView(item.id);
                    }
                  }}
                  className={`inline-flex h-9 shrink-0 items-center gap-2 rounded-md px-3 text-sm font-semibold transition ${
                    active
                      ? "bg-white text-[#071426] shadow-[0_16px_30px_-20px_rgba(56,189,248,0.8)]"
                      : "text-slate-100/85 hover:bg-white/12 hover:text-white"
                  }`}
                >
                  <Icon className="h-4 w-4" />
                  <span>{item.label}</span>
                  {masterLocked && <LockKeyhole aria-label="Master terkunci" className="h-3.5 w-3.5 opacity-70" />}
                </button>
              );
            })}
            </nav>
          </div>
        </div>
      </header>

      <div className="mx-auto max-w-[1920px] px-3 py-5 sm:px-5 lg:px-6">
          {renderView === "home" && (
            <DashboardHome
              onPeriodChange={selectPeriod}
              onPeriodFilterChange={updatePeriodFilter}
              periodOptions={periodOptions}
              periodFilter={periodFilter}
              periodValidationMessage={filteredReportState.validationMessage}
              report={displayReport}
              selectedPeriodKey={selectedPeriodKey}
            />
          )}
          {renderView === "master" && permissions.canAccessMaster && (
            <MasterPage
              autoSaveStatus={autoSaveStatus}
              canGoBack={(saveSnapshots[selectedPeriodKey] ?? []).length > 0}
              headersCleared={masterHeadersCleared}
              periodLabel={displayReport.periodLabel}
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
          {renderView === "report" && <ReportPage report={displayReport} variant="sementara" />}
          {renderView === "stf" && <ReportPage report={displayReport} variant="stf" />}
      </div>
      <MasterUnlockModal
        open={masterUnlockOpen}
        theme="hsi"
        onClose={() => setMasterUnlockOpen(false)}
        onUnlocked={(token) => {
          setMasterUnlockToken(token);
          setMasterUnlockOpen(false);
          setView("master");
        }}
      />
    </main>
  );
}


