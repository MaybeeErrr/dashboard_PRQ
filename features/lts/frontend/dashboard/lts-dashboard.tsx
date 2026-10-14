"use client";

import Image from "next/image";
import { useRouter } from "next/navigation";
import {
  ArrowLeft,
  ArrowUpRight,
  BarChart3,
  Database,
  Download,
  FileSpreadsheet,
  Home,
  Loader2,
  LockKeyhole,
  RefreshCw,
  ShieldCheck,
  Target,
  TrendingDown,
  TrendingUp,
  Upload,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import type { DashboardPermissions } from "@/features/hsi/backend/auth/session";
import { dateRangeLabel, formatDataCutoffDate } from "@/features/shared/dashboard/date-label";
import { DashboardMetricCard, type DashboardMetricTone } from "@/features/shared/dashboard/metric-card";
import { MasterUnlockModal } from "@/features/shared/dashboard/master-unlock-modal";
import { DashboardPeriodFilter, defaultPeriodFilter, type DashboardPeriodFilterState } from "@/features/shared/dashboard/period-filter";
import { DashboardStatusPill, type DashboardStatusTone } from "@/features/shared/dashboard/status-pill";
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
  computeLtsReport,
  ltsDatasetDefinitions,
  ltsReportTabs,
  periodKeyFromLtsRows,
  periodLabelFromLtsPeriodKey,
  validateLtsDataset,
  type LtsCell,
  type LtsComputedReport,
  type LtsDatasetKey,
  type LtsMasterRow,
  type LtsReportRow,
  type LtsReportTabKey,
  type LtsReportTable,
  type LtsUploadedDataset,
} from "@/features/lts/domain/reporting/reporting";
import { ltsSeedData } from "@/features/lts/domain/reporting/seed-data";

type View = "home" | "master" | "report";

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

type DatasetState = Partial<Record<LtsDatasetKey, LtsUploadedDataset>>;
type DatasetArchive = Partial<Record<LtsDatasetKey, Record<string, LtsUploadedDataset>>>;
type DashboardPersistedState = {
  datasets: DatasetState;
  datasetArchive?: DatasetArchive;
  selectedPeriodKey?: string;
  updatedAt: string | null;
};

const STORAGE_KEY = "lts-dashboard-state-v1";
const DATA_VERSION_KEY = "lts-seed-data-version-v1";
const navItems: { id: View; label: string; icon: React.ComponentType<{ className?: string }> }[] = [
  { id: "home", label: "Home", icon: Home },
  { id: "master", label: "MASTER", icon: Database },
  { id: "report", label: "Report LTS", icon: FileSpreadsheet },
];

const compactColumns: Record<LtsDatasetKey, string[]> = {
  sales: ["WITEL", "TELKOM_DAERAH", "Jml_Sales", "COUNT_SID", "PRODUK", "PRODUCT", "TYPE_LAYANAN", "NPER", "PERIOD", "LAST_UPDATED_DATE"],
  caps: ["L_WITEL", "WITEL (NEW_TERRITORY)", "TWITEL", "TELKOM_DAERAH", "L_PRODUK", "PRODUK", "PRODUCT", "COUNT_SID", "SSL_CT0", "SSL_DO", "PERIODE", "PERIODE_CT0_BAYAR"],
  ct0: ["Witel", "Witel Bill", "OLD_WITEL", "Datel", "Sto Desc", "Produk", "Bayar Desc", "STATUS_PAYMENT", "JML LYN", "Number of Records", "JML_BLN_NUNGGAK", "TUNGGAKAN", "Nper", "TAHUN", "LAST_PERIODE"],
};

const pageBackground =
  "min-h-screen bg-[#071426] bg-[radial-gradient(circle_at_90%_8%,rgba(124,74,196,0.18),transparent_32%),radial-gradient(circle_at_8%_24%,rgba(109,58,168,0.16),transparent_28%),linear-gradient(135deg,#061426_0%,#0a1f3b_50%,#07182d_100%)] text-slate-950";
const glassPanel =
  "rounded-xl border border-violet-200/10 bg-[#100d1d]/82 shadow-[0_24px_70px_-48px_rgba(0,0,0,0.95)] ring-1 ring-violet-300/[0.04] backdrop-blur-xl transition-shadow duration-200";
const lightPanel =
  "rounded-xl border border-white/15 bg-white text-slate-950 shadow-[0_18px_44px_-30px_rgba(0,0,0,0.58)] ring-1 ring-slate-950/[0.04] transition-shadow duration-200";
const primaryButton =
  "inline-flex h-10 items-center justify-center gap-2 rounded-lg bg-[linear-gradient(135deg,#43206a,#6d3aa8,#8a58c7)] px-4 text-sm font-semibold text-white shadow-[0_18px_36px_-22px_rgba(124,74,196,0.95)] transition hover:-translate-y-0.5 hover:brightness-110 disabled:cursor-not-allowed disabled:opacity-60";
const ltsMetricTheme = {
  accentLineClass: "bg-gradient-to-r from-violet-500 via-fuchsia-300 to-violet-700",
  detailToneClass: {
    good: "text-emerald-700",
    neutral: "text-slate-500",
    warn: "text-amber-700",
  },
  hoverClass: "hover:-translate-y-0.5",
  iconToneClass: {
    good: "border-emerald-100 bg-white text-emerald-700",
    neutral: "border-violet-100 bg-white text-violet-700",
    warn: "border-amber-100 bg-white text-amber-700",
  },
  labelClass: "text-slate-500",
  panelClass: lightPanel,
  toneClass: {
    good: "from-emerald-50/90 to-white text-emerald-700",
    neutral: "from-violet-50/90 to-white text-violet-700",
    warn: "from-amber-50/90 to-white text-amber-700",
  },
  valueClass: "text-slate-950",
} as const;

function formatNumber(value: number | null | undefined) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return new Intl.NumberFormat("id-ID", { maximumFractionDigits: 0 }).format(value);
}

function formatDecimal(value: number | null | undefined) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return new Intl.NumberFormat("id-ID", { maximumFractionDigits: 2 }).format(value);
}

function formatSavedAt(value: string | null | undefined) {
  if (!value) return "Belum pernah diupload";
  return new Intl.DateTimeFormat("id-ID", {
    dateStyle: "medium",
    timeStyle: "short",
    timeZone: "Asia/Jakarta",
  })
    .format(new Date(value))
    .replace(/\./g, ":");
}

function normalizeHeader(value: string) {
  return value.trim().toUpperCase().replace(/\s+/g, " ");
}

function cleanWorkbookCell(value: unknown): LtsCell {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date) {
    const pad = (input: number) => String(input).padStart(2, "0");
    return `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
  }
  if (typeof value === "number" || typeof value === "boolean") return value;
  return String(value);
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

function findHeaderIndexes(headers: string[], selectedColumns: string[]) {
  const headerMap = new Map(headers.map((header, index) => [normalizeHeader(header), index]));
  return selectedColumns.map((column) => headerMap.get(normalizeHeader(column)) ?? -1);
}

function rowsIncludeHeader(rows: unknown[][], requiredColumns: string[]) {
  const first = rows[0]?.map((item) => normalizeHeader(String(item ?? ""))) ?? [];
  return requiredColumns.every((column) => first.includes(normalizeHeader(column)));
}

async function readLtsDatasetFromWorkbook(file: File, datasetKey: LtsDatasetKey): Promise<LtsUploadedDataset> {
  const definition = ltsDatasetDefinitions[datasetKey];
  const buffer = await file.arrayBuffer();
  const XLSX = await loadSheetJs();
  const workbook = XLSX.read(buffer, { cellDates: true, type: "array" });
  const sheetName = workbook.SheetNames.find((name) => normalizeHeader(name) === normalizeHeader(definition.sheetName)) ?? workbook.SheetNames[0];
  if (!sheetName) throw new Error("Workbook tidak memiliki sheet yang bisa dibaca.");
  const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    blankrows: false,
    defval: null,
    header: 1,
    raw: true,
  });
  if (sheetRows.length === 0) throw new Error(`Sheet ${sheetName} kosong.`);

  const rawHeaders = (sheetRows[0] ?? []).map((value) => String(value ?? "").trim());
  const dataRows = rowsIncludeHeader(sheetRows, definition.requiredColumns) ? sheetRows.slice(1) : sheetRows;
  validateLtsDataset(datasetKey, rawHeaders, dataRows as LtsMasterRow[]);

  const compactHeaders = compactColumns[datasetKey];
  const indexes = findHeaderIndexes(rawHeaders, compactHeaders);
  const compactRows = dataRows
    .map((row) => indexes.map((index) => (index >= 0 ? cleanWorkbookCell(row[index]) : null)))
    .filter((row) => row.some((cell) => cell !== null && String(cell).trim() !== ""));
  if (compactRows.length === 0) throw new Error(`${definition.label} tidak memiliki baris data valid.`);
  validateLtsDataset(datasetKey, compactHeaders, compactRows);
  return {
    datasetKey,
    fileName: file.name,
    headers: compactHeaders,
    periodKey: periodKeyFromLtsRows(datasetKey, compactHeaders, compactRows),
    rows: compactRows,
    rowCount: compactRows.length,
    uploadedAt: new Date().toISOString(),
  };
}

function normalizePersistedDataset(datasetKey: LtsDatasetKey, value: Partial<LtsUploadedDataset> | undefined): LtsUploadedDataset | null {
  if (!value || !Array.isArray(value.headers) || !Array.isArray(value.rows)) return null;
  const headers = value.headers.map(String);
  const rows = value.rows.map((row) => (Array.isArray(row) ? row : []));
  return {
    datasetKey,
    fileName: typeof value.fileName === "string" ? value.fileName : ltsDatasetDefinitions[datasetKey].sheetName,
    headers,
    periodKey: typeof value.periodKey === "string" ? value.periodKey : periodKeyFromLtsRows(datasetKey, headers, rows),
    rows,
    rowCount: typeof value.rowCount === "number" ? value.rowCount : rows.length,
    uploadedAt: typeof value.uploadedAt === "string" ? value.uploadedAt : new Date().toISOString(),
  };
}

function seedDatasets() {
  const output: DatasetState = {};
  for (const item of ltsSeedData.datasets) {
    const datasetKey = item.datasetKey as LtsDatasetKey;
    output[datasetKey] = {
      datasetKey,
      fileName: item.fileName,
      headers: [...item.headers],
      periodKey: item.periodKey,
      rows: item.rows.map((row) => [...row]) as LtsMasterRow[],
      rowCount: item.rowCount,
      uploadedAt: item.uploadedAt,
    };
  }
  return output;
}

function initialSeedArchive() {
  return archiveFromDatasets(seedDatasets());
}

function archiveFromDatasets(datasets: DatasetState) {
  const archive: DatasetArchive = {};
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    const dataset = datasets[key];
    if (!dataset) continue;
    archive[key] = { ...(archive[key] ?? {}), [dataset.periodKey]: dataset };
  }
  return archive;
}

function activeDatasetsForPeriod(archive: DatasetArchive, periodKey: string) {
  const datasets: DatasetState = {};
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    const dataset = archive[key]?.[periodKey];
    if (dataset) datasets[key] = dataset;
  }
  return datasets;
}

function mergeDatasetArchives(first: DatasetArchive, second: DatasetArchive) {
  const output: DatasetArchive = { ...first };
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    output[key] = { ...(first[key] ?? {}), ...(second[key] ?? {}) };
  }
  return output;
}

function latestLtsPeriodKey(archive: DatasetArchive) {
  const keys = new Set<string>();
  for (const datasetArchive of Object.values(archive)) {
    Object.keys(datasetArchive ?? {}).forEach((periodKey) => keys.add(periodKey));
  }
  return [...keys].sort().at(-1) ?? "202609";
}

function periodOptionsFromArchive(archive: DatasetArchive) {
  const keys = new Set<string>();
  for (const datasetArchive of Object.values(archive)) {
    Object.keys(datasetArchive ?? {}).forEach((periodKey) => keys.add(periodKey));
  }
  return [...keys]
    .sort((a, b) => b.localeCompare(a))
    .map((periodKey) => ({
      periodKey,
      label: periodLabelFromLtsPeriodKey(periodKey),
    }));
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

function normalizeLtsPeriodFilter(value: DashboardPeriodFilterState) {
  if (value.mode === "current-month") return { ...value, endMonth: value.startMonth };
  if (value.mode === "ytd") return { ...value, startMonth: `${value.endMonth.slice(0, 4)}01` };
  if (value.mode === "date-range") {
    return {
      ...value,
      endMonth: value.endDate ? value.endDate.slice(0, 7).replace("-", "") : value.endMonth,
      startMonth: value.startDate ? value.startDate.slice(0, 7).replace("-", "") : value.startMonth,
    };
  }
  return value;
}

function validateLtsPeriodFilter(value: DashboardPeriodFilterState) {
  const normalized = normalizeLtsPeriodFilter(value);
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

function combineDatasetsForRange(archive: DatasetArchive, filter: DashboardPeriodFilterState) {
  const normalized = normalizeLtsPeriodFilter(filter);
  const validationMessage = validateLtsPeriodFilter(normalized);
  if (validationMessage) return { datasets: activeDatasetsForPeriod(archive, normalized.endMonth), validationMessage };
  const periodKeys = monthRangeKeys(normalized.startMonth, normalized.endMonth);
  const datasets: DatasetState = {};
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    const entries = periodKeys.map((periodKey) => archive[key]?.[periodKey]).filter(Boolean) as LtsUploadedDataset[];
    if (entries.length === 0) continue;
    const first = entries[0];
    datasets[key] = {
      ...first,
      fileName: entries.length === 1 ? first.fileName : `${first.fileName} + ${entries.length - 1} periode`,
      periodKey: normalized.endMonth,
      rows: entries.flatMap((entry) => entry.rows),
      rowCount: entries.reduce((sum, entry) => sum + entry.rowCount, 0),
      uploadedAt: entries.map((entry) => entry.uploadedAt).sort().at(-1) ?? first.uploadedAt,
    };
  }
  return {
    datasets,
    validationMessage: Object.keys(datasets).length === 0 ? "Data periode yang dipilih belum tersedia." : null,
  };
}

function ltsPeriodFilterLabel(value: DashboardPeriodFilterState) {
  const normalized = normalizeLtsPeriodFilter(value);
  if (normalized.mode === "date-range") {
    return dateRangeLabel(normalized.startDate, normalized.endDate);
  }
  if (normalized.startMonth === normalized.endMonth) return periodLabelFromLtsPeriodKey(normalized.endMonth);
  const startLabel = periodLabelFromLtsPeriodKey(normalized.startMonth);
  const endLabel = periodLabelFromLtsPeriodKey(normalized.endMonth);
  return startLabel.slice(-4) === endLabel.slice(-4)
    ? `${startLabel.replace(` ${startLabel.slice(-4)}`, "")} - ${endLabel}`
    : `${startLabel} - ${endLabel}`;
}

function ltsPeriodCutoffLabel(value: DashboardPeriodFilterState) {
  const normalized = normalizeLtsPeriodFilter(value);
  const cutoffDate = normalized.mode === "date-range"
    ? new Date(`${normalized.endDate}T00:00:00`)
    : new Date(`${normalized.endMonth.slice(0, 4)}-${normalized.endMonth.slice(4, 6)}-${String(new Date(Number(normalized.endMonth.slice(0, 4)), Number(normalized.endMonth.slice(4, 6)), 0).getDate()).padStart(2, "0")}T00:00:00`);
  return Number.isNaN(cutoffDate.getTime()) ? "DATA S.D. -" : formatDataCutoffDate(cutoffDate);
}

function parsePersistedState(value: unknown): DashboardPersistedState | null {
  if (!value || typeof value !== "object") return null;
  const candidate = value as Partial<DashboardPersistedState>;
  const datasets: DatasetState = {};
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    const dataset = normalizePersistedDataset(key, candidate.datasets?.[key]);
    if (dataset) datasets[key] = dataset;
  }
  let datasetArchive = archiveFromDatasets(datasets);
  for (const key of Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]) {
    const archiveEntries = candidate.datasetArchive?.[key];
    if (!archiveEntries || typeof archiveEntries !== "object") continue;
    for (const [periodKey, datasetValue] of Object.entries(archiveEntries)) {
      const dataset = normalizePersistedDataset(key, datasetValue);
      if (!dataset) continue;
      datasetArchive = {
        ...datasetArchive,
        [key]: { ...(datasetArchive[key] ?? {}), [periodKey]: { ...dataset, periodKey } },
      };
    }
  }
  const selectedPeriodKey =
    typeof candidate.selectedPeriodKey === "string" && periodOptionsFromArchive(datasetArchive).some((option) => option.periodKey === candidate.selectedPeriodKey)
      ? candidate.selectedPeriodKey
      : latestLtsPeriodKey(datasetArchive);
  return {
    datasets: Object.keys(datasets).length > 0 ? datasets : activeDatasetsForPeriod(datasetArchive, selectedPeriodKey),
    datasetArchive,
    selectedPeriodKey,
    updatedAt: typeof candidate.updatedAt === "string" ? candidate.updatedAt : null,
  };
}

function persistLocally(state: DashboardPersistedState) {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch {
    // Large workbooks can exceed browser quota; server sync and current memory state still continue.
  }
}

function readLocalState() {
  try {
    return parsePersistedState(JSON.parse(window.localStorage.getItem(STORAGE_KEY) ?? "null"));
  } catch {
    return null;
  }
}

async function loadStateFromServer(): Promise<{ configured: boolean; state: DashboardPersistedState | null }> {
  const response = await fetch("/api/lts/dashboard-state", { cache: "no-store" });
  if (!response.ok) throw new Error("Gagal memuat data LTS dari server.");
  const payload = (await response.json()) as { configured?: boolean; state?: unknown };
  return {
    configured: Boolean(payload.configured),
    state: parsePersistedState(payload.state),
  };
}

async function saveStateToServer(state: DashboardPersistedState, masterUnlockToken: string | null): Promise<{ configured: boolean }> {
  const response = await fetch("/api/lts/dashboard-state", {
    body: JSON.stringify(state),
    headers: {
      "Content-Type": "application/json",
      ...(masterUnlockToken ? { "x-master-unlock-token": masterUnlockToken } : {}),
    },
    method: "PUT",
  });
  if (response.status === 503) return { configured: false };
  if (!response.ok) throw new Error("Gagal menyimpan data LTS ke server.");
  const payload = (await response.json()) as { configured?: boolean };
  return { configured: Boolean(payload.configured) };
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

async function createLtsReportPng(report: LtsComputedReport, timestamp: Date) {
  const exportTableWidth = 1320;
  const columns = [
    { key: "label", label: "WILAYAH", width: 178 },
    { key: "coordinator", label: "KOORDINATOR", width: 190 },
    { key: "sales", label: "SALES", width: 76 },
    { key: "ct0", label: "CT0", width: 76 },
    { key: "ct0Ratio", label: "CT0/SALES", width: 88 },
    { key: "ct0Rank", label: "RANK", width: 62 },
    { key: "caps", label: "CAPS", width: 76 },
    { key: "capsRatio", label: "CAPS/SALES", width: 92 },
    { key: "capsRank", label: "RANK", width: 62 },
    { key: "score", label: "SCORE", width: 76 },
    { key: "scoreRank", label: "RANK", width: 62 },
    { key: "capsRankOnly", label: "RANK LTS CAPS", width: 94 },
    { key: "churn", label: "CHURN", width: 86 },
    { key: "gap", label: "GAP", width: 76 },
  ];
  const width = Math.max(1400, exportTableWidth + 80);
  const rowHeight = 30;
  const groupHeaderHeight = 38;
  const subHeaderHeight = 36;
  const tableTitleHeight = 92;
  const headerHeight = 112;
  const tableHeaderHeight = groupHeaderHeight + subHeaderHeight;
  const height = headerHeight + tableTitleHeight * 2 + tableHeaderHeight * 2 + ([...report.witel.rows, report.witel.total, ...report.hotda.rows, report.hotda.total].length) * rowHeight + 92;
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = width * scale;
  canvas.height = height * scale;
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Canvas tidak tersedia.");
  const ctx = context;
  ctx.scale(scale, scale);
  ctx.fillStyle = "#f8f5fb";
  ctx.fillRect(0, 0, width, height);
  ctx.fillStyle = "#090612";
  ctx.fillRect(0, 0, width, headerHeight);
  ctx.fillStyle = "#ffffff";
  ctx.font = "700 28px Arial";
  ctx.fillText(report.label, 40, 46);
  ctx.font = "600 13px Arial";
  ctx.fillStyle = "#d8d3e3";
  ctx.fillText(`${report.periodLabel} | Source: ${report.sourceLabel} | ${report.cutoffLabel}`, 40, 75);
  ctx.fillStyle = "#a89ec0";
  ctx.fillText(`Rendered ${formatSavedAt(timestamp.toISOString())} WIB`, 40, 94);

  function drawCell(x: number, y: number, w: number, h: number, fill: string, stroke = "#e7e2ed") {
    ctx.fillStyle = fill;
    ctx.fillRect(x, y, w, h);
    ctx.strokeStyle = stroke;
    ctx.strokeRect(x, y, w, h);
  }

  function drawText(text: string, x: number, y: number, w: number, align: CanvasTextAlign, fill: string, baseline: CanvasTextBaseline = "middle") {
    ctx.save();
    ctx.beginPath();
    ctx.rect(x, y - 18, w, 36);
    ctx.clip();
    ctx.textAlign = align;
    ctx.textBaseline = baseline;
    ctx.fillStyle = fill;
    ctx.fillText(text, align === "left" ? x + 7 : align === "right" ? x + w - 7 : x + w / 2, y);
    ctx.restore();
  }

  function drawPill(text: string, x: number, y: number, color: string) {
    if (!text) return;
    const pillWidth = Math.max(34, ctx.measureText(text).width + 18);
    const pillHeight = 22;
    const left = x - pillWidth / 2;
    const top = y - pillHeight / 2;
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.roundRect(left, top, pillWidth, pillHeight, 11);
    ctx.fill();
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText(text, x, y);
  }

  function textValue(row: LtsReportRow, key: string) {
    if (key === "label") return row.label;
    if (key === "coordinator") return row.coordinator ?? "";
    if (key === "ct0Ratio" || key === "capsRatio" || key === "score") return formatDecimal(Number(row[key as keyof LtsReportRow]));
    if (key === "ct0Rank" || key === "capsRank" || key === "scoreRank" || key === "capsRankOnly") return formatNumber(row[key as keyof LtsReportRow] as number | null);
    return formatNumber(row[key as keyof LtsReportRow] as number);
  }

  function exportColumnsForTable(table: LtsReportTable) {
    const visibleColumns = columns.filter((column) => !(table.scope === "witel" && column.key === "coordinator") && !(table.scope === "hotda" && column.key === "capsRankOnly"));
    const baseWidth = visibleColumns.reduce((sum, column) => sum + column.width, 0);
    const scaleFactor = exportTableWidth / Math.max(baseWidth, 1);
    return visibleColumns.map((column) => ({
      ...column,
      label: column.key === "label" ? (table.scope === "witel" ? "WITEL" : "BRANCH") : column.label,
      width: column.width * scaleFactor,
    }));
  }

  function drawTable(table: LtsReportTable, y: number) {
    const exportColumns = exportColumnsForTable(table);
    const columnPositions = new Map<string, { width: number; x: number }>();
    let currentX = 40;
    for (const column of exportColumns) {
      columnPositions.set(column.key, { width: column.width, x: currentX });
      currentX += column.width;
    }
    const groupWidth = (keys: string[]) => keys.reduce((sum, key) => sum + (columnPositions.get(key)?.width ?? 0), 0);
    const groupX = (key: string) => columnPositions.get(key)?.x ?? 40;
    const gradient = ctx.createLinearGradient(40, y, 40 + exportTableWidth, y);
    gradient.addColorStop(0, "#07040f");
    gradient.addColorStop(0.62, "#24103f");
    gradient.addColorStop(1, "#43206a");
    ctx.fillStyle = gradient;
    ctx.fillRect(40, y, exportTableWidth, tableTitleHeight);
    ctx.textAlign = "center";
    ctx.font = "700 10px Arial";
    ctx.fillStyle = "#c9b8ee";
    ctx.fillText("MAIN TABLE", 40 + exportTableWidth / 2, y + 23);
    ctx.font = "700 19px Arial";
    ctx.fillStyle = "#ffffff";
    ctx.fillText(table.title, 40 + exportTableWidth / 2, y + 50);
    ctx.font = "600 12px Arial";
    ctx.fillStyle = "#d8d3e3";
    ctx.fillText(table.subtitle, 40 + exportTableWidth / 2, y + 70);
    y += tableTitleHeight;

    const drawHeader = (key: string, label: string, fill: string, height: number, yOffset = 0, textColor = "#ffffff") => {
      const position = columnPositions.get(key);
      if (!position) return;
      drawCell(position.x, y + yOffset, position.width, height, fill, "#4b356d");
      ctx.font = "700 11px Arial";
      drawText(label, position.x, y + yOffset + height / 2, position.width, "center", textColor);
    };
    const drawHeaderGroup = (keys: string[], label: string, fill: string) => {
      const x = groupX(keys[0]);
      const w = groupWidth(keys);
      drawCell(x, y, w, groupHeaderHeight, fill, "#4b356d");
      ctx.font = "700 12px Arial";
      drawText(label, x, y + groupHeaderHeight / 2, w, "center", "#ffffff");
    };
    drawHeader("label", table.scope === "witel" ? "WITEL" : "BRANCH", "#120824", tableHeaderHeight);
    drawHeader("coordinator", "KOORDINATOR", "#1a0d2f", tableHeaderHeight);
    drawHeader("sales", "Sales", "#43206a", tableHeaderHeight);
    drawHeaderGroup(["ct0", "ct0Ratio", "ct0Rank"], "Bobot 50% CT0", "#3730a3");
    drawHeaderGroup(["caps", "capsRatio", "capsRank"], "Bobot 50% CAPS", "#0e7490");
    drawHeaderGroup(["score", "scoreRank"], "Score", "#5b21b6");
    drawHeader("capsRankOnly", "Rank LTS CAPS", "#92400e", tableHeaderHeight);
    drawHeaderGroup(["churn", "gap"], "CAPS 27/8", "#047857");

    const subHeaderStyles: Record<string, { fill: string; text: string }> = {
      caps: { fill: "#ecfeff", text: "#0e7490" },
      capsRatio: { fill: "#ecfeff", text: "#0e7490" },
      capsRank: { fill: "#ecfeff", text: "#0e7490" },
      churn: { fill: "#fff7ed", text: "#92400e" },
      ct0: { fill: "#eef2ff", text: "#3730a3" },
      ct0Rank: { fill: "#eef2ff", text: "#3730a3" },
      ct0Ratio: { fill: "#eef2ff", text: "#3730a3" },
      gap: { fill: "#ecfdf5", text: "#047857" },
      score: { fill: "#f5f1ff", text: "#5b21b6" },
      scoreRank: { fill: "#f5f1ff", text: "#5b21b6" },
    };
    for (const column of exportColumns) {
      if (["label", "coordinator", "sales", "capsRankOnly"].includes(column.key)) continue;
      const style = subHeaderStyles[column.key] ?? { fill: "#ffffff", text: "#111827" };
      drawCell(columnPositions.get(column.key)?.x ?? 40, y + groupHeaderHeight, column.width, subHeaderHeight, style.fill, "#dbe4f0");
      ctx.font = "700 11px Arial";
      drawText(column.label, columnPositions.get(column.key)?.x ?? 40, y + groupHeaderHeight + subHeaderHeight / 2, column.width, "center", style.text);
    }
    y += tableHeaderHeight;

    [...table.rows, table.total].forEach((row, index, list) => {
      const total = index === list.length - 1;
      ctx.font = total ? "700 11px Arial" : "600 11px Arial";
      for (const column of exportColumns) {
        const position = columnPositions.get(column.key);
        if (!position) continue;
        const isNegative = column.key === "gap" && row.gap < 0;
        const rowBackground = index % 2 === 0 ? "#ffffff" : "#f7f4ff";
        const fill =
          column.key === "label"
            ? total ? "#0b0614" : rowBackground
            : column.key === "coordinator"
              ? total ? "#160d24" : rowBackground
              : column.key === "sales"
                ? total ? "#160d24" : "#f3edff"
                : column.key === "ct0" || column.key === "ct0Rank"
                  ? "#eef2ff"
                  : column.key === "ct0Ratio"
                    ? "#e0e7ff"
                    : column.key === "caps" || column.key === "capsRank"
                      ? "#ecfeff"
                      : column.key === "capsRatio"
                        ? "#cffafe"
                        : column.key === "score"
                          ? "#f5f1ff"
                          : column.key === "scoreRank"
                            ? "#ede9fe"
                            : column.key === "capsRankOnly"
                              ? "#fffbeb"
                              : column.key === "churn"
                                ? "#fff7ed"
                                : isNegative
                                  ? "#fff1f2"
                                  : "#ecfdf5";
        drawCell(position.x, y, column.width, rowHeight, fill);
        const value = textValue(row, column.key);
        if (["ct0Rank", "capsRank", "scoreRank", "capsRankOnly"].includes(column.key) && value && !total) {
          const pillColor = column.key === "ct0Rank" ? "#312e81" : column.key === "capsRank" ? "#164e63" : column.key === "capsRankOnly" ? "#92400e" : "#2e1065";
          drawPill(value, position.x + column.width / 2, y + rowHeight / 2, pillColor);
          continue;
        }
        const textColor =
          total && (column.key === "label" || column.key === "coordinator" || column.key === "sales")
            ? "#ffffff"
            : isNegative
              ? "#be123c"
              : column.key === "sales"
                ? "#321653"
                : column.key === "ct0Ratio"
                  ? "#3730a3"
                  : column.key === "capsRatio"
                    ? "#0e7490"
                    : column.key === "score"
                      ? "#5b21b6"
                      : column.key === "churn"
                        ? "#92400e"
                        : column.key === "gap"
                          ? "#047857"
                          : "#111827";
        const align = column.key === "label" || column.key === "coordinator" ? "left" : "right";
        drawText(value, position.x, y + rowHeight / 2, column.width, align, textColor);
      }
      y += rowHeight;
    });
    ctx.textAlign = "left";
    return y + 22;
  }

  const afterWitel = drawTable(report.witel, headerHeight + 26);
  drawTable(report.hotda, afterWitel);
  const blob = await new Promise<Blob>((resolve, reject) => {
    canvas.toBlob((result) => (result ? resolve(result) : reject(new Error("PNG gagal dibuat."))), "image/png");
  });
  downloadBlob(blob, `${report.label.replace(/\s+/g, "-")}.png`);
}

function StatusPill({ children, tone = "neutral" }: { children: React.ReactNode; tone?: "neutral" | "good" | "warn" | "bad" }) {
  return (
    <span className="inline-flex min-h-8 items-center uppercase">
      <DashboardStatusPill tone={tone as DashboardStatusTone}>{children}</DashboardStatusPill>
    </span>
  );
}

function MetricCard({
  detail,
  icon: Icon,
  label,
  tone = "neutral",
  value,
}: {
  detail: string;
  icon?: React.ComponentType<{ className?: string }>;
  label: string;
  tone?: "neutral" | "good" | "warn" | "bad";
  value: string;
}) {
  return (
    <DashboardMetricCard
      detail={detail}
      icon={Icon}
      label={label}
      theme={ltsMetricTheme}
      tone={(tone === "bad" ? "warn" : tone) as DashboardMetricTone}
      value={value}
    />
  );
}

type LtsWilayahFilter = "all-witel" | "all-hotda" | `witel:${string}` | `hotda:${string}`;

function formatRatio(value: number | null | undefined) {
  return formatDecimal(value);
}

function chartLabelNumber(value: unknown) {
  return typeof value === "number" && value > 0 ? formatNumber(value) : "";
}

function chartLabelDecimal(value: unknown) {
  return typeof value === "number" && Number.isFinite(value) ? value.toFixed(2) : "";
}

function tooltipValue(value: unknown) {
  return typeof value === "number" ? formatDecimal(value) : String(value ?? "");
}

function activeLtsRow(report: LtsComputedReport, filter: LtsWilayahFilter) {
  if (filter === "all-witel") return { label: "ALL WITEL", row: report.witel.total, rows: report.witel.rows, table: report.witel, type: "WITEL" };
  if (filter === "all-hotda") return { label: "ALL HOTDA", row: report.hotda.total, rows: report.hotda.rows, table: report.hotda, type: "HOTDA" };
  const [kind, label] = filter.split(":");
  const table = kind === "hotda" ? report.hotda : report.witel;
  return {
    label,
    row: table.rows.find((item) => item.label === label) ?? table.total,
    rows: table.rows,
    table,
    type: kind === "hotda" ? "HOTDA" : "WITEL",
  };
}

function bestByScore(rows: LtsReportRow[]) {
  return [...rows].sort((a, b) => a.score - b.score)[0] ?? null;
}

function DashboardHome({
  activeTab,
  onGoMaster,
  onTabChange,
  onPeriodFilterChange,
  onPeriodChange,
  periodFilter,
  periodOptions,
  periodValidationMessage,
  readyCount,
  reports,
  selectedPeriodKey,
}: {
  activeTab: LtsReportTabKey;
  onGoMaster: () => void;
  onTabChange: (tab: LtsReportTabKey) => void;
  onPeriodFilterChange: (value: DashboardPeriodFilterState) => void;
  onPeriodChange: (periodKey: string) => void;
  periodFilter: DashboardPeriodFilterState;
  periodOptions: { label: string; periodKey: string }[];
  periodValidationMessage: string | null;
  readyCount: number;
  reports: Partial<Record<LtsReportTabKey, LtsComputedReport>>;
  selectedPeriodKey: string;
}) {
  const availableTabs = ltsReportTabs.filter((tab) => reports[tab.key]);
  const fallbackTab = availableTabs[0]?.key;
  const selectedTab = reports[activeTab] ? activeTab : fallbackTab;
  const report = selectedTab ? reports[selectedTab] : null;
  const [wilayahFilter, setWilayahFilter] = useState<LtsWilayahFilter>("all-witel");
  const [rankingScope, setRankingScope] = useState<"witel" | "hotda">("witel");
  const [chartMetric, setChartMetric] = useState<"score" | "ct0" | "caps">("score");
  const [trendMetric, setTrendMetric] = useState<"score" | "ct0" | "caps">("score");

  if (!report) {
    return (
      <div className="space-y-5">
        <section className={`${glassPanel} p-8 text-center`}>
          <ShieldCheck className="mx-auto h-10 w-10 text-violet-200" />
          <h2 className="mt-4 text-xl font-bold text-white">LTS data belum tersedia</h2>
          <p className="mx-auto mt-2 max-w-xl text-sm leading-6 text-violet-100/65">
            Upload SALES_MASTER, CAPS_MASTER, dan CT0_MASTER di Master Data LTS untuk mulai menghitung laporan.
          </p>
          <button className={`${primaryButton} mt-5`} onClick={onGoMaster}>
            Go to Master Data
          </button>
        </section>
      </div>
    );
  }

  const active = activeLtsRow(report, wilayahFilter);
  const rankingRows = rankingScope === "witel" ? report.witel.rows : report.hotda.rows;
  const ranking = [...rankingRows].sort((a, b) => a.score - b.score);
  const topWilayah = ranking[0];
  const chartRows = wilayahFilter === "all-witel" || wilayahFilter === "all-hotda" ? active.rows : [active.row];
  const metricLabel = chartMetric === "score" ? "SCORE" : chartMetric === "ct0" ? "CT0/SALES" : "CAPS/SALES";
  const chartData = chartRows.map((row) => ({
    caps: row.caps,
    capsRatio: Number(row.capsRatio.toFixed(2)),
    ct0: row.ct0,
    ct0Ratio: Number(row.ct0Ratio.toFixed(2)),
    name: row.label.replace("YOGYA JATENG SELATAN", "YJS").replace("SEMARANG JATENG UTARA", "SJU").replace("SOLO JATENG TIMUR", "SJT"),
    sales: row.sales,
    score: Number(row.score.toFixed(2)),
  }));
  const trendValue = (row: LtsReportRow) => (trendMetric === "score" ? row.score : trendMetric === "ct0" ? row.ct0Ratio : row.capsRatio);
  const trendLabel = trendMetric === "score" ? "SCORE LTS" : trendMetric === "ct0" ? "CT0 / SALES" : "CAPS / SALES";
  const trendAverage = active.rows.reduce((sum, row) => sum + trendValue(row), 0) / Math.max(active.rows.length, 1);
  const trendData = active.rows.map((row) => ({
    average: Number(trendAverage.toFixed(2)),
    name: row.label.replace("YOGYA JATENG SELATAN", "YJS").replace("SEMARANG JATENG UTARA", "SJU").replace("SOLO JATENG TIMUR", "SJT"),
    value: Number(trendValue(row).toFixed(2)),
  }));
  const heroTitle = wilayahFilter === "all-witel" || wilayahFilter === "all-hotda" ? `LTS PERFORMA ${active.label}` : `LTS PERFORMA ${active.type} ${active.label}`;
  const validationReady = readyCount === 3 ? "VALIDASI MATCH" : `${readyCount} / 3 READY`;
  const validationTone = readyCount === 3 ? "good" : readyCount > 0 ? "warn" : "bad";
  const bestWitel = bestByScore(report.witel.rows);
  const bestHotda = bestByScore(report.hotda.rows);

  return (
    <div className="space-y-5">
      <section className={`${lightPanel} relative overflow-hidden`}>
        <div className="absolute inset-0 bg-[radial-gradient(circle_at_88%_10%,rgba(154,106,214,0.18),transparent_28%),radial-gradient(circle_at_8%_18%,rgba(109,58,168,0.08),transparent_24%),linear-gradient(135deg,#ffffff_0%,#fbfaff_52%,#f4edff_100%)]" />
        <div className="absolute inset-x-0 top-0 h-1 bg-gradient-to-r from-[#24103f] via-[#7c4ac4] to-[#b28ae5]" />
        <div className="relative flex flex-wrap items-center justify-between gap-3 px-4 py-3">
          <div className="min-w-[280px] flex-1">
            <p className="inline-flex items-center gap-2 text-[11px] font-semibold uppercase text-violet-700">
              <BarChart3 className="h-4 w-4" />
              DASHBOARD REPORTING LTS 2026
            </p>
            <h1 className="mt-1 text-xl font-semibold uppercase tracking-normal text-[#24103f] sm:text-2xl">
              LTS PERFORMANCE
            </h1>
            <p className="mt-1 max-w-2xl text-xs font-medium uppercase leading-5 text-slate-500">
              {report.label} | {active.label} | SOURCE {report.sourceLabel}
            </p>
          </div>
          <div className="flex flex-wrap items-end gap-2">
            <DashboardPeriodFilter
              onChange={(value) => {
                const normalized = normalizeLtsPeriodFilter(value);
                onPeriodFilterChange(normalized);
                const anchorPeriod = normalized.endMonth;
                if (anchorPeriod !== selectedPeriodKey && periodOptions.some((option) => option.periodKey === anchorPeriod)) {
                  onPeriodChange(anchorPeriod);
                }
              }}
              options={periodOptions}
              theme="lts"
              validationMessage={periodValidationMessage}
              value={periodFilter}
            />
            <label className="grid gap-1 text-[11px] font-semibold uppercase text-slate-500">
              REPORT
              <select
                className="h-9 w-[170px] rounded-md border border-violet-100 bg-white px-3 text-xs font-semibold text-[#24103f] shadow-sm outline-none transition hover:border-violet-300 focus:border-violet-400 focus:ring-4 focus:ring-violet-100"
                value={selectedTab}
                onChange={(event) => onTabChange(event.target.value as LtsReportTabKey)}
              >
                {ltsReportTabs.map((tab) => (
                  <option key={tab.key} value={tab.key} disabled={!reports[tab.key]}>
                    {tab.label}
                  </option>
                ))}
              </select>
            </label>
            <label className="grid gap-1 text-[11px] font-semibold uppercase text-slate-500">
              WILAYAH
              <select
                className="h-9 w-[210px] rounded-md border border-violet-100 bg-white px-3 text-xs font-semibold text-[#24103f] shadow-sm outline-none transition hover:border-violet-300 focus:border-violet-400 focus:ring-4 focus:ring-violet-100"
                value={wilayahFilter}
                onChange={(event) => setWilayahFilter(event.target.value as LtsWilayahFilter)}
              >
                <option value="all-witel">ALL WITEL</option>
                <option value="all-hotda">ALL HOTDA</option>
                <optgroup label="WITEL">
                  {report.witel.rows.map((row) => (
                    <option key={row.label} value={`witel:${row.label}`}>
                      {row.label}
                    </option>
                  ))}
                </optgroup>
                <optgroup label="HOTDA">
                  {report.hotda.rows.map((row) => (
                    <option key={row.label} value={`hotda:${row.label}`}>
                      {row.label}
                    </option>
                  ))}
                </optgroup>
              </select>
            </label>
            <StatusPill tone={validationTone}>{validationReady}</StatusPill>
          </div>
        </div>
      </section>

      <section className={`${lightPanel} overflow-hidden`}>
        <div className="relative grid items-start gap-4 bg-[linear-gradient(135deg,#ffffff_0%,#fbfaff_48%,#f6f1ff_100%)] p-4 xl:grid-cols-[1.35fr_0.65fr]">
          <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(109,58,168,0.04)_1px,transparent_1px),linear-gradient(90deg,rgba(109,58,168,0.03)_1px,transparent_1px)] bg-[size:44px_44px] opacity-50" />
          <div className="relative overflow-hidden rounded-xl border border-violet-100 bg-white shadow-[0_16px_36px_-28px_rgba(36,16,63,0.65)]">
            <div className="flex flex-wrap items-center justify-between gap-3 border-b border-violet-100 bg-[#24103f] px-3 py-2.5 text-white">
              <div>
                <p className="text-[11px] font-semibold uppercase text-violet-200">LTS PERFORMANCE SUMMARY</p>
                <h2 className="mt-0.5 text-sm font-semibold tracking-normal">PERIODE {report.periodLabel} | {active.label}</h2>
              </div>
              <StatusPill>LOWER SCORE BETTER</StatusPill>
            </div>
            <div className="grid gap-2.5 p-2.5 xl:grid-cols-2">
              <div className="overflow-hidden rounded-md border border-violet-100 bg-white shadow-sm">
                <div className="flex items-center justify-between gap-3 border-b border-violet-50 px-2.5 py-2">
                  <div>
                    <p className="text-[10px] font-semibold uppercase text-violet-700">LTS PERFORMA SCORE ({report.label})</p>
                    <p className="mt-0.5 text-xs font-semibold uppercase text-slate-950">{report.periodLabel} | {active.label}</p>
                  </div>
                  <span className="rounded-md border border-violet-100 bg-violet-50 px-2.5 py-1 text-[11px] font-semibold uppercase text-violet-700">
                    SCORE
                  </span>
                </div>
                <div className="grid gap-2 p-2.5">
                  <div className="grid gap-2 sm:grid-cols-2">
                    <MetricCard label="SALES" value={formatNumber(active.row.sales)} detail="BASE LTS" icon={Target} />
                    <MetricCard label="CT0" value={formatNumber(active.row.ct0)} detail="BOBOT 50%" tone="warn" icon={TrendingDown} />
                    <MetricCard label="CAPS" value={formatNumber(active.row.caps)} detail="BOBOT 50%" tone="warn" icon={FileSpreadsheet} />
                    <MetricCard label="SCORE" value={formatRatio(active.row.score)} detail="CT0 + CAPS / SALES" tone="good" icon={ArrowUpRight} />
                  </div>
                  <div className="rounded-md border border-violet-100 bg-violet-50/70 p-2">
                    <div className="flex items-center justify-between text-[11px] font-semibold uppercase">
                      <span className="text-slate-500">Score Indicator</span>
                      <span className="text-violet-700">{formatRatio(active.row.score)}</span>
                    </div>
                    <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-violet-100">
                      <div className="h-full rounded-full bg-gradient-to-r from-emerald-400 via-amber-400 to-rose-500" style={{ width: `${Math.min(active.row.score * 18, 100)}%` }} />
                    </div>
                  </div>
                </div>
              </div>

              <div className="overflow-hidden rounded-md border border-violet-100 bg-white shadow-sm">
                <div className="flex items-center justify-between gap-3 border-b border-violet-50 px-2.5 py-2">
                  <div>
                    <p className="text-[10px] font-semibold uppercase text-violet-700">LTS RATIO DETAIL</p>
                    <p className="mt-0.5 text-xs font-semibold uppercase text-slate-950">CT0/SALES + CAPS/SALES | {active.label}</p>
                  </div>
                  <span className="rounded-md border border-violet-100 bg-violet-50 px-2.5 py-1 text-[11px] font-semibold uppercase text-violet-700">
                    RATIO
                  </span>
                </div>
                <div className="grid gap-2 p-2.5">
                  <div className="grid gap-2 sm:grid-cols-2">
                    <MetricCard label="CT0 / SALES" value={formatRatio(active.row.ct0Ratio)} detail={`RANK ${active.row.ct0Rank ?? "-"}`} icon={BarChart3} />
                    <MetricCard label="CAPS / SALES" value={formatRatio(active.row.capsRatio)} detail={`RANK ${active.row.capsRank ?? "-"}`} icon={BarChart3} />
                    <MetricCard label="RANK SCORE" value={active.row.scoreRank ? `#${active.row.scoreRank}` : "-"} detail={active.type} tone="good" icon={ShieldCheck} />
                    <MetricCard label="GAP CAPS" value={formatNumber(active.row.gap)} detail="CAPS 27/8 - CAPS" tone={active.row.gap < 0 ? "bad" : "good"} icon={TrendingUp} />
                  </div>
                  <div className="rounded-md border border-violet-100 bg-violet-50/70 p-2">
                    <div className="flex items-center justify-between text-[11px] font-semibold uppercase">
                      <span className="text-slate-500">Data Completeness</span>
                      <span className="text-violet-700">{readyCount} / 3</span>
                    </div>
                    <div className="mt-2 h-1.5 overflow-hidden rounded-full bg-violet-100">
                      <div className="h-full rounded-full bg-gradient-to-r from-violet-700 to-violet-300" style={{ width: `${(readyCount / 3) * 100}%` }} />
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div className="relative self-start overflow-hidden rounded-xl border border-violet-300/20 bg-[linear-gradient(145deg,#07040f_0%,#24103f_52%,#43206a_100%)] p-5 text-white shadow-[0_30px_78px_-42px_rgba(154,106,214,0.7)]">
            <div className="absolute inset-0 bg-[radial-gradient(circle_at_86%_4%,rgba(178,138,229,0.22),transparent_30%)]" />
            <div className="absolute inset-x-0 top-0 h-px bg-gradient-to-r from-transparent via-violet-200 to-transparent" />
            <div className="relative">
              <div className="flex items-center justify-between gap-4">
                <p className="max-w-[260px] text-sm font-semibold uppercase leading-5 text-violet-50">{heroTitle}</p>
                <span className="rounded-md border border-violet-200/30 bg-violet-200/10 px-2 py-1 text-[11px] font-semibold text-violet-100">
                  SOURCE: {report.sourceLabel}
                </span>
              </div>
              <p className="mt-3 text-[11px] font-semibold uppercase text-violet-200/60">CAKUPAN: {active.label}</p>
              <p className="mt-3 text-xs font-semibold uppercase text-violet-200/72">TOTAL SCORE</p>
              <p className="mt-1.5 text-4xl font-semibold tracking-normal text-white">{formatRatio(active.row.score)}</p>
              <div className="mt-4 h-2 overflow-hidden rounded-full bg-white/10">
                <div className="h-full rounded-full bg-gradient-to-r from-emerald-400 via-amber-400 to-rose-500" style={{ width: `${Math.min(active.row.score * 18, 100)}%` }} />
              </div>
              <div className="mt-4 grid grid-cols-2 gap-2.5 text-sm">
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-violet-200/60">BEST WITEL</p>
                  <p className="mt-1 text-base font-semibold text-white">{bestWitel?.label ?? "-"}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-violet-200/60">BEST HOTDA</p>
                  <p className="mt-1 text-base font-semibold text-white">{bestHotda?.label ?? "-"}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-violet-200/60">RANK SCORE</p>
                  <p className="mt-1 text-base font-semibold text-violet-100">{active.row.scoreRank ? `#${active.row.scoreRank}` : "-"}</p>
                </div>
                <div className="rounded-md border border-white/10 bg-white/[0.04] p-3">
                  <p className="text-xs font-semibold uppercase text-violet-200/60">DATA READY</p>
                  <p className="mt-1 text-base font-semibold text-violet-100">{readyCount} / 3</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <div className="grid gap-2.5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6">
        <MetricCard label="TOTAL SALES" value={formatNumber(active.row.sales)} detail={report.label} icon={Target} />
        <MetricCard label="TOTAL CT0" value={formatNumber(active.row.ct0)} detail="BOBOT 50% CT0" tone="warn" icon={TrendingDown} />
        <MetricCard label="TOTAL CAPS" value={formatNumber(active.row.caps)} detail="BOBOT 50% CAPS" tone="warn" icon={FileSpreadsheet} />
        <MetricCard label="SCORE LTS" value={formatRatio(active.row.score)} detail="LOWER IS BETTER" tone="good" icon={ArrowUpRight} />
        <MetricCard label="GAP CAPS" value={formatNumber(active.row.gap)} detail="CAPS 27/8 - CAPS" tone={active.row.gap < 0 ? "bad" : "good"} icon={BarChart3} />
        <MetricCard label="RANKING WILAYAH" value={topWilayah?.label ?? "N/A"} detail={topWilayah ? `TOP ${rankingScope.toUpperCase()} SCORE ${formatRatio(topWilayah.score)}` : "RANK SCORE"} icon={ShieldCheck} />
      </div>

      <section className={`${lightPanel} overflow-hidden bg-white`}>
        <div className="flex flex-wrap items-center justify-between gap-4 border-b border-[#eaeff5] px-6 py-5">
          <div>
            <p className="text-[11px] font-semibold uppercase tracking-[0.08em] text-violet-600">Performance Analytics</p>
            <h2 className="mt-1 text-[28px] font-bold uppercase leading-tight tracking-[-0.02em] text-slate-950">Regional Performance</h2>
            <p className="mt-1 text-sm font-medium text-slate-500">Ringkasan performa dan ranking achievement per wilayah</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <select
              className="h-11 rounded-xl border border-violet-100 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-violet-400"
              value={chartMetric}
              onChange={(event) => setChartMetric(event.target.value as "score" | "ct0" | "caps")}
            >
              <option value="score">SCORE</option>
              <option value="ct0">CT0 / SALES</option>
              <option value="caps">CAPS / SALES</option>
            </select>
            <select
              className="h-11 rounded-xl border border-violet-100 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-violet-400"
              value={rankingScope}
              onChange={(event) => setRankingScope(event.target.value as "witel" | "hotda")}
            >
              <option value="witel">WITEL</option>
              <option value="hotda">HOTDA</option>
            </select>
            <StatusPill>{report.periodLabel}</StatusPill>
          </div>
        </div>
        <div className="grid xl:grid-cols-[1.45fr_0.95fr]">
        <div className="p-4">
          <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-xl font-bold uppercase text-slate-950">SALES VS CT0 / CAPS PER WILAYAH</h2>
              <p className="mt-1 text-[13px] font-medium uppercase text-slate-500">INDIKATOR LTS: SALES, CT0, CAPS, DAN SCORE PER {active.type}</p>
            </div>
          </div>
          <div className="mt-3 h-[360px]">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={chartData} margin={{ left: 0, right: 22, top: 24, bottom: 72 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e9d5ff" />
                <XAxis dataKey="name" interval={0} tick={{ angle: -32, dy: 18, fontSize: 9, fontWeight: 700, textAnchor: "end" }} height={62} />
                <YAxis yAxisId="left" tick={{ fontSize: 12 }} />
                <YAxis yAxisId="right" orientation="right" tick={{ fontSize: 12 }} />
                <Tooltip formatter={(value) => tooltipValue(value)} />
                <Legend verticalAlign="bottom" align="center" wrapperStyle={{ bottom: 0, fontSize: 12, fontWeight: 700, paddingTop: 18 }} />
                <Bar yAxisId="left" dataKey="sales" name="SALES" fill="#7c4ac4" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="sales" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Bar yAxisId="left" dataKey="ct0" name="CT0" fill="#b28ae5" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="ct0" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Bar yAxisId="left" dataKey="caps" name="CAPS" fill="#ef4444" radius={[4, 4, 0, 0]}>
                  <LabelList dataKey="caps" position="top" formatter={chartLabelNumber} fontSize={10} />
                </Bar>
                <Line yAxisId="right" type="monotone" dataKey={chartMetric === "score" ? "score" : chartMetric === "ct0" ? "ct0Ratio" : "capsRatio"} name={metricLabel} stroke="#059669" strokeWidth={3} dot={{ r: 3 }}>
                  <LabelList dataKey={chartMetric === "score" ? "score" : chartMetric === "ct0" ? "ct0Ratio" : "capsRatio"} position="top" formatter={chartLabelDecimal} fontSize={10} fill="#047857" />
                </Line>
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="border-t border-violet-100 p-4 xl:border-l xl:border-t-0">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h2 className="text-xl font-bold uppercase text-slate-950">RANKING SCORE LTS</h2>
              <p className="mt-1 text-[13px] font-medium uppercase text-slate-500">RANKING BERDASARKAN SCORE {rankingScope.toUpperCase()}</p>
            </div>
            <ShieldCheck className="h-5 w-5 text-emerald-600" />
          </div>
          <div className="mt-3 divide-y divide-violet-100">
            {ranking.map((row, index) => (
              <div className={`grid grid-cols-[34px_1fr_auto] items-center gap-3 py-2.5 text-sm ${row.label === active.row.label ? "rounded-md bg-violet-50 px-2" : ""}`} key={row.label}>
                <span className="flex h-7 w-7 items-center justify-center rounded-full bg-[#24103f] text-xs font-semibold text-white shadow-sm">{index + 1}</span>
                <div className="min-w-0">
                  <p className="truncate font-semibold text-slate-700">{row.label}</p>
                  <div className="mt-1 h-1.5 overflow-hidden rounded-full bg-violet-100">
                    <div className="h-full rounded-full bg-gradient-to-r from-emerald-400 via-amber-400 to-rose-500" style={{ width: `${Math.min(row.score * 18, 100)}%` }} />
                  </div>
                </div>
                <p className="text-xs font-semibold text-violet-700">{formatRatio(row.score)}</p>
              </div>
            ))}
          </div>
        </div>
        </div>
      </section>

      <section className={`${lightPanel} p-4`}>
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-lg font-semibold uppercase text-slate-950">TREN SCORE LTS PER WILAYAH</h2>
            <p className="mt-1 text-xs font-medium uppercase text-slate-500">{trendLabel} PER {active.type} VS RATA-RATA {active.type}</p>
          </div>
          <select
            className="h-9 rounded-md border border-violet-100 bg-white px-3 text-xs font-semibold uppercase text-slate-700 shadow-sm outline-none focus:border-violet-400 focus:ring-4 focus:ring-violet-100"
            value={trendMetric}
            onChange={(event) => setTrendMetric(event.target.value as "score" | "ct0" | "caps")}
          >
            <option value="score">SCORE</option>
            <option value="ct0">CT0 / SALES</option>
            <option value="caps">CAPS / SALES</option>
          </select>
        </div>
        <div className="mt-3 h-[340px]">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={trendData} margin={{ left: 0, right: 22, top: 24, bottom: 58 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e9d5ff" />
              <XAxis dataKey="name" interval={0} tick={{ angle: -32, dy: 18, fontSize: 9, fontWeight: 700, textAnchor: "end" }} height={62} />
              <YAxis tick={{ fontSize: 12 }} />
              <Tooltip formatter={(value) => tooltipValue(value)} />
              <Legend verticalAlign="bottom" align="center" wrapperStyle={{ bottom: 0, fontSize: 12, fontWeight: 700, paddingTop: 18 }} />
              <Line type="monotone" dataKey="value" name={trendLabel} stroke="#0f766e" strokeWidth={3} dot={{ r: 3 }} />
              <Line type="monotone" dataKey="average" name={`RATA-RATA ${active.type}`} stroke="#64748b" strokeWidth={2.4} strokeDasharray="6 5" dot={{ r: 2 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      </section>
    </div>
  );
}

function MasterUploadCard({
  dataset,
  datasetKey,
  error,
  onUpload,
  processing,
}: {
  dataset?: LtsUploadedDataset;
  datasetKey: LtsDatasetKey;
  error?: string;
  onUpload: (datasetKey: LtsDatasetKey, file: File | undefined) => void;
  processing: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const definition = ltsDatasetDefinitions[datasetKey];
  const ready = Boolean(dataset);
  return (
    <article className={`${glassPanel} flex min-h-[300px] flex-col justify-between p-5 transition hover:-translate-y-1 hover:border-violet-300/25 hover:shadow-[0_26px_80px_-48px_rgba(124,74,196,0.95)]`}>
      <div>
        <div className="flex items-start justify-between gap-3">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.12em] text-violet-200/70">{definition.shortLabel}</p>
            <h2 className="mt-2 text-xl font-bold tracking-normal text-white">{definition.label}</h2>
          </div>
          <StatusPill tone={processing ? "warn" : ready ? "good" : "bad"}>{processing ? "PROCESSING" : ready ? "READY" : "NOT UPLOADED"}</StatusPill>
        </div>
        <p className="mt-4 text-sm leading-6 text-violet-100/68">{definition.purpose}</p>
        <div className="mt-4 space-y-2 rounded-xl border border-violet-200/10 bg-black/15 p-3">
          {definition.outputs.map((output) => (
            <p key={output} className="text-xs font-semibold uppercase text-violet-100/70">
              {output}
            </p>
          ))}
        </div>
      </div>
      <div className="mt-5 space-y-4">
        <div className="grid gap-2 text-xs font-semibold uppercase text-violet-100/65">
          <span>File: {dataset?.fileName ?? "-"}</span>
          <span>Rows: {dataset ? formatNumber(dataset.rowCount) : "-"}</span>
          <span>Last updated: {formatSavedAt(dataset?.uploadedAt)}</span>
        </div>
        {error && <p className="rounded-lg border border-rose-300/20 bg-rose-400/10 px-3 py-2 text-sm font-medium text-rose-100">{error}</p>}
        <button className={primaryButton} disabled={processing} onClick={() => inputRef.current?.click()}>
          {processing ? <Loader2 className="h-4 w-4 animate-spin" /> : <Upload className="h-4 w-4" />}
          {ready ? "Replace Excel" : "Upload Excel"}
        </button>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv"
          className="hidden"
          onChange={(event) => onUpload(datasetKey, event.target.files?.[0])}
        />
      </div>
    </article>
  );
}

function ReportDataTable({ table }: { table: LtsReportTable }) {
  const allRows = [...table.rows, table.total];
  const showCoordinator = table.scope === "hotda";
  const subHeaders = [
    { label: "CT0", className: "bg-[#eef2ff] text-indigo-800" },
    { label: "CT0/Sales", className: "bg-[#eef2ff] text-indigo-800" },
    { label: "Rank", className: "bg-[#eef2ff] text-indigo-800" },
    { label: "CAPS", className: "bg-[#ecfeff] text-cyan-800" },
    { label: "CAPS/Sales", className: "bg-[#ecfeff] text-cyan-800" },
    { label: "Rank", className: "bg-[#ecfeff] text-cyan-800" },
    { label: "Score", className: "bg-[#f5f1ff] text-violet-800" },
    { label: "Rank", className: "bg-[#f5f1ff] text-violet-800" },
    { label: "Churn", className: "bg-[#fff7ed] text-amber-800" },
    { label: "Gap", className: "bg-[#ecfdf5] text-emerald-800" },
  ];
  return (
    <section className={`${glassPanel} overflow-hidden bg-white`}>
      <div className="border-b border-violet-200/10 bg-[linear-gradient(135deg,#07040f,#24103f_60%,#43206a)] px-5 py-5 text-center">
        <p className="text-[11px] font-semibold uppercase tracking-[0.16em] text-violet-200">Main Table</p>
        <h2 className="mt-2 text-xl font-bold tracking-normal text-white">{table.title}</h2>
        <p className="mt-1 text-sm font-semibold uppercase text-violet-100/70">{table.subtitle}</p>
      </div>
      <div className="max-h-[72vh] overflow-auto bg-white">
        <table className="w-full min-w-[1320px] border-separate border-spacing-0 text-[11px] text-slate-800">
          <thead>
            <tr>
              <th rowSpan={2} className="sticky left-0 top-0 z-30 w-[180px] min-w-[180px] border-b border-r border-violet-700/60 bg-[#120824] px-2 py-2 text-center font-semibold text-white">
                {table.scope === "witel" ? "WITEL" : "BRANCH"}
              </th>
              {showCoordinator && (
                <th rowSpan={2} className="sticky top-0 z-20 w-[190px] min-w-[190px] border-b border-r border-violet-700/60 bg-[#1a0d2f] px-2 py-2 text-center font-semibold text-white">
                  KOORDINATOR
                </th>
              )}
              <th rowSpan={2} className="sticky top-0 z-20 w-[84px] border-b border-r border-violet-700/60 bg-[#43206a] px-2 py-2 text-center font-semibold text-white">Sales</th>
              <th colSpan={3} className="sticky top-0 z-20 border-b border-r border-indigo-900/35 bg-[#3730a3] px-2 py-2 text-center font-semibold text-white">Bobot 50% CT0</th>
              <th colSpan={3} className="sticky top-0 z-20 border-b border-r border-cyan-900/35 bg-[#0e7490] px-2 py-2 text-center font-semibold text-white">Bobot 50% CAPS</th>
              <th colSpan={2} className="sticky top-0 z-20 border-b border-r border-violet-900/40 bg-[#5b21b6] px-2 py-2 text-center font-semibold text-white">Score</th>
              {!showCoordinator && <th rowSpan={2} className="sticky top-0 z-20 w-[92px] border-b border-r border-amber-900/35 bg-[#92400e] px-2 py-2 text-center font-semibold text-white">Rank LTS CAPS</th>}
              <th colSpan={2} className="sticky top-0 z-20 border-b border-r border-emerald-900/35 bg-[#047857] px-2 py-2 text-center font-semibold text-white">CAPS 27/8</th>
            </tr>
            <tr>
              {subHeaders.map((header, headerIndex) => (
                <th key={`${header.label}-${headerIndex}`} className={`sticky top-[33px] z-20 w-[82px] border-b border-r border-slate-200 px-2 py-2 text-center font-semibold ${header.className}`}>
                  {header.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {allRows.map((row, index) => {
              const total = index === allRows.length - 1;
              const rowBackground = index % 2 === 0 ? "bg-white" : "bg-[#f7f4ff]";
              return (
                <tr key={`${row.label}-${total ? "total" : "row"}`} className={total ? "bg-[#0b0614] font-bold text-white" : `${rowBackground} text-slate-800 hover:bg-[#efe8ff]`}>
                  <th className={`sticky left-0 z-10 border-b border-r border-violet-100/80 px-3 py-2 text-left ${total ? "bg-[#0b0614] text-white" : `${rowBackground} text-slate-800`}`}>{row.label}</th>
                  {showCoordinator && <td className={`border-b border-r border-violet-100/80 px-3 py-2 ${total ? "bg-[#160d24] text-white" : `${rowBackground} text-slate-700`}`}>{row.coordinator ?? ""}</td>}
                  <td className={`border-b border-r border-violet-100/80 px-2 py-2 text-right ${total ? "bg-[#160d24] text-white" : "bg-[#f3edff] font-semibold text-[#321653]"}`}>{formatNumber(row.sales)}</td>
                  <td className="border-b border-r border-indigo-100 bg-indigo-50/45 px-2 py-2 text-right">{formatNumber(row.ct0)}</td>
                  <td className="border-b border-r border-indigo-100 bg-indigo-50/80 px-2 py-2 text-right font-semibold text-indigo-800">{formatDecimal(row.ct0Ratio)}</td>
                  <td className="border-b border-r border-indigo-100 bg-indigo-50/45 px-2 py-2 text-center">{row.ct0Rank ? <span className="inline-flex min-w-8 justify-center rounded-full bg-indigo-900 px-2 py-0.5 text-white shadow-sm">{row.ct0Rank}</span> : ""}</td>
                  <td className="border-b border-r border-cyan-100 bg-cyan-50/45 px-2 py-2 text-right">{formatNumber(row.caps)}</td>
                  <td className="border-b border-r border-cyan-100 bg-cyan-50/80 px-2 py-2 text-right font-semibold text-cyan-800">{formatDecimal(row.capsRatio)}</td>
                  <td className="border-b border-r border-cyan-100 bg-cyan-50/45 px-2 py-2 text-center">{row.capsRank ? <span className="inline-flex min-w-8 justify-center rounded-full bg-cyan-900 px-2 py-0.5 text-white shadow-sm">{row.capsRank}</span> : ""}</td>
                  <td className="border-b border-r border-violet-100 bg-violet-50/90 px-2 py-2 text-right font-bold text-violet-900">{formatDecimal(row.score)}</td>
                  <td className="border-b border-r border-violet-100 bg-violet-50/50 px-2 py-2 text-center">{row.scoreRank ? <span className="inline-flex min-w-8 justify-center rounded-full bg-violet-950 px-2 py-0.5 text-white shadow-sm">{row.scoreRank}</span> : ""}</td>
                  {!showCoordinator && <td className="border-b border-r border-amber-100 bg-amber-50/70 px-2 py-2 text-center">{row.capsRankOnly ? <span className="inline-flex min-w-8 justify-center rounded-full bg-amber-800 px-2 py-0.5 text-white shadow-sm">{row.capsRankOnly}</span> : ""}</td>}
                  <td className="border-b border-r border-amber-100 bg-amber-50/60 px-2 py-2 text-right text-amber-900">{formatNumber(row.churn)}</td>
                  <td className={`border-b border-r border-violet-100/80 px-2 py-2 text-right font-semibold ${row.gap < 0 ? "bg-rose-50 text-rose-700" : "bg-teal-50 text-teal-700"}`}>{formatNumber(row.gap)}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}

const fullAccessPermissions: DashboardPermissions = {
  canAccessMaster: true,
  canEditMaster: true,
  canViewDashboard: true,
};

export function LtsDashboard({ permissions = fullAccessPermissions }: { permissions?: DashboardPermissions }) {
  const router = useRouter();
  const [view, setView] = useState<View>("home");
  const [datasets, setDatasets] = useState<DatasetState>(() => {
    const archive = initialSeedArchive();
    return activeDatasetsForPeriod(archive, latestLtsPeriodKey(archive));
  });
  const [datasetArchive, setDatasetArchive] = useState<DatasetArchive>(() => initialSeedArchive());
  const [selectedPeriodKey, setSelectedPeriodKey] = useState(() => latestLtsPeriodKey(initialSeedArchive()));
  const [periodFilter, setPeriodFilter] = useState<DashboardPeriodFilterState>(() => defaultPeriodFilter(latestLtsPeriodKey(initialSeedArchive())));
  const [updatedAt, setUpdatedAt] = useState<string | null>(null);
  const [activeTab, setActiveTab] = useState<LtsReportTabKey>("all");
  const [autoSaveStatus, setAutoSaveStatus] = useState("Memuat data LTS");
  const [processingKey, setProcessingKey] = useState<LtsDatasetKey | null>(null);
  const [uploadErrors, setUploadErrors] = useState<Partial<Record<LtsDatasetKey, string>>>({});
  const [downloadState, setDownloadState] = useState<"idle" | "working" | "error">("idle");
  const [masterUnlockToken, setMasterUnlockToken] = useState<string | null>(null);
  const [masterUnlockOpen, setMasterUnlockOpen] = useState(false);
  const visibleNavItems = useMemo(
    () => navItems.filter((item) => item.id !== "master" || permissions.canAccessMaster),
    [permissions.canAccessMaster],
  );
  const renderView: View = view === "master" && (!permissions.canAccessMaster || !masterUnlockToken) ? "home" : view;

  const filteredDatasetState = useMemo(() => combineDatasetsForRange(datasetArchive, periodFilter), [datasetArchive, periodFilter]);
  const reportPeriodLabel = ltsPeriodFilterLabel(periodFilter);
  const reportCutoffLabel = ltsPeriodCutoffLabel(periodFilter);
  const reports = useMemo(() => {
    const output: Partial<Record<LtsReportTabKey, LtsComputedReport>> = {};
    for (const tab of ltsReportTabs) {
      const dataset = filteredDatasetState.datasets[tab.datasetKey];
      if (!dataset) continue;
      output[tab.key] = { ...computeLtsReport(tab.key, filteredDatasetState.datasets), cutoffLabel: reportCutoffLabel, periodLabel: reportPeriodLabel };
    }
    return output;
  }, [filteredDatasetState.datasets, reportCutoffLabel, reportPeriodLabel]);

  const activeReport = reports[activeTab];
  const readyCount = Object.values(ltsDatasetDefinitions).filter((definition) => datasets[definition.datasetKey]).length;
  const periodOptions = useMemo(() => {
    const options = periodOptionsFromArchive(datasetArchive);
    return options.length > 0 ? options : [{ label: periodLabelFromLtsPeriodKey(selectedPeriodKey), periodKey: selectedPeriodKey }];
  }, [datasetArchive, selectedPeriodKey]);

  const syncState = useCallback((state: DashboardPersistedState, successStatus: string) => {
    persistLocally(state);
    void saveStateToServer(state, masterUnlockToken)
      .then((result) => setAutoSaveStatus(result.configured ? successStatus : "Data tersimpan di browser; database server belum dikonfigurasi"))
      .catch(() => setAutoSaveStatus("Data tersimpan di browser; sync server gagal"));
  }, [masterUnlockToken]);

  useEffect(() => {
    const seedArchive = initialSeedArchive();
    const shouldRefreshSeed = window.localStorage.getItem(DATA_VERSION_KEY) !== ltsSeedData.dataVersion;
    const stored = readLocalState();
    let currentArchive = seedArchive;
    let currentPeriodKey = latestLtsPeriodKey(seedArchive);
    let currentUpdatedAt: string | null = null;

    if (stored) {
      const storedArchive = stored.datasetArchive ?? archiveFromDatasets(stored.datasets);
      const nextArchive = shouldRefreshSeed ? mergeDatasetArchives(storedArchive, seedArchive) : mergeDatasetArchives(seedArchive, storedArchive);
      const nextPeriodKey = shouldRefreshSeed ? latestLtsPeriodKey(seedArchive) : stored.selectedPeriodKey ?? latestLtsPeriodKey(nextArchive);
      currentArchive = nextArchive;
      currentPeriodKey = nextPeriodKey;
      currentUpdatedAt = shouldRefreshSeed ? null : stored.updatedAt;
      setDatasetArchive(nextArchive);
      setSelectedPeriodKey(nextPeriodKey);
      setPeriodFilter(defaultPeriodFilter(nextPeriodKey));
      setDatasets(activeDatasetsForPeriod(nextArchive, nextPeriodKey));
      setUpdatedAt(currentUpdatedAt);
      setAutoSaveStatus("Data lokal dimuat");
    } else {
      setDatasetArchive(seedArchive);
      setSelectedPeriodKey(currentPeriodKey);
      setPeriodFilter(defaultPeriodFilter(currentPeriodKey));
      setDatasets(activeDatasetsForPeriod(seedArchive, currentPeriodKey));
      setUpdatedAt(null);
      setAutoSaveStatus("Data seed LTS dimuat");
    }

    persistLocally({
      datasets: activeDatasetsForPeriod(currentArchive, currentPeriodKey),
      datasetArchive: currentArchive,
      selectedPeriodKey: currentPeriodKey,
      updatedAt: currentUpdatedAt,
    });
    window.localStorage.setItem(DATA_VERSION_KEY, ltsSeedData.dataVersion);

    void loadStateFromServer()
      .then(({ configured, state }) => {
        if (state) {
          setDatasetArchive((currentArchive) => {
            const serverArchive = state.datasetArchive ?? archiveFromDatasets(state.datasets);
            const mergedArchive = shouldRefreshSeed ? mergeDatasetArchives(serverArchive, currentArchive) : mergeDatasetArchives(currentArchive, serverArchive);
            const nextPeriodKey = shouldRefreshSeed ? currentPeriodKey : state.selectedPeriodKey ?? latestLtsPeriodKey(mergedArchive);
            setSelectedPeriodKey(nextPeriodKey);
            setPeriodFilter(defaultPeriodFilter(nextPeriodKey));
            setDatasets(activeDatasetsForPeriod(mergedArchive, nextPeriodKey));
            return mergedArchive;
          });
          setUpdatedAt(shouldRefreshSeed ? currentUpdatedAt : state.updatedAt);
          setAutoSaveStatus("Data server dimuat");
        } else if (!configured) {
          setAutoSaveStatus("Database server belum dikonfigurasi; data tersimpan di browser ini");
        }
      })
      .catch(() => setAutoSaveStatus("Data lokal dimuat; server belum bisa diakses"));
  }, []);

  async function handleUpload(datasetKey: LtsDatasetKey, file: File | undefined) {
    if (!file) return;
    setProcessingKey(datasetKey);
    setUploadErrors((current) => ({ ...current, [datasetKey]: undefined }));
    try {
      const dataset = await readLtsDatasetFromWorkbook(file, datasetKey);
      const timestamp = new Date().toISOString();
      const nextArchive = {
        ...datasetArchive,
        [datasetKey]: {
          ...(datasetArchive[datasetKey] ?? {}),
          [dataset.periodKey]: dataset,
        },
      };
      const nextDatasets = activeDatasetsForPeriod(nextArchive, dataset.periodKey);
      const state = { datasets: nextDatasets, datasetArchive: nextArchive, selectedPeriodKey: dataset.periodKey, updatedAt: timestamp };
      setDatasetArchive(nextArchive);
      setSelectedPeriodKey(dataset.periodKey);
      setPeriodFilter(defaultPeriodFilter(dataset.periodKey));
      setDatasets(nextDatasets);
      setUpdatedAt(timestamp);
      syncState(state, `${ltsDatasetDefinitions[datasetKey].label} tersimpan dan tersinkron`);
    } catch (error) {
      setUploadErrors((current) => ({
        ...current,
        [datasetKey]: error instanceof Error ? error.message : "File Excel tidak bisa dibaca.",
      }));
    } finally {
      setProcessingKey(null);
    }
  }

  async function downloadPng() {
    if (!activeReport) return;
    setDownloadState("working");
    try {
      await createLtsReportPng(activeReport, new Date());
      setDownloadState("idle");
    } catch {
      setDownloadState("error");
    }
  }

  function selectPeriod(periodKey: string) {
    const nextPeriodFilter = defaultPeriodFilter(periodKey);
    setSelectedPeriodKey(periodKey);
    setPeriodFilter(nextPeriodFilter);
    setDatasets(activeDatasetsForPeriod(datasetArchive, periodKey));
    const state = {
      datasets: activeDatasetsForPeriod(datasetArchive, periodKey),
      datasetArchive,
      selectedPeriodKey: periodKey,
      updatedAt,
    };
    persistLocally(state);
    setAutoSaveStatus("Periode LTS dipilih");
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
    <main className={pageBackground}>
      <header className="sticky top-0 z-40 border-b border-violet-300/20 bg-[#071426]/95 bg-[radial-gradient(circle_at_12%_0%,rgba(124,74,196,0.22),transparent_28%),linear-gradient(135deg,#061426_0%,#0a1f3b_54%,#07182d_100%)] shadow-[0_22px_48px_-34px_rgba(0,0,0,0.92)] backdrop-blur-xl">
        <div className="mx-auto flex max-w-[1920px] flex-col gap-3 px-4 py-3 sm:px-6 xl:flex-row xl:items-center xl:justify-between">
          <div className="flex items-center gap-3 rounded-xl border border-white/10 bg-white/[0.07] px-2.5 py-2 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)]">
            <div className="flex h-11 w-11 items-center justify-center overflow-hidden rounded-lg bg-white/[0.08] shadow-[0_14px_30px_-16px_rgba(154,106,214,0.9)] ring-1 ring-violet-300/25">
              <Image alt="LTS" className="h-full w-full object-cover" height={80} priority src="/module-covers/lts-design-card.png" width={80} />
            </div>
            <div>
              <p className="text-sm font-bold uppercase tracking-normal text-white">Reporting LTS 2026</p>
              <p className="text-xs font-semibold uppercase text-violet-100/72">{readyCount} / 3 DATASETS READY</p>
              <p className="mt-0.5 max-w-[280px] truncate text-[10px] font-semibold uppercase text-violet-100/50">
                {autoSaveStatus} | {formatSavedAt(updatedAt)}
              </p>
            </div>
          </div>
          <nav className="scrollbar-thin flex min-w-0 gap-1.5 overflow-x-auto rounded-xl border border-white/10 bg-white/[0.08] p-1.5 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)] backdrop-blur">
            <button onClick={() => router.push("/portal")} className="inline-flex h-9 shrink-0 items-center gap-2 rounded-md px-3 text-sm font-semibold text-slate-100/85 transition hover:bg-white/12 hover:text-white">
              <ArrowLeft className="h-4 w-4" />
              Portal
            </button>
            {visibleNavItems.map((item) => {
              const Icon = item.icon;
              const active = renderView === item.id;
              const masterLocked = item.id === "master" && !masterUnlockToken;
              return (
                <button
                  className={`inline-flex h-9 shrink-0 items-center gap-2 rounded-md px-3 text-sm font-semibold transition ${
                    active ? "bg-white text-[#24103f] shadow-[0_16px_30px_-20px_rgba(154,106,214,0.95)]" : "text-slate-100/85 hover:bg-white/12 hover:text-white"
                  }`}
                  key={item.id}
                  onClick={() => {
                    if (item.id === "master") {
                      openMasterView();
                    } else {
                      setMasterUnlockToken(null);
                      setView(item.id);
                    }
                  }}
                >
                  <Icon className="h-4 w-4" />
                  <span>{item.label}</span>
                  {masterLocked && <LockKeyhole aria-label="Master terkunci" className="h-3.5 w-3.5 opacity-70" />}
                </button>
              );
            })}
          </nav>
        </div>
      </header>

      <div className="mx-auto max-w-[1920px] px-3 py-5 sm:px-5 lg:px-6">
        {renderView === "home" && (
          <DashboardHome
            activeTab={activeTab}
            onGoMaster={openMasterView}
            onTabChange={setActiveTab}
            onPeriodFilterChange={setPeriodFilter}
            onPeriodChange={selectPeriod}
            periodFilter={periodFilter}
            periodOptions={periodOptions}
            periodValidationMessage={filteredDatasetState.validationMessage}
            readyCount={readyCount}
            reports={reports}
            selectedPeriodKey={selectedPeriodKey}
          />
        )}

        {renderView === "master" && permissions.canAccessMaster && (
          <div className="space-y-5">
            <section className={`${glassPanel} flex flex-wrap items-end justify-between gap-4 p-5`}>
              <div>
                <p className="text-sm font-semibold uppercase text-violet-200">MASTER DATA LTS</p>
                <h1 className="mt-1 text-3xl font-bold tracking-normal text-white">Upload Master LTS</h1>
                <p className="mt-2 text-sm font-medium text-violet-100/62">Setiap master berdiri sendiri dan hanya memperbarui report category terkait.</p>
              </div>
              <StatusPill tone={readyCount === 3 ? "good" : "warn"}>Data completeness {readyCount} / 3</StatusPill>
            </section>
            <div className="grid gap-4 xl:grid-cols-3">
              {(Object.keys(ltsDatasetDefinitions) as LtsDatasetKey[]).map((datasetKey) => (
                <MasterUploadCard
                  key={datasetKey}
                  dataset={datasets[datasetKey]}
                  datasetKey={datasetKey}
                  error={uploadErrors[datasetKey]}
                  onUpload={handleUpload}
                  processing={processingKey === datasetKey}
                />
              ))}
            </div>
          </div>
        )}

        {renderView === "report" && (
          <div className="space-y-5">
            <section className={`${glassPanel} flex flex-wrap items-end justify-between gap-4 p-5`}>
              <div>
                <p className="text-sm font-semibold uppercase text-violet-200">REPORT LTS</p>
                <h1 className="mt-1 text-3xl font-bold tracking-normal text-white">{activeReport?.label ?? "LTS REPORT"}</h1>
                <p className="mt-2 text-sm font-medium text-violet-100/62">{activeReport ? `${activeReport.sourceLabel} | ${activeReport.periodLabel} | ${activeReport.cutoffLabel}` : "Pilih tab dan upload master terkait."}</p>
              </div>
              <button className={primaryButton} disabled={!activeReport || downloadState === "working"} onClick={downloadPng}>
                {downloadState === "working" ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
                PNG Report
              </button>
            </section>
            {downloadState === "error" && <div className="rounded-xl border border-amber-300/20 bg-amber-400/10 px-4 py-3 text-sm font-semibold text-amber-100">PNG belum berhasil dibuat. Coba tekan PNG Report sekali lagi.</div>}
            <div className="flex flex-wrap gap-2">
              {ltsReportTabs.map((tab) => {
                const ready = Boolean(reports[tab.key]);
                const active = activeTab === tab.key;
                return (
                  <button
                    key={tab.key}
                    className={`inline-flex h-11 items-center gap-2 rounded-xl border px-4 text-sm font-bold uppercase transition ${
                      active ? "border-violet-300/35 bg-[#321653] text-white shadow-[0_18px_36px_-24px_rgba(124,74,196,0.95)]" : "border-violet-200/12 bg-white/[0.05] text-violet-100/70 hover:bg-white/[0.09] hover:text-white"
                    }`}
                    onClick={() => setActiveTab(tab.key)}
                  >
                    {ready ? <ShieldCheck className="h-4 w-4 text-emerald-300" /> : <RefreshCw className="h-4 w-4 text-violet-200/60" />}
                    {tab.label}
                  </button>
                );
              })}
            </div>
            {activeReport ? (
              <>
                <ReportDataTable table={activeReport.witel} />
                <ReportDataTable table={activeReport.hotda} />
              </>
            ) : (
              <section className={`${glassPanel} p-8 text-center`}>
                <FileSpreadsheet className="mx-auto h-10 w-10 text-violet-200" />
                <h2 className="mt-4 text-xl font-bold text-white">{ltsDatasetDefinitions[ltsReportTabs.find((tab) => tab.key === activeTab)?.datasetKey ?? "sales"].label} belum diupload</h2>
                <p className="mx-auto mt-2 max-w-xl text-sm leading-6 text-violet-100/65">Report ini membutuhkan master yang sesuai dengan mapping LTS. Upload dataset tersebut di Master Data LTS.</p>
                <button className={`${primaryButton} mt-5`} onClick={openMasterView}>Go to Master Data</button>
              </section>
            )}
          </div>
        )}
      </div>
      <MasterUnlockModal
        open={masterUnlockOpen}
        theme="lts"
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
