import { readDashboardStateByKey, writeDashboardStateByKey } from "@/features/shared/backend/supabase/dashboard-state";
import { supabaseGet, supabaseUpsert } from "@/features/shared/backend/supabase/rest";
import {
  legacyStorageDefinitionForKey,
  LOCAL_STORAGE_MIGRATION_VERSION,
  type LegacyStorageArea,
} from "@/features/shared/storage/legacy-storage-keys";
import { masterHeaders as hsiMasterHeaders } from "@/features/hsi/domain/reporting/reporting";
import { masterHeaders as wmsMasterHeaders } from "@/features/wms/domain/reporting/reporting";

type LegacyStorageSnapshotEntry = {
  area: LegacyStorageArea;
  key: string;
  value: string;
};

export type LegacyStorageSnapshot = {
  createdAt: string;
  keys: LegacyStorageSnapshotEntry[];
  migrationVersion: number;
  source: "legacy-localStorage";
};

export type MigrationProductResult = {
  conflicts: string[];
  localRows: number;
  migratedRows: number;
  status: "success" | "skipped" | "conflict" | "error";
  verifiedRows: number;
};

export type LegacyStorageMigrationReport = {
  backupId: string | null;
  conflicts: string[];
  errors: string[];
  migratedKeys: string[];
  migrationVersion: number;
  products: Record<"HSI" | "WMS" | "LTS" | "METABRIGHT", MigrationProductResult>;
  skippedKeys: string[];
  success: boolean;
};

type DashboardRecord = {
  businessKey: string;
  datasetKind: string;
  payload: unknown;
  periodKey: string | null;
  product: "HSI" | "WMS" | "LTS";
  rowIndex: number | null;
};

type LegacyStorageEntryRow = {
  backup_id?: string | null;
  migrated_at?: string;
  parsed_value?: unknown;
  product?: string;
  raw_value?: string | null;
  source_hash?: string;
  source_key?: string;
  storage_area?: LegacyStorageArea;
};

const hsiBusinessKeys = ["hsi-master-rows-v1", "hsi-period-archive-v1", "hsi-save-snapshots-v1", "hsi-report-history-v1"];
const wmsBusinessKeys = ["wms-master-rows-v1", "wms-period-archive-v1", "wms-save-snapshots-v1"];

const emptyProductResult = (): MigrationProductResult => ({
  conflicts: [],
  localRows: 0,
  migratedRows: 0,
  status: "skipped",
  verifiedRows: 0,
});

function stableJson(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableJson).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([key, entry]) => `${JSON.stringify(key)}:${stableJson(entry)}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
}

async function sha256(value: string) {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return Array.from(new Uint8Array(digest))
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function parseJson(value: string): unknown {
  try {
    return JSON.parse(value);
  } catch {
    return value;
  }
}

function entryValue(snapshot: LegacyStorageSnapshot, key: string) {
  return snapshot.keys.find((entry) => entry.area === "localStorage" && entry.key === key)?.value ?? null;
}

function parsedEntry(snapshot: LegacyStorageSnapshot, key: string) {
  const value = entryValue(snapshot, key);
  return value === null ? null : parseJson(value);
}

function periodKeyFromRows(rows: unknown, fallback: string | null) {
  if (!Array.isArray(rows)) return fallback;
  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    const nper = row[hsiMasterHeaders.indexOf("NPER")] ?? row[wmsMasterHeaders.indexOf("NPER")] ?? row[wmsMasterHeaders.indexOf("PERIOD")];
    const normalized = String(nper ?? "").replace(/\D/g, "").slice(0, 6);
    if (/^\d{6}$/.test(normalized)) return normalized;
  }
  return fallback;
}

function normalizePeriodKey(value: unknown): string | null {
  const normalized = String(value ?? "").replace(/\D/g, "").slice(0, 6);
  return /^\d{6}$/.test(normalized) ? normalized : null;
}

function rawFieldsFromRow(headers: readonly string[], row: unknown) {
  if (!Array.isArray(row)) return {};
  return Object.fromEntries(headers.map((header, index) => [header, row[index] ?? null]));
}

function fieldValue(headers: readonly string[], row: unknown, header: string) {
  if (!Array.isArray(row)) return null;
  const index = headers.indexOf(header);
  return index >= 0 ? row[index] ?? null : null;
}

function hasMeaningfulValue(value: unknown): boolean {
  if (value === null || value === undefined) return false;
  if (typeof value === "string") return value.trim().length > 0;
  if (typeof value === "number") return Number.isFinite(value);
  if (typeof value === "boolean") return true;
  if (Array.isArray(value)) return value.some(hasMeaningfulValue);
  if (typeof value === "object") return Object.values(value as Record<string, unknown>).some(hasMeaningfulValue);
  return false;
}

function isMeaningfulRow(row: unknown): boolean {
  return Array.isArray(row) ? row.some(hasMeaningfulValue) : hasMeaningfulValue(row);
}

function normalizeProductPayload(product: "HSI" | "WMS", row: unknown, periodKey: string | null) {
  const headers = product === "HSI" ? hsiMasterHeaders : wmsMasterHeaders;
  const fields = rawFieldsFromRow(headers, row);
  if (product === "HSI") {
    return {
      datel: fieldValue(headers, row, "DATEL"),
      lastUpdated: fieldValue(headers, row, "LAST UPDATED DATE"),
      ndem: fieldValue(headers, row, "NDEM"),
      nper: fieldValue(headers, row, "NPER"),
      orderId: fieldValue(headers, row, "ORDER ID"),
      periodKey,
      product,
      rawFields: fields,
      rawRow: row,
      source: fieldValue(headers, row, "PROVIDER") ?? fieldValue(headers, row, "CHANNEL"),
      statusResume: fieldValue(headers, row, "STATUS RESUME"),
      telda: fieldValue(headers, row, "DATEL"),
      witel: fieldValue(headers, row, "WITEL"),
    };
  }

  return {
    lastUpdated: fieldValue(headers, row, "LAST_UPDATED_DATE") ?? fieldValue(headers, row, "LAST_UPDATE_DATE"),
    nper: fieldValue(headers, row, "NPER") ?? fieldValue(headers, row, "PERIOD"),
    orderId: fieldValue(headers, row, "ORDER_ID") ?? fieldValue(headers, row, "ORDERID_WMS_NCX"),
    periodKey,
    product,
    rawFields: fields,
    rawRow: row,
    realization: fieldValue(headers, row, "Jml_Sales") ?? fieldValue(headers, row, "COUNT_SID"),
    source: fieldValue(headers, row, "PROVIDER") ?? fieldValue(headers, row, "CHANNEL"),
    statusResume: fieldValue(headers, row, "STATUS_RESUME"),
    targetOrTariff: fieldValue(headers, row, "TARIF_TOT") ?? fieldValue(headers, row, "PACKAGE_PRICE"),
    telda: fieldValue(headers, row, "TELKOM_DAERAH") ?? fieldValue(headers, row, "TELDA_NCX"),
    witel: fieldValue(headers, row, "WITEL"),
  };
}

function countRows(value: unknown): number {
  if (Array.isArray(value)) return value.length;
  if (!value || typeof value !== "object") return 0;
  return Object.values(value as Record<string, unknown>).reduce<number>((total, entry) => total + countRows(entry), 0);
}

function tableRowsFromObject(value: unknown): unknown[] {
  if (!value || typeof value !== "object") return [];
  const rows = (value as { rows?: unknown }).rows;
  return Array.isArray(rows) ? rows.filter(isMeaningfulRow) : [];
}

function tableTotalFromObject(value: unknown): unknown | null {
  if (!value || typeof value !== "object") return null;
  const total = (value as { total?: unknown }).total;
  return hasMeaningfulValue(total) ? total : null;
}

function pushComputedTableRecords(
  records: DashboardRecord[],
  params: {
    datasetKind: string;
    periodKey: string | null;
    product: "HSI" | "WMS";
    table: unknown;
  },
) {
  const rows = tableRowsFromObject(params.table);
  rows.forEach((row, rowIndex) => {
    records.push({
      businessKey: `${params.product}:${params.datasetKind}:${params.periodKey ?? "unknown"}:${rowIndex}:${stableJson(row).slice(0, 96)}`,
      datasetKind: params.datasetKind,
      payload: {
        periodKey: params.periodKey,
        product: params.product,
        row,
        source: "computed_report",
      },
      periodKey: params.periodKey,
      product: params.product,
      rowIndex,
    });
  });

  const total = tableTotalFromObject(params.table);
  if (total) {
    records.push({
      businessKey: `${params.product}:${params.datasetKind}_total:${params.periodKey ?? "unknown"}:${stableJson(total).slice(0, 96)}`,
      datasetKind: `${params.datasetKind}_total`,
      payload: {
        periodKey: params.periodKey,
        product: params.product,
        row: total,
        source: "computed_report",
      },
      periodKey: params.periodKey,
      product: params.product,
      rowIndex: null,
    });
  }
}

function recordsFromComputedReport(product: "HSI" | "WMS", periodKey: string | null, reportValue: unknown): DashboardRecord[] {
  if (!reportValue || typeof reportValue !== "object") return [];
  const report = reportValue as Record<string, unknown>;
  const records: DashboardRecord[] = [];

  pushComputedTableRecords(records, {
    datasetKind: "computed_witel_rows",
    periodKey,
    product,
    table: report.witel,
  });

  pushComputedTableRecords(records, {
    datasetKind: product === "HSI" ? "computed_branch_rows" : "computed_telda_rows",
    periodKey,
    product,
    table: product === "HSI" ? report.branch : report.telda,
  });

  if (product === "HSI" && report.targetStf && typeof report.targetStf === "object") {
    const targetStf = report.targetStf as Record<string, unknown>;
    pushComputedTableRecords(records, {
      datasetKind: "computed_stf_witel_rows",
      periodKey,
      product,
      table: targetStf.witel,
    });
    pushComputedTableRecords(records, {
      datasetKind: "computed_stf_branch_rows",
      periodKey,
      product,
      table: targetStf.branch,
    });
  }

  const dailyTrend = report.dailyTrend;
  if (Array.isArray(dailyTrend)) {
    dailyTrend.filter(isMeaningfulRow).forEach((row, rowIndex) => {
      records.push({
        businessKey: `${product}:computed_daily_trend:${periodKey ?? "unknown"}:${rowIndex}:${stableJson(row).slice(0, 96)}`,
        datasetKind: "computed_daily_trend",
        payload: {
          periodKey,
          product,
          row,
          source: "computed_report",
        },
        periodKey,
        product,
        rowIndex,
      });
    });
  }

  const validation = report.validation;
  if (hasMeaningfulValue(validation)) {
    records.push({
      businessKey: `${product}:computed_validation:${periodKey ?? "unknown"}:${stableJson(validation).slice(0, 96)}`,
      datasetKind: "computed_validation",
      payload: {
        periodKey,
        product,
        row: validation,
        source: "computed_report",
      },
      periodKey,
      product,
      rowIndex: null,
    });
  }

  return records;
}

function dashboardStateForHsi(snapshot: LegacyStorageSnapshot) {
  const periodArchive = parsedEntry(snapshot, "hsi-period-archive-v1");
  const selectedPeriodKey = parsedEntry(snapshot, "hsi-selected-period-v1");
  const rows = parsedEntry(snapshot, "hsi-master-rows-v1");
  const fallbackPeriodKey = normalizePeriodKey(selectedPeriodKey) ?? periodKeyFromRows(rows, null);
  const normalizedArchive =
    periodArchive && typeof periodArchive === "object"
      ? periodArchive
      : fallbackPeriodKey && Array.isArray(rows)
        ? {
            [fallbackPeriodKey]: {
              label: fallbackPeriodKey,
              periodKey: fallbackPeriodKey,
              reportHistory: parsedEntry(snapshot, "hsi-report-history-v1"),
              reportSnapshot: null,
              rows,
              savedAt: entryValue(snapshot, "hsi-master-saved-at-v1"),
              updatedAt: entryValue(snapshot, "hsi-master-saved-at-v1"),
            },
          }
        : null;
  if (!normalizedArchive || typeof fallbackPeriodKey !== "string") return null;

  return {
    dataVersion: entryValue(snapshot, "hsi-seed-data-version-v1"),
    masterHeadersCleared: entryValue(snapshot, "hsi-master-headers-cleared-v1") === "true",
    periodArchive: normalizedArchive,
    saveSnapshots: parsedEntry(snapshot, "hsi-save-snapshots-v1") ?? {},
    selectedPeriodKey: fallbackPeriodKey,
    updatedAt: entryValue(snapshot, "hsi-master-saved-at-v1"),
  };
}

function dashboardStateForWms(snapshot: LegacyStorageSnapshot) {
  const periodArchive = parsedEntry(snapshot, "wms-period-archive-v1");
  const selectedPeriodKey = parsedEntry(snapshot, "wms-selected-period-v1");
  const rows = parsedEntry(snapshot, "wms-master-rows-v1");
  const fallbackPeriodKey = normalizePeriodKey(selectedPeriodKey) ?? periodKeyFromRows(rows, null);
  const normalizedArchive =
    periodArchive && typeof periodArchive === "object"
      ? periodArchive
      : fallbackPeriodKey && Array.isArray(rows)
        ? {
            [fallbackPeriodKey]: {
              label: fallbackPeriodKey,
              periodKey: fallbackPeriodKey,
              rows,
              savedAt: entryValue(snapshot, "wms-master-saved-at-v1"),
              updatedAt: entryValue(snapshot, "wms-master-saved-at-v1"),
            },
          }
        : null;
  if (!normalizedArchive || typeof fallbackPeriodKey !== "string") return null;

  return {
    periodArchive: normalizedArchive,
    saveSnapshots: parsedEntry(snapshot, "wms-save-snapshots-v1") ?? {},
    selectedPeriodKey: fallbackPeriodKey,
    updatedAt: entryValue(snapshot, "wms-master-saved-at-v1"),
  };
}

function dashboardStateForLts(snapshot: LegacyStorageSnapshot) {
  const state = parsedEntry(snapshot, "lts-dashboard-state-v1");
  return state && typeof state === "object" ? state : null;
}

function recordsFromPeriodArchive(product: "HSI" | "WMS", archiveValue: unknown): DashboardRecord[] {
  if (!archiveValue || typeof archiveValue !== "object") return [];
  const records: DashboardRecord[] = [];
  for (const [periodKey, entry] of Object.entries(archiveValue as Record<string, unknown>)) {
    if (!entry || typeof entry !== "object") continue;
    const archiveEntry = entry as { reportSnapshot?: unknown; rows?: unknown };
    const rows = Array.isArray(archiveEntry.rows) ? archiveEntry.rows.filter(isMeaningfulRow) : [];
    rows.forEach((row, rowIndex) => {
      const payload = normalizeProductPayload(product, row, periodKey);
      records.push({
        businessKey: `${product}:${periodKey}:${rowIndex}:${stableJson(payload).slice(0, 96)}`,
        datasetKind: "master_rows",
        payload,
        periodKey,
        product,
        rowIndex,
      });
    });
    records.push(...recordsFromComputedReport(product, periodKey, archiveEntry.reportSnapshot));
  }
  return records;
}

function recordsFromMasterRows(product: "HSI" | "WMS", rowsValue: unknown, periodKey: string | null, datasetKind = "master_rows"): DashboardRecord[] {
  if (!Array.isArray(rowsValue)) return [];
  const rows = rowsValue.filter(isMeaningfulRow);
  const resolvedPeriodKey = periodKeyFromRows(rows, periodKey);
  return rows.map((row, rowIndex) => {
    const payload = normalizeProductPayload(product, row, resolvedPeriodKey);
    return {
      businessKey: `${product}:${resolvedPeriodKey ?? "unknown"}:${rowIndex}:${stableJson(payload).slice(0, 96)}`,
      datasetKind,
      payload,
      periodKey: resolvedPeriodKey,
      product,
      rowIndex,
    };
  });
}

function recordsFromSaveSnapshots(product: "HSI" | "WMS", snapshotsValue: unknown): DashboardRecord[] {
  if (!snapshotsValue || typeof snapshotsValue !== "object") return [];
  const records: DashboardRecord[] = [];
  for (const [periodKey, versions] of Object.entries(snapshotsValue as Record<string, unknown>)) {
    if (!Array.isArray(versions)) continue;
    versions.forEach((version, versionIndex) => {
      const rows = version && typeof version === "object" ? (version as { rows?: unknown }).rows : null;
      if (!Array.isArray(rows)) return;
      records.push(
        ...recordsFromMasterRows(product, rows, periodKey, `snapshot_v${versionIndex + 1}`).map((record) => ({
          ...record,
          businessKey: `${record.businessKey}:snapshot:${versionIndex + 1}`,
        })),
      );
    });
  }
  return records;
}

function recordsFromLtsState(state: unknown): DashboardRecord[] {
  if (!state || typeof state !== "object") return [];
  const archive = (state as { datasetArchive?: unknown }).datasetArchive;
  if (!archive || typeof archive !== "object") return [];
  const records: DashboardRecord[] = [];

  for (const [datasetKind, periods] of Object.entries(archive as Record<string, unknown>)) {
    if (!periods || typeof periods !== "object") continue;
    for (const [periodKey, dataset] of Object.entries(periods as Record<string, unknown>)) {
      const rows = dataset && typeof dataset === "object" ? (dataset as { rows?: unknown }).rows : null;
      if (!Array.isArray(rows)) continue;
      rows.forEach((row, rowIndex) => {
        records.push({
          businessKey: `${datasetKind}:${periodKey}:${rowIndex}:${stableJson(row).slice(0, 96)}`,
          datasetKind,
          payload: row,
          periodKey,
          product: "LTS",
          rowIndex,
        });
      });
    }
  }

  return records;
}

function latestEntryByKey(entries: LegacyStorageEntryRow[], key: string) {
  return entries
    .filter((entry) => entry.source_key === key)
    .sort((left, right) => String(right.migrated_at ?? "").localeCompare(String(left.migrated_at ?? "")))[0] ?? null;
}

function parsedLegacyEntryValue(entry: LegacyStorageEntryRow | null) {
  if (!entry) return null;
  if (entry.parsed_value !== null && entry.parsed_value !== undefined) return entry.parsed_value;
  return entry.raw_value ? parseJson(entry.raw_value) : null;
}

function dashboardStateForHsiEntries(entries: LegacyStorageEntryRow[]) {
  const archive = parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-period-archive-v1"));
  const rows = parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-rows-v1"));
  const selectedPeriod = parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-selected-period-v1"));
  const periodKey = normalizePeriodKey(selectedPeriod) ?? periodKeyFromRows(rows, null);
  const periodArchive =
    archive && typeof archive === "object"
      ? archive
      : periodKey && Array.isArray(rows)
        ? {
            [periodKey]: {
              label: periodKey,
              periodKey,
              reportHistory: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-report-history-v1")),
              reportSnapshot: null,
              rows,
              savedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-saved-at-v1")),
              updatedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-saved-at-v1")),
            },
          }
        : null;
  if (!periodArchive || !periodKey) return null;
  return {
    dataVersion: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-seed-data-version-v1")),
    masterHeadersCleared: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-headers-cleared-v1")) === "true",
    periodArchive,
    saveSnapshots: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-save-snapshots-v1")) ?? {},
    selectedPeriodKey: periodKey,
    updatedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-saved-at-v1")),
  };
}

function dashboardStateForWmsEntries(entries: LegacyStorageEntryRow[]) {
  const archive = parsedLegacyEntryValue(latestEntryByKey(entries, "wms-period-archive-v1"));
  const rows = parsedLegacyEntryValue(latestEntryByKey(entries, "wms-master-rows-v1"));
  const selectedPeriod = parsedLegacyEntryValue(latestEntryByKey(entries, "wms-selected-period-v1"));
  const periodKey = normalizePeriodKey(selectedPeriod) ?? periodKeyFromRows(rows, null);
  const periodArchive =
    archive && typeof archive === "object"
      ? archive
      : periodKey && Array.isArray(rows)
        ? {
            [periodKey]: {
              label: periodKey,
              periodKey,
              rows,
              savedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "wms-master-saved-at-v1")),
              updatedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "wms-master-saved-at-v1")),
            },
          }
        : null;
  if (!periodArchive || !periodKey) return null;
  return {
    periodArchive,
    saveSnapshots: parsedLegacyEntryValue(latestEntryByKey(entries, "wms-save-snapshots-v1")) ?? {},
    selectedPeriodKey: periodKey,
    updatedAt: parsedLegacyEntryValue(latestEntryByKey(entries, "wms-master-saved-at-v1")),
  };
}

function recordsFromLegacyEntries(entries: LegacyStorageEntryRow[]) {
  const hsiState = dashboardStateForHsiEntries(entries);
  const wmsState = dashboardStateForWmsEntries(entries);
  const hsiArchiveRecords = recordsFromPeriodArchive(
    "HSI",
    hsiState && typeof hsiState === "object" ? (hsiState as { periodArchive?: unknown }).periodArchive : null,
  );
  const wmsArchiveRecords = recordsFromPeriodArchive(
    "WMS",
    wmsState && typeof wmsState === "object" ? (wmsState as { periodArchive?: unknown }).periodArchive : null,
  );
  return {
    hsiState,
    records: [
      ...hsiArchiveRecords,
      ...(hsiArchiveRecords.length === 0
        ? recordsFromMasterRows(
            "HSI",
            parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-master-rows-v1")),
            hsiState && typeof hsiState === "object" ? String((hsiState as { selectedPeriodKey?: unknown }).selectedPeriodKey ?? "") : null,
            "active_master_rows",
          )
        : []),
      ...recordsFromSaveSnapshots("HSI", parsedLegacyEntryValue(latestEntryByKey(entries, "hsi-save-snapshots-v1"))),
      ...wmsArchiveRecords,
      ...(wmsArchiveRecords.length === 0
        ? recordsFromMasterRows(
            "WMS",
            parsedLegacyEntryValue(latestEntryByKey(entries, "wms-master-rows-v1")),
            wmsState && typeof wmsState === "object" ? String((wmsState as { selectedPeriodKey?: unknown }).selectedPeriodKey ?? "") : null,
            "active_master_rows",
          )
        : []),
      ...recordsFromSaveSnapshots("WMS", parsedLegacyEntryValue(latestEntryByKey(entries, "wms-save-snapshots-v1"))),
    ],
    wmsState,
  };
}

async function maybeWriteDashboardState(
  key: "default" | "wms" | "lts",
  payload: unknown,
  report: LegacyStorageMigrationReport,
  product: "HSI" | "WMS" | "LTS",
) {
  if (!payload) return;
  const existing = await readDashboardStateByKey(key);
  const existingState = "state" in existing.body ? existing.body.state : null;
  if (existingState && stableJson(existingState) !== stableJson(payload)) {
    const conflict = `${product}: dashboard_state.${key} already contains different data; legacy state was preserved as backup and row records, but dashboard_state was not overwritten.`;
    report.conflicts.push(conflict);
    report.products[product].conflicts.push(conflict);
    report.products[product].status = "conflict";
    return;
  }

  const result = await writeDashboardStateByKey(key, payload);
  if (result.status && result.status >= 400) {
    const message = `${product}: failed to write dashboard_state.${key}`;
    report.errors.push(message);
    report.products[product].status = "error";
  }
}

async function upsertRows(table: string, rows: unknown[], onConflict = "id") {
  if (rows.length === 0) return { ok: true, status: 204 };
  const result = await supabaseUpsert(table, rows, onConflict);
  return { ok: !result.error, status: result.status, error: result.error };
}

async function persistDashboardRecords(
  records: DashboardRecord[],
  report: LegacyStorageMigrationReport,
  options: { backupId: string | null; source: string },
) {
  const recordsByProduct = records.reduce(
    (accumulator, record) => ({ ...accumulator, [record.product]: (accumulator[record.product] ?? 0) + 1 }),
    {} as Partial<Record<"HSI" | "WMS" | "LTS", number>>,
  );

  const importRows = await Promise.all(
    (["HSI", "WMS", "LTS"] as const).map(async (product) => ({
      backup_id: options.backupId,
      id: `${options.source}-${product.toLowerCase()}-${await sha256(`${options.source}:${product}`)}`,
      imported_at: new Date().toISOString(),
      metadata: { source: options.source },
      product,
      row_count: recordsByProduct[product] ?? 0,
      source_hash: await sha256(`${options.source}:${product}`),
      source_key: `${product.toLowerCase()}-${options.source}`,
      updated_at: new Date().toISOString(),
    })),
  );
  const importResult = await upsertRows("reporting_imports", importRows);
  if (!importResult.ok) report.errors.push(`Gagal menyimpan reporting_imports: ${importResult.error ?? importResult.status}`);

  const importIdByProduct = Object.fromEntries(importRows.map((row) => [row.product, row.id])) as Record<"HSI" | "WMS" | "LTS", string>;
  const recordRows = await Promise.all(
    records.map(async (record) => {
      const sourceHash = await sha256(stableJson(record.payload));
      return {
        business_key: record.businessKey,
        dataset_kind: record.datasetKind,
        id: `${record.product}-${record.datasetKind}-${record.periodKey ?? "none"}-${record.rowIndex ?? "x"}-${sourceHash.slice(0, 24)}`,
        import_id: importIdByProduct[record.product],
        imported_at: new Date().toISOString(),
        payload: record.payload,
        period_key: record.periodKey,
        product: record.product,
        row_index: record.rowIndex,
        source_hash: sourceHash,
        updated_at: new Date().toISOString(),
      };
    }),
  );
  const recordResult = await upsertRows("dashboard_records", recordRows);
  if (!recordResult.ok) report.errors.push(`Gagal menyimpan dashboard_records: ${recordResult.error ?? recordResult.status}`);

  for (const product of ["HSI", "WMS", "LTS"] as const) {
    const rowCount = recordsByProduct[product] ?? 0;
    report.products[product].localRows = rowCount;
    report.products[product].migratedRows = recordResult.ok ? rowCount : 0;
    report.products[product].verifiedRows = recordResult.ok ? rowCount : 0;
    if (report.products[product].status !== "conflict" && report.products[product].status !== "error") {
      report.products[product].status = rowCount > 0 ? "success" : "skipped";
    }
  }
}

export async function migrateLegacyStorageSnapshot(snapshot: LegacyStorageSnapshot): Promise<LegacyStorageMigrationReport> {
  const report: LegacyStorageMigrationReport = {
    backupId: null,
    conflicts: [],
    errors: [],
    migratedKeys: [],
    migrationVersion: LOCAL_STORAGE_MIGRATION_VERSION,
    products: {
      HSI: emptyProductResult(),
      LTS: emptyProductResult(),
      METABRIGHT: emptyProductResult(),
      WMS: emptyProductResult(),
    },
    skippedKeys: [],
    success: false,
  };

  if (snapshot.migrationVersion !== LOCAL_STORAGE_MIGRATION_VERSION || snapshot.source !== "legacy-localStorage") {
    report.errors.push("Snapshot migrasi tidak valid atau versinya tidak cocok.");
    return report;
  }

  const snapshotHash = await sha256(stableJson(snapshot));
  const backupId = `legacy-${LOCAL_STORAGE_MIGRATION_VERSION}-${snapshotHash.slice(0, 24)}`;
  report.backupId = backupId;

  const backupResult = await upsertRows("legacy_storage_backups", [
    {
      created_at: snapshot.createdAt,
      id: backupId,
      key_count: snapshot.keys.length,
      migration_version: snapshot.migrationVersion,
      payload: snapshot,
      source: snapshot.source,
      source_hash: snapshotHash,
    },
  ]);
  if (!backupResult.ok) {
    report.errors.push(`Gagal menyimpan backup localStorage: ${backupResult.error ?? backupResult.status}`);
    return report;
  }

  const entryRows = [];
  const logRows = [];
  for (const entry of snapshot.keys) {
    const definition = legacyStorageDefinitionForKey(entry.key, entry.area);
    const parsedValue = parseJson(entry.value);
    const sourceHash = await sha256(`${entry.area}:${entry.key}:${entry.value}`);
    const product = definition?.product ?? "PORTAL";
    const classification = definition?.classification ?? "TEMPORARY_SESSION";

    if (classification === "TEMPORARY_SESSION") {
      report.skippedKeys.push(`${entry.area}.${entry.key}`);
      continue;
    }

    entryRows.push({
      backup_id: backupId,
      classification,
      id: `${backupId}-${entry.area}-${entry.key}`,
      migrated_at: new Date().toISOString(),
      migration_version: snapshot.migrationVersion,
      parsed_value: typeof parsedValue === "string" ? null : parsedValue,
      product,
      raw_value: entry.value,
      source_hash: sourceHash,
      source_key: entry.key,
      storage_area: entry.area,
    });
    report.migratedKeys.push(`${entry.area}.${entry.key}`);

    logRows.push({
      details: { proposedTable: definition?.proposedTable ?? "legacy_storage_entries" },
      id: `${LOCAL_STORAGE_MIGRATION_VERSION}-${entry.key}-${sourceHash.slice(0, 24)}`,
      migrated_at: new Date().toISOString(),
      migration_version: snapshot.migrationVersion,
      product,
      row_count: countRows(parsedValue),
      source_hash: sourceHash,
      source_key: entry.key,
      status: "success",
    });
  }

  const entryResult = await upsertRows("legacy_storage_entries", entryRows);
  if (!entryResult.ok) report.errors.push(`Gagal menyimpan detail key localStorage: ${entryResult.error ?? entryResult.status}`);

  const hsiState = dashboardStateForHsi(snapshot);
  const wmsState = dashboardStateForWms(snapshot);
  const ltsState = dashboardStateForLts(snapshot);
  await maybeWriteDashboardState("default", hsiState, report, "HSI");
  await maybeWriteDashboardState("wms", wmsState, report, "WMS");
  await maybeWriteDashboardState("lts", ltsState, report, "LTS");

  const hsiArchiveRecords = recordsFromPeriodArchive(
    "HSI",
    hsiState && typeof hsiState === "object" ? (hsiState as { periodArchive?: unknown }).periodArchive : null,
  );
  const wmsArchiveRecords = recordsFromPeriodArchive(
    "WMS",
    wmsState && typeof wmsState === "object" ? (wmsState as { periodArchive?: unknown }).periodArchive : null,
  );
  const records = [
    ...hsiArchiveRecords,
    ...(hsiArchiveRecords.length === 0
      ? recordsFromMasterRows(
          "HSI",
          parsedEntry(snapshot, "hsi-master-rows-v1"),
          hsiState && typeof hsiState === "object" ? String((hsiState as { selectedPeriodKey?: unknown }).selectedPeriodKey ?? "") : null,
          "active_master_rows",
        )
      : []),
    ...recordsFromSaveSnapshots("HSI", parsedEntry(snapshot, "hsi-save-snapshots-v1")),
    ...wmsArchiveRecords,
    ...(wmsArchiveRecords.length === 0
      ? recordsFromMasterRows(
          "WMS",
          parsedEntry(snapshot, "wms-master-rows-v1"),
          wmsState && typeof wmsState === "object" ? String((wmsState as { selectedPeriodKey?: unknown }).selectedPeriodKey ?? "") : null,
          "active_master_rows",
        )
      : []),
    ...recordsFromSaveSnapshots("WMS", parsedEntry(snapshot, "wms-save-snapshots-v1")),
    ...recordsFromLtsState(ltsState),
  ];
  await persistDashboardRecords(records, report, { backupId, source: "legacy-localStorage" });

  const logResult = await upsertRows("migration_logs", logRows);
  if (!logResult.ok) report.errors.push(`Gagal menyimpan migration_logs: ${logResult.error ?? logResult.status}`);

  report.success = report.errors.length === 0;
  return report;
}

export async function normalizeExistingHsiWmsLegacyEntries(): Promise<LegacyStorageMigrationReport> {
  const report: LegacyStorageMigrationReport = {
    backupId: null,
    conflicts: [],
    errors: [],
    migratedKeys: [],
    migrationVersion: LOCAL_STORAGE_MIGRATION_VERSION,
    products: {
      HSI: emptyProductResult(),
      LTS: emptyProductResult(),
      METABRIGHT: emptyProductResult(),
      WMS: emptyProductResult(),
    },
    skippedKeys: [],
    success: false,
  };

  const keys = [...hsiBusinessKeys, ...wmsBusinessKeys, "hsi-selected-period-v1", "wms-selected-period-v1", "hsi-master-saved-at-v1", "wms-master-saved-at-v1"];
  const result = await supabaseGet<LegacyStorageEntryRow[]>(
    "legacy_storage_entries",
    `select=backup_id,migrated_at,parsed_value,product,raw_value,source_hash,source_key,storage_area&source_key=in.(${keys.join(",")})`,
  );

  if (result.error) {
    report.errors.push(`Gagal membaca legacy_storage_entries: ${result.error}`);
    return report;
  }

  const entries = result.data ?? [];
  report.migratedKeys = Array.from(new Set(entries.map((entry) => `legacy_storage_entries.${entry.source_key ?? "unknown"}`)));
  const { hsiState, records, wmsState } = recordsFromLegacyEntries(entries);

  await maybeWriteDashboardState("default", hsiState, report, "HSI");
  await maybeWriteDashboardState("wms", wmsState, report, "WMS");
  await persistDashboardRecords(
    records,
    report,
    { backupId: entries.find((entry) => entry.backup_id)?.backup_id ?? null, source: "legacy-storage-renormalize" },
  );

  const logRows = await Promise.all(
    (["HSI", "WMS"] as const).map(async (product) => ({
      details: { source: "legacy_storage_entries", normalizedRecords: report.products[product].migratedRows },
      id: `${LOCAL_STORAGE_MIGRATION_VERSION}-${product.toLowerCase()}-renormalize-${await sha256(stableJson(records.filter((record) => record.product === product)))}`,
      migrated_at: new Date().toISOString(),
      migration_version: LOCAL_STORAGE_MIGRATION_VERSION,
      product,
      row_count: report.products[product].migratedRows,
      source_hash: await sha256(`${product}:legacy-storage-renormalize:${report.products[product].migratedRows}`),
      source_key: `${product.toLowerCase()}-legacy-storage-renormalize`,
      status: report.products[product].status === "error" ? "error" : "success",
    })),
  );
  const logResult = await upsertRows("migration_logs", logRows);
  if (!logResult.ok) report.errors.push(`Gagal menyimpan migration_logs renormalize: ${logResult.error ?? logResult.status}`);

  report.success = report.errors.length === 0;
  return report;
}
