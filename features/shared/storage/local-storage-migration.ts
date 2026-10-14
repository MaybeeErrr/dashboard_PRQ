"use client";

import {
  legacyStorageDefinitionForKey,
  legacyStorageKeys,
  LOCAL_STORAGE_MIGRATION_VERSION,
  type LegacyStorageArea,
  type LegacyStorageClassification,
  type LegacyStorageProduct,
} from "@/features/shared/storage/legacy-storage-keys";

export type LocalStorageBackupEntry = {
  area: LegacyStorageArea;
  classification: LegacyStorageClassification;
  key: string;
  product: LegacyStorageProduct;
  proposedTable: string;
  value: string;
};

export type LocalStorageBackup = {
  createdAt: string;
  keys: LocalStorageBackupEntry[];
  migrationVersion: number;
  source: "legacy-localStorage";
};

export type LocalStorageMigrationReport = {
  backupId: string | null;
  conflicts: string[];
  errors: string[];
  migratedKeys: string[];
  migrationVersion: number;
  products: Record<"HSI" | "WMS" | "LTS" | "METABRIGHT", {
    conflicts: string[];
    localRows: number;
    migratedRows: number;
    status: "success" | "skipped" | "conflict" | "error";
    verifiedRows: number;
  }>;
  skippedKeys: string[];
  success: boolean;
};

function storageEntries(storage: Storage, area: LegacyStorageArea): LocalStorageBackupEntry[] {
  return Array.from({ length: storage.length }, (_, index) => storage.key(index))
    .filter((key): key is string => Boolean(key))
    .sort((left, right) => left.localeCompare(right))
    .map((key) => {
      const definition = legacyStorageDefinitionForKey(key, area);
      return {
        area,
        classification: definition?.classification ?? "TEMPORARY_SESSION",
        key,
        product: definition?.product ?? "PORTAL",
        proposedTable: definition?.proposedTable ?? "not migrated",
        value: storage.getItem(key) ?? "",
      };
    });
}

export function scanLegacyLocalStorage(): LocalStorageBackupEntry[] {
  if (typeof window === "undefined") return [];
  const knownKeys = new Set(legacyStorageKeys.map((definition) => `${definition.area}.${definition.key}`));
  const entries = [...storageEntries(window.localStorage, "localStorage"), ...storageEntries(window.sessionStorage, "sessionStorage")];

  for (const definition of legacyStorageKeys) {
    const storage = definition.area === "localStorage" ? window.localStorage : window.sessionStorage;
    if (knownKeys.has(`${definition.area}.${definition.key}`) && storage.getItem(definition.key) === null) {
      continue;
    }
  }

  return entries;
}

export function createLocalStorageBackup(): LocalStorageBackup {
  return {
    createdAt: new Date().toISOString(),
    keys: scanLegacyLocalStorage(),
    migrationVersion: LOCAL_STORAGE_MIGRATION_VERSION,
    source: "legacy-localStorage",
  };
}

export function exportLocalStorageBackup() {
  const backup = createLocalStorageBackup();
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `prime-local-storage-backup-${backup.createdAt.replace(/[:.]/g, "-")}.json`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
  return backup;
}

export async function migrateLegacyDataToSupabase(backup = createLocalStorageBackup()): Promise<LocalStorageMigrationReport> {
  const response = await fetch("/api/migrations/legacy-storage", {
    body: JSON.stringify({
      createdAt: backup.createdAt,
      keys: backup.keys.map(({ area, key, value }) => ({ area, key, value })),
      migrationVersion: backup.migrationVersion,
      source: backup.source,
    }),
    headers: { "Content-Type": "application/json" },
    method: "POST",
  });

  const payload = (await response.json().catch(() => null)) as LocalStorageMigrationReport | { error?: string } | null;
  if (!response.ok && response.status !== 207) {
    throw new Error(payload && "error" in payload && payload.error ? payload.error : "Migrasi Supabase gagal.");
  }
  return payload as LocalStorageMigrationReport;
}

export async function normalizeExistingLegacyEntriesToSupabase(): Promise<LocalStorageMigrationReport> {
  const response = await fetch("/api/migrations/legacy-storage", { method: "PUT" });
  const payload = (await response.json().catch(() => null)) as LocalStorageMigrationReport | { error?: string } | null;
  if (!response.ok && response.status !== 207) {
    throw new Error(payload && "error" in payload && payload.error ? payload.error : "Normalisasi HSI/WMS dari legacy entries gagal.");
  }
  return payload as LocalStorageMigrationReport;
}
