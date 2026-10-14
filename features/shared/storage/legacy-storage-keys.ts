export const LOCAL_STORAGE_MIGRATION_VERSION = 1;

export type LegacyStorageArea = "localStorage" | "sessionStorage";
export type LegacyStorageClassification = "BUSINESS_PERSISTENT" | "USER_UI_PREFERENCE" | "TEMPORARY_SESSION";
export type LegacyStorageProduct = "HSI" | "WMS" | "LTS" | "METABRIGHT" | "PORTAL";

export type LegacyStorageKeyDefinition = {
  area: LegacyStorageArea;
  classification: LegacyStorageClassification;
  description: string;
  key: string;
  product: LegacyStorageProduct;
  proposedTable: string;
};

export const legacyStorageKeys = [
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI active master rows for the selected reporting period.",
    key: "hsi-master-rows-v1",
    product: "HSI",
    proposedTable: "dashboard_state, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI active master data save timestamp.",
    key: "hsi-master-saved-at-v1",
    product: "HSI",
    proposedTable: "dashboard_state, legacy_storage_entries",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI imported report history and computed report snapshots.",
    key: "hsi-report-history-v1",
    product: "HSI",
    proposedTable: "dashboard_state, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI period archive keyed by reporting period.",
    key: "hsi-period-archive-v1",
    product: "HSI",
    proposedTable: "dashboard_state, reporting_imports, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "USER_UI_PREFERENCE",
    description: "HSI currently selected reporting period.",
    key: "hsi-selected-period-v1",
    product: "HSI",
    proposedTable: "dashboard_state",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI seed workbook version used to decide refresh behavior.",
    key: "hsi-seed-data-version-v1",
    product: "HSI",
    proposedTable: "dashboard_state, migration_logs",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "HSI saved snapshots per reporting period.",
    key: "hsi-save-snapshots-v1",
    product: "HSI",
    proposedTable: "dashboard_state, reporting_imports",
  },
  {
    area: "localStorage",
    classification: "USER_UI_PREFERENCE",
    description: "HSI master header cleanup flag.",
    key: "hsi-master-headers-cleared-v1",
    product: "HSI",
    proposedTable: "dashboard_state",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "WMS active master rows for the selected reporting period.",
    key: "wms-master-rows-v1",
    product: "WMS",
    proposedTable: "dashboard_state, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "WMS active master data save timestamp.",
    key: "wms-master-saved-at-v1",
    product: "WMS",
    proposedTable: "dashboard_state, legacy_storage_entries",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "WMS period archive keyed by reporting period.",
    key: "wms-period-archive-v1",
    product: "WMS",
    proposedTable: "dashboard_state, reporting_imports, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "WMS saved snapshots per reporting period.",
    key: "wms-save-snapshots-v1",
    product: "WMS",
    proposedTable: "dashboard_state, reporting_imports",
  },
  {
    area: "localStorage",
    classification: "USER_UI_PREFERENCE",
    description: "WMS currently selected reporting period.",
    key: "wms-selected-period-v1",
    product: "WMS",
    proposedTable: "dashboard_state",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "WMS seed workbook version used to decide refresh behavior.",
    key: "wms-seed-data-version-v1",
    product: "WMS",
    proposedTable: "dashboard_state, migration_logs",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "LTS full persisted dashboard state, including datasets and dataset archive.",
    key: "lts-dashboard-state-v1",
    product: "LTS",
    proposedTable: "dashboard_state, reporting_imports, dashboard_records",
  },
  {
    area: "localStorage",
    classification: "BUSINESS_PERSISTENT",
    description: "LTS seed workbook version used to decide refresh behavior.",
    key: "lts-seed-data-version-v1",
    product: "LTS",
    proposedTable: "dashboard_state, migration_logs",
  },
  {
    area: "sessionStorage",
    classification: "TEMPORARY_SESSION",
    description: "Per-tab login/session guard marker. This is intentionally not migrated.",
    key: "prime_session_active",
    product: "PORTAL",
    proposedTable: "not migrated",
  },
] satisfies LegacyStorageKeyDefinition[];

export function legacyStorageDefinitionForKey(key: string, area: LegacyStorageArea) {
  return legacyStorageKeys.find((definition) => definition.key === key && definition.area === area) ?? null;
}
