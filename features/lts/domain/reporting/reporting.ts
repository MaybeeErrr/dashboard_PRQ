export type LtsDatasetKey = "sales" | "caps" | "ct0";
export type LtsReportTabKey = "all" | "wms" | "hsi";
export type LtsScope = "witel" | "hotda";
export type LtsCell = string | number | boolean | null;
export type LtsMasterRow = LtsCell[];

export type LtsUploadedDataset = {
  datasetKey: LtsDatasetKey;
  fileName: string;
  headers: string[];
  periodKey: string;
  rows: LtsMasterRow[];
  rowCount: number;
  uploadedAt: string;
};

export type LtsReportRow = {
  label: string;
  coordinator?: string;
  sales: number;
  ct0: number;
  ct0Ratio: number;
  ct0Rank: number | null;
  caps: number;
  capsRatio: number;
  capsRank: number | null;
  score: number;
  scoreRank: number | null;
  capsRankOnly: number | null;
  churn: number;
  gap: number;
};

export type LtsReportTable = {
  scope: LtsScope;
  title: string;
  subtitle: string;
  rows: LtsReportRow[];
  total: LtsReportRow;
};

export type LtsComputedReport = {
  tabKey: LtsReportTabKey;
  datasetKey: LtsDatasetKey;
  label: string;
  sourceLabel: string;
  periodLabel: string;
  cutoffLabel: string;
  witel: LtsReportTable;
  hotda: LtsReportTable;
  summary: {
    sales: number;
    ct0: number;
    caps: number;
    score: number;
    bestWitel: LtsReportRow | null;
    bestHotda: LtsReportRow | null;
  };
};

type DatasetDefinition = {
  datasetKey: LtsDatasetKey;
  sheetName: string;
  label: string;
  shortLabel: string;
  reportTab: LtsReportTabKey;
  reportLabel: string;
  purpose: string;
  outputs: string[];
  requiredColumns: string[];
};

export const ltsDatasetDefinitions: Record<LtsDatasetKey, DatasetDefinition> = {
  sales: {
    datasetKey: "sales",
    sheetName: "SALES_MASTER",
    label: "SALES MASTER",
    shortLabel: "LTS ALL",
    reportTab: "all",
    reportLabel: "LTS ALL",
    purpose: "Source untuk report LTS all.",
    outputs: ["LTS WITEL (all)", "LTS HOTDA (all)"],
    requiredColumns: ["WITEL", "TELKOM_DAERAH", "Jml_Sales"],
  },
  caps: {
    datasetKey: "caps",
    sheetName: "CAPS_MASTER",
    label: "CAPS MASTER",
    shortLabel: "LTS WMS ONLY",
    reportTab: "wms",
    reportLabel: "LTS WMS ONLY",
    purpose: "Source untuk report LTS WMS only.",
    outputs: ["LTS WITEL (WMS only)", "LTS HOTDA (WMS only)"],
    requiredColumns: ["L_WITEL", "TELKOM_DAERAH", "COUNT_SID", "SSL_CT0", "SSL_DO"],
  },
  ct0: {
    datasetKey: "ct0",
    sheetName: "CT0_MASTER",
    label: "CT0 MASTER",
    shortLabel: "LTS HSI ONLY",
    reportTab: "hsi",
    reportLabel: "LTS HSI ONLY",
    purpose: "Source untuk report LTS HSI only.",
    outputs: ["LTS WITEL (HSI only)", "LTS HOTDA (HSI only)"],
    requiredColumns: ["Produk", "Bayar Desc", "Datel", "Witel"],
  },
};

export const ltsReportTabs: { key: LtsReportTabKey; label: string; datasetKey: LtsDatasetKey }[] = [
  { key: "all", label: "LTS ALL", datasetKey: "sales" },
  { key: "wms", label: "LTS WMS ONLY", datasetKey: "caps" },
  { key: "hsi", label: "LTS HSI ONLY", datasetKey: "ct0" },
];

const WITEL_ORDER = [
  "BALI",
  "JATIM BARAT",
  "JATIM TIMUR",
  "NUSA TENGGARA",
  "SEMARANG JATENG UTARA",
  "SOLO JATENG TIMUR",
  "SURAMADU",
  "YOGYA JATENG SELATAN",
];

const HOTDA_GROUPS: { label: string; coordinator: string; members: string[] }[] = [
  { label: "BANTUL", coordinator: "MURTADO", members: ["BANTUL", "GUNUNG KIDUL"] },
  { label: "CILACAP", coordinator: "P. CH. DWI PRASETYO M.S.", members: ["CILACAP"] },
  { label: "KEBUMEN", coordinator: "ODHY ANUGRAH AGUSTA", members: ["KEBUMEN", "PURWOREJO"] },
  { label: "MAGELANG", coordinator: "TRINIL ANJAR LESTARI, ST", members: ["MAGELANG", "MUNTILAN"] },
  { label: "PURBALINGGA", coordinator: "IVAN FERNALDY", members: ["PURBALINGGA", "BANJARNEGARA"] },
  { label: "PURWOKERTO", coordinator: "YOGI BAHTIAR, ST", members: ["PURWOKERTO"] },
  { label: "WONOSOBO", coordinator: "YOGANDA RAGIL PRANATA", members: ["WONOSOBO", "TEMANGGUNG"] },
  { label: "YOGYAKARTA", coordinator: "HERU SETIAWAN", members: ["YOGYAKARTA", "INNER - YOGYA JATENG SELATAN", "SLEMAN"] },
];

// Filter produk disamakan dengan rumus xlsx (sheet "Untuk manggil data (3 versi)"):
// - SALES  : kolom PRODUK, WMS = {WMS, WMS FIT, WMS LITE, WMS RESELLER}, HSI = {HSI BASIC/BISNIS (+MTC), HSI RESELLER}
// - CAPS   : kolom L_PRODUK, WMS = "WMS" saja, HSI = {HSI BASIC, HSI BISNIS, HSI RESELLER}
// - CT0    : kolom Produk, WMS = "Wifi Manage Service", HSI = "Internet" (persis sama)
const WMS_SALES_PRODUCTS = new Set(["WMS", "WMS FIT", "WMS LITE", "WMS RESELLER"]);
const HSI_SALES_PRODUCTS = new Set(["HSI BASIC", "HSI BASIC MTC", "HSI BISNIS", "HSI BISNIS MTC", "HSI RESELLER"]);
const WMS_CAPS_PRODUCTS = new Set(["WMS"]);
const HSI_CAPS_PRODUCTS = new Set(["HSI BASIC", "HSI BISNIS", "HSI RESELLER"]);
const WMS_CT0_PRODUCTS = new Set(["WIFI MANAGE SERVICE"]);
const HSI_CT0_PRODUCTS = new Set(["INTERNET"]);

const WITEL_CHURN: Record<string, number> = {
  BALI: 51,
  "JATIM BARAT": 123,
  "JATIM TIMUR": 91,
  "NUSA TENGGARA": 108,
  "SEMARANG JATENG UTARA": 126,
  "SOLO JATENG TIMUR": 164,
  SURAMADU: 231,
  "YOGYA JATENG SELATAN": 96,
};

const HOTDA_CHURN: Record<string, number> = {
  BANTUL: 0,
  CILACAP: 2,
  KEBUMEN: 3,
  MAGELANG: 6,
  PURBALINGGA: 4,
  PURWOKERTO: 8,
  WONOSOBO: 2,
  YOGYAKARTA: 10,
};

function normalizeHeader(value: string) {
  return value.trim().toUpperCase().replace(/\s+/g, " ");
}

function normalizeLabel(value: unknown) {
  return String(value ?? "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function stripTerritoryCode(value: unknown) {
  return normalizeLabel(value)
    .replace(/^\d+\s*-\s*/, "")
    .replace(/^YK JATENG SELATAN$/, "YOGYA JATENG SELATAN")
    .replace(/^YOGYAKARTA JATENG SELATAN$/, "YOGYA JATENG SELATAN")
    .replace(/PURWOREJJO/g, "PURWOREJO");
}

function headerMap(headers: string[]) {
  const map = new Map<string, number>();
  headers.forEach((header, index) => {
    const normalized = normalizeHeader(header);
    if (normalized && !map.has(normalized)) map.set(normalized, index);
  });
  return map;
}

function findColumn(headers: string[], candidates: string[]) {
  const map = headerMap(headers);
  for (const candidate of candidates) {
    const index = map.get(normalizeHeader(candidate));
    if (index !== undefined) return index;
  }
  return -1;
}

function getCell(row: LtsMasterRow, headers: string[], candidates: string[]) {
  const index = findColumn(headers, candidates);
  return index >= 0 ? row[index] : null;
}

function toNumber(value: unknown, fallback = 0) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value === "boolean") return value ? 1 : 0;
  const raw = String(value ?? "").trim();
  // Angka biasa ("1", "1.0", "-2.5") dibaca apa adanya; sebelumnya "1.0" salah terbaca 10.
  if (/^-?\d+(\.\d+)?$/.test(raw)) return Number(raw);
  const text = raw.replace(/\./g, "").replace(",", ".").trim();
  if (!text) return fallback;
  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function rowValue(row: LtsMasterRow, headers: string[], candidates: string[], fallback = 1) {
  const value = getCell(row, headers, candidates);
  const parsed = toNumber(value, Number.NaN);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function hasAnyColumn(headers: string[], candidates: string[]) {
  return findColumn(headers, candidates) >= 0;
}

// xlsx: COUNTIFS(..., "Belum Bayar") -> kecocokan penuh, bukan "mengandung".
function isBelumBayar(value: unknown) {
  return normalizeLabel(value) === "BELUM BAYAR";
}

function productMatchesMode(tabKey: LtsReportTabKey, source: LtsDatasetKey, product: unknown, hasProductColumn: boolean) {
  if (tabKey === "all") return true;
  const normalized = normalizeLabel(product);
  if (!hasProductColumn && source === "caps" && tabKey === "wms") return true;
  if (source === "caps") return (tabKey === "wms" ? WMS_CAPS_PRODUCTS : HSI_CAPS_PRODUCTS).has(normalized);
  if (source === "ct0") return (tabKey === "wms" ? WMS_CT0_PRODUCTS : HSI_CT0_PRODUCTS).has(normalized);
  return (tabKey === "wms" ? WMS_SALES_PRODUCTS : HSI_SALES_PRODUCTS).has(normalized);
}

function normalizeWitel(value: unknown) {
  const normalized = stripTerritoryCode(value);
  if (normalized === "YK JATENG SELATAN" || normalized === "YOGYAKARTA JATENG SELATAN") return "YOGYA JATENG SELATAN";
  return WITEL_ORDER.includes(normalized) ? normalized : "";
}

function normalizeHotda(value: unknown) {
  const normalized = stripTerritoryCode(value);
  if (!normalized) return "";
  for (const group of HOTDA_GROUPS) {
    if (group.members.includes(normalized) || normalized === group.label) return group.label;
  }
  return "";
}

function periodColumnsForDataset(datasetKey: LtsDatasetKey) {
  // Kolom angka YYYYMM didahulukan; nama bulan (NPER) hanya cadangan.
  return datasetKey === "sales"
    ? ["PERIOD", "PERIOD_U", "NPER", "LAST_UPDATED_DATE"]
    : datasetKey === "caps"
      ? ["PERIODE", "PERIODE_CT0_BAYAR"]
      : ["Nper", "TAHUN", "LAST_PERIODE"];
}

export function periodLabelFromLtsPeriodKey(periodKey: string) {
  const match = periodKey.match(/^(\d{4})(\d{2})$/);
  if (!match) return periodKey;
  const monthIndex = Number(match[2]) - 1;
  if (monthIndex < 0 || monthIndex > 11) return periodKey;
  return `${MONTHS[monthIndex]} ${match[1]}`;
}

function periodKey(value: unknown) {
  if (value instanceof Date) {
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${value.getFullYear()}${month}`;
  }
  const text = String(value ?? "").trim();
  if (!text) return "";
  const sixDigit = text.match(/^(\d{4})(\d{2})/);
  if (sixDigit) return `${sixDigit[1]}${sixDigit[2]}`;
  const isoDate = text.match(/^(\d{4})-(\d{2})-\d{2}/);
  if (isoDate) return `${isoDate[1]}${isoDate[2]}`;
  // Nama bulan bahasa Indonesia ATAU Inggris ("October" sebelumnya tidak dikenali -> periode salah).
  const upper = text.toUpperCase();
  const monthIndex = MONTHS.findIndex((month) => month === upper);
  const englishIndex = CUTOFF_MONTHS.findIndex((month) => month.toUpperCase() === upper);
  const index = monthIndex >= 0 ? monthIndex : englishIndex;
  if (index >= 0) return `${new Date().getFullYear()}${String(index + 1).padStart(2, "0")}`;
  return "";
}

export function periodKeyFromLtsRows(datasetKey: LtsDatasetKey, headers: string[], rows: LtsMasterRow[]) {
  const periodColumns = periodColumnsForDataset(datasetKey);
  const values = new Map<string, number>();
  for (const row of rows.slice(0, 500)) {
    // Coba setiap kolom periode berurutan sampai ada yang bisa dibaca.
    let key = "";
    for (const column of periodColumns) {
      key = periodKey(getCell(row, headers, [column]));
      if (key) break;
    }
    if (key) values.set(key, (values.get(key) ?? 0) + 1);
  }
  return [...values.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? "202609";
}

function periodFromRows(datasetKey: LtsDatasetKey, headers: string[], rows: LtsMasterRow[]) {
  return periodLabelFromLtsPeriodKey(periodKeyFromLtsRows(datasetKey, headers, rows));
}

const MONTHS = ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"];
const CUTOFF_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function cutoffLabelFromDatasets(datasets: LtsReportSourceDatasets) {
  const timestamps = Object.values(datasets)
    .map((dataset) => (dataset?.uploadedAt ? new Date(dataset.uploadedAt).getTime() : Number.NaN))
    .filter(Number.isFinite);
  if (timestamps.length === 0) return "Cut Off: -";
  const date = new Date(Math.max(...timestamps));
  return `Cut Off: ${String(date.getDate()).padStart(2, "0")} ${CUTOFF_MONTHS[date.getMonth()]} ${date.getFullYear()}`;
}

type Accumulator = {
  sales: number;
  ct0: number;
  caps: number;
  churn: number;
};

function blankAccumulator(): Accumulator {
  return { sales: 0, ct0: 0, caps: 0, churn: 0 };
}

function addMetric(target: Accumulator, metric: Accumulator) {
  target.sales += metric.sales;
  target.ct0 += metric.ct0;
  target.caps += metric.caps;
  target.churn += metric.churn;
}

function salesMetricForRow(tabKey: LtsReportTabKey, headers: string[], row: LtsMasterRow): Accumulator {
  const product = getCell(row, headers, ["PRODUK", "PRODUCT", "TYPE_LAYANAN", "L_PRODUK"]);
  const hasProductColumn = hasAnyColumn(headers, ["PRODUK", "PRODUCT", "TYPE_LAYANAN", "L_PRODUK"]);
  if (!productMatchesMode(tabKey, "sales", product, hasProductColumn)) return blankAccumulator();
  // xlsx: tab ALL = SUM(Jml_Sales); tab HSI/WMS = COUNTIFS (jumlah baris, 1 per baris).
  const sales = tabKey === "all" ? rowValue(row, headers, ["Jml_Sales"], 0) : 1;
  return { sales, ct0: 0, caps: 0, churn: 0 };
}

function ct0MetricForRow(tabKey: LtsReportTabKey, headers: string[], row: LtsMasterRow): Accumulator {
  if (!isBelumBayar(getCell(row, headers, ["Bayar Desc"]))) return blankAccumulator();
  const product = getCell(row, headers, ["Produk"]);
  const hasProductColumn = hasAnyColumn(headers, ["Produk"]);
  if (!productMatchesMode(tabKey, "ct0", product, hasProductColumn)) return blankAccumulator();
  // xlsx: COUNTIFS -> hitung jumlah baris (bukan menjumlahkan kolom JML LYN).
  return { sales: 0, ct0: 1, caps: 0, churn: 0 };
}

function capsMetricForRow(tabKey: LtsReportTabKey, headers: string[], row: LtsMasterRow): Accumulator {
  // xlsx: COUNTIFS(SSL_DO, 1) -> baris dengan SSL_DO tepat 1, dihitung per baris.
  if (rowValue(row, headers, ["SSL_DO"], 0) !== 1) return blankAccumulator();
  const product = getCell(row, headers, ["L_PRODUK", "PRODUK", "PRODUCT"]);
  const hasProductColumn = hasAnyColumn(headers, ["L_PRODUK", "PRODUK", "PRODUCT"]);
  if (!productMatchesMode(tabKey, "caps", product, hasProductColumn)) return blankAccumulator();
  return { sales: 0, ct0: 0, caps: 1, churn: 0 };
}

function scopeLabel(datasetKey: LtsDatasetKey, scope: LtsScope, headers: string[], row: LtsMasterRow) {
  if (scope === "witel") {
    if (datasetKey === "sales") return normalizeWitel(getCell(row, headers, ["WITEL"]));
    // xlsx CAPS memakai kolom "WITEL (NEW_TERRITORY)" (kolom AB), bukan L_WITEL.
    if (datasetKey === "caps") return normalizeWitel(getCell(row, headers, ["WITEL (NEW_TERRITORY)", "L_WITEL", "TWITEL"]));
    return normalizeWitel(getCell(row, headers, ["Witel", "Witel Bill", "OLD_WITEL"]));
  }
  if (datasetKey === "sales") return normalizeHotda(getCell(row, headers, ["TELKOM_DAERAH", "TELDA_NCX", "CTELDA"]));
  if (datasetKey === "caps") return normalizeHotda(getCell(row, headers, ["TELKOM_DAERAH"]));
  return normalizeHotda(getCell(row, headers, ["Datel"]));
}

// xlsx: SALES untuk WITEL "YOGYA JATENG SELATAN" hanya menghitung baris dengan TELKOM_DAERAH terisi (TELKOM_DAERAH <> "").
function isExcludedFromWitelSales(datasetKey: LtsDatasetKey, scope: LtsScope, label: string, headers: string[], row: LtsMasterRow) {
  if (datasetKey !== "sales" || scope !== "witel" || label !== "YOGYA JATENG SELATAN") return false;
  return String(getCell(row, headers, ["TELKOM_DAERAH"]) ?? "").trim() === "";
}

function rankRows(rows: LtsReportRow[], selector: (row: LtsReportRow) => number, setter: (row: LtsReportRow, rank: number) => void) {
  const values = rows.map(selector).filter((value) => Number.isFinite(value));
  rows.forEach((row) => {
    const value = selector(row);
    if (!Number.isFinite(value)) return;
    setter(row, values.filter((other) => other < value).length + 1);
  });
}

function safeRatio(numerator: number, denominator: number) {
  return denominator > 0 ? numerator / denominator : 0;
}

function calculateLtsScore(ct0Ratio: number, capsRatio: number) {
  return ct0Ratio * 0.5 + capsRatio * 0.5;
}

function churnForLabel(scope: LtsScope, label: string) {
  return scope === "witel" ? (WITEL_CHURN[label] ?? 0) : (HOTDA_CHURN[label] ?? 0);
}

export type LtsReportSourceDatasets = Partial<Record<LtsDatasetKey, LtsUploadedDataset>>;

function addDatasetRows(targets: Map<string, Accumulator>, scope: LtsScope, datasetKey: LtsDatasetKey, dataset: LtsUploadedDataset | undefined, tabKey: LtsReportTabKey) {
  if (!dataset) return;
  for (const row of dataset.rows) {
    const label = scopeLabel(datasetKey, scope, dataset.headers, row);
    const target = targets.get(label);
    if (!target) continue;
    if (isExcludedFromWitelSales(datasetKey, scope, label, dataset.headers, row)) continue;
    const metric =
      datasetKey === "sales"
        ? salesMetricForRow(tabKey, dataset.headers, row)
        : datasetKey === "ct0"
          ? ct0MetricForRow(tabKey, dataset.headers, row)
          : capsMetricForRow(tabKey, dataset.headers, row);
    addMetric(target, metric);
  }
}

function buildReportRows(scope: LtsScope, tabKey: LtsReportTabKey, datasets: LtsReportSourceDatasets) {
  const labels = scope === "witel" ? WITEL_ORDER : HOTDA_GROUPS.map((group) => group.label);
  const coordinators = new Map(HOTDA_GROUPS.map((group) => [group.label, group.coordinator]));
  const grouped = new Map(labels.map((label) => [label, { ...blankAccumulator(), churn: churnForLabel(scope, label) }]));

  addDatasetRows(grouped, scope, "sales", datasets.sales, tabKey);
  addDatasetRows(grouped, scope, "ct0", datasets.ct0, tabKey);
  addDatasetRows(grouped, scope, "caps", datasets.caps, tabKey);

  const reportRows: LtsReportRow[] = labels.map((label) => {
    const value = grouped.get(label) ?? blankAccumulator();
    const ct0Ratio = safeRatio(value.ct0, value.sales);
    const capsRatio = safeRatio(value.caps, value.sales);
    return {
      label,
      coordinator: coordinators.get(label),
      sales: value.sales,
      ct0: value.ct0,
      ct0Ratio,
      ct0Rank: null,
      caps: value.caps,
      capsRatio,
      capsRank: null,
      score: calculateLtsScore(ct0Ratio, capsRatio),
      scoreRank: null,
      capsRankOnly: null,
      churn: value.churn,
      gap: value.churn - value.caps,
    };
  });

  rankRows(reportRows, (row) => row.ct0Ratio, (row, rank) => (row.ct0Rank = rank));
  rankRows(reportRows, (row) => row.capsRatio, (row, rank) => {
    row.capsRank = rank;
    row.capsRankOnly = rank;
  });
  rankRows(reportRows, (row) => row.score, (row, rank) => (row.scoreRank = rank));

  const totalValue = reportRows.reduce(
    (acc, row) => ({
      sales: acc.sales + row.sales,
      ct0: acc.ct0 + row.ct0,
      caps: acc.caps + row.caps,
      churn: acc.churn + row.churn,
    }),
    blankAccumulator(),
  );
  const totalCt0Ratio = safeRatio(totalValue.ct0, totalValue.sales);
  const totalCapsRatio = safeRatio(totalValue.caps, totalValue.sales);
  const total: LtsReportRow = {
    label: scope === "witel" ? "TREG 3" : "YOGYAKARTA JATENG SELATAN",
    sales: totalValue.sales,
    ct0: totalValue.ct0,
    ct0Ratio: totalCt0Ratio,
    ct0Rank: null,
    caps: totalValue.caps,
    capsRatio: totalCapsRatio,
    capsRank: null,
    score: calculateLtsScore(totalCt0Ratio, totalCapsRatio),
    scoreRank: null,
    capsRankOnly: null,
    churn: totalValue.churn,
    gap: totalValue.churn - totalValue.caps,
  };

  return { rows: reportRows, total };
}

export function validateLtsDataset(datasetKey: LtsDatasetKey, headers: string[], rows: LtsMasterRow[]) {
  const definition = ltsDatasetDefinitions[datasetKey];
  const normalized = new Set(headers.map(normalizeHeader));
  const missing = definition.requiredColumns.filter((column) => !normalized.has(normalizeHeader(column)));
  if (missing.length > 0) {
    throw new Error(`${definition.label} is missing required column: ${missing[0]}`);
  }
  if (rows.length === 0) throw new Error(`${definition.label} tidak berisi data.`);
}

export function computeLtsReport(tabKey: LtsReportTabKey, datasets: LtsReportSourceDatasets): LtsComputedReport {
  const tab = ltsReportTabs.find((item) => item.key === tabKey);
  if (!tab) throw new Error("Tab LTS tidak valid.");
  const definition = ltsDatasetDefinitions[tab.datasetKey];
  const primaryDataset = datasets[definition.datasetKey];
  if (!primaryDataset) throw new Error(`${definition.reportLabel} membutuhkan ${definition.label}.`);

  const witelRows = buildReportRows("witel", tabKey, datasets);
  const hotdaRows = buildReportRows("hotda", tabKey, datasets);
  const period = periodFromRows(primaryDataset.datasetKey, primaryDataset.headers, primaryDataset.rows);
  const bestWitel = [...witelRows.rows].sort((a, b) => a.score - b.score)[0] ?? null;
  const bestHotda = [...hotdaRows.rows].sort((a, b) => a.score - b.score)[0] ?? null;

  return {
    tabKey,
    datasetKey: definition.datasetKey,
    label: definition.reportLabel,
    sourceLabel: "BRIGHT",
    periodLabel: period,
    cutoffLabel: cutoffLabelFromDatasets(datasets),
    witel: {
      scope: "witel",
      title: `LTS WITEL - ${definition.reportLabel.replace("LTS ", "")}`,
      subtitle: definition.datasetKey === "sales" ? "TREG 3" : definition.shortLabel,
      rows: witelRows.rows,
      total: witelRows.total,
    },
    hotda: {
      scope: "hotda",
      title: `LTS HOTDA - ${definition.reportLabel.replace("LTS ", "")}`,
      subtitle: "WITEL YOGYA JATENG SELATAN",
      rows: hotdaRows.rows,
      total: hotdaRows.total,
    },
    summary: {
      sales: witelRows.total.sales,
      ct0: witelRows.total.ct0,
      caps: witelRows.total.caps,
      score: witelRows.total.score,
      bestWitel,
      bestHotda,
    },
  };
}
