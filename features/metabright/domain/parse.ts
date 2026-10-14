import * as XLSX from "xlsx";
import type { Dataset, Row, Indicator } from "./types";
import { num } from "./utils";

type Grid = unknown[][];

export function parseSheet(rows: Grid): Dataset {
  const h = rows.findIndex(
    (r) => r && r.some((c) => String(c).trim().toUpperCase() === "TELDA")
  );
  if (h < 2) throw new Error('baris judul "TELDA" tidak ditemukan');

  const H = rows[h];
  const names = rows[h - 2];
  const w = rows[h - 1];
  const cols: number[] = [];
  H.forEach((x, c) => {
    if (String(x).trim().toUpperCase() === "TARGET") cols.push(c);
  });

  const indicators: Indicator[] = cols.map((c, idx) => {
    let nm = "";
    for (let k = c; k >= c - 3 && k >= 0; k--) {
      if (names[k]) {
        nm = String(names[k]).trim();
        break;
      }
    }
    return { name: nm || "Indikator " + (idx + 1), weight: num(w[c]) || num(w[c + 1]) };
  });

  const find = (re: RegExp) => {
    for (const k of [h, h - 1, h - 2]) {
      const j = (rows[k] || []).findIndex((x) => re.test(String(x).trim()));
      if (j >= 0) return j;
    }
    return -1;
  };
  const tc = find(/^(total\s*score|skor\s*total|total\s*skor|total)$/i);
  const rc = find(/^(rank|ranking|peringkat)$/i);

  const out: Row[] = [];
  let warn = "";
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[2]) continue;
    out.push({
      no: num(r[0]),
      witel: String(r[1]).trim(),
      telda: String(r[2]).trim(),
      v: cols.map((c) => [num(r[c]), num(r[c + 1]), num(r[c + 2]), num(r[c + 3])]),
      total: tc >= 0 ? num(r[tc]) : 0,
      rank: rc >= 0 ? num(r[rc]) || 0 : 0,
    });
  }
  if (!out.length) throw new Error("tidak ada baris data");

  if (out.every((o) => !o.total)) {
    out.forEach((o) => (o.total = +o.v.reduce((a, v) => a + v[3], 0).toFixed(2)));
    warn =
      "kolom Total Score tidak terbaca di file ini, jadi total dihitung dari jumlah skor semua indikator. Periksa apakah angkanya sesuai.";
  }
  out.sort((a, b) => b.total - a.total);
  out.forEach((o, i) => {
    if (!o.rank) o.rank = i + 1;
  });
  return { indicators, rows: out, warn };
}

export async function parseFile(file: File): Promise<Dataset> {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, raw: false, defval: "" });
  return parseSheet(grid as Grid);
}
