export const MN = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember",
];

export const COL = [
  "#2563eb", "#0ea5e9", "#14b8a6", "#f59e0b", "#8b5cf6",
  "#e11d48", "#64748b", "#84cc16", "#ec4899", "#f97316",
];

export const nf = (n: number, d = 2) =>
  Number(n).toLocaleString("id-ID", { maximumFractionDigits: d });

export const fmt = (n: number) => {
  const a = Math.abs(n);
  return a >= 1e9 ? nf(n / 1e9) + " miliar" : a >= 1e6 ? nf(n / 1e6) + " juta" : nf(n);
};

export const num = (v: unknown): number => {
  if (v == null || v === "") return 0;
  if (typeof v === "number") return v;
  const x = parseFloat(String(v).replace(/%/g, "").replace(/,/g, ""));
  return isNaN(x) ? 0 : x;
};

export const avg = (a: number[]) => (a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0);

export const periodLabel = (n: string) => {
  const m = String(n).match(/(20\d{2})(0[1-9]|1[0-2])/);
  return m ? MN[+m[2] - 1] + " " + m[1] : String(n).replace(/\.[^.]+$/, "");
};

/** Indikator yang dihitung sebagai rata-rata (bukan penjumlahan) antar telda. */
export const isRatio = (name: string) => /collection|loss to sales|visiting/i.test(name);

export const slug = (x: string) => String(x).replace(/[^\w\-]+/g, "_").replace(/^_+|_+$/g, "");

export const wait = (t: number) => new Promise<void>((r) => setTimeout(r, t));

/** Warna pill berdasarkan persentase (>=80 hijau, >=50 kuning, selain itu merah). */
export const pillStyle = (p: number): { background: string; color: string } =>
  p >= 80
    ? { background: "#d1fae5", color: "#047857" }
    : p >= 50
    ? { background: "#fef3c7", color: "#b45309" }
    : { background: "#ffe4e6", color: "#be123c" };
