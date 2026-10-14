import * as XLSX from "xlsx";
import type { Dataset, Row } from "./types";

export function toAoa(rows: Row[], d: Dataset): (string | number)[][] {
  const h: (string | number)[] = ["Rank", "No", "Witel", "Telda"];
  d.indicators.forEach((x) =>
    h.push(x.name + " - Target", x.name + " - Realisasi", x.name + " - ACH (%)", x.name + " - Skor")
  );
  h.push("Total score");
  return [
    h,
    ...[...rows]
      .sort((a, b) => a.rank - b.rank)
      .map((r) => [r.rank, r.no, r.witel, r.telda, ...r.v.flat(), r.total]),
  ];
}

export function buildWorkbook(sets: [string, (string | number)[][]][]): Blob {
  const wb = XLSX.utils.book_new();
  const used: Record<string, number> = {};
  sets.forEach(([n, a]) => {
    const nm = n.replace(/[\\/?*[\]:]/g, " ").slice(0, 28);
    let k = nm;
    let c = 1;
    while (used[k]) k = nm + " " + ++c;
    used[k] = 1;
    const ws = XLSX.utils.aoa_to_sheet(a);
    ws["!cols"] = a[0].map((_, i) => ({
      wch: i === 0 || i === 1 ? 7 : i === 2 ? 22 : i === 3 ? 18 : i === a[0].length - 1 ? 12 : 17,
    }));
    XLSX.utils.book_append_sheet(wb, ws, k);
  });
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}

export function saveBlob(filename: string, blob: Blob) {
  const a = document.createElement("a");
  a.download = filename;
  a.href = URL.createObjectURL(blob);
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
