"use client";

import type { Dataset, Row } from "@/features/metabright/domain/types";
import { avg, nf, pillStyle } from "@/features/metabright/domain/utils";

/* Peta capaian witel per indikator */
export function Heatmap({ rows, d }: { rows: Row[]; d: Dataset }) {
  const ws = [...new Set(rows.map((r) => r.witel))].sort();

  const cell = (rs: Row[], k: number) => {
    const w = d.indicators[k].weight;
    if (!w || !rs.length) return <td key={k}>-</td>;
    const v = avg(rs.map((r) => (r.v[k][3] / w) * 100));
    return (
      <td key={k} title={nf(v, 1) + "%"}>
        <span className="pill" style={pillStyle(v)}>{nf(v, 0)}%</span>
      </td>
    );
  };

  const line = (rs: Row[], name: string, sm: string, cls?: string) => (
    <tr key={name} className={cls}>
      <td><b>{name}</b><small>{sm}</small></td>
      {d.indicators.map((_, k) => cell(rs, k))}
      <td><b>{nf(avg(rs.map((r) => r.total)))}</b></td>
    </tr>
  );

  return (
    <table className="hm">
      <thead>
        <tr>
          <th>Witel</th>
          {d.indicators.map((x, k) => (
            <th key={k} title={x.name}>{x.name}<small>bobot {x.weight}</small></th>
          ))}
          <th>Rata-rata total score</th>
        </tr>
      </thead>
      <tbody>
        {ws.map((w) => {
          const rs = rows.filter((r) => r.witel === w);
          return line(rs, w, rs.length + " telda");
        })}
        {ws.length > 1 && line(rows, "Semua witel", rows.length + " telda", "all")}
      </tbody>
    </table>
  );
}

/* Tabel data skor per indikator */
export function DataTable({ rows, d }: { rows: Row[]; d: Dataset }) {
  const sorted = [...rows].sort((a, b) => a.rank - b.rank);
  return (
    <>
      <div className="sub" style={{ marginBottom: 12 }}>
        {rows.length} telda. Warna sel menunjukkan skor terhadap maksimum indikator:{" "}
        <span className="lgc" style={pillStyle(90)}>≥ 80%</span>
        <span className="lgc" style={pillStyle(60)}>50–79%</span>
        <span className="lgc" style={pillStyle(10)}>&lt; 50%</span>
      </div>
      <div className="tw">
        <table>
          <thead>
            <tr>
              <th>Rank</th><th>No</th><th>Witel</th><th>Telda</th>
              {d.indicators.map((x, k) => (
                <th key={k} title={x.name}>{x.name} ({x.weight})</th>
              ))}
              <th>Total score</th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((r) => (
              <tr key={r.telda + r.no}>
                <td><span className={"rk r" + (r.rank <= 3 ? r.rank : "x")}>{r.rank}</span></td>
                <td>{r.no}</td>
                <td>{r.witel}</td>
                <td><b>{r.telda}</b></td>
                {r.v.map((v, k) => {
                  const wt = d.indicators[k].weight || 1;
                  const p = Math.min(v[3] / (wt * 1.1), 1) * 100;
                  return (
                    <td key={k} title={"ACH " + nf(v[2]) + "%"}>
                      <span className="pill" style={pillStyle(p)}>{nf(v[3])}</span>
                    </td>
                  );
                })}
                <td>
                  <div className="ts">
                    <i style={{ ["--w" as string]: Math.min((r.total / 110) * 100, 100) + "%" }} />
                    <b>{nf(r.total)}</b>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  );
}
