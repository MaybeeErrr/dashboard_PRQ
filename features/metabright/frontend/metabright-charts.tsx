"use client";

import {
  Bar, BarChart, CartesianGrid, Cell, ComposedChart, LabelList, Legend, Line,
  PolarAngleAxis, PolarGrid, PolarRadiusAxis, Radar, RadarChart, ResponsiveContainer,
  Tooltip, XAxis, YAxis,
} from "recharts";
import type { Dataset, Row } from "@/features/metabright/domain/types";
import { COL, avg, fmt, nf } from "@/features/metabright/domain/utils";

const GRID = "#eef2f7";
const AXIS = { fontSize: 11, fill: "#475569" };
const LBL = { fontSize: 11, fontWeight: 700, fill: "#1e293b" };
const tipStyle = { background: "#0f172a", border: 0, borderRadius: 8, color: "#fff", fontSize: 12 };
const tipItem = { color: "#fff" };
const tv = (value: unknown) => (typeof value === "number" ? value : Number(value ?? 0));

export const witelColor = (ws: string[], w: string) => COL[ws.indexOf(w) % COL.length];

/* 1. Peringkat total score */
export function RankChart({ rows, ws, topN }: { rows: Row[]; ws: string[]; topN: number }) {
  const top = [...rows].sort((a, b) => b.total - a.total).slice(0, topN || rows.length);
  const data = top.map((r, k) => ({ label: `${k + 1}. ${r.telda}`, telda: r.telda, witel: r.witel, total: r.total }));
  const h = top.length * (top.length <= 12 ? 44 : 28) + 54;
  return (
    <ResponsiveContainer width="100%" height={h}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 48, bottom: 4, left: 4 }}>
        <CartesianGrid horizontal={false} stroke={GRID} />
        <XAxis type="number" domain={[0, "auto"]} tick={AXIS} />
        <YAxis type="category" dataKey="label" width={170} interval={0} tick={{ ...AXIS, fontWeight: 600, fill: "#334155" }} axisLine={false} tickLine={false} />
        <Tooltip
          cursor={{ fill: "rgba(148,163,184,.12)" }}
          contentStyle={tipStyle}
          itemStyle={tipItem}
          labelStyle={{ color: "#fff", fontWeight: 700 }}
          formatter={(v: unknown, _n: unknown, p: unknown) => [
            nf(tv(v)),
            (p as { payload?: { witel?: string } }).payload?.witel ?? "Total score",
          ]}
        />
        <Bar dataKey="total" radius={[0, 8, 8, 0]} maxBarSize={top.length <= 12 ? 30 : 20} isAnimationActive={false}>
          {data.map((d, i) => <Cell key={i} fill={witelColor(ws, d.witel)} />)}
          <LabelList dataKey="total" position="right" formatter={(v: unknown) => nf(Number(v))} style={LBL} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

/* 2. Rata-rata total score per witel */
export function WitelChart({ rows, ws }: { rows: Row[]; ws: string[] }) {
  const data = ws
    .map((w) => {
      const rs = rows.filter((r) => r.witel === w);
      return { witel: w, label: `${w} (${rs.length})`, v: +avg(rs.map((r) => r.total)).toFixed(2), n: rs.length };
    })
    .filter((x) => x.n)
    .sort((a, b) => b.v - a.v);
  return (
    <ResponsiveContainer width="100%" height={data.length * 34 + 40}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 48, bottom: 4, left: 4 }}>
        <CartesianGrid horizontal={false} stroke={GRID} />
        <XAxis type="number" domain={[0, "auto"]} tick={AXIS} />
        <YAxis type="category" dataKey="label" width={170} interval={0} tick={AXIS} axisLine={false} tickLine={false} />
        <Tooltip cursor={{ fill: "rgba(148,163,184,.12)" }} contentStyle={tipStyle} itemStyle={tipItem} formatter={(v: unknown) => [nf(tv(v)), "Rata-rata"]} />
        <Bar dataKey="v" radius={[0, 6, 6, 0]} isAnimationActive={false}>
          {data.map((d, i) => <Cell key={i} fill={witelColor(ws, d.witel)} />)}
          <LabelList dataKey="v" position="right" formatter={(v: unknown) => nf(Number(v))} style={LBL} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

/* helper: capaian rata-rata per indikator (% dari bobot) */
export const indicatorAvg = (rows: Row[], d: Dataset) =>
  d.indicators.map((x, k) => (x.weight ? avg(rows.map((r) => (r.v[k][3] / x.weight) * 100)) : 0));

/* 3. Capaian rata-rata per indikator */
export function IndicatorChart({ rows, d }: { rows: Row[]; d: Dataset }) {
  const iv = indicatorAvg(rows, d);
  const data = d.indicators.map((x, k) => ({ name: x.name, v: +iv[k].toFixed(1) }));
  return (
    <ResponsiveContainer width="100%" height={data.length * 26 + 40}>
      <BarChart data={data} layout="vertical" margin={{ top: 4, right: 56, bottom: 4, left: 4 }}>
        <CartesianGrid horizontal={false} stroke={GRID} />
        <XAxis type="number" domain={[0, 110]} tickFormatter={(v) => v + "%"} tick={AXIS} />
        <YAxis type="category" dataKey="name" width={200} interval={0} tick={AXIS} axisLine={false} tickLine={false} />
        <Tooltip cursor={{ fill: "rgba(148,163,184,.12)" }} contentStyle={tipStyle} itemStyle={tipItem} formatter={(v: unknown) => [nf(tv(v)) + "% dari bobot", "Capaian"]} />
        <Bar dataKey="v" radius={[0, 6, 6, 0]} isAnimationActive={false}>
          {data.map((d2, i) => <Cell key={i} fill={d2.v >= 80 ? "#059669" : d2.v >= 50 ? "#f59e0b" : "#e11d48"} />)}
          <LabelList dataKey="v" position="right" formatter={(v: unknown) => nf(Number(v), 1) + "%"} style={LBL} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

/* 4. Profil capaian telda (radar) */
export function RadarProfile({ t, rows, d }: { t?: Row; rows: Row[]; d: Dataset }) {
  const iv = indicatorAvg(rows, d);
  const short = (s: string) => (s.length > 16 ? s.slice(0, 15) + "…" : s);
  const data = d.indicators.map((x, k) => ({
    name: short(x.name),
    full: x.name,
    telda: t ? +((t.v[k][3] / (x.weight || 1)) * 100).toFixed(1) : 0,
    rata: +iv[k].toFixed(1),
  }));
  return (
    <ResponsiveContainer width="100%" height="100%">
      <RadarChart data={data} outerRadius="68%">
        <PolarGrid stroke="#e2e8f0" />
        <PolarAngleAxis dataKey="name" tick={{ fontSize: 10, fill: "#475569" }} />
        <PolarRadiusAxis angle={90} domain={[0, 110]} tickCount={5} tickFormatter={(v) => v + "%"} tick={{ fontSize: 10, fill: "#94a3b8" }} />
        <Radar name={t?.telda ?? "-"} dataKey="telda" stroke="#2563eb" fill="#2563eb" fillOpacity={0.22} strokeWidth={2.5} isAnimationActive={false} />
        <Radar name="Rata-rata filter" dataKey="rata" stroke="#f59e0b" fill="#f59e0b" fillOpacity={0.08} strokeWidth={2} strokeDasharray="5 4" isAnimationActive={false} />
        <Legend verticalAlign="bottom" iconType="circle" iconSize={8} />
        <Tooltip contentStyle={tipStyle} itemStyle={tipItem} labelStyle={{ color: "#fff", fontWeight: 700 }} formatter={(v: unknown) => nf(tv(v), 1) + "%"} />
      </RadarChart>
    </ResponsiveContainer>
  );
}

/* 5. Distribusi pencapaian */
const BUCKETS: [string, number, number, string][] = [
  ["< 50%", 0, 50, "#e11d48"],
  ["50–79%", 50, 80, "#f59e0b"],
  ["80–99%", 80, 100, "#38bdf8"],
  ["≥ 100%", 100, 1e9, "#059669"],
];

export function DistChart({ rows, i }: { rows: Row[]; i: number }) {
  const data = BUCKETS.map(([label, lo, hi, color]) => ({
    label, color, n: rows.filter((r) => r.v[i][2] >= lo && r.v[i][2] < hi).length,
  }));
  return (
    <ResponsiveContainer width="100%" height="100%">
      <BarChart data={data} margin={{ top: 26, right: 8, bottom: 4, left: 0 }}>
        <CartesianGrid vertical={false} stroke={GRID} />
        <XAxis dataKey="label" tick={AXIS} axisLine={false} tickLine={false} />
        <YAxis allowDecimals={false} tick={AXIS} label={{ value: "Jumlah telda", angle: -90, position: "insideLeft", style: { fontSize: 11, fill: "#475569" } }} />
        <Tooltip cursor={{ fill: "rgba(148,163,184,.12)" }} contentStyle={tipStyle} itemStyle={tipItem} formatter={(v: unknown) => [tv(v), "Telda"]} />
        <Bar dataKey="n" radius={[8, 8, 0, 0]} isAnimationActive={false}>
          {data.map((d, k) => <Cell key={k} fill={d.color} />)}
          <LabelList dataKey="n" position="top" style={LBL} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

/* 6. Target, realisasi, achievement */
export function TargetChart({ rows, i }: { rows: Row[]; i: number }) {
  const data = rows.map((r) => {
    const [tg, re] = r.v[i];
    return { telda: r.telda, Target: tg, Realisasi: re, Shortage: Math.max(tg - re, 0), ach: tg ? +((re / tg) * 100).toFixed(1) : 0 };
  });
  return (
    <ResponsiveContainer width="100%" height="100%">
      <ComposedChart data={data} margin={{ top: 8, right: 12, bottom: 4, left: 8 }}>
        <CartesianGrid vertical={false} stroke={GRID} />
        <XAxis dataKey="telda" interval={0} angle={-55} textAnchor="end" height={90} tick={AXIS} />
        <YAxis tickFormatter={(v) => fmt(v)} tick={AXIS} width={90} />
        <YAxis yAxisId="r" orientation="right" domain={[0, "auto"]} tickFormatter={(v) => v + "%"} tick={AXIS} />
        <Tooltip
          contentStyle={tipStyle}
          itemStyle={tipItem}
          labelStyle={{ color: "#fff", fontWeight: 700 }}
          formatter={(v: unknown, name: unknown) => (name === "Achievement (%)" ? nf(tv(v)) + "%" : fmt(tv(v)))}
        />
        <Legend verticalAlign="bottom" iconType="circle" iconSize={8} wrapperStyle={{ paddingTop: 8 }} />
        <Bar dataKey="Target" fill="#cbd5e1" radius={[6, 6, 0, 0]} isAnimationActive={false} />
        <Bar dataKey="Realisasi" fill="#2563eb" radius={[6, 6, 0, 0]} isAnimationActive={false} />
        <Bar dataKey="Shortage" fill="#fb7185" radius={[6, 6, 0, 0]} isAnimationActive={false} />
        <Line yAxisId="r" type="monotone" dataKey="ach" name="Achievement (%)" stroke="#059669" strokeWidth={2.5} dot={{ r: 3, fill: "#fff", stroke: "#059669" }} isAnimationActive={false} />
      </ComposedChart>
    </ResponsiveContainer>
  );
}
