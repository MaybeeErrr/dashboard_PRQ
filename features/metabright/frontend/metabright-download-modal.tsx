"use client";

import { useMemo, useState } from "react";
import JSZip from "jszip";
import type { Dataset, Tab } from "@/features/metabright/domain/types";
import { buildWorkbook, saveBlob, toAoa } from "@/features/metabright/domain/export";
import { slug } from "@/features/metabright/domain/utils";
import { Icon, type IconName } from "./metabright-icons";

type Scope = "all" | "one" | "each";

interface Props {
  period: string;
  data: Dataset;
  capture: (witel: string, mode: Tab) => Promise<Blob>;
  begin: () => void;
  end: () => void;
  onClose: () => void;
  onMsg: (text: string, kind?: "okk" | "err") => void;
}

interface OutFile { n: string; z: string; b: Blob }

function ScopeOpt({ v, cur, set, busy, ic, t, s, children }: {
  v: Scope; cur: Scope; set: (v: Scope) => void; busy: boolean;
  ic: IconName; t: string; s: string; children?: React.ReactNode;
}) {
  return (
    <label className="sc">
      <input type="radio" name="sc" checked={cur === v} onChange={() => set(v)} disabled={busy} />
      <span className="sc-b">
        <span className="ic"><Icon name={ic} /></span>
        <span className="tx"><b>{t}</b><small>{s}</small>{children}</span>
        <span className="rd" />
      </span>
    </label>
  );
}

function Content({ id, ic, t, tag, s, on, set, busy }: {
  id: string; ic: IconName; t: string; tag: string; s: string;
  on: boolean; set: (v: boolean) => void; busy: boolean;
}) {
  return (
    <label className="ct">
      <input type="checkbox" id={id} checked={on} onChange={(e) => set(e.target.checked)} disabled={busy} />
      <span className="ct-b">
        <span className="ic"><Icon name={ic} /></span>
        <span className="tx"><b>{t}</b><small>{s}</small><span className="tag">{tag}</span></span>
        <span className="ck"><Icon name="check" sw={3.5} /></span>
      </span>
    </label>
  );
}

export default function DownloadModal({ period, data: d, capture, begin, end, onClose, onMsg }: Props) {
  const ws = useMemo(() => [...new Set(d.rows.map((r) => r.witel))].sort(), [d]);
  const [scope, setScope] = useState<Scope>("all");
  const [mw, setMw] = useState(ws[0] ?? "");
  const [kd, setKd] = useState(true);
  const [kt, setKt] = useState(false);
  const [kv, setKv] = useState(true);
  const [busy, setBusy] = useState(false);
  const [status, setStatus] = useState("");

  const n = scope === "each" ? ws.length : 1;
  const files = (kd ? 1 : 0) + n * ((kt ? 1 : 0) + (kv ? 1 : 0));
  const where =
    scope === "all" ? "semua witel digabung" : scope === "one" ? "witel " + mw : n + " witel dipisah per folder";
  const summary = !files
    ? "Pilih minimal satu isi unduhan."
    : (files === 1 ? "1 file akan diunduh" : files + " file akan dikemas dalam satu ZIP") + " — " + where + ".";

  async function run() {
    if (!files) return;
    const targets = scope === "all" ? ["ALL"] : scope === "one" ? [mw] : ws;
    const nm = (w: string) => (w === "ALL" ? "Semua_witel" : slug(w));
    const out: OutFile[] = [];
    setBusy(true);
    try {
      if (kd) {
        const sets: [string, (string | number)[][]][] = targets.map((w) => [
          w === "ALL" ? "Semua witel" : w,
          toAoa(w === "ALL" ? d.rows : d.rows.filter((r) => r.witel === w), d),
        ]);
        if (scope === "each") sets.unshift(["Semua witel", toAoa(d.rows, d)]);
        const tag = targets.length === 1 ? nm(targets[0]) : "per_witel";
        out.push({ n: `Data_${tag}_${slug(period)}.xlsx`, z: `Data/Data_${tag}.xlsx`, b: buildWorkbook(sets) });
      }
      const modes: Tab[] = [];
      if (kt) modes.push("data");
      if (kv) modes.push("viz");
      const total = targets.length * modes.length;
      if (total) {
        begin();
        let k = 0;
        for (const w of targets) {
          for (const m of modes) {
            setStatus(`Menyiapkan gambar ${++k} dari ${total}…`);
            const b = await capture(w, m);
            const t = m === "data" ? "Tabel" : "Visualisasi";
            out.push({ n: `${t}_${nm(w)}_${slug(period)}.png`, z: `${t}/${t}_${nm(w)}.png`, b });
          }
        }
      }
    } catch (e) {
      onMsg("Gagal menyiapkan unduhan: " + (e instanceof Error ? e.message : String(e)), "err");
      out.length = 0;
    } finally {
      end();
    }
    if (!out.length) {
      setBusy(false);
      onClose();
      return;
    }
    try {
      setStatus("Menyimpan…");
      if (out.length === 1) saveBlob(out[0].n, out[0].b);
      else {
        const z = new JSZip();
        out.forEach((f) => z.file(f.z, f.b));
        saveBlob(`Impactful_Award_Telda_${slug(period)}.zip`, await z.generateAsync({ type: "blob" }));
      }
      onMsg("Unduhan siap disimpan.", "okk");
    } catch (e) {
      onMsg("Gagal menyimpan: " + (e instanceof Error ? e.message : String(e)), "err");
    }
    setBusy(false);
    onClose();
  }

  return (
    <div className="ov" onClick={(e) => { if (e.target === e.currentTarget && !busy) onClose(); }}>
      <div className="md wide" role="dialog" aria-modal="true" aria-label="Unduh data dan visualisasi">
        <h3>Unduh data dan visualisasi</h3>
        <p className="dm-sub">Periode <b>{period}</b>. Pilih cakupan dan isi unduhan.</p>

        <div className="dm-t">Cakupan witel</div>
        <ScopeOpt v="all" cur={scope} set={setScope} busy={busy} ic="all" t="Semua witel" s="Satu file gabungan dari seluruh witel" />
        <ScopeOpt v="one" cur={scope} set={setScope} busy={busy} ic="pin" t="Satu witel" s="Hanya data dari witel yang dipilih">
          <select
            aria-label="Pilih witel"
            value={mw}
            disabled={busy}
            onFocus={() => setScope("one")}
            onChange={(e) => { setMw(e.target.value); setScope("one"); }}
          >
            {ws.map((w) => <option key={w}>{w}</option>)}
          </select>
        </ScopeOpt>
        <ScopeOpt v="each" cur={scope} set={setScope} busy={busy} ic="folder" t="Setiap witel dipisah" s="Satu folder per witel, dikemas dalam ZIP" />

        <div className="dm-t">Isi unduhan<small>boleh pilih lebih dari satu</small></div>
        <div className="cts">
          <Content busy={busy} id="kd" ic="data" t="Data" tag=".xlsx" s="Angka target, realisasi, dan skor" on={kd} set={setKd} />
          <Content busy={busy} id="kt" ic="grid" t="Tabel" tag=".png" s="Gambar tabel skor per indikator" on={kt} set={setKt} />
          <Content busy={busy} id="kv" ic="bars" t="Visualisasi" tag=".png" s="Gambar KPI dan semua grafik" on={kv} set={setKv} />
        </div>

        <div className={"dm-sum" + (files ? "" : " bad")}>
          <Icon name="info" />
          <span>{summary}</span>
        </div>
        <div className={"dm-st" + (busy ? " run" : "")}>{status}</div>

        <div className="mf">
          <button className="btn ol" onClick={() => !busy && onClose()}>Batal</button>
          <button className="btn pr" onClick={run} disabled={busy || !files}>
            <Icon name="download" sw={2.2} />Unduh
          </button>
        </div>
      </div>
    </div>
  );
}
