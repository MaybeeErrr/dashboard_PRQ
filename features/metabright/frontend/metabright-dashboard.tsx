"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Dataset, Tab } from "@/features/metabright/domain/types";
import defaultData from "@/features/metabright/data/default.json";
import { parseFile } from "@/features/metabright/domain/parse";
import { avg, fmt, isRatio, nf, periodLabel, wait } from "@/features/metabright/domain/utils";
import { Icon } from "./metabright-icons";
import {
  DistChart, IndicatorChart, RadarProfile, RankChart, TargetChart, WitelChart,
  indicatorAvg, witelColor,
} from "./metabright-charts";
import { DataTable, Heatmap } from "./metabright-tables";
import DownloadModal from "./metabright-download-modal";

const DEFAULT = defaultData as unknown as Dataset;
const DEFAULT_LABEL = periodLabel("202610");

type Msg = { t: string; c?: "okk" | "err" } | null;

export function MetabrightDashboard() {
  const [periods, setPeriods] = useState<Record<string, Dataset>>({ [DEFAULT_LABEL]: DEFAULT });
  const [cur, setCur] = useState<string | null>(DEFAULT_LABEL);
  const [witel, setWitel] = useState("ALL");
  const [iIdx, setIIdx] = useState(0);
  const [q, setQ] = useState("");
  const [topN, setTopN] = useState(20);
  const [radarTelda, setRadarTelda] = useState("");
  const [tab, setTab] = useState<Tab>("data");
  const [cap, setCap] = useState<Tab | null>(null);
  const [msg, setMsg] = useState<Msg>(null);
  const [over, setOver] = useState(false);
  const [del, setDel] = useState<string | null>(null);
  const [dlOpen, setDlOpen] = useState(false);

  const fileRef = useRef<HTMLInputElement>(null);
  const capRef = useRef<HTMLDivElement>(null);
  const saved = useRef({ witel: "ALL", q: "" });

  const flash = useCallback((t: string, c?: "okk" | "err") => setMsg({ t, c }), []);
  useEffect(() => {
    if (msg?.c !== "okk") return;
    const id = setTimeout(() => setMsg(null), 6000);
    return () => clearTimeout(id);
  }, [msg]);

  const d = cur ? periods[cur] : null;
  const shown: Tab = cap ?? tab;

  /* ---------- turunan ---------- */
  const ws = useMemo(() => (d ? [...new Set(d.rows.map((r) => r.witel))].sort() : []), [d]);
  const rows = useMemo(() => {
    if (!d) return [];
    const s = q.trim().toLowerCase();
    return d.rows.filter((r) => (witel === "ALL" || r.witel === witel) && (!s || r.telda.toLowerCase().includes(s)));
  }, [d, witel, q]);
  const sorted = useMemo(() => [...rows].sort((a, b) => b.total - a.total), [rows]);
  const byRank = useMemo(() => [...rows].sort((a, b) => a.rank - b.rank), [rows]);

  /* ---------- aksi ---------- */
  function setup(label: string | null) {
    setCur(label);
    setWitel("ALL");
    setQ("");
    setIIdx(0);
    setRadarTelda("");
  }

  async function loadMany(files: File[]) {
    let last: string | undefined;
    let n = 0;
    let warn = "";
    const next = { ...periods };
    for (const f of files) {
      try {
        const data = await parseFile(f);
        let lb = periodLabel(f.name);
        if (next[lb]) lb += " (" + f.name.slice(0, 18) + ")";
        next[lb] = data;
        last = lb;
        n++;
        if (data.warn) warn = data.warn;
      } catch (e) {
        flash("Gagal membaca " + f.name + ": " + (e instanceof Error ? e.message : String(e)), "err");
      }
    }
    if (last) {
      setPeriods(next);
      setup(last);
      flash(n + " file berhasil dimuat." + (warn ? " Catatan: " + warn : ""), warn ? undefined : "okk");
    }
  }

  function removePeriod(p: string) {
    const next = { ...periods };
    delete next[p];
    setPeriods(next);
    const keys = Object.keys(next);
    setup(keys.length ? (keys.includes(cur ?? "") ? cur : keys[0]) : null);
    setDel(null);
    flash('Data "' + p + '" telah dihapus.');
  }

  const begin = useCallback(() => { saved.current = { witel, q }; setQ(""); }, [witel, q]);
  const end = useCallback(() => {
    setCap(null);
    setWitel(saved.current.witel);
    setQ(saved.current.q);
  }, []);

  const capture = useCallback(async (w: string, mode: Tab): Promise<Blob> => {
    setWitel(w);
    setCap(mode);
    await wait(1000);
    const el = capRef.current;
    if (!el) throw new Error("area tangkapan tidak ditemukan");
    if (mode === "data") {
      const t = el.querySelector("#vData table") as HTMLElement | null;
      el.style.width = Math.max((t?.scrollWidth ?? 900) + 60, 900) + "px";
      await wait(250);
    }
    const html2canvas = (await import("html2canvas")).default;
    const W = el.offsetWidth;
    try {
      const c = await html2canvas(el, {
        backgroundColor: "#f7faff",
        scale: mode === "data" ? 2 : 1.5,
        useCORS: true,
        width: W,
        windowWidth: Math.max(W + 40, window.innerWidth),
      });
      return await new Promise<Blob>((res, rej) =>
        c.toBlob((b) => (b ? res(b) : rej(new Error("gagal membuat gambar"))), "image/png")
      );
    } finally {
      el.style.width = "";
    }
  }, []);

  /* ---------- render ---------- */
  const ind = d?.indicators[iIdx] ?? d?.indicators[0];
  const radarT = rows.find((r) => r.telda === radarTelda) ?? byRank[0];

  let kpis: { label: string; value: string; sub: string; bar?: number }[] = [];
  let sor: { cls: string; title: string; text: string }[] = [];
  let ach = 0;
  if (d && ind && rows.length) {
    const agg = (a: number[]) => (isRatio(ind.name) ? avg(a) : a.reduce((x, y) => x + y, 0));
    const T = agg(rows.map((r) => r.v[iIdx][0]));
    const R = agg(rows.map((r) => r.v[iIdx][1]));
    ach = T ? (R / T) * 100 : 0;
    const hi = sorted[0];
    const lo = sorted[sorted.length - 1];
    kpis = [
      { label: "Telda dianalisis", value: String(rows.length), sub: "dari " + d.rows.length + " telda" },
      { label: "Rata-rata total score", value: nf(avg(rows.map((r) => r.total))), sub: `Rentang ${nf(lo.total)} – ${nf(hi.total)}` },
      { label: "Skor tertinggi", value: nf(hi.total), sub: hi.telda },
      { label: "Skor terendah", value: nf(lo.total), sub: lo.telda },
      { label: "Achievement " + ind.name, value: nf(ach) + "%", sub: `Target ${fmt(T)} · Realisasi ${fmt(R)}`, bar: Math.min(ach, 100) },
    ];
    const iv = indicatorAvg(rows, d).map((p, k) => ({ n: d.indicators[k].name, p }));
    const by = [...iv].sort((a, b) => b.p - a.p);
    const best = by[0];
    const worst = by[by.length - 1];
    const ok = rows.filter((r) => r.v[iIdx][2] >= 100).length;
    sor = [
      { cls: "ok", title: "Indikator terkuat", text: `${best.n}, rata-rata ${nf(best.p, 0)}% dari bobot.` },
      { cls: "bad", title: "Paling perlu dikejar", text: `${worst.n}, rata-rata baru ${nf(worst.p, 0)}% dari bobot.` },
      { cls: "", title: ind.name, text: `${ok} dari ${rows.length} telda sudah mencapai ACH 100% atau lebih.` },
      { cls: "mid", title: "Selisih skor", text: `${nf(hi.total - lo.total)} poin antara ${hi.telda} (tertinggi) dan ${lo.telda} (terendah).` },
    ];
  }

  const scopeLabel = witel === "ALL" ? "Semua witel" : witel;
  const hasFilter = !!d && (witel !== "ALL" || q.trim() !== "");
  const keys = Object.keys(periods);

  return (
    <main className="metabright-page wrap">
      {/* Header */}
      <header className="top">
        <Link className="back-home" href="/portal">
          <span aria-hidden="true">←</span>
          Back to Home
        </Link>
        <div className="brand">
          <div className="logo">MB</div>
          <div>
            <p className="eyebrow">PRIME • METABRIGHT</p>
            <h2>Impactful Award Telda</h2>
            <p>Monitoring performa dan pencapaian Telda berdasarkan indikator MetaBright.</p>
            <p>{d ? `${d.rows.length} telda · ${d.indicators.length} indikator` : "Belum ada data"}</p>
          </div>
        </div>
        <nav aria-label="Pilih tampilan">
          <div className="seg">
            <button className={"tabbtn" + (tab === "data" ? " on" : "")} onClick={() => setTab("data")}>
              <Icon name="data" />Tabel Data
            </button>
            <button className={"tabbtn" + (tab === "viz" ? " on" : "")} onClick={() => setTab("viz")}>
              <Icon name="bars" />Visualisasi
            </button>
          </div>
          <div className="dl">
            <button className="btn gh" onClick={() => (d ? setDlOpen(true) : flash("Belum ada data untuk diunduh.", "err"))}>
              <Icon name="download" />Unduh data / visualisasi
            </button>
          </div>
        </nav>
      </header>

      {/* Hero + filter */}
      <section className="card hero">
        <div className="in">
          <div className="head">
            <div className="ttl">
              <h1>Impactful Award Telda<span className="pp">{cur ?? ""}</span></h1>
              {d && ind && (
                <div className="sums">
                  <span className="st"><Icon name="pin" /><span>{scopeLabel}</span></span>
                  <span className="st"><Icon name="grid" /><span>{rows.length} dari {d.rows.length} telda</span></span>
                  <span className="st" title={ind.name}><Icon name="bars" /><span>Indikator: {ind.name}</span></span>
                </div>
              )}
            </div>
            <button className="btn pr upb" onClick={() => fileRef.current?.click()}>
              <Icon name="upload" />Upload Data
            </button>
            <input
              ref={fileRef}
              type="file"
              accept=".xls,.xlsx,.csv"
              multiple
              hidden
              onChange={(e) => { loadMany([...(e.target.files ?? [])]); e.target.value = ""; }}
            />
          </div>

          <div className="fbar">
            <div className="fh">
              <span className="ft"><Icon name="filter" />Filter tampilan</span>
              {hasFilter && <button className="lnk" onClick={() => { setWitel("ALL"); setQ(""); }}>Atur ulang</button>}
            </div>
            <div className="fg">
              <div className="ff">
                <label htmlFor="fw"><Icon name="pin" />Witel</label>
                <select id="fw" value={witel} onChange={(e) => setWitel(e.target.value)} disabled={!d}>
                  <option value="ALL">Semua witel</option>
                  {ws.map((w) => <option key={w}>{w}</option>)}
                </select>
              </div>
              <div className="ff">
                <label htmlFor="fi"><Icon name="target" />Indikator</label>
                <select id="fi" value={iIdx} onChange={(e) => setIIdx(+e.target.value)} disabled={!d}>
                  {d?.indicators.map((x, i) => <option key={i} value={i}>{x.name} (bobot {x.weight})</option>)}
                </select>
              </div>
              <div className="ff">
                <label htmlFor="fq"><Icon name="search" />Cari telda</label>
                <input id="fq" type="search" placeholder="Ketik nama telda" value={q} onChange={(e) => setQ(e.target.value)} />
              </div>
            </div>
          </div>

          <div className="meta">
            <div className="chips">
              {keys.length > 0 && <span>Data tersimpan:</span>}
              {keys.map((p) => (
                <span key={p} className={"chip" + (p === cur ? " cur" : "")} onClick={() => setup(p)}>
                  {p}
                  <u onClick={(e) => { e.stopPropagation(); setDel(p); }}>Hapus</u>
                </span>
              ))}
            </div>
            <div
              className={"drop" + (over ? " over" : "")}
              onClick={() => fileRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setOver(true); }}
              onDragEnter={(e) => { e.preventDefault(); setOver(true); }}
              onDragLeave={(e) => { e.preventDefault(); setOver(false); }}
              onDrop={(e) => { e.preventDefault(); setOver(false); loadMany([...e.dataTransfer.files]); }}
            >
              <Icon name="upload" />
              <span><b>Tarik dan lepas</b> file Excel ke sini, atau klik area ini untuk memilih file.</span>
            </div>
          </div>
          {msg && <div className={"msg" + (msg.c ? " " + msg.c : "")} role="status">{msg.t}</div>}
        </div>
      </section>

      {/* Area tampilan / tangkapan gambar */}
      <div id="cap" ref={capRef} className={cap ? "capmode" : ""} style={cap === "viz" ? { width: 1440 } : undefined}>
        <div className="cph">
          <b>Impactful Award Telda</b><i>{cur}</i>
          <span>
            {scopeLabel} · {rows.length} telda{cap === "viz" && ind ? " · Indikator: " + ind.name : ""}
          </span>
        </div>

        {!d || !ind ? (
          <div className="card empty">Belum ada data. Upload file Excel Impactful Award Telda untuk menampilkan tabel dan visualisasi.</div>
        ) : shown === "data" ? (
          <div id="vData">
            <section className="card">
              <div className="ch"><div><h3>Data skor per indikator</h3></div></div>
              <DataTable rows={rows} d={d} />
            </section>
          </div>
        ) : (
          <div id="vViz">
            <div className="kpis">
              {kpis.map((k, j) => (
                <div key={j} className={"k k" + j}>
                  <small>{k.label}</small>
                  <strong>{k.value}</strong>
                  <em>{k.sub}</em>
                  {k.bar != null && <div className="kb"><i style={{ width: k.bar + "%" }} /></div>}
                </div>
              ))}
            </div>

            <div className="g3">
              <section className="card">
                <div className="ch"><div><h3>Tiga telda teratas</h3><div className="sub">Berdasarkan total score, skor maksimum 110</div></div></div>
                <div className="pod">
                  {sorted.length ? [1, 0, 2].filter((k) => sorted[k]).map((k) => {
                    const r = sorted[k];
                    return (
                      <div className="pd" key={r.telda}>
                        <div className="pdn" title={r.telda}>{r.telda}</div>
                        <div className="pdw">{r.witel}</div>
                        <div className="pds">{nf(r.total)}</div>
                        <div className={"pdb p" + (k + 1)} style={{ height: [132, 100, 78][k] }}>{k + 1}</div>
                      </div>
                    );
                  }) : <div className="empty">Tidak ada telda pada filter ini.</div>}
                </div>
              </section>
              <section className="card">
                <h3>Sorotan</h3>
                <div className="sub">Dihitung dari hasil filter saat ini</div>
                <div className="sor">
                  {sor.length ? sor.map((s, k) => (
                    <div key={k} className={"si " + s.cls}><b>{s.title}</b><span>{s.text}</span></div>
                  )) : <div className="empty">Tidak ada data.</div>}
                </div>
              </section>
            </div>

            {rows.length > 0 && (
              <>
                <div className="g2">
                  <section className="card fill">
                    <div className="ch">
                      <div><h3>Peringkat total score</h3><div className="sub">Warna menunjukkan witel</div></div>
                      <select value={topN} onChange={(e) => setTopN(+e.target.value)}>
                        <option value={10}>10 teratas</option>
                        <option value={20}>20 teratas</option>
                        <option value={0}>Semua telda</option>
                      </select>
                    </div>
                    <div className="box"><RankChart rows={rows} ws={ws} topN={topN} /></div>
                    <div className="leg">
                      {ws.map((w) => <span key={w}><i style={{ background: witelColor(ws, w) }} />{w}</span>)}
                    </div>
                  </section>
                  <div>
                    <section className="card">
                      <h3>Rata-rata total score per witel</h3>
                      <div className="sub">Skor maksimum 110</div>
                      <div className="box" style={{ marginTop: 10 }}><WitelChart rows={rows} ws={ws} /></div>
                    </section>
                    <section className="card">
                      <h3>Capaian rata-rata per indikator</h3>
                      <div className="sub">Skor diperoleh terhadap bobot indikator (%)</div>
                      <div className="box" style={{ marginTop: 10 }}><IndicatorChart rows={rows} d={d} /></div>
                    </section>
                  </div>
                </div>

                <div className="g2">
                  <section className="card">
                    <div className="ch">
                      <div><h3>Profil capaian telda</h3><div className="sub">Dibandingkan dengan rata-rata hasil filter</div></div>
                      <select value={radarT?.telda ?? ""} onChange={(e) => setRadarTelda(e.target.value)}>
                        {byRank.map((r) => <option key={r.telda}>{r.telda}</option>)}
                      </select>
                    </div>
                    <div className="box" style={{ height: 430 }}><RadarProfile t={radarT} rows={rows} d={d} /></div>
                  </section>
                  <section className="card">
                    <h3>Distribusi pencapaian indikator</h3>
                    <div className="sub">Jumlah telda per rentang ACH — {ind.name}</div>
                    <div className="box" style={{ height: 430, marginTop: 10 }}><DistChart rows={rows} i={iIdx} /></div>
                  </section>
                </div>

                <section className="card">
                  <div className="ch">
                    <div>
                      <h3>Peta capaian witel per indikator</h3>
                      <div className="sub">Rata-rata skor diperoleh terhadap bobot indikator. Sel merah menunjukkan area yang perlu dikejar.</div>
                    </div>
                  </div>
                  <div className="scroll"><Heatmap rows={rows} d={d} /></div>
                </section>

                <section className="card">
                  <div className="ch"><div><h3>Target, realisasi, dan achievement per telda</h3><div className="sub">Indikator: {ind.name}</div></div></div>
                  <div className="scroll">
                    <div style={{ height: 400, width: Math.max(rows.length * 60, 300), minWidth: "100%" }}>
                      <TargetChart rows={rows} i={iIdx} />
                    </div>
                  </div>
                </section>
              </>
            )}
          </div>
        )}

        <div className="cpf">
          Dibuat {new Date().toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" })} · Dashboard Impactful Award Telda
        </div>
      </div>

      {/* Modal hapus */}
      {del && (
        <div className="ov" onClick={(e) => { if (e.target === e.currentTarget) setDel(null); }}>
          <div className="md" role="dialog" aria-modal="true">
            <h3>Hapus data?</h3>
            <p>Data periode <b>{del}</b> akan dihapus dari dashboard. Anda dapat mengupload ulang file tersebut kapan saja.</p>
            <div className="mf">
              <button className="btn ol" onClick={() => setDel(null)}>Batal</button>
              <button className="btn dn" onClick={() => removePeriod(del)}>Hapus</button>
            </div>
          </div>
        </div>
      )}

      {/* Modal unduh */}
      {dlOpen && d && cur && (
        <DownloadModal
          period={cur}
          data={d}
          capture={capture}
          begin={begin}
          end={end}
          onClose={() => setDlOpen(false)}
          onMsg={flash}
        />
      )}
    </main>
  );
}

export default MetabrightDashboard;
