"use client";

import { useMemo, useState } from "react";
import {
  createLocalStorageBackup,
  exportLocalStorageBackup,
  migrateLegacyDataToSupabase,
  normalizeExistingLegacyEntriesToSupabase,
  scanLegacyLocalStorage,
  type LocalStorageBackupEntry,
  type LocalStorageMigrationReport,
} from "@/features/shared/storage/local-storage-migration";

function productCounts(entries: LocalStorageBackupEntry[]) {
  return entries.reduce(
    (counts, entry) => ({
      ...counts,
      [entry.product]: (counts[entry.product] ?? 0) + 1,
    }),
    {} as Partial<Record<LocalStorageBackupEntry["product"], number>>,
  );
}

export function LegacyStorageMigrationPanel() {
  const [entries, setEntries] = useState<LocalStorageBackupEntry[]>(() => scanLegacyLocalStorage());
  const [error, setError] = useState<string | null>(null);
  const [isMigrating, setIsMigrating] = useState(false);
  const [report, setReport] = useState<LocalStorageMigrationReport | null>(null);
  const counts = useMemo(() => productCounts(entries), [entries]);
  const persistentEntries = entries.filter((entry) => entry.classification === "BUSINESS_PERSISTENT");

  async function handleBackup() {
    exportLocalStorageBackup();
    setEntries(scanLegacyLocalStorage());
  }

  async function handleMigration() {
    setIsMigrating(true);
    setError(null);
    setReport(null);
    try {
      const backup = createLocalStorageBackup();
      setReport(await migrateLegacyDataToSupabase(backup));
      setEntries(scanLegacyLocalStorage());
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Migrasi gagal.");
    } finally {
      setIsMigrating(false);
    }
  }

  async function handleNormalizeExisting() {
    setIsMigrating(true);
    setError(null);
    setReport(null);
    try {
      setReport(await normalizeExistingLegacyEntriesToSupabase());
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Normalisasi HSI/WMS gagal.");
    } finally {
      setIsMigrating(false);
    }
  }

  return (
    <main className="min-h-screen bg-[linear-gradient(135deg,#ffffff_0%,#f5f9ff_45%,#edf5ff_100%)] px-6 py-8 text-slate-950">
      <section className="mx-auto max-w-6xl">
        <div className="rounded-[24px] border border-slate-200/80 bg-white/90 p-8 shadow-[0_18px_50px_rgba(15,23,42,0.08)]">
          <p className="text-xs font-bold uppercase tracking-[0.22em] text-blue-500">PRIME Data Migration</p>
          <h1 className="mt-3 text-4xl font-bold tracking-normal text-slate-950">Legacy localStorage to Supabase</h1>
          <p className="mt-3 max-w-3xl text-sm leading-6 text-slate-600">
            Tool ini membuat backup lengkap localStorage/sessionStorage browser, lalu mengirim data persistent ke Supabase. Data lokal lama tetap
            dipertahankan sebagai fallback dan tidak dihapus otomatis.
          </p>

          <div className="mt-6 flex flex-wrap gap-3">
            <button
              className="h-11 rounded-xl border border-slate-200 bg-white px-5 text-sm font-bold text-slate-700 shadow-sm transition hover:bg-slate-50"
              onClick={handleBackup}
              type="button"
            >
              Export Local Backup
            </button>
            <button
              className="h-11 rounded-xl bg-gradient-to-r from-cyan-400 via-blue-500 to-violet-600 px-5 text-sm font-bold text-white shadow-[0_12px_28px_rgba(37,99,235,0.22)] transition hover:-translate-y-0.5 disabled:cursor-wait disabled:opacity-60"
              disabled={isMigrating || persistentEntries.length === 0}
              onClick={handleMigration}
              type="button"
            >
              {isMigrating ? "Migrating..." : "Backup & Migrate to Supabase"}
            </button>
            <button
              className="h-11 rounded-xl border border-blue-200 bg-blue-50 px-5 text-sm font-bold text-blue-700 shadow-sm transition hover:bg-blue-100 disabled:cursor-wait disabled:opacity-60"
              disabled={isMigrating}
              onClick={handleNormalizeExisting}
              type="button"
            >
              Normalize Existing HSI/WMS
            </button>
          </div>
        </div>

        <div className="mt-6 grid gap-4 md:grid-cols-4">
          {(["HSI", "WMS", "LTS", "METABRIGHT"] as const).map((product) => (
            <div key={product} className="rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
              <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-400">{product}</p>
              <p className="mt-3 text-3xl font-bold text-slate-950">{counts[product] ?? 0}</p>
              <p className="mt-1 text-sm text-slate-500">legacy keys found</p>
            </div>
          ))}
        </div>

        {error ? (
          <div className="mt-6 rounded-2xl border border-rose-200 bg-rose-50 p-5 text-sm font-semibold text-rose-700">{error}</div>
        ) : null}

        {report ? (
          <div className="mt-6 rounded-2xl border border-slate-200 bg-white p-5 shadow-sm">
            <h2 className="text-xl font-bold text-slate-950">Migration Report</h2>
            <div className="mt-4 grid gap-3 md:grid-cols-4">
              {Object.entries(report.products).map(([product, result]) => (
                <div key={product} className="rounded-xl border border-slate-100 bg-slate-50 p-4">
                  <p className="text-xs font-bold uppercase tracking-[0.18em] text-slate-400">{product}</p>
                  <p className="mt-2 text-sm font-semibold text-slate-700">Status: {result.status}</p>
                  <p className="mt-1 text-sm text-slate-500">Rows: {result.migratedRows} / {result.localRows}</p>
                </div>
              ))}
            </div>
            {report.conflicts.length > 0 ? (
              <div className="mt-4 rounded-xl border border-amber-200 bg-amber-50 p-4 text-sm text-amber-800">
                <p className="font-bold">Conflicts detected</p>
                <ul className="mt-2 list-disc space-y-1 pl-5">
                  {report.conflicts.map((conflict) => (
                    <li key={conflict}>{conflict}</li>
                  ))}
                </ul>
              </div>
            ) : null}
          </div>
        ) : null}

        <section className="mt-6 overflow-hidden rounded-2xl border border-slate-200 bg-white shadow-sm">
          <div className="border-b border-slate-100 px-5 py-4">
            <h2 className="text-lg font-bold text-slate-950">Storage Audit</h2>
            <p className="mt-1 text-sm text-slate-500">{entries.length} key ditemukan di browser ini.</p>
          </div>
          <div className="overflow-auto">
            <table className="min-w-full text-left text-sm">
              <thead className="bg-slate-50 text-xs uppercase tracking-[0.12em] text-slate-500">
                <tr>
                  <th className="px-4 py-3">Area</th>
                  <th className="px-4 py-3">Key</th>
                  <th className="px-4 py-3">Product</th>
                  <th className="px-4 py-3">Classification</th>
                  <th className="px-4 py-3">Target</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {entries.map((entry) => (
                  <tr key={`${entry.area}.${entry.key}`}>
                    <td className="px-4 py-3 font-medium text-slate-700">{entry.area}</td>
                    <td className="px-4 py-3 font-mono text-xs text-slate-700">{entry.key}</td>
                    <td className="px-4 py-3 text-slate-600">{entry.product}</td>
                    <td className="px-4 py-3 text-slate-600">{entry.classification}</td>
                    <td className="px-4 py-3 text-slate-500">{entry.proposedTable}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </section>
    </main>
  );
}
