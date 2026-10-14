"use client";

import { useState } from "react";
import type React from "react";

type MasterUnlockModalProps = {
  onClose: () => void;
  onUnlocked: (token: string) => void;
  open: boolean;
  theme?: "hsi" | "wms" | "lts";
};

const themeClass = {
  hsi: "from-[#071426] via-[#0a1f3b] to-[#123f68] focus:border-cyan-300 focus:ring-cyan-300/20",
  lts: "from-[#080712] via-[#24103f] to-[#43206a] focus:border-violet-300 focus:ring-violet-300/20",
  wms: "from-[#21130b] via-[#3b2414] to-[#5a3519] focus:border-amber-300 focus:ring-amber-300/20",
} as const;

export function MasterUnlockModal({ onClose, onUnlocked, open, theme = "hsi" }: MasterUnlockModalProps) {
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [submitting, setSubmitting] = useState(false);

  if (!open) return null;

  async function submit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    setSubmitting(true);
    setError(null);
    try {
      const response = await fetch("/api/master-unlock", {
        body: JSON.stringify({ password }),
        headers: { "Content-Type": "application/json" },
        method: "POST",
      });
      if (!response.ok) {
        const payload = (await response.json().catch(() => null)) as { error?: string } | null;
        throw new Error(payload?.error ?? "Password MASTER salah.");
      }
      const payload = (await response.json().catch(() => null)) as { token?: unknown } | null;
      const token = typeof payload?.token === "string" ? payload.token : "";
      if (!token) throw new Error("Token MASTER tidak valid.");
      setPassword("");
      onUnlocked(token);
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "Password MASTER salah.");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <div className="fixed inset-0 z-[90] flex items-center justify-center bg-slate-950/72 px-4 backdrop-blur-sm">
      <form
        className={`w-full max-w-[420px] rounded-2xl border border-white/12 bg-gradient-to-br ${themeClass[theme]} p-5 text-white shadow-[0_30px_90px_-42px_rgba(0,0,0,0.95)]`}
        onSubmit={submit}
      >
        <p className="text-xs font-semibold uppercase text-white/55">MASTER ACCESS</p>
        <h2 className="mt-1 text-2xl font-bold tracking-normal">Masukkan Password</h2>
        <p className="mt-2 text-sm leading-6 text-white/65">Halaman MASTER terkunci untuk menjaga data upload tetap aman.</p>
        <label className="mt-5 grid gap-2 text-xs font-semibold uppercase text-white/65">
          Password
          <input
            autoFocus
            className="h-11 rounded-xl border border-white/12 bg-black/24 px-3 text-sm font-semibold text-white outline-none ring-4 ring-transparent transition placeholder:text-white/35"
            placeholder="Masukkan password MASTER"
            type="password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        {error && <p className="mt-3 rounded-lg border border-rose-300/20 bg-rose-500/12 px-3 py-2 text-sm font-semibold text-rose-100">{error}</p>}
        <div className="mt-5 flex justify-end gap-2">
          <button className="h-10 rounded-lg border border-white/12 px-4 text-sm font-semibold text-white/78 transition hover:bg-white/10" type="button" onClick={onClose}>
            Batal
          </button>
          <button className="h-10 rounded-lg bg-white px-4 text-sm font-bold text-slate-950 transition hover:brightness-95 disabled:cursor-not-allowed disabled:opacity-60" disabled={submitting} type="submit">
            {submitting ? "Memeriksa..." : "Unlock"}
          </button>
        </div>
      </form>
    </div>
  );
}
