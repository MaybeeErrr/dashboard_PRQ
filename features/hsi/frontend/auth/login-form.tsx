"use client";

import { useId, useState, type FormEvent } from "react";
import { Eye, EyeOff, Loader2, LogIn } from "lucide-react";
import { useRouter } from "next/navigation";
import { TAB_SESSION_KEY } from "@/features/shared/dashboard/session-tab-guard";

const ERROR_MESSAGES: Record<string, string> = {
  config: "Credential dashboard belum dikonfigurasi di environment variable server.",
  credentials: "Username or password is incorrect.",
};

export function LoginForm({
  configured,
  error,
  returnTo,
}: {
  configured: boolean;
  error?: string;
  returnTo: string;
}) {
  const router = useRouter();
  const usernameId = useId();
  const passwordId = useId();
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [clientError, setClientError] = useState<string | null>(null);

  const message = clientError ?? (!configured
    ? "Credential dashboard belum dikonfigurasi."
    : error
      ? (ERROR_MESSAGES[error] ?? "Login gagal. Coba lagi.")
      : null);

  async function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!configured || loading) return;

    const formData = new FormData(event.currentTarget);
    setLoading(true);
    setClientError(null);

    try {
      const response = await fetch("/api/login", {
        body: JSON.stringify({
          password: String(formData.get("password") ?? ""),
          returnTo,
          username: String(formData.get("username") ?? ""),
        }),
        headers: {
          "Content-Type": "application/json",
        },
        method: "POST",
      });

      if (!response.ok) {
        setClientError(response.status === 401 ? "Username or password is incorrect." : "Login gagal. Coba lagi.");
        return;
      }

      window.sessionStorage.setItem(TAB_SESSION_KEY, "1");
      router.push("/portal");
      router.refresh();
    } catch {
      setClientError("Login gagal. Periksa koneksi lalu coba lagi.");
    } finally {
      setLoading(false);
    }
  }

  return (
    <form action="/api/login" className="mt-8 space-y-5" method="post" onSubmit={handleSubmit}>
      <input name="returnTo" type="hidden" value={returnTo} />
      <div className="block">
        <label className="text-sm font-semibold text-slate-700" htmlFor={usernameId}>
          Username
        </label>
        <input
          autoComplete="username"
          className="mt-2 h-12 w-full rounded-xl border border-[#dce5f0] bg-slate-50 px-4 text-sm font-medium text-slate-950 outline-none transition duration-200 placeholder:text-slate-400 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10 disabled:cursor-not-allowed disabled:opacity-60"
          disabled={loading || !configured}
          id={usernameId}
          name="username"
          placeholder="Enter your username"
          required
        />
      </div>
      <div className="block">
        <label className="text-sm font-semibold text-slate-700" htmlFor={passwordId}>
          Password
        </label>
        <div className="relative mt-2">
          <input
            autoComplete="current-password"
            className="h-12 w-full rounded-xl border border-[#dce5f0] bg-slate-50 px-4 pr-12 text-sm font-medium text-slate-950 outline-none transition duration-200 placeholder:text-slate-400 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/10 disabled:cursor-not-allowed disabled:opacity-60"
            disabled={loading || !configured}
            id={passwordId}
            name="password"
            placeholder="Enter your password"
            required
            type={showPassword ? "text" : "password"}
          />
          <button
            aria-label={showPassword ? "Hide password" : "Show password"}
            className="absolute right-3 top-1/2 inline-flex h-8 w-8 -translate-y-1/2 items-center justify-center rounded-lg text-slate-500 transition hover:bg-blue-50 hover:text-blue-600 focus:outline-none focus:ring-2 focus:ring-blue-300/50 disabled:cursor-not-allowed disabled:opacity-50"
            disabled={loading || !configured}
            onClick={() => setShowPassword((value) => !value)}
            type="button"
          >
            {showPassword ? <EyeOff className="h-4 w-4" /> : <Eye className="h-4 w-4" />}
          </button>
        </div>
      </div>
      {message && (
        <div className="rounded-xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm font-medium text-rose-700">
          {message}
        </div>
      )}
      <button
        aria-busy={loading}
        className="inline-flex h-14 w-full items-center justify-center gap-2 rounded-xl bg-[linear-gradient(115deg,#38bdf8_0%,#3b82f6_36%,#7c3aed_68%,#f472b6_100%)] px-4 text-sm font-bold uppercase tracking-[0.12em] text-white shadow-[0_18px_36px_-22px_rgba(59,130,246,0.85)] transition duration-200 hover:-translate-y-0.5 hover:brightness-105 hover:shadow-[0_22px_46px_-24px_rgba(124,58,237,0.75)] active:translate-y-0 disabled:cursor-not-allowed disabled:opacity-50 disabled:hover:translate-y-0"
        disabled={!configured || loading}
        type="submit"
      >
        {loading ? <Loader2 className="h-4 w-4 animate-spin" /> : <LogIn className="h-4 w-4" />}
        {loading ? "Signing in..." : "Sign in"}
      </button>
    </form>
  );
}
