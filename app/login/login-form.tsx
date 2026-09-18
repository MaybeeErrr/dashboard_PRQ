import { LogIn } from "lucide-react";

const ERROR_MESSAGES: Record<string, string> = {
  config: "Credential dashboard belum dikonfigurasi di environment variable server.",
  credentials: "Username atau password tidak sesuai.",
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
  const message = !configured
    ? "Credential dashboard belum dikonfigurasi."
    : error
      ? (ERROR_MESSAGES[error] ?? "Login gagal. Coba lagi.")
      : null;

  return (
    <form action="/api/login" className="mt-8 space-y-4" method="post">
      <input name="returnTo" type="hidden" value={returnTo} />
      <label className="block">
        <span className="text-sm font-medium text-slate-700">Username</span>
        <input
          autoComplete="username"
          className="mt-2 h-11 w-full rounded-md border border-slate-200/90 bg-white/90 px-3 text-sm text-slate-950 shadow-sm outline-none transition focus:border-slate-400 focus:ring-4 focus:ring-slate-200"
          name="username"
          placeholder="Masukkan username"
          required
        />
      </label>
      <label className="block">
        <span className="text-sm font-medium text-slate-700">Password</span>
        <input
          autoComplete="current-password"
          className="mt-2 h-11 w-full rounded-md border border-slate-200/90 bg-white/90 px-3 text-sm text-slate-950 shadow-sm outline-none transition focus:border-slate-400 focus:ring-4 focus:ring-slate-200"
          name="password"
          placeholder="Masukkan password"
          required
          type="password"
        />
      </label>
      {message && (
        <div className="rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-sm text-amber-800">
          {message}
        </div>
      )}
      <button
        className="inline-flex h-11 w-full items-center justify-center gap-2 rounded-md bg-slate-950 px-4 text-sm font-semibold text-white shadow-[0_18px_34px_-22px_rgba(15,23,42,0.9)] transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300"
        disabled={!configured}
        type="submit"
      >
        <LogIn className="h-4 w-4" />
        Login
      </button>
    </form>
  );
}
