import { cookies } from "next/headers";
import { redirect } from "next/navigation";
import { credentialsConfigured, SESSION_COOKIE, verifySessionToken } from "@/lib/auth-session";
import { LoginForm } from "./login-form";

export const dynamic = "force-dynamic";

function safeReturnTo(value: string | undefined) {
  if (!value || !value.startsWith("/") || value.startsWith("//")) return "/";
  return value;
}

export default async function LoginPage({
  searchParams,
}: {
  searchParams?: Promise<{ error?: string; returnTo?: string }>;
}) {
  const cookieStore = await cookies();
  const authenticated = await verifySessionToken(cookieStore.get(SESSION_COOKIE)?.value);
  if (authenticated) redirect("/");
  const params = await searchParams;
  const returnTo = safeReturnTo(params?.returnTo);

  return (
    <main className="relative flex min-h-screen items-center justify-center overflow-hidden bg-[linear-gradient(180deg,#f8fafc_0%,#eef2f7_48%,#fff7f7_100%)] px-4 py-10 text-slate-950">
      <div className="pointer-events-none absolute inset-0 bg-[linear-gradient(rgba(15,23,42,0.045)_1px,transparent_1px),linear-gradient(90deg,rgba(15,23,42,0.045)_1px,transparent_1px)] bg-[size:46px_46px] opacity-70" />
      <section className="relative w-full max-w-md rounded-lg border border-white/70 bg-white/90 p-6 shadow-[0_28px_90px_-50px_rgba(15,23,42,0.8)] ring-1 ring-slate-950/[0.03] backdrop-blur sm:p-8">
        <div className="flex items-center gap-3">
          <div className="flex h-11 w-11 items-center justify-center rounded-lg bg-slate-950 text-sm font-bold text-white shadow-[0_16px_32px_-18px_rgba(15,23,42,0.95)]">HSI</div>
          <div>
            <p className="text-sm font-semibold text-slate-950">Reporting HSI 2026</p>
            <p className="text-xs text-slate-500">Secure dashboard access</p>
          </div>
        </div>
        <div className="mt-8">
          <p className="text-sm font-semibold uppercase text-rose-700">Login</p>
          <h1 className="mt-2 text-2xl font-semibold tracking-normal text-slate-950">Masuk ke dashboard</h1>
          <p className="mt-2 text-sm leading-6 text-slate-500">
            Gunakan credential yang dikonfigurasi di environment variable server.
          </p>
        </div>
        <LoginForm configured={credentialsConfigured()} error={params?.error} returnTo={returnTo} />
      </section>
    </main>
  );
}
