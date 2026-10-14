import Image from "next/image";
import { credentialsConfigured } from "@/features/hsi/backend/auth/session";
import { LoginForm } from "@/features/hsi/frontend/auth/login-form";
import { LoginSessionReset } from "@/features/shared/dashboard/session-tab-guard";

export const dynamic = "force-dynamic";

export default async function LoginPage({
  searchParams,
}: {
  searchParams?: Promise<{ error?: string; returnTo?: string }>;
}) {
  const params = await searchParams;
  const returnTo = "/portal";

  return (
    <main className="relative isolate min-h-screen overflow-hidden bg-[linear-gradient(135deg,#ffffff_0%,#f5f9ff_45%,#edf5ff_100%)] px-5 py-7 text-slate-950 sm:px-8 lg:px-12">
      <LoginSessionReset />
      <style>
        {`
          @keyframes primeLoginBrand {
            from { opacity: 0; transform: translateY(8px) scale(0.985); }
            to { opacity: 1; transform: translateY(0) scale(1); }
          }
          @keyframes primeLoginPanel {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
          }
        `}
      </style>

      <div className="pointer-events-none absolute inset-0 opacity-[0.42] [background-image:linear-gradient(rgba(37,99,235,0.05)_1px,transparent_1px),linear-gradient(90deg,rgba(37,99,235,0.04)_1px,transparent_1px)] [background-size:96px_96px]" />
      <div className="pointer-events-none absolute -left-[18rem] top-28 h-[36rem] w-[54rem] rotate-[-18deg] rounded-[50%] border border-sky-200/50 bg-sky-100/20" />
      <div className="pointer-events-none absolute -right-[20rem] top-10 h-[42rem] w-[58rem] rotate-[20deg] rounded-[50%] border border-blue-200/45 bg-white/35" />
      <div className="pointer-events-none absolute bottom-[-20rem] left-1/2 h-[34rem] w-[86rem] -translate-x-1/2 rounded-[50%] bg-blue-100/35 blur-3xl" />

      <header className="relative z-20 flex items-start justify-between">
        <div className="flex h-20 w-64 items-center md:h-24 md:w-72">
          <Image
            alt="Telkom Indonesia"
            className="h-full w-full object-contain"
            height={64}
            priority
            src="/logo-telkom-latest.png"
            unoptimized
            width={210}
          />
        </div>
        <button
          aria-label="User profile"
          className="grid h-12 w-12 place-items-center rounded-full border border-slate-200 bg-white text-base font-semibold text-slate-950 shadow-[0_10px_24px_rgba(15,23,42,0.14)]"
          type="button"
        >
          N
        </button>
      </header>

      <div className="relative z-10 mx-auto grid min-h-[calc(100vh-7rem)] w-full max-w-[1400px] grid-cols-1 items-center gap-10 py-8 lg:grid-cols-[55fr_45fr] lg:gap-16">
        <section className="relative flex items-center justify-center px-2 py-6 sm:px-6 lg:px-10">
          <div className="pointer-events-none absolute left-1/2 top-1/2 h-[min(70vw,660px)] w-[min(70vw,660px)] -translate-x-1/2 -translate-y-1/2 rounded-full bg-[radial-gradient(circle,rgba(59,130,246,0.13),rgba(56,189,248,0.06)_38%,transparent_70%)] blur-3xl" />
          <div className="relative w-full max-w-[560px] animate-[primeLoginBrand_560ms_ease-out_both] text-center">
            <div className="mx-auto mb-7 h-1.5 w-16 rounded-full bg-[linear-gradient(90deg,#38bdf8,#3b82f6,#7c3aed,#f472b6)]" />
            <Image
              alt="PRIME"
              className="mx-auto h-auto w-[min(62vw,360px)] select-none object-contain drop-shadow-[0_28px_58px_rgba(59,130,246,0.16)]"
              draggable={false}
              height={900}
              priority
              src="/module-covers/prime-p-logo.png"
              width={1000}
            />
            <h1 className="mt-7 text-5xl font-extrabold tracking-[0.22em] text-slate-950 sm:text-6xl">PRIME</h1>
            <p className="mt-5 text-base font-medium leading-7 text-slate-600">
              PRQ Reporting, Integration, Monitoring &amp; Evaluation
            </p>
            <p className="mt-4 text-sm font-semibold uppercase tracking-[0.28em] text-slate-400">
              HSI • WMS • LTS • METABRIGHT
            </p>
          </div>
        </section>

        <div className="pointer-events-none absolute bottom-[10%] left-[55%] top-[10%] hidden w-px bg-[linear-gradient(180deg,transparent_0%,rgba(56,189,248,0.18)_28%,rgba(148,163,184,0.18)_52%,rgba(244,114,182,0.16)_76%,transparent_100%)] lg:block" />

        <section className="flex items-center justify-center px-1 pb-6 pt-0 sm:px-6 lg:px-10 lg:py-10">
          <div className="w-full max-w-[440px] animate-[primeLoginPanel_520ms_ease-out_120ms_both] rounded-[24px] border border-[#e5edf7] bg-white p-6 shadow-[0_20px_60px_rgba(15,23,42,0.08)] sm:p-8">
            <div className="mb-7">
              <div className="mb-5 h-1.5 w-16 rounded-full bg-[linear-gradient(90deg,#38bdf8,#3b82f6,#7c3aed,#f472b6)]" />
              <p className="mb-3 text-xs font-bold uppercase tracking-[0.22em] text-sky-500">PRIME PLATFORM</p>
              <h1 className="text-3xl font-extrabold tracking-normal text-slate-950">Welcome to PRIME</h1>
              <p className="mt-3 text-sm leading-6 text-slate-500">
                Sign in to access the PRQ Integrated Reporting Portal
              </p>
            </div>
            <LoginForm configured={credentialsConfigured()} error={params?.error} returnTo={returnTo} />
          </div>
        </section>
      </div>
    </main>
  );
}
