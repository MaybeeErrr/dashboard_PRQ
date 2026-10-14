"use client";

import Link from "next/link";
import Image from "next/image";
import { ArrowRight, LogOut } from "lucide-react";
import { useRouter } from "next/navigation";
import { useState } from "react";
import { reportingModules, type ReportingModule, type ReportingModuleId } from "@/features/portal/modules";
import type { SessionUser } from "@/features/hsi/backend/auth/session";
import { TAB_SESSION_KEY } from "@/features/shared/dashboard/session-tab-guard";

const coverImages: Record<ReportingModuleId, string> = {
  hsi: "/module-covers/hsi-design-card.png",
  lts: "/module-covers/lts-design-card.png",
  metabright: "/module-covers/metabright-logo-clean.png",
  wms: "/module-covers/wms-design-card.png",
};

const moduleThemes = {
  hsi: {
    accentLine: "bg-blue-500",
    arrow: "border-blue-200 bg-blue-50 text-blue-600 group-hover:bg-blue-100",
    description: "Home Service Infrastructure",
    imagePanel: "from-blue-50 via-sky-50 to-white ring-blue-100",
    imageClass: "object-cover",
    imageOverlay: "bg-[radial-gradient(circle_at_70%_8%,rgba(255,255,255,0.92),transparent_34%)]",
    title: "HSI",
  },
  wms: {
    accentLine: "bg-orange-500",
    arrow: "border-orange-200 bg-orange-50 text-orange-600 group-hover:bg-orange-100",
    description: "Warehouse Management System",
    imagePanel: "from-orange-50 via-amber-50 to-white ring-orange-100",
    imageClass: "object-cover",
    imageOverlay: "bg-[radial-gradient(circle_at_70%_8%,rgba(255,255,255,0.92),transparent_34%)]",
    title: "WMS",
  },
  lts: {
    accentLine: "bg-violet-500",
    arrow: "border-violet-200 bg-violet-50 text-violet-600 group-hover:bg-violet-100",
    description: "Lastmile Transport System",
    imagePanel: "from-violet-50 via-purple-50 to-white ring-violet-100",
    imageClass: "object-cover",
    imageOverlay: "bg-[radial-gradient(circle_at_70%_8%,rgba(255,255,255,0.92),transparent_34%)]",
    title: "LTS",
  },
  metabright: {
    accentLine: "bg-gradient-to-r from-cyan-500 via-blue-500 to-fuchsia-500",
    arrow: "border-cyan-200 bg-cyan-50 text-blue-600 group-hover:bg-cyan-100",
    description: "MetaBright",
    imagePanel: "border border-cyan-100/80 from-white via-white to-sky-50 ring-cyan-200/70",
    imageClass: "scale-110 object-contain p-2 drop-shadow-[0_16px_24px_rgba(37,99,235,0.14)]",
    imageOverlay: "bg-[radial-gradient(circle_at_center,rgba(56,189,248,0.10),transparent_64%)]",
    title: "METABRIGHT",
  },
} satisfies Record<ReportingModuleId, {
  accentLine: string;
  arrow: string;
  description: string;
  imageClass: string;
  imageOverlay: string;
  imagePanel: string;
  title: string;
}>;

function ModuleCard({ module }: { module: ReportingModule }) {
  const theme = moduleThemes[module.id];

  return (
    <Link
      aria-label={`Open ${module.initials}`}
      className="group flex h-[430px] overflow-hidden rounded-[24px] border border-slate-200/80 bg-white/90 p-3 shadow-[0_12px_32px_rgba(15,23,42,0.08)] outline-none transition duration-300 ease-out hover:-translate-y-1 hover:border-slate-300 hover:bg-white hover:shadow-[0_18px_42px_rgba(15,23,42,0.12)] focus:ring-4 focus:ring-sky-200/60"
      href={module.href}
    >
      <article className="flex min-h-0 w-full flex-col">
        <div className={`relative min-h-0 flex-1 overflow-hidden rounded-[20px] bg-gradient-to-br ${theme.imagePanel} ring-1`}>
          <div className={`pointer-events-none absolute inset-0 ${theme.imageOverlay}`} />
          <Image
            alt={module.initials}
            className={`${theme.imageClass} transition duration-300 group-hover:brightness-[1.03]`}
            fill
            priority={module.id === "hsi"}
            sizes="(max-width: 640px) 92vw, (max-width: 1024px) 46vw, 24vw"
            src={coverImages[module.id]}
          />
        </div>

        <div className="flex items-end justify-between gap-4 px-3 pb-4 pt-5">
          <div className="min-w-0">
            <h2 className="text-3xl font-bold uppercase tracking-normal text-slate-950">{theme.title}</h2>
            <div className={`mt-2 h-1 w-11 rounded-full ${theme.accentLine}`} />
            <p className="mt-3 text-sm font-medium leading-5 text-slate-500">{theme.description}</p>
          </div>
          <span className={`grid h-12 w-12 shrink-0 place-items-center rounded-full border transition duration-300 ${theme.arrow}`}>
            <ArrowRight className="h-5 w-5 transition duration-300 group-hover:translate-x-0.5" />
          </span>
        </div>
      </article>
    </Link>
  );
}

function UserMenu({ user }: { user: SessionUser | null }) {
  const router = useRouter();
  const [isLoggingOut, setIsLoggingOut] = useState(false);
  const displayName = user?.username?.trim() || "User";
  const initial = displayName.charAt(0).toUpperCase() || "U";

  async function handleLogout() {
    setIsLoggingOut(true);
    window.sessionStorage.removeItem(TAB_SESSION_KEY);
    await fetch("/api/logout", { method: "POST" });
    router.replace("/login?returnTo=%2Fportal");
    router.refresh();
  }

  return (
    <div className="flex items-center gap-2 rounded-full border border-slate-200/80 bg-white/85 px-2 py-2 shadow-[0_12px_28px_rgba(15,23,42,0.10)] backdrop-blur">
      <div className="grid h-10 w-10 place-items-center rounded-full bg-slate-950 text-sm font-bold text-white shadow-[0_8px_20px_rgba(15,23,42,0.16)]">
        {initial}
      </div>
      <div className="hidden min-w-0 pr-1 sm:block">
        <p className="max-w-[150px] truncate text-sm font-bold text-slate-950">{displayName}</p>
        <p className="text-[10px] font-semibold uppercase tracking-[0.18em] text-slate-400">Logged in</p>
      </div>
      <button
        aria-label="Logout"
        className="inline-flex h-9 items-center gap-1.5 rounded-full border border-slate-200 bg-white px-3 text-xs font-bold text-slate-600 transition hover:border-rose-200 hover:bg-rose-50 hover:text-rose-600 disabled:cursor-wait disabled:opacity-65"
        disabled={isLoggingOut}
        onClick={handleLogout}
        type="button"
      >
        <LogOut className="h-3.5 w-3.5" />
        <span className="hidden md:inline">{isLoggingOut ? "Keluar..." : "Logout"}</span>
      </button>
    </div>
  );
}

export function ReportingPortal({ user }: { user: SessionUser | null }) {
  return (
    <main className="relative isolate min-h-screen overflow-hidden bg-[linear-gradient(135deg,#ffffff_0%,#f5f9ff_45%,#edf5ff_100%)] px-5 py-7 text-slate-950 sm:px-8 lg:px-12">
      <div className="pointer-events-none absolute inset-0 opacity-[0.42] [background-image:linear-gradient(rgba(37,99,235,0.05)_1px,transparent_1px),linear-gradient(90deg,rgba(37,99,235,0.04)_1px,transparent_1px)] [background-size:96px_96px]" />
      <div className="pointer-events-none absolute -left-[18rem] top-28 h-[36rem] w-[54rem] rotate-[-18deg] rounded-[50%] border border-sky-200/50 bg-sky-100/20" />
      <div className="pointer-events-none absolute -right-[20rem] top-10 h-[42rem] w-[58rem] rotate-[20deg] rounded-[50%] border border-blue-200/45 bg-white/35" />
      <div className="pointer-events-none absolute bottom-[-20rem] left-1/2 h-[34rem] w-[86rem] -translate-x-1/2 rounded-[50%] bg-blue-100/35 blur-3xl" />

      <header className="relative z-20 flex items-start justify-between">
        <div className="flex h-28 w-80 items-center md:h-32 md:w-96">
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
        <UserMenu user={user} />
      </header>

      <section className="relative z-10 mx-auto flex min-h-[calc(100vh-9rem)] w-full max-w-[1400px] flex-col justify-start pb-10 pt-0">
        <div className="mx-auto max-w-3xl text-center">
          <div className="mx-auto mb-6 h-1.5 w-16 rounded-full bg-gradient-to-r from-sky-400 via-blue-500 to-violet-500" />
          <p className="text-sm font-semibold uppercase tracking-[0.38em] text-blue-300">Prime Platform</p>
          <h1 className="mt-5 text-5xl font-bold tracking-normal text-slate-950 sm:text-6xl">Welcome</h1>
          <p className="mt-4 text-lg font-medium text-slate-500">Pilih modul untuk melanjutkan</p>
        </div>

        <div className="mt-10 grid w-full gap-7 sm:grid-cols-2 xl:grid-cols-4">
          {reportingModules.map((module) => (
            <ModuleCard key={module.id} module={module} />
          ))}
        </div>
      </section>
    </main>
  );
}
