export type ReportingModuleId = "hsi" | "wms" | "lts" | "metabright";

export type ReportingModule = {
  accent: string;
  description: string;
  href: string;
  id: ReportingModuleId;
  initials: string;
  name: string;
  status: "ready" | "setup";
  summary: string;
};

export const reportingModules: ReportingModule[] = [
  {
    accent: "from-rose-500 via-amber-400 to-emerald-400",
    description: "Dashboard existing untuk MASTER, KPI, chart, Report Target RKAP, dan Report Target STF.",
    href: "/hsi",
    id: "hsi",
    initials: "HSI",
    name: "HSI Reporting",
    status: "ready",
    summary: "Existing module",
  },
  {
    accent: "from-sky-500 via-cyan-400 to-emerald-400",
    description: "Scaffold reporting WMS siap dihubungkan dengan file dan perhitungan khusus WMS.",
    href: "/wms",
    id: "wms",
    initials: "WMS",
    name: "WMS Reporting",
    status: "setup",
    summary: "Configuration pending",
  },
  {
    accent: "from-slate-500 via-indigo-400 to-teal-400",
    description: "Scaffold reporting LTS siap dihubungkan dengan file dan perhitungan khusus LTS.",
    href: "/lts",
    id: "lts",
    initials: "LTS",
    name: "LTS Reporting",
    status: "setup",
    summary: "Configuration pending",
  },
  {
    accent: "from-cyan-400 via-blue-500 to-violet-500",
    description: "Impactful Award Telda untuk monitoring pencapaian indikator MetaBright.",
    href: "/metabright",
    id: "metabright",
    initials: "METABRIGHT",
    name: "Metabright",
    status: "ready",
    summary: "Impactful Award",
  },
];
