import type { Metadata } from "next";
import { SessionTabGuard } from "@/features/shared/dashboard/session-tab-guard";
import "./globals.css";

export const metadata: Metadata = {
  title: "Reporting Portal",
  description: "Integrated reporting portal for HSI, WMS, and LTS dashboards.",
  other: {
    "codex-preview": "development",
  },
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="id">
      <body className="antialiased">
        <SessionTabGuard />
        {children}
      </body>
    </html>
  );
}
