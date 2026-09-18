import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Reporting HSI 2026",
  description: "Dashboard, MASTER input, Report Sementara, and Report HSI Target STF SEPT for HSI reporting.",
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
      <body className="antialiased">{children}</body>
    </html>
  );
}
