"use client";

import { useEffect } from "react";
import { usePathname, useRouter } from "next/navigation";

export const TAB_SESSION_KEY = "prime_session_active";

export function SessionTabGuard() {
  const pathname = usePathname();
  const router = useRouter();

  useEffect(() => {
    if (pathname === "/login") return;
    if (window.sessionStorage.getItem(TAB_SESSION_KEY) === "1") return;

    const returnTo = `${window.location.pathname}${window.location.search}`;
    void fetch("/api/logout", { method: "POST" }).finally(() => {
      router.replace(`/login?returnTo=${encodeURIComponent(returnTo)}`);
      router.refresh();
    });
  }, [pathname, router]);

  return null;
}

export function LoginSessionReset() {
  useEffect(() => {
    window.sessionStorage.removeItem(TAB_SESSION_KEY);
    void fetch("/api/logout", { method: "POST" });
  }, []);

  return null;
}
