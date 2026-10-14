import { NextRequest, NextResponse } from "next/server";
import { getSessionUser, hasPermission, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";

const PUBLIC_PATHS = [
  "/login",
  "/api/login",
  "/api/logout",
  "/api/telegram/webhook",
  "/favicon.svg",
  "/logo-telkom-latest.png",
  "/module-covers",
];
const MASTER_PATHS = ["/master", "/hsi/master", "/wms/master", "/lts/master"];

function getRequestOrigin(request: NextRequest) {
  const host = request.headers.get("x-forwarded-host") ?? request.headers.get("host") ?? request.nextUrl.host;
  const protocol = request.headers.get("x-forwarded-proto") ?? request.nextUrl.protocol.replace(":", "");
  return `${protocol}://${host}`;
}

export async function proxy(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const isPublic =
    PUBLIC_PATHS.some((path) => pathname === path || pathname.startsWith(`${path}/`)) ||
    pathname.startsWith("/_next/") ||
    pathname.startsWith("/assets/");

  const user = await getSessionUser(request.cookies.get(SESSION_COOKIE)?.value);
  const authenticated = Boolean(user);

  if (!isPublic && !authenticated) {
    const loginUrl = new URL("/login", getRequestOrigin(request));
    loginUrl.searchParams.set("returnTo", `${pathname}${request.nextUrl.search}`);
    return NextResponse.redirect(loginUrl);
  }

  const masterPath = MASTER_PATHS.some((path) => pathname === path || pathname.startsWith(`${path}/`));
  if (masterPath && !hasPermission(user, "MASTER_ACCESS")) {
    return NextResponse.redirect(new URL("/portal", getRequestOrigin(request)));
  }

  return NextResponse.next();
}

export const config = {
  matcher: ["/((?!_next/static|_next/image|favicon.ico).*)"],
};
