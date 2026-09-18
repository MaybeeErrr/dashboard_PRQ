import { NextRequest, NextResponse } from "next/server";
import {
  createSessionToken,
  credentialsConfigured,
  SESSION_COOKIE,
  SESSION_MAX_AGE_SECONDS,
  validateCredentials,
} from "@/lib/auth-session";

type LoginPayload = {
  password: string;
  returnTo: string;
  username: string;
  wantsJson: boolean;
};

function safeReturnTo(value: string | null | undefined) {
  if (!value || !value.startsWith("/") || value.startsWith("//")) return "/";
  return value;
}

function getRequestOrigin(request: NextRequest) {
  const origin = request.headers.get("origin");
  if (origin) return origin;

  const referer = request.headers.get("referer");
  if (referer) return new URL(referer).origin;

  const host = request.headers.get("x-forwarded-host") ?? request.headers.get("host") ?? new URL(request.url).host;
  const protocol = request.headers.get("x-forwarded-proto") ?? new URL(request.url).protocol.replace(":", "");
  return `${protocol}://${host}`;
}

function redirectToLogin(request: NextRequest, error: string, returnTo: string) {
  const loginUrl = new URL("/login", getRequestOrigin(request));
  loginUrl.searchParams.set("error", error);
  const destination = safeReturnTo(returnTo);
  if (destination !== "/") loginUrl.searchParams.set("returnTo", destination);
  return NextResponse.redirect(loginUrl, 303);
}

async function readLoginPayload(request: NextRequest): Promise<LoginPayload> {
  const contentType = request.headers.get("content-type") ?? "";
  const wantsJson = contentType.includes("application/json");

  if (wantsJson) {
    const body = (await request.json().catch(() => null)) as {
      password?: string;
      returnTo?: string;
      username?: string;
    } | null;
    return {
      password: body?.password ?? "",
      returnTo: safeReturnTo(body?.returnTo),
      username: body?.username?.trim() ?? "",
      wantsJson,
    };
  }

  const formData = await request.formData();
  return {
    password: String(formData.get("password") ?? ""),
    returnTo: safeReturnTo(String(formData.get("returnTo") ?? "")),
    username: String(formData.get("username") ?? "").trim(),
    wantsJson,
  };
}

export async function POST(request: NextRequest) {
  const payload = await readLoginPayload(request);

  if (!credentialsConfigured()) {
    if (!payload.wantsJson) return redirectToLogin(request, "config", payload.returnTo);
    return NextResponse.json({ error: "Credential dashboard belum dikonfigurasi." }, { status: 500 });
  }

  if (!(await validateCredentials(payload.username, payload.password))) {
    if (!payload.wantsJson) return redirectToLogin(request, "credentials", payload.returnTo);
    return NextResponse.json({ error: "Username atau password tidak sesuai." }, { status: 401 });
  }

  const response = payload.wantsJson
    ? NextResponse.json({ ok: true })
    : NextResponse.redirect(new URL(payload.returnTo, getRequestOrigin(request)), 303);
  response.cookies.set(SESSION_COOKIE, await createSessionToken(payload.username), {
    httpOnly: true,
    maxAge: SESSION_MAX_AGE_SECONDS,
    path: "/",
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
  });
  return response;
}
