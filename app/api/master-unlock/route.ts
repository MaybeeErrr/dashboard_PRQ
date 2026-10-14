import { NextRequest, NextResponse } from "next/server";
import {
  createMasterUnlockToken,
  getSessionUser,
  hasPermission,
  SESSION_COOKIE,
  validateMasterPassword,
} from "@/features/hsi/backend/auth/session";

export const dynamic = "force-dynamic";

export async function POST(request: NextRequest) {
  const user = await getSessionUser(request.cookies.get(SESSION_COOKIE)?.value);
  if (!hasPermission(user, "MASTER_ACCESS") || !user) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const payload = (await request.json().catch(() => null)) as { password?: unknown } | null;
  const password = typeof payload?.password === "string" ? payload.password : "";
  if (!(await validateMasterPassword(password))) {
    return NextResponse.json({ error: "Password MASTER salah." }, { status: 403 });
  }

  return NextResponse.json({ ok: true, token: await createMasterUnlockToken(user.username) });
}
