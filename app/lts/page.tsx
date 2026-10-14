import { cookies } from "next/headers";
import { getSessionUser, permissionsForUser, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import { LtsDashboard } from "@/features/lts/frontend/dashboard/lts-dashboard";

export const dynamic = "force-dynamic";

export default async function LtsPage() {
  const cookieStore = await cookies();
  const user = await getSessionUser(cookieStore.get(SESSION_COOKIE)?.value);
  return <LtsDashboard permissions={permissionsForUser(user)} />;
}
