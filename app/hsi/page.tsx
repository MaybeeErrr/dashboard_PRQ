import { cookies } from "next/headers";
import { getSessionUser, permissionsForUser, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import { HsiDashboard } from "@/features/hsi/frontend/dashboard/hsi-dashboard";

export const dynamic = "force-dynamic";

export default async function HsiPage() {
  const cookieStore = await cookies();
  const user = await getSessionUser(cookieStore.get(SESSION_COOKIE)?.value);
  return <HsiDashboard permissions={permissionsForUser(user)} />;
}
