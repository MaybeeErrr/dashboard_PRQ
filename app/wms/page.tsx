import { cookies } from "next/headers";
import { getSessionUser, permissionsForUser, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import { WmsDashboard } from "@/features/wms/frontend/dashboard/wms-dashboard";

export const dynamic = "force-dynamic";

export default async function WmsPage() {
  const cookieStore = await cookies();
  const user = await getSessionUser(cookieStore.get(SESSION_COOKIE)?.value);
  return <WmsDashboard permissions={permissionsForUser(user)} />;
}
