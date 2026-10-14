import { ReportingPortal } from "@/features/portal/frontend/reporting-portal";
import { getSessionUser, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import { cookies } from "next/headers";

export default async function PortalPage() {
  const cookieStore = await cookies();
  const user = await getSessionUser(cookieStore.get(SESSION_COOKIE)?.value);

  return <ReportingPortal user={user} />;
}
