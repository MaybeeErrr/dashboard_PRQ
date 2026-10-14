import { redirect } from "next/navigation";
import { cookies } from "next/headers";
import { getSessionUser, hasPermission, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import { LegacyStorageMigrationPanel } from "@/features/shared/storage/legacy-storage-migration-panel";

export default async function MigrationPage() {
  const cookieStore = await cookies();
  const user = await getSessionUser(cookieStore.get(SESSION_COOKIE)?.value);

  if (!hasPermission(user, "MASTER_ACCESS")) {
    redirect("/login?returnTo=%2Fmigration");
  }

  return <LegacyStorageMigrationPanel />;
}
