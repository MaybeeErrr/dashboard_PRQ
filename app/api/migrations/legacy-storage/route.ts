import { NextRequest, NextResponse } from "next/server";
import { getSessionUser, hasPermission, SESSION_COOKIE } from "@/features/hsi/backend/auth/session";
import {
  migrateLegacyStorageSnapshot,
  normalizeExistingHsiWmsLegacyEntries,
  type LegacyStorageSnapshot,
} from "@/features/shared/backend/migrations/legacy-storage-migration";

export const dynamic = "force-dynamic";

async function authorizedMigrationUser(request: NextRequest) {
  const user = await getSessionUser(request.cookies.get(SESSION_COOKIE)?.value);
  return hasPermission(user, "MASTER_ACCESS");
}

function isSnapshot(value: unknown): value is LegacyStorageSnapshot {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<LegacyStorageSnapshot>;
  return (
    candidate.source === "legacy-localStorage" &&
    typeof candidate.createdAt === "string" &&
    typeof candidate.migrationVersion === "number" &&
    Array.isArray(candidate.keys) &&
    candidate.keys.every(
      (entry) =>
        entry &&
        typeof entry === "object" &&
        (entry as { area?: unknown }).area &&
        ((entry as { area?: unknown }).area === "localStorage" || (entry as { area?: unknown }).area === "sessionStorage") &&
        typeof (entry as { key?: unknown }).key === "string" &&
        typeof (entry as { value?: unknown }).value === "string",
    )
  );
}

export async function POST(request: NextRequest) {
  if (!(await authorizedMigrationUser(request))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const snapshot = await request.json().catch(() => null);
  if (!isSnapshot(snapshot)) {
    return NextResponse.json({ error: "Snapshot localStorage tidak valid." }, { status: 400 });
  }

  const report = await migrateLegacyStorageSnapshot(snapshot);
  return NextResponse.json(report, { status: report.errors.length > 0 ? 207 : 200 });
}

export async function PUT(request: NextRequest) {
  if (!(await authorizedMigrationUser(request))) {
    return NextResponse.json({ error: "Unauthorized" }, { status: 401 });
  }

  const report = await normalizeExistingHsiWmsLegacyEntries();
  return NextResponse.json(report, { status: report.errors.length > 0 ? 207 : 200 });
}
