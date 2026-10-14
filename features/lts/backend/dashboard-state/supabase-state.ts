import { readDashboardStateByKey, writeDashboardStateByKey } from "@/features/shared/backend/supabase/dashboard-state";

const STATE_KEY = "lts";

export function readLtsDashboardState() {
  return readDashboardStateByKey(STATE_KEY);
}

export function writeLtsDashboardState(payload: unknown) {
  return writeDashboardStateByKey(STATE_KEY, payload);
}
