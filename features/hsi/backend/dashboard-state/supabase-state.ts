import { readDashboardStateByKey, writeDashboardStateByKey } from "@/features/shared/backend/supabase/dashboard-state";

const STATE_KEY = "default";

export function readDashboardState() {
  return readDashboardStateByKey(STATE_KEY);
}

export function writeDashboardState(payload: unknown) {
  return writeDashboardStateByKey(STATE_KEY, payload);
}
