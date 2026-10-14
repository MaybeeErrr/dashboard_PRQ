import { readDashboardStateByKey, writeDashboardStateByKey } from "@/features/shared/backend/supabase/dashboard-state";

const STATE_KEY = "wms";

export function readWmsDashboardState() {
  return readDashboardStateByKey(STATE_KEY);
}

export function writeWmsDashboardState(payload: unknown) {
  return writeDashboardStateByKey(STATE_KEY, payload);
}
