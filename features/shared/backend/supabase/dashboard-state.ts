import { supabaseGet, supabaseUpsert } from "@/features/shared/backend/supabase/rest";

export type DashboardStateResult =
  | { body: Record<string, unknown>; status?: number }
  | { body: { configured: boolean; state: unknown; updatedAt?: string | null }; status?: number };

type DashboardStateRow = {
  payload?: unknown;
  updated_at?: string;
};

export async function readDashboardStateByKey(stateKey: string): Promise<DashboardStateResult> {
  const result = await supabaseGet<DashboardStateRow[]>(
    process.env.SUPABASE_DASHBOARD_TABLE || "dashboard_state",
    `select=payload,updated_at&key=eq.${encodeURIComponent(stateKey)}&limit=1`,
  );

  if (result.status === 503) return { body: { configured: false, state: null } };
  if (result.error) {
    return {
      body: { configured: true, error: "Gagal membaca state dashboard dari Supabase." },
      status: result.status,
    };
  }

  return {
    body: {
      configured: true,
      state: result.data?.[0]?.payload ?? null,
      updatedAt: result.data?.[0]?.updated_at ?? null,
    },
  };
}

export async function writeDashboardStateByKey(stateKey: string, payload: unknown): Promise<DashboardStateResult> {
  if (!payload || typeof payload !== "object") {
    return { body: { error: "Payload dashboard tidak valid." }, status: 400 };
  }

  const result = await supabaseUpsert(
    process.env.SUPABASE_DASHBOARD_TABLE || "dashboard_state",
    {
      key: stateKey,
      payload,
      updated_at: new Date().toISOString(),
    },
    "key",
  );

  if (result.status === 503) {
    return {
      body: { configured: false, error: "Supabase belum dikonfigurasi di Vercel." },
      status: 503,
    };
  }

  if (result.error) {
    return {
      body: { configured: true, error: "Gagal menyimpan state dashboard ke Supabase." },
      status: result.status,
    };
  }

  return { body: { configured: true, ok: true } };
}
