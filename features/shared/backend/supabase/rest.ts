type SupabaseConfig = {
  key: string;
  table: string;
  url: string;
};

export type SupabaseResult<T> =
  | { data: T; error: null; status: number }
  | { data: null; error: string; status: number };

export function supabaseConfig(table = process.env.SUPABASE_DASHBOARD_TABLE || "dashboard_state"): SupabaseConfig | null {
  const url = process.env.SUPABASE_URL?.replace(/\/$/, "");
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY;
  if (!url || !key) return null;
  return { key, table, url };
}

export function supabaseHeaders(key: string) {
  return {
    apikey: key,
    Authorization: `Bearer ${key}`,
    "Content-Type": "application/json",
  };
}

export async function supabaseGet<T>(table: string, query: string): Promise<SupabaseResult<T>> {
  const config = supabaseConfig(table);
  if (!config) return { data: null, error: "Supabase belum dikonfigurasi.", status: 503 };

  const response = await fetch(`${config.url}/rest/v1/${config.table}?${query}`, {
    cache: "no-store",
    headers: supabaseHeaders(config.key),
  });

  if (!response.ok) {
    return { data: null, error: await response.text(), status: response.status };
  }

  return { data: (await response.json()) as T, error: null, status: response.status };
}

export async function supabaseUpsert<T>(
  table: string,
  payload: unknown,
  onConflict: string,
  options: { returning?: "minimal" | "representation" } = {},
): Promise<SupabaseResult<T>> {
  const config = supabaseConfig(table);
  if (!config) return { data: null, error: "Supabase belum dikonfigurasi.", status: 503 };

  const response = await fetch(`${config.url}/rest/v1/${config.table}?on_conflict=${encodeURIComponent(onConflict)}`, {
    body: JSON.stringify(payload),
    headers: {
      ...supabaseHeaders(config.key),
      Prefer: `resolution=merge-duplicates,return=${options.returning ?? "minimal"}`,
    },
    method: "POST",
  });

  if (!response.ok) {
    return { data: null, error: await response.text(), status: response.status };
  }

  if (options.returning === "representation") {
    return { data: (await response.json()) as T, error: null, status: response.status };
  }

  return { data: null as T, error: null, status: response.status };
}
