import { createClient, type SupabaseClient } from "@supabase/supabase-js";

// The publishable key is safe to expose: all access is enforced by
// row-level-security policies in the database.
export const SUPABASE_URL =
  process.env.NEXT_PUBLIC_SUPABASE_URL ?? "https://wqgjqljupslpzklmrwqm.supabase.co";
export const SUPABASE_KEY =
  process.env.NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY ??
  "sb_publishable_eglpkiylAZ5O1d-YZSzDcQ_dShoQtws";

const BUCKET = "gazette";

/** Per-request client for server components (no session persistence). */
export function getServerSupabase(): SupabaseClient {
  return createClient(SUPABASE_URL, SUPABASE_KEY, {
    auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false },
  });
}

let browserClient: SupabaseClient | undefined;

/** Shared browser client; keeps the admin session in localStorage. */
export function getBrowserSupabase(): SupabaseClient {
  if (!browserClient) {
    browserClient = createClient(SUPABASE_URL, SUPABASE_KEY);
  }
  return browserClient;
}

/** Public CDN URL for an object in the gazette bucket. */
export function publicUrl(path: string): string {
  return `${SUPABASE_URL}/storage/v1/object/public/${BUCKET}/${path}`;
}

export const GAZETTE_BUCKET = BUCKET;
