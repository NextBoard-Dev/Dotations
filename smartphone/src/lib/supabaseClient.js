import { createClient } from "@supabase/supabase-js";

const DEFAULT_SUPABASE_URL = "https://dphrvdhqhgycmllietuk.supabase.co";
const DEFAULT_SUPABASE_PUBLISHABLE_KEY = "sb_publishable_2wYXnIDj4-c8daQZW8D5hA_2Py6k7z6";

const supabaseUrl = String(import.meta.env.VITE_SUPABASE_URL || DEFAULT_SUPABASE_URL).trim();
const supabaseAnonKey = String(
  import.meta.env.VITE_SUPABASE_ANON_KEY || DEFAULT_SUPABASE_PUBLISHABLE_KEY
).trim();

if (!supabaseUrl || !supabaseAnonKey) {
  throw new Error("Supabase non configure: VITE_SUPABASE_URL / VITE_SUPABASE_ANON_KEY manquants");
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);
