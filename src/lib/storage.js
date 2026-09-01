import { createClient } from "@supabase/supabase-js";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  // eslint-disable-next-line no-console
  console.error(
    "Missing Supabase config. Create a .env file with VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY (see .env.example)."
  );
}

const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Drop-in replacement for the Claude artifact's window.storage API.
// Same shape: get/set/delete/list, all backed by a single `app_storage` table in Supabase.
// This means the rest of the app (all the .get/.set calls) doesn't need to change at all —
// only the import at the top of App.jsx changes from `window.storage` to this `storage` object.
export const storage = {
  async get(key, shared = false) {
    const { data, error } = await supabase
      .from("app_storage")
      .select("value")
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    if (!data) throw new Error(`Key "${key}" not found`);
    return { key, value: data.value, shared };
  },

  async set(key, value, shared = false) {
    const { error } = await supabase
      .from("app_storage")
      .upsert({ key, value, shared }, { onConflict: "key" });
    if (error) {
      console.error("storage.set failed:", error);
      return null;
    }
    return { key, value, shared };
  },

  async delete(key, shared = false) {
    const { error } = await supabase.from("app_storage").delete().eq("key", key);
    if (error) {
      console.error("storage.delete failed:", error);
      return null;
    }
    return { key, deleted: true, shared };
  },

  async list(prefix = "", shared = false) {
    let query = supabase.from("app_storage").select("key");
    if (prefix) query = query.like("key", `${prefix}%`);
    const { data, error } = await query;
    if (error) {
      console.error("storage.list failed:", error);
      return null;
    }
    return { keys: (data || []).map((row) => row.key), prefix, shared };
  },
};
