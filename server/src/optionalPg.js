import pg from "pg";

const DATABASE_URL = String(process.env.DATABASE_URL || "").trim();
const NETWORK_ERROR_CODES = new Set([
  "ENETUNREACH",
  "EHOSTUNREACH",
  "ECONNREFUSED",
  "ETIMEDOUT",
  "EAI_AGAIN",
]);

function isSupabaseUrl(url) {
  return url.includes("supabase.co");
}

function isNetworkFailure(error) {
  const code = String(error?.code || "").toUpperCase();
  if (NETWORK_ERROR_CODES.has(code)) return true;
  const message = String(error?.message || "");
  return /ENETUNREACH|EHOSTUNREACH|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN|connect timeout/i.test(message);
}

export function createOptionalPgStore(label) {
  let pool = null;
  let dbDisabled = false;
  let disabledReason = "";

  if (DATABASE_URL) {
    console.log(`[${label}] Using PostgreSQL/Supabase persistence.`);
    pool = new pg.Pool({
      connectionString: DATABASE_URL,
      ssl: isSupabaseUrl(DATABASE_URL) ? { rejectUnauthorized: false } : false,
    });
  } else {
    console.log(`[${label}] Using local JSON file persistence.`);
  }

  async function disablePool(reason) {
    if (dbDisabled) return;
    dbDisabled = true;
    disabledReason = reason;
    const currentPool = pool;
    pool = null;
    console.warn(`[${label}] PostgreSQL unavailable; falling back to local persistence. ${reason}`);
    if (currentPool) {
      try {
        await currentPool.end();
      } catch {
        // Ignore shutdown errors for optional persistence.
      }
    }
  }

  async function query(text, params = []) {
    if (!pool || dbDisabled) return null;
    try {
      return await pool.query(text, params);
    } catch (error) {
      if (isNetworkFailure(error)) {
        error.optionalDbFallback = true;
        await disablePool(`${error.code || "NETWORK_ERROR"} ${error.message || ""}`.trim());
      }
      throw error;
    }
  }

  return {
    query,
    isEnabled() {
      return Boolean(pool) && !dbDisabled;
    },
    getStatus() {
      if (pool && !dbDisabled) return "postgres";
      if (!DATABASE_URL) return "file";
      return `file:fallback:${disabledReason || "disabled"}`;
    },
  };
}
