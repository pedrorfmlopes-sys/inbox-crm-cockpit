import pg from "pg";

const DATABASE_URL = String(process.env.DATABASE_URL || "").trim();
const NETWORK_ERROR_CODES = new Set([
  "ENETUNREACH",
  "EHOSTUNREACH",
  "ECONNREFUSED",
  "ETIMEDOUT",
  "EAI_AGAIN",
]);
const MAX_CONSECUTIVE_FAILURES = 3;
const DISABLE_COOLDOWN_MS = 60_000;
const RETRYABLE_WINDOW_MS = 30_000;

function safeParseDatabaseUrl(url) {
  try {
    return new URL(url);
  } catch {
    return null;
  }
}

function isLocalHost(hostname) {
  return hostname === "localhost" || hostname === "127.0.0.1" || hostname === "::1";
}

function isSupabaseHost(hostname) {
  return hostname.endsWith(".supabase.co") || hostname.endsWith(".supabase.com") || hostname === "supabase.co" || hostname === "supabase.com";
}

function getSslConfig(connectionString) {
  const parsed = safeParseDatabaseUrl(connectionString);
  const hostname = String(parsed?.hostname || "").toLowerCase();
  const sslmode = String(parsed?.searchParams?.get("sslmode") || "").toLowerCase();

  if (sslmode === "disable") {
    return false;
  }

  if (sslmode === "verify-full" || sslmode === "verify-ca") {
    return { rejectUnauthorized: true };
  }

  if (sslmode === "require" || sslmode === "prefer") {
    return { rejectUnauthorized: false };
  }

  if (isSupabaseHost(hostname)) {
    return { rejectUnauthorized: false };
  }

  if (hostname && !isLocalHost(hostname)) {
    return { rejectUnauthorized: false };
  }

  return false;
}

function isNetworkFailure(error) {
  const code = String(error?.code || "").toUpperCase();
  if (NETWORK_ERROR_CODES.has(code)) return true;
  const message = String(error?.message || "");
  return /ENETUNREACH|EHOSTUNREACH|ECONNREFUSED|ETIMEDOUT|EAI_AGAIN|connect timeout/i.test(message);
}

export function createOptionalPgStore(label) {
  let pool = null;
  let dbDisabledUntil = 0;
  let disabledReason = "";
  let consecutiveFailures = 0;
  let lastFailureAt = 0;
  const ssl = getSslConfig(DATABASE_URL);

  function canUseDb() {
    return Boolean(DATABASE_URL) && Date.now() >= dbDisabledUntil;
  }

  function createPool() {
    if (!DATABASE_URL || pool) return;
    console.log(`[${label}] Using PostgreSQL persistence via DATABASE_URL.`);
    pool = new pg.Pool({
      connectionString: DATABASE_URL,
      ssl,
    });
  }

  async function closePool() {
    if (!pool) return;
    const currentPool = pool;
    pool = null;
    try {
      await currentPool.end();
    } catch {
      // Ignore shutdown errors for optional persistence.
    }
  }

  async function temporarilyDisable(reason) {
    disabledReason = reason;
    dbDisabledUntil = Date.now() + DISABLE_COOLDOWN_MS;
    console.warn(
      `[${label}] PostgreSQL temporarily disabled for ${Math.round(DISABLE_COOLDOWN_MS / 1000)}s; falling back to local persistence. ${reason}`
    );
    await closePool();
  }

  if (DATABASE_URL) {
    createPool();
  } else {
    console.log(`[${label}] Using local JSON file persistence.`);
  }

  async function query(text, params = []) {
    if (!canUseDb()) return null;
    if (!pool) createPool();
    if (!pool) return null;

    try {
      const result = await pool.query(text, params);
      consecutiveFailures = 0;
      lastFailureAt = 0;
      disabledReason = "";
      return result;
    } catch (error) {
      if (isNetworkFailure(error)) {
        const now = Date.now();
        consecutiveFailures = (now - lastFailureAt) <= RETRYABLE_WINDOW_MS ? consecutiveFailures + 1 : 1;
        lastFailureAt = now;
        error.optionalDbFallback = true;
        error.optionalDbFailureCount = consecutiveFailures;

        if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
          await temporarilyDisable(`${error.code || "NETWORK_ERROR"} ${error.message || ""}`.trim());
        } else {
          console.warn(
            `[${label}] PostgreSQL network failure (${consecutiveFailures}/${MAX_CONSECUTIVE_FAILURES}); keeping DATABASE_URL active and falling back for this operation only. ${error.code || ""} ${error.message || ""}`.trim()
          );
        }
      }
      throw error;
    }
  }

  return {
    query,
    isEnabled() {
      return canUseDb();
    },
    getStatus() {
      if (!DATABASE_URL) return "file";
      if (!canUseDb()) return `file:fallback:${disabledReason || "cooldown"}`;
      if (pool) return "postgres";
      return "postgres:lazy";
    },
  };
}
