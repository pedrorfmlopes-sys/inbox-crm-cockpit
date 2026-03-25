import net from "node:net";
import pg from "pg";

const DATABASE_URL = String(process.env.DATABASE_URL || "").trim();
const PG_FORCE_IPV4 = /^(1|true|yes|on)$/i.test(String(process.env.PG_FORCE_IPV4 || "").trim());
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

function isSupabasePoolerHost(hostname) {
  return hostname.includes(".pooler.supabase.");
}

function isSupabaseDirectHost(hostname) {
  return isSupabaseHost(hostname) && hostname.startsWith("db.") && !isSupabasePoolerHost(hostname);
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

function createIpv4OnlyStream() {
  const socket = new net.Socket();
  const originalConnect = socket.connect.bind(socket);

  socket.connect = function connect(port, host) {
    if (typeof port === "object" && port !== null) {
      return originalConnect({ ...port, family: 4, autoSelectFamily: false });
    }
    if (typeof port === "number" && typeof host === "string" && !host.startsWith("/")) {
      return originalConnect({ port, host, family: 4, autoSelectFamily: false });
    }
    return originalConnect(port, host);
  };

  return socket;
}

export function createOptionalPgStore(label) {
  let pool = null;
  let dbDisabledUntil = 0;
  let disabledReason = "";
  let consecutiveFailures = 0;
  let lastFailureAt = 0;
  let warnedAboutSupabaseDirect = false;
  const ssl = getSslConfig(DATABASE_URL);
  const parsedUrl = safeParseDatabaseUrl(DATABASE_URL);
  const hostname = String(parsedUrl?.hostname || "").toLowerCase();
  const port = Number(parsedUrl?.port || 5432);
  let forceIpv4 = PG_FORCE_IPV4 || isSupabasePoolerHost(hostname);

  function canUseDb() {
    return Boolean(DATABASE_URL) && Date.now() >= dbDisabledUntil;
  }

  function shouldPreferIpv4() {
    return Boolean(hostname) && !isLocalHost(hostname) && forceIpv4;
  }

  function buildPoolConfig() {
    const poolConfig = {
      connectionString: DATABASE_URL,
      ssl,
      connectionTimeoutMillis: 10_000,
      keepAlive: true,
      keepAliveInitialDelayMillis: 30_000,
    };

    if (shouldPreferIpv4()) {
      poolConfig.stream = createIpv4OnlyStream;
    }

    return poolConfig;
  }

  function maybeLogSupabaseDirectHint(error) {
    if (warnedAboutSupabaseDirect || !isSupabaseDirectHost(hostname)) return;
    warnedAboutSupabaseDirect = true;
    console.warn(
      `[${label}] Supabase direct connection detected at ${hostname}:${port}. Direct Supabase URLs are typically IPv6-only. On Render/non-IPv6 runtimes, use the Supabase Session pooler connection string in DATABASE_URL. ${error?.code || ""} ${error?.message || ""}`.trim()
    );
  }

  function createPool() {
    if (!DATABASE_URL || pool) return;
    const resolutionMode = shouldPreferIpv4() ? "IPv4-preferred" : "default DNS";
    console.log(`[${label}] Using PostgreSQL persistence via DATABASE_URL (${resolutionMode}).`);
    pool = new pg.Pool(buildPoolConfig());
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
      const code = String(error?.code || "").toUpperCase();

      if (!forceIpv4 && (code === "ENETUNREACH" || code === "EHOSTUNREACH") && isSupabaseHost(hostname)) {
        console.warn(
          `[${label}] PostgreSQL connection hit ${code} for ${hostname}:${port}; retrying once with IPv4-only DNS resolution before falling back.`
        );
        forceIpv4 = true;
        await closePool();
        createPool();
        if (pool) {
          try {
            const retryResult = await pool.query(text, params);
            consecutiveFailures = 0;
            lastFailureAt = 0;
            disabledReason = "";
            return retryResult;
          } catch (retryError) {
            error = retryError;
          }
        }
      }

      if (isNetworkFailure(error)) {
        const now = Date.now();
        consecutiveFailures = (now - lastFailureAt) <= RETRYABLE_WINDOW_MS ? consecutiveFailures + 1 : 1;
        lastFailureAt = now;
        error.optionalDbFallback = true;
        error.optionalDbFailureCount = consecutiveFailures;
        maybeLogSupabaseDirectHint(error);

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
    isConfigured() {
      return Boolean(DATABASE_URL);
    },
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
