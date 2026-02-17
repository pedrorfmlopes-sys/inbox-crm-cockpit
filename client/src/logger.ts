export type LogLevel = "debug" | "info" | "warn" | "error";

export type ClientLogEntry = {
  ts: string;
  level: LogLevel;
  message: string;
  data?: any;
};

const MAX = 200;
const KEY = "icc_client_logs_v1";

function nowIso() {
  return new Date().toISOString();
}

function safeStringify(v: any) {
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
}

function pushEntry(entry: ClientLogEntry) {
  // persist (best-effort)
  try {
    const raw = localStorage.getItem(KEY);
    const arr: ClientLogEntry[] = raw ? JSON.parse(raw) : [];
    arr.push(entry);
    while (arr.length > MAX) arr.shift();
    localStorage.setItem(KEY, safeStringify(arr));
  } catch {
    // ignore
  }

  // notify UI listeners (best-effort)
  try {
    window.dispatchEvent(new CustomEvent("icc:log", { detail: entry }));
  } catch {
    // ignore
  }
}

function logImpl(level: LogLevel, message: string, data?: any) {
  const entry: ClientLogEntry = { ts: nowIso(), level, message, data };
  pushEntry(entry);

  // also mirror to console for debugging
  try {
    const fn =
      level === "error" ? console.error :
      level === "warn" ? console.warn :
      level === "debug" ? console.debug :
      console.log;
    fn(message, data ?? "");
  } catch {
    // ignore
  }
}

export function getClientLogs(): ClientLogEntry[] {
  try {
    const raw = localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

export function clearClientLogs() {
  try {
    localStorage.removeItem(KEY);
  } catch {
    // ignore
  }
}

/**
 * clientLog is both:
 *  - a function: clientLog(level, message, data?)
 *  - an object with helpers: clientLog.log/warn/error/debug
 */
export type ClientLogFn = ((level: LogLevel, message: string, data?: any) => void) & {
  log: (message: string, data?: any) => void;
  warn: (message: string, data?: any) => void;
  error: (message: string, data?: any) => void;
  debug: (message: string, data?: any) => void;
};

export const clientLog: ClientLogFn = (function (level: LogLevel, message: string, data?: any) {
  logImpl(level, message, data);
} as any);

clientLog.log = (message: string, data?: any) => logImpl("info", message, data);
clientLog.warn = (message: string, data?: any) => logImpl("warn", message, data);
clientLog.error = (message: string, data?: any) => logImpl("error", message, data);
clientLog.debug = (message: string, data?: any) => logImpl("debug", message, data);
