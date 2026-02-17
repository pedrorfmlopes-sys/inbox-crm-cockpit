// client/src/settings.ts
// Settings storage (RoamingSettings preferred; localStorage fallback for non-Office contexts)

import type { AiTone } from "./ai/aiClient";

export type AppLocale = "pt-PT" | "es-ES" | "en-GB" | "it-IT" | "de-DE";
export type LangOption = AppLocale | "auto";
export type ReplyLength = "xs" | "s" | "m" | "l";
export type SkinId = "classic" | "mailmaestro";

export type CockpitSettingsV1 = {
  version: 1;

  // UI skin/theme
  skinId: SkinId;

  // UI language (labels, future i18n)
  appLanguage: AppLocale;

  // Used for summaries/quick replies detection; "auto" tries to infer from email
  readingLanguage: LangOption;

  // Output language (reply/summary). "auto" defaults to readingLanguage
  replyLanguage: LangOption;

  // Default tone for AI
  tone: AiTone;

  // Default length for reply generation
  length: ReplyLength;

  // Which languages are shown in the quick language picker (bottom bar).
  // If empty/undefined, we fall back to all supported languages.
  enabledLanguages?: AppLocale[];

  // Optional signature blocks per language
  signatures: Partial<Record<AppLocale, string>>;

  // Optional signature blocks in HTML per language (preferred)
  signaturesHtml?: Partial<Record<AppLocale, string>>;

  // Signature image (URL) + max width per language
  // NOTE: dataURL of uploaded image is stored locally via helper functions (NOT roaming settings).
  signatureImageUrl?: Partial<Record<AppLocale, string>>;
  signatureImageMaxWidth?: Partial<Record<AppLocale, number>>;

  // Freeform notes/instructions that the AI should always consider
  aiKnowledge: string[];
};

const KEY_API_BASE = "apiBaseUrl";
const KEY_SETTINGS = "cockpitSettingsV1";

// Local-only keys for uploaded signature images (dataURL)
// Stored outside roaming settings to avoid size limits.
const KEY_SIGIMG_DATA_PREFIX = "icc.sigimg.data.v1:";

const DEFAULT_SETTINGS: CockpitSettingsV1 = {
  version: 1,
  skinId: "classic",
  appLanguage: "pt-PT",
  readingLanguage: "auto",
  replyLanguage: "auto",
  tone: "neutro",
  length: "m",
  enabledLanguages: ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"],
  signatures: {
    "pt-PT": "",
    "es-ES": "",
    "en-GB": "",
    "it-IT": "",
    "de-DE": "",
  },
  signaturesHtml: {
    "pt-PT": "",
    "es-ES": "",
    "en-GB": "",
    "it-IT": "",
    "de-DE": "",
  },
  signatureImageUrl: {
    "pt-PT": "",
    "es-ES": "",
    "en-GB": "",
    "it-IT": "",
    "de-DE": "",
  },
  signatureImageMaxWidth: {
    "pt-PT": 260,
    "es-ES": 260,
    "en-GB": 260,
    "it-IT": 260,
    "de-DE": 260,
  },
  aiKnowledge: [],
};

function hasOffice(): boolean {
  return typeof (globalThis as any).Office !== "undefined";
}

async function officeReady(): Promise<void> {
  if (!hasOffice()) return;
  await new Promise<void>((resolve) => {
    // @ts-ignore Office é global (office.js)
    Office.onReady(() => resolve());
  });
}

function getRoamingSettings(): any | null {
  try {
    // @ts-ignore Office global
    return Office?.context?.roamingSettings || null;
  } catch {
    return null;
  }
}

async function saveRoamingSettings(rs: any): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    rs.saveAsync((asyncResult: any) => {
      if (asyncResult.status === "succeeded") resolve();
      else reject(asyncResult.error?.message || "Falha ao guardar settings");
    });
  });
}

function safeJsonParse<T>(value: any): T | null {
  if (typeof value !== "string") return null;
  try {
    return JSON.parse(value) as T;
  } catch {
    return null;
  }
}

function mergeSettings(base: CockpitSettingsV1, incoming: Partial<CockpitSettingsV1> | null): CockpitSettingsV1 {
  if (!incoming) return base;

  const merged: CockpitSettingsV1 = {
    ...base,
    ...incoming,
    signatures: { ...base.signatures, ...(incoming.signatures || {}) },
    signaturesHtml: { ...(base.signaturesHtml || {}), ...((incoming as any).signaturesHtml || {}) },
    signatureImageUrl: { ...(base.signatureImageUrl || {}), ...((incoming as any).signatureImageUrl || {}) },
    signatureImageMaxWidth: { ...(base.signatureImageMaxWidth || {}), ...((incoming as any).signatureImageMaxWidth || {}) },
    aiKnowledge: Array.isArray(incoming.aiKnowledge) ? incoming.aiKnowledge : base.aiKnowledge,
  };

  // guard against wrong versions
  merged.version = 1;
  return merged;
}

export async function getSettings(): Promise<CockpitSettingsV1> {
  await officeReady();

  const rs = getRoamingSettings();
  if (rs) {
    const raw = rs.get(KEY_SETTINGS);
    const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
    return mergeSettings(DEFAULT_SETTINGS, parsed);
  }

  // fallback (dev / non-office)
  const raw = globalThis.localStorage?.getItem(KEY_SETTINGS);
  const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
  return mergeSettings(DEFAULT_SETTINGS, parsed);
}

export async function saveSettings(patch: Partial<CockpitSettingsV1>): Promise<CockpitSettingsV1> {
  await officeReady();
  const current = await getSettings();
  const next = mergeSettings(current, patch);
  const json = JSON.stringify(next);

  const rs = getRoamingSettings();
  if (rs) {
    rs.set(KEY_SETTINGS, json);
    await saveRoamingSettings(rs);
    return next;
  }

  globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  return next;
}

export async function resetSettings(): Promise<CockpitSettingsV1> {
  await officeReady();
  const rs = getRoamingSettings();
  const json = JSON.stringify(DEFAULT_SETTINGS);
  if (rs) {
    rs.set(KEY_SETTINGS, json);
    await saveRoamingSettings(rs);
    return DEFAULT_SETTINGS;
  }
  globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  return DEFAULT_SETTINGS;
}

// ---------------------------
// Signature Image (local-only) helpers
// ---------------------------

function sigImgKey(loc: AppLocale): string {
  return `${KEY_SIGIMG_DATA_PREFIX}${loc}`;
}

// Returns the stored dataURL for uploaded signature image (per language).
export function getSignatureImageDataUrl(loc: AppLocale): string {
  try {
    return globalThis.localStorage?.getItem(sigImgKey(loc)) || "";
  } catch {
    return "";
  }
}

// Stores a dataURL (from upload) for signature image (per language).
// Pass empty string to clear.
export function setSignatureImageDataUrl(loc: AppLocale, dataUrl: string): void {
  try {
    const v = String(dataUrl || "").trim();
    if (!v) globalThis.localStorage?.removeItem(sigImgKey(loc));
    else globalThis.localStorage?.setItem(sigImgKey(loc), v);
  } catch {
    // ignore
  }
}

export function clearSignatureImageDataUrl(loc: AppLocale): void {
  try {
    globalThis.localStorage?.removeItem(sigImgKey(loc));
  } catch {
    // ignore
  }
}

// ---------------------------
// Existing API base helpers
// ---------------------------

export async function getApiBaseUrl(): Promise<string> {
  await officeReady();
  const rs = getRoamingSettings();
  const v = rs ? rs.get(KEY_API_BASE) : globalThis.localStorage?.getItem(KEY_API_BASE);
  if (typeof v === "string" && v.trim()) return v.trim();
  return "http://localhost:7071"; // default DEV
}

export async function setApiBaseUrl(url: string): Promise<void> {
  await officeReady();
  const u = url.trim();
  const rs = getRoamingSettings();

  if (rs) {
    rs.set(KEY_API_BASE, u);
    await saveRoamingSettings(rs);
    return;
  }

  globalThis.localStorage?.setItem(KEY_API_BASE, u);
}
