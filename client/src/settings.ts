// client/src/settings.ts
// Settings storage (RoamingSettings preferred; localStorage fallback for non-Office contexts)

import type { AiTone } from "./ai/aiClient";
import {
  DEFAULT_GROUP_STORAGE_SETTINGS,
  type GroupStorageMode,
  type GroupStorageLegacyProvider as GroupStorageProvider,
  type GroupStorageSettings,
} from "./modules/crm/groups-v1/storage/settings";
import {
  DEFAULT_GROUPS_TAB_SETTINGS,
  type GroupsTabSettings,
} from "./modules/crm/groups-v1/settings/groupsTabSettings";
import {
  DEFAULT_GROUPS_MODULE_SETTINGS,
  buildGroupsLegacyAliases,
  normalizeGroupsModuleSettings,
  type GroupLabelCatalogEntry,
  type GroupLabelStatus,
  type GroupOutlookCategorySettings,
  type GroupsModuleSettings,
  type GroupTicketAutoLinkMode,
  type GroupTicketUiSettings,
} from "./modules/crm/groups-v1/settings/groupsModuleSettings";
export { normalizeGroupStorageSettings } from "./modules/crm/groups-v1/storage/settings";
export { normalizeGroupsTabSettings } from "./modules/crm/groups-v1/settings/groupsTabSettings";
export {
  findGroupLabelCatalogEntry,
  getGroupLabelCatalogLabels,
  normalizeGroupLabelCatalog,
} from "./modules/crm/groups-v1/settings/groupsModuleSettings";

export type AppLocale = "pt-PT" | "es-ES" | "en-GB" | "it-IT" | "de-DE";
export type LangOption = AppLocale | "auto";
export type ReplyLength = "xs" | "s" | "m" | "l";
export type SkinId = "classic" | "mailmaestro" | "vibrant";

export type ResponsePreset = {
  id: string;
  name: string;
  prompt: string;
};

export type SettingsMigrations = {
  legacyResponsePresetsV1?: boolean;
};

export type ContactAlias = {
  id: string;
  name: string;
  email: string;
};

export type AiCustomTone = {
  id: string;
  name: string;
  instructions: string;
};

export type AiTextShortcut = {
  id: string;
  trigger: string;
  content: string;
};

export type AiAutoLabelId =
  | "to_respond"
  | "meeting"
  | "fyi"
  | "notification"
  | "internal_update"
  | "awaiting_reply"
  | "marketing"
  | "done";

export type AiAutoLabelSettings = {
  enabled: boolean;
  autoDraftEnabled: boolean;
  labels: Record<AiAutoLabelId, boolean>;
};

export type AiFontPreference = {
  family: string;
  size: number;
  color: string;
};

export type ReferenceEntityKey = "lead" | "project" | "task" | "ticket";
export type ReferenceCounterMode = "per_type" | "global";
export type ReferenceCodePosition = "prefix" | "suffix";
export type Crm2OdooLayoutMode = "description_only" | "structured_project";
export type Crm2OdooLayoutTarget = "project" | "lead" | "task" | "ticket";
export type { GroupStorageMode, GroupStorageProvider, GroupStorageSettings };
export type { GroupsTabSettings };
export type {
  GroupLabelCatalogEntry,
  GroupLabelStatus,
  GroupOutlookCategorySettings,
  GroupsModuleSettings,
  GroupTicketAutoLinkMode,
  GroupTicketUiSettings,
};

export type ReferenceCodeSettings = {
  enabled: boolean;
  prefixes: Record<ReferenceEntityKey, string>;
  counterMode: ReferenceCounterMode;
  includeYear: boolean;
  position: ReferenceCodePosition;
  counters: {
    global: number;
    perType: Record<ReferenceEntityKey, number>;
  };
};

export type InvoiceStudioSettings = {
  enabled: boolean;
  baseUrl: string;
  email: string;
  password: string;
  project: string;
};

export type Crm2StructuredLayoutSettings<TModel extends string> = {
  model: TModel;
  mode: Crm2OdooLayoutMode;
  descriptionField: string;
  fixedInfoField: string;
  historyField: string;
  documentsField: string;
  fixedInfoTabLabel: string;
  historyTabLabel: string;
  documentsTabLabel: string;
  fallbackToDescription: boolean;
};

export type Crm2ProjectStructuredLayoutSettings = Crm2StructuredLayoutSettings<"project.project">;
export type Crm2LeadStructuredLayoutSettings = Crm2StructuredLayoutSettings<"crm.lead">;
export type Crm2TaskStructuredLayoutSettings = Crm2StructuredLayoutSettings<"project.task">;
export type Crm2TicketStructuredLayoutSettings = Crm2StructuredLayoutSettings<"helpdesk.ticket">;

export type Crm2OdooLayoutSettings = {
  // Legacy/global default kept for backwards compatibility.
  mode: Crm2OdooLayoutMode;
  includeAnchorIndex: boolean;
  showBackToTopLinks: boolean;
  project: Crm2ProjectStructuredLayoutSettings;
  lead: Crm2LeadStructuredLayoutSettings;
  task: Crm2TaskStructuredLayoutSettings;
  ticket: Crm2TicketStructuredLayoutSettings;
};

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

  // When true, blocks automatic AI runs and keeps the app in on-demand mode.
  aiManualOnly: boolean;

  // Persona & Style Mimic
  userRole?: string;
  styleContext?: string;
  styleExamples?: string;

  // Personal Meeting Links
  meetingLinks?: {
    teams?: string;
    zoom?: string;
    meet?: string;
  };

  // Odoo Credentials (Optional, stored for persistence)
  odooUrl?: string;
  odooDb?: string;
  odooLogin?: string;
  odooPassword?: string;
  odooSessionToken?: string;

  // AI Credentials
  geminiApiKey?: string;
  openaiApiKey?: string;

  // AI Models
  openaiModelFast?: string;
  openaiModelQuality?: string;
  geminiModel?: string;
  invoiceStudio: InvoiceStudioSettings;

  // AI Context Scope
  bodyScope?: "main" | "full";

  // New: Custom Response Presets
  responsePresets: ResponsePreset[];

  // One-shot local/roaming migrations already processed.
  migrations?: SettingsMigrations;

  // New: Contact Aliases (Forwarding Shortcuts)
  contactAliases: ContactAlias[];

  // AI module settings (MailMaestro-like per-module configuration)
  aiCustomTones: AiCustomTone[];
  aiTextShortcuts: AiTextShortcut[];
  aiAutoLabel: AiAutoLabelSettings;
  aiFontPreference: AiFontPreference;

  // Configurable reference codes for Odoo-created records
  referenceCodes: ReferenceCodeSettings;

  // Canonical Groups module settings / data source
  groups: GroupsModuleSettings;

  // Legacy compatibility aliases for Groups
  groupStorage: GroupStorageSettings;
  groupsTabSettings: GroupsTabSettings;
  groupLabelsManagerEnabled: boolean;
  groupLabelCatalog: GroupLabelCatalogEntry[];
  groupFavoriteIds: string[];
  groupTicketsEnabled: boolean;
  groupTicketUi: GroupTicketUiSettings;
  groupOutlookCategories: GroupOutlookCategorySettings;

  // CRM2 Odoo layout strategy for multi-company deployments
  crm2OdooLayout: Crm2OdooLayoutSettings;
};

const KEY_API_BASE = "apiBaseUrl";
const KEY_SETTINGS = "cockpitSettingsV1";
export const SETTINGS_UPDATED_EVENT = "iccc:settings-updated";

const SETTINGS_STORAGE_KEYS: Array<keyof CockpitSettingsV1> = [
  "version",
  "skinId",
  "appLanguage",
  "readingLanguage",
  "replyLanguage",
  "tone",
  "length",
  "enabledLanguages",
  "signatures",
  "signaturesHtml",
  "signatureImageUrl",
  "signatureImageMaxWidth",
  "aiKnowledge",
  "aiManualOnly",
  "userRole",
  "styleContext",
  "styleExamples",
  "meetingLinks",
  "odooUrl",
  "odooDb",
  "odooLogin",
  "odooPassword",
  "odooSessionToken",
  "geminiApiKey",
  "openaiApiKey",
  "openaiModelFast",
  "openaiModelQuality",
  "geminiModel",
  "invoiceStudio",
  "bodyScope",
  "responsePresets",
  "migrations",
  "contactAliases",
  "aiCustomTones",
  "aiTextShortcuts",
  "aiAutoLabel",
  "aiFontPreference",
  "referenceCodes",
  "groups",
  "crm2OdooLayout",
];

function pickKnownSettings(input: Partial<CockpitSettingsV1> | null | undefined): Partial<CockpitSettingsV1> {
  if (!input || typeof input !== "object") return {};
  const compact: Partial<CockpitSettingsV1> = {};
  for (const key of SETTINGS_STORAGE_KEYS) {
    if (Object.prototype.hasOwnProperty.call(input, key)) {
      (compact as Record<string, unknown>)[key] = (input as Record<string, unknown>)[key as string];
    }
  }
  return compact;
}

function compactSettingsForStorage(settings: CockpitSettingsV1): CockpitSettingsV1 {
  const compact = pickKnownSettings(settings) as CockpitSettingsV1;
  compact.version = 1;
  return compact;
}

function serializeSettings(settings: CockpitSettingsV1): string {
  return JSON.stringify(compactSettingsForStorage(settings));
}

function writeLocalSettingsCache(json: string): void {
  try {
    globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  } catch {
    try {
      globalThis.localStorage?.removeItem(KEY_SETTINGS);
      globalThis.localStorage?.setItem(KEY_SETTINGS, json);
    } catch {
      // Office roaming settings remain the source of truth; the local mirror is best effort.
    }
  }
}

function writeLocalSettingsRequired(json: string): void {
  try {
    globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  } catch (error) {
    try {
      globalThis.localStorage?.removeItem(KEY_SETTINGS);
      globalThis.localStorage?.setItem(KEY_SETTINGS, json);
    } catch {
      throw error;
    }
  }
}

// Local-only keys for uploaded signature images (dataURL)
// Stored outside roaming settings to avoid size limits.
const KEY_SIGIMG_DATA_PREFIX = "icc.sigimg.data.v1:";
const LEGACY_SIGNATURE_KEYS = {
  mode: "icc.sig.mode",
  text: "icc.sig.text",
  html: "icc.sig.html",
  image: "icc.sig.img",
  imageWidth: "icc.sig.img.w",
} as const;
const LEGACY_RESPONSE_PRESETS_KEY = "crmCockpit.templates.v1";
const SIGNATURE_LOCALES: AppLocale[] = ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"];

const DEFAULT_RESPONSE_PRESETS: ResponsePreset[] = [
  { id: "p1", name: "Pedido de Dados", prompt: "Agradece o contacto e solicita os dados de faturação (NIF, Morada) para podermos proceder." },
  { id: "p2", name: "Agendamento Carga", prompt: "Informa que a mercadoria está pronta e solicita confirmação de data/hora para a recolha no nosso armazém." },
  { id: "p3", name: "Follow-up Proposta", prompt: "Faz um follow-up cortês sobre a última proposta enviada, perguntando se restam dúvidas técnicas." }
];

const DEFAULT_GROUPS_SETTINGS = normalizeGroupsModuleSettings(
  DEFAULT_GROUPS_MODULE_SETTINGS,
  null,
  DEFAULT_GROUPS_MODULE_SETTINGS
);
const DEFAULT_GROUPS_LEGACY_ALIASES = buildGroupsLegacyAliases(DEFAULT_GROUPS_SETTINGS);

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
  aiManualOnly: true,
  userRole: "",
  styleContext: "",
  styleExamples: "",
  meetingLinks: {
    teams: "",
    zoom: "",
    meet: "",
  },
  odooUrl: "",
  odooDb: "",
  odooLogin: "",
  odooPassword: "",
  odooSessionToken: "",
  geminiApiKey: "",
  openaiApiKey: "",
  openaiModelFast: "",
  openaiModelQuality: "",
  geminiModel: "",
  invoiceStudio: {
    enabled: false,
    baseUrl: "https://invoice-studio-backend.onrender.com",
    email: "",
    password: "",
    project: "",
  },
  bodyScope: "main",
  responsePresets: [...DEFAULT_RESPONSE_PRESETS],
  migrations: {},
  contactAliases: [
    { id: "c1", name: "Ragno", email: "info@ragno.it" },
    { id: "c2", name: "Marazzi", email: "contact@marazzi.it" }
  ],
  aiCustomTones: [],
  aiTextShortcuts: [],
  aiAutoLabel: {
    enabled: false,
    autoDraftEnabled: false,
    labels: {
      to_respond: true,
      meeting: true,
      fyi: true,
      notification: true,
      internal_update: true,
      awaiting_reply: true,
      marketing: false,
      done: true,
    },
  },
  aiFontPreference: {
    family: "Segoe UI",
    size: 12,
    color: "#172B4D",
  },
  referenceCodes: {
    enabled: false,
    prefixes: {
      lead: "",
      project: "",
      task: "",
      ticket: "",
    },
    counterMode: "per_type",
    includeYear: false,
    position: "prefix",
    counters: {
      global: 0,
      perType: {
        lead: 0,
        project: 0,
        task: 0,
        ticket: 0,
      },
    },
  },
  groups: {
    ...DEFAULT_GROUPS_SETTINGS,
  },
  // Legacy top-level aliases remain derived only for compatibility during the
  // migration window. Runtime consumers must read from settings.groups.
  groupStorage: {
    ...(DEFAULT_GROUPS_LEGACY_ALIASES.groupStorage || DEFAULT_GROUP_STORAGE_SETTINGS),
  },
  groupsTabSettings: {
    ...(DEFAULT_GROUPS_LEGACY_ALIASES.groupsTabSettings || DEFAULT_GROUPS_TAB_SETTINGS),
  },
  groupLabelsManagerEnabled: DEFAULT_GROUPS_LEGACY_ALIASES.groupLabelsManagerEnabled ?? true,
  groupLabelCatalog: Array.isArray(DEFAULT_GROUPS_LEGACY_ALIASES.groupLabelCatalog)
    ? [...DEFAULT_GROUPS_LEGACY_ALIASES.groupLabelCatalog]
    : [],
  groupFavoriteIds: Array.isArray(DEFAULT_GROUPS_LEGACY_ALIASES.groupFavoriteIds)
    ? [...DEFAULT_GROUPS_LEGACY_ALIASES.groupFavoriteIds]
    : [],
  groupTicketsEnabled: DEFAULT_GROUPS_LEGACY_ALIASES.groupTicketsEnabled ?? true,
  groupTicketUi: {
    ...(DEFAULT_GROUPS_LEGACY_ALIASES.groupTicketUi || DEFAULT_GROUPS_SETTINGS.tickets.ui),
  },
  groupOutlookCategories: {
    ...(DEFAULT_GROUPS_LEGACY_ALIASES.groupOutlookCategories || DEFAULT_GROUPS_SETTINGS.outlookCategories),
  },
  crm2OdooLayout: {
    mode: "description_only",
    includeAnchorIndex: true,
    showBackToTopLinks: true,
    project: {
      model: "project.project",
      mode: "description_only",
      descriptionField: "description",
      fixedInfoField: "x_studio_iccc_project_brief",
      historyField: "x_studio_iccc_project_history",
      documentsField: "x_studio_iccc_project_documents",
      fixedInfoTabLabel: "Informacao fixa",
      historyTabLabel: "Historico",
      documentsTabLabel: "Documentos",
      fallbackToDescription: true,
    },
    lead: {
      model: "crm.lead",
      mode: "description_only",
      descriptionField: "description",
      fixedInfoField: "x_studio_iccc_lead_brief",
      historyField: "x_studio_iccc_lead_history",
      documentsField: "x_studio_iccc_lead_documents",
      fixedInfoTabLabel: "Informacao fixa",
      historyTabLabel: "Historico",
      documentsTabLabel: "Documentos",
      fallbackToDescription: true,
    },
    task: {
      model: "project.task",
      mode: "description_only",
      descriptionField: "description",
      fixedInfoField: "x_studio_iccc_task_brief",
      historyField: "x_studio_iccc_task_history",
      documentsField: "x_studio_iccc_task_documents",
      fixedInfoTabLabel: "Informacao fixa",
      historyTabLabel: "Historico",
      documentsTabLabel: "Documentos",
      fallbackToDescription: true,
    },
    ticket: {
      model: "helpdesk.ticket",
      mode: "description_only",
      descriptionField: "description",
      fixedInfoField: "x_studio_iccc_ticket_brief",
      historyField: "x_studio_iccc_ticket_history",
      documentsField: "x_studio_iccc_ticket_documents",
      fixedInfoTabLabel: "Informacao fixa",
      historyTabLabel: "Historico",
      documentsTabLabel: "Documentos",
      fallbackToDescription: true,
    },
  },
};

function normalizeCrm2OdooLayoutMode(value: unknown): Crm2OdooLayoutMode {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "structured_project" ? "structured_project" : "description_only";
}

function hasOffice(): boolean {
  return typeof (globalThis as any).Office !== "undefined";
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForOffice(maxWaitMs = 5000): Promise<any | null> {
  const startedAt = Date.now();
  while (true) {
    const OfficeAny = (globalThis as any).Office;
    if (OfficeAny) return OfficeAny;
    if (Date.now() - startedAt >= maxWaitMs) return null;
    await sleep(50);
  }
}

async function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | null = null;
  const timeoutPromise = new Promise<T>((resolve) => {
    timer = setTimeout(() => resolve(fallback), ms);
  });
  const result = await Promise.race([promise, timeoutPromise]);
  if (timer) clearTimeout(timer);
  return result;
}

async function officeReady(): Promise<void> {
  if (!hasOffice()) return;
  const OfficeAny = await waitForOffice(5000);
  if (!OfficeAny) return;
  if (OfficeAny.context?.roamingSettings) return;

  const readyPromise = new Promise<void>((resolve) => {
    let settled = false;
    const finish = () => {
      if (settled) return;
      settled = true;
      resolve();
    };

    try {
      const maybePromise = OfficeAny.onReady?.(() => finish());
      if (maybePromise?.then) {
        maybePromise.then(() => finish()).catch(() => finish());
      }
    } catch {
      finish();
    }
  });

  await withTimeout(readyPromise, 5000, undefined);
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

function emitSettingsUpdated(settings: CockpitSettingsV1): void {
  try {
    globalThis.dispatchEvent?.(new CustomEvent(SETTINGS_UPDATED_EVENT, { detail: settings }));
  } catch {
    // ignore
  }
}

function readLocalString(key: string): string {
  try {
    return String(globalThis.localStorage?.getItem(key) || "").trim();
  } catch {
    return "";
  }
}

function removeLegacySignatureSettings(): void {
  try {
    for (const key of Object.values(LEGACY_SIGNATURE_KEYS)) {
      globalThis.localStorage?.removeItem(key);
    }
  } catch {
    // ignore
  }
}

function hasOfficialSignature(settings: CockpitSettingsV1): boolean {
  for (const loc of SIGNATURE_LOCALES) {
    if (String(settings.signatures?.[loc] || "").trim()) return true;
    if (String(settings.signaturesHtml?.[loc] || "").trim()) return true;
    if (String(settings.signatureImageUrl?.[loc] || "").trim()) return true;
    if (getSignatureImageDataUrl(loc)) return true;
  }
  return false;
}

function applyLegacySignatureMigration(settings: CockpitSettingsV1, removeLegacy = false): CockpitSettingsV1 {
  const legacyMode = readLocalString(LEGACY_SIGNATURE_KEYS.mode).toLowerCase();
  const legacyText = readLocalString(LEGACY_SIGNATURE_KEYS.text);
  const legacyHtml = readLocalString(LEGACY_SIGNATURE_KEYS.html);
  const legacyImage = readLocalString(LEGACY_SIGNATURE_KEYS.image);
  const legacyWidth = Number(readLocalString(LEGACY_SIGNATURE_KEYS.imageWidth));

  if (!legacyMode && !legacyText && !legacyHtml && !legacyImage) return settings;

  if (hasOfficialSignature(settings)) {
    if (removeLegacy) removeLegacySignatureSettings();
    return settings;
  }

  const targetLocale: AppLocale =
    settings.replyLanguage && settings.replyLanguage !== "auto"
      ? settings.replyLanguage
      : settings.appLanguage || "pt-PT";
  const next: CockpitSettingsV1 = {
    ...settings,
    signatures: { ...settings.signatures },
    signaturesHtml: { ...(settings.signaturesHtml || {}) },
    signatureImageUrl: { ...(settings.signatureImageUrl || {}) },
    signatureImageMaxWidth: { ...(settings.signatureImageMaxWidth || {}) },
  };

  if ((legacyMode === "html" || (!legacyMode && legacyHtml)) && legacyHtml) {
    next.signaturesHtml![targetLocale] = legacyHtml;
  } else if ((legacyMode === "text" || (!legacyMode && legacyText)) && legacyText) {
    next.signatures[targetLocale] = legacyText;
  } else if ((legacyMode === "image" || (!legacyMode && legacyImage)) && legacyImage) {
    if (legacyImage.startsWith("data:")) {
      setSignatureImageDataUrl(targetLocale, legacyImage);
      next.signatureImageUrl![targetLocale] = "";
    } else {
      next.signatureImageUrl![targetLocale] = legacyImage;
    }
    if (Number.isFinite(legacyWidth) && legacyWidth > 0) {
      next.signatureImageMaxWidth![targetLocale] = Math.max(80, Math.min(800, legacyWidth));
    }
  }

  if (removeLegacy) removeLegacySignatureSettings();
  return next;
}

function responsePresetKey(entry: Pick<ResponsePreset, "name" | "prompt">): string {
  return `${String(entry.name || "").trim().toLowerCase()}\n${String(entry.prompt || "").trim().toLowerCase()}`;
}

function normalizeResponsePresets(input: any[]): ResponsePreset[] {
  const usedIds = new Set<string>();
  const seen = new Set<string>();
  const next: ResponsePreset[] = [];

  for (const raw of input || []) {
    const name = String(raw?.name || "").trim();
    const prompt = String(raw?.prompt ?? raw?.body ?? "").trim();
    if (!name && !prompt) continue;

    const key = responsePresetKey({ name, prompt });
    if (seen.has(key)) continue;
    seen.add(key);

    const rawId = String(raw?.id || "").trim();
    const baseId = rawId || `preset-${next.length + 1}`;
    let id = baseId;
    let suffix = 2;
    while (usedIds.has(id)) {
      id = `${baseId}-${suffix}`;
      suffix += 1;
    }
    usedIds.add(id);

    next.push({
      id,
      name: name || "MOD sem nome",
      prompt,
    });
  }

  return next;
}

function responsePresetsAreDefaultOrEmpty(input: ResponsePreset[] | undefined): boolean {
  const presets = normalizeResponsePresets(Array.isArray(input) ? input : []);
  if (!presets.length) return true;
  const defaults = normalizeResponsePresets(DEFAULT_RESPONSE_PRESETS);
  if (presets.length !== defaults.length) return false;
  const defaultKeys = new Set(defaults.map(responsePresetKey));
  return presets.every((entry) => defaultKeys.has(responsePresetKey(entry)));
}

function readLegacyResponsePresets(): ResponsePreset[] {
  const raw = readLocalString(LEGACY_RESPONSE_PRESETS_KEY);
  const parsed = safeJsonParse<any[]>(raw);
  if (!Array.isArray(parsed)) return [];
  return normalizeResponsePresets(
    parsed.map((entry: any) => ({
      id: String(entry?.id || "").trim(),
      name: String(entry?.name || "").trim(),
      prompt: String(entry?.prompt ?? entry?.body ?? "").trim(),
    }))
  );
}

function removeLegacyResponsePresets(): void {
  try {
    globalThis.localStorage?.removeItem(LEGACY_RESPONSE_PRESETS_KEY);
  } catch {
    // ignore
  }
}

function applyLegacyResponsePresetMigration(settings: CockpitSettingsV1, removeLegacy = false): CockpitSettingsV1 {
  if (settings.migrations?.legacyResponsePresetsV1) {
    if (removeLegacy) removeLegacyResponsePresets();
    return settings;
  }

  const legacyPresets = readLegacyResponsePresets();
  if (!legacyPresets.length) return settings;

  const next: CockpitSettingsV1 = {
    ...settings,
    migrations: {
      ...(settings.migrations || {}),
      legacyResponsePresetsV1: true,
    },
  };

  if (responsePresetsAreDefaultOrEmpty(settings.responsePresets)) {
    next.responsePresets = normalizeResponsePresets([
      ...legacyPresets,
      ...DEFAULT_RESPONSE_PRESETS,
    ]);
  }

  if (removeLegacy) removeLegacyResponsePresets();
  return next;
}

function applyLegacySettingsMigrations(settings: CockpitSettingsV1, removeLegacy = false): CockpitSettingsV1 {
  return applyLegacyResponsePresetMigration(
    applyLegacySignatureMigration(settings, removeLegacy),
    removeLegacy
  );
}

function mergeSettings(base: CockpitSettingsV1, incoming: Partial<CockpitSettingsV1> | null): CockpitSettingsV1 {
  if (!incoming) return compactSettingsForStorage(applyLegacySettingsMigrations(base));
  const knownIncoming = pickKnownSettings(incoming);
  const incomingLayout = ((incoming as any).crm2OdooLayout || {});
  const incomingLayoutMode = normalizeCrm2OdooLayoutMode(incomingLayout.mode ?? base.crm2OdooLayout.mode);
  const normalizedGroups = normalizeGroupsModuleSettings(
    (incoming as any).groups || null,
    {
      groupStorage: (incoming as any).groupStorage,
      groupsTabSettings: (incoming as any).groupsTabSettings,
      groupLabelsManagerEnabled: (incoming as any).groupLabelsManagerEnabled,
      groupLabelCatalog: (incoming as any).groupLabelCatalog,
      groupFavoriteIds: (incoming as any).groupFavoriteIds,
      groupTicketsEnabled: (incoming as any).groupTicketsEnabled,
      groupTicketUi: (incoming as any).groupTicketUi,
      groupOutlookCategories: (incoming as any).groupOutlookCategories,
    },
    base.groups || DEFAULT_GROUPS_SETTINGS
  );
  const normalizedGroupAliases = buildGroupsLegacyAliases(normalizedGroups);

  const merged: CockpitSettingsV1 = {
    ...base,
    ...knownIncoming,
    signatures: { ...base.signatures, ...(incoming.signatures || {}) },
    signaturesHtml: { ...(base.signaturesHtml || {}), ...((incoming as any).signaturesHtml || {}) },
    signatureImageUrl: { ...(base.signatureImageUrl || {}), ...((incoming as any).signatureImageUrl || {}) },
    signatureImageMaxWidth: { ...(base.signatureImageMaxWidth || {}), ...((incoming as any).signatureImageMaxWidth || {}) },
    aiKnowledge: Array.isArray(incoming.aiKnowledge) ? incoming.aiKnowledge : base.aiKnowledge,
    aiManualOnly: typeof incoming.aiManualOnly === "boolean" ? incoming.aiManualOnly : base.aiManualOnly,
    responsePresets: Array.isArray((incoming as any).responsePresets)
      ? normalizeResponsePresets((incoming as any).responsePresets)
      : base.responsePresets,
    migrations: {
      ...(base.migrations || {}),
      ...(((incoming as any).migrations || {}) as SettingsMigrations),
    },
    aiCustomTones: Array.isArray((incoming as any).aiCustomTones)
      ? (incoming as any).aiCustomTones
        .map((entry: any) => ({
          id: String(entry?.id || "").trim(),
          name: String(entry?.name || "").trim(),
          instructions: String(entry?.instructions || "").trim(),
        }))
        .filter((entry: AiCustomTone) => entry.id && entry.name)
      : base.aiCustomTones,
    aiTextShortcuts: Array.isArray((incoming as any).aiTextShortcuts)
      ? (incoming as any).aiTextShortcuts
        .map((entry: any) => ({
          id: String(entry?.id || "").trim(),
          trigger: String(entry?.trigger || "").trim(),
          content: String(entry?.content || "").trim(),
        }))
        .filter((entry: AiTextShortcut) => entry.id && entry.trigger)
      : base.aiTextShortcuts,
    aiAutoLabel: {
      ...base.aiAutoLabel,
      ...((incoming as any).aiAutoLabel || {}),
      enabled: typeof ((incoming as any).aiAutoLabel || {}).enabled === "boolean"
        ? Boolean(((incoming as any).aiAutoLabel || {}).enabled)
        : base.aiAutoLabel.enabled,
      autoDraftEnabled: typeof ((incoming as any).aiAutoLabel || {}).autoDraftEnabled === "boolean"
        ? Boolean(((incoming as any).aiAutoLabel || {}).autoDraftEnabled)
        : base.aiAutoLabel.autoDraftEnabled,
      labels: {
        ...base.aiAutoLabel.labels,
        ...((((incoming as any).aiAutoLabel || {}).labels) || {}),
      },
    },
    aiFontPreference: {
      ...base.aiFontPreference,
      ...((incoming as any).aiFontPreference || {}),
      family: String((((incoming as any).aiFontPreference || {}).family ?? base.aiFontPreference.family) || "").trim() || base.aiFontPreference.family,
      size: Math.max(9, Math.min(20, Number((((incoming as any).aiFontPreference || {}).size ?? base.aiFontPreference.size) || base.aiFontPreference.size))),
      color: String((((incoming as any).aiFontPreference || {}).color ?? base.aiFontPreference.color) || "").trim() || base.aiFontPreference.color,
    },
    invoiceStudio: {
      ...base.invoiceStudio,
      ...((incoming as any).invoiceStudio || {}),
      enabled: typeof ((incoming as any).invoiceStudio || {}).enabled === "boolean"
        ? Boolean(((incoming as any).invoiceStudio || {}).enabled)
        : base.invoiceStudio.enabled,
      baseUrl: String((((incoming as any).invoiceStudio || {}).baseUrl ?? base.invoiceStudio.baseUrl) || "").trim(),
      email: String((((incoming as any).invoiceStudio || {}).email ?? base.invoiceStudio.email) || "").trim(),
      password: String((((incoming as any).invoiceStudio || {}).password ?? base.invoiceStudio.password) || "").trim(),
      project: String((((incoming as any).invoiceStudio || {}).project ?? base.invoiceStudio.project) || "").trim(),
    },
    referenceCodes: {
      ...base.referenceCodes,
      ...((incoming as any).referenceCodes || {}),
      prefixes: {
        ...base.referenceCodes.prefixes,
        ...(((incoming as any).referenceCodes || {}).prefixes || {}),
      },
      counters: {
        ...base.referenceCodes.counters,
        ...(((incoming as any).referenceCodes || {}).counters || {}),
        perType: {
          ...base.referenceCodes.counters.perType,
          ...((((incoming as any).referenceCodes || {}).counters || {}).perType || {}),
        },
      },
    },
    groups: normalizedGroups,
    // Legacy top-level aliases are written as compatibility mirrors only.
    // settings.groups remains the sole canonical source of truth for Grupos.
    groupStorage: {
      ...(normalizedGroupAliases.groupStorage || normalizedGroups.storage),
    },
    groupsTabSettings: {
      ...(normalizedGroupAliases.groupsTabSettings || normalizedGroups.tab),
    },
    groupLabelsManagerEnabled: normalizedGroupAliases.groupLabelsManagerEnabled ?? normalizedGroups.labels.managerEnabled,
    groupLabelCatalog: Array.isArray(normalizedGroupAliases.groupLabelCatalog)
      ? [...normalizedGroupAliases.groupLabelCatalog]
      : [...normalizedGroups.labels.catalog],
    groupFavoriteIds: Array.isArray(normalizedGroupAliases.groupFavoriteIds)
      ? [...normalizedGroupAliases.groupFavoriteIds]
      : [...normalizedGroups.labels.favoriteIds],
    groupTicketsEnabled: normalizedGroupAliases.groupTicketsEnabled ?? normalizedGroups.tickets.enabled,
    groupTicketUi: {
      ...(normalizedGroupAliases.groupTicketUi || normalizedGroups.tickets.ui),
    },
    groupOutlookCategories: {
      ...(normalizedGroupAliases.groupOutlookCategories || normalizedGroups.outlookCategories),
    },
    crm2OdooLayout: {
      ...base.crm2OdooLayout,
      ...incomingLayout,
      mode: incomingLayoutMode,
      includeAnchorIndex: typeof incomingLayout.includeAnchorIndex === "boolean"
        ? incomingLayout.includeAnchorIndex
        : base.crm2OdooLayout.includeAnchorIndex,
      showBackToTopLinks: typeof incomingLayout.showBackToTopLinks === "boolean"
        ? incomingLayout.showBackToTopLinks
        : base.crm2OdooLayout.showBackToTopLinks,
      project: {
        ...base.crm2OdooLayout.project,
        ...(incomingLayout.project || {}),
        model: "project.project",
        mode: normalizeCrm2OdooLayoutMode((incomingLayout.project || {}).mode ?? incomingLayoutMode),
        fallbackToDescription: typeof ((incomingLayout.project || {}).fallbackToDescription) === "boolean"
          ? (incomingLayout.project || {}).fallbackToDescription
          : base.crm2OdooLayout.project.fallbackToDescription,
      },
      lead: {
        ...base.crm2OdooLayout.lead,
        ...(incomingLayout.lead || {}),
        model: "crm.lead",
        mode: normalizeCrm2OdooLayoutMode((incomingLayout.lead || {}).mode ?? incomingLayoutMode),
        fallbackToDescription: typeof ((incomingLayout.lead || {}).fallbackToDescription) === "boolean"
          ? (incomingLayout.lead || {}).fallbackToDescription
          : base.crm2OdooLayout.lead.fallbackToDescription,
      },
      task: {
        ...base.crm2OdooLayout.task,
        ...(incomingLayout.task || {}),
        model: "project.task",
        mode: normalizeCrm2OdooLayoutMode((incomingLayout.task || {}).mode ?? incomingLayoutMode),
        fallbackToDescription: typeof ((incomingLayout.task || {}).fallbackToDescription) === "boolean"
          ? (incomingLayout.task || {}).fallbackToDescription
          : base.crm2OdooLayout.task.fallbackToDescription,
      },
      ticket: {
        ...base.crm2OdooLayout.ticket,
        ...(incomingLayout.ticket || {}),
        model: "helpdesk.ticket",
        mode: normalizeCrm2OdooLayoutMode((incomingLayout.ticket || {}).mode ?? incomingLayoutMode),
        fallbackToDescription: typeof ((incomingLayout.ticket || {}).fallbackToDescription) === "boolean"
          ? (incomingLayout.ticket || {}).fallbackToDescription
          : base.crm2OdooLayout.ticket.fallbackToDescription,
      },
    },
  };

  // guard against wrong versions
  merged.version = 1;
  return compactSettingsForStorage(applyLegacySettingsMigrations(merged));
}

export function getCachedSettingsSnapshot(): CockpitSettingsV1 {
  const raw = globalThis.localStorage?.getItem(KEY_SETTINGS);
  const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
  const merged = mergeSettings(DEFAULT_SETTINGS, parsed);
  if (raw) {
    const compactJson = serializeSettings(merged);
    if (compactJson !== raw) writeLocalSettingsCache(compactJson);
  }
  return merged;
}

export async function getSettings(): Promise<CockpitSettingsV1> {
  await officeReady();

  const rs = getRoamingSettings();
  if (rs) {
    const raw = rs.get(KEY_SETTINGS);
    const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
    const merged = mergeSettings(getCachedSettingsSnapshot(), parsed);
    const migrated = compactSettingsForStorage(applyLegacySettingsMigrations(merged, true));
    if (serializeSettings(migrated) !== serializeSettings(merged)) {
      const json = serializeSettings(migrated);
      try {
        rs.set(KEY_SETTINGS, json);
        await saveRoamingSettings(rs);
      } catch {
        // Keep returning the migrated value; persistence can be retried on next settings save.
      }
      writeLocalSettingsCache(json);
    }
    return migrated;
  }

  // fallback (dev / non-office)
  const merged = getCachedSettingsSnapshot();
  const migrated = compactSettingsForStorage(applyLegacySettingsMigrations(merged, true));
  if (serializeSettings(migrated) !== serializeSettings(merged)) {
    writeLocalSettingsRequired(serializeSettings(migrated));
  }
  return migrated;
}

export async function saveSettings(patch: Partial<CockpitSettingsV1>): Promise<CockpitSettingsV1> {
  await officeReady();
  const current = await getSettings();
  const next = mergeSettings(current, patch);
  const json = serializeSettings(next);

  const rs = getRoamingSettings();
  if (rs) {
    rs.set(KEY_SETTINGS, json);
    await saveRoamingSettings(rs);
    writeLocalSettingsCache(json);
    emitSettingsUpdated(next);
    return next;
  }

  writeLocalSettingsRequired(json);
  emitSettingsUpdated(next);
  return next;
}

export async function resetSettings(): Promise<CockpitSettingsV1> {
  await officeReady();
  const rs = getRoamingSettings();
  const next = compactSettingsForStorage(DEFAULT_SETTINGS);
  const json = serializeSettings(next);
  if (rs) {
    rs.set(KEY_SETTINGS, json);
    await saveRoamingSettings(rs);
    writeLocalSettingsCache(json);
    emitSettingsUpdated(next);
    return next;
  }
  writeLocalSettingsRequired(json);
  emitSettingsUpdated(next);
  return next;
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
  if (typeof globalThis.location?.origin === "string" && globalThis.location.origin.trim()) return globalThis.location.origin.trim();
  return "";
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
