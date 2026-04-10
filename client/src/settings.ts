// client/src/settings.ts
// Settings storage (RoamingSettings preferred; localStorage fallback for non-Office contexts)

import type { AiTone } from "./ai/aiClient";
import {
  DEFAULT_GROUP_STORAGE_SETTINGS,
  normalizeGroupStorageSettings,
  type GroupStorageMode,
  type GroupStorageLegacyProvider as GroupStorageProvider,
  type GroupStorageSettings,
} from "./modules/crm/groups-v1/storage/settings";

export type AppLocale = "pt-PT" | "es-ES" | "en-GB" | "it-IT" | "de-DE";
export type LangOption = AppLocale | "auto";
export type ReplyLength = "xs" | "s" | "m" | "l";
export type SkinId = "classic" | "mailmaestro" | "vibrant";

export type ResponsePreset = {
  id: string;
  name: string;
  prompt: string;
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
export type GroupTicketAutoLinkMode = "confirm" | "auto";
export type { GroupStorageMode, GroupStorageProvider, GroupStorageSettings };

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

export type GroupTicketUiSettings = {
  autoLinkMode: GroupTicketAutoLinkMode;
  suggestDraftOnCreate: boolean;
  useAiDrafts: boolean;
  includeTicketCodeInSubject: boolean;
  aiInstructions: string;
};

export type GroupLabelStatus = "em_analise" | "em_progresso" | "concluido" | string;

export type GroupLabelCatalogEntry = {
  label: string;
  categorize: boolean;
  hasStatus: boolean;
  status?: GroupLabelStatus;
};

export type InvoiceStudioSettings = {
  enabled: boolean;
  baseUrl: string;
  email: string;
  password: string;
  project: string;
};

export type GroupOutlookCategorySettings = {
  enabled: boolean;
  includeGroups: boolean;
  includeTickets: boolean;
  includeStatuses: boolean;
  includeLabels: boolean;
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

  // New: Contact Aliases (Forwarding Shortcuts)
  contactAliases: ContactAlias[];

  // AI module settings (MailMaestro-like per-module configuration)
  aiCustomTones: AiCustomTone[];
  aiTextShortcuts: AiTextShortcut[];
  aiAutoLabel: AiAutoLabelSettings;
  aiFontPreference: AiFontPreference;

  // Configurable reference codes for Odoo-created records
  referenceCodes: ReferenceCodeSettings;

  // Group document storage configuration
  groupStorage: GroupStorageSettings;

  // Optional extra: central label manager for group labels
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
  responsePresets: [
    { id: "p1", name: "Pedido de Dados", prompt: "Agradece o contacto e solicita os dados de faturação (NIF, Morada) para podermos proceder." },
    { id: "p2", name: "Agendamento Carga", prompt: "Informa que a mercadoria está pronta e solicita confirmação de data/hora para a recolha no nosso armazém." },
    { id: "p3", name: "Follow-up Proposta", prompt: "Faz um follow-up cortês sobre a última proposta enviada, perguntando se restam dúvidas técnicas." }
  ],
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
  groupStorage: {
    ...DEFAULT_GROUP_STORAGE_SETTINGS,
  },
  groupLabelsManagerEnabled: true,
  groupLabelCatalog: [],
  groupFavoriteIds: [],
  groupTicketsEnabled: true,
  groupTicketUi: {
    autoLinkMode: "confirm",
    suggestDraftOnCreate: true,
    useAiDrafts: true,
    includeTicketCodeInSubject: true,
    aiInstructions: "Escreve em tom profissional e claro. Indica o numero do ticket e pede que todas as respostas futuras mantenham esse numero no assunto.",
  },
  groupOutlookCategories: {
    enabled: true,
    includeGroups: true,
    includeTickets: true,
    includeStatuses: true,
    includeLabels: false,
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

function normalizeGroupTicketAutoLinkMode(value: unknown): GroupTicketAutoLinkMode {
  const normalized = String(value || "").trim().toLowerCase();
  return normalized === "auto" ? "auto" : "confirm";
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

function normalizeGroupLabelStatus(value: any): GroupLabelStatus | undefined {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === "em_progresso") return "em_progresso";
  if (raw === "concluido") return "concluido";
  if (raw === "em_analise") return "em_analise";
  return undefined;
}

export function normalizeGroupLabelCatalog(
  entries: any,
  fallback: GroupLabelCatalogEntry[] = []
): GroupLabelCatalogEntry[] {
  const byKey = new Map<string, GroupLabelCatalogEntry>();

  const commit = (raw: any) => {
    const rawLabel = typeof raw === "string" ? raw : raw?.label ?? raw?.name ?? raw?.value ?? raw?.id;
    const label = String(rawLabel || "").trim();
    if (!label) return;
    const key = label.toLowerCase();
    const previous = byKey.get(key);
    const categorize = typeof raw?.categorize === "boolean"
      ? Boolean(raw.categorize)
      : previous?.categorize ?? false;
    const hasStatus = typeof raw?.hasStatus === "boolean"
      ? Boolean(raw.hasStatus)
      : previous?.hasStatus ?? false;
    const status = normalizeGroupLabelStatus(raw?.status ?? raw?.defaultStatus ?? previous?.status);
    byKey.set(key, {
      label: previous?.label || label,
      categorize,
      hasStatus,
      status: hasStatus ? (status || "em_analise") : undefined,
    });
  };

  for (const entry of Array.isArray(fallback) ? fallback : []) {
    commit(entry);
  }
  for (const entry of Array.isArray(entries) ? entries : []) {
    commit(entry);
  }

  return Array.from(byKey.values()).sort((a, b) => a.label.localeCompare(b.label, "pt-PT"));
}

export function getGroupLabelCatalogLabels(catalog: GroupLabelCatalogEntry[] | null | undefined): string[] {
  return normalizeGroupLabelCatalog(catalog || []).map((entry) => entry.label);
}

export function findGroupLabelCatalogEntry(
  catalog: GroupLabelCatalogEntry[] | null | undefined,
  label: string
): GroupLabelCatalogEntry | null {
  const normalized = String(label || "").trim().toLowerCase();
  if (!normalized) return null;
  return normalizeGroupLabelCatalog(catalog || []).find((entry) => entry.label.toLowerCase() === normalized) || null;
}

function mergeSettings(base: CockpitSettingsV1, incoming: Partial<CockpitSettingsV1> | null): CockpitSettingsV1 {
  if (!incoming) return base;
  const incomingLayout = ((incoming as any).crm2OdooLayout || {});
  const incomingLayoutMode = normalizeCrm2OdooLayoutMode(incomingLayout.mode ?? base.crm2OdooLayout.mode);

  const merged: CockpitSettingsV1 = {
    ...base,
    ...incoming,
    signatures: { ...base.signatures, ...(incoming.signatures || {}) },
    signaturesHtml: { ...(base.signaturesHtml || {}), ...((incoming as any).signaturesHtml || {}) },
    signatureImageUrl: { ...(base.signatureImageUrl || {}), ...((incoming as any).signatureImageUrl || {}) },
    signatureImageMaxWidth: { ...(base.signatureImageMaxWidth || {}), ...((incoming as any).signatureImageMaxWidth || {}) },
    aiKnowledge: Array.isArray(incoming.aiKnowledge) ? incoming.aiKnowledge : base.aiKnowledge,
    aiManualOnly: typeof incoming.aiManualOnly === "boolean" ? incoming.aiManualOnly : base.aiManualOnly,
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
    groupStorage: {
      ...normalizeGroupStorageSettings({
        ...base.groupStorage,
        ...((incoming as any).groupStorage || {}),
      }),
    },
    groupLabelsManagerEnabled: typeof (incoming as any).groupLabelsManagerEnabled === "boolean"
      ? Boolean((incoming as any).groupLabelsManagerEnabled)
      : base.groupLabelsManagerEnabled,
    groupLabelCatalog: Array.isArray((incoming as any).groupLabelCatalog)
      ? normalizeGroupLabelCatalog((incoming as any).groupLabelCatalog, base.groupLabelCatalog)
      : base.groupLabelCatalog,
    groupFavoriteIds: Array.isArray((incoming as any).groupFavoriteIds)
      ? Array.from(new Set((incoming as any).groupFavoriteIds.map((entry: any) => String(entry || "").trim()).filter(Boolean)))
      : base.groupFavoriteIds,
    groupTicketsEnabled: typeof (incoming as any).groupTicketsEnabled === "boolean"
      ? Boolean((incoming as any).groupTicketsEnabled)
      : base.groupTicketsEnabled,
    groupTicketUi: {
      ...base.groupTicketUi,
      ...((incoming as any).groupTicketUi || {}),
      autoLinkMode: normalizeGroupTicketAutoLinkMode(((incoming as any).groupTicketUi || {}).autoLinkMode ?? base.groupTicketUi.autoLinkMode),
      suggestDraftOnCreate: typeof ((incoming as any).groupTicketUi || {}).suggestDraftOnCreate === "boolean"
        ? Boolean(((incoming as any).groupTicketUi || {}).suggestDraftOnCreate)
        : base.groupTicketUi.suggestDraftOnCreate,
      useAiDrafts: typeof ((incoming as any).groupTicketUi || {}).useAiDrafts === "boolean"
        ? Boolean(((incoming as any).groupTicketUi || {}).useAiDrafts)
        : base.groupTicketUi.useAiDrafts,
      includeTicketCodeInSubject: typeof ((incoming as any).groupTicketUi || {}).includeTicketCodeInSubject === "boolean"
        ? Boolean(((incoming as any).groupTicketUi || {}).includeTicketCodeInSubject)
        : base.groupTicketUi.includeTicketCodeInSubject,
      aiInstructions: String((((incoming as any).groupTicketUi || {}).aiInstructions ?? base.groupTicketUi.aiInstructions) || "").trim(),
    },
    groupOutlookCategories: {
      ...base.groupOutlookCategories,
      ...((incoming as any).groupOutlookCategories || {}),
      enabled: typeof ((incoming as any).groupOutlookCategories || {}).enabled === "boolean"
        ? Boolean(((incoming as any).groupOutlookCategories || {}).enabled)
        : base.groupOutlookCategories.enabled,
      includeGroups: typeof ((incoming as any).groupOutlookCategories || {}).includeGroups === "boolean"
        ? Boolean(((incoming as any).groupOutlookCategories || {}).includeGroups)
        : base.groupOutlookCategories.includeGroups,
      includeTickets: typeof ((incoming as any).groupOutlookCategories || {}).includeTickets === "boolean"
        ? Boolean(((incoming as any).groupOutlookCategories || {}).includeTickets)
        : base.groupOutlookCategories.includeTickets,
      includeStatuses: typeof ((incoming as any).groupOutlookCategories || {}).includeStatuses === "boolean"
        ? Boolean(((incoming as any).groupOutlookCategories || {}).includeStatuses)
        : base.groupOutlookCategories.includeStatuses,
      includeLabels: typeof ((incoming as any).groupOutlookCategories || {}).includeLabels === "boolean"
        ? Boolean(((incoming as any).groupOutlookCategories || {}).includeLabels)
        : base.groupOutlookCategories.includeLabels,
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
  return merged;
}

export function getCachedSettingsSnapshot(): CockpitSettingsV1 {
  const raw = globalThis.localStorage?.getItem(KEY_SETTINGS);
  const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
  return mergeSettings(DEFAULT_SETTINGS, parsed);
}

export async function getSettings(): Promise<CockpitSettingsV1> {
  await officeReady();

  const rs = getRoamingSettings();
  if (rs) {
    const raw = rs.get(KEY_SETTINGS);
    const parsed = safeJsonParse<Partial<CockpitSettingsV1>>(raw);
    return mergeSettings(getCachedSettingsSnapshot(), parsed);
  }

  // fallback (dev / non-office)
  return getCachedSettingsSnapshot();
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
    globalThis.localStorage?.setItem(KEY_SETTINGS, json);
    emitSettingsUpdated(next);
    return next;
  }

  globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  emitSettingsUpdated(next);
  return next;
}

export async function resetSettings(): Promise<CockpitSettingsV1> {
  await officeReady();
  const rs = getRoamingSettings();
  const json = JSON.stringify(DEFAULT_SETTINGS);
  if (rs) {
    rs.set(KEY_SETTINGS, json);
    await saveRoamingSettings(rs);
    globalThis.localStorage?.setItem(KEY_SETTINGS, json);
    emitSettingsUpdated(DEFAULT_SETTINGS);
    return DEFAULT_SETTINGS;
  }
  globalThis.localStorage?.setItem(KEY_SETTINGS, json);
  emitSettingsUpdated(DEFAULT_SETTINGS);
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
