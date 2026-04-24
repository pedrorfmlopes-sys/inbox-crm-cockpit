import {
  DEFAULT_GROUP_STORAGE_SETTINGS,
  normalizeGroupStorageSettings,
  type GroupStorageSettings,
} from "../storage/settings";
import {
  DEFAULT_GROUPS_TAB_SETTINGS,
  normalizeGroupsTabSettings,
  type GroupsTabSettings,
} from "./groupsTabSettings";

export type GroupTicketAutoLinkMode = "confirm" | "auto";

export type GroupTicketUiSettings = {
  autoLinkMode: GroupTicketAutoLinkMode;
  suggestDraftOnCreate: boolean;
  useAiDrafts: boolean;
  includeTicketCodeInSubject: boolean;
  aiInstructions: string;
};

export type GroupStateDefinition = {
  name: string;
  color?: string;
};

export type GroupStateCatalogSettings = {
  states: GroupStateDefinition[];
};

export type GroupLabelCatalogEntry = {
  label: string;
};

export type GroupOutlookCategorySettings = {
  enabled: boolean;
  includeGroups: boolean;
  includeTickets: boolean;
  includeStatuses: boolean;
  includeLabels: boolean;
};

export type GroupsModuleSettings = {
  storage: GroupStorageSettings;
  tab: GroupsTabSettings;
  groups: GroupStateCatalogSettings;
  references: GroupStateCatalogSettings;
  labels: {
    managerEnabled: boolean;
    catalog: GroupLabelCatalogEntry[];
    favoriteIds: string[];
    states: GroupStateDefinition[];
  };
  tickets: {
    enabled: boolean;
    ui: GroupTicketUiSettings;
    states: GroupStateDefinition[];
  };
  outlookCategories: GroupOutlookCategorySettings;
};

export type GroupsLegacySettingsInput = {
  groupStorage?: Partial<GroupStorageSettings> | null;
  groupsTabSettings?: Partial<GroupsTabSettings> | null;
  groupLabelsManagerEnabled?: boolean | null;
  groupLabelCatalog?: unknown;
  groupFavoriteIds?: unknown;
  groupTicketsEnabled?: boolean | null;
  groupTicketUi?: Partial<GroupTicketUiSettings> | null;
  groupOutlookCategories?: Partial<GroupOutlookCategorySettings> | null;
};
// Compatibilidade transitória: aceita blobs antigos durante a migração para
// `settings.groups`, mas não deve voltar a ser lida como fonte ativa de runtime.

export type GroupsSettingsLike =
  | ({
      groups?: Partial<GroupsModuleSettings> | null;
    } & GroupsLegacySettingsInput)
  | Partial<GroupsModuleSettings>
  | null
  | undefined;

export const DEFAULT_GROUP_TICKET_UI_SETTINGS: GroupTicketUiSettings = {
  autoLinkMode: "confirm",
  suggestDraftOnCreate: true,
  useAiDrafts: true,
  includeTicketCodeInSubject: true,
  aiInstructions:
    "Escreve em tom profissional e claro. Indica o numero do ticket e pede que todas as respostas futuras mantenham esse numero no assunto.",
};

export const DEFAULT_GROUP_OUTLOOK_CATEGORY_SETTINGS: GroupOutlookCategorySettings = {
  enabled: true,
  includeGroups: true,
  includeTickets: true,
  includeStatuses: true,
  includeLabels: false,
};

export const DEFAULT_GROUP_ENTITY_STATES: GroupStateDefinition[] = [
  { name: "em_analise", color: "#f59e0b" },
  { name: "em_progresso", color: "#3b82f6" },
  { name: "concluido", color: "#10b981" },
];

export const DEFAULT_REFERENCE_ENTITY_STATES: GroupStateDefinition[] = [
  ...DEFAULT_GROUP_ENTITY_STATES,
];

export const DEFAULT_LABEL_ENTITY_STATES: GroupStateDefinition[] = [
  { name: "em_analise", color: "#f59e0b" },
  { name: "respondido", color: "#3b82f6" },
  { name: "confirmado", color: "#10b981" },
  { name: "arquivado", color: "#6b7280" },
  { name: "cancelado", color: "#ef4444" },
];

export const DEFAULT_TICKET_ENTITY_STATES: GroupStateDefinition[] = [
  { name: "open", color: "#3b82f6" },
  { name: "closed", color: "#10b981" },
];

export const DEFAULT_GROUPS_MODULE_SETTINGS: GroupsModuleSettings = {
  storage: {
    ...DEFAULT_GROUP_STORAGE_SETTINGS,
  },
  tab: {
    ...DEFAULT_GROUPS_TAB_SETTINGS,
  },
  groups: {
    states: [...DEFAULT_GROUP_ENTITY_STATES],
  },
  references: {
    states: [...DEFAULT_REFERENCE_ENTITY_STATES],
  },
  labels: {
    managerEnabled: true,
    catalog: [],
    favoriteIds: [],
    states: [...DEFAULT_LABEL_ENTITY_STATES],
  },
  tickets: {
    enabled: true,
    ui: {
      ...DEFAULT_GROUP_TICKET_UI_SETTINGS,
    },
    states: [...DEFAULT_TICKET_ENTITY_STATES],
  },
  outlookCategories: {
    ...DEFAULT_GROUP_OUTLOOK_CATEGORY_SETTINGS,
  },
};

function normalizeGroupTicketAutoLinkMode(value: unknown): GroupTicketAutoLinkMode {
  return String(value || "").trim().toLowerCase() === "auto" ? "auto" : "confirm";
}

function normalizeGroupStateName(value: unknown): string {
  return String(value || "").trim().toLowerCase();
}

function normalizeGroupStateDefinition(value: unknown): GroupStateDefinition | null {
  const record = value && typeof value === "object"
    ? (value as Record<string, unknown>)
    : {};
  const name = normalizeGroupStateName(record.name ?? record.value ?? record.id ?? value);
  if (!name) return null;
  const color = String(record.color || "").trim();
  return {
    name,
    color: color || undefined,
  };
}

function normalizeGroupStateCatalogSettings(
  input: unknown,
  fallback: GroupStateDefinition[]
): GroupStateCatalogSettings {
  const seen = new Map<string, GroupStateDefinition>();
  for (const entry of Array.isArray(fallback) ? fallback : []) {
    const normalized = normalizeGroupStateDefinition(entry);
    if (!normalized) continue;
    seen.set(normalized.name, normalized);
  }
  const rawStates =
    input && typeof input === "object" && !Array.isArray(input)
      ? (input as Record<string, unknown>).states
      : input;
  for (const entry of Array.isArray(rawStates) ? rawStates : []) {
    const normalized = normalizeGroupStateDefinition(entry);
    if (!normalized) continue;
    const previous = seen.get(normalized.name);
    seen.set(normalized.name, {
      ...previous,
      ...normalized,
    });
  }
  return {
    states: Array.from(seen.values()),
  };
}

export function normalizeGroupLabelCatalog(
  entries: unknown,
  fallback: GroupLabelCatalogEntry[] = []
): GroupLabelCatalogEntry[] {
  const byKey = new Map<string, GroupLabelCatalogEntry>();

  const commit = (raw: unknown) => {
    const value = raw as Record<string, unknown> | string | null | undefined;
    const rawLabel =
      typeof value === "string"
        ? value
        : value?.label ?? value?.name ?? value?.value ?? value?.id;
    const label = String(rawLabel || "").trim();
    if (!label) return;
    const key = label.toLowerCase();
    const previous = byKey.get(key);
    byKey.set(key, {
      label: previous?.label || label,
    });
  };

  for (const entry of Array.isArray(fallback) ? fallback : []) commit(entry);
  for (const entry of Array.isArray(entries) ? entries : []) commit(entry);

  return Array.from(byKey.values()).sort((a, b) => a.label.localeCompare(b.label, "pt-PT"));
}

function normalizeFavoriteIds(values: unknown): string[] {
  return Array.from(
    new Set(
      (Array.isArray(values) ? values : [])
        .map((entry) => String(entry || "").trim())
        .filter(Boolean)
    )
  );
}

function normalizeGroupTicketUiSettings(
  input: Partial<GroupTicketUiSettings> | null | undefined
): GroupTicketUiSettings {
  const value = input || {};
  return {
    autoLinkMode: normalizeGroupTicketAutoLinkMode(value.autoLinkMode),
    suggestDraftOnCreate: value.suggestDraftOnCreate !== false,
    useAiDrafts: value.useAiDrafts !== false,
    includeTicketCodeInSubject: value.includeTicketCodeInSubject !== false,
    aiInstructions:
      String(value.aiInstructions || "").trim() || DEFAULT_GROUP_TICKET_UI_SETTINGS.aiInstructions,
  };
}

function normalizeGroupOutlookCategorySettings(
  input: Partial<GroupOutlookCategorySettings> | null | undefined
): GroupOutlookCategorySettings {
  const value = input || {};
  return {
    enabled: value.enabled === true,
    includeGroups: value.includeGroups !== false,
    includeTickets: value.includeTickets !== false,
    includeStatuses: value.includeStatuses !== false,
    includeLabels: value.includeLabels === true,
  };
}

export function getGroupLabelCatalogLabels(
  catalog: GroupLabelCatalogEntry[] | null | undefined
): string[] {
  return normalizeGroupLabelCatalog(catalog || []).map((entry) => entry.label);
}

export function getGroupStateCatalogLabels(
  catalog: GroupStateDefinition[] | null | undefined
): string[] {
  return normalizeGroupStateCatalogSettings(catalog || [], []).states.map((entry) => entry.name);
}

export function findGroupStateDefinition(
  catalog: GroupStateDefinition[] | null | undefined,
  state: string
): GroupStateDefinition | null {
  const normalized = normalizeGroupStateName(state);
  if (!normalized) return null;
  return (
    normalizeGroupStateCatalogSettings(catalog || [], []).states.find(
      (entry) => entry.name === normalized
    ) || null
  );
}

export function findGroupLabelCatalogEntry(
  catalog: GroupLabelCatalogEntry[] | null | undefined,
  label: string
): GroupLabelCatalogEntry | null {
  const normalized = String(label || "").trim().toLowerCase();
  if (!normalized) return null;
  return (
    normalizeGroupLabelCatalog(catalog || []).find(
      (entry) => entry.label.toLowerCase() === normalized
    ) || null
  );
}

export function normalizeGroupsModuleSettings(
  input: Partial<GroupsModuleSettings> | null | undefined,
  legacy?: GroupsLegacySettingsInput | null,
  fallback: GroupsModuleSettings = DEFAULT_GROUPS_MODULE_SETTINGS
): GroupsModuleSettings {
  const next = input || {};
  const legacyInput = legacy || {};
  const storage = normalizeGroupStorageSettings({
    ...fallback.storage,
    ...(legacyInput.groupStorage || {}),
    ...(next.storage || {}),
  });
  const tab = normalizeGroupsTabSettings({
    ...fallback.tab,
    ...(legacyInput.groupsTabSettings || {}),
    ...(next.tab || {}),
  });
  const groups = normalizeGroupStateCatalogSettings(
    next.groups,
    fallback.groups.states
  );
  const references = normalizeGroupStateCatalogSettings(
    next.references,
    fallback.references.states
  );
  const labels = {
    managerEnabled:
      typeof next.labels?.managerEnabled === "boolean"
        ? next.labels.managerEnabled
        : typeof legacyInput.groupLabelsManagerEnabled === "boolean"
          ? legacyInput.groupLabelsManagerEnabled
          : fallback.labels.managerEnabled,
    catalog: normalizeGroupLabelCatalog(
      next.labels?.catalog ?? legacyInput.groupLabelCatalog ?? [],
      fallback.labels.catalog
    ),
    favoriteIds: normalizeFavoriteIds(
      next.labels?.favoriteIds ?? legacyInput.groupFavoriteIds ?? fallback.labels.favoriteIds
    ),
    states: normalizeGroupStateCatalogSettings(
      next.labels?.states,
      fallback.labels.states
    ).states,
  };
  const tickets = {
    enabled:
      typeof next.tickets?.enabled === "boolean"
        ? next.tickets.enabled
        : typeof legacyInput.groupTicketsEnabled === "boolean"
          ? legacyInput.groupTicketsEnabled
          : fallback.tickets.enabled,
    ui: normalizeGroupTicketUiSettings({
      ...fallback.tickets.ui,
      ...(legacyInput.groupTicketUi || {}),
      ...(next.tickets?.ui || {}),
    }),
    states: normalizeGroupStateCatalogSettings(
      next.tickets?.states,
      fallback.tickets.states
    ).states,
  };
  const outlookCategories = normalizeGroupOutlookCategorySettings({
    ...fallback.outlookCategories,
    ...(legacyInput.groupOutlookCategories || {}),
    ...(next.outlookCategories || {}),
  });

  return {
    storage,
    tab,
    groups,
    references,
    labels,
    tickets,
    outlookCategories,
  };
}

export function getGroupsModuleSettings(
  input: GroupsSettingsLike,
  fallback: GroupsModuleSettings = DEFAULT_GROUPS_MODULE_SETTINGS
): GroupsModuleSettings {
  const value = input || null;
  if (value && typeof value === "object" && "groups" in value) {
    const scoped = value as { groups?: Partial<GroupsModuleSettings> | null } & GroupsLegacySettingsInput;
    return normalizeGroupsModuleSettings(scoped.groups || null, scoped, fallback);
  }
  return normalizeGroupsModuleSettings(value as Partial<GroupsModuleSettings> | null, null, fallback);
}

export function buildGroupsSettingsPatch(
  current: GroupsSettingsLike,
  patch: Partial<GroupsModuleSettings>,
  fallback: GroupsModuleSettings = DEFAULT_GROUPS_MODULE_SETTINGS
): { groups: GroupsModuleSettings } {
  const base = getGroupsModuleSettings(current, fallback);
  return {
    groups: normalizeGroupsModuleSettings(
      {
        ...base,
        ...patch,
      },
      null,
      base
    ),
  };
}
