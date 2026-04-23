export type GroupsSettingsStorageMode = "local_indexeddb" | "disabled";
export type GroupsSettingsAttachmentStrategy = "server" | "outside" | "by_size";
export type GroupsSettingsFrequency = "manual" | "daily" | "weekly";
export type GroupsSettingsMigrationMode = "always_ask" | "move" | "copy";

export type GroupsTabSettings = {
  groupsTabEnabled: boolean;
  storageMode: GroupsSettingsStorageMode;
  baseFolderPath: string;
  locationStatus: string;
  autoCreateCaseOnNewEmail: boolean;
  reopenExistingCase: boolean;
  recreateIntermediateCopy: boolean;
  validateLocationOnOpen: boolean;
  blockTabIfUnavailable: boolean;
  warnIfUnavailable: boolean;
  autoRetryValidation: boolean;
  attachmentStrategy: GroupsSettingsAttachmentStrategy;
  saveAttachmentsOnServer: boolean;
  saveAttachmentsOutsideServer: boolean;
  attachmentServerLimitMb: number;
  attachmentIntermediateLimitMb: number;
  externalAttachmentFolder: string;
  showAttachmentMetadataOnServer: boolean;
  requireImmediatePreview: boolean;
  mixedCaseWarningDays: number;
  localAbandonedWarningDays: number;
  cleanupClosedCaseDays: number;
  cleanupAbandonedCaseDays: number;
  cleanupFrequency: GroupsSettingsFrequency;
  neverDeleteMixedSilently: boolean;
  warnUnclassifiedEmails: boolean;
  warnMixedCases: boolean;
  warningFrequency: GroupsSettingsFrequency;
  prepareTasksBridge: boolean;
  migrationTarget: string;
  migrationMode: GroupsSettingsMigrationMode;
  allowMoveExistingData: boolean;
  strictMigrationSafety: boolean;
  mergeExistingData: boolean;
  explorerServerPrimary: boolean;
  explorerOpenStoredAttachments: boolean;
  explorerGenerateReply: boolean;
  groupsVersion: string;
  quickDiagnostic: string;
};

export const DEFAULT_GROUPS_TAB_SETTINGS: GroupsTabSettings = {
  groupsTabEnabled: true,
  storageMode: "local_indexeddb",
  baseFolderPath: "",
  locationStatus: "Storage local do add-in",
  autoCreateCaseOnNewEmail: true,
  reopenExistingCase: true,
  recreateIntermediateCopy: true,
  validateLocationOnOpen: true,
  blockTabIfUnavailable: true,
  warnIfUnavailable: true,
  autoRetryValidation: true,
  attachmentStrategy: "by_size",
  saveAttachmentsOnServer: true,
  saveAttachmentsOutsideServer: true,
  attachmentServerLimitMb: 10,
  attachmentIntermediateLimitMb: 50,
  externalAttachmentFolder: "",
  showAttachmentMetadataOnServer: true,
  requireImmediatePreview: true,
  mixedCaseWarningDays: 15,
  localAbandonedWarningDays: 30,
  cleanupClosedCaseDays: 15,
  cleanupAbandonedCaseDays: 90,
  cleanupFrequency: "daily",
  neverDeleteMixedSilently: true,
  warnUnclassifiedEmails: true,
  warnMixedCases: true,
  warningFrequency: "daily",
  prepareTasksBridge: false,
  migrationTarget: "",
  migrationMode: "always_ask",
  allowMoveExistingData: false,
  strictMigrationSafety: true,
  mergeExistingData: false,
  explorerServerPrimary: true,
  explorerOpenStoredAttachments: true,
  explorerGenerateReply: true,
  groupsVersion: "Grupos v1",
  quickDiagnostic: "Sem pasta intermédia definida; o fallback principal é o storage local do add-in.",
};

function clampInteger(value: unknown, fallback: number, min: number, max: number): number {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return fallback;
  return Math.min(Math.max(Math.round(numeric), min), max);
}

function normalizeStorageMode(value: unknown): GroupsSettingsStorageMode {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "disabled") return "disabled";
  if (normalized === "local_indexeddb" || normalized === "onedrive_sharepoint") return "local_indexeddb";
  return "local_indexeddb";
}

function normalizeAttachmentStrategy(value: unknown): GroupsSettingsAttachmentStrategy {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "server" || normalized === "outside" || normalized === "by_size") return normalized;
  return DEFAULT_GROUPS_TAB_SETTINGS.attachmentStrategy;
}

function normalizeFrequency(value: unknown, fallback: GroupsSettingsFrequency): GroupsSettingsFrequency {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "manual" || normalized === "daily" || normalized === "weekly") return normalized;
  return fallback;
}

function normalizeMigrationMode(value: unknown): GroupsSettingsMigrationMode {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "move" || normalized === "copy" || normalized === "always_ask") return normalized;
  return DEFAULT_GROUPS_TAB_SETTINGS.migrationMode;
}

export function deriveGroupsLocationStatus(
  settings: Pick<GroupsTabSettings, "storageMode" | "baseFolderPath">
): string {
  if (settings.storageMode === "disabled") return "Intermedio desativado";
  return String(settings.baseFolderPath || "").trim()
    ? "Pasta local intermédia configurada"
    : "Storage local do add-in";
}

export function deriveGroupsQuickDiagnostic(
  settings: Pick<
    GroupsTabSettings,
    "storageMode" | "baseFolderPath" | "warnUnclassifiedEmails" | "warnMixedCases" | "validateLocationOnOpen"
  >
): string {
  if (settings.storageMode === "disabled") {
    return "Storage intermedio desligado; a aba trabalha sem persistencia local do caso";
  }
  if (!String(settings.baseFolderPath || "").trim()) {
    return "Sem pasta intermédia definida; o caso abre no storage local do add-in e só cai para memória como último fallback técnico";
  }
  if (!settings.validateLocationOnOpen) {
    return "Pasta intermédia configurada sem validação leve ao abrir";
  }
  if (!settings.warnUnclassifiedEmails && !settings.warnMixedCases) {
    return "Pasta intermédia configurada com avisos leves desativados";
  }
  return "Pasta intermédia configurada; o caso grava logo aí e a persistência final continua a ser feita no apply";
}

export function normalizeGroupsTabSettings(input: Partial<GroupsTabSettings> | null | undefined): GroupsTabSettings {
  const merged = {
    ...DEFAULT_GROUPS_TAB_SETTINGS,
    ...(input || {}),
  };

  const normalized: GroupsTabSettings = {
    groupsTabEnabled: merged.groupsTabEnabled !== false,
    storageMode: normalizeStorageMode(merged.storageMode),
    baseFolderPath: String(merged.baseFolderPath || "").trim(),
    locationStatus: "",
    autoCreateCaseOnNewEmail: merged.autoCreateCaseOnNewEmail !== false,
    reopenExistingCase: merged.reopenExistingCase !== false,
    recreateIntermediateCopy: merged.recreateIntermediateCopy !== false,
    validateLocationOnOpen: merged.validateLocationOnOpen !== false,
    blockTabIfUnavailable: merged.blockTabIfUnavailable !== false,
    warnIfUnavailable: merged.warnIfUnavailable !== false,
    autoRetryValidation: merged.autoRetryValidation !== false,
    attachmentStrategy: normalizeAttachmentStrategy(merged.attachmentStrategy),
    saveAttachmentsOnServer: merged.saveAttachmentsOnServer !== false,
    saveAttachmentsOutsideServer: merged.saveAttachmentsOutsideServer !== false,
    attachmentServerLimitMb: clampInteger(
      merged.attachmentServerLimitMb,
      DEFAULT_GROUPS_TAB_SETTINGS.attachmentServerLimitMb,
      1,
      2048
    ),
    attachmentIntermediateLimitMb: clampInteger(
      merged.attachmentIntermediateLimitMb,
      DEFAULT_GROUPS_TAB_SETTINGS.attachmentIntermediateLimitMb,
      1,
      4096
    ),
    externalAttachmentFolder: String(merged.externalAttachmentFolder || "").trim(),
    showAttachmentMetadataOnServer: merged.showAttachmentMetadataOnServer !== false,
    requireImmediatePreview: merged.requireImmediatePreview !== false,
    mixedCaseWarningDays: clampInteger(
      merged.mixedCaseWarningDays,
      DEFAULT_GROUPS_TAB_SETTINGS.mixedCaseWarningDays,
      1,
      3650
    ),
    localAbandonedWarningDays: clampInteger(
      merged.localAbandonedWarningDays,
      DEFAULT_GROUPS_TAB_SETTINGS.localAbandonedWarningDays,
      1,
      3650
    ),
    cleanupClosedCaseDays: clampInteger(
      merged.cleanupClosedCaseDays,
      DEFAULT_GROUPS_TAB_SETTINGS.cleanupClosedCaseDays,
      1,
      3650
    ),
    cleanupAbandonedCaseDays: clampInteger(
      merged.cleanupAbandonedCaseDays,
      DEFAULT_GROUPS_TAB_SETTINGS.cleanupAbandonedCaseDays,
      1,
      3650
    ),
    cleanupFrequency: normalizeFrequency(merged.cleanupFrequency, DEFAULT_GROUPS_TAB_SETTINGS.cleanupFrequency),
    neverDeleteMixedSilently: merged.neverDeleteMixedSilently !== false,
    warnUnclassifiedEmails: merged.warnUnclassifiedEmails !== false,
    warnMixedCases: merged.warnMixedCases !== false,
    warningFrequency: normalizeFrequency(merged.warningFrequency, DEFAULT_GROUPS_TAB_SETTINGS.warningFrequency),
    prepareTasksBridge: merged.prepareTasksBridge === true,
    migrationTarget: String(merged.migrationTarget || "").trim(),
    migrationMode: normalizeMigrationMode(merged.migrationMode),
    allowMoveExistingData: merged.allowMoveExistingData === true,
    strictMigrationSafety: merged.strictMigrationSafety !== false,
    mergeExistingData: merged.mergeExistingData === true,
    explorerServerPrimary: merged.explorerServerPrimary !== false,
    explorerOpenStoredAttachments: merged.explorerOpenStoredAttachments !== false,
    explorerGenerateReply: merged.explorerGenerateReply !== false,
    groupsVersion:
      String(merged.groupsVersion || DEFAULT_GROUPS_TAB_SETTINGS.groupsVersion).trim() ||
      DEFAULT_GROUPS_TAB_SETTINGS.groupsVersion,
    quickDiagnostic: "",
  };

  normalized.locationStatus = deriveGroupsLocationStatus(normalized);
  normalized.quickDiagnostic = deriveGroupsQuickDiagnostic(normalized);
  return normalized;
}
