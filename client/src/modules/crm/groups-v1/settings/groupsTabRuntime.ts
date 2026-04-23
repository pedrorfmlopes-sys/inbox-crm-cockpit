import type { IntermediateCase, IntermediateCaseSummary } from "../storage/intermediateCaseTypes";
import type { ResolvedIntermediateCaseStorage } from "../storage/resolveIntermediateCaseStorage";
import { resolveGroupStorageRuntime, type ResolvedGroupStorageRuntime } from "../storage/resolveStorageMode";
import {
  normalizeGroupsTabSettings,
  type GroupsSettingsFrequency,
  type GroupsTabSettings,
} from "./groupsTabSettings";
import {
  normalizeGroupStorageSettings,
  type GroupStorageSettings,
} from "../storage/settings";

const BYTES_PER_MB = 1024 * 1024;

export type GroupsTabRuntimeSettings = {
  tab: GroupsTabSettings;
  storage: GroupStorageSettings;
  storageRuntime: ResolvedGroupStorageRuntime;
};

export type GroupsTabAttachmentTarget = "server" | "outside" | "metadata_only";

export type GroupsTabAttachmentPolicy = {
  strategy: GroupsTabSettings["attachmentStrategy"];
  serverEnabled: boolean;
  outsideEnabled: boolean;
  serverLimitBytes: number;
  intermediateLimitBytes: number;
  ignoreInlineAttachments: boolean;
  requireImmediatePreview: boolean;
  showAttachmentMetadataOnServer: boolean;
  externalFolder: string;
  outsideProvider: string;
  outsideBasePath: string;
};

export type GroupsTabAttachmentDecision = {
  selectable: boolean;
  selectedByDefault: boolean;
  requiresImmediatePreview: boolean;
  target: GroupsTabAttachmentTarget;
  storageDecision: "pending" | "local" | "server" | "hybrid" | "metadata_only" | "skip_inline";
  includeMetadataOnServer: boolean;
  includeBinaryInPayload: boolean;
  storageProvider?: string;
  storageBasePath?: string;
  storagePathHint?: string;
};

export type GroupsTabStorageValidation = {
  available: boolean;
  blocked: boolean;
  warning: boolean;
  retrySuggested: boolean;
  reason: string;
};

export type GroupsTabWarningMessage = {
  kind: "mixed_case" | "unclassified" | "local_abandoned" | "storage_unavailable";
  message: string;
};

function clampBytesFromMb(value: unknown, fallbackMb: number, minMb: number, maxMb: number): number {
  const raw = Number(value);
  const safeMb = Number.isFinite(raw) && raw > 0
    ? Math.min(Math.max(raw, minMb), maxMb)
    : fallbackMb;
  return Math.round(safeMb * BYTES_PER_MB);
}

function normalizeText(value: unknown): string {
  return String(value || "").trim();
}

function getAgeInDays(isoValue: string | undefined, nowMs: number): number {
  const raw = normalizeText(isoValue);
  if (!raw) return 0;
  const parsed = Date.parse(raw);
  if (!Number.isFinite(parsed)) return 0;
  return Math.floor((nowMs - parsed) / (24 * 60 * 60 * 1000));
}

export function resolveGroupsTabRuntimeSettings(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null; storage?: Partial<GroupStorageSettings> | null } | null;
} | null): GroupsTabRuntimeSettings {
  const tab = normalizeGroupsTabSettings(settingsLike?.groups?.tab || null);
  const storage = normalizeGroupStorageSettings(settingsLike?.groups?.storage || null);
  const storageRuntime = resolveGroupStorageRuntime({ groups: { storage } });
  return { tab, storage, storageRuntime };
}

export function resolveGroupsTabAttachmentPolicy(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null; storage?: Partial<GroupStorageSettings> | null } | null;
} | null): GroupsTabAttachmentPolicy {
  const runtime = resolveGroupsTabRuntimeSettings(settingsLike);
  const tab = runtime.tab;
  const externalFolder = normalizeText(tab.externalAttachmentFolder);
  const outsideBasePath = externalFolder || normalizeText(runtime.storageRuntime.legacyBridge.baseFolderPath);
  const outsideProvider = externalFolder
    ? "local"
    : runtime.storageRuntime.legacyBridge.provider;
  return {
    strategy: tab.attachmentStrategy,
    serverEnabled: tab.saveAttachmentsOnServer,
    outsideEnabled: tab.saveAttachmentsOutsideServer,
    serverLimitBytes: clampBytesFromMb(tab.attachmentServerLimitMb, 10, 1, 2048),
    intermediateLimitBytes: clampBytesFromMb(tab.attachmentIntermediateLimitMb, 50, 1, 4096),
    ignoreInlineAttachments: runtime.storage.ignoreInlineAttachments,
    requireImmediatePreview: tab.requireImmediatePreview,
    showAttachmentMetadataOnServer: tab.showAttachmentMetadataOnServer,
    externalFolder,
    outsideProvider,
    outsideBasePath,
  };
}

export function resolveGroupsTabAttachmentDecision(
  input: {
    key?: string;
    name?: string;
    size?: number;
    isInline?: boolean;
    hasContent?: boolean;
  },
  settingsLike?: {
    groups?: { tab?: Partial<GroupsTabSettings> | null; storage?: Partial<GroupStorageSettings> | null } | null;
  } | null
): GroupsTabAttachmentDecision {
  const policy = resolveGroupsTabAttachmentPolicy(settingsLike);
  const size = Number(input.size || 0);
  const hasContent = input.hasContent === true;
  const storagePathHint = normalizeText(input.name || "");

  if (policy.ignoreInlineAttachments && input.isInline) {
    return {
      selectable: false,
      selectedByDefault: false,
      requiresImmediatePreview: false,
      target: "metadata_only",
      storageDecision: "skip_inline",
      includeMetadataOnServer: false,
      includeBinaryInPayload: false,
    };
  }

  const target = (() => {
    if (policy.strategy === "server") {
      if (policy.serverEnabled) return "server";
      if (policy.outsideEnabled) return "outside";
      return "metadata_only";
    }
    if (policy.strategy === "outside") {
      if (policy.outsideEnabled) return "outside";
      if (policy.serverEnabled) return "server";
      return "metadata_only";
    }
    if (policy.serverEnabled && size <= policy.serverLimitBytes) return "server";
    if (policy.outsideEnabled) return "outside";
    if (policy.serverEnabled) return "server";
    return "metadata_only";
  })();

  const exceedsIntermediateLimit = size > policy.intermediateLimitBytes;
  const selectedByDefault = hasContent && (!exceedsIntermediateLimit || policy.requireImmediatePreview);
  const includeMetadataOnServer = target === "server"
    ? true
    : policy.showAttachmentMetadataOnServer;
  const includeBinaryInPayload = hasContent && (
    target === "server"
      || (target === "outside" && Boolean(policy.outsideBasePath))
  );

  if (target === "outside") {
    return {
      selectable: true,
      selectedByDefault,
      requiresImmediatePreview: policy.requireImmediatePreview,
      target,
      storageDecision: includeBinaryInPayload ? "hybrid" : "metadata_only",
      includeMetadataOnServer,
      includeBinaryInPayload,
      storageProvider: includeBinaryInPayload ? policy.outsideProvider : undefined,
      storageBasePath: includeBinaryInPayload ? policy.outsideBasePath : undefined,
      storagePathHint: includeBinaryInPayload ? storagePathHint || undefined : undefined,
    };
  }

  if (target === "server") {
    return {
      selectable: true,
      selectedByDefault,
      requiresImmediatePreview: policy.requireImmediatePreview,
      target,
      storageDecision: hasContent ? "server" : "metadata_only",
      includeMetadataOnServer: true,
      includeBinaryInPayload: hasContent,
      storageProvider: "cloud",
      storageBasePath: "",
    };
  }

  return {
    selectable: true,
    selectedByDefault,
    requiresImmediatePreview: policy.requireImmediatePreview,
    target,
    storageDecision: "metadata_only",
    includeMetadataOnServer,
    includeBinaryInPayload: false,
  };
}

export function buildGroupsTabAttachmentStorageOptions(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null; storage?: Partial<GroupStorageSettings> | null } | null;
} | null): {
  attachmentStorageProvider: string;
  attachmentStorageBasePath: string;
} {
  const { storageRuntime } = resolveGroupsTabRuntimeSettings(settingsLike);
  const policy = resolveGroupsTabAttachmentPolicy(settingsLike);
  if (
    policy.outsideEnabled
    && policy.outsideBasePath
    && (policy.strategy === "outside" || policy.strategy === "by_size")
  ) {
    return {
      attachmentStorageProvider: policy.outsideProvider,
      attachmentStorageBasePath: policy.outsideBasePath,
    };
  }
  return {
    attachmentStorageProvider: storageRuntime.legacyBridge.provider,
    attachmentStorageBasePath: storageRuntime.legacyBridge.baseFolderPath,
  };
}

export function shouldReopenGroupsExistingCase(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).reopenExistingCase;
}

export function shouldRecreateGroupsIntermediateCopy(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).recreateIntermediateCopy;
}

export function shouldAutoCreateGroupsCase(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).autoCreateCaseOnNewEmail;
}

export function shouldPersistGroupsPrepareCase(args: {
  settingsLike?: { groups?: { tab?: Partial<GroupsTabSettings> | null } | null } | null;
  hasHydratedCase: boolean;
  hasLocalCheckpoint: boolean;
}): boolean {
  return shouldAutoCreateGroupsCase(args.settingsLike)
    || args.hasHydratedCase
    || args.hasLocalCheckpoint;
}

export function shouldProjectServerCopyIntoIntermediate(args: {
  settingsLike?: { groups?: { tab?: Partial<GroupsTabSettings> | null } | null } | null;
  hasHydratedCase: boolean;
}): boolean {
  return shouldRecreateGroupsIntermediateCopy(args.settingsLike) || args.hasHydratedCase;
}

export function shouldUsePrepareTasksBridge(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).prepareTasksBridge;
}

export function shouldUseExplorerServerPrimary(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).explorerServerPrimary;
}

export function canOpenStoredAttachmentsFromGroups(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).explorerOpenStoredAttachments;
}

export function canGenerateReplyFromGroups(settingsLike?: {
  groups?: { tab?: Partial<GroupsTabSettings> | null } | null;
} | null): boolean {
  return normalizeGroupsTabSettings(settingsLike?.groups?.tab || null).explorerGenerateReply;
}

export async function validateGroupsTabStorageAvailability(args: {
  settings: GroupsTabSettings;
  storage: ResolvedIntermediateCaseStorage;
}): Promise<GroupsTabStorageValidation> {
  const settings = normalizeGroupsTabSettings(args.settings);
  if (!settings.validateLocationOnOpen) {
    return {
      available: true,
      blocked: false,
      warning: false,
      retrySuggested: false,
      reason: "Validacao ao abrir desativada por configuracao.",
    };
  }
  if (args.storage.availability === "disabled") {
    return {
      available: false,
      blocked: settings.blockTabIfUnavailable,
      warning: settings.warnIfUnavailable,
      retrySuggested: false,
      reason: args.storage.reason,
    };
  }
  if (args.storage.availability === "missing_location") {
    return {
      available: false,
      blocked: settings.blockTabIfUnavailable,
      warning: settings.warnIfUnavailable,
      retrySuggested: false,
      reason: "Sem namespace persistente configurado para o storage intermédio.",
    };
  }
  try {
    await args.storage.repository.listCases();
    return {
      available: true,
      blocked: false,
      warning: false,
      retrySuggested: false,
      reason: args.storage.reason,
    };
  } catch (error) {
    const reason = normalizeText((error as Error)?.message) || "Falha a validar o IndexedDB do modulo Groups.";
    return {
      available: false,
      blocked: settings.blockTabIfUnavailable,
      warning: settings.warnIfUnavailable,
      retrySuggested: settings.autoRetryValidation,
      reason,
    };
  }
}

export function isGroupsTabFrequencyDue(
  frequency: GroupsSettingsFrequency,
  lastRunIso: string | null | undefined,
  nowMs: number = Date.now()
): boolean {
  if (frequency === "manual") return false;
  const lastRun = Date.parse(normalizeText(lastRunIso));
  if (!Number.isFinite(lastRun)) return true;
  const intervalMs = frequency === "weekly"
    ? 7 * 24 * 60 * 60 * 1000
    : 24 * 60 * 60 * 1000;
  return nowMs - lastRun >= intervalMs;
}

export function buildGroupsTabWarningMessages(args: {
  settings: GroupsTabSettings;
  caseValue?: IntermediateCase | null;
  summary?: IntermediateCaseSummary | null;
  validation?: GroupsTabStorageValidation | null;
  nowMs?: number;
}): GroupsTabWarningMessage[] {
  const settings = normalizeGroupsTabSettings(args.settings);
  const nowMs = Number.isFinite(Number(args.nowMs)) ? Number(args.nowMs) : Date.now();
  const messages: GroupsTabWarningMessage[] = [];
  const summary = args.summary || null;
  const caseValue = args.caseValue || null;

  if (args.validation && !args.validation.available && args.validation.warning) {
    messages.push({
      kind: "storage_unavailable",
      message: args.validation.reason,
    });
  }

  if (caseValue?.classificationSummary?.mixedCase && settings.warnMixedCases) {
    const ageDays = getAgeInDays(caseValue.lastAccessedAt || caseValue.updatedAt, nowMs);
    if (ageDays >= settings.mixedCaseWarningDays) {
      messages.push({
        kind: "mixed_case",
        message: `O caso atual continua misto ha ${ageDays} dia(s); convem fechar a classificacao pendente.`,
      });
    }
  }

  if (caseValue?.classificationSummary?.unclassifiedEmails > 0 && settings.warnUnclassifiedEmails) {
    messages.push({
      kind: "unclassified",
      message: `O caso atual ainda tem ${caseValue.classificationSummary.unclassifiedEmails} email(s) por classificar.`,
    });
  }

  if (summary?.retentionState === "local_only") {
    const ageDays = getAgeInDays(summary.lastAccessedAt || summary.updatedAt, nowMs);
    if (ageDays >= settings.localAbandonedWarningDays) {
      messages.push({
        kind: "local_abandoned",
        message: `Existe trabalho local abandonado ha ${ageDays} dia(s) neste namespace.`,
      });
    }
  }

  return messages;
}
