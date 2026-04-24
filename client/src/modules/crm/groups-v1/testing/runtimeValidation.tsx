import React from "react";
import ReactDOM from "react-dom/client";
import {
  getRelatedEmailContext,
  type GroupTicketEntry,
  type LinkGroupEntry,
  type RelatedEmailEntry,
  type RelevantEmailPayload,
} from "@/api";
import {
  getSettings,
  resetSettings,
  saveSettings,
  type CockpitSettingsV1,
} from "@/settings";
import { buildResolvedRemoteApplyExecutionPlan, buildResolvedStudioApplySelection } from "@/modules/crm/group-classification/applyResolution";
import {
  buildAttachmentStorageOptions,
  buildRelevantEmailPayloadFromRelatedEmail,
  makeEmailKey,
} from "@/modules/crm/group-classification/documentUtils";
import {
  executeLegacyBaseTicketApply,
  executeLegacyRemoteApplyForTarget,
} from "@/modules/crm/group-classification/legacyRemoteApply";
import { projectApplyIntoIntermediateCase } from "@/modules/crm/group-classification/localCaseProjection";
import { persistAndRefreshClassificationCase } from "@/modules/crm/group-classification/casePersistence";
import type { ClassificationMetaDraft } from "@/modules/crm/group-classification/types";
import { buildOutlookCategoryPlan, buildOutlookCategorySourceFromRelatedContext } from "@/outlookCategories";
import { GroupsSettingsPanel } from "../settings/GroupsSettingsPanel";
import { DEFAULT_GROUPS_MODULE_SETTINGS, type GroupsModuleSettings } from "../settings/groupsModuleSettings";
import { DEFAULT_GROUPS_TAB_SETTINGS, normalizeGroupsTabSettings } from "../settings/groupsTabSettings";
import {
  buildGroupsTabAttachmentStorageOptions,
  buildGroupsTabWarningMessages,
  canGenerateReplyFromGroups,
  canOpenStoredAttachmentsFromGroups,
  isGroupsTabFrequencyDue,
  resolveGroupsTabAttachmentDecision,
  shouldPersistGroupsPrepareCase,
  shouldProjectServerCopyIntoIntermediate,
  shouldUseExplorerServerPrimary,
  shouldUsePrepareTasksBridge,
  validateGroupsTabStorageAvailability,
} from "../settings/groupsTabRuntime";
import { resolveGroupAttachmentStoragePolicy, resolvePreparedAttachmentStorageDecision } from "../storage/attachmentPolicy";
import { buildPrepareWorksetManifest } from "../storage/buildPrepareWorksetManifest";
import {
  cleanupIntermediateCases,
  migrateIntermediateCaseNamespace,
} from "../storage/intermediateCaseMaintenance";
import { hydrateIntermediateCaseEmailsToRelatedEntries } from "../storage/intermediateCaseAdapters";
import { INTERMEDIATE_CASE_DB_NAME } from "../storage/intermediateCaseIndexedDbAdapter";
import { buildPrepareIntermediateCaseFromSources } from "../storage/prepareIntermediateCaseResolution";
import {
  resolveClassificationIntermediateCase,
} from "../storage/resolveClassificationIntermediateCase";
import { resolveIntermediateCaseStorage } from "../storage/resolveIntermediateCaseStorage";
import { resolveGroupStorageRuntime } from "../storage/resolveStorageMode";
import { savePrimaryGroupWorkset } from "../storage/saveWorkset";
import { loadPrimaryGroupWorkset } from "../storage/loadWorkset";
import { DEFAULT_GROUP_STORAGE_SETTINGS } from "../storage/settings";
import type { IntermediateCase } from "../storage/intermediateCaseTypes";
import type { GroupWorksetManifest } from "../storage/types";
import { GROUPS_SETTINGS_MATRIX, type GroupsSettingsMatrixEntry } from "./settingsMatrix";

type ValidationArea =
  | "settings"
  | "prepare"
  | "classify"
  | "storage"
  | "migration"
  | "cleanup"
  | "attachments"
  | "outlook_categories";

type ValidationStatus = "passed" | "failed" | "corrected" | "pending";

type ValidationScenarioResult = {
  id: string;
  area: ValidationArea;
  title: string;
  status: ValidationStatus;
  details: string;
};

type MockFetchCall = {
  url: string;
  method: string;
  body?: string;
  responseStatus?: number;
  responseBody?: unknown;
};

type FinalWriteProofSnapshot = {
  emails: Array<{
    emailKey: string;
    subject?: string;
    labels: string[];
    principalGroupId?: string;
    referenceGroupIds: string[];
    ticketIds: string[];
    attachments: Array<{
      attachmentKey: string;
      name: string;
      documentState?: string;
      isHidden?: boolean;
      storageDecision?: string;
      hasContent?: boolean;
      storageProvider?: string;
      storageBasePath?: string;
      storagePathHint?: string;
    }>;
  }>;
  groups: Array<{
    id: string;
    name: string;
    memberEmailKeys: string[];
  }>;
  tickets: Array<{
    id: string;
    code: string;
    status?: string;
    groupIds: string[];
    emailKeys: string[];
  }>;
};

type FinalWriteProofArtifact = {
  scenarioId: string;
  title: string;
  inputContext: Record<string, unknown>;
  relevantSettings: Record<string, unknown>;
  intermediateState: {
    beforeWrite: unknown;
    afterProjection: unknown;
    reopenedLocalCase: unknown;
  };
  finalPayloads: Array<{
    url: string;
    method: string;
    body: unknown;
  }>;
  backendResponses: Array<{
    url: string;
    method: string;
    status: number;
    body: unknown;
  }>;
  persistedBackendState: FinalWriteProofSnapshot;
  reopenedState: {
    serverContext: unknown;
    refreshedContext: unknown;
  };
  expected: Record<string, unknown>;
  actual: Record<string, unknown>;
  pass: boolean;
  failureReason?: string;
};

export type GroupsBrowserValidationReport = {
  generatedAtIso: string;
  settingsMatrix: GroupsSettingsMatrixEntry[];
  scenarios: ValidationScenarioResult[];
  writeProofs: FinalWriteProofArtifact[];
  passed: number;
  failed: number;
};

type RelatedEmailAttachmentFixture = NonNullable<RelatedEmailEntry["attachments"]>[number];
type RelatedEmailGroupFixture = NonNullable<RelatedEmailEntry["relatedGroups"]>[number];

const FIXED_NOW_ISO = "2026-04-23T12:00:00.000Z";
const FIXED_NOW_MS = Date.parse(FIXED_NOW_ISO);

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function buildGroupsSettings(patch?: Partial<GroupsModuleSettings>): GroupsModuleSettings {
  return {
    ...DEFAULT_GROUPS_MODULE_SETTINGS,
    ...patch,
    storage: {
      ...DEFAULT_GROUPS_MODULE_SETTINGS.storage,
      ...(patch?.storage || {}),
      localDevice: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.storage.localDevice,
        ...(patch?.storage?.localDevice || {}),
      },
      chosenFolder: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.storage.chosenFolder,
        ...(patch?.storage?.chosenFolder || {}),
      },
      supabase: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.storage.supabase,
        ...(patch?.storage?.supabase || {}),
      },
      hybrid: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.storage.hybrid,
        ...(patch?.storage?.hybrid || {}),
      },
    },
    tab: normalizeGroupsTabSettings({
      ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
      ...(patch?.tab || {}),
    }),
    labels: {
      ...DEFAULT_GROUPS_MODULE_SETTINGS.labels,
      ...(patch?.labels || {}),
      catalog: patch?.labels?.catalog || DEFAULT_GROUPS_MODULE_SETTINGS.labels.catalog,
      favoriteIds: patch?.labels?.favoriteIds || DEFAULT_GROUPS_MODULE_SETTINGS.labels.favoriteIds,
    },
    tickets: {
      ...DEFAULT_GROUPS_MODULE_SETTINGS.tickets,
      ...(patch?.tickets || {}),
      ui: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.tickets.ui,
        ...(patch?.tickets?.ui || {}),
      },
    },
    outlookCategories: {
      ...DEFAULT_GROUPS_MODULE_SETTINGS.outlookCategories,
      ...(patch?.outlookCategories || {}),
    },
  };
}

function buildRelatedEmail(overrides?: Partial<RelatedEmailEntry>): RelatedEmailEntry {
  const conversationId = String(overrides?.conversationId || "conv-anchor");
  const subject = String(overrides?.subject || "Assunto Base");
  const fromEmail = String(overrides?.fromEmail || "cliente@example.com");
  const receivedAtIso = String(overrides?.messageDateIso || overrides?.receivedAtIso || FIXED_NOW_ISO);
  const itemId = String(overrides?.itemId || "");
  const internetMessageId = String(overrides?.internetMessageId || "");
  const emailKey = String(overrides?.emailKey || itemId || internetMessageId || `${conversationId}|${subject}|${fromEmail}|${receivedAtIso}`);
  return {
    emailKey,
    itemId: itemId || undefined,
    internetMessageId: internetMessageId || undefined,
    conversationId,
    subject,
    fromEmail,
    fromName: String(overrides?.fromName || "Cliente"),
    receivedAtIso,
    messageDateIso: receivedAtIso,
    bodyText: String(overrides?.bodyText || "Corpo de teste"),
    bodyHtml: String(overrides?.bodyHtml || "<p>Corpo de teste</p>"),
    attachments: overrides?.attachments || [],
    relatedGroups: overrides?.relatedGroups || [],
    relatedReasons: overrides?.relatedReasons || [],
    labels: overrides?.labels || [],
    removedInheritedLabels: overrides?.removedInheritedLabels || [],
    labelStates: overrides?.labelStates || {},
    classificationMeta: overrides?.classificationMeta || {},
    toRecipients: overrides?.toRecipients || [{ email: "to@example.com", name: "To" }],
    ccRecipients: overrides?.ccRecipients || [{ email: "cc@example.com", name: "Cc" }],
    ...overrides,
  } as RelatedEmailEntry;
}

function buildPrepareEmailInput(overrides?: {
  emailKey?: string;
  subject?: string;
  sourceOrigin?: "server" | "intermediate" | "outlook";
  principalGroupId?: string;
  referenceGroupIds?: string[];
  labels?: string[];
  ticketIds?: string[];
  attachments?: Array<{ key: string; name: string; hasContent?: boolean; documentState?: string }>;
}): Parameters<typeof buildPrepareIntermediateCaseFromSources>[0]["outlookEmails"][number] {
  const emailKey = String(overrides?.emailKey || "msg-anchor");
  return {
    emailKey,
    itemId: `${emailKey}-item`,
    internetMessageId: `<${emailKey}@example.com>`,
    conversationId: "conv-prepare",
    subject: overrides?.subject || "Prepare fixture",
    fromEmail: "cliente@example.com",
    fromName: "Cliente",
    to: ["to@example.com"],
    cc: ["cc@example.com"],
    receivedAtIso: FIXED_NOW_ISO,
    bodyText: "Corpo fixture",
    bodyHtml: "<p>Corpo fixture</p>",
    sourceOrigin: overrides?.sourceOrigin || "outlook",
    visibilityState: "draft",
    serverPresence: "none",
    localPresence: "case_only",
    classification: {
      principalGroupId: overrides?.principalGroupId,
      principalGroupName: overrides?.principalGroupId ? `Grupo ${overrides.principalGroupId}` : undefined,
      referenceGroupIds: overrides?.referenceGroupIds || [],
      labels: overrides?.labels || [],
      removedInheritedLabels: [],
      labelStates: {},
      categorizedLabelNames: overrides?.labels || [],
      ticketIds: overrides?.ticketIds || [],
      ticketCodes: (overrides?.ticketIds || []).map((ticketId) => `TK-${ticketId}`),
    },
    attachments: (overrides?.attachments || []).map((attachment, index) => ({
      attachmentKey: attachment.key,
      id: `att-${index + 1}`,
      name: attachment.name,
      hasContent: attachment.hasContent !== false,
      documentState: attachment.documentState || "ingested",
      selected: true,
      contentBase64: attachment.hasContent === false ? undefined : "ZmFrZQ==",
    })),
  };
}

function buildGroup(overrides?: Partial<LinkGroupEntry>): LinkGroupEntry {
  return {
    id: String(overrides?.id || "grp-main"),
    name: String(overrides?.name || "Grupo Principal"),
    kind: String(overrides?.kind || "custom"),
    status: String(overrides?.status || "em_analise"),
    labels: overrides?.labels || ["Urgente"],
    ...overrides,
  } as LinkGroupEntry;
}

function buildTicket(overrides?: Partial<GroupTicketEntry>): GroupTicketEntry {
  return {
    id: String(overrides?.id || "ticket-1"),
    code: String(overrides?.code || "TK-001"),
    status: String(overrides?.status || "em_analise"),
    groupId: String(overrides?.groupId || "grp-main"),
    ...overrides,
  } as GroupTicketEntry;
}

async function resetBrowserPersistence(): Promise<void> {
  try {
    localStorage.clear();
  } catch {
    // ignore
  }
  try {
    sessionStorage.clear();
  } catch {
    // ignore
  }
  await new Promise<void>((resolve) => {
    const request = indexedDB.deleteDatabase(INTERMEDIATE_CASE_DB_NAME);
    request.onsuccess = () => resolve();
    request.onerror = () => resolve();
    request.onblocked = () => resolve();
  });
}

async function withIndexedDbAvailabilityOverride<T>(
  nextValue: IDBFactory | undefined,
  run: () => Promise<T>
): Promise<T> {
  const ownDescriptor = Object.getOwnPropertyDescriptor(globalThis, "indexedDB");
  Object.defineProperty(globalThis, "indexedDB", {
    configurable: true,
    enumerable: ownDescriptor?.enumerable ?? true,
    writable: true,
    value: nextValue,
  });
  try {
    return await run();
  } finally {
    if (ownDescriptor) {
      Object.defineProperty(globalThis, "indexedDB", ownDescriptor);
    } else {
      delete (globalThis as Record<string, unknown>).indexedDB;
    }
  }
}

async function renderSettingsPanelScenario(): Promise<void> {
  const originalPrompt = window.prompt;
  const originalAlert = window.alert;
  const originalConfirm = window.confirm;
  window.prompt = (() => {
    throw new Error("window.prompt should not be called");
  }) as typeof window.prompt;
  window.alert = (() => {
    throw new Error("window.alert should not be called");
  }) as typeof window.alert;
  window.confirm = (() => {
    throw new Error("window.confirm should not be called");
  }) as typeof window.confirm;

  const host = document.createElement("div");
  document.body.appendChild(host);
  const root = ReactDOM.createRoot(host);

  try {
    root.render(
      <GroupsSettingsPanel
        open
        value={normalizeGroupsTabSettings(DEFAULT_GROUPS_TAB_SETTINGS)}
        onClose={() => undefined}
        onSave={() => undefined}
      />
    );
    await new Promise((resolve) => setTimeout(resolve, 20));

    const storageButton = Array.from(host.querySelectorAll("button")).find((button) =>
      button.textContent?.includes("Storage intermedio")
    );
    assert(storageButton, "A secao de storage intermedio nao foi renderizada.");
    storageButton?.click();
    await new Promise((resolve) => setTimeout(resolve, 20));

    const locationInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: C:/dados/grupos/intermedio"]');
    assert(locationInput, "O editor inline da pasta intermédia nao foi renderizado.");
    locationInput.value = "C:/dados/grupos/runtime";
    locationInput.dispatchEvent(new Event("input", { bubbles: true }));
    locationInput.dispatchEvent(new Event("change", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    const updatedLocationInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: C:/dados/grupos/intermedio"]');
    assert(updatedLocationInput?.value === "C:/dados/grupos/runtime", "A pasta intermédia inline nao refletiu a edicao local.");

    const enabledToggles = Array.from(host.querySelectorAll('button[aria-pressed]')).filter((button) => !button.hasAttribute("disabled"));
    assert(enabledToggles.length >= 7, "Os toggles executaveis da secao intermedia nao ficaram interativos.");
    enabledToggles[0]?.click();
    await new Promise((resolve) => setTimeout(resolve, 20));

    const attachmentsButton = Array.from(host.querySelectorAll("button")).find((button) =>
      button.textContent?.includes("Anexos")
    );
    assert(attachmentsButton, "A secao de anexos nao foi renderizada.");
    attachmentsButton?.click();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const externalPathInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: C:/dados/grupos/anexos"]');
    assert(externalPathInput, "O campo do destino externo nao ficou editavel.");

    const exploreButton = Array.from(host.querySelectorAll("button")).find((button) =>
      button.textContent?.includes("Explorar")
    );
    assert(exploreButton, "A secao de bridges internas nao foi renderizada.");
    exploreButton?.click();
    await new Promise((resolve) => setTimeout(resolve, 20));
    const enabledExploreToggles = Array.from(host.querySelectorAll('button[aria-pressed]')).filter((button) => !button.hasAttribute("disabled"));
    assert(enabledExploreToggles.length >= 3, "Os toggles explorer* nao ficaram interativos.");

    const migrationButton = Array.from(host.querySelectorAll("button")).find((button) =>
      button.textContent?.includes("Migracao")
    );
    assert(migrationButton, "A secao de migracao nao foi renderizada.");
    migrationButton?.click();
    await new Promise((resolve) => setTimeout(resolve, 20));

    const migrationSelect = Array.from(host.querySelectorAll("select")).find((select) =>
      Array.from(select.querySelectorAll("option")).some((option) => option.value === "always_ask")
    );
    assert(migrationSelect, "O seletor de migracao nao foi renderizado.");
    const askOption = Array.from(migrationSelect.querySelectorAll("option")).find((option) => option.value === "always_ask");
    assert(askOption?.disabled, "A opcao 'always_ask' devia estar explicitamente indisponivel.");
  } finally {
    root.unmount();
    host.remove();
    window.prompt = originalPrompt;
    window.alert = originalAlert;
    window.confirm = originalConfirm;
  }
}

function buildManifestForRuntime(runtimeMode: ReturnType<typeof resolveGroupStorageRuntime>["mode"]): GroupWorksetManifest {
  const runtime = resolveGroupStorageRuntime({
    ...DEFAULT_GROUP_STORAGE_SETTINGS,
    mode: runtimeMode,
    localDevice: { rootPath: "C:/tmp/groups" },
    chosenFolder: { path: "C:/tmp/groups/chosen", kind: "filesystem" },
    hybrid: { primaryTarget: "chosen_folder", promoteManifestOnSave: true, promoteAttachmentMetadataOnSave: false },
  });
  const manifest = buildPrepareWorksetManifest({
    anchorEmailKey: "anchor@email|base",
    settings: runtime.settings,
    runtime,
    selectedEmailKeys: ["anchor@email|base", "known@email|related"],
    selectedAttachmentKeys: ["doc-1"],
    attachmentRows: [
      {
        key: "doc-1",
        emailKey: "anchor@email|base",
        name: "doc.pdf",
        contentType: "application/pdf",
        size: 1024,
        hasContent: true,
      },
    ],
    workingGroupId: "grp-main",
    workingGroupName: "Grupo Main",
    filterQuery: "cliente",
    attachmentMode: "with",
    groupMode: "with_group",
  });
  assert(manifest, `Nao foi possivel construir manifesto de workset para ${runtimeMode}.`);
  return manifest!;
}

type MockServerTicketState = {
  ticket: GroupTicketEntry;
  emailKeys: Set<string>;
  groupIds: Set<string>;
};

type MockGroupsServerState = {
  emailsByKey: Map<string, RelatedEmailEntry>;
  groupsById: Map<string, LinkGroupEntry>;
  groupMembersById: Map<string, Set<string>>;
  ticketsById: Map<string, MockServerTicketState>;
  nextTicketSequence: number;
};

type MockFetchHarness = Array<MockFetchCall> & {
  serverState: MockGroupsServerState;
  pickedFolderPath: string;
  getServerSnapshot: () => FinalWriteProofSnapshot;
};

function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value));
}

function normalizeString(value: unknown): string {
  return String(value || "").trim();
}

function normalizeStringList(values: unknown[]): string[] {
  return Array.from(new Set(values.map((value) => normalizeString(value)).filter(Boolean)));
}

function normalizeRecipientEntries(values: RelevantEmailPayload["toRecipients"] | RelevantEmailPayload["ccRecipients"]): Array<{ email: string; name?: string }> {
  if (!Array.isArray(values)) return [];
  return values
    .map((recipient) => ({
      email: normalizeString(recipient?.email).toLowerCase(),
      name: normalizeString(recipient?.name) || undefined,
    }))
    .filter((recipient) => recipient.email);
}

function buildMockGroup(groupId: string, groupName?: string, relationKind?: string): LinkGroupEntry {
  return buildGroup({
    id: groupId,
    name: groupName || groupId,
    kind: "custom",
    status: "em_analise",
    labels: [],
    ...(relationKind ? { relationKind } : {}),
  } as LinkGroupEntry);
}

function buildMockServerState(): MockGroupsServerState {
  return {
    emailsByKey: new Map(),
    groupsById: new Map(),
    groupMembersById: new Map(),
    ticketsById: new Map(),
    nextTicketSequence: 1,
  };
}

function ensureMockGroup(serverState: MockGroupsServerState, groupId: string, groupName?: string): LinkGroupEntry {
  const normalizedGroupId = normalizeString(groupId);
  const existing = serverState.groupsById.get(normalizedGroupId);
  if (existing) return existing;
  const next = buildMockGroup(normalizedGroupId, groupName);
  serverState.groupsById.set(normalizedGroupId, next);
  serverState.groupMembersById.set(normalizedGroupId, new Set());
  return next;
}

function normalizeMockAttachments(
  attachments: RelevantEmailPayload["attachments"],
  existing: NonNullable<RelatedEmailEntry["attachments"]> = [],
  defaults?: {
    storageProvider?: string;
    storageBasePath?: string;
  }
): NonNullable<RelatedEmailEntry["attachments"]> {
  const existingByKey = new Map(
    existing.map((attachment) => [normalizeString(attachment.key || attachment.id || attachment.name), attachment] as const)
  );
  return (attachments || []).map((attachment, index) => {
    const attachmentKey = normalizeString(attachment?.key || attachment?.id || attachment?.name || `attachment-${index + 1}`);
    const current = existingByKey.get(attachmentKey);
    return {
      key: attachmentKey,
      id: normalizeString(attachment?.id) || current?.id || attachmentKey,
      name: normalizeString(attachment?.name) || current?.name || attachmentKey,
      contentType: normalizeString(attachment?.contentType) || current?.contentType || "application/octet-stream",
      size: Number(attachment?.size || current?.size || 0) || 0,
      isInline: attachment?.isInline === true || current?.isInline === true,
      contentId: normalizeString(attachment?.contentId) || current?.contentId || undefined,
      content: normalizeString(attachment?.content) || current?.content || undefined,
      storageProvider:
        normalizeString(attachment?.storageProvider) ||
        normalizeString(defaults?.storageProvider) ||
        current?.storageProvider ||
        undefined,
      storageBasePath:
        normalizeString(attachment?.storageBasePath) ||
        normalizeString(defaults?.storageBasePath) ||
        current?.storageBasePath ||
        undefined,
      storagePathHint: normalizeString(attachment?.storagePathHint) || current?.storagePathHint || undefined,
      documentState: normalizeString(attachment?.documentState) || current?.documentState || "ingested",
      hasContent: attachment?.hasContent === true || Boolean(normalizeString(attachment?.content)) || current?.hasContent === true,
      isHidden: typeof attachment?.isHidden === "boolean" ? attachment.isHidden : current?.isHidden,
    };
  });
}

function upsertMockServerEmail(
  serverState: MockGroupsServerState,
  payload: RelevantEmailPayload
): RelatedEmailEntry {
  const emailKey = makeEmailKey(payload as RelatedEmailEntry);
  const existing = serverState.emailsByKey.get(emailKey);
  const next: RelatedEmailEntry = {
    ...(existing || {}),
    emailKey,
    itemId: normalizeString(payload.itemId) || existing?.itemId,
    internetMessageId: normalizeString(payload.internetMessageId) || existing?.internetMessageId,
    conversationId: normalizeString(payload.conversationId) || existing?.conversationId,
    subject: normalizeString(payload.subject) || existing?.subject || "(sem assunto)",
    fromEmail: normalizeString(payload.fromEmail) || existing?.fromEmail || undefined,
    fromName: normalizeString(payload.fromName) || existing?.fromName || undefined,
    emailWebLink: normalizeString(payload.emailWebLink) || existing?.emailWebLink || undefined,
    sentAtIso: normalizeString(payload.sentAtIso) || existing?.sentAtIso || undefined,
    receivedAtIso: normalizeString(payload.receivedAtIso) || normalizeString(payload.messageDateIso) || existing?.receivedAtIso || FIXED_NOW_ISO,
    messageDateIso: normalizeString(payload.messageDateIso) || normalizeString(payload.receivedAtIso) || existing?.messageDateIso || FIXED_NOW_ISO,
    toRecipients: normalizeRecipientEntries(payload.toRecipients).length
      ? normalizeRecipientEntries(payload.toRecipients)
      : (existing?.toRecipients || []),
    ccRecipients: normalizeRecipientEntries(payload.ccRecipients).length
      ? normalizeRecipientEntries(payload.ccRecipients)
      : (existing?.ccRecipients || []),
    bodyText: normalizeString(payload.bodyText) || existing?.bodyText || "",
    bodyHtml: normalizeString(payload.bodyHtml) || existing?.bodyHtml || "",
    status: normalizeString(payload.status) || existing?.status || undefined,
    labels: payload.labels ? normalizeStringList(payload.labels) : (existing?.labels || []),
    removedInheritedLabels: payload.removedInheritedLabels
      ? normalizeStringList(payload.removedInheritedLabels)
      : (existing?.removedInheritedLabels || []),
    labelStates: payload.labelStates
      ? Object.fromEntries(
          Object.entries(payload.labelStates)
            .map(([label, status]) => [normalizeString(label), normalizeString(status)])
            .filter(([label, status]) => label && status)
        )
      : (existing?.labelStates || {}),
    classificationMeta: payload.classificationMeta
      ? cloneJson(payload.classificationMeta)
      : (existing?.classificationMeta || {}),
    relatedGroups: existing?.relatedGroups || [],
    relatedReasons: existing?.relatedReasons || [],
    attachments: payload.attachments
      ? normalizeMockAttachments(payload.attachments, existing?.attachments || [], {
          storageProvider: normalizeString(payload.attachmentStorageProvider),
          storageBasePath: normalizeString(payload.attachmentStorageBasePath),
        })
      : (existing?.attachments || []),
  };
  serverState.emailsByKey.set(emailKey, next);
  return next;
}

function applyMockGroupMembership(args: {
  serverState: MockGroupsServerState;
  email: RelatedEmailEntry;
  groupId: string;
  groupName?: string;
  relationKind: string;
}): RelatedEmailEntry {
  const group = ensureMockGroup(args.serverState, args.groupId, args.groupName);
  const currentEmail = args.serverState.emailsByKey.get(args.email.emailKey || "") || args.email;
  const nextGroups = [
    ...(currentEmail.relatedGroups || []).filter((entry) => normalizeString(entry.id) !== group.id),
    {
      id: group.id,
      name: group.name,
      kind: group.kind,
      relationKind: args.relationKind,
    },
  ];
  args.serverState.groupMembersById.get(group.id)?.add(currentEmail.emailKey || "");
  const nextEmail = {
    ...currentEmail,
    relatedGroups: nextGroups,
  };
  args.serverState.emailsByKey.set(currentEmail.emailKey || "", nextEmail);
  return nextEmail;
}

function removeMockGroupMembership(args: {
  serverState: MockGroupsServerState;
  email: RelatedEmailEntry;
  groupId: string;
}): RelatedEmailEntry {
  const normalizedGroupId = normalizeString(args.groupId);
  const currentEmail = args.serverState.emailsByKey.get(args.email.emailKey || "") || args.email;
  const nextEmail = {
    ...currentEmail,
    relatedGroups: (currentEmail.relatedGroups || []).filter((entry) => normalizeString(entry.id) !== normalizedGroupId),
  };
  args.serverState.groupMembersById.get(normalizedGroupId)?.delete(currentEmail.emailKey || "");
  args.serverState.emailsByKey.set(currentEmail.emailKey || "", nextEmail);
  return nextEmail;
}

function ensureMockTicket(
  serverState: MockGroupsServerState,
  ticketId: string,
  overrides?: Partial<GroupTicketEntry>
): MockServerTicketState {
  const normalizedTicketId = normalizeString(ticketId);
  const existing = serverState.ticketsById.get(normalizedTicketId);
  if (existing) return existing;
  const nextTicket = buildTicket({
    id: normalizedTicketId,
    code: overrides?.code || `TK-${String(serverState.nextTicketSequence).padStart(3, "0")}`,
    sequenceNumber: serverState.nextTicketSequence,
    seriesId: overrides?.seriesId || "series-default",
    title: overrides?.title || `Ticket ${serverState.nextTicketSequence}`,
    status: overrides?.status || "em_analise",
    groupIds: overrides?.groupIds || [],
    ...overrides,
  });
  serverState.nextTicketSequence += 1;
  const state: MockServerTicketState = {
    ticket: nextTicket,
    emailKeys: new Set(),
    groupIds: new Set(normalizeStringList(nextTicket.groupIds || [])),
  };
  serverState.ticketsById.set(normalizedTicketId, state);
  return state;
}

function buildMockTicketResponse(ticketState: MockServerTicketState, emailLinked?: boolean): GroupTicketEntry {
  return {
    ...ticketState.ticket,
    groupIds: Array.from(ticketState.groupIds),
    emailCount: ticketState.emailKeys.size,
    emailLinked,
  };
}

function findMockServerEmail(serverState: MockGroupsServerState, payload: Partial<RelevantEmailPayload>): RelatedEmailEntry | null {
  const emailKey = makeEmailKey(payload as RelatedEmailEntry);
  if (emailKey && serverState.emailsByKey.has(emailKey)) {
    return serverState.emailsByKey.get(emailKey) || null;
  }
  for (const email of serverState.emailsByKey.values()) {
    if (normalizeString(payload.itemId) && normalizeString(payload.itemId) === normalizeString(email.itemId)) return email;
    if (normalizeString(payload.internetMessageId) && normalizeString(payload.internetMessageId) === normalizeString(email.internetMessageId)) return email;
    if (
      normalizeString(payload.conversationId)
      && normalizeString(payload.subject)
      && normalizeString(payload.fromEmail).toLowerCase()
      && normalizeString(payload.receivedAtIso || payload.messageDateIso)
      && normalizeString(payload.conversationId) === normalizeString(email.conversationId)
      && normalizeString(payload.subject) === normalizeString(email.subject)
      && normalizeString(payload.fromEmail).toLowerCase() === normalizeString(email.fromEmail).toLowerCase()
      && normalizeString(payload.receivedAtIso || payload.messageDateIso) === normalizeString(email.receivedAtIso || email.messageDateIso)
    ) {
      return email;
    }
  }
  return null;
}

function buildMockServerSnapshot(serverState: MockGroupsServerState): FinalWriteProofSnapshot {
  const emails = Array.from(serverState.emailsByKey.values())
    .map((email) => ({
      emailKey: normalizeString(email.emailKey),
      subject: email.subject,
      labels: normalizeStringList(email.labels || []),
      principalGroupId: (email.relatedGroups || []).find((group) => group.relationKind === "principal")?.id,
      referenceGroupIds: (email.relatedGroups || [])
        .filter((group) => group.relationKind !== "principal")
        .map((group) => normalizeString(group.id))
        .filter(Boolean),
      ticketIds: Array.from(serverState.ticketsById.values())
        .filter((ticketState) => ticketState.emailKeys.has(normalizeString(email.emailKey)))
        .map((ticketState) => normalizeString(ticketState.ticket.id)),
      attachments: (email.attachments || []).map((attachment) => ({
        attachmentKey: normalizeString(attachment.key || attachment.id || attachment.name),
        name: normalizeString(attachment.name),
        documentState: normalizeString(attachment.documentState) || undefined,
        isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : undefined,
        hasContent: attachment.hasContent === true,
        storageProvider: normalizeString(attachment.storageProvider) || undefined,
        storageBasePath: normalizeString(attachment.storageBasePath) || undefined,
        storagePathHint: normalizeString(attachment.storagePathHint) || undefined,
      })),
    }))
    .sort((left, right) => left.emailKey.localeCompare(right.emailKey));
  const groups = Array.from(serverState.groupsById.values())
    .map((group) => ({
      id: normalizeString(group.id),
      name: normalizeString(group.name),
      memberEmailKeys: Array.from(serverState.groupMembersById.get(normalizeString(group.id)) || []).sort(),
    }))
    .sort((left, right) => left.id.localeCompare(right.id));
  const tickets = Array.from(serverState.ticketsById.values())
    .map((ticketState) => ({
      id: normalizeString(ticketState.ticket.id),
      code: normalizeString(ticketState.ticket.code),
      status: normalizeString(ticketState.ticket.status) || undefined,
      groupIds: Array.from(ticketState.groupIds).sort(),
      emailKeys: Array.from(ticketState.emailKeys).sort(),
    }))
    .sort((left, right) => left.id.localeCompare(right.id));
  return { emails, groups, tickets };
}

function buildMockRelatedContext(serverState: MockGroupsServerState, payload: Partial<RelevantEmailPayload>) {
  const targetEmail = findMockServerEmail(serverState, payload);
  if (!targetEmail) {
    return { email: null, emails: [], groups: [], tickets: [] };
  }
  const relatedEmails = Array.from(serverState.emailsByKey.values())
    .filter((email) =>
      normalizeString(email.emailKey) === normalizeString(targetEmail.emailKey)
      || (
        normalizeString(email.conversationId)
        && normalizeString(email.conversationId) === normalizeString(targetEmail.conversationId)
      )
    )
    .sort((left, right) => normalizeString(left.emailKey).localeCompare(normalizeString(right.emailKey)));
  const relatedEmailKeys = new Set(relatedEmails.map((email) => normalizeString(email.emailKey)));
  const groups = Array.from(
    new Map(
      relatedEmails.flatMap((email) => (email.relatedGroups || []).map((group) => [
        normalizeString(group.id),
        ensureMockGroup(serverState, normalizeString(group.id), normalizeString(group.name) || undefined),
      ]))
    ).values()
  );
  const tickets = Array.from(serverState.ticketsById.values())
    .filter((ticketState) => Array.from(ticketState.emailKeys).some((emailKey) => relatedEmailKeys.has(emailKey)))
    .map((ticketState) => buildMockTicketResponse(ticketState, relatedEmailKeys.has(normalizeString(targetEmail.emailKey))))
    .sort((left, right) => normalizeString(left.id).localeCompare(normalizeString(right.id)));
  return {
    email: cloneJson(targetEmail),
    emails: cloneJson(relatedEmails),
    groups: cloneJson(groups),
    tickets: cloneJson(tickets),
  };
}

function simplifyIntermediateCase(caseValue: IntermediateCase | null) {
  if (!caseValue) return null;
  return {
    caseId: caseValue.caseId,
    anchorEmailKey: caseValue.anchorEmailKey,
    source: caseValue.sourceSummary.primarySource,
    retentionState: caseValue.retentionSummary.state,
    emails: caseValue.emails.map((email) => ({
      emailKey: email.emailKey,
      principalGroupId: email.classification.principalGroupId || null,
      referenceGroupIds: [...email.classification.referenceGroupIds],
      labels: [...email.classification.labels],
      ticketIds: [...email.classification.ticketIds],
      attachments: email.attachments.map((attachment) => ({
        attachmentKey: attachment.attachmentKey,
        name: attachment.name,
        documentState: attachment.documentState || null,
        isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : null,
        storageDecision: attachment.storageDecision,
      })),
    })),
  };
}

function simplifyRelatedContext(context: Awaited<ReturnType<typeof getRelatedEmailContext>> | null) {
  if (!context) return null;
  return {
    email: context.email
      ? {
          emailKey: context.email.emailKey,
          labels: context.email.labels || [],
          relatedGroups: (context.email.relatedGroups || [])
            .map((group) => ({
              id: group.id,
              relationKind: group.relationKind,
            }))
            .sort((left, right) => String(left.id).localeCompare(String(right.id))),
          attachments: (context.email.attachments || [])
            .map((attachment) => ({
              key: attachment.key,
              name: attachment.name,
              documentState: attachment.documentState || null,
              isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : null,
              storageProvider: attachment.storageProvider || null,
              storageBasePath: attachment.storageBasePath || null,
            }))
            .sort((left, right) => String(left.key || "").localeCompare(String(right.key || ""))),
        }
      : null,
    emails: context.emails
      .map((email) => ({
        emailKey: email.emailKey,
        labels: email.labels || [],
        relatedGroups: (email.relatedGroups || [])
          .map((group) => ({
            id: group.id,
            relationKind: group.relationKind,
          }))
          .sort((left, right) => String(left.id).localeCompare(String(right.id))),
        attachments: (email.attachments || [])
          .map((attachment) => ({
            key: attachment.key,
            name: attachment.name,
            documentState: attachment.documentState || null,
            isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : null,
            storageProvider: attachment.storageProvider || null,
            storageBasePath: attachment.storageBasePath || null,
          }))
          .sort((left, right) => String(left.key || "").localeCompare(String(right.key || ""))),
      }))
      .sort((left, right) => String(left.emailKey || "").localeCompare(String(right.emailKey || ""))),
    groups: context.groups
      .map((group) => ({ id: group.id, name: group.name }))
      .sort((left, right) => String(left.id).localeCompare(String(right.id))),
    tickets: context.tickets
      .map((ticket) => ({
        id: ticket.id,
        code: ticket.code,
        status: ticket.status || null,
        groupIds: (ticket.groupIds || []).slice().sort(),
        emailLinked: ticket.emailLinked === true,
      }))
      .sort((left, right) => String(left.id).localeCompare(String(right.id))),
  };
}

function parseCallBody(body: string | undefined): unknown {
  if (!body) return null;
  try {
    return JSON.parse(body);
  } catch {
    return body;
  }
}

function isCentralFinalWriteCall(url: string, method: string): boolean {
  if (method === "POST" && /\/api\/links\/email(?:\?|$)/i.test(url)) return true;
  if (method === "POST" && /\/api\/links\/groups\/[^/]+\/emails(?:\?|$)/i.test(url)) return true;
  if (method === "DELETE" && /\/api\/links\/groups\/[^/]+\/emails(?:\?|$)/i.test(url)) return true;
  if (method === "POST" && /\/api\/links\/group-tickets(?:\?|$)/i.test(url)) return true;
  if (method === "PATCH" && /\/api\/links\/group-tickets\/[^/]+(?:\?|$)/i.test(url)) return true;
  if (method === "POST" && /\/api\/links\/group-tickets\/[^/]+\/email(?:\?|$)/i.test(url)) return true;
  if (method === "DELETE" && /\/api\/links\/group-tickets\/[^/]+\/email(?:\?|$)/i.test(url)) return true;
  return false;
}

async function withMockedFetch<T>(
  run: (calls: MockFetchHarness) => Promise<T>,
  options?: {
    pickedFolderPath?: string;
  }
): Promise<T> {
  const originalFetch = window.fetch.bind(window);
  const rawCalls: MockFetchCall[] = [];
  const serverState = buildMockServerState();
  const manifestByKey = new Map<string, GroupWorksetManifest>();
  const intermediateTextByKey = new Map<string, string>();
  const intermediateBinaryByKey = new Map<string, { contentBase64: string; contentType?: string }>();
  const pickedFolderPath = String(options?.pickedFolderPath || "C:/Users/test/OneDrive - Demo/Groups").trim();
  const calls = Object.assign(rawCalls, {
    serverState,
    pickedFolderPath,
    getServerSnapshot: () => buildMockServerSnapshot(serverState),
  }) as MockFetchHarness;

  const makeIntermediateKey = (basePath: string, relativePath: string) => `${basePath}::${relativePath}`;
  const listIntermediatePaths = (basePath: string, prefix: string) =>
    Array.from(new Set([
      ...Array.from(intermediateTextByKey.keys()),
      ...Array.from(intermediateBinaryByKey.keys()),
    ]))
      .filter((entry) => entry.startsWith(`${basePath}::`))
      .map((entry) => entry.slice(`${basePath}::`.length))
      .filter((relativePath) => !prefix || relativePath === prefix || relativePath.startsWith(`${prefix}/`));

  const respondJson = (call: MockFetchCall, status: number, payload: unknown) => {
    call.responseStatus = status;
    call.responseBody = cloneJson(payload);
    return new Response(JSON.stringify(payload), {
      status,
      headers: { "Content-Type": "application/json" },
    });
  };

  window.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
    const method = String(init?.method || "GET").toUpperCase();
    const body = typeof init?.body === "string" ? init.body : undefined;
    const call: MockFetchCall = { url, method, body };
    calls.push(call);

    if (url.includes("/api/links/groups/worksets") && method === "POST") {
      const parsed = JSON.parse(body || "{}") as { manifest?: GroupWorksetManifest };
      if (parsed.manifest?.worksetKey) {
        manifestByKey.set(parsed.manifest.worksetKey, parsed.manifest);
      }
      return respondJson(call, 200, { ok: true, manifest: parsed.manifest || null });
    }

    if (url.includes("/api/links/groups/worksets/") && method === "GET") {
      const match = url.match(/\/api\/links\/groups\/worksets\/([^?]+)/i);
      const worksetKey = match ? decodeURIComponent(match[1]) : "";
      return respondJson(call, 200, { ok: true, manifest: manifestByKey.get(worksetKey) || null });
    }

    if (url.includes("/api/links/groups/storage/validate")) {
      const parsed = JSON.parse(body || "{}") as {
        mode?: string;
        baseFolderPath?: string;
        chosenFolder?: { path?: string; kind?: string };
      };
      const basePath = String(parsed?.chosenFolder?.path || parsed?.baseFolderPath || "").trim();
      const blocked = basePath.includes("INVALID");
      return respondJson(call, 200, {
        ok: true,
        result: {
          mode: parsed?.mode || (basePath ? "chosen_folder" : "supabase"),
          provider: basePath ? "local" : "cloud",
          fileBacked: Boolean(basePath),
          supported: !blocked,
          basePath,
          normalizedBasePath: basePath,
          isWebUrl: false,
          requiresServerAccessiblePath: Boolean(basePath),
          canStoreManifest: !blocked,
          canStoreBinary: !blocked,
          pickerAvailable: true,
          notes: ["mocked"],
          architecturalBlocker: null,
          requiredChange: null,
          blockingReason: blocked ? "Mocked invalid folder path." : null,
        },
      });
    }

    if (url.includes("/api/links/groups/intermediate-storage/")) {
      if (url.includes("/pick-folder")) {
        return respondJson(call, 200, {
          ok: true,
          result: {
            supported: true,
            selected: true,
            cancelled: false,
            path: pickedFolderPath,
            normalizedPath: pickedFolderPath,
            picker: "windows_folder_browser",
            reason: "mocked picker",
            validation: {
              supported: true,
              normalizedBasePath: pickedFolderPath,
              notes: ["mocked"],
            },
          },
        });
      }
      const parsed = JSON.parse(body || "{}") as {
        basePath?: string;
        path?: string;
        prefix?: string;
        content?: string;
        contentBase64?: string;
        contentType?: string;
      };
      const basePath = String(parsed.basePath || "").trim();
      const relativePath = String(parsed.path || "").trim();
      const prefix = String(parsed.prefix || "").trim();
      if (basePath.includes("INVALID")) {
        return respondJson(call, 500, {
          ok: false,
          error: "group_intermediate_mock_invalid_path",
          details: "Mocked invalid folder path.",
        });
      }
      if (url.includes("/read-text")) {
        return respondJson(call, 200, {
          ok: true,
          content: intermediateTextByKey.get(makeIntermediateKey(basePath, relativePath)) || null,
        });
      }
      if (url.includes("/write-text")) {
        intermediateTextByKey.set(makeIntermediateKey(basePath, relativePath), String(parsed.content || ""));
        return respondJson(call, 200, { ok: true });
      }
      if (url.includes("/read-binary")) {
        const payload = intermediateBinaryByKey.get(makeIntermediateKey(basePath, relativePath)) || null;
        return respondJson(call, 200, { ok: true, ...payload });
      }
      if (url.includes("/write-binary")) {
        intermediateBinaryByKey.set(makeIntermediateKey(basePath, relativePath), {
          contentBase64: String(parsed.contentBase64 || ""),
          contentType: String(parsed.contentType || "").trim() || undefined,
        });
        return respondJson(call, 200, { ok: true });
      }
      if (url.includes("/delete-tree")) {
        const normalizedPrefix = `${basePath}::${relativePath.replace(/\/+$/, "")}`;
        for (const key of Array.from(intermediateTextByKey.keys())) {
          if (key === normalizedPrefix || key.startsWith(`${normalizedPrefix}/`)) {
            intermediateTextByKey.delete(key);
          }
        }
        for (const key of Array.from(intermediateBinaryByKey.keys())) {
          if (key === normalizedPrefix || key.startsWith(`${normalizedPrefix}/`)) {
            intermediateBinaryByKey.delete(key);
          }
        }
        return respondJson(call, 200, { ok: true, deleted: true });
      }
      if (url.includes("/list-paths")) {
        return respondJson(call, 200, { ok: true, paths: listIntermediatePaths(basePath, prefix) });
      }
    }

    if (/\/api\/links\/email(?:\?|$)/i.test(url) && method === "POST") {
      const parsed = JSON.parse(body || "{}") as RelevantEmailPayload;
      const email = upsertMockServerEmail(serverState, parsed);
      return respondJson(call, 200, {
        email,
        groups: (email.relatedGroups || []).map((group) => ensureMockGroup(serverState, normalizeString(group.id), normalizeString(group.name) || undefined)),
      });
    }

    if (/\/api\/links\/related\?/i.test(url) && method === "GET") {
      const parsedUrl = new URL(url, window.location.origin);
      const payload: Partial<RelevantEmailPayload> = {
        conversationId: parsedUrl.searchParams.get("conversationId") || undefined,
        internetMessageId: parsedUrl.searchParams.get("internetMessageId") || undefined,
        itemId: parsedUrl.searchParams.get("itemId") || undefined,
        subject: parsedUrl.searchParams.get("subject") || undefined,
        fromEmail: parsedUrl.searchParams.get("fromEmail") || undefined,
        receivedAtIso: parsedUrl.searchParams.get("receivedAtIso") || undefined,
      };
      return respondJson(call, 200, buildMockRelatedContext(serverState, payload));
    }

    if (/\/api\/links\/groups\/[^/]+\/emails(?:\?|$)/i.test(url)) {
      const match = url.match(/\/api\/links\/groups\/([^/]+)\/emails/i);
      const groupId = match ? decodeURIComponent(match[1]) : "";
      if (method === "POST") {
        const parsed = JSON.parse(body || "{}") as RelevantEmailPayload;
        const email = upsertMockServerEmail(serverState, parsed);
        const updatedEmail = applyMockGroupMembership({
          serverState,
          email,
          groupId,
          relationKind: normalizeString(parsed.membershipKind) || "referencia",
        });
        return respondJson(call, 200, {
          group: ensureMockGroup(serverState, groupId),
          email: updatedEmail,
        });
      }
      if (method === "DELETE") {
        const parsed = JSON.parse(body || "{}") as RelevantEmailPayload & { emailKey?: string };
        const email = findMockServerEmail(serverState, {
          ...parsed,
          internetMessageId: parsed.internetMessageId,
          itemId: parsed.itemId,
        }) || (parsed.emailKey ? serverState.emailsByKey.get(parsed.emailKey) || null : null);
        if (email) {
          removeMockGroupMembership({ serverState, email, groupId });
        }
        return respondJson(call, 200, { ok: true });
      }
    }

    if (/\/api\/links\/group-tickets(?:\?|$)/i.test(url) && method === "POST") {
      const parsed = JSON.parse(body || "{}") as {
        seriesId?: string;
        title?: string;
        description?: string;
        labels?: string[];
        groupIds?: string[];
      };
      const state = ensureMockTicket(serverState, `ticket-${serverState.nextTicketSequence}`, {
        seriesId: normalizeString(parsed.seriesId) || "series-default",
        title: normalizeString(parsed.title) || "Ticket",
        description: normalizeString(parsed.description) || undefined,
        labels: normalizeStringList(parsed.labels || []),
        groupIds: normalizeStringList(parsed.groupIds || []),
      });
      return respondJson(call, 200, { ticket: buildMockTicketResponse(state, false) });
    }

    if (/\/api\/links\/group-tickets\/[^/]+(?:\?|$)/i.test(url) && method === "PATCH") {
      const match = url.match(/\/api\/links\/group-tickets\/([^/]+)(?:\?|$)/i);
      const ticketId = match ? decodeURIComponent(match[1]) : "";
      const parsed = JSON.parse(body || "{}") as Partial<GroupTicketEntry>;
      const state = ensureMockTicket(serverState, ticketId);
      state.ticket = {
        ...state.ticket,
        ...parsed,
        groupIds: parsed.groupIds ? normalizeStringList(parsed.groupIds) : Array.from(state.groupIds),
      };
      state.groupIds = new Set(normalizeStringList(state.ticket.groupIds || []));
      return respondJson(call, 200, { ticket: buildMockTicketResponse(state, false) });
    }

    if (/\/api\/links\/group-tickets\/[^/]+\/email(?:\?|$)/i.test(url)) {
      const match = url.match(/\/api\/links\/group-tickets\/([^/]+)\/email/i);
      const ticketId = match ? decodeURIComponent(match[1]) : "";
      const ticketState = ensureMockTicket(serverState, ticketId);
      if (method === "POST") {
        const parsed = JSON.parse(body || "{}") as {
          email?: RelevantEmailPayload;
          applyGroups?: boolean;
          groupIds?: string[];
          membershipKind?: string;
        };
        const email = upsertMockServerEmail(serverState, parsed.email || {});
        ticketState.emailKeys.add(normalizeString(email.emailKey));
        const appliedGroups: LinkGroupEntry[] = [];
        if (parsed.applyGroups) {
          for (const groupId of normalizeStringList(parsed.groupIds || [])) {
            const updatedEmail = applyMockGroupMembership({
              serverState,
              email,
              groupId,
              relationKind: normalizeString(parsed.membershipKind) || "referencia",
            });
            appliedGroups.push(
              ensureMockGroup(serverState, groupId, (updatedEmail.relatedGroups || []).find((group) => group.id === groupId)?.name)
            );
          }
        }
        for (const groupId of normalizeStringList(parsed.groupIds || [])) {
          ticketState.groupIds.add(groupId);
        }
        return respondJson(call, 200, {
          ok: true,
          ticket: buildMockTicketResponse(ticketState, true),
          appliedGroups,
          email: serverState.emailsByKey.get(normalizeString(email.emailKey)) || email,
        });
      }
      if (method === "DELETE") {
        const parsed = JSON.parse(body || "{}") as { email?: RelevantEmailPayload; emailKey?: string };
        const email = parsed.email
          ? findMockServerEmail(serverState, parsed.email)
          : (parsed.emailKey ? serverState.emailsByKey.get(normalizeString(parsed.emailKey)) || null : null);
        if (email?.emailKey) {
          ticketState.emailKeys.delete(normalizeString(email.emailKey));
        }
        return respondJson(call, 200, {
          ok: true,
          removed: true,
          ticket: buildMockTicketResponse(ticketState, false),
          emailKey: normalizeString(email?.emailKey),
        });
      }
    }

    return originalFetch(input, init);
  }) as typeof window.fetch;

  try {
    return await run(calls);
  } finally {
    window.fetch = originalFetch;
  }
}

type FinalPersistenceScenarioArgs = {
  mock: MockFetchHarness;
  scenarioId: string;
  title: string;
  classificationCase: IntermediateCase;
  targetEmails: RelatedEmailEntry[];
  selectedEmail: RelatedEmailEntry;
  currentContext: {
    itemId?: string;
    internetMessageId?: string;
    conversationId?: string;
    subject?: string;
    fromEmail?: string;
    fromName?: string;
    receivedAtIso?: string;
    toRecipients?: RelevantEmailPayload["toRecipients"];
    ccRecipients?: RelevantEmailPayload["ccRecipients"];
  };
  resolvedApplySelection: ReturnType<typeof buildResolvedStudioApplySelection>;
  classificationMetaDraft: ClassificationMetaDraft;
  emailContextMeta?: Map<string, { groupIds: string[]; labels: string[]; ticketIds: string[] }>;
  relevantSettings: Record<string, unknown>;
  expected: Record<string, unknown>;
  writeProofs: FinalWriteProofArtifact[];
  attachmentStorageOptions?: ReturnType<typeof buildAttachmentStorageOptions>;
  createTicketTitle?: string;
  currentOutlookTicket?: GroupTicketEntry | null;
};

async function executeFinalPersistenceProof(args: FinalPersistenceScenarioArgs): Promise<string> {
  const {
    mock,
    scenarioId,
    title,
    classificationCase,
    targetEmails,
    selectedEmail,
    currentContext,
    resolvedApplySelection,
    classificationMetaDraft,
    emailContextMeta,
    relevantSettings,
    expected,
    writeProofs,
    attachmentStorageOptions,
    createTicketTitle,
    currentOutlookTicket,
  } = args;

  await saveSettings({
    groups: buildGroupsSettings({
      storage: {
        ...DEFAULT_GROUP_STORAGE_SETTINGS,
        mode: "supabase",
      },
      tab: {
        ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
        storageMode: "local_indexeddb",
        baseFolderPath: "",
      },
    }),
  });

  const intermediateStorage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
  assert(intermediateStorage.mode === "indexeddb", "O modo add-in fallback devia usar IndexedDB como storage intermédio principal.");
  await intermediateStorage.repository.writeCase(classificationCase);

  const remoteApplyPlan = buildResolvedRemoteApplyExecutionPlan({
    targetEmails,
    resolvedApplySelection,
    currentContext,
    emailContextMeta: emailContextMeta || new Map(),
  });

  const baseTicketApply = await executeLegacyBaseTicketApply({
    remoteApplyPlan,
    resolvedApplySelection,
    currentContext,
    createTicketTitle,
    currentOutlookTicket: currentOutlookTicket || null,
    attachmentStorageOptions,
  });

  let finalTicket = baseTicketApply.finalTicket;
  for (const targetPlan of remoteApplyPlan.targetPlans) {
    finalTicket = await executeLegacyRemoteApplyForTarget({
      targetPlan,
      resolvedApplySelection,
      finalTicket,
      attachmentStorageOptions,
      skipTicketLink: false,
    });
  }

  const projected = projectApplyIntoIntermediateCase({
    classificationCase,
    resolvedApplySelection,
    resolvedCaseTicket: finalTicket,
    targetEmails,
    classificationMetaDraft,
  });

  const selectedEmailPayload = buildRelevantEmailPayloadFromRelatedEmail(selectedEmail);
  assert(selectedEmailPayload, "Nao foi possivel construir payload de reabertura para o email selecionado.");

  const persisted = await persistAndRefreshClassificationCase({
    classificationCase,
    nextClassificationCase: projected.nextClassificationCase,
    preferredSelectedEmailKey: makeEmailKey(selectedEmail),
    preferredTargetEmailKeys: targetEmails.map((email) => makeEmailKey(email)),
    syncClassificationCaseEmails: () => undefined,
    refreshSelectedEmailContext: async () => getRelatedEmailContext(selectedEmailPayload),
  });

  const reopenedLocal = await resolveClassificationIntermediateCase({
    caseId: classificationCase.caseId,
    anchorEmailKey: classificationCase.anchorEmailKey,
  });
  const reopenedServer = await getRelatedEmailContext(selectedEmailPayload);
  const finalCalls = mock
    .filter((call) => isCentralFinalWriteCall(call.url, call.method))
    .map((call) => ({
      url: call.url,
      method: call.method,
      body: parseCallBody(call.body),
      status: Number(call.responseStatus || 0),
      responseBody: cloneJson(call.responseBody ?? null),
    }));

  const actual = {
    intermediateMode: intermediateStorage.mode,
    localCase: simplifyIntermediateCase(reopenedLocal.caseValue),
    serverContext: simplifyRelatedContext(reopenedServer),
    persistedGroups: mock.getServerSnapshot().groups,
    persistedTickets: mock.getServerSnapshot().tickets,
  };
  const pass = JSON.stringify(expected) === JSON.stringify(actual.serverContext ? {
    serverContext: actual.serverContext,
    persistedGroups: actual.persistedGroups,
    persistedTickets: actual.persistedTickets,
  } : actual);

  writeProofs.push({
    scenarioId,
    title,
    inputContext: {
      anchorEmailKey: classificationCase.anchorEmailKey,
      caseId: classificationCase.caseId,
      targetEmailKeys: targetEmails.map((email) => makeEmailKey(email)),
      selectedEmailKey: makeEmailKey(selectedEmail),
    },
    relevantSettings,
    intermediateState: {
      beforeWrite: simplifyIntermediateCase(classificationCase),
      afterProjection: simplifyIntermediateCase(projected.nextClassificationCase),
      reopenedLocalCase: simplifyIntermediateCase(reopenedLocal.caseValue),
    },
    finalPayloads: finalCalls.map((call) => ({
      url: call.url,
      method: call.method,
      body: call.body,
    })),
    backendResponses: finalCalls.map((call) => ({
      url: call.url,
      method: call.method,
      status: call.status,
      body: call.responseBody,
    })),
    persistedBackendState: mock.getServerSnapshot(),
    reopenedState: {
      serverContext: simplifyRelatedContext(reopenedServer),
      refreshedContext: simplifyRelatedContext(persisted.refreshedContext),
    },
    expected,
    actual,
    pass,
    failureReason: pass ? undefined : "Expected vs actual diverged on persisted server state or reopened context.",
  });

  assert(pass, "A reabertura nao refletiu o estado final esperado do backend central.");
  return `Persistencia final provada com ${finalCalls.length} writes centrais e reabertura coerente a partir do backend.`;
}

async function runScenario(
  scenarios: ValidationScenarioResult[],
  id: string,
  area: ValidationArea,
  title: string,
  fn: () => Promise<string> | string
): Promise<void> {
  try {
    const details = await fn();
    scenarios.push({ id, area, title, status: "passed", details });
  } catch (error) {
    scenarios.push({
      id,
      area,
      title,
      status: "failed",
      details: error instanceof Error ? error.message : String(error || "Erro desconhecido"),
    });
  }
}

export async function runGroupsV1BrowserValidation(): Promise<GroupsBrowserValidationReport> {
  await resetBrowserPersistence();
  await resetSettings();

  const scenarios: ValidationScenarioResult[] = [];
  const writeProofs: FinalWriteProofArtifact[] = [];

  await runScenario(scenarios, "settings-groups-roundtrip", "settings", "settings.groups roundtrip sem aliases runtime", async () => {
    const groups = buildGroupsSettings({
      storage: {
        ...DEFAULT_GROUP_STORAGE_SETTINGS,
        mode: "hybrid",
        chosenFolder: { path: "C:/tmp/groups/chosen", kind: "filesystem" },
        hybrid: { primaryTarget: "chosen_folder", promoteManifestOnSave: true, promoteAttachmentMetadataOnSave: true },
      },
        tab: {
          ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
          groupsTabEnabled: false,
          storageMode: "local_indexeddb",
          baseFolderPath: "C:/dados/grupos/runtime",
        cleanupClosedCaseDays: 12,
        cleanupAbandonedCaseDays: 34,
        neverDeleteMixedSilently: true,
      },
      labels: {
        managerEnabled: true,
        catalog: [{ label: "Financeiro", categorize: true, hasStatus: true, status: "em_analise" }],
        favoriteIds: ["Financeiro"],
      },
      tickets: {
        enabled: false,
        ui: {
          ...DEFAULT_GROUPS_MODULE_SETTINGS.tickets.ui,
          autoLinkMode: "auto",
          aiInstructions: "Responde com o codigo do ticket.",
        },
      },
      outlookCategories: {
        enabled: true,
        includeGroups: true,
        includeTickets: false,
        includeStatuses: true,
        includeLabels: true,
      },
    });
    await saveSettings({ groups });
    const next = await getSettings();
    assert(next.groups.storage.mode === "hybrid", "O modo de storage canonico nao foi persistido.");
    assert(next.groups.tab.baseFolderPath === "C:/dados/grupos/runtime", "A pasta intermédia canónica nao foi persistida.");
    assert(next.groups.labels.catalog[0]?.label === "Financeiro", "O catalogo de labels canonico nao foi persistido.");
    assert(next.groups.tickets.enabled === false, "O flag canonico de tickets nao foi persistido.");
    assert(next.groups.outlookCategories.includeLabels === true, "O flag canonico de Outlook categories nao foi persistido.");
    const runtimeSnapshot = next as Record<string, unknown>;
    assert(!Object.prototype.hasOwnProperty.call(runtimeSnapshot, "groupStorage"), "O runtime ainda expôs groupStorage.");
    assert(!Object.prototype.hasOwnProperty.call(runtimeSnapshot, "groupsTabSettings"), "O runtime ainda expôs groupsTabSettings.");
    assert(!Object.prototype.hasOwnProperty.call(runtimeSnapshot, "groupLabelCatalog"), "O runtime ainda expôs groupLabelCatalog.");
    return "Roundtrip canonico confirmado; o runtime devolveu apenas settings.groups sem aliases legacy ativos.";
  });

  await runScenario(scenarios, "settings-panel-host-safe-shells", "settings", "Painel de settings host-safe e shells honestamente desativadas", async () => {
    await renderSettingsPanelScenario();
    return "GroupsSettingsPanel renderizou editores inline e os toggles canonicos ficaram interativos sem prompt/alert/confirm.";
  });

  await runScenario(scenarios, "settings-tab-folder-picker", "settings", "Picker real de pasta preenche e valida a pasta intermédia", async () => {
    await withMockedFetch(async () => {
      const host = document.createElement("div");
      document.body.appendChild(host);
      const root = ReactDOM.createRoot(host);

      try {
        root.render(
          <GroupsSettingsPanel
            open
            value={normalizeGroupsTabSettings(DEFAULT_GROUPS_TAB_SETTINGS)}
            onClose={() => undefined}
            onSave={() => undefined}
            initialSection="intermediate_storage"
          />
        );
        await new Promise((resolve) => setTimeout(resolve, 20));

        const browseButton = Array.from(host.querySelectorAll("button")).find((button) =>
          button.textContent?.includes("Procurar pasta")
        );
        assert(browseButton, "O botão de picker real nao foi renderizado.");
        browseButton?.click();
        await new Promise((resolve) => setTimeout(resolve, 40));

        const pathInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: C:/dados/grupos/intermedio"]');
        assert(pathInput?.value === "C:/Users/test/OneDrive - Demo/Groups", "O picker nao preencheu a pasta intermédia escolhida.");
        const calls = Array.from(host.querySelectorAll("*"));
        assert(calls.length > 0, "O painel deixou de renderizar depois do picker.");
      } finally {
        root.unmount();
        host.remove();
      }
    });

    return "O botão 'Procurar pasta' preencheu um caminho local real vindo do picker, incluindo um exemplo de pasta OneDrive sincronizada localmente.";
  });

  await runScenario(scenarios, "settings-tab-case-behavior", "settings", "Settings de caso mudam bootstrap e reabertura", async () => {
    await withMockedFetch(async () => {
      const folderPath = "C:/dados/grupos/settings-case";
      const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: folderPath }));
      const caseValue = buildPrepareIntermediateCaseFromSources({
        caseId: "case-settings-runtime",
        anchorEmailKey: "msg-settings-runtime",
        outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-settings-runtime" })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      await storage.repository.writeCase(caseValue);

      await saveSettings({
        groups: buildGroupsSettings({
          tab: {
            ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
            storageMode: "local_indexeddb",
            baseFolderPath: folderPath,
            reopenExistingCase: true,
          },
        }),
      });
      const reopened = await resolveClassificationIntermediateCase({ anchorEmailKey: "msg-settings-runtime" });
      assert(reopened.caseValue?.caseId === caseValue.caseId, "Com reopenExistingCase=true o caso devia reabrir por anchor.");

      await saveSettings({
        groups: buildGroupsSettings({
          tab: {
            ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
            storageMode: "local_indexeddb",
            baseFolderPath: folderPath,
            reopenExistingCase: false,
            autoCreateCaseOnNewEmail: false,
            recreateIntermediateCopy: false,
          },
        }),
      });
      const blockedReopen = await resolveClassificationIntermediateCase({ anchorEmailKey: "msg-settings-runtime" });
      assert(blockedReopen.caseValue === null, "Com reopenExistingCase=false o runtime nao devia reabrir por anchor.");
    });
    assert(
      shouldPersistGroupsPrepareCase({
        settingsLike: { groups: { tab: { autoCreateCaseOnNewEmail: false } } },
        hasHydratedCase: false,
        hasLocalCheckpoint: false,
      }) === false,
      "autoCreateCaseOnNewEmail=false devia impedir checkpoint automatico novo."
    );
    assert(
      shouldProjectServerCopyIntoIntermediate({
        settingsLike: { groups: { tab: { recreateIntermediateCopy: false } } },
        hasHydratedCase: false,
      }) === false,
      "recreateIntermediateCopy=false devia impedir a projecao remota quando nao existe caso local."
    );
    return "autoCreate/reopen/recreate passaram a governar o bootstrap e a reabertura, incluindo o caso file-backed.";
  });

  await runScenario(scenarios, "settings-tab-storage-validation", "settings", "Settings de validacao mandam no acesso ao intermédio", async () => {
    const addinLocalSettings = normalizeGroupsTabSettings({
      storageMode: "local_indexeddb",
      baseFolderPath: "",
      validateLocationOnOpen: true,
      blockTabIfUnavailable: true,
      warnIfUnavailable: true,
      autoRetryValidation: true,
    });
    const validation = await validateGroupsTabStorageAvailability({
      settings: addinLocalSettings,
      storage: resolveIntermediateCaseStorage(addinLocalSettings),
    });
    assert(validation.available === true, "Sem pasta definida o fallback principal devia ser o storage local do add-in.");

    await withIndexedDbAvailabilityOverride(undefined, async () => {
      const fallbackSettings = normalizeGroupsTabSettings({
        storageMode: "local_indexeddb",
        baseFolderPath: "",
        validateLocationOnOpen: true,
        blockTabIfUnavailable: true,
        warnIfUnavailable: true,
        autoRetryValidation: true,
      });
      const fallbackValidation = await validateGroupsTabStorageAvailability({
        settings: fallbackSettings,
        storage: resolveIntermediateCaseStorage(fallbackSettings),
      });
      assert(fallbackValidation.available === false, "Sem pasta e sem IndexedDB a validação devia acusar fallback técnico.");
      assert(fallbackValidation.blocked === true, "blockTabIfUnavailable=true devia bloquear a aba.");
      assert(fallbackValidation.warning === true, "warnIfUnavailable=true devia marcar warning.");
    });

    const relaxedValidation = await validateGroupsTabStorageAvailability({
      settings: normalizeGroupsTabSettings({
        ...addinLocalSettings,
        validateLocationOnOpen: false,
        blockTabIfUnavailable: false,
        warnIfUnavailable: false,
      }),
      storage: resolveIntermediateCaseStorage(addinLocalSettings),
    });
    assert(relaxedValidation.available === true, "Com validateLocationOnOpen=false a validacao devia ser neutralizada.");

    await withMockedFetch(async () => {
      const invalidFolderSettings = normalizeGroupsTabSettings({
        storageMode: "local_indexeddb",
        baseFolderPath: "C:/INVALID/groups",
        validateLocationOnOpen: true,
        blockTabIfUnavailable: true,
        warnIfUnavailable: true,
        autoRetryValidation: true,
      });
      const invalidFolderValidation = await validateGroupsTabStorageAvailability({
        settings: invalidFolderSettings,
        storage: resolveIntermediateCaseStorage(invalidFolderSettings),
      });
      assert(invalidFolderValidation.available === false, "Pasta inválida devia falhar na validacao real.");
      assert(invalidFolderValidation.blocked === true, "Pasta inválida devia obedecer a blockTabIfUnavailable.");
      assert(invalidFolderValidation.warning === true, "Pasta inválida devia obedecer a warnIfUnavailable.");
      assert(invalidFolderValidation.retrySuggested === true, "Pasta inválida devia obedecer a autoRetryValidation.");
    });

    return "validateLocationOnOpen, blockTabIfUnavailable, warnIfUnavailable e autoRetryValidation passaram a distinguir add-in local pronto, pasta inválida e fallback técnico em memória.";
  });

  await runScenario(scenarios, "settings-tab-attachment-runtime", "attachments", "Settings de anexos mudam selecao, destino e payload", async () => {
    const settingsLike = {
      groups: {
        tab: normalizeGroupsTabSettings({
          attachmentStrategy: "by_size",
          saveAttachmentsOnServer: true,
          saveAttachmentsOutsideServer: true,
          attachmentServerLimitMb: 1,
          attachmentIntermediateLimitMb: 2,
          externalAttachmentFolder: "C:/tmp/groups/outside",
          showAttachmentMetadataOnServer: false,
          requireImmediatePreview: false,
        }),
        storage: {
          ...DEFAULT_GROUP_STORAGE_SETTINGS,
          mode: "chosen_folder",
          chosenFolder: { path: "C:/tmp/groups/chosen", kind: "filesystem" },
        },
      },
    };
    const small = resolveGroupsTabAttachmentDecision({
      key: "doc-small",
      name: "small.pdf",
      size: 200 * 1024,
      hasContent: true,
    }, settingsLike);
    const large = resolveGroupsTabAttachmentDecision({
      key: "doc-large",
      name: "large.pdf",
      size: 4 * 1024 * 1024,
      hasContent: true,
    }, settingsLike);
    assert(small.target === "server", "Anexo pequeno devia ir para server no modo by_size.");
    assert(large.target === "outside", "Anexo grande devia ir para outside no modo by_size.");
    assert(large.selectedByDefault === false, "attachmentIntermediateLimitMb devia tirar anexos grandes da selecao por defeito quando nao ha preview imediato.");
    const forcedPreview = resolveGroupsTabAttachmentDecision({
      key: "doc-large",
      name: "large.pdf",
      size: 4 * 1024 * 1024,
      hasContent: true,
    }, {
      groups: {
        ...settingsLike.groups,
        tab: normalizeGroupsTabSettings({
          ...settingsLike.groups.tab,
          requireImmediatePreview: true,
        }),
      },
    });
    assert(forcedPreview.selectedByDefault === true, "requireImmediatePreview=true devia voltar a selecionar o anexo grande.");
    const payload = buildRelevantEmailPayloadFromRelatedEmail(buildRelatedEmail({
      emailKey: "msg-attachments-runtime",
      attachments: [{
        key: "doc-large",
        id: "doc-large",
        name: "large.pdf",
        contentType: "application/pdf",
        size: 4 * 1024 * 1024,
        hasContent: true,
        content: "ZmFrZQ==",
      } as RelatedEmailAttachmentFixture],
    }), settingsLike);
    assert(payload?.attachments?.[0]?.storageBasePath === "C:/tmp/groups/outside", "externalAttachmentFolder devia seguir para o payload final.");
    const options = buildGroupsTabAttachmentStorageOptions(settingsLike);
    assert(options.attachmentStorageBasePath === "C:/tmp/groups/outside", "A bridge final de anexos devia respeitar o destino externo.");
    return "attachmentStrategy/saveAttachments*/limits/externalAttachmentFolder/showAttachmentMetadataOnServer/requireImmediatePreview passaram a influenciar o runtime.";
  });

  await runScenario(scenarios, "settings-tab-warning-runtime", "cleanup", "Warnings e cadencia de limpeza obedecem aos settings", async () => {
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-warning-runtime",
      anchorEmailKey: "msg-warning-runtime",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-warning-runtime" })],
      serverEmails: [],
      nowIso: "2026-04-01T12:00:00.000Z",
    });
    caseValue.classificationSummary = {
      ...caseValue.classificationSummary,
      mixedCase: true,
      unclassifiedEmails: 2,
    };
    caseValue.lastAccessedAt = "2026-03-01T12:00:00.000Z";
    const messages = buildGroupsTabWarningMessages({
      settings: normalizeGroupsTabSettings({
        mixedCaseWarningDays: 15,
        localAbandonedWarningDays: 20,
        warnUnclassifiedEmails: true,
        warnMixedCases: true,
      }),
      caseValue,
      summary: {
        caseId: caseValue.caseId,
        anchorEmailKey: caseValue.anchorEmailKey,
        subject: "warning",
        updatedAt: caseValue.updatedAt,
        lastAccessedAt: caseValue.lastAccessedAt,
        totalEmails: caseValue.emails.length,
        classifiedEmails: 0,
        unclassifiedEmails: 2,
        visibleState: "draft",
        retentionState: "local_only",
        quickState: "draft",
      },
      nowMs: FIXED_NOW_MS,
    });
    assert(messages.some((entry) => entry.kind === "mixed_case"), "warnMixedCases/mixedCaseWarningDays nao geraram aviso.");
    assert(messages.some((entry) => entry.kind === "unclassified"), "warnUnclassifiedEmails nao gerou aviso.");
    assert(messages.some((entry) => entry.kind === "local_abandoned"), "localAbandonedWarningDays nao gerou aviso.");
    assert(isGroupsTabFrequencyDue("daily", "2026-04-20T12:00:00.000Z", FIXED_NOW_MS) === true, "warningFrequency/cleanupFrequency deviam disparar quando a janela expirou.");
    assert(isGroupsTabFrequencyDue("weekly", FIXED_NOW_ISO, FIXED_NOW_MS) === false, "warningFrequency/cleanupFrequency nao deviam disparar na mesma semana.");
    return "mixedCaseWarningDays/localAbandonedWarningDays/cleanupFrequency/warnUnclassifiedEmails/warnMixedCases/warningFrequency passaram a ter efeito real.";
  });

  await runScenario(scenarios, "settings-tab-prepare-bridge", "settings", "prepareTasksBridge governa a ponte local Prepare -> Classificar", async () => {
    assert(shouldUsePrepareTasksBridge({ groups: { tab: { prepareTasksBridge: true } } }) === true, "prepareTasksBridge=true devia ligar a ponte local.");
    assert(shouldUsePrepareTasksBridge({ groups: { tab: { prepareTasksBridge: false } } }) === false, "prepareTasksBridge=false devia desligar a ponte local.");
    return "prepareTasksBridge passou a governar a escrita do contexto local entre Preparar e Classificar.";
  });

  await runScenario(scenarios, "settings-tab-explorer-bridges", "settings", "explorer* governa as bridges internas do studio", async () => {
    assert(shouldUseExplorerServerPrimary({ groups: { tab: { explorerServerPrimary: true } } }) === true, "explorerServerPrimary=true devia preferir o lado server.");
    assert(shouldUseExplorerServerPrimary({ groups: { tab: { explorerServerPrimary: false } } }) === false, "explorerServerPrimary=false devia permitir preferencia local/intermédia.");
    assert(canOpenStoredAttachmentsFromGroups({ groups: { tab: { explorerOpenStoredAttachments: true } } }) === true, "explorerOpenStoredAttachments=true devia permitir hidratar anexos guardados.");
    assert(canOpenStoredAttachmentsFromGroups({ groups: { tab: { explorerOpenStoredAttachments: false } } }) === false, "explorerOpenStoredAttachments=false devia bloquear hidratação remota.");
    assert(canGenerateReplyFromGroups({ groups: { tab: { explorerGenerateReply: true } } }) === true, "explorerGenerateReply=true devia permitir reply/forward.");
    assert(canGenerateReplyFromGroups({ groups: { tab: { explorerGenerateReply: false } } }) === false, "explorerGenerateReply=false devia bloquear reply/forward.");
    return "explorerServerPrimary/explorerOpenStoredAttachments/explorerGenerateReply passaram a governar o bootstrap e as acoes do studio.";
  });

  await runScenario(scenarios, "prepare-new-email-no-group", "prepare", "Preparar email novo sem grupo", async () => {
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-prepare-1",
      anchorEmailKey: "msg-anchor",
      conversationId: "conv-prepare",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor", subject: "Novo email" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    assert(caseValue.emails.length === 1, "Um email novo sem relacoes nao devia abrir mais emails no caso.");
    assert(caseValue.sourceSummary.primarySource === "outlook", "A origem primaria devia ser Outlook.");
    return "Caso preparado com um unico email ancora e sem grupo previo.";
  });

  await runScenario(scenarios, "prepare-related-history-priority", "prepare", "Preparar email com historico relacionado do servidor", async () => {
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-prepare-2",
      anchorEmailKey: "msg-anchor",
      conversationId: "conv-prepare",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor", subject: "Historico" })],
      serverEmails: [
        buildPrepareEmailInput({ emailKey: "msg-server-1", sourceOrigin: "server", subject: "Historico 1" }),
        buildPrepareEmailInput({ emailKey: "msg-server-2", sourceOrigin: "server", subject: "Historico 2" }),
      ],
      nowIso: FIXED_NOW_ISO,
    });
    assert(caseValue.sourceSummary.primarySource === "server", "Com historico remoto a origem primaria devia passar para server.");
    assert(caseValue.emails.length === 3, "O caso devia combinar ancora + historico remoto.");
    return "Historico relacionado do servidor entrou no caso e passou a ser a origem primaria.";
  });

  await runScenario(scenarios, "prepare-same-subject-no-implicit-merge", "prepare", "Mesmo assunto sem relacao forte nao entra implicitamente", async () => {
    const anchor = buildPrepareEmailInput({ emailKey: "msg-anchor", subject: "Mesmo assunto" });
    const unrelatedSameSubject = buildPrepareEmailInput({ emailKey: "msg-other", subject: "Mesmo assunto" });
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-prepare-3",
      anchorEmailKey: anchor.emailKey,
      conversationId: "conv-prepare",
      outlookEmails: [anchor],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    assert(caseValue.emails.every((email) => email.emailKey !== unrelatedSameSubject.emailKey), "Um email so com assunto igual nao devia entrar sem ser passado como relacionado.");
    return "O miolo de Preparar nao faz merge por assunto sozinho.";
  });

  await runScenario(scenarios, "prepare-existing-group-preserved", "prepare", "Preparar email com grupo ja existente", async () => {
    const existingCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-prepare-4",
      anchorEmailKey: "msg-anchor",
      conversationId: "conv-prepare",
      outlookEmails: [
        buildPrepareEmailInput({
          emailKey: "msg-anchor",
          principalGroupId: "grp-main",
          referenceGroupIds: ["grp-ref"],
          labels: ["Financeiro"],
        }),
      ],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    const nextCase = buildPrepareIntermediateCaseFromSources({
      caseId: existingCase.caseId,
      anchorEmailKey: existingCase.anchorEmailKey,
      conversationId: existingCase.conversationId,
      existingCase,
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor", subject: "Reabertura" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    assert(nextCase.emails[0]?.classification.principalGroupId === "grp-main", "A classificacao existente devia ser preservada na reabertura.");
    return "A reabertura preservou grupo principal, referencias e labels ja guardados.";
  });

  await runScenario(scenarios, "prepare-multiple-related-emails", "prepare", "Preparar com multiplos emails relacionados", async () => {
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-prepare-5",
      anchorEmailKey: "msg-anchor",
      conversationId: "conv-prepare",
      existingCase: buildPrepareIntermediateCaseFromSources({
        caseId: "case-prepare-5",
        anchorEmailKey: "msg-anchor",
        conversationId: "conv-prepare",
        outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor" })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      }),
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor" })],
      serverEmails: [
        buildPrepareEmailInput({ emailKey: "msg-rel-1", sourceOrigin: "server" }),
        buildPrepareEmailInput({ emailKey: "msg-rel-2", sourceOrigin: "server" }),
      ],
      nowIso: FIXED_NOW_ISO,
    });
    assert(caseValue.emails.length === 3, "A combinacao de email ancora + relacionados devia produzir tres emails unicos.");
    return "Preparar combinou ancora, existingCase e historico remoto sem duplicar emails.";
  });

  await runScenario(scenarios, "prepare-storage-folder-primary", "storage", "Com pasta definida o intermédio grava logo na pasta local", async () => {
    await withMockedFetch(async () => {
      const settings = normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "C:/dados/grupos/intermedio" });
      const storage = resolveIntermediateCaseStorage(settings);
      assert(storage.mode === "filesystem", "Com pasta definida o storage intermédio devia ser file-backed.");
      const caseValue = buildPrepareIntermediateCaseFromSources({
        caseId: "case-storage-folder",
        anchorEmailKey: "msg-anchor",
        outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor", attachments: [{ key: "doc-1", name: "doc.pdf", hasContent: true }] })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      await storage.repository.writeCase(caseValue);
      const attachmentPath = caseValue.emails[0]?.attachments[0]?.localRef?.value;
      if (attachmentPath && "writeBinary" in storage.adapter) {
        await storage.adapter.writeBinary(attachmentPath, new Blob(["folder-binary"]));
      }
      const readBack = await storage.repository.readCase(caseValue.caseId);
      assert(readBack?.caseId === caseValue.caseId, "O caso intermédio nao voltou da pasta local.");
      if (attachmentPath && "readBinary" in storage.adapter) {
        assert(await storage.adapter.readBinary(attachmentPath), "O binário intermédio nao voltou da pasta local.");
      }
    });
    return "Pasta local definida confirmada como storage intermédio principal desde o arranque.";
  });

  await runScenario(scenarios, "prepare-storage-addin-local-fallback", "storage", "Sem pasta definida o intermédio usa o add-in local", async () => {
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
    assert(storage.mode === "indexeddb", "Sem pasta definida o fallback principal devia ser IndexedDB do add-in.");
    assert(storage.availability === "ready", "O add-in local devia ficar pronto para reabertura.");
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-storage-indexeddb",
      anchorEmailKey: "msg-anchor",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    await storage.repository.writeCase(caseValue);
    const readBack = await storage.repository.readCase(caseValue.caseId);
    assert(readBack?.caseId === caseValue.caseId, "O caso intermédio nao voltou do storage local do add-in.");
    return "Sem pasta definida, o add-in local ficou como fallback intermédio principal.";
  });

  await runScenario(scenarios, "prepare-storage-technical-memory-fallback", "storage", "Sem pasta e sem add-in local o runtime cai para memória técnica", async () => {
    await withIndexedDbAvailabilityOverride(undefined, async () => {
      const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
      assert(storage.mode === "memory", "Sem pasta e sem IndexedDB o runtime devia cair para memória.");
      assert(storage.availability === "fallback_memory", "A disponibilidade devia marcar fallback técnico.");
    });
    return "Memória ficou confirmada apenas como fallback técnico temporário.";
  });

  await runScenario(scenarios, "prepare-storage-disabled", "storage", "Storage intermédio desativado", async () => {
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "disabled" }));
    assert(storage.mode === "memory", "Storage disabled devia usar adapter in-memory.");
    assert(storage.availability === "disabled", "A disponibilidade devia ser disabled.");
    return "Modo disabled confirmado sem promessa falsa de persistencia local.";
  });

  await runScenario(scenarios, "prepare-workset-manifest-filters", "prepare", "Workset de Preparar preserva filtros, emails e anexos selecionados", async () => {
    const runtime = resolveGroupStorageRuntime({ ...DEFAULT_GROUP_STORAGE_SETTINGS, mode: "supabase" });
    const manifest = buildPrepareWorksetManifest({
      anchorEmailKey: "anchor@email|base",
      settings: runtime.settings,
      runtime,
      selectedEmailKeys: ["anchor@email|base", "known@email|related"],
      selectedAttachmentKeys: ["doc-1"],
      attachmentRows: [
        { key: "doc-1", emailKey: "anchor@email|base", name: "doc.pdf", size: 1024, hasContent: true },
        { key: "doc-2", emailKey: "known@email|related", name: "foto.png", size: 256, hasContent: true, isInline: true },
      ],
      workingGroupId: "grp-main",
      workingGroupName: "Grupo Main",
      filterQuery: "cliente",
      attachmentMode: "with",
      groupMode: "with_group",
    });
    assert(manifest?.includedEmailKeys.length === 2, "O manifesto devia guardar os emails selecionados.");
    assert(manifest?.filters.query === "cliente", "O filtro textual devia persistir no manifesto.");
    assert(manifest?.attachments[0]?.selection === "selected", "A selecao do anexo principal devia persistir.");
    return "Manifesto de workset confirmou conjunto interno, filtros e anexos selecionados.";
  });

  await runScenario(scenarios, "classification-principal-group-single-email", "classify", "Classificar grupo principal num email", async () => {
    const email = buildRelatedEmail({ emailKey: "msg-classify-1", attachments: [] });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [email],
      principalGroupId: "grp-main",
      principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
      referenceGroupIds: [],
      referenceGroups: [],
      selectedLabels: [],
      inheritedLabels: [],
      selectedLabelStates: {},
      categorizedLabelNames: [],
      selectedTicketId: "",
      selectedSeriesId: "",
      selectedTicket: null,
      ticketStatusDraft: "",
      classificationMetaDraft: {},
    });
    const baseCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-classify-1",
      anchorEmailKey: email.emailKey,
      outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    const projected = projectApplyIntoIntermediateCase({
      classificationCase: baseCase,
      resolvedApplySelection,
      resolvedCaseTicket: null,
      targetEmails: [email],
      classificationMetaDraft: {},
    });
    assert(projected.nextClassificationCase.emails[0]?.classification.principalGroupId === "grp-main", "O grupo principal nao foi projetado para o caso.");
    return "Grupo principal aplicado localmente ao email alvo.";
  });

  await runScenario(scenarios, "classification-principal-plus-references", "classify", "Classificar grupo principal e referencias", async () => {
    const email = buildRelatedEmail({ emailKey: "msg-classify-2" });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [email],
      principalGroupId: "grp-main",
      principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
      referenceGroupIds: ["grp-ref-1", "grp-ref-2"],
      referenceGroups: [buildGroup({ id: "grp-ref-1", name: "Ref 1" }), buildGroup({ id: "grp-ref-2", name: "Ref 2" })],
      selectedLabels: [],
      inheritedLabels: [],
      selectedLabelStates: {},
      categorizedLabelNames: [],
      selectedTicketId: "",
      selectedSeriesId: "",
      selectedTicket: null,
      ticketStatusDraft: "",
      classificationMetaDraft: { referenceCategorize: true },
    });
    const projected = projectApplyIntoIntermediateCase({
      classificationCase: buildPrepareIntermediateCaseFromSources({
        caseId: "case-classify-2",
        anchorEmailKey: email.emailKey,
        outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      }),
      resolvedApplySelection,
      resolvedCaseTicket: null,
      targetEmails: [email],
      classificationMetaDraft: { referenceCategorize: true },
    });
    const classification = projected.nextClassificationCase.emails[0]?.classification;
    assert(classification?.principalGroupId === "grp-main", "O grupo principal nao ficou persistido.");
    assert(classification?.referenceGroupIds.length === 2, "As referencias nao ficaram persistidas.");
    return "Grupo principal e referencias projetados no caso intermédio.";
  });

  await runScenario(scenarios, "classification-labels-and-ticket", "classify", "Classificar labels e ticket de Grupos", async () => {
    const email = buildRelatedEmail({ emailKey: "msg-classify-3" });
    const ticket = buildTicket({ id: "ticket-77", code: "TK-077", status: "em_progresso" });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [email],
      principalGroupId: "",
      principalGroup: null,
      referenceGroupIds: [],
      referenceGroups: [],
      selectedLabels: ["Financeiro", "Urgente"],
      inheritedLabels: ["Base"],
      selectedLabelStates: { Financeiro: "em_analise" },
      categorizedLabelNames: ["Financeiro"],
      selectedTicketId: ticket.id,
      selectedSeriesId: "",
      selectedTicket: ticket,
      ticketStatusDraft: "em_progresso",
      classificationMetaDraft: { ticketStatusEnabled: true, ticketStatusCategorize: true },
    });
    const projected = projectApplyIntoIntermediateCase({
      classificationCase: buildPrepareIntermediateCaseFromSources({
        caseId: "case-classify-3",
        anchorEmailKey: email.emailKey,
        outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      }),
      resolvedApplySelection,
      resolvedCaseTicket: ticket,
      targetEmails: [email],
      classificationMetaDraft: { ticketStatusEnabled: true, ticketStatusCategorize: true },
    });
    const classification = projected.nextClassificationCase.emails[0]?.classification;
    assert(classification?.labels.includes("Financeiro"), "A label Financeiro nao ficou persistida.");
    assert(classification?.ticketIds.includes(ticket.id), "O ticket de Grupos nao ficou persistido.");
    return "Labels, estados de label e ticket de Grupos projetados no caso.";
  });

  await runScenario(scenarios, "classification-multi-scope-no-global-blind-apply", "classify", "Apply multiplo continua por email alvo e por scope", async () => {
    const emailA = buildRelatedEmail({ emailKey: "msg-classify-a", subject: "A" });
    const emailB = buildRelatedEmail({ emailKey: "msg-classify-b", subject: "B" });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [emailA, emailB],
      principalGroupId: "grp-main",
      principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
      referenceGroupIds: [],
      referenceGroups: [],
      selectedLabels: ["Financeiro"],
      inheritedLabels: [],
      selectedLabelStates: {},
      categorizedLabelNames: [],
      selectedTicketId: "",
      selectedSeriesId: "",
      selectedTicket: null,
      ticketStatusDraft: "",
      classificationMetaDraft: {},
    });
    const plan = buildResolvedRemoteApplyExecutionPlan({
      targetEmails: [emailA, emailB],
      resolvedApplySelection,
      currentContext: {
        itemId: emailA.itemId,
        internetMessageId: emailA.internetMessageId,
        conversationId: emailA.conversationId,
        subject: emailA.subject,
        fromEmail: emailA.fromEmail,
        fromName: emailA.fromName,
        receivedAtIso: emailA.receivedAtIso,
      },
      emailContextMeta: new Map(),
    });
    assert(plan.targetPlans.length === 2, "O plano remoto devia manter dois alvos explicitos.");
    assert(plan.targetEmailKeys.every((key) => key === emailA.emailKey || key === emailB.emailKey), "O plano remoto nao devia expandir para o caso inteiro.");
    return "O plano remoto ficou restrito aos emails alvo escolhidos.";
  });

  await runScenario(scenarios, "classification-attachments-owner-state", "classify", "Anexos persistem por email dono", async () => {
    const emailA = buildRelatedEmail({
      emailKey: "msg-classify-attach",
      attachments: [{
        key: "doc-1",
        id: "doc-1",
        name: "doc.pdf",
        contentType: "application/pdf",
        size: 2048,
        hasContent: true,
        documentState: "processed",
        isHidden: true,
      } as RelatedEmailAttachmentFixture],
    });
    const baseCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-classify-attach",
      anchorEmailKey: emailA.emailKey,
      outlookEmails: [buildPrepareEmailInput({
        emailKey: emailA.emailKey,
        attachments: [{ key: "doc-1", name: "doc.pdf", hasContent: true, documentState: "ingested" }],
      })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [emailA],
      principalGroupId: "",
      principalGroup: null,
      referenceGroupIds: [],
      referenceGroups: [],
      selectedLabels: [],
      inheritedLabels: [],
      selectedLabelStates: {},
      categorizedLabelNames: [],
      selectedTicketId: "",
      selectedSeriesId: "",
      selectedTicket: null,
      ticketStatusDraft: "",
      classificationMetaDraft: {},
    });
    const projected = projectApplyIntoIntermediateCase({
      classificationCase: baseCase,
      resolvedApplySelection,
      resolvedCaseTicket: null,
      targetEmails: [emailA],
      classificationMetaDraft: {},
    });
    const attachment = projected.nextClassificationCase.emails[0]?.attachments[0];
    assert(attachment?.documentState === "processed", "O documentState do anexo do email dono nao foi preservado.");
    assert(attachment?.isHidden === true, "O isHidden do anexo do email dono nao foi preservado.");
    return "documentState/isHidden seguiram o email dono sem contaminar outros emails.";
  });

  await runScenario(scenarios, "classification-reopen-and-rehydrate", "classify", "Reabrir caso e reidratar a partir do add-in local", async () => {
    await saveSettings({
      groups: buildGroupsSettings({
        tab: { ...DEFAULT_GROUPS_MODULE_SETTINGS.tab, storageMode: "local_indexeddb", baseFolderPath: "" },
      }),
    });
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-reopen",
      anchorEmailKey: "msg-reopen",
      outlookEmails: [buildPrepareEmailInput({
        emailKey: "msg-reopen",
        principalGroupId: "grp-main",
        labels: ["Financeiro"],
        attachments: [{ key: "doc-1", name: "doc.pdf", hasContent: true }],
      })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    const attachmentPath = caseValue.emails[0]?.attachments[0]?.localRef?.value;
    await storage.repository.writeCase(caseValue);
    if (attachmentPath && "writeBinary" in storage.adapter) {
      await storage.adapter.writeBinary(attachmentPath, new Blob(["fixture-binary"]));
    }
    const resolved = await resolveClassificationIntermediateCase({
      caseId: "case-reopen",
      anchorEmailKey: "msg-reopen",
    });
    assert(resolved.caseValue?.caseId === "case-reopen", "O caso nao foi resolvido por caseId.");
    const hydratedEmails = await hydrateIntermediateCaseEmailsToRelatedEntries({
      caseValue: resolved.caseValue,
      adapter: resolved.storage.adapter,
    });
    const hydrated = hydratedEmails[0];
    assert(hydrated.classificationMeta?.principalGroupId === "grp-main", "A reidratação nao preservou o grupo principal.");
    assert(Boolean(hydrated.attachments?.[0]?.content), "A reidratacao nao recuperou o binario local do anexo.");
    return "Classificar reabriu o caso pelo storage local do add-in e reidratou classificacao e anexo local.";
  });

  await runScenario(scenarios, "classification-folder-backed-handoff", "classify", "Classificar mantém o estado quando o intermédio arranca numa pasta local", async () => {
    await withMockedFetch(async () => {
      const folderPath = "C:/dados/grupos/intermedio-classify";
      await saveSettings({
        groups: buildGroupsSettings({
          storage: {
            ...DEFAULT_GROUP_STORAGE_SETTINGS,
            mode: "supabase",
          },
          tab: {
            ...DEFAULT_GROUPS_MODULE_SETTINGS.tab,
            storageMode: "local_indexeddb",
            baseFolderPath: folderPath,
          },
        }),
      });
      const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: folderPath }));
      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-folder-classify",
        anchorEmailKey: "msg-folder-classify",
        outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-folder-classify" })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      await storage.repository.writeCase(classificationCase);
      const resolvedBefore = await resolveClassificationIntermediateCase({
        caseId: classificationCase.caseId,
        anchorEmailKey: classificationCase.anchorEmailKey,
      });
      assert(resolvedBefore.storage.mode === "filesystem", "O Classificar devia reabrir a partir da pasta intermédia.");
      const nextClassificationCase = {
        ...classificationCase,
        emails: classificationCase.emails.map((email) => ({
          ...email,
          classification: {
            ...email.classification,
            principalGroupId: "grp-main",
            principalGroupName: "Grupo Main",
          },
        })),
      };
      await resolvedBefore.storage.repository.writeCase(nextClassificationCase);
      const readBack = await storage.repository.readCase(classificationCase.caseId);
      assert(readBack?.emails[0]?.classification.principalGroupId === "grp-main", "O estado classificado nao ficou preservado na pasta intermédia antes da promoção final.");
    });
    return "O Classificar reusou a mesma pasta intermédia e preservou o estado antes da gravação final.";
  });

  await runScenario(scenarios, "classification-repeat-apply-no-duplicates", "classify", "Reaplicar nao duplica classificacao", async () => {
    const email = buildRelatedEmail({ emailKey: "msg-repeat" });
    const resolvedApplySelection = buildResolvedStudioApplySelection({
      targetEmails: [email],
      principalGroupId: "grp-main",
      principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
      referenceGroupIds: ["grp-ref"],
      referenceGroups: [buildGroup({ id: "grp-ref", name: "Grupo Ref" })],
      selectedLabels: ["Financeiro"],
      inheritedLabels: [],
      selectedLabelStates: { Financeiro: "em_analise" },
      categorizedLabelNames: ["Financeiro"],
      selectedTicketId: "ticket-1",
      selectedSeriesId: "",
      selectedTicket: buildTicket({ id: "ticket-1", code: "TK-1" }),
      ticketStatusDraft: "em_analise",
      classificationMetaDraft: {},
    });
    let caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-repeat",
      anchorEmailKey: email.emailKey,
      outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    caseValue = projectApplyIntoIntermediateCase({
      classificationCase: caseValue,
      resolvedApplySelection,
      resolvedCaseTicket: buildTicket({ id: "ticket-1", code: "TK-1" }),
      targetEmails: [email],
      classificationMetaDraft: {},
    }).nextClassificationCase;
    caseValue = projectApplyIntoIntermediateCase({
      classificationCase: caseValue,
      resolvedApplySelection,
      resolvedCaseTicket: buildTicket({ id: "ticket-1", code: "TK-1" }),
      targetEmails: [email],
      classificationMetaDraft: {},
    }).nextClassificationCase;
    const classification = caseValue.emails[0]?.classification;
    assert(classification?.referenceGroupIds.length === 1, "As referencias ficaram duplicadas.");
    assert(classification?.labels.length === 1, "As labels ficaram duplicadas.");
    assert(classification?.ticketIds.length === 1, "Os tickets ficaram duplicados.");
    return "Reaplicar a mesma selecao manteve arrays canonicos sem duplicacao.";
  });

  await runScenario(scenarios, "classification-controlled-missing-case", "classify", "Lookup sem caso devolve estado controlado", async () => {
    await saveSettings({
      groups: buildGroupsSettings({
        tab: { ...DEFAULT_GROUPS_MODULE_SETTINGS.tab, storageMode: "disabled", baseFolderPath: "" },
      }),
    });
    const resolved = await resolveClassificationIntermediateCase({
      caseId: "case-missing",
      anchorEmailKey: "msg-missing",
    });
    assert(resolved.caseValue === null, "Sem caso persistido o lookup devia devolver null.");
    assert(resolved.lookup === "none", "Sem caso persistido o lookup devia ficar em 'none'.");
    assert(resolved.storage.availability === "disabled", "O storage devia refletir o modo disabled.");
    return "Falha controlada confirmada quando nao existe caso e o storage intermédio esta desligado.";
  });

  await runScenario(scenarios, "classification-server-write-principal-reopen", "classify", "Preparar -> Classificar grava grupo principal no backend e reabre coerente", async () => {
    return await withMockedFetch(async (mock) => {
      const email = buildRelatedEmail({
        emailKey: "item-proof-1",
        itemId: "item-proof-1",
        internetMessageId: "<proof-1@example.com>",
        conversationId: "conv-proof-1",
        subject: "Proof 1",
        receivedAtIso: FIXED_NOW_ISO,
      });
      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-proof-1",
        anchorEmailKey: email.emailKey,
        conversationId: email.conversationId,
        outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey, subject: email.subject })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      const resolvedApplySelection = buildResolvedStudioApplySelection({
        targetEmails: [email],
        principalGroupId: "grp-main",
        principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
        referenceGroupIds: [],
        referenceGroups: [],
        selectedLabels: [],
        inheritedLabels: [],
        selectedLabelStates: {},
        categorizedLabelNames: [],
        selectedTicketId: "",
        selectedSeriesId: "",
        selectedTicket: null,
        ticketStatusDraft: "",
        classificationMetaDraft: {},
      });
      return await executeFinalPersistenceProof({
        mock,
        scenarioId: "classification-server-write-principal-reopen",
        title: "Preparar -> Classificar grava grupo principal no backend e reabre coerente",
        classificationCase,
        targetEmails: [email],
        selectedEmail: email,
        currentContext: {
          itemId: email.itemId,
          internetMessageId: email.internetMessageId,
          conversationId: email.conversationId,
          subject: email.subject,
          fromEmail: email.fromEmail,
          fromName: email.fromName,
          receivedAtIso: email.receivedAtIso,
          toRecipients: email.toRecipients,
          ccRecipients: email.ccRecipients,
        },
        resolvedApplySelection,
        classificationMetaDraft: {},
        relevantSettings: {
          intermediate: "indexeddb_addin_primary",
          final: "server_backend",
        },
        expected: {
          serverContext: {
            email: {
              emailKey: "item-proof-1",
              labels: [],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [],
            },
            emails: [{
              emailKey: "item-proof-1",
              labels: [],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [],
            }],
            groups: [{ id: "grp-main", name: "grp-main" }],
            tickets: [],
          },
          persistedGroups: [{ id: "grp-main", name: "grp-main", memberEmailKeys: ["item-proof-1"] }],
          persistedTickets: [],
        },
        writeProofs,
      });
    });
  });

  await runScenario(scenarios, "classification-server-write-groups-labels-reopen", "classify", "Grupo principal, referencias e etiquetas reaparecem na reabertura", async () => {
    return await withMockedFetch(async (mock) => {
      const email = buildRelatedEmail({
        emailKey: "item-proof-2",
        itemId: "item-proof-2",
        internetMessageId: "<proof-2@example.com>",
        conversationId: "conv-proof-2",
        subject: "Proof 2",
        receivedAtIso: FIXED_NOW_ISO,
      });
      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-proof-2",
        anchorEmailKey: email.emailKey,
        conversationId: email.conversationId,
        outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey, subject: email.subject })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      const resolvedApplySelection = buildResolvedStudioApplySelection({
        targetEmails: [email],
        principalGroupId: "grp-main",
        principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
        referenceGroupIds: ["grp-ref-1", "grp-ref-2"],
        referenceGroups: [
          buildGroup({ id: "grp-ref-1", name: "Ref 1" }),
          buildGroup({ id: "grp-ref-2", name: "Ref 2" }),
        ],
        selectedLabels: ["Financeiro", "Urgente"],
        inheritedLabels: [],
        selectedLabelStates: { Financeiro: "em_analise" },
        categorizedLabelNames: ["Financeiro"],
        selectedTicketId: "",
        selectedSeriesId: "",
        selectedTicket: null,
        ticketStatusDraft: "",
        classificationMetaDraft: { referenceCategorize: true },
      });
      return await executeFinalPersistenceProof({
        mock,
        scenarioId: "classification-server-write-groups-labels-reopen",
        title: "Grupo principal, referencias e etiquetas reaparecem na reabertura",
        classificationCase,
        targetEmails: [email],
        selectedEmail: email,
        currentContext: {
          itemId: email.itemId,
          internetMessageId: email.internetMessageId,
          conversationId: email.conversationId,
          subject: email.subject,
          fromEmail: email.fromEmail,
          fromName: email.fromName,
          receivedAtIso: email.receivedAtIso,
          toRecipients: email.toRecipients,
          ccRecipients: email.ccRecipients,
        },
        resolvedApplySelection,
        classificationMetaDraft: { referenceCategorize: true },
        relevantSettings: {
          intermediate: "indexeddb_addin_primary",
          final: "server_backend",
          labels: true,
        },
        expected: {
          serverContext: {
            email: {
              emailKey: "item-proof-2",
              labels: ["Financeiro", "Urgente"],
              relatedGroups: [
                { id: "grp-main", relationKind: "principal" },
                { id: "grp-ref-1", relationKind: "referencia" },
                { id: "grp-ref-2", relationKind: "referencia" },
              ],
              attachments: [],
            },
            emails: [{
              emailKey: "item-proof-2",
              labels: ["Financeiro", "Urgente"],
              relatedGroups: [
                { id: "grp-main", relationKind: "principal" },
                { id: "grp-ref-1", relationKind: "referencia" },
                { id: "grp-ref-2", relationKind: "referencia" },
              ],
              attachments: [],
            }],
            groups: [
              { id: "grp-main", name: "grp-main" },
              { id: "grp-ref-1", name: "grp-ref-1" },
              { id: "grp-ref-2", name: "grp-ref-2" },
            ],
            tickets: [],
          },
          persistedGroups: [
            { id: "grp-main", name: "grp-main", memberEmailKeys: ["item-proof-2"] },
            { id: "grp-ref-1", name: "grp-ref-1", memberEmailKeys: ["item-proof-2"] },
            { id: "grp-ref-2", name: "grp-ref-2", memberEmailKeys: ["item-proof-2"] },
          ],
          persistedTickets: [],
        },
        writeProofs,
      });
    });
  });

  await runScenario(scenarios, "classification-server-write-multi-scope-targeted", "classify", "Scope multiplo altera so os alvos e reabre coerente", async () => {
    return await withMockedFetch(async (mock) => {
      const emailA = buildRelatedEmail({
        emailKey: "item-proof-3a",
        itemId: "item-proof-3a",
        internetMessageId: "<proof-3a@example.com>",
        conversationId: "conv-proof-3",
        subject: "Proof 3A",
        receivedAtIso: FIXED_NOW_ISO,
      });
      const emailB = buildRelatedEmail({
        emailKey: "item-proof-3b",
        itemId: "item-proof-3b",
        internetMessageId: "<proof-3b@example.com>",
        conversationId: "conv-proof-3",
        subject: "Proof 3B",
        receivedAtIso: FIXED_NOW_ISO,
      });
      const emailUntouched = buildRelatedEmail({
        emailKey: "item-proof-3c",
        itemId: "item-proof-3c",
        internetMessageId: "<proof-3c@example.com>",
        conversationId: "conv-proof-3",
        subject: "Proof 3C",
        receivedAtIso: FIXED_NOW_ISO,
        relatedGroups: [{ id: "grp-legacy", name: "Legacy", kind: "custom", relationKind: "principal" }],
      });
      mock.serverState.emailsByKey.set(emailUntouched.emailKey || "", cloneJson(emailUntouched));
      ensureMockGroup(mock.serverState, "grp-legacy", "Legacy");
      mock.serverState.groupMembersById.get("grp-legacy")?.add(emailUntouched.emailKey || "");

      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-proof-3",
        anchorEmailKey: emailA.emailKey,
        conversationId: emailA.conversationId,
        outlookEmails: [
          buildPrepareEmailInput({ emailKey: emailA.emailKey, subject: emailA.subject }),
          buildPrepareEmailInput({ emailKey: emailB.emailKey, subject: emailB.subject }),
          buildPrepareEmailInput({ emailKey: emailUntouched.emailKey, subject: emailUntouched.subject }),
        ],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      const resolvedApplySelection = buildResolvedStudioApplySelection({
        targetEmails: [emailA, emailB],
        principalGroupId: "grp-main",
        principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
        referenceGroupIds: [],
        referenceGroups: [],
        selectedLabels: ["Financeiro"],
        inheritedLabels: [],
        selectedLabelStates: {},
        categorizedLabelNames: [],
        selectedTicketId: "",
        selectedSeriesId: "",
        selectedTicket: null,
        ticketStatusDraft: "",
        classificationMetaDraft: {},
      });
      return await executeFinalPersistenceProof({
        mock,
        scenarioId: "classification-server-write-multi-scope-targeted",
        title: "Scope multiplo altera so os alvos e reabre coerente",
        classificationCase,
        targetEmails: [emailA, emailB],
        selectedEmail: emailA,
        currentContext: {
          itemId: emailA.itemId,
          internetMessageId: emailA.internetMessageId,
          conversationId: emailA.conversationId,
          subject: emailA.subject,
          fromEmail: emailA.fromEmail,
          fromName: emailA.fromName,
          receivedAtIso: emailA.receivedAtIso,
          toRecipients: emailA.toRecipients,
          ccRecipients: emailA.ccRecipients,
        },
        resolvedApplySelection,
        classificationMetaDraft: {},
        relevantSettings: {
          intermediate: "indexeddb_addin_primary",
          final: "server_backend",
          multiScope: true,
        },
        expected: {
          serverContext: {
            email: {
              emailKey: "item-proof-3a",
              labels: ["Financeiro"],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [],
            },
            emails: [
              {
                emailKey: "item-proof-3a",
                labels: ["Financeiro"],
                relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
                attachments: [],
              },
              {
                emailKey: "item-proof-3b",
                labels: ["Financeiro"],
                relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
                attachments: [],
              },
              {
                emailKey: "item-proof-3c",
                labels: [],
                relatedGroups: [{ id: "grp-legacy", relationKind: "principal" }],
                attachments: [],
              },
            ],
            groups: [
              { id: "grp-legacy", name: "Legacy" },
              { id: "grp-main", name: "grp-main" },
            ],
            tickets: [],
          },
          persistedGroups: [
            { id: "grp-legacy", name: "Legacy", memberEmailKeys: ["item-proof-3c"] },
            { id: "grp-main", name: "grp-main", memberEmailKeys: ["item-proof-3a", "item-proof-3b"] },
          ],
          persistedTickets: [],
        },
        writeProofs,
      });
    });
  });

  await runScenario(scenarios, "classification-server-write-attachments-reopen", "classify", "Anexos persistem no write final e reabrem por email dono", async () => {
    return await withMockedFetch(async (mock) => {
      const email = buildRelatedEmail({
        emailKey: "item-proof-4",
        itemId: "item-proof-4",
        internetMessageId: "<proof-4@example.com>",
        conversationId: "conv-proof-4",
        subject: "Proof 4",
        receivedAtIso: FIXED_NOW_ISO,
        attachments: [{
          key: "doc-proof-4",
          id: "doc-proof-4",
          name: "doc-proof-4.pdf",
          contentType: "application/pdf",
          size: 2048,
          hasContent: true,
          content: "ZmFrZQ==",
          documentState: "processed",
          isHidden: true,
        } as RelatedEmailAttachmentFixture],
      });
      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-proof-4",
        anchorEmailKey: email.emailKey,
        conversationId: email.conversationId,
        outlookEmails: [buildPrepareEmailInput({
          emailKey: email.emailKey,
          subject: email.subject,
          attachments: [{ key: "doc-proof-4", name: "doc-proof-4.pdf", hasContent: true, documentState: "processed" }],
        })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      const resolvedApplySelection = buildResolvedStudioApplySelection({
        targetEmails: [email],
        principalGroupId: "grp-main",
        principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
        referenceGroupIds: [],
        referenceGroups: [],
        selectedLabels: [],
        inheritedLabels: [],
        selectedLabelStates: {},
        categorizedLabelNames: [],
        selectedTicketId: "",
        selectedSeriesId: "",
        selectedTicket: null,
        ticketStatusDraft: "",
        classificationMetaDraft: {},
      });
      return await executeFinalPersistenceProof({
        mock,
        scenarioId: "classification-server-write-attachments-reopen",
        title: "Anexos persistem no write final e reabrem por email dono",
        classificationCase,
        targetEmails: [email],
        selectedEmail: email,
        currentContext: {
          itemId: email.itemId,
          internetMessageId: email.internetMessageId,
          conversationId: email.conversationId,
          subject: email.subject,
          fromEmail: email.fromEmail,
          fromName: email.fromName,
          receivedAtIso: email.receivedAtIso,
          toRecipients: email.toRecipients,
          ccRecipients: email.ccRecipients,
        },
        resolvedApplySelection,
        classificationMetaDraft: {},
        relevantSettings: {
          intermediate: "indexeddb_addin_primary",
          final: "server_backend",
          attachments: true,
        },
        attachmentStorageOptions: {
          attachmentStorageProvider: "cloud",
          attachmentStorageBasePath: "groups/final",
        },
        expected: {
          serverContext: {
            email: {
              emailKey: "item-proof-4",
              labels: [],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [{
                key: "doc-proof-4",
                name: "doc-proof-4.pdf",
                documentState: "processed",
                isHidden: true,
                storageProvider: "cloud",
                storageBasePath: "groups/final",
              }],
            },
            emails: [{
              emailKey: "item-proof-4",
              labels: [],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [{
                key: "doc-proof-4",
                name: "doc-proof-4.pdf",
                documentState: "processed",
                isHidden: true,
                storageProvider: "cloud",
                storageBasePath: "groups/final",
              }],
            }],
            groups: [{ id: "grp-main", name: "grp-main" }],
            tickets: [],
          },
          persistedGroups: [{ id: "grp-main", name: "grp-main", memberEmailKeys: ["item-proof-4"] }],
          persistedTickets: [],
        },
        writeProofs,
      });
    });
  });

  await runScenario(scenarios, "classification-server-write-ticket-reopen", "classify", "Ticket de Grupos persiste no backend e reaparece na reabertura", async () => {
    return await withMockedFetch(async (mock) => {
      const email = buildRelatedEmail({
        emailKey: "item-proof-5",
        itemId: "item-proof-5",
        internetMessageId: "<proof-5@example.com>",
        conversationId: "conv-proof-5",
        subject: "Proof 5",
        receivedAtIso: FIXED_NOW_ISO,
      });
      const classificationCase = buildPrepareIntermediateCaseFromSources({
        caseId: "case-proof-5",
        anchorEmailKey: email.emailKey,
        conversationId: email.conversationId,
        outlookEmails: [buildPrepareEmailInput({ emailKey: email.emailKey, subject: email.subject })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      const resolvedApplySelection = buildResolvedStudioApplySelection({
        targetEmails: [email],
        principalGroupId: "grp-main",
        principalGroup: buildGroup({ id: "grp-main", name: "Grupo Main" }),
        referenceGroupIds: [],
        referenceGroups: [],
        selectedLabels: ["Financeiro"],
        inheritedLabels: [],
        selectedLabelStates: { Financeiro: "em_analise" },
        categorizedLabelNames: ["Financeiro"],
        selectedTicketId: "",
        selectedSeriesId: "series-proof",
        selectedTicket: null,
        ticketStatusDraft: "em_progresso",
        classificationMetaDraft: { ticketStatusEnabled: true, ticketStatusCategorize: true },
      });
      return await executeFinalPersistenceProof({
        mock,
        scenarioId: "classification-server-write-ticket-reopen",
        title: "Ticket de Grupos persiste no backend e reaparece na reabertura",
        classificationCase,
        targetEmails: [email],
        selectedEmail: email,
        currentContext: {
          itemId: email.itemId,
          internetMessageId: email.internetMessageId,
          conversationId: email.conversationId,
          subject: email.subject,
          fromEmail: email.fromEmail,
          fromName: email.fromName,
          receivedAtIso: email.receivedAtIso,
          toRecipients: email.toRecipients,
          ccRecipients: email.ccRecipients,
        },
        resolvedApplySelection,
        classificationMetaDraft: { ticketStatusEnabled: true, ticketStatusCategorize: true },
        relevantSettings: {
          intermediate: "indexeddb_addin_primary",
          final: "server_backend",
          tickets: true,
        },
        createTicketTitle: "Ticket Proof 5",
        expected: {
          serverContext: {
            email: {
              emailKey: "item-proof-5",
              labels: ["Financeiro"],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [],
            },
            emails: [{
              emailKey: "item-proof-5",
              labels: ["Financeiro"],
              relatedGroups: [{ id: "grp-main", relationKind: "principal" }],
              attachments: [],
            }],
            groups: [{ id: "grp-main", name: "grp-main" }],
            tickets: [{
              id: "ticket-1",
              code: "TK-001",
              status: "em_progresso",
              groupIds: ["grp-main"],
              emailLinked: true,
            }],
          },
          persistedGroups: [{ id: "grp-main", name: "grp-main", memberEmailKeys: ["item-proof-5"] }],
          persistedTickets: [{
            id: "ticket-1",
            code: "TK-001",
            status: "em_progresso",
            groupIds: ["grp-main"],
            emailKeys: ["item-proof-5"],
          }],
        },
        writeProofs,
      });
    });
  });

  await runScenario(scenarios, "workset-cloud-roundtrip", "storage", "Workset cloud roundtrip com fetch mockado", async () => {
    await withMockedFetch(async (calls) => {
      const runtime = resolveGroupStorageRuntime({ ...DEFAULT_GROUP_STORAGE_SETTINGS, mode: "supabase" });
      const manifest = buildManifestForRuntime("supabase");
      const saved = await savePrimaryGroupWorkset({ runtime, manifest });
      const loaded = await loadPrimaryGroupWorkset({ runtime, anchorEmailKey: manifest.anchorEmailKey });
      assert(saved?.worksetKey === manifest.worksetKey, "O save do manifesto cloud falhou.");
      assert(loaded?.worksetKey === manifest.worksetKey, "O load do manifesto cloud falhou.");
      assert(calls.some((call) => call.method === "POST" && call.url.includes("/api/links/groups/worksets")), "O save do workset cloud nao chamou o endpoint.");
    });
    return "Save/load do manifesto cloud confirmado com fetch mockado.";
  });

  await runScenario(scenarios, "workset-local-device-roundtrip", "storage", "Workset local_device usa runtime path-based correto", async () => {
    await withMockedFetch(async (calls) => {
      const runtime = resolveGroupStorageRuntime({
        ...DEFAULT_GROUP_STORAGE_SETTINGS,
        mode: "local_device",
        localDevice: { rootPath: "C:/tmp/groups/local-device" },
      });
      const manifest = buildManifestForRuntime("local_device");
      await savePrimaryGroupWorkset({ runtime, manifest });
      await loadPrimaryGroupWorkset({ runtime, anchorEmailKey: manifest.anchorEmailKey });
      const getCall = calls.find((call) => call.method === "GET");
      assert(getCall?.url.includes("mode=local_device"), "O load do workset local_device nao passou o modo certo.");
      assert(getCall?.url.includes("basePath=C%3A%2Ftmp%2Fgroups%2Flocal-device"), "O load do workset local_device nao passou o path certo.");
    });
    return "Save/load do manifesto local_device confirmou parametrizacao do runtime.";
  });

  await runScenario(scenarios, "workset-chosen-folder-roundtrip", "storage", "Workset chosen_folder usa runtime path-based correto", async () => {
    await withMockedFetch(async (calls) => {
      const runtime = resolveGroupStorageRuntime({
        ...DEFAULT_GROUP_STORAGE_SETTINGS,
        mode: "chosen_folder",
        chosenFolder: { path: "C:/tmp/groups/chosen", kind: "filesystem" },
      });
      const manifest = buildManifestForRuntime("chosen_folder");
      await savePrimaryGroupWorkset({ runtime, manifest });
      await loadPrimaryGroupWorkset({ runtime, anchorEmailKey: manifest.anchorEmailKey });
      const getCall = calls.find((call) => call.method === "GET");
      assert(getCall?.url.includes("mode=chosen_folder"), "O load do workset chosen_folder nao passou o modo certo.");
      assert(getCall?.url.includes("chosenFolderKind=filesystem"), "O load do workset chosen_folder nao passou o kind certo.");
    });
    return "Save/load do manifesto chosen_folder confirmou parametrizacao do runtime.";
  });

  await runScenario(scenarios, "workset-hybrid-roundtrip", "storage", "Workset hybrid usa runtime path-based correto", async () => {
    await withMockedFetch(async (calls) => {
      const runtime = resolveGroupStorageRuntime({
        ...DEFAULT_GROUP_STORAGE_SETTINGS,
        mode: "hybrid",
        chosenFolder: { path: "C:/tmp/groups/hybrid", kind: "filesystem" },
        hybrid: { primaryTarget: "chosen_folder", promoteManifestOnSave: true, promoteAttachmentMetadataOnSave: true },
      });
      const manifest = buildManifestForRuntime("hybrid");
      await savePrimaryGroupWorkset({ runtime, manifest });
      await loadPrimaryGroupWorkset({ runtime, anchorEmailKey: manifest.anchorEmailKey });
      const getCall = calls.find((call) => call.method === "GET");
      assert(getCall?.url.includes("mode=hybrid"), "O load do workset hybrid nao passou o modo certo.");
      assert(getCall?.url.includes("primaryTarget=chosen_folder"), "O load do workset hybrid nao passou o primaryTarget.");
    });
    return "Save/load do manifesto hybrid confirmou parametrizacao do runtime.";
  });

  await runScenario(scenarios, "migration-copy-and-safety", "migration", "Migracao por copia e gates de seguranca", async () => {
    await withMockedFetch(async () => {
      const sourceNamespace = "C:/dados/grupos/migrate-source";
      const targetNamespace = "C:/dados/grupos/migrate-target";
      const sourceStorage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: sourceNamespace }));
      const targetStorage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: targetNamespace }));
      const caseValue = buildPrepareIntermediateCaseFromSources({
        caseId: "case-migrate",
        anchorEmailKey: "msg-migrate",
        outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-migrate", attachments: [{ key: "doc-1", name: "doc.pdf", hasContent: true }] })],
        serverEmails: [],
        nowIso: FIXED_NOW_ISO,
      });
      await sourceStorage.repository.writeCase(caseValue);
      const attachmentPath = caseValue.emails[0]?.attachments[0]?.localRef?.value;
      if (attachmentPath && "writeBinary" in sourceStorage.adapter) {
        await sourceStorage.adapter.writeBinary(attachmentPath, new Blob(["binary-fixture"]));
      }

      const copyResult = await migrateIntermediateCaseNamespace({
        sourceNamespace,
        targetNamespace,
        mode: "copy",
        mergeExistingData: false,
        strictMigrationSafety: true,
      });
      assert(copyResult.migratedCases === 1, "A migracao por copia devia migrar um caso.");
      assert(await targetStorage.repository.readCase(caseValue.caseId), "O caso nao apareceu na localizacao de destino.");
      assert(await sourceStorage.repository.readCase(caseValue.caseId), "O caso de origem devia continuar no modo copy.");
      if (attachmentPath && "readBinary" in targetStorage.adapter) {
        assert(await targetStorage.adapter.readBinary(attachmentPath), "Os binarios do caso nao foram copiados.");
      }

      let blockedMove = false;
      try {
        await migrateIntermediateCaseNamespace({
          sourceNamespace,
          targetNamespace,
          mode: "move",
          allowMoveExistingData: false,
          mergeExistingData: true,
          strictMigrationSafety: false,
        });
      } catch {
        blockedMove = true;
      }
      assert(blockedMove, "O move devia ser bloqueado quando allowMoveExistingData=false.");
    });
    return "Migracao copy funcionou entre localizacoes intermédias e os gates de seguranca bloquearam movimento indevido.";
  });

  await runScenario(scenarios, "cleanup-retention-and-mixed-protection", "cleanup", "Limpeza real respeita retention e mixed cases", async () => {
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
    const repository = storage.repository;

    const promotedCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-promoted",
      anchorEmailKey: "msg-promoted",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-promoted" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    promotedCase.emails[0].serverPresence = "classified";
    promotedCase.emails[0].localPresence = "none";
    promotedCase.lastAccessedAt = "2026-03-01T00:00:00.000Z";
    promotedCase.updatedAt = promotedCase.lastAccessedAt;

    const localOnlyCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-local",
      anchorEmailKey: "msg-local",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-local" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    localOnlyCase.lastAccessedAt = "2025-12-01T00:00:00.000Z";
    localOnlyCase.updatedAt = localOnlyCase.lastAccessedAt;

    const mixedCase = buildPrepareIntermediateCaseFromSources({
      caseId: "case-mixed",
      anchorEmailKey: "msg-mixed",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-mixed" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    mixedCase.emails[0].serverPresence = "classified";
    mixedCase.emails[0].localPresence = "complete";
    mixedCase.lastAccessedAt = "2026-01-01T00:00:00.000Z";
    mixedCase.updatedAt = mixedCase.lastAccessedAt;

    await repository.writeCase(promotedCase);
    await repository.writeCase(localOnlyCase);
    await repository.writeCase(mixedCase);

    const resultProtected = await cleanupIntermediateCases(
      normalizeGroupsTabSettings({
        storageMode: "local_indexeddb",
        baseFolderPath: "",
        cleanupClosedCaseDays: 10,
        cleanupAbandonedCaseDays: 10,
        neverDeleteMixedSilently: true,
      }),
      { nowMs: FIXED_NOW_MS }
    );
    assert(resultProtected.deletedCases === 2, "Promoted/local_only antigos deviam ser limpos.");
    assert(resultProtected.skippedMixedCases === 1, "O mixed case devia ser preservado quando neverDeleteMixedSilently=true.");

    const resultDeleteMixed = await cleanupIntermediateCases(
      normalizeGroupsTabSettings({
        storageMode: "local_indexeddb",
        baseFolderPath: "",
        cleanupClosedCaseDays: 10,
        cleanupAbandonedCaseDays: 10,
        neverDeleteMixedSilently: false,
      }),
      { nowMs: FIXED_NOW_MS }
    );
    assert(resultDeleteMixed.deletedCases === 1, "O mixed case antigo devia ser limpo quando a protecao esta desligada.");
    return "Limpeza real confirmou thresholds, promoted/local_only e protecao de mixed cases.";
  });

  await runScenario(scenarios, "settings-influence-storage-runtime", "storage", "Mudar storage settings altera o runtime resolvido", async () => {
    const cloudRuntime = resolveGroupStorageRuntime({ ...DEFAULT_GROUP_STORAGE_SETTINGS, mode: "supabase" });
    const localRuntime = resolveGroupStorageRuntime({ ...DEFAULT_GROUP_STORAGE_SETTINGS, mode: "local_device", localDevice: { rootPath: "C:/tmp/groups" } });
    assert(cloudRuntime.mode === "supabase", "O runtime cloud nao foi resolvido.");
    assert(localRuntime.mode === "local_device", "O runtime local_device nao foi resolvido.");
    assert(localRuntime.legacyBridge.baseFolderPath === "C:/tmp/groups", "O local_device nao propagou o root path.");
    return "O runtime de storage respondeu corretamente a mudancas de modo e path.";
  });

  await runScenario(scenarios, "settings-influence-attachment-policy", "attachments", "Mudar settings de anexos altera a politica aplicada", async () => {
    const settings = {
      ...DEFAULT_GROUP_STORAGE_SETTINGS,
      ignoreInlineAttachments: true,
      attachmentPromptThresholdMb: 1,
    };
    const policy = resolveGroupAttachmentStoragePolicy(settings);
    const inlineDecision = resolvePreparedAttachmentStorageDecision({
      size: 500,
      isInline: true,
      hasContent: true,
    }, settings);
    const largeDecision = resolvePreparedAttachmentStorageDecision({
      size: 3 * 1024 * 1024,
      isInline: false,
      hasContent: true,
    }, settings);
    assert(policy.thresholdBytes === 1024 * 1024, "O threshold em bytes nao acompanhou o setting.");
    assert(inlineDecision.mainDisposition === "skip", "O ignoreInlineAttachments nao foi aplicado.");
    assert(largeDecision.requiresDecision === true, "Anexos acima do threshold deviam exigir decisao.");
    const legacyAttachmentOptions = buildAttachmentStorageOptions({ groups: { storage: settings } });
    assert(legacyAttachmentOptions.attachmentStorageProvider === "cloud", "A bridge legacy de anexos devia refletir o provider atual.");
    return "Threshold, inline policy e bridge legacy de anexos responderam aos settings.";
  });

  await runScenario(scenarios, "settings-influence-outlook-categories", "outlook_categories", "Mudar outlookCategories altera o plano logico local", async () => {
    const principalGroup = buildGroup({ id: "grp-main", name: "Grupo Main", status: "em_analise", labels: ["Financeiro"] });
    const ticket = buildTicket({ id: "ticket-1", code: "TK-001", status: "em_progresso", groupId: principalGroup.id });
    const email = buildRelatedEmail({
      emailKey: "msg-categories",
      relatedGroups: [{
        id: principalGroup.id,
        name: principalGroup.name,
        relationKind: "principal",
        kind: "custom",
      } as RelatedEmailGroupFixture],
      labels: ["Financeiro"],
      labelStates: { Financeiro: "em_analise" },
      classificationMeta: {
        principalStatusEnabled: true,
        principalStatusCategorize: true,
        ticketStatusEnabled: true,
        ticketStatusCategorize: true,
        ticketId: ticket.id,
        categorizedLabelNames: ["Financeiro"],
      },
    });
    const sourceWithLabels = buildOutlookCategorySourceFromRelatedContext({
      email,
      groups: [principalGroup],
      tickets: [ticket],
      settings: {
        groups: buildGroupsSettings({
          labels: { managerEnabled: true, catalog: [{ label: "Financeiro", categorize: true, hasStatus: true, status: "em_analise" }], favoriteIds: [] },
          outlookCategories: {
            enabled: true,
            includeGroups: true,
            includeTickets: true,
            includeStatuses: true,
            includeLabels: true,
          },
        }),
      } as Pick<CockpitSettingsV1, "groups">,
      currentOutlookLabelNames: ["Legacy Label"],
    });
    const sourceWithoutLabels = buildOutlookCategorySourceFromRelatedContext({
      email,
      groups: [principalGroup],
      tickets: [ticket],
      settings: {
        groups: buildGroupsSettings({
          outlookCategories: {
            enabled: true,
            includeGroups: true,
            includeTickets: true,
            includeStatuses: true,
            includeLabels: false,
          },
        }),
      } as Pick<CockpitSettingsV1, "groups">,
      currentOutlookLabelNames: ["Legacy Label"],
    });
    const plan = buildOutlookCategoryPlan(sourceWithLabels);
    assert(sourceWithLabels.labelNames.includes("Financeiro"), "includeLabels=true devia incluir labels no source.");
    assert(sourceWithoutLabels.labelNames.length === 0, "includeLabels=false devia retirar labels do source.");
    assert(plan.desiredCategories.some((entry) => entry.includes("Grupo Main")), "O plano devia incluir a categoria do grupo.");
    assert(plan.desiredCategories.some((entry) => entry.includes("TK-001")), "O plano devia incluir a categoria do ticket.");
    return "O plano logico de Outlook categories respondeu aos toggles canonicos.";
  });

  const failed = scenarios.filter((scenario) => scenario.status === "failed").length;
  return {
    generatedAtIso: new Date().toISOString(),
    settingsMatrix: GROUPS_SETTINGS_MATRIX,
    scenarios,
    writeProofs,
    passed: scenarios.length - failed,
    failed,
  };
}
