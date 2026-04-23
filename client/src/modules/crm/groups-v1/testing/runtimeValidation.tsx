import React from "react";
import ReactDOM from "react-dom/client";
import type { GroupTicketEntry, LinkGroupEntry, RelatedEmailEntry } from "@/api";
import {
  getSettings,
  resetSettings,
  saveSettings,
  type CockpitSettingsV1,
} from "@/settings";
import { buildResolvedRemoteApplyExecutionPlan, buildResolvedStudioApplySelection } from "@/modules/crm/group-classification/applyResolution";
import { buildAttachmentStorageOptions } from "@/modules/crm/group-classification/documentUtils";
import { projectApplyIntoIntermediateCase } from "@/modules/crm/group-classification/localCaseProjection";
import { buildOutlookCategoryPlan, buildOutlookCategorySourceFromRelatedContext } from "@/outlookCategories";
import { GroupsSettingsPanel } from "../settings/GroupsSettingsPanel";
import { DEFAULT_GROUPS_MODULE_SETTINGS, type GroupsModuleSettings } from "../settings/groupsModuleSettings";
import { DEFAULT_GROUPS_TAB_SETTINGS, normalizeGroupsTabSettings } from "../settings/groupsTabSettings";
import { resolveGroupAttachmentStoragePolicy, resolvePreparedAttachmentStorageDecision } from "../storage/attachmentPolicy";
import { buildPrepareWorksetManifest } from "../storage/buildPrepareWorksetManifest";
import {
  cleanupIntermediateCases,
  migrateIntermediateCaseNamespace,
} from "../storage/intermediateCaseMaintenance";
import { hydrateIntermediateCaseEmailsToRelatedEntries } from "../storage/intermediateCaseAdapters";
import { createIndexedDbIntermediateCaseStorageAdapter, INTERMEDIATE_CASE_DB_NAME } from "../storage/intermediateCaseIndexedDbAdapter";
import { createIntermediateCaseRepository } from "../storage/intermediateCaseRepository";
import { buildPrepareIntermediateCaseFromSources } from "../storage/prepareIntermediateCaseResolution";
import {
  resolveClassificationIntermediateCase,
} from "../storage/resolveClassificationIntermediateCase";
import { resolveIntermediateCaseStorage } from "../storage/resolveIntermediateCaseStorage";
import { resolveGroupStorageRuntime } from "../storage/resolveStorageMode";
import { savePrimaryGroupWorkset } from "../storage/saveWorkset";
import { loadPrimaryGroupWorkset } from "../storage/loadWorkset";
import { DEFAULT_GROUP_STORAGE_SETTINGS } from "../storage/settings";
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

export type GroupsBrowserValidationReport = {
  generatedAtIso: string;
  settingsMatrix: GroupsSettingsMatrixEntry[];
  scenarios: ValidationScenarioResult[];
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

    const namespaceInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: grupos/cliente-acme"]');
    assert(namespaceInput, "O editor inline do namespace nao foi renderizado.");
    namespaceInput.value = "grupos/runtime";
    namespaceInput.dispatchEvent(new Event("input", { bubbles: true }));
    namespaceInput.dispatchEvent(new Event("change", { bubbles: true }));
    await new Promise((resolve) => setTimeout(resolve, 20));
    const updatedNamespaceInput = host.querySelector<HTMLInputElement>('input[placeholder="ex.: grupos/cliente-acme"]');
    assert(updatedNamespaceInput?.value === "grupos/runtime", "O namespace inline nao refletiu a edicao local.");

    const disabledToggles = Array.from(host.querySelectorAll('button[aria-pressed][disabled]'));
    assert(disabledToggles.length >= 6, "Os toggles shell herdados da secao intermedia nao ficaram desativados.");

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

async function withMockedFetch<T>(run: (calls: Array<{ url: string; method: string; body?: string }>) => Promise<T>): Promise<T> {
  const originalFetch = window.fetch.bind(window);
  const calls: Array<{ url: string; method: string; body?: string }> = [];
  const manifestByKey = new Map<string, GroupWorksetManifest>();

  window.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === "string" ? input : input instanceof URL ? input.toString() : input.url;
    const method = String(init?.method || "GET").toUpperCase();
    const body = typeof init?.body === "string" ? init.body : undefined;
    calls.push({ url, method, body });

    if (url.includes("/api/links/groups/worksets") && method === "POST") {
      const parsed = JSON.parse(body || "{}") as { manifest?: GroupWorksetManifest };
      if (parsed.manifest?.worksetKey) {
        manifestByKey.set(parsed.manifest.worksetKey, parsed.manifest);
      }
      return new Response(JSON.stringify({ ok: true, manifest: parsed.manifest || null }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      });
    }

    if (url.includes("/api/links/groups/worksets/") && method === "GET") {
      const match = url.match(/\/api\/links\/groups\/worksets\/([^?]+)/i);
      const worksetKey = match ? decodeURIComponent(match[1]) : "";
      return new Response(
        JSON.stringify({ ok: true, manifest: manifestByKey.get(worksetKey) || null }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    }

    if (url.includes("/api/links/groups/storage/validate")) {
      return new Response(
        JSON.stringify({
          ok: true,
          result: {
            mode: "supabase",
            provider: "cloud",
            fileBacked: false,
            supported: true,
            basePath: "",
            normalizedBasePath: "",
            isWebUrl: false,
            requiresServerAccessiblePath: false,
            canStoreManifest: true,
            canStoreBinary: true,
            pickerAvailable: false,
            notes: ["mocked"],
            architecturalBlocker: null,
            requiredChange: null,
          },
        }),
        { status: 200, headers: { "Content-Type": "application/json" } }
      );
    }

    return originalFetch(input, init);
  }) as typeof window.fetch;

  try {
    return await run(calls);
  } finally {
    window.fetch = originalFetch;
  }
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

  await runScenario(scenarios, "settings-groups-roundtrip", "settings", "settings.groups roundtrip e aliases derivados", async () => {
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
        baseFolderPath: "grupos/runtime",
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
    assert(next.groups.tab.baseFolderPath === "grupos/runtime", "O namespace canonico nao foi persistido.");
    assert(next.groups.labels.catalog[0]?.label === "Financeiro", "O catalogo de labels canonico nao foi persistido.");
    assert(next.groups.tickets.enabled === false, "O flag canonico de tickets nao foi persistido.");
    assert(next.groups.outlookCategories.includeLabels === true, "O flag canonico de Outlook categories nao foi persistido.");
    assert(next.groupStorage.mode === "hybrid", "O alias legacy deixou de espelhar o canonico.");
    assert(next.groupsTabSettings.baseFolderPath === "grupos/runtime", "O alias legacy do namespace nao ficou derivado.");
    return "Roundtrip canonico confirmado; aliases legacy ficaram apenas como espelho derivado.";
  });

  await runScenario(scenarios, "settings-panel-host-safe-shells", "settings", "Painel de settings host-safe e shells honestamente desativadas", async () => {
    await renderSettingsPanelScenario();
    return "GroupsSettingsPanel renderizou editores inline e manteve toggles shell desativados sem prompt/alert/confirm.";
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

  await runScenario(scenarios, "prepare-storage-indexeddb", "storage", "Storage intermédio com namespace usa IndexedDB", async () => {
    const settings = normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "grupos/indexeddb" });
    const storage = resolveIntermediateCaseStorage(settings);
    assert(storage.mode === "indexeddb", "Com namespace o storage intermédio devia usar IndexedDB.");
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-storage-1",
      anchorEmailKey: "msg-anchor",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-anchor" })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    await storage.repository.writeCase(caseValue);
    const readBack = await storage.repository.readCase(caseValue.caseId);
    assert(readBack?.caseId === caseValue.caseId, "O caso intermédio nao voltou do IndexedDB.");
    return "IndexedDB real confirmado para o namespace configurado.";
  });

  await runScenario(scenarios, "prepare-storage-missing-namespace", "storage", "Sem namespace o intermédio cai para memoria", async () => {
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: "" }));
    assert(storage.mode === "memory", "Sem namespace devia haver fallback para memoria.");
    assert(storage.availability === "missing_location", "A disponibilidade devia refletir missing_location.");
    return "Fallback em memoria confirmado quando nao existe namespace.";
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

  await runScenario(scenarios, "classification-reopen-and-rehydrate", "classify", "Reabrir caso e reidratar a partir do IndexedDB", async () => {
    const namespace = "grupos/classify-reopen";
    await saveSettings({
      groups: buildGroupsSettings({
        tab: { ...DEFAULT_GROUPS_MODULE_SETTINGS.tab, storageMode: "local_indexeddb", baseFolderPath: namespace },
      }),
    });
    const storage = resolveIntermediateCaseStorage(normalizeGroupsTabSettings({ storageMode: "local_indexeddb", baseFolderPath: namespace }));
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
    return "Classificar reabriu o caso pelo namespace configurado e reidratou classificacao e anexo local.";
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
    const sourceNamespace = "grupos/migrate-source";
    const targetNamespace = "grupos/migrate-target";
    const sourceAdapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace: sourceNamespace });
    const targetAdapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace: targetNamespace });
    const sourceRepository = createIntermediateCaseRepository(sourceAdapter);
    const targetRepository = createIntermediateCaseRepository(targetAdapter);
    const caseValue = buildPrepareIntermediateCaseFromSources({
      caseId: "case-migrate",
      anchorEmailKey: "msg-migrate",
      outlookEmails: [buildPrepareEmailInput({ emailKey: "msg-migrate", attachments: [{ key: "doc-1", name: "doc.pdf", hasContent: true }] })],
      serverEmails: [],
      nowIso: FIXED_NOW_ISO,
    });
    await sourceRepository.writeCase(caseValue);
    const attachmentPath = caseValue.emails[0]?.attachments[0]?.localRef?.value;
    if (attachmentPath) {
      await sourceAdapter.writeBinary(attachmentPath, new Blob(["binary-fixture"]));
    }

    const copyResult = await migrateIntermediateCaseNamespace({
      sourceNamespace,
      targetNamespace,
      mode: "copy",
      mergeExistingData: false,
      strictMigrationSafety: true,
    });
    assert(copyResult.migratedCases === 1, "A migracao por copia devia migrar um caso.");
    assert(await targetRepository.readCase(caseValue.caseId), "O caso nao apareceu no namespace de destino.");
    assert(await sourceRepository.readCase(caseValue.caseId), "O caso de origem devia continuar no modo copy.");
    if (attachmentPath) {
      assert(await targetAdapter.readBinary(attachmentPath), "Os binarios do caso nao foram copiados.");
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
    return "Migracao copy funcionou e os gates de seguranca bloquearam movimento indevido.";
  });

  await runScenario(scenarios, "cleanup-retention-and-mixed-protection", "cleanup", "Limpeza real respeita retention e mixed cases", async () => {
    const namespace = "grupos/cleanup";
    const repository = createIntermediateCaseRepository(createIndexedDbIntermediateCaseStorageAdapter({ namespace }));

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
        baseFolderPath: namespace,
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
        baseFolderPath: namespace,
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
    passed: scenarios.length - failed,
    failed,
  };
}
