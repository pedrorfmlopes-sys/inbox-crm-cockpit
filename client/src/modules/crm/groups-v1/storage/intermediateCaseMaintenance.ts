import { normalizeGroupsTabSettings, type GroupsTabSettings } from "../settings/groupsTabSettings";
import {
  supportsIntermediateCaseBinaryStorage,
  type IntermediateCaseStorageAdapter,
} from "./intermediateCaseRepository";
import { resolveIntermediateCaseStorage } from "./resolveIntermediateCaseStorage";
import type { IntermediateCase, IntermediateCaseSummary } from "./intermediateCaseTypes";

function normalizeText(value: string | undefined): string {
  return String(value || "").trim();
}

function ageInDays(isoValue: string | undefined, nowMs: number): number {
  const raw = normalizeText(isoValue);
  if (!raw) return 0;
  const parsed = Date.parse(raw);
  if (!Number.isFinite(parsed)) return 0;
  return Math.floor((nowMs - parsed) / (24 * 60 * 60 * 1000));
}

async function copyCaseBinaryPayloads(args: {
  caseValue: IntermediateCase;
  sourceAdapter: IntermediateCaseStorageAdapter;
  targetAdapter: IntermediateCaseStorageAdapter;
}) {
  if (!supportsIntermediateCaseBinaryStorage(args.sourceAdapter) || !supportsIntermediateCaseBinaryStorage(args.targetAdapter)) {
    return 0;
  }

  let copied = 0;
  for (const email of args.caseValue.emails) {
    for (const attachment of email.attachments) {
      const localPath = normalizeText(attachment.localRef?.value);
      if (!localPath || !attachment.hasContent) continue;
      const blob = await args.sourceAdapter.readBinary(localPath);
      if (!blob) continue;
      await args.targetAdapter.writeBinary(localPath, blob);
      copied += 1;
    }
  }
  return copied;
}

function resolveMaintenanceStorage(locationPath: string) {
  const storage = resolveIntermediateCaseStorage(
    normalizeGroupsTabSettings({
      storageMode: "local_indexeddb",
      baseFolderPath: locationPath,
    })
  );
  if (storage.availability === "disabled") {
    throw new Error("O storage intermédio está desligado para esta operação.");
  }
  if (storage.availability === "fallback_memory") {
    throw new Error("A operação de manutenção foi bloqueada porque o runtime caiu para memória como fallback técnico.");
  }
  return storage;
}

export async function migrateIntermediateCaseNamespace(input: {
  sourceNamespace: string;
  targetNamespace: string;
  mode: "move" | "copy";
  allowMoveExistingData?: boolean;
  mergeExistingData?: boolean;
  strictMigrationSafety?: boolean;
}): Promise<{
  migratedCases: number;
  copiedAttachments: number;
  skippedCases: number;
}> {
  const sourceLocation = normalizeText(input.sourceNamespace);
  const targetLocation = normalizeText(input.targetNamespace);
  if (sourceLocation === targetLocation) {
    return { migratedCases: 0, copiedAttachments: 0, skippedCases: 0 };
  }

  const sourceStorage = resolveMaintenanceStorage(sourceLocation);
  const targetStorage = resolveMaintenanceStorage(targetLocation);
  const sourceAdapter = sourceStorage.adapter;
  const targetAdapter = targetStorage.adapter;
  const sourceRepository = sourceStorage.repository;
  const targetRepository = targetStorage.repository;
  const summaries = await sourceRepository.listCases();
  const targetSummaries = await targetRepository.listCases();
  const strictMigrationSafety = input.strictMigrationSafety !== false;
  const mergeExistingData = input.mergeExistingData === true;
  const allowMoveExistingData = input.allowMoveExistingData === true;

  if (input.mode === "move" && !allowMoveExistingData) {
    throw new Error("O modo de movimento está bloqueado pelos settings atuais. Ativa 'Permitir mover dados existentes' para continuar.");
  }

  if (targetSummaries.length && !mergeExistingData) {
    throw new Error("A localização intermédia de destino já contém casos. Ativa a fusão de dados existentes ou escolhe um destino vazio.");
  }

  const targetCaseIds = new Set(targetSummaries.map((summary) => summary.caseId));
  const conflictingCaseIds = summaries
    .map((summary) => summary.caseId)
    .filter((caseId) => targetCaseIds.has(caseId));
  if (strictMigrationSafety && conflictingCaseIds.length) {
    throw new Error(
      `A migração foi bloqueada pela segurança estrita porque o destino já contém ${conflictingCaseIds.length} caso(s) com o mesmo identificador.`
    );
  }

  let migratedCases = 0;
  let copiedAttachments = 0;
  let skippedCases = 0;

  for (const summary of summaries) {
    const caseValue = await sourceRepository.readCase(summary.caseId);
    if (!caseValue) {
      skippedCases += 1;
      continue;
    }
    await targetRepository.writeCase(caseValue);
    copiedAttachments += await copyCaseBinaryPayloads({
      caseValue,
      sourceAdapter,
      targetAdapter,
    });
    migratedCases += 1;
    if (input.mode === "move") {
      await sourceRepository.deleteCase(caseValue.caseId);
    }
  }

  return {
    migratedCases,
    copiedAttachments,
    skippedCases,
  };
}

function shouldDeleteIntermediateCase(
  summary: IntermediateCaseSummary,
  settings: GroupsTabSettings,
  nowMs: number
): boolean {
  const days = ageInDays(summary.lastAccessedAt || summary.updatedAt, nowMs);
  if (summary.retentionState === "promoted") {
    return days >= settings.cleanupClosedCaseDays;
  }
  if (summary.retentionState === "local_only") {
    return days >= settings.cleanupAbandonedCaseDays;
  }
  if (summary.retentionState === "mixed") {
    if (settings.neverDeleteMixedSilently) return false;
    return days >= settings.cleanupClosedCaseDays;
  }
  return false;
}

export async function cleanupIntermediateCases(
  settingsLike: GroupsTabSettings | null | undefined,
  options?: {
    nowMs?: number;
  }
): Promise<{
  namespace: string;
  deletedCases: number;
  skippedMixedCases: number;
  inspectedCases: number;
}> {
  const settings = normalizeGroupsTabSettings(settingsLike || null);
  const namespace = normalizeText(settings.baseFolderPath);
  if (settings.storageMode === "disabled") {
    return {
      namespace,
      deletedCases: 0,
      skippedMixedCases: 0,
      inspectedCases: 0,
    };
  }

  const storage = resolveIntermediateCaseStorage(settings);
  if (storage.availability === "fallback_memory") {
    return {
      namespace,
      deletedCases: 0,
      skippedMixedCases: 0,
      inspectedCases: 0,
    };
  }

  const repository = storage.repository;
  const summaries = await repository.listCases();
  const nowMs = Number.isFinite(Number(options?.nowMs)) ? Number(options?.nowMs) : Date.now();
  let deletedCases = 0;
  let skippedMixedCases = 0;

  for (const summary of summaries) {
    if (summary.retentionState === "mixed" && settings.neverDeleteMixedSilently) {
      skippedMixedCases += 1;
      continue;
    }
    if (!shouldDeleteIntermediateCase(summary, settings, nowMs)) continue;
    await repository.deleteCase(summary.caseId);
    deletedCases += 1;
  }

  return {
    namespace,
    deletedCases,
    skippedMixedCases,
    inspectedCases: summaries.length,
  };
}
