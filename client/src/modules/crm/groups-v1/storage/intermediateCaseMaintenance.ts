import { normalizeGroupsTabSettings, type GroupsTabSettings } from "../settings/groupsTabSettings";
import { createIndexedDbIntermediateCaseStorageAdapter } from "./intermediateCaseIndexedDbAdapter";
import { createIntermediateCaseRepository, supportsIntermediateCaseBinaryStorage } from "./intermediateCaseRepository";
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
  sourceAdapter: ReturnType<typeof createIndexedDbIntermediateCaseStorageAdapter>;
  targetAdapter: ReturnType<typeof createIndexedDbIntermediateCaseStorageAdapter>;
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

export async function migrateIntermediateCaseNamespace(input: {
  sourceNamespace: string;
  targetNamespace: string;
  mode: "move" | "copy";
}): Promise<{
  migratedCases: number;
  copiedAttachments: number;
  skippedCases: number;
}> {
  const sourceNamespace = normalizeText(input.sourceNamespace);
  const targetNamespace = normalizeText(input.targetNamespace);
  if (!sourceNamespace || !targetNamespace) {
    throw new Error("Indica namespaces de origem e destino para a migracao do intermédio.");
  }
  if (sourceNamespace === targetNamespace) {
    return { migratedCases: 0, copiedAttachments: 0, skippedCases: 0 };
  }

  const sourceAdapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace: sourceNamespace });
  const targetAdapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace: targetNamespace });
  const sourceRepository = createIntermediateCaseRepository(sourceAdapter);
  const targetRepository = createIntermediateCaseRepository(targetAdapter);
  const summaries = await sourceRepository.listCases();

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
  settingsLike: GroupsTabSettings | null | undefined
): Promise<{
  namespace: string;
  deletedCases: number;
  skippedMixedCases: number;
  inspectedCases: number;
}> {
  const settings = normalizeGroupsTabSettings(settingsLike || null);
  const namespace = normalizeText(settings.baseFolderPath);
  if (settings.storageMode === "disabled" || !namespace) {
    return {
      namespace,
      deletedCases: 0,
      skippedMixedCases: 0,
      inspectedCases: 0,
    };
  }

  const adapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace });
  const repository = createIntermediateCaseRepository(adapter);
  const summaries = await repository.listCases();
  const nowMs = Date.now();
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
