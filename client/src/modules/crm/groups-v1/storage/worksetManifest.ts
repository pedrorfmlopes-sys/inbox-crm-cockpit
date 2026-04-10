import type {
  GroupStorageLocationPointer,
  GroupWorksetManifest,
  GroupWorksetPromotionStatus,
} from "./types";
import { buildGroupStorageSessionDraft } from "./sessionDraft";

function normalizeString(value: unknown): string {
  return String(value || "").trim();
}

function normalizePromotionStatus(input: Partial<GroupWorksetPromotionStatus> | null | undefined): GroupWorksetPromotionStatus {
  const promotedScopes = Array.isArray(input?.promotedScopes) ? input.promotedScopes : [];
  const blockedScopes = Array.isArray(input?.blockedScopes) ? input.blockedScopes : [];
  return {
    state: input?.state || "not_requested",
    lastAttemptAtIso: normalizeString(input?.lastAttemptAtIso) || undefined,
    promotedScopes,
    blockedScopes,
    note: normalizeString(input?.note) || undefined,
  };
}

function normalizeLocation(input: GroupStorageLocationPointer | null | undefined): GroupStorageLocationPointer {
  return {
    kind: input?.kind || "session",
    provider: input?.provider || "cloud",
    label: normalizeString(input?.label) || "Sessao local",
    basePath: normalizeString(input?.basePath) || undefined,
    relativePath: normalizeString(input?.relativePath) || undefined,
    folderHint: normalizeString(input?.folderHint) || undefined,
    isRemote: input?.isRemote === true,
    isConfigured: input?.isConfigured === true,
  };
}

export function buildGroupWorksetManifest(input: Partial<GroupWorksetManifest> & {
  worksetKey: string;
  storageMode: GroupWorksetManifest["storageMode"];
  anchorEmailKey: string;
  mainLocation: GroupStorageLocationPointer;
}): GroupWorksetManifest {
  const draft = buildGroupStorageSessionDraft({
    anchorEmailKey: input.anchorEmailKey,
    storageMode: input.storageMode,
    selectedEmailKeys: input.includedEmailKeys,
    workingGroupId: input.workingGroupId,
    workingGroupName: input.workingGroupName,
    filters: input.filters,
    preparedAttachments: input.attachments,
  });
  return {
    kind: "groups_v1_workset_manifest",
    version: 1,
    worksetKey: normalizeString(input.worksetKey),
    createdAtIso: normalizeString(input.createdAtIso) || new Date().toISOString(),
    updatedAtIso: normalizeString(input.updatedAtIso) || new Date().toISOString(),
    storageMode: input.storageMode,
    anchorEmailKey: draft.anchorEmailKey,
    includedEmailKeys: draft.selectedEmailKeys,
    workingGroupId: draft.workingGroupId,
    workingGroupName: draft.workingGroupName,
    filters: draft.filters,
    attachments: draft.preparedAttachments,
    mainLocation: normalizeLocation(input.mainLocation),
    remotePromotionLocation: input.remotePromotionLocation ? normalizeLocation(input.remotePromotionLocation) : null,
    promotion: normalizePromotionStatus(input.promotion),
  };
}

export function normalizeGroupWorksetManifest(input: Partial<GroupWorksetManifest> | null | undefined): GroupWorksetManifest | null {
  const worksetKey = normalizeString(input?.worksetKey);
  const anchorEmailKey = normalizeString(input?.anchorEmailKey);
  const storageMode = input?.storageMode;
  if (!worksetKey || !anchorEmailKey || !storageMode) return null;
  return buildGroupWorksetManifest({
    ...input,
    worksetKey,
    anchorEmailKey,
    storageMode,
    mainLocation: normalizeLocation(input?.mainLocation),
  });
}

export function getGroupWorksetManifestSignature(input: Partial<GroupWorksetManifest> | null | undefined): string {
  const manifest = normalizeGroupWorksetManifest(input);
  if (!manifest) return "";
  return JSON.stringify({
    ...manifest,
    createdAtIso: "",
    updatedAtIso: "",
  });
}
