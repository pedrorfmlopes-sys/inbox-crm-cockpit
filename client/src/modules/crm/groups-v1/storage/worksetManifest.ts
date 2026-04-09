import type {
  GroupStorageLocationPointer,
  GroupWorksetManifest,
  GroupWorksetPromotionStatus,
} from "./types";
import { buildGroupStorageSessionDraft } from "./sessionDraft";

function buildPromotionStatus(input: Partial<GroupWorksetPromotionStatus> | null | undefined): GroupWorksetPromotionStatus {
  const promotedScopes = Array.isArray(input?.promotedScopes) ? input.promotedScopes : [];
  const blockedScopes = Array.isArray(input?.blockedScopes) ? input.blockedScopes : [];
  return {
    state: input?.state || "not_requested",
    lastAttemptAtIso: String(input?.lastAttemptAtIso || "").trim() || undefined,
    promotedScopes,
    blockedScopes,
    note: String(input?.note || "").trim() || undefined,
  };
}

function normalizeLocation(input: GroupStorageLocationPointer | null | undefined): GroupStorageLocationPointer {
  return {
    kind: input?.kind || "session",
    provider: input?.provider || "cloud",
    label: String(input?.label || "Sessao local").trim(),
    basePath: String(input?.basePath || "").trim() || undefined,
    relativePath: String(input?.relativePath || "").trim() || undefined,
    folderHint: String(input?.folderHint || "").trim() || undefined,
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
    worksetKey: String(input.worksetKey || "").trim(),
    createdAtIso: String(input.createdAtIso || new Date().toISOString()),
    updatedAtIso: String(input.updatedAtIso || new Date().toISOString()),
    storageMode: input.storageMode,
    anchorEmailKey: draft.anchorEmailKey,
    includedEmailKeys: draft.selectedEmailKeys,
    workingGroupId: draft.workingGroupId,
    workingGroupName: draft.workingGroupName,
    filters: draft.filters,
    attachments: draft.preparedAttachments,
    mainLocation: normalizeLocation(input.mainLocation),
    remotePromotionLocation: input.remotePromotionLocation ? normalizeLocation(input.remotePromotionLocation) : null,
    promotion: buildPromotionStatus(input.promotion),
  };
}
