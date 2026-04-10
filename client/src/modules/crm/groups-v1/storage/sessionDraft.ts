import type {
  GroupPreparedAttachmentDescriptor,
  GroupStorageMode,
  GroupStorageSessionDraft,
  GroupWorksetFilterSnapshot,
} from "./types";

function normalizeStringArray(values: string[] | null | undefined): string[] {
  return Array.from(new Set((Array.isArray(values) ? values : []).map((value) => String(value || "").trim()).filter(Boolean)));
}

function normalizeFilters(input: GroupWorksetFilterSnapshot | null | undefined): GroupWorksetFilterSnapshot {
  return {
    query: String(input?.query || "").trim() || undefined,
    fromEmail: String(input?.fromEmail || "").trim() || undefined,
    labels: normalizeStringArray(input?.labels),
    dateFromIso: String(input?.dateFromIso || "").trim() || undefined,
    dateToIso: String(input?.dateToIso || "").trim() || undefined,
    attachmentMode: input?.attachmentMode === "with" || input?.attachmentMode === "without" ? input.attachmentMode : "all",
    groupMode: input?.groupMode === "with_group" || input?.groupMode === "without_group" ? input.groupMode : "all",
  };
}

function normalizeAttachments(
  entries: GroupPreparedAttachmentDescriptor[] | null | undefined
): GroupPreparedAttachmentDescriptor[] {
  const rows = Array.isArray(entries) ? entries : [];
  return rows
    .map((entry) => ({
      key: String(entry?.key || "").trim(),
      emailKey: String(entry?.emailKey || "").trim(),
      name: String(entry?.name || "").trim(),
      contentType: String(entry?.contentType || "").trim() || undefined,
      size: Number(entry?.size || 0) || undefined,
      isInline: entry?.isInline === true,
      hasContent: entry?.hasContent === true,
      selection: entry?.selection === "rejected" ? "rejected" : entry?.selection === "pending" ? "pending" : "selected",
      storageDisposition: entry?.storageDisposition === "reference" ? "reference" : entry?.storageDisposition === "skip" ? "skip" : entry?.storageDisposition === "binary" ? "binary" : undefined,
      requiresDecision: entry?.requiresDecision === true,
    }))
    .filter((entry) => entry.key && entry.emailKey && entry.name);
}

export function buildGroupStorageSessionDraft(input: Partial<GroupStorageSessionDraft> & {
  anchorEmailKey: string;
  storageMode: GroupStorageMode;
}): GroupStorageSessionDraft {
  return {
    kind: "groups_v1_storage_session_draft",
    version: 1,
    savedAtIso: String(input.savedAtIso || new Date().toISOString()),
    storageMode: input.storageMode,
    anchorEmailKey: String(input.anchorEmailKey || "").trim(),
    selectedEmailKeys: normalizeStringArray(input.selectedEmailKeys),
    expandedEmailKeys: normalizeStringArray(input.expandedEmailKeys),
    workingGroupId: String(input.workingGroupId || "").trim() || undefined,
    workingGroupName: String(input.workingGroupName || "").trim() || undefined,
    filters: normalizeFilters(input.filters),
    preparedAttachments: normalizeAttachments(input.preparedAttachments),
  };
}
