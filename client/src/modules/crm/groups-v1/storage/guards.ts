import type { GroupStorageMode, GroupWorksetManifest } from "./types";

function normalizeKeyPart(value: string | null | undefined): string {
  return String(value || "").trim();
}

export function buildGroupWorksetKey(anchorEmailKey: string | null | undefined): string {
  const anchor = normalizeKeyPart(anchorEmailKey);
  return anchor ? `groups_v1_workset:${anchor}` : "";
}

export function supportsPrimaryGroupWorksetPersistence(mode: GroupStorageMode): boolean {
  return mode === "supabase" || mode === "hybrid";
}

export function hasMeaningfulGroupWorksetPayload(input: Partial<GroupWorksetManifest> | null | undefined): boolean {
  const manifest = input || null;
  if (!normalizeKeyPart(manifest?.worksetKey) || !normalizeKeyPart(manifest?.anchorEmailKey)) {
    return false;
  }
  return Boolean(
    Array.isArray(manifest?.includedEmailKeys) && manifest!.includedEmailKeys.length
    || Array.isArray(manifest?.attachments) && manifest!.attachments.length
    || normalizeKeyPart(manifest?.workingGroupId)
    || normalizeKeyPart(manifest?.workingGroupName)
    || normalizeKeyPart(manifest?.filters?.query)
    || normalizeKeyPart(manifest?.filters?.fromEmail)
    || Array.isArray(manifest?.filters?.labels) && manifest!.filters!.labels!.length
    || normalizeKeyPart(manifest?.filters?.dateFromIso)
    || normalizeKeyPart(manifest?.filters?.dateToIso)
    || normalizeKeyPart(manifest?.filters?.attachmentMode) && manifest?.filters?.attachmentMode !== "all"
    || normalizeKeyPart(manifest?.filters?.groupMode) && manifest?.filters?.groupMode !== "all"
  );
}
