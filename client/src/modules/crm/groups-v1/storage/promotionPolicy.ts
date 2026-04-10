import type {
  GroupPromotionPolicy,
  GroupPromotionScope,
  GroupStorageLocationPointer,
  GroupStorageSettings,
} from "./types";

export function resolveGroupPromotionPolicy(
  settings: GroupStorageSettings,
  mainPersistence: GroupStorageLocationPointer,
  remotePromotionLocation?: GroupStorageLocationPointer | null
): GroupPromotionPolicy {
  const remoteAllowed = settings.mode === "supabase" || settings.mode === "hybrid";
  return {
    mainPersistence,
    remotePromotionLocation: remoteAllowed ? remotePromotionLocation || null : null,
    allowRemotePromotion: remoteAllowed && settings.supabase.allowPromotion,
    promoteManifestOnPrimarySave: settings.mode === "supabase"
      ? settings.supabase.promoteManifestOnSave
      : settings.mode === "hybrid"
        ? settings.hybrid.promoteManifestOnSave
        : false,
    promoteAttachmentMetadataOnPrimarySave: settings.mode === "supabase"
      ? settings.supabase.promoteAttachmentMetadataOnSave
      : settings.mode === "hybrid"
        ? settings.hybrid.promoteAttachmentMetadataOnSave
        : false,
    promoteAttachmentBinaryOnPrimarySave: settings.mode === "supabase"
      ? settings.supabase.promoteAttachmentBinaryOnSave
      : false,
    requireExplicitRemotePromotion: settings.mode !== "supabase",
    requireFreshPayloadBeforeOverwrite: true,
    saveSessionBeforeContextChange: true,
    saveSessionBeforeExit: true,
  };
}

export function shouldAutoPromoteGroupScope(policy: GroupPromotionPolicy, scope: GroupPromotionScope): boolean {
  if (!policy.allowRemotePromotion) return false;
  if (scope === "manifest") return policy.promoteManifestOnPrimarySave;
  if (scope === "attachment_metadata" || scope === "email_metadata") {
    return policy.promoteAttachmentMetadataOnPrimarySave;
  }
  if (scope === "attachment_binary") return policy.promoteAttachmentBinaryOnPrimarySave;
  return false;
}
