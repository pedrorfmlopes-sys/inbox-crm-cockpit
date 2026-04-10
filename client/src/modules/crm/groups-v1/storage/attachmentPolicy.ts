import type { GroupAttachmentStoragePolicy, GroupPreparedAttachmentDescriptor, GroupStorageSettings } from "./types";

const BYTES_PER_MB = 1024 * 1024;

export function getGroupAttachmentPromptThresholdBytes(settings: Pick<GroupStorageSettings, "attachmentPromptThresholdMb">): number {
  const raw = Number(settings.attachmentPromptThresholdMb || 0);
  const safeMb = Number.isFinite(raw) && raw > 0 ? Math.min(Math.max(raw, 1), 250) : 10;
  return Math.round(safeMb * BYTES_PER_MB);
}

export function resolveGroupAttachmentStoragePolicy(settings: GroupStorageSettings): GroupAttachmentStoragePolicy {
  const thresholdBytes = getGroupAttachmentPromptThresholdBytes(settings);
  const mainBinaryStrategy = settings.mode === "supabase" ? "embed_binary" : "store_reference";
  const remoteBinaryStrategy = settings.mode === "supabase" && settings.supabase.promoteAttachmentBinaryOnSave
    ? "prompt_user"
    : "store_reference";
  return {
    thresholdBytes,
    ignoreInlineAttachments: settings.ignoreInlineAttachments,
    mainBinaryStrategy,
    remoteBinaryStrategy,
    promptWhenAboveThreshold: true,
    neverAutoPromoteBinary: true,
  };
}

export function resolvePreparedAttachmentStorageDecision(
  attachment: Pick<GroupPreparedAttachmentDescriptor, "size" | "isInline" | "hasContent">,
  settings: GroupStorageSettings
): {
  requiresDecision: boolean;
  mainDisposition: "binary" | "reference" | "skip";
  remoteDisposition: "manual" | "blocked" | "reference_only";
} {
  const policy = resolveGroupAttachmentStoragePolicy(settings);
  if (policy.ignoreInlineAttachments && attachment.isInline) {
    return {
      requiresDecision: false,
      mainDisposition: "skip",
      remoteDisposition: "blocked",
    };
  }

  const size = Number(attachment.size || 0);
  const requiresDecision = Boolean(attachment.hasContent) && size > policy.thresholdBytes;
  return {
    requiresDecision,
    mainDisposition: policy.mainBinaryStrategy === "embed_binary" ? "binary" : "reference",
    remoteDisposition: requiresDecision ? "manual" : "reference_only",
  };
}
