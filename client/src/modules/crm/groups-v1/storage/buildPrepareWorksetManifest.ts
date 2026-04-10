import { resolvePreparedAttachmentStorageDecision } from "./attachmentPolicy";
import { buildGroupWorksetKey } from "./guards";
import { buildGroupWorksetManifest } from "./worksetManifest";
import type { ResolvedGroupStorageRuntime } from "./resolveStorageMode";
import type { GroupStorageSettings, GroupWorksetManifest } from "./types";

type PrepareAttachmentInput = {
  key: string;
  emailKey: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  hasContent: boolean;
};

type PrepareBuildArgs = {
  anchorEmailKey: string;
  settings: GroupStorageSettings;
  runtime: ResolvedGroupStorageRuntime;
  selectedEmailKeys: string[];
  selectedAttachmentKeys: string[];
  attachmentRows: PrepareAttachmentInput[];
  workingGroupId?: string;
  workingGroupName?: string;
  filterQuery?: string;
  attachmentMode?: "all" | "with" | "without";
  groupMode?: "all" | "with_group" | "without_group";
  previous?: GroupWorksetManifest | null;
};

function buildPromotionNote(args: PrepareBuildArgs): string {
  if (args.runtime.mode === "hybrid") {
    return "Manifesto persistido com pointers locais; promocao remota de anexos continua separada e controlada.";
  }
  return "Manifesto persistido no destino principal; anexos continuam metadata/reference-first sem promocao binaria automatica.";
}

export function buildPrepareWorksetManifest(args: PrepareBuildArgs): GroupWorksetManifest | null {
  const worksetKey = buildGroupWorksetKey(args.anchorEmailKey);
  if (!worksetKey) return null;

  const attachments = args.attachmentRows.map((attachment) => {
    const decision = resolvePreparedAttachmentStorageDecision({
      size: attachment.size,
      isInline: attachment.isInline,
      hasContent: attachment.hasContent,
    }, args.settings);
    return {
      key: attachment.key,
      emailKey: attachment.emailKey,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      hasContent: attachment.hasContent,
      selection: args.selectedAttachmentKeys.includes(attachment.key) ? "selected" : "rejected",
      storageDisposition: decision.mainDisposition,
      requiresDecision: decision.requiresDecision,
    } as const;
  });

  return buildGroupWorksetManifest({
    createdAtIso: args.previous?.createdAtIso,
    updatedAtIso: new Date().toISOString(),
    worksetKey,
    storageMode: args.runtime.mode,
    anchorEmailKey: args.anchorEmailKey,
    includedEmailKeys: args.selectedEmailKeys,
    workingGroupId: args.workingGroupId,
    workingGroupName: args.workingGroupName,
    filters: {
      query: String(args.filterQuery || "").trim() || undefined,
      attachmentMode: args.attachmentMode || "all",
      groupMode: args.groupMode || "all",
    },
    attachments,
    mainLocation: args.runtime.primaryLocation,
    remotePromotionLocation: args.runtime.remotePromotionLocation,
    promotion: {
      state: "not_requested",
      promotedScopes: [],
      blockedScopes: args.runtime.attachmentPolicy.neverAutoPromoteBinary ? ["attachment_binary"] : [],
      note: buildPromotionNote(args),
    },
  });
}
