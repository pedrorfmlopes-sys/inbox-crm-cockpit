import type { RelatedEmailEntry } from "@/api";
import {
  supportsIntermediateCaseBinaryStorage,
  type IntermediateCaseStorageAdapter,
} from "./intermediateCaseRepository";
import type { IntermediateCase, IntermediateCaseAttachment, IntermediateCaseEmail } from "./intermediateCaseTypes";

function makeAttachmentKeyForUi(email: IntermediateCaseEmail, attachment: IntermediateCaseAttachment): string {
  if (attachment.attachmentKey.startsWith(`${email.emailKey}:`)) {
    return attachment.attachmentKey.slice(email.emailKey.length + 1);
  }
  return attachment.attachmentKey;
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => {
      const result = String(reader.result || "");
      const base64 = result.includes(",") ? result.slice(result.indexOf(",") + 1) : result;
      resolve(base64);
    };
    reader.onerror = () => reject(reader.error || new Error("Nao foi possivel ler o blob do anexo."));
    reader.readAsDataURL(blob);
  });
}

export async function readIntermediateCaseAttachmentContentMap(args: {
  caseValue: IntermediateCase | null | undefined;
  adapter: IntermediateCaseStorageAdapter;
}): Promise<Map<string, string>> {
  const contentByPath = new Map<string, string>();
  if (!args.caseValue || !supportsIntermediateCaseBinaryStorage(args.adapter)) {
    return contentByPath;
  }

  for (const email of args.caseValue.emails) {
    for (const attachment of email.attachments) {
      const localPath = String(attachment.localRef?.value || "").trim();
      if (!localPath || !attachment.hasContent) continue;
      try {
        const blob = await args.adapter.readBinary(localPath);
        if (!blob) continue;
        const base64 = await blobToBase64(blob);
        if (base64) contentByPath.set(localPath, base64);
      } catch {
        // best effort: keep attachment metadata even if binary hydration fails
      }
    }
  }

  return contentByPath;
}

export function mapIntermediateEmailToRelatedEmailEntry(
  email: IntermediateCaseEmail,
  options?: {
    attachmentContentByPath?: Map<string, string>;
  }
): RelatedEmailEntry {
  const attachmentContentByPath = options?.attachmentContentByPath || new Map<string, string>();
  const principalGroupId = String(email.classification.principalGroupId || "").trim();
  const principalGroupName = String(email.classification.principalGroupName || principalGroupId).trim();
  const referenceGroups: Array<{ id: string; name: string; relationKind: "reference" }> = email.classification.referenceGroupIds.map((groupId) => ({
    id: String(groupId || "").trim(),
    name: String(groupId || "").trim(),
    relationKind: "reference" as const,
  })).filter((group) => group.id);

  return {
    emailKey: email.emailKey,
    itemId: email.itemId,
    internetMessageId: email.internetMessageId,
    conversationId: email.conversationId,
    subject: email.subject,
    fromName: email.fromName,
    fromEmail: email.fromEmail,
    toRecipients: email.to.map((entry) => ({ email: entry })),
    ccRecipients: email.cc.map((entry) => ({ email: entry })),
    receivedAtIso: email.receivedAtIso,
    messageDateIso: email.receivedAtIso,
    bodyText: email.bodyText,
    bodyHtml: email.bodyHtml,
    membershipKind: principalGroupId ? "principal" : undefined,
    groupId: principalGroupId || undefined,
    groupName: principalGroupId ? principalGroupName || principalGroupId : undefined,
    status: email.classification.status,
    labels: email.classification.labels,
    removedInheritedLabels: email.classification.removedInheritedLabels,
    labelStates: email.classification.labelStates,
    relatedGroups: referenceGroups,
    relatedReasons: [],
    classificationMeta: {
      principalGroupId: principalGroupId || "",
      referenceGroupIds: email.classification.referenceGroupIds,
      ticketId: email.classification.ticketIds[0] || "",
      categorizedLabelNames: email.classification.categorizedLabelNames,
    },
    attachments: email.attachments.map((attachment) => {
      const localPath = String(attachment.localRef?.value || "").trim();
      return {
        key: makeAttachmentKeyForUi(email, attachment),
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        contentId: attachment.contentId,
        hasContent: attachment.hasContent,
        documentState: attachment.documentState,
        isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : undefined,
        storagePathHint: localPath || undefined,
        content: localPath ? attachmentContentByPath.get(localPath) || undefined : undefined,
      };
    }),
  } as RelatedEmailEntry;
}

export async function hydrateIntermediateCaseEmailsToRelatedEntries(args: {
  caseValue: IntermediateCase | null | undefined;
  adapter: IntermediateCaseStorageAdapter;
}): Promise<RelatedEmailEntry[]> {
  if (!args.caseValue) return [];
  const attachmentContentByPath = await readIntermediateCaseAttachmentContentMap(args);
  return args.caseValue.emails.map((email) =>
    mapIntermediateEmailToRelatedEmailEntry(email, { attachmentContentByPath })
  );
}
