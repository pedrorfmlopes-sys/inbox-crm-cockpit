import type { RelatedEmailEntry } from "@/api";
import { normalizeStudioAttachment, makeEmailKey } from "@/modules/crm/group-classification/documentUtils";
import { normalizeIntermediateCase } from "./intermediateCaseNormalization";
import type {
  IntermediateCase,
  IntermediateCaseAttachment,
  IntermediateCaseEmail,
  IntermediateClassificationSource,
  IntermediateEmailClassification,
  IntermediateLocalPresence,
  IntermediateServerPresence,
  IntermediateVisibleState,
} from "./intermediateCaseTypes";

type StudioAttachmentLike = {
  key?: string;
  id?: string;
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
  hasContent?: boolean;
  content?: string;
  documentState?: string;
  isHidden?: boolean;
};

export type IntermediateCaseClassificationDraft = {
  principalGroupId?: string;
  principalGroupName?: string;
  referenceGroupIds: string[];
  labels: string[];
  ticketIds: string[];
  ticketCodes: string[];
  state?: string;
  status?: string;
  classifiedAt: string;
  classifiedSource: IntermediateClassificationSource;
};

function normalizeString(value: unknown): string {
  return String(value || "").trim();
}

function normalizeStringArray(values: unknown): string[] {
  if (!Array.isArray(values)) return [];
  return Array.from(new Set(values.map((value) => normalizeString(value)).filter(Boolean)));
}

function buildClassificationFromDraft(
  draft: IntermediateCaseClassificationDraft
): IntermediateEmailClassification {
  return {
    principalGroupId: normalizeString(draft.principalGroupId) || undefined,
    principalGroupName: normalizeString(draft.principalGroupName) || undefined,
    referenceGroupIds: normalizeStringArray(draft.referenceGroupIds),
    labels: normalizeStringArray(draft.labels),
    ticketIds: normalizeStringArray(draft.ticketIds),
    ticketCodes: normalizeStringArray(draft.ticketCodes),
    state: normalizeString(draft.state) || undefined,
    status: normalizeString(draft.status) || undefined,
    classifiedAt: normalizeString(draft.classifiedAt) || undefined,
    classifiedSource: draft.classifiedSource,
  };
}

function buildUiAttachmentKey(emailKey: string, attachment: { key?: string; id?: string; name?: string }): string {
  const bareKey = normalizeString(attachment.key) || normalizeString(attachment.id) || normalizeString(attachment.name);
  if (!bareKey) return "";
  return `${emailKey}:${bareKey}`;
}

function resolveAttachmentIdentity(
  emailKey: string,
  attachment: { key?: string; id?: string; name?: string },
  currentAttachments: IntermediateCaseAttachment[]
): string {
  const uiKey = buildUiAttachmentKey(emailKey, attachment);
  if (!uiKey) return "";
  const exact = currentAttachments.find((entry) => entry.attachmentKey === uiKey);
  if (exact) return exact.attachmentKey;
  const bareKey = normalizeString(attachment.key) || normalizeString(attachment.id) || normalizeString(attachment.name);
  const suffixMatch = currentAttachments.find((entry) => entry.attachmentKey === bareKey || entry.attachmentKey.endsWith(`:${bareKey}`));
  return suffixMatch?.attachmentKey || uiKey;
}

function mapRelatedEmailEntryToIntermediateEmailBase(
  email: RelatedEmailEntry,
  current?: IntermediateCaseEmail | null
): IntermediateCaseEmail {
  const emailKey = makeEmailKey(email);
  const currentAttachments = current?.attachments || [];
  const attachments = Array.isArray(email.attachments)
    ? email.attachments
        .map((attachment) => normalizeStudioAttachment(attachment))
        .filter(Boolean)
        .map((attachment) => {
          const normalizedAttachment = attachment as StudioAttachmentLike;
          const attachmentKey = resolveAttachmentIdentity(emailKey, normalizedAttachment, currentAttachments);
          const existing = currentAttachments.find((entry) => entry.attachmentKey === attachmentKey) || null;
          return {
            attachmentKey,
            id: normalizeString(normalizedAttachment.id) || existing?.id,
            name: normalizeString(normalizedAttachment.name) || existing?.name || attachmentKey,
            contentType: normalizeString(normalizedAttachment.contentType) || existing?.contentType,
            size: Number.isFinite(Number(normalizedAttachment.size)) ? Number(normalizedAttachment.size) : existing?.size,
            isInline: normalizedAttachment.isInline === true || existing?.isInline === true,
            contentId: normalizeString(normalizedAttachment.contentId) || existing?.contentId,
            hasContent: normalizedAttachment.hasContent === true || Boolean(normalizeString(normalizedAttachment.content)) || existing?.hasContent === true,
            documentState: normalizeString(normalizedAttachment.documentState) || existing?.documentState,
            isHidden: typeof normalizedAttachment.isHidden === "boolean" ? normalizedAttachment.isHidden : existing?.isHidden,
            storageDecision: existing?.storageDecision || "pending",
            localRef: existing?.localRef,
            serverRef: existing?.serverRef,
            previewReady: normalizedAttachment.hasContent === true || Boolean(normalizeString(normalizedAttachment.content)) || existing?.previewReady === true,
          } satisfies IntermediateCaseAttachment;
        })
    : currentAttachments;

  return {
    emailKey,
    itemId: normalizeString(email.itemId) || current?.itemId,
    internetMessageId: normalizeString(email.internetMessageId) || current?.internetMessageId,
    conversationId: normalizeString(email.conversationId) || current?.conversationId,
    subject: normalizeString(email.subject) || current?.subject,
    fromName: normalizeString(email.fromName) || current?.fromName,
    fromEmail: normalizeString(email.fromEmail) || current?.fromEmail,
    to: current?.to || [],
    cc: current?.cc || [],
    receivedAtIso: normalizeString(email.messageDateIso || email.receivedAtIso) || current?.receivedAtIso,
    bodyText: normalizeString(email.bodyText) || current?.bodyText,
    bodyHtml: normalizeString(email.bodyHtml) || current?.bodyHtml,
    sourceOrigin: current?.sourceOrigin || "intermediate",
    visibilityState: current?.visibilityState || "draft",
    serverPresence: current?.serverPresence || "none",
    localPresence: current?.localPresence || "case_only",
    classification: current?.classification || {
      referenceGroupIds: [],
      labels: [],
      ticketIds: [],
      ticketCodes: [],
    },
    attachments,
  };
}

function resolveVisibilityState(current?: IntermediateCaseEmail | null): IntermediateVisibleState {
  const serverPresence = current?.serverPresence || "none";
  return serverPresence === "classified" || serverPresence === "attachments" || serverPresence === "complete"
    ? "server"
    : "local";
}

function resolveLocalPresence(current?: IntermediateCaseEmail | null): IntermediateLocalPresence {
  if (!current) return "complete";
  if (current.localPresence === "attachments") return "attachments";
  return "complete";
}

function resolveServerPresence(current?: IntermediateCaseEmail | null): IntermediateServerPresence {
  return current?.serverPresence || "none";
}

export function applyClassificationToIntermediateCase(args: {
  caseValue: IntermediateCase;
  targetEmails: RelatedEmailEntry[];
  draft: IntermediateCaseClassificationDraft;
}): IntermediateCase {
  const classification = buildClassificationFromDraft(args.draft);
  const targetEmailKeys = new Set(args.targetEmails.map((email) => makeEmailKey(email)).filter(Boolean));
  const currentByKey = new Map(args.caseValue.emails.map((email) => [email.emailKey, email]));
  const incomingByKey = new Map(args.targetEmails.map((email) => [makeEmailKey(email), email] as const));

  const nextEmails = args.caseValue.emails.map((email) => {
    if (!targetEmailKeys.has(email.emailKey)) return email;
    const incoming = incomingByKey.get(email.emailKey);
    if (!incoming) return email;
    const base = mapRelatedEmailEntryToIntermediateEmailBase(incoming, email);
    return {
      ...base,
      visibilityState: resolveVisibilityState(email),
      localPresence: resolveLocalPresence(email),
      serverPresence: resolveServerPresence(email),
      classification,
    } satisfies IntermediateCaseEmail;
  });

  for (const incoming of args.targetEmails) {
    const emailKey = makeEmailKey(incoming);
    if (!emailKey || currentByKey.has(emailKey)) continue;
    const base = mapRelatedEmailEntryToIntermediateEmailBase(incoming, null);
    nextEmails.push({
      ...base,
      sourceOrigin: "intermediate",
      visibilityState: "local",
      localPresence: "complete",
      serverPresence: "none",
      classification,
    } satisfies IntermediateCaseEmail);
  }

  return normalizeIntermediateCase({
    ...args.caseValue,
    updatedAt: new Date().toISOString(),
    lastAccessedAt: new Date().toISOString(),
    emails: nextEmails,
  }) || args.caseValue;
}
