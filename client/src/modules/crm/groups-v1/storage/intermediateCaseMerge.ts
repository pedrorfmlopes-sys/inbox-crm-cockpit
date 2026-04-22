import type {
  IntermediateCase,
  IntermediateCaseAttachment,
  IntermediateCaseEmail,
} from "./intermediateCaseTypes";
import { normalizeIntermediateCase } from "./intermediateCaseNormalization";

function pickString(current?: string, incoming?: string): string | undefined {
  return String(incoming || "").trim() || String(current || "").trim() || undefined;
}

function pickNumber(current?: number, incoming?: number): number | undefined {
  return Number.isFinite(Number(incoming)) ? Number(incoming) : current;
}

function mergeAttachment(
  current: IntermediateCaseAttachment | undefined,
  incoming: IntermediateCaseAttachment
): IntermediateCaseAttachment {
  if (!current) return incoming;
  return {
    attachmentKey: incoming.attachmentKey,
    id: pickString(current.id, incoming.id),
    name: pickString(current.name, incoming.name) || incoming.name,
    contentType: pickString(current.contentType, incoming.contentType),
    size: pickNumber(current.size, incoming.size),
    isInline: incoming.isInline === true || current.isInline === true,
    contentId: pickString(current.contentId, incoming.contentId),
    hasContent: incoming.hasContent === true || current.hasContent === true,
    documentState: pickString(current.documentState, incoming.documentState),
    isHidden: typeof incoming.isHidden === "boolean" ? incoming.isHidden : current.isHidden,
    storageDecision: incoming.storageDecision || current.storageDecision,
    localRef: incoming.localRef || current.localRef,
    serverRef: incoming.serverRef || current.serverRef,
    previewReady: incoming.previewReady === true || current.previewReady === true,
  };
}

function mergeEmail(current: IntermediateCaseEmail | undefined, incoming: IntermediateCaseEmail): IntermediateCaseEmail {
  if (!current) return incoming;
  const attachments = new Map<string, IntermediateCaseAttachment>();
  for (const attachment of current.attachments) attachments.set(attachment.attachmentKey, attachment);
  for (const attachment of incoming.attachments) {
    attachments.set(attachment.attachmentKey, mergeAttachment(attachments.get(attachment.attachmentKey), attachment));
  }
  return {
    emailKey: incoming.emailKey,
    itemId: pickString(current.itemId, incoming.itemId),
    internetMessageId: pickString(current.internetMessageId, incoming.internetMessageId),
    conversationId: pickString(current.conversationId, incoming.conversationId),
    subject: pickString(current.subject, incoming.subject),
    fromName: pickString(current.fromName, incoming.fromName),
    fromEmail: pickString(current.fromEmail, incoming.fromEmail),
    to: Array.from(new Set([...(current.to || []), ...(incoming.to || [])])),
    cc: Array.from(new Set([...(current.cc || []), ...(incoming.cc || [])])),
    receivedAtIso: pickString(current.receivedAtIso, incoming.receivedAtIso),
    bodyText: pickString(current.bodyText, incoming.bodyText),
    bodyHtml: pickString(current.bodyHtml, incoming.bodyHtml),
    sourceOrigin: incoming.sourceOrigin || current.sourceOrigin,
    visibilityState: incoming.visibilityState || current.visibilityState,
    serverPresence: incoming.serverPresence !== "none" ? incoming.serverPresence : current.serverPresence,
    localPresence: incoming.localPresence !== "none" ? incoming.localPresence : current.localPresence,
    classification: {
      principalGroupId: pickString(current.classification.principalGroupId, incoming.classification.principalGroupId),
      principalGroupName: pickString(current.classification.principalGroupName, incoming.classification.principalGroupName),
      referenceGroupIds: Array.from(new Set([
        ...(current.classification.referenceGroupIds || []),
        ...(incoming.classification.referenceGroupIds || []),
      ])),
      labels: Array.from(new Set([...(current.classification.labels || []), ...(incoming.classification.labels || [])])),
      removedInheritedLabels: Array.from(new Set([
        ...(current.classification.removedInheritedLabels || []),
        ...(incoming.classification.removedInheritedLabels || []),
      ])),
      labelStates: {
        ...(current.classification.labelStates || {}),
        ...(incoming.classification.labelStates || {}),
      },
      categorizedLabelNames: Array.from(new Set([
        ...(current.classification.categorizedLabelNames || []),
        ...(incoming.classification.categorizedLabelNames || []),
      ])),
      ticketIds: Array.from(new Set([...(current.classification.ticketIds || []), ...(incoming.classification.ticketIds || [])])),
      ticketCodes: Array.from(new Set([...(current.classification.ticketCodes || []), ...(incoming.classification.ticketCodes || [])])),
      state: pickString(current.classification.state, incoming.classification.state),
      status: pickString(current.classification.status, incoming.classification.status),
      classifiedAt: pickString(current.classification.classifiedAt, incoming.classification.classifiedAt),
      classifiedSource: incoming.classification.classifiedSource || current.classification.classifiedSource,
    },
    attachments: Array.from(attachments.values()),
  };
}

export function mergeEmailIntoIntermediateCase(caseValue: IntermediateCase, incoming: IntermediateCaseEmail): IntermediateCase {
  const emails = new Map(caseValue.emails.map((email) => [email.emailKey, email]));
  emails.set(incoming.emailKey, mergeEmail(emails.get(incoming.emailKey), incoming));
  return normalizeIntermediateCase({
    ...caseValue,
    updatedAt: new Date().toISOString(),
    emails: Array.from(emails.values()),
  }) || caseValue;
}

export function mergeAttachmentsIntoIntermediateCase(
  caseValue: IntermediateCase,
  emailKey: string,
  incoming: IntermediateCaseAttachment[]
): IntermediateCase {
  const target = caseValue.emails.find((email) => email.emailKey === emailKey);
  if (!target) return caseValue;
  return mergeEmailIntoIntermediateCase(caseValue, {
    ...target,
    attachments: incoming,
  });
}
