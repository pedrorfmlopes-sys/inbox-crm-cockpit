import { buildIntermediateCaseFromSeed, touchIntermediateCaseAccess } from "./intermediateCaseNormalization";
import { mergeEmailIntoIntermediateCase } from "./intermediateCaseMerge";
import type {
  IntermediateCase,
  IntermediateCaseAttachment,
  IntermediateCaseEmail,
  IntermediateEmailClassification,
  IntermediateLocalPresence,
  IntermediateServerPresence,
  IntermediateVisibleState,
} from "./intermediateCaseTypes";

export type PrepareIntermediateAttachmentInput = {
  attachmentKey: string;
  id?: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
  hasContent?: boolean;
  documentState?: string;
  previewReady?: boolean;
  selected?: boolean;
};

export type PrepareIntermediateEmailInput = {
  emailKey: string;
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromName?: string;
  fromEmail?: string;
  to?: string[];
  cc?: string[];
  receivedAtIso?: string;
  bodyText?: string;
  bodyHtml?: string;
  sourceOrigin: IntermediateCaseEmail["sourceOrigin"];
  visibilityState: IntermediateVisibleState;
  serverPresence: IntermediateServerPresence;
  localPresence: IntermediateLocalPresence;
  classification?: Partial<IntermediateEmailClassification>;
  attachments?: PrepareIntermediateAttachmentInput[];
};

type BuildPrepareIntermediateCaseArgs = {
  caseId: string;
  anchorEmailKey: string;
  conversationId?: string;
  existingCase?: IntermediateCase | null;
  emails: PrepareIntermediateEmailInput[];
  nowIso?: string;
};

function normalizeString(value: unknown): string | undefined {
  const normalized = String(value || "").trim();
  return normalized || undefined;
}

function normalizeAttachment(input: PrepareIntermediateAttachmentInput): IntermediateCaseAttachment {
  return {
    attachmentKey: input.attachmentKey,
    id: normalizeString(input.id),
    name: String(input.name || "").trim(),
    contentType: normalizeString(input.contentType),
    size: Number.isFinite(Number(input.size)) ? Number(input.size) : undefined,
    isInline: input.isInline === true,
    contentId: normalizeString(input.contentId),
    hasContent: input.hasContent === true,
    documentState: normalizeString(input.documentState),
    storageDecision: input.selected ? (input.hasContent === true ? "local" : "metadata_only") : "pending",
    localRef: input.selected
      ? {
          kind: "relative_path",
          value: `attachments/${input.attachmentKey}`,
          label: String(input.name || "").trim() || undefined,
        }
      : undefined,
    previewReady: input.previewReady === true || (input.selected === true && input.hasContent === true),
  };
}

function normalizeEmail(input: PrepareIntermediateEmailInput): IntermediateCaseEmail {
  return {
    emailKey: input.emailKey,
    itemId: normalizeString(input.itemId),
    internetMessageId: normalizeString(input.internetMessageId),
    conversationId: normalizeString(input.conversationId),
    subject: normalizeString(input.subject),
    fromName: normalizeString(input.fromName),
    fromEmail: normalizeString(input.fromEmail),
    to: Array.isArray(input.to) ? input.to.filter(Boolean) : [],
    cc: Array.isArray(input.cc) ? input.cc.filter(Boolean) : [],
    receivedAtIso: normalizeString(input.receivedAtIso),
    bodyText: normalizeString(input.bodyText),
    bodyHtml: normalizeString(input.bodyHtml),
    sourceOrigin: input.sourceOrigin,
    visibilityState: input.visibilityState,
    serverPresence: input.serverPresence,
    localPresence: input.localPresence,
    classification: {
      principalGroupId: normalizeString(input.classification?.principalGroupId),
      principalGroupName: normalizeString(input.classification?.principalGroupName),
      referenceGroupIds: Array.isArray(input.classification?.referenceGroupIds)
        ? input.classification.referenceGroupIds.filter(Boolean)
        : [],
      labels: Array.isArray(input.classification?.labels) ? input.classification.labels.filter(Boolean) : [],
      ticketIds: Array.isArray(input.classification?.ticketIds) ? input.classification.ticketIds.filter(Boolean) : [],
      ticketCodes: Array.isArray(input.classification?.ticketCodes) ? input.classification.ticketCodes.filter(Boolean) : [],
      state: normalizeString(input.classification?.state),
      status: normalizeString(input.classification?.status),
      classifiedAt: normalizeString(input.classification?.classifiedAt),
      classifiedSource: input.classification?.classifiedSource,
    },
    attachments: Array.isArray(input.attachments) ? input.attachments.map((attachment) => normalizeAttachment(attachment)) : [],
  };
}

export function buildPrepareIntermediateCase(args: BuildPrepareIntermediateCaseArgs): IntermediateCase {
  const nowIso = String(args.nowIso || new Date().toISOString()).trim();
  const existingCase = args.existingCase && args.existingCase.caseId === args.caseId ? args.existingCase : null;
  const currentEmailKeys = new Set(args.emails.map((email) => email.emailKey).filter(Boolean));

  let caseValue = buildIntermediateCaseFromSeed({
    caseId: args.caseId,
    anchorEmailKey: args.anchorEmailKey,
    conversationId: args.conversationId,
    createdAt: existingCase?.createdAt || nowIso,
    updatedAt: nowIso,
    lastAccessedAt: nowIso,
    emails: (existingCase?.emails || [])
      .filter((email) => currentEmailKeys.has(email.emailKey))
      .map((email) => ({
        ...email,
        attachments: email.attachments,
        classification: email.classification,
      })),
  });

  for (const email of args.emails) {
    caseValue = mergeEmailIntoIntermediateCase(caseValue, normalizeEmail(email));
  }

  return touchIntermediateCaseAccess(caseValue, nowIso);
}
