import type { GroupTicketEntry, LinkGroupEntry, RelatedEmailEntry, RelevantEmailPayload } from "@/api";
import type { IntermediateCaseClassificationDraft } from "@/modules/crm/groups-v1/storage/intermediateCaseClassification";

import type { ClassificationMetaDraft } from "./types";
import { isCurrentContextEmail, makeEmailKey, normalizeDocumentLifecycleState } from "./documentUtils";

export type ApplyCurrentContext = {
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  receivedAtIso?: string;
};

export type ResolvedStudioApplySelection = {
  targetEmails: RelatedEmailEntry[];
  targetEmailKeys: string[];
  principalGroupId: string;
  principalGroupName: string;
  principalGroup: LinkGroupEntry | null;
  referenceGroupIds: string[];
  referenceGroups: LinkGroupEntry[];
  allGroupIds: string[];
  emailLabelStatus: string;
  labels: string[];
  emailOwnedSelectedLabels: string[];
  removedInheritedLabels: string[];
  labelStates: Record<string, string>;
  categorizedLabelNames: string[];
  selectedTicketId: string;
  selectedSeriesId: string;
  selectedTicket: GroupTicketEntry | null;
  desiredTicketStatus: string;
  baseClassificationMeta: ClassificationMetaDraft;
  targetMembershipKind: "principal" | "referencia";
  hasAnyClassificationValue: boolean;
};

function normalizeString(value: unknown): string {
  return String(value || "").trim();
}

function normalizeStringList(values: unknown[]): string[] {
  return Array.from(new Set(values.map((value) => normalizeString(value)).filter(Boolean)));
}

export function buildResolvedStudioApplySelection(args: {
  targetEmails: RelatedEmailEntry[];
  principalGroupId?: string;
  principalGroup?: LinkGroupEntry | null;
  referenceGroupIds: string[];
  referenceGroups: LinkGroupEntry[];
  selectedLabels: string[];
  inheritedLabels: string[];
  selectedLabelStates: Record<string, string>;
  categorizedLabelNames: string[];
  selectedTicketId?: string;
  selectedSeriesId?: string;
  selectedTicket?: GroupTicketEntry | null;
  ticketStatusDraft?: string;
  classificationMetaDraft: ClassificationMetaDraft;
  existingSelectedEmailGroupIds?: string[];
  existingSelectedEmailTicketIds?: string[];
  existingSelectedEmailLabels?: string[];
  existingSelectedEmailStatus?: string;
}): ResolvedStudioApplySelection {
  const principalGroupId = normalizeString(args.principalGroupId);
  const referenceGroupIds = normalizeStringList(args.referenceGroupIds);
  const selectedLabels = normalizeStringList(args.selectedLabels);
  const inheritedLabels = normalizeStringList(args.inheritedLabels);
  const removedInheritedLabels = inheritedLabels.filter((label) => !selectedLabels.includes(label));
  const emailOwnedSelectedLabels = selectedLabels.filter((label) => !inheritedLabels.includes(label));
  const targetEmails = args.targetEmails.filter(Boolean);
  const targetEmailKeys = targetEmails.map((email) => makeEmailKey(email)).filter(Boolean);
  return {
    targetEmails,
    targetEmailKeys,
    principalGroupId,
    principalGroupName: normalizeString(args.principalGroup?.name || principalGroupId),
    principalGroup: args.principalGroup || null,
    referenceGroupIds,
    referenceGroups: args.referenceGroups.filter(Boolean),
    allGroupIds: normalizeStringList([principalGroupId, ...referenceGroupIds]),
    emailLabelStatus: normalizeString(Object.values(args.selectedLabelStates || {}).find(Boolean)),
    labels: selectedLabels,
    emailOwnedSelectedLabels,
    removedInheritedLabels,
    labelStates: Object.fromEntries(
      Object.entries(args.selectedLabelStates || {})
        .map(([label, status]) => [normalizeString(label), normalizeString(status)])
        .filter(([label, status]) => label && status)
    ),
    categorizedLabelNames: normalizeStringList(args.categorizedLabelNames),
    selectedTicketId: normalizeString(args.selectedTicketId),
    selectedSeriesId: normalizeString(args.selectedSeriesId),
    selectedTicket: args.selectedTicket || null,
    desiredTicketStatus: normalizeString(args.ticketStatusDraft),
    baseClassificationMeta: {
      ...args.classificationMetaDraft,
      categorizedLabelNames: normalizeStringList(args.categorizedLabelNames),
    },
    targetMembershipKind: principalGroupId ? "principal" : "referencia",
    hasAnyClassificationValue: Boolean(
      principalGroupId
      || referenceGroupIds.length
      || normalizeString(args.selectedTicketId)
      || normalizeString(args.selectedSeriesId)
      || (args.existingSelectedEmailGroupIds || []).length
      || (args.existingSelectedEmailTicketIds || []).length
      || selectedLabels.length
      || (args.existingSelectedEmailLabels || []).length
      || normalizeString(args.existingSelectedEmailStatus)
    ),
  };
}

export function buildResolvedApplyTargetPayload(args: {
  targetEmail: RelatedEmailEntry;
  currentContext: ApplyCurrentContext;
}): RelevantEmailPayload {
  const { targetEmail, currentContext } = args;
  const targetIsCurrent = isCurrentContextEmail(targetEmail, currentContext);
  const targetAttachments = (targetEmail.attachments || []).map((attachment) => ({
    key: attachment.key,
    id: attachment.id,
    name: attachment.name,
    contentType: String(attachment.contentType || "application/octet-stream"),
    content: String(attachment.content || ""),
    size: attachment.size,
    isInline: attachment.isInline,
    contentId: attachment.contentId,
    storageProvider: (attachment as any).storageProvider,
    storageBasePath: (attachment as any).storageBasePath,
    storagePathHint: (attachment as any).storagePathHint,
    documentState: normalizeDocumentLifecycleState((attachment as any)?.documentState, "ingested"),
    hasContent: (attachment as any)?.hasContent === true || Boolean(String(attachment.content || "").trim()),
    isHidden: typeof (attachment as any)?.isHidden === "boolean" ? (attachment as any).isHidden : undefined,
  }));
  return {
    itemId: normalizeString(targetEmail?.itemId || (targetIsCurrent ? currentContext.itemId : "")) || undefined,
    internetMessageId: normalizeString(targetEmail?.internetMessageId || (targetIsCurrent ? currentContext.internetMessageId : "")) || undefined,
    conversationId: normalizeString(targetEmail?.conversationId || (targetIsCurrent ? currentContext.conversationId : "")) || undefined,
    subject: normalizeString(targetEmail?.subject || (targetIsCurrent ? currentContext.subject : "")) || undefined,
    fromEmail: normalizeString(targetEmail?.fromEmail || (targetIsCurrent ? currentContext.fromEmail : "")) || undefined,
    fromName: normalizeString(targetEmail?.fromName || (targetIsCurrent ? currentContext.fromName : "")) || undefined,
    receivedAtIso: normalizeString(targetEmail?.receivedAtIso || targetEmail?.messageDateIso || (targetIsCurrent ? currentContext.receivedAtIso : "")) || undefined,
    messageDateIso: normalizeString(targetEmail?.messageDateIso || targetEmail?.receivedAtIso || (targetIsCurrent ? currentContext.receivedAtIso : "")) || undefined,
    bodyText: normalizeString(targetEmail?.bodyText) || undefined,
    bodyHtml: normalizeString(targetEmail?.bodyHtml) || undefined,
    attachments: targetAttachments.map((attachment) => ({
      key: attachment.key,
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId,
      content: attachment.content,
      storageProvider: (attachment as any).storageProvider,
      storageBasePath: (attachment as any).storageBasePath,
      storagePathHint: (attachment as any).storagePathHint,
      documentState: (attachment as any).documentState,
      hasContent: (attachment as any).hasContent === true || Boolean(String(attachment.content || "").trim()),
      isHidden: typeof (attachment as any)?.isHidden === "boolean" ? (attachment as any).isHidden : undefined,
    })),
  };
}

export function buildResolvedClassifiedEmailPayload(args: {
  targetEmail: RelatedEmailEntry;
  currentContext: ApplyCurrentContext;
  resolvedApplySelection: ResolvedStudioApplySelection;
}): RelevantEmailPayload {
  const basePayload = buildResolvedApplyTargetPayload({
    targetEmail: args.targetEmail,
    currentContext: args.currentContext,
  });
  return {
    ...basePayload,
    status: args.resolvedApplySelection.emailLabelStatus,
    labels: args.resolvedApplySelection.emailOwnedSelectedLabels,
    removedInheritedLabels: args.resolvedApplySelection.removedInheritedLabels,
    labelStates: args.resolvedApplySelection.labelStates,
    classificationMeta: args.resolvedApplySelection.baseClassificationMeta,
  };
}

export function buildResolvedIntermediateCaseClassificationDraft(args: {
  resolvedApplySelection: ResolvedStudioApplySelection;
  resolvedCaseTicket?: GroupTicketEntry | null;
  localClassificationState?: string;
}): IntermediateCaseClassificationDraft {
  const resolvedCaseTicket = args.resolvedCaseTicket || null;
  return {
    principalGroupId: args.resolvedApplySelection.principalGroupId || undefined,
    principalGroupName: args.resolvedApplySelection.principalGroupName || undefined,
    referenceGroupIds: args.resolvedApplySelection.referenceGroupIds,
    labels: args.resolvedApplySelection.labels,
    removedInheritedLabels: args.resolvedApplySelection.removedInheritedLabels,
    labelStates: args.resolvedApplySelection.labelStates,
    categorizedLabelNames: args.resolvedApplySelection.categorizedLabelNames,
    ticketIds: [normalizeString(resolvedCaseTicket?.id || args.resolvedApplySelection.selectedTicketId)].filter(Boolean),
    ticketCodes: [normalizeString(resolvedCaseTicket?.code)].filter(Boolean),
    state: normalizeString(args.localClassificationState) || undefined,
    status: args.resolvedApplySelection.emailLabelStatus || undefined,
    classifiedAt: new Date().toISOString(),
    classifiedSource: "user",
  };
}
