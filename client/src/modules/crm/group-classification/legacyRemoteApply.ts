import {
  addEmailToLinkGroup,
  createGroupTicket,
  linkEmailToGroupTicket,
  registerRelevantEmail,
  removeEmailFromLinkGroup,
  unlinkEmailFromGroupTicket,
  updateGroupTicket,
  type GroupTicketEntry,
  type RelevantEmailPayload,
} from "@/api";

import type {
  ApplyCurrentContext,
  RemoteApplyTargetPlan,
  ResolvedRemoteApplyExecutionPlan,
  ResolvedStudioApplySelection,
} from "./applyResolution";
import { buildResolvedClassifiedEmailPayload } from "./applyResolution";

export type LegacyRemoteApplyAttachmentStorageOptions = Pick<
  RelevantEmailPayload,
  "attachmentStorageProvider" | "attachmentStorageBasePath"
>;

export type ExecuteLegacyBaseTicketApplyResult = {
  finalTicket: GroupTicketEntry | null;
  createdTicket: boolean;
  updatedTicketStatus: boolean;
};

export async function executeLegacyBaseTicketApply(args: {
  remoteApplyPlan: ResolvedRemoteApplyExecutionPlan;
  resolvedApplySelection: ResolvedStudioApplySelection;
  currentContext: ApplyCurrentContext;
  createTicketTitle?: string;
  currentOutlookTicket: GroupTicketEntry | null;
  attachmentStorageOptions?: LegacyRemoteApplyAttachmentStorageOptions;
}): Promise<ExecuteLegacyBaseTicketApplyResult> {
  const {
    remoteApplyPlan,
    resolvedApplySelection,
    currentContext,
    createTicketTitle,
    currentOutlookTicket,
    attachmentStorageOptions,
  } = args;

  const desiredTicketStatus = resolvedApplySelection.desiredTicketStatus;
  let finalTicket: GroupTicketEntry | null = null;
  let createdTicket = false;
  let updatedTicketStatus = false;

  if (remoteApplyPlan.shouldCreateTicket) {
    const baseClassifiedEmailPayload = remoteApplyPlan.targetPlans[0]?.classifiedEmailPayload
      || buildResolvedClassifiedEmailPayload({
        targetEmail: remoteApplyPlan.baseTargetEmail,
        currentContext,
        resolvedApplySelection,
      });
    finalTicket = await createGroupTicket({
      seriesId: resolvedApplySelection.selectedSeriesId,
      title: String(createTicketTitle || remoteApplyPlan.baseTargetEmail?.subject || "Ticket").trim(),
      description: String(remoteApplyPlan.baseTargetEmail?.bodyText || "").trim().slice(0, 4000),
      labels: resolvedApplySelection.labels,
      groupIds: resolvedApplySelection.allGroupIds,
      email: {
        ...baseClassifiedEmailPayload,
        ...(attachmentStorageOptions || {}),
      },
      membershipKind: resolvedApplySelection.targetMembershipKind,
    });
    createdTicket = true;

    if (desiredTicketStatus && desiredTicketStatus !== String(finalTicket?.status || "").trim()) {
      finalTicket = await updateGroupTicket(finalTicket.id, { status: desiredTicketStatus });
      updatedTicketStatus = true;
    }
  } else if (
    remoteApplyPlan.shouldUpdateTicketStatus
    && desiredTicketStatus !== String(currentOutlookTicket?.status || "").trim()
  ) {
    finalTicket = await updateGroupTicket(resolvedApplySelection.selectedTicketId, { status: desiredTicketStatus });
    updatedTicketStatus = true;
  }

  return {
    finalTicket,
    createdTicket,
    updatedTicketStatus,
  };
}

export async function executeLegacyRemoteApplyForTarget(args: {
  targetPlan: RemoteApplyTargetPlan;
  resolvedApplySelection: ResolvedStudioApplySelection;
  finalTicket: GroupTicketEntry | null;
  attachmentStorageOptions?: LegacyRemoteApplyAttachmentStorageOptions;
  skipTicketLink: boolean;
}): Promise<GroupTicketEntry | null> {
  const {
    targetPlan,
    resolvedApplySelection,
    attachmentStorageOptions,
    skipTicketLink,
  } = args;

  const { targetEmail, targetEmailPayload, classifiedEmailPayload } = targetPlan;
  const emailKey = String(targetEmail?.emailKey || "").trim() || undefined;
  const ticketIdsToRemove = targetPlan.ticketIdsToRemove.filter(
    (ticketId) => ticketId !== args.finalTicket?.id
  );

  for (const groupId of targetPlan.groupsToRemove) {
    await removeEmailFromLinkGroup(groupId, {
      ...targetEmailPayload,
      emailKey,
    });
  }

  if (resolvedApplySelection.principalGroupId) {
    await addEmailToLinkGroup(resolvedApplySelection.principalGroupId, {
      ...classifiedEmailPayload,
      membershipKind: "principal",
    });
  }

  for (const groupId of resolvedApplySelection.referenceGroupIds) {
    await addEmailToLinkGroup(groupId, {
      ...classifiedEmailPayload,
      membershipKind: "referencia",
    });
  }

  for (const ticketId of ticketIdsToRemove) {
    await unlinkEmailFromGroupTicket(ticketId, {
      email: targetEmailPayload,
      emailKey,
    });
  }

  await registerRelevantEmail({
    ...classifiedEmailPayload,
    ...(attachmentStorageOptions || {}),
  });

  const targetTicketId = args.finalTicket?.id || resolvedApplySelection.selectedTicketId;
  if (!targetTicketId || skipTicketLink) {
    return args.finalTicket;
  }

  const linked = await linkEmailToGroupTicket(targetTicketId, {
    email: classifiedEmailPayload,
    // Group memberships are already written explicitly above as principal/referencia.
    // Ticket linking in the final store should only attach the email to the ticket and
    // expand ticket.groupIds when needed, without reclassifying memberships.
    applyGroups: false,
    groupIds: resolvedApplySelection.allGroupIds,
    membershipKind: resolvedApplySelection.targetMembershipKind,
  });

  return linked.ticket || args.finalTicket;
}
