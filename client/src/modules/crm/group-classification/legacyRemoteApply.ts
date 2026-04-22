import {
  addEmailToLinkGroup,
  linkEmailToGroupTicket,
  registerRelevantEmail,
  removeEmailFromLinkGroup,
  unlinkEmailFromGroupTicket,
  type GroupTicketEntry,
} from "@/api";

import type {
  RemoteApplyTargetPlan,
  ResolvedStudioApplySelection,
} from "./applyResolution";

export type LegacyRemoteApplyAttachmentStorageOptions = Record<string, unknown>;

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
    applyGroups: resolvedApplySelection.allGroupIds.length > 0,
    groupIds: resolvedApplySelection.allGroupIds,
    membershipKind: resolvedApplySelection.targetMembershipKind,
  });

  return linked.ticket || args.finalTicket;
}
