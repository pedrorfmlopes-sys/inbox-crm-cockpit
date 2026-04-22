import type { GroupTicketEntry, RelatedEmailEntry } from "@/api";
import {
  applyClassificationToIntermediateCase,
  type IntermediateCaseClassificationDraft,
} from "@/modules/crm/groups-v1/storage/intermediateCaseClassification";
import type { IntermediateCase } from "@/modules/crm/groups-v1/storage/intermediateCaseTypes";

import {
  buildResolvedIntermediateCaseClassificationDraft,
  type ResolvedStudioApplySelection,
} from "./applyResolution";
import type { ClassificationMetaDraft } from "./types";

export type ApplyLocalCaseProjectionResult = {
  nextClassificationCase: IntermediateCase;
  localClassificationDraft: IntermediateCaseClassificationDraft;
  localClassificationState: string;
};

export function projectApplyIntoIntermediateCase(args: {
  classificationCase: IntermediateCase;
  resolvedApplySelection: ResolvedStudioApplySelection;
  resolvedCaseTicket: GroupTicketEntry | null;
  targetEmails: RelatedEmailEntry[];
  classificationMetaDraft: ClassificationMetaDraft;
}): ApplyLocalCaseProjectionResult {
  const localClassificationState = String(
    (args.classificationMetaDraft.ticketStatusEnabled
      ? args.resolvedApplySelection.desiredTicketStatus || args.resolvedCaseTicket?.status
      : "")
    || (args.classificationMetaDraft.principalStatusEnabled
      ? args.resolvedApplySelection.principalGroup?.status
      : "")
    || ""
  ).trim();

  const localClassificationDraft = buildResolvedIntermediateCaseClassificationDraft({
    resolvedApplySelection: args.resolvedApplySelection,
    resolvedCaseTicket: args.resolvedCaseTicket,
    localClassificationState,
  });

  const nextClassificationCase = applyClassificationToIntermediateCase({
    caseValue: args.classificationCase,
    targetEmails: args.targetEmails,
    draft: localClassificationDraft,
  });

  return {
    nextClassificationCase,
    localClassificationDraft,
    localClassificationState,
  };
}
