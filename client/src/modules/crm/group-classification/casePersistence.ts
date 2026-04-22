import { resolveClassificationIntermediateCase } from "@/modules/crm/groups-v1/storage/resolveClassificationIntermediateCase";
import type { IntermediateCase } from "@/modules/crm/groups-v1/storage/intermediateCaseTypes";

type SyncClassificationCaseEmails = (
  nextCaseValue: IntermediateCase,
  options?: {
    preferredSelectedEmailKey?: string;
    preferredTargetEmailKeys?: string[];
    rehydrateSelectedEmail?: boolean;
  }
) => void;

export async function persistAndRefreshClassificationCase<TRefreshedContext>(args: {
  classificationCase: IntermediateCase;
  nextClassificationCase: IntermediateCase;
  preferredSelectedEmailKey?: string;
  preferredTargetEmailKeys?: string[];
  syncClassificationCaseEmails: SyncClassificationCaseEmails;
  refreshSelectedEmailContext: () => Promise<TRefreshedContext | null>;
}): Promise<{
  appliedClassificationCase: IntermediateCase;
  refreshedContext: TRefreshedContext | null;
}> {
  const classificationStorage = await resolveClassificationIntermediateCase({
    caseId: args.classificationCase.caseId,
    anchorEmailKey: args.classificationCase.anchorEmailKey,
  });

  await classificationStorage.storage.repository.writeCase(args.nextClassificationCase);

  args.syncClassificationCaseEmails(args.nextClassificationCase, {
    preferredSelectedEmailKey: args.preferredSelectedEmailKey,
    preferredTargetEmailKeys: args.preferredTargetEmailKeys,
  });

  const refreshedContext = await args.refreshSelectedEmailContext().catch(() => null);

  args.syncClassificationCaseEmails(args.nextClassificationCase, {
    preferredSelectedEmailKey: args.preferredSelectedEmailKey,
    preferredTargetEmailKeys: args.preferredTargetEmailKeys,
  });

  return {
    appliedClassificationCase: args.nextClassificationCase,
    refreshedContext,
  };
}
