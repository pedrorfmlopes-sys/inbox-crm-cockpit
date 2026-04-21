import { buildIntermediateCaseFromSeed, touchIntermediateCaseAccess } from "./intermediateCaseNormalization";
import { mergeEmailIntoIntermediateCase } from "./intermediateCaseMerge";
import type {
  IntermediateCase,
  IntermediateCaseSourceOrigin,
} from "./intermediateCaseTypes";
import {
  applyPrepareIntermediateCaseSourcePriority,
  prepareIntermediateEmailToCaseEmail,
  type PrepareIntermediateEmailInput,
} from "./prepareIntermediateCase";

export type PrepareIntermediateCasePrimarySource = IntermediateCaseSourceOrigin;

type BuildPrepareIntermediateCaseFromSourcesArgs = {
  caseId: string;
  anchorEmailKey: string;
  conversationId?: string;
  existingCase?: IntermediateCase | null;
  outlookEmails: PrepareIntermediateEmailInput[];
  serverEmails: PrepareIntermediateEmailInput[];
  nowIso?: string;
};

export function resolvePrepareIntermediateCasePrimarySource(args: {
  serverEmails: PrepareIntermediateEmailInput[];
  existingCase?: IntermediateCase | null;
}): PrepareIntermediateCasePrimarySource {
  if (args.serverEmails.length) return "server";
  if (Array.isArray(args.existingCase?.emails) && args.existingCase!.emails.length) return "intermediate";
  return "outlook";
}

export function buildPrepareIntermediateCaseFromSources(
  args: BuildPrepareIntermediateCaseFromSourcesArgs
): IntermediateCase {
  const nowIso = String(args.nowIso || new Date().toISOString()).trim();
  const primarySource = resolvePrepareIntermediateCasePrimarySource({
    serverEmails: args.serverEmails,
    existingCase: args.existingCase,
  });
  let caseValue = buildIntermediateCaseFromSeed({
    caseId: args.caseId,
    anchorEmailKey: args.anchorEmailKey,
    conversationId: args.conversationId,
    createdAt: args.existingCase?.createdAt || nowIso,
    updatedAt: nowIso,
    lastAccessedAt: nowIso,
    emails: [],
  });

  for (const email of args.outlookEmails) {
    caseValue = mergeEmailIntoIntermediateCase(
      caseValue,
      prepareIntermediateEmailToCaseEmail(args.caseId, email)
    );
  }

  for (const email of args.existingCase?.emails || []) {
    caseValue = mergeEmailIntoIntermediateCase(caseValue, email);
  }

  for (const email of args.serverEmails) {
    caseValue = mergeEmailIntoIntermediateCase(
      caseValue,
      prepareIntermediateEmailToCaseEmail(args.caseId, email)
    );
  }

  return applyPrepareIntermediateCaseSourcePriority({
    caseValue: touchIntermediateCaseAccess(caseValue, nowIso),
    primarySource,
  });
}
