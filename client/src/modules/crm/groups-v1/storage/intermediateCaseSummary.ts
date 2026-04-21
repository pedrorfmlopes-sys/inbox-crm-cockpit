import type {
  IntermediateCase,
  IntermediateCaseClassificationSummary,
  IntermediateCaseDiagnosticSummary,
  IntermediateCaseSourceOrigin,
  IntermediateCaseRetentionSummary,
  IntermediateCaseSourceSummary,
  IntermediateCaseSummary,
  IntermediateVisibleState,
} from "./intermediateCaseTypes";

function pickVisibleState(caseValue: IntermediateCase): IntermediateVisibleState {
  if (caseValue.emails.some((email) => email.visibilityState === "server")) return "server";
  if (caseValue.emails.some((email) => email.visibilityState === "local")) return "local";
  return "draft";
}

export function buildIntermediateCaseSourceSummary(caseValue: Pick<IntermediateCase, "emails">): IntermediateCaseSourceSummary {
  const emails = caseValue.emails;
  const serverEmailCount = emails.filter((email) => email.sourceOrigin === "server").length;
  const intermediateEmailCount = emails.filter((email) => email.sourceOrigin === "intermediate").length;
  const outlookEmailCount = emails.filter((email) => email.sourceOrigin === "outlook").length;
  const primarySource: IntermediateCaseSourceOrigin = serverEmailCount
    ? "server"
    : intermediateEmailCount
      ? "intermediate"
      : "outlook";
  return {
    precedence: ["server", "intermediate", "outlook"],
    primarySource,
    anchorOrigin: emails[0]?.sourceOrigin || "outlook",
    hasServerBackedEmails: serverEmailCount > 0,
    hasIntermediateBackedEmails: intermediateEmailCount > 0,
    hasOutlookBackedEmails: outlookEmailCount > 0,
    serverEmailCount,
    intermediateEmailCount,
    outlookEmailCount,
  };
}

export function buildIntermediateCaseClassificationSummary(
  caseValue: Pick<IntermediateCase, "emails">
): IntermediateCaseClassificationSummary {
  const totalEmails = caseValue.emails.length;
  const classifiedEmails = caseValue.emails.filter((email) =>
    Boolean(
      email.classification.principalGroupId
      || email.classification.referenceGroupIds.length
      || email.classification.labels.length
      || email.classification.ticketIds.length
      || email.classification.ticketCodes.length
      || email.classification.status
      || email.classification.state
    )
  ).length;
  const unclassifiedEmails = Math.max(0, totalEmails - classifiedEmails);
  return {
    totalEmails,
    classifiedEmails,
    unclassifiedEmails,
    mixedCase: classifiedEmails > 0 && unclassifiedEmails > 0,
    visibleState: pickVisibleState(caseValue as IntermediateCase),
  };
}

export function buildIntermediateCaseRetentionSummary(
  caseValue: Pick<IntermediateCase, "emails" | "lastAccessedAt">
): IntermediateCaseRetentionSummary {
  const hasServer = caseValue.emails.some((email) => email.serverPresence !== "none");
  const hasLocal = caseValue.emails.some((email) => email.localPresence !== "none");
  return {
    state: hasServer && hasLocal ? "mixed" : hasServer ? "promoted" : "local_only",
    lastAccessedAt: caseValue.lastAccessedAt,
    canCleanupLater: hasLocal,
  };
}

export function buildIntermediateCaseDiagnosticSummary(
  caseValue: Pick<IntermediateCase, "emails" | "retentionSummary" | "classificationSummary">
): IntermediateCaseDiagnosticSummary {
  const notes: string[] = [];
  if (caseValue.classificationSummary.mixedCase) notes.push("Caso misto com emails classificados e por classificar.");
  if (caseValue.emails.some((email) => email.attachments.some((attachment) => attachment.storageDecision === "pending"))) {
    notes.push("Existem anexos com decisao de storage pendente.");
  }
  if (caseValue.retentionSummary.state === "mixed") {
    notes.push("O caso combina dados locais com dados ja promovidos.");
  }
  const quickState = caseValue.classificationSummary.visibleState === "server"
    ? "Caso com classificacao funcional no servidor."
    : caseValue.classificationSummary.visibleState === "local"
      ? "Caso intermedio local pronto para retoma."
      : "Caso ainda em rascunho local.";
  return { quickState, notes };
}

export function buildIntermediateCaseSummary(caseValue: IntermediateCase): IntermediateCaseSummary {
  return {
    caseId: caseValue.caseId,
    anchorEmailKey: caseValue.anchorEmailKey,
    conversationId: caseValue.conversationId,
    updatedAt: caseValue.updatedAt,
    lastAccessedAt: caseValue.lastAccessedAt,
    emailCount: caseValue.emails.length,
    visibleState: caseValue.classificationSummary.visibleState,
    retentionState: caseValue.retentionSummary.state,
    quickState: caseValue.diagnosticSummary.quickState,
  };
}
