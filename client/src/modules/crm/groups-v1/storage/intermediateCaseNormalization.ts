import {
  INTERMEDIATE_CASE_SCHEMA_VERSION,
  type IntermediateCase,
  type IntermediateCaseAttachment,
  type IntermediateCaseDiagnosticSummary,
  type IntermediateCaseEmail,
  type IntermediateCaseSeed,
  type IntermediateEmailClassification,
  type IntermediateStorageRef,
} from "./intermediateCaseTypes";
import {
  buildIntermediateCaseClassificationSummary,
  buildIntermediateCaseDiagnosticSummary,
  buildIntermediateCaseRetentionSummary,
  buildIntermediateCaseSourceSummary,
} from "./intermediateCaseSummary";

function normalizeString(value: unknown): string {
  return String(value || "").trim();
}

function normalizeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return Array.from(new Set(value.map((entry) => normalizeString(entry)).filter(Boolean)));
}

function pickDate(value: unknown, fallback?: string): string | undefined {
  const normalized = normalizeString(value);
  return normalized || fallback;
}

function normalizeStorageRef(input: Partial<IntermediateStorageRef> | null | undefined): IntermediateStorageRef | undefined {
  const value = normalizeString(input?.value);
  if (!value) return undefined;
  return {
    kind: input?.kind || "relative_path",
    value,
    label: normalizeString(input?.label) || undefined,
  };
}

function normalizeClassification(input: Partial<IntermediateEmailClassification> | null | undefined): IntermediateEmailClassification {
  return {
    principalGroupId: normalizeString(input?.principalGroupId) || undefined,
    principalGroupName: normalizeString(input?.principalGroupName) || undefined,
    referenceGroupIds: normalizeStringArray(input?.referenceGroupIds),
    labels: normalizeStringArray(input?.labels),
    ticketIds: normalizeStringArray(input?.ticketIds),
    ticketCodes: normalizeStringArray(input?.ticketCodes),
    state: normalizeString(input?.state) || undefined,
    status: normalizeString(input?.status) || undefined,
    classifiedAt: pickDate(input?.classifiedAt),
    classifiedSource: input?.classifiedSource || undefined,
  };
}

function normalizeAttachment(input: Partial<IntermediateCaseAttachment> | null | undefined): IntermediateCaseAttachment | null {
  const attachmentKey = normalizeString(input?.attachmentKey);
  const name = normalizeString(input?.name);
  if (!attachmentKey || !name) return null;
  return {
    attachmentKey,
    id: normalizeString(input?.id) || undefined,
    name,
    contentType: normalizeString(input?.contentType) || undefined,
    size: Number.isFinite(Number(input?.size)) ? Number(input?.size) : undefined,
    isInline: input?.isInline === true,
    contentId: normalizeString(input?.contentId) || undefined,
    hasContent: input?.hasContent === true,
    documentState: normalizeString(input?.documentState) || undefined,
    storageDecision: input?.storageDecision || "pending",
    localRef: normalizeStorageRef(input?.localRef),
    serverRef: normalizeStorageRef(input?.serverRef),
    previewReady: input?.previewReady === true,
  };
}

function normalizeEmail(input: Partial<IntermediateCaseEmail> | null | undefined): IntermediateCaseEmail | null {
  const emailKey = normalizeString(input?.emailKey);
  if (!emailKey) return null;
  const attachments = Array.isArray(input?.attachments)
    ? input.attachments.map((attachment) => normalizeAttachment(attachment)).filter(Boolean) as IntermediateCaseAttachment[]
    : [];
  return {
    emailKey,
    itemId: normalizeString(input?.itemId) || undefined,
    internetMessageId: normalizeString(input?.internetMessageId) || undefined,
    conversationId: normalizeString(input?.conversationId) || undefined,
    subject: normalizeString(input?.subject) || undefined,
    fromName: normalizeString(input?.fromName) || undefined,
    fromEmail: normalizeString(input?.fromEmail) || undefined,
    to: normalizeStringArray(input?.to),
    cc: normalizeStringArray(input?.cc),
    receivedAtIso: pickDate(input?.receivedAtIso),
    bodyText: normalizeString(input?.bodyText) || undefined,
    bodyHtml: normalizeString(input?.bodyHtml) || undefined,
    sourceOrigin: input?.sourceOrigin || "outlook",
    visibilityState: input?.visibilityState || "draft",
    serverPresence: input?.serverPresence || "none",
    localPresence: input?.localPresence || "case_only",
    classification: normalizeClassification(input?.classification),
    attachments,
  };
}

function ensureAnchorFirst(emails: IntermediateCaseEmail[], anchorEmailKey: string): IntermediateCaseEmail[] {
  const unique = new Map<string, IntermediateCaseEmail>();
  for (const email of emails) unique.set(email.emailKey, email);
  const rows = Array.from(unique.values());
  rows.sort((left, right) => {
    if (left.emailKey === anchorEmailKey) return -1;
    if (right.emailKey === anchorEmailKey) return 1;
    return String(right.receivedAtIso || "").localeCompare(String(left.receivedAtIso || ""));
  });
  return rows;
}

export function normalizeIntermediateCase(input: Partial<IntermediateCase> | null | undefined): IntermediateCase | null {
  const caseId = normalizeString(input?.caseId);
  const anchorEmailKey = normalizeString(input?.anchorEmailKey);
  if (!caseId || !anchorEmailKey) return null;
  const rawEmails = Array.isArray(input?.emails) ? input.emails : [];
  const normalizedEmails = ensureAnchorFirst(
    rawEmails
      .map((email) => normalizeEmail(email))
      .filter(Boolean) as IntermediateCaseEmail[],
    anchorEmailKey
  );
  const createdAt = pickDate(input?.createdAt, new Date().toISOString())!;
  const updatedAt = pickDate(input?.updatedAt, createdAt)!;
  const lastAccessedAt = pickDate(input?.lastAccessedAt, updatedAt)!;
  const baseCase: IntermediateCase = {
    schemaVersion: INTERMEDIATE_CASE_SCHEMA_VERSION,
    caseId,
    anchorEmailKey,
    conversationId: normalizeString(input?.conversationId) || undefined,
    createdAt,
    updatedAt,
    lastAccessedAt,
    sourceSummary: {
      precedence: ["server", "intermediate", "outlook"],
      primarySource: "outlook",
      anchorOrigin: "outlook",
      hasServerBackedEmails: false,
      hasIntermediateBackedEmails: false,
      hasOutlookBackedEmails: false,
      serverEmailCount: 0,
      intermediateEmailCount: 0,
      outlookEmailCount: 0,
    },
    emails: normalizedEmails,
    classificationSummary: {
      totalEmails: 0,
      classifiedEmails: 0,
      unclassifiedEmails: 0,
      mixedCase: false,
      visibleState: "draft",
    },
    retentionSummary: {
      state: "local_only",
      lastAccessedAt,
      canCleanupLater: true,
    },
    diagnosticSummary: {
      quickState: "Caso intermedio preparado.",
      notes: [],
    },
  };
  baseCase.sourceSummary = buildIntermediateCaseSourceSummary(baseCase);
  if (
    input?.sourceSummary?.primarySource === "server"
    || input?.sourceSummary?.primarySource === "intermediate"
    || input?.sourceSummary?.primarySource === "outlook"
  ) {
    baseCase.sourceSummary = {
      ...baseCase.sourceSummary,
      primarySource: input.sourceSummary.primarySource,
    };
  }
  baseCase.classificationSummary = buildIntermediateCaseClassificationSummary(baseCase);
  baseCase.retentionSummary = buildIntermediateCaseRetentionSummary(baseCase);
  baseCase.diagnosticSummary = buildIntermediateCaseDiagnosticSummary(baseCase);
  return baseCase;
}

export function createEmptyIntermediateCase(input: {
  caseId: string;
  anchorEmailKey: string;
  conversationId?: string;
  nowIso?: string;
}): IntermediateCase {
  const nowIso = pickDate(input.nowIso, new Date().toISOString())!;
  return normalizeIntermediateCase({
    caseId: input.caseId,
    anchorEmailKey: input.anchorEmailKey,
    conversationId: input.conversationId,
    createdAt: nowIso,
    updatedAt: nowIso,
    lastAccessedAt: nowIso,
    emails: [
      {
        emailKey: input.anchorEmailKey,
        conversationId: input.conversationId,
        sourceOrigin: "outlook",
        visibilityState: "draft",
        serverPresence: "none",
        localPresence: "case_only",
      },
    ],
  })!;
}

export function buildIntermediateCaseFromSeed(seed: IntermediateCaseSeed): IntermediateCase {
  const nowIso = pickDate(seed.updatedAt, new Date().toISOString())!;
  const fallbackCaseId = normalizeString(seed.caseId) || normalizeString(seed.anchorEmailKey);
  return normalizeIntermediateCase({
    caseId: fallbackCaseId,
    anchorEmailKey: seed.anchorEmailKey,
    conversationId: seed.conversationId,
    createdAt: pickDate(seed.createdAt, nowIso),
    updatedAt: nowIso,
    lastAccessedAt: pickDate(seed.lastAccessedAt, nowIso),
    emails: (seed.emails || []).map((email) => ({
      ...email,
      classification: normalizeClassification(email.classification),
      attachments: (email.attachments || []).map((attachment) => ({
        ...attachment,
      })),
    })),
  })!;
}

export function touchIntermediateCaseAccess(caseValue: IntermediateCase, nowIso = new Date().toISOString()): IntermediateCase {
  const next = normalizeIntermediateCase({
    ...caseValue,
    lastAccessedAt: nowIso,
    updatedAt: caseValue.updatedAt,
  });
  return next || caseValue;
}

export function mergeIntermediateDiagnosticSummary(
  current: IntermediateCaseDiagnosticSummary,
  next: Partial<IntermediateCaseDiagnosticSummary>
): IntermediateCaseDiagnosticSummary {
  return {
    quickState: normalizeString(next.quickState) || current.quickState,
    notes: Array.from(new Set([...(current.notes || []), ...normalizeStringArray(next.notes)])),
  };
}
