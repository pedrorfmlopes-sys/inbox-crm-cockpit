export const INTERMEDIATE_CASE_SCHEMA_VERSION = 1;

export type IntermediateCaseSourceOrigin = "server" | "intermediate" | "outlook";
export type IntermediateVisibleState = "draft" | "local" | "server";
export type IntermediateServerPresence = "none" | "metadata" | "classified" | "attachments" | "complete";
export type IntermediateLocalPresence = "none" | "case_only" | "metadata" | "attachments" | "complete";
export type IntermediateClassificationSource = IntermediateCaseSourceOrigin | "user";
export type IntermediateAttachmentStorageDecision =
  | "pending"
  | "local"
  | "server"
  | "hybrid"
  | "metadata_only"
  | "skip_inline";
export type IntermediateRetentionState = "local_only" | "mixed" | "promoted";
export type IntermediateStorageRefKind = "path" | "relative_path" | "uri" | "storage_key";

export type IntermediateStorageRef = {
  kind: IntermediateStorageRefKind;
  value: string;
  label?: string;
};

export type IntermediateEmailClassification = {
  principalGroupId?: string;
  principalGroupName?: string;
  referenceGroupIds: string[];
  labels: string[];
  ticketIds: string[];
  ticketCodes: string[];
  state?: string;
  status?: string;
  classifiedAt?: string;
  classifiedSource?: IntermediateClassificationSource;
};

export type IntermediateCaseAttachment = {
  attachmentKey: string;
  id?: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
  hasContent: boolean;
  documentState?: string;
  storageDecision: IntermediateAttachmentStorageDecision;
  localRef?: IntermediateStorageRef;
  serverRef?: IntermediateStorageRef;
  previewReady: boolean;
};

export type IntermediateCaseEmail = {
  emailKey: string;
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromName?: string;
  fromEmail?: string;
  to: string[];
  cc: string[];
  receivedAtIso?: string;
  bodyText?: string;
  bodyHtml?: string;
  sourceOrigin: IntermediateCaseSourceOrigin;
  visibilityState: IntermediateVisibleState;
  serverPresence: IntermediateServerPresence;
  localPresence: IntermediateLocalPresence;
  classification: IntermediateEmailClassification;
  attachments: IntermediateCaseAttachment[];
};

export type IntermediateCaseSourceSummary = {
  precedence: IntermediateCaseSourceOrigin[];
  primarySource: IntermediateCaseSourceOrigin;
  anchorOrigin: IntermediateCaseSourceOrigin;
  hasServerBackedEmails: boolean;
  hasIntermediateBackedEmails: boolean;
  hasOutlookBackedEmails: boolean;
  serverEmailCount: number;
  intermediateEmailCount: number;
  outlookEmailCount: number;
};

export type IntermediateCaseClassificationSummary = {
  totalEmails: number;
  classifiedEmails: number;
  unclassifiedEmails: number;
  mixedCase: boolean;
  visibleState: IntermediateVisibleState;
};

export type IntermediateCaseRetentionSummary = {
  state: IntermediateRetentionState;
  lastAccessedAt: string;
  canCleanupLater: boolean;
};

export type IntermediateCaseDiagnosticSummary = {
  quickState: string;
  notes: string[];
};

export type IntermediateCase = {
  schemaVersion: number;
  caseId: string;
  anchorEmailKey: string;
  conversationId?: string;
  createdAt: string;
  updatedAt: string;
  lastAccessedAt: string;
  sourceSummary: IntermediateCaseSourceSummary;
  emails: IntermediateCaseEmail[];
  classificationSummary: IntermediateCaseClassificationSummary;
  retentionSummary: IntermediateCaseRetentionSummary;
  diagnosticSummary: IntermediateCaseDiagnosticSummary;
};

export type IntermediateCaseSeedAttachment = Partial<IntermediateCaseAttachment> & {
  attachmentKey: string;
  name: string;
};

export type IntermediateCaseSeedEmail = Partial<IntermediateCaseEmail> & {
  emailKey: string;
  attachments?: IntermediateCaseSeedAttachment[];
  classification?: Partial<IntermediateEmailClassification>;
};

export type IntermediateCaseSeed = {
  caseId?: string;
  anchorEmailKey: string;
  conversationId?: string;
  createdAt?: string;
  updatedAt?: string;
  lastAccessedAt?: string;
  emails?: IntermediateCaseSeedEmail[];
};

export type IntermediateCaseSummary = Pick<
  IntermediateCase,
  "caseId" | "anchorEmailKey" | "conversationId" | "updatedAt" | "lastAccessedAt"
> & {
  emailCount: number;
  visibleState: IntermediateVisibleState;
  retentionState: IntermediateRetentionState;
  quickState: string;
};
