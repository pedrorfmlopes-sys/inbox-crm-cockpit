export interface GroupLabelDraft {
  categorize: boolean;
  hasStatus: boolean;
  status?: string;
}

export type EmailLabelStatus = "em_analise" | "respondido" | "confirmado" | "arquivado" | "cancelado";

export interface LabelDraft {
  categorize: boolean;
  hasStatus: boolean;
  status?: EmailLabelStatus;
}

export type ClassificationFocus = "summary" | "principal" | "labels" | "ticket" | "references";

export type SectionId = "emails" | "details" | "managed-group" | "classification" | "groups" | "labels" | "filters" | "ticket" | "principal" | "references" | "summary";

export interface ClassificationMetaDraft {
  principalGroupId: string;
  principalCategorize: boolean;
  principalStatusEnabled: boolean;
  principalStatusCategorize: boolean;
  ticketId: string;
  ticketCategorize: boolean;
  ticketStatusEnabled: boolean;
  ticketStatusCategorize: boolean;
  categorizedLabelNames: string[];
  labelStates: Record<string, LabelDraft>;
  referenceGroupIds: string[];
  referenceCategorize: boolean;
  referenceStatusEnabled: boolean;
  referenceStatusCategorize: boolean;
}

export type DocumentLifecycleState = "ingested" | "processed" | "accepted" | "rejected" | "reread_requested";

export type ApplyDialogScopeMode = "current" | "selected" | "all" | "case_all";

export type ClassificationLayoutMode = "normal" | "advanced";

export type ScopeMode = "related" | "all";

export type ApplyScopeMode = "current" | "selected" | "all" | "principal_group";

export interface GroupContactDraft {
  key: string;
  name: string;
  email: string;
  company?: string;
  source?: string;
  role: string;
  isPrincipal: boolean;
}

export interface GroupEntityDraft {
  key: string;
  id?: string;
  name: string;
  kind?: string;
  source?: string;
  role: string;
  isPrincipal: boolean;
}

export type PreviewMode = "email" | "document" | "reply" | "forward";

export interface AttachmentPreviewState {
  kind: "office" | "text" | "image" | "pdf" | "unsupported";
  url?: string;
  src?: string;
  text?: string;
}

export interface ReadingSuggestionChip {
  key: string;
  kind: "group" | "ticket" | "label";
  label: string;
  value: string;
  active?: boolean;
  score?: number;
}

export type TicketEditorMode = "existing" | "new";

export interface StudioParams {
  seedKey?: string;
  prepareSeedKey?: string;
  caseId?: string;
  anchorEmailKey?: string;
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  receivedAtIso?: string;
}
