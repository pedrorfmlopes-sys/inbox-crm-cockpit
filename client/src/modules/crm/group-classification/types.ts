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

export type SectionId = "emails" | "details" | "managed-group";

export interface ClassificationMetaDraft {
  principalGroupId: string;
  ticketId: string;
  categorizedLabelNames: string[];
  labelStates: Record<string, LabelDraft>;
  referenceGroupIds: string[];
}

export type DocumentLifecycleState = "ingested" | "processed" | "accepted" | "rejected" | "reread_requested";

export type ApplyDialogScopeMode = "current" | "selected" | "all";

export type ClassificationLayoutMode = "normal" | "advanced";

export type ScopeMode = "related" | "all";

export type ApplyScopeMode = "current" | "selected" | "all";

export type TicketEditorMode = "existing" | "new";

export type PreviewMode = "email" | "document" | "reply" | "forward";
