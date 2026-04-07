import { ClassificationMetaDraft, ClassificationFocus } from "./types";

export const GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX = "group_classification_seed_";

export const MENU: Array<{ key: ClassificationFocus; title: string }> = [
  { key: "summary", title: "Resumo" },
  { key: "principal", title: "Principal" },
  { key: "labels", title: "Etiquetas" },
  { key: "ticket", title: "Ticket" },
  { key: "references", title: "Referências" },
];

export const LABEL_STATUS_OPTIONS: Array<{ value: string; label: string; color?: string }> = [
  { value: "em_analise", label: "Em análise", color: "#f59e0b" },
  { value: "respondido", label: "Respondido", color: "#3b82f6" },
  { value: "confirmado", label: "Confirmado", color: "#10b981" },
  { value: "arquivado", label: "Arquivado", color: "#6b7280" },
  { value: "cancelado", label: "Cancelado", color: "#ef4444" },
];

export const TICKET_STATUS_OPTIONS: Array<{ value: string; label: string; color?: string }> = [
  { value: "open", label: "Aberto", color: "#3b82f6" },
  { value: "closed", label: "Fechado", color: "#10b981" },
];

export const DOCUMENT_STATE_OPTIONS: Array<{ value: string; label: string }> = [
  { value: "ingested", label: "Ingerido" },
  { value: "processed", label: "Processado" },
  { value: "accepted", label: "Aceite" },
  { value: "rejected", label: "Rejeitado" },
  { value: "reread_requested", label: "Re-leitura" },
];

export const EMPTY_CLASSIFICATION_META: ClassificationMetaDraft = {
  principalGroupId: "",
  principalCategorize: false,
  principalStatusEnabled: false,
  principalStatusCategorize: false,
  ticketId: "",
  ticketCategorize: false,
  ticketStatusEnabled: false,
  ticketStatusCategorize: false,
  categorizedLabelNames: [],
  labelStates: {},
  referenceGroupIds: [],
  referenceCategorize: false,
  referenceStatusEnabled: false,
  referenceStatusCategorize: false,
};
