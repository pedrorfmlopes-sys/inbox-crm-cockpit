import { ClassificationMetaDraft, ClassificationFocus } from "./types";

export const GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX = "group_classification_seed_";

export const MENU: Array<{ key: ClassificationFocus; title: string }> = [
  { key: "summary", title: "Resumo" },
  { key: "principal", title: "Principal" },
  { key: "labels", title: "Etiquetas" },
  { key: "ticket", title: "Ticket" },
  { key: "references", title: "Referencias" },
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
  ticketId: "",
  labelStates: {},
  referenceGroupIds: [],
};
