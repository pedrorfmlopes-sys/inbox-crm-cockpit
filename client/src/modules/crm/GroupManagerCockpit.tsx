import React, { useEffect, useMemo, useState } from "react";
import {
  addEmailToLinkGroup,
  createGroupTicket,
  createGroupTicketSeries,
  createLinkGroup,
  deleteLinkGroup,
  deleteGroupTicketSeries,
  detectGroupTicketsForEmail,
  getRelatedEmailContext,
  getGroupEmails,
  listLinkGroups,
  listGroupTicketSeries,
  removeEmailFromLinkGroup,
  registerRelevantEmail,
  saveGroupDocuments,
  searchKnownEmails,
  searchGroupTickets,
  type GroupTicketDetectionMatch,
  type GroupTicketEntry,
  type GroupTicketSeriesEntry,
  updateGroupTicket,
  updateLinkGroup,
  updateGroupTicketSeries,
  linkEmailToGroupTicket,
  type LinkGroupEntry,
  type RelevantEmailPayload,
  type RelatedEmailEntry,
} from "@/api";
import { aiGenerate } from "@/ai/aiClient";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { displayNewMessageForm, displayReplyForm, getManagedOutlookCategorySnapshot, openGroupClassificationStudio, openGroupExplorer, openGroupSettings, openLinkedOutlookEmail, setSubjectInComposeDraft, syncOutlookCategorySource } from "@/office";
import { buildOutlookCategorySourceFromRelatedContext } from "@/outlookCategories";
import {
  findGroupLabelCatalogEntry,
  getGroupLabelCatalogLabels,
  normalizeGroupLabelCatalog,
  saveSettings,
  type GroupLabelCatalogEntry,
  type GroupLabelStatus,
} from "@/settings";
import { HelpHint } from "@/ui/HelpHint";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type GroupManagerView = "groups" | "detail" | "library" | "quicklink" | "settings" | "labels" | "tickets";
type GroupStatusFilter = "all" | "em_analise" | "em_progresso" | "concluido";
type GroupArchiveFilter = "active" | "archived" | "all";
type MembershipKind = "principal" | "referencia";
type ManagedLabelDraft = {
  categorize: boolean;
  hasStatus: boolean;
  status: GroupLabelStatus;
};

type TicketSeriesDraft = {
  name: string;
  prefix: string;
  replyInstructions: string;
  yearMode: "none" | "yy" | "yyyy";
  separator: "-" | "/" | "_" | " " | "";
  nextNumber: string;
  padding: string;
  isActive: boolean;
};

const GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX = "iccc_group_classification_seed_v1:";

type TicketUiDraft = {
  autoLinkMode: "confirm" | "auto";
  suggestDraftOnCreate: boolean;
  useAiDrafts: boolean;
  includeTicketCodeInSubject: boolean;
  aiInstructions: string;
};

type QuickLinkDraft = {
  ticketMode: "none" | "existing" | "new";
  ticketId: string;
  ticketSeriesId: string;
  ticketStatus: "open" | "closed" | string;
  principalGroupId: string;
  secondaryGroupIds: string[];
  labelsText: string;
};

type GroupDraft = {
  name: string;
  description: string;
  status: "" | "em_analise" | "em_progresso" | "concluido";
  labelsText: string;
  documentsEnabled: boolean;
  isArchived: boolean;
};

const STATUS_OPTIONS: Array<{ value: GroupDraft["status"]; label: string }> = [
  { value: "", label: "Sem estado" },
  { value: "em_analise", label: "Em analise" },
  { value: "em_progresso", label: "Em progresso" },
  { value: "concluido", label: "Concluido" },
];

const LABEL_STATUS_OPTIONS: Array<{ value: GroupLabelStatus; label: string }> = [
  { value: "em_analise", label: "Em analise" },
  { value: "em_progresso", label: "Em progresso" },
  { value: "concluido", label: "Concluido" },
];

const MEMBERSHIP_OPTIONS: Array<{ value: MembershipKind; label: string }> = [
  { value: "principal", label: "Principal" },
  { value: "referencia", label: "Referencia" },
];

const TICKET_YEAR_MODE_OPTIONS: Array<{ value: TicketSeriesDraft["yearMode"]; label: string }> = [
  { value: "none", label: "Sem ano" },
  { value: "yy", label: "Ano curto" },
  { value: "yyyy", label: "Ano longo" },
];

const TICKET_SEPARATOR_OPTIONS: Array<{ value: TicketSeriesDraft["separator"]; label: string }> = [
  { value: "-", label: "-" },
  { value: "/", label: "/" },
  { value: "_", label: "_" },
  { value: " ", label: "Espaco" },
  { value: "", label: "Sem separador" },
];

const TICKET_STATUS_OPTIONS: Array<{ value: "open" | "closed"; label: string }> = [
  { value: "open", label: "Aberto" },
  { value: "closed", label: "Fechado" },
];

function normalizeText(value: string | undefined): string {
  return String(value || "").trim().toLowerCase();
}

function normalizeTicketPrefixInput(value: string | undefined): string {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

function confirmAction(message: string): boolean {
  if (typeof window === "undefined" || typeof window.confirm !== "function") {
    return true;
  }
  try {
    return window.confirm(message);
  } catch {
    return true;
  }
}

function buildTicketSeriesPreview(draft: Pick<TicketSeriesDraft, "prefix" | "yearMode" | "separator" | "padding" | "nextNumber">): string {
  const prefix = normalizeTicketPrefixInput(draft.prefix) || "TCK";
  const separator = draft.separator ?? "-";
  const currentYear = String(new Date().getUTCFullYear());
  const yearValue = draft.yearMode === "yyyy" ? currentYear : draft.yearMode === "yy" ? currentYear.slice(-2) : "";
  const number = String(Math.max(1, Number(draft.nextNumber || 1) || 1)).padStart(Math.max(2, Number(draft.padding || 4) || 4), "0");
  return [prefix, yearValue, number].filter(Boolean).join(separator);
}

function parseLabels(value: string | string[] | undefined): string[] {
  const raw = Array.isArray(value)
    ? value
    : String(value || "")
      .split(/[,\n;]/g)
      .map((entry) => entry.trim());
  const seen = new Set<string>();
  const labels: string[] = [];
  for (const item of raw) {
    const label = String(item || "").trim();
    if (!label) continue;
    const key = label.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    labels.push(label);
  }
  return labels.sort((a, b) => a.localeCompare(b, "pt-PT"));
}

function stripEmailPayloadAttachmentContent<T extends { attachments?: RelevantEmailPayload["attachments"] }>(payload: T): T {
  if (!payload || !Array.isArray(payload.attachments) || !payload.attachments.length) {
    return payload;
  }
  return {
    ...payload,
    attachments: payload.attachments.map(({ content: _content, ...attachment }) => attachment),
  };
}

function splitDocumentsIntoBatches<T extends { contentBase64?: string }>(documents: T[], maxBatchChars = 8_000_000): T[][] {
  const batches: T[][] = [];
  let currentBatch: T[] = [];
  let currentSize = 0;

  for (const document of documents) {
    const size = String(document?.contentBase64 || "").length;
    if (currentBatch.length && currentSize + size > maxBatchChars) {
      batches.push(currentBatch);
      currentBatch = [];
      currentSize = 0;
    }
    currentBatch.push(document);
    currentSize += size;
  }

  if (currentBatch.length) batches.push(currentBatch);
  return batches;
}

function makeAttachmentSelectionKey(attachment: {
  id?: string;
  name?: string;
  size?: number;
  contentId?: string;
  contentType?: string;
}): string {
  return (
    String(attachment?.id || "").trim()
    || [
      String(attachment?.name || "").trim(),
      String(attachment?.size || "").trim(),
      String(attachment?.contentId || "").trim(),
      String(attachment?.contentType || "").trim(),
    ].join("|")
  );
}

function mergeLabelCatalog(...sources: Array<string[] | undefined>): string[] {
  const seen = new Map<string, string>();
  for (const source of sources) {
    for (const raw of source || []) {
      const label = String(raw || "").trim();
      if (!label) continue;
      const key = label.toLowerCase();
      if (!seen.has(key)) seen.set(key, label);
    }
  }
  return Array.from(seen.values()).sort((a, b) => a.localeCompare(b, "pt-PT"));
}

function canonicalizeLabelsWithCatalog(labels: string[], catalog: string[]): string[] {
  const map = new Map<string, string>();
  for (const label of catalog || []) {
    const normalized = String(label || "").trim();
    if (!normalized) continue;
    map.set(normalized.toLowerCase(), normalized);
  }
  return parseLabels(labels.map((label) => map.get(String(label || "").trim().toLowerCase()) || label));
}

function labelsToText(labels: string[] | undefined): string {
  return (labels || []).join(", ");
}

function createManagedLabelDraft(entry?: Partial<GroupLabelCatalogEntry> | null): ManagedLabelDraft {
  const hasStatus = entry?.hasStatus === true;
  return {
    categorize: entry?.categorize === true,
    hasStatus,
    status: hasStatus ? (entry?.status || "em_analise") : "em_analise",
  };
}

function buildGroupLabelCatalogEntry(label: string, draft?: Partial<ManagedLabelDraft> | null): GroupLabelCatalogEntry {
  const normalizedLabel = String(label || "").trim();
  const hasStatus = draft?.hasStatus === true;
  return {
    label: normalizedLabel,
    categorize: draft?.categorize === true,
    hasStatus,
    status: hasStatus ? (draft?.status || "em_analise") : undefined,
  };
}

function statusLabel(value: string | undefined): string {
  if (!String(value || "").trim()) return "Sem estado";
  if (value === "concluido") return "Concluido";
  if (value === "em_progresso") return "Em progresso";
  return "Em analise";
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function membershipKindLabel(value: string | undefined): string {
  return String(value || "").trim().toLowerCase() === "referencia" ? "Referencia" : "Principal";
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return (
    String(email?.emailKey || "").trim()
    || String(email?.id || "").trim()
    || String(email?.itemId || "").trim()
    || String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
    || [
      String(email?.conversationId || "").trim(),
      String(email?.subject || "").trim().toLowerCase(),
      String(email?.fromEmail || "").trim().toLowerCase(),
      String(email?.messageDateIso || email?.receivedAtIso || "").trim(),
    ].join("|")
  );
}

function createDraft(group: LinkGroupEntry | null): GroupDraft {
  return {
    name: String(group?.name || "").trim(),
    description: String(group?.description || "").trim(),
    status: group?.status === "em_analise" || group?.status === "em_progresso" || group?.status === "concluido"
      ? group.status
      : "",
    labelsText: labelsToText(group?.labels),
    documentsEnabled: group?.documentsEnabled !== false,
    isArchived: group?.isArchived === true,
  };
}

function isCurrentContextEmail(email: Partial<RelatedEmailEntry>, ctx: ReturnType<typeof useCockpit>["ctx"]): boolean {
  const currentItemId = String(ctx.itemId || "").trim();
  const emailItemId = String(email.itemId || "").trim();
  if (currentItemId && emailItemId && currentItemId === emailItemId) return true;

  const currentMessageId = String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  const emailMessageId = String(email.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  if (currentMessageId && emailMessageId && currentMessageId === emailMessageId) return true;

  return Boolean(
    String(ctx.conversationId || "").trim()
    && String(email.conversationId || "").trim()
    && String(ctx.conversationId || "").trim() === String(email.conversationId || "").trim()
    && normalizeText(ctx.subject) === normalizeText(email.subject)
  );
}

function draftChanged(group: LinkGroupEntry | null, draft: GroupDraft): boolean {
  if (!group) return false;
  return (
    String(group.name || "") !== draft.name
    || String(group.description || "") !== draft.description
    || String(group.status || "") !== draft.status
    || labelsToText(group.labels) !== labelsToText(parseLabels(draft.labelsText))
    || (group.documentsEnabled !== false) !== draft.documentsEnabled
    || (group.isArchived === true) !== draft.isArchived
  );
}

function createTicketSeriesDraft(series: GroupTicketSeriesEntry | null): TicketSeriesDraft {
  return {
    name: String(series?.name || "").trim(),
    prefix: normalizeTicketPrefixInput(series?.prefix),
    replyInstructions: String(series?.replyInstructions || "").trim(),
    yearMode: series?.yearMode === "yy" || series?.yearMode === "yyyy" ? series.yearMode : "none",
    separator: series?.separator === "/" || series?.separator === "_" || series?.separator === " " || series?.separator === ""
      ? series.separator
      : "-",
    nextNumber: String(Number(series?.nextNumber || 1) || 1),
    padding: String(Number(series?.padding || 4) || 4),
    isActive: series?.isActive !== false,
  };
}

function ticketSeriesDraftChanged(series: GroupTicketSeriesEntry | null, draft: TicketSeriesDraft): boolean {
  if (!series) return false;
  return (
    String(series.name || "").trim() !== String(draft.name || "").trim()
    || String(series.prefix || "").trim() !== String(draft.prefix || "").trim()
    || String(series.replyInstructions || "").trim() !== String(draft.replyInstructions || "").trim()
    || String(series.yearMode || "none") !== String(draft.yearMode || "none")
    || String(series.separator || "-") !== String(draft.separator || "-")
    || Number(series.nextNumber || 1) !== Math.max(1, Number(draft.nextNumber || 1) || 1)
    || Number(series.padding || 4) !== Math.max(2, Number(draft.padding || 4) || 4)
    || (series.isActive !== false) !== Boolean(draft.isActive)
  );
}

function createTicketUiDraft(value: any): TicketUiDraft {
  return {
    autoLinkMode: value?.autoLinkMode === "auto" ? "auto" : "confirm",
    suggestDraftOnCreate: value?.suggestDraftOnCreate !== false,
    useAiDrafts: value?.useAiDrafts !== false,
    includeTicketCodeInSubject: value?.includeTicketCodeInSubject !== false,
    aiInstructions: String(value?.aiInstructions || "").trim(),
  };
}

function ticketUiDraftChanged(current: TicketUiDraft, next: TicketUiDraft): boolean {
  return (
    current.autoLinkMode !== next.autoLinkMode
    || current.suggestDraftOnCreate !== next.suggestDraftOnCreate
    || current.useAiDrafts !== next.useAiDrafts
    || current.includeTicketCodeInSubject !== next.includeTicketCodeInSubject
    || current.aiInstructions !== next.aiInstructions
  );
}

function buildTicketEmailSubject(baseSubject: string | undefined, ticketCode: string | undefined, includeTicketCode: boolean): string {
  const code = String(ticketCode || "").trim();
  const raw = String(baseSubject || "").trim();
  if (!includeTicketCode || !code) return raw;
  if (raw.toLowerCase().includes(code.toLowerCase())) return raw;

  const prefixMatch = raw.match(/^((?:(?:re|fw|fwd)\s*:\s*)+)/i);
  const prefixes = prefixMatch?.[1] || "";
  const remainder = raw.slice(prefixes.length).trim();

  if (prefixes) {
    return `${prefixes}[${code}] ${remainder || "Ticket"}`.trim();
  }
  return `[${code}] ${raw || "Ticket"}`.trim();
}

function createQuickLinkDraft(partial: Partial<QuickLinkDraft> = {}): QuickLinkDraft {
  return {
    ticketMode: partial.ticketMode === "existing" || partial.ticketMode === "new" ? partial.ticketMode : "none",
    ticketId: String(partial.ticketId || "").trim(),
    ticketSeriesId: String(partial.ticketSeriesId || "").trim(),
    ticketStatus: String(partial.ticketStatus || "").trim() || "open",
    principalGroupId: String(partial.principalGroupId || "").trim(),
    secondaryGroupIds: Array.from(new Set((partial.secondaryGroupIds || []).map((entry) => String(entry || "").trim()).filter(Boolean))),
    labelsText: String(partial.labelsText || "").trim(),
  };
}

function createQuickLinkDraftFromTicket(ticket: GroupTicketEntry | null): QuickLinkDraft {
  const groupIds = Array.from(new Set((ticket?.groupIds || []).map((entry) => String(entry || "").trim()).filter(Boolean)));
  return createQuickLinkDraft({
    ticketMode: ticket ? "existing" : "none",
    ticketId: String(ticket?.id || "").trim(),
    ticketStatus: String(ticket?.status || "").trim() || "open",
    principalGroupId: groupIds[0] || "",
    secondaryGroupIds: groupIds.slice(1),
    labelsText: labelsToText(ticket?.labels),
  });
}

function uniqueEmails(values: Array<string | undefined>): string[] {
  return Array.from(new Set(values.map((value) => String(value || "").trim()).filter(Boolean)));
}

function buildSeriesReplyGuidance(series: GroupTicketSeriesEntry | null, ticket: GroupTicketEntry | null): string {
  const explicit = String(series?.replyInstructions || "").trim();
  if (explicit) return explicit;

  const raw = [
    String(series?.name || "").trim(),
    String(series?.prefix || "").trim(),
    String(ticket?.title || "").trim(),
  ].join(" ").toLowerCase();

  if (/(encomenda|pedido|orcamento|cotacao)/.test(raw)) {
    return "Agradece o pedido ou encomenda recebida, confirma que vamos dar seguimento e informa que, assim que houver atualizacoes, iremos comunicar.";
  }
  if (/(reclam|incid|suporte|assistencia|pos[- ]?venda)/.test(raw)) {
    return "Agradece o contacto, confirma a rececao da situacao reportada e informa que vamos analisar o caso, dar seguimento e comunicar assim que existirem atualizacoes.";
  }
  return "Agradece o contacto, confirma a rececao do email e informa que vamos dar seguimento ao assunto, comunicando assim que existirem atualizacoes.";
}

const LabelPicker: React.FC<{
  value: string;
  onChange: (next: string) => void;
  catalog: string[];
  enabled: boolean;
  busy: boolean;
  placeholder?: string;
  helpText?: string;
  onCreateLabel: (label: string) => Promise<string | null>;
}> = ({ value, onChange, catalog, enabled, busy, placeholder, helpText, onCreateLabel }) => {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);

  const selectedLabels = useMemo(() => parseLabels(value), [value]);

  const suggestions = useMemo(() => {
    const q = normalizeText(query);
    return (catalog || [])
      .filter((label) => !selectedLabels.some((entry) => normalizeText(entry) === normalizeText(label)))
      .filter((label) => !q || normalizeText(label).includes(q))
      .slice(0, 8);
  }, [catalog, query, selectedLabels]);

  const exactCatalogMatch = useMemo(
    () => (catalog || []).some((label) => normalizeText(label) === normalizeText(query)),
    [catalog, query]
  );

  const canCreate = enabled && Boolean(String(query || "").trim()) && !exactCatalogMatch && !suggestions.length;

  function commitLabel(label: string) {
    onChange(labelsToText(canonicalizeLabelsWithCatalog([...selectedLabels, label], catalog)));
    setQuery("");
    setOpen(false);
  }

  async function handleAction() {
    if (canCreate) {
      const created = await onCreateLabel(query);
      if (created) commitLabel(created);
      return;
    }
    setOpen(true);
  }

  return (
    <div style={S.labelPicker}>
      <div style={S.sectionTitleRow}>
        <div style={S.fieldLabel}>Etiquetas</div>
        {helpText ? <HelpHint text={helpText} title="Ajuda: Etiquetas" /> : null}
      </div>
      <div style={S.labelSearchRow}>
        <input
          style={S.input}
          value={query}
          onChange={(event) => {
            setQuery(event.target.value);
            setOpen(true);
          }}
          onFocus={() => setOpen(true)}
          onBlur={() => window.setTimeout(() => setOpen(false), 120)}
          placeholder={placeholder || "Pesquisar etiqueta"}
        />
        <button
          type="button"
          style={canCreate ? S.primaryBtn : S.iconGhostBtn}
          onClick={() => void handleAction()}
          disabled={busy || (!canCreate && !suggestions.length)}
          title={canCreate ? "Criar etiqueta e adicionar ao grupo" : "Pesquisar etiquetas existentes"}
        >
          {canCreate ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
        </button>
      </div>
      {open && suggestions.length ? (
        <div style={S.labelSuggestionList}>
          {suggestions.map((label) => (
            <button
              key={label}
              type="button"
              style={S.labelSuggestion}
              onMouseDown={(event) => {
                event.preventDefault();
                commitLabel(label);
              }}
            >
              {label}
            </button>
          ))}
        </div>
      ) : null}
      {selectedLabels.length ? (
        <div style={S.selectedLabelRow}>
          {selectedLabels.map((label) => (
            <button
              key={label}
              type="button"
              style={S.selectedLabelChip}
              onClick={() =>
                onChange(labelsToText(selectedLabels.filter((entry) => normalizeText(entry) !== normalizeText(label))))
              }
              title="Remover etiqueta"
            >
              <span>{label}</span>
              <span style={S.selectedLabelRemove}>×</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
};

type GroupManagerCockpitProps = {
  initialView?: GroupManagerView;
  standaloneSettings?: boolean;
};

function normalizeGroupSettingsView(view: GroupManagerView | undefined, standaloneSettings: boolean): GroupManagerView {
  if (standaloneSettings) {
    return view === "labels" || view === "tickets" ? view : "settings";
  }
  return view || "groups";
}

export const GroupManagerCockpit: React.FC<GroupManagerCockpitProps> = ({
  initialView,
  standaloneSettings = false,
}) => {
  const { ctx, bodyText, bodyHtml, attachments, setMsg, setActiveGroupForCurrentEmail, settings, openSettingsSection } = useCockpit();
  const [view, setView] = useState<GroupManagerView>(() => normalizeGroupSettingsView(initialView, standaloneSettings));
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [groupsLoading, setGroupsLoading] = useState(false);
  const [groupsError, setGroupsError] = useState("");
  const [groupQuery, setGroupQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState<GroupStatusFilter>("all");
  const [archiveFilter, setArchiveFilter] = useState<GroupArchiveFilter>("active");
  const [activeLabelFilters, setActiveLabelFilters] = useState<string[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState("");
  const [draft, setDraft] = useState<GroupDraft>(createDraft(null));
  const [busy, setBusy] = useState(false);
  const [reloadToken, setReloadToken] = useState(0);

  const [groupEmails, setGroupEmails] = useState<RelatedEmailEntry[]>([]);
  const [groupEmailsLoading, setGroupEmailsLoading] = useState(false);
  const [groupEmailQuery, setGroupEmailQuery] = useState("");
  const [selectedGroupEmailKeys, setSelectedGroupEmailKeys] = useState<string[]>([]);
  const [selectedCurrentAttachmentKeys, setSelectedCurrentAttachmentKeys] = useState<string[]>([]);
  const [linkKind, setLinkKind] = useState<MembershipKind>("principal");

  const [libraryQuery, setLibraryQuery] = useState("");
  const [libraryEmails, setLibraryEmails] = useState<RelatedEmailEntry[]>([]);
  const [libraryLoading, setLibraryLoading] = useState(false);
  const [selectedLibraryKeys, setSelectedLibraryKeys] = useState<string[]>([]);
  const [newCatalogLabel, setNewCatalogLabel] = useState("");
  const [newCatalogLabelDraft, setNewCatalogLabelDraft] = useState<ManagedLabelDraft>(createManagedLabelDraft());
  const [selectedManagedLabel, setSelectedManagedLabel] = useState("");
  const [renameLabelValue, setRenameLabelValue] = useState("");
  const [managedLabelDraft, setManagedLabelDraft] = useState<ManagedLabelDraft>(createManagedLabelDraft());
  const [ticketSeries, setTicketSeries] = useState<GroupTicketSeriesEntry[]>([]);
  const [ticketSeriesLoading, setTicketSeriesLoading] = useState(false);
  const [selectedTicketSeriesId, setSelectedTicketSeriesId] = useState("");
  const [ticketSeriesDraft, setTicketSeriesDraft] = useState<TicketSeriesDraft>(createTicketSeriesDraft(null));
  const [newTicketSeriesDraft, setNewTicketSeriesDraft] = useState<TicketSeriesDraft>({ name: "", prefix: "", replyInstructions: "", yearMode: "none", separator: "-", nextNumber: "1", padding: "4", isActive: true });
  const [ticketSearchQuery, setTicketSearchQuery] = useState("");
  const [ticketSearchResults, setTicketSearchResults] = useState<GroupTicketEntry[]>([]);
  const [ticketSearchLoading, setTicketSearchLoading] = useState(false);
  const [currentEmailTickets, setCurrentEmailTickets] = useState<GroupTicketEntry[]>([]);
  const [ticketMatches, setTicketMatches] = useState<GroupTicketDetectionMatch[]>([]);
  const [ticketDetectionLoading, setTicketDetectionLoading] = useState(false);
  const [ticketMatchGroupSelection, setTicketMatchGroupSelection] = useState<Record<string, string[]>>({});
  const [ticketUiDraft, setTicketUiDraft] = useState<TicketUiDraft>(createTicketUiDraft(settings?.groupTicketUi));
  const [quickLinkDraft, setQuickLinkDraft] = useState<QuickLinkDraft>(createQuickLinkDraft());
  const [quickLinkTicketQuery, setQuickLinkTicketQuery] = useState("");
  const [quickLinkTicketResults, setQuickLinkTicketResults] = useState<GroupTicketEntry[]>([]);
  const [quickLinkTicketLoading, setQuickLinkTicketLoading] = useState(false);
  const [quickLinkPrimaryQuery, setQuickLinkPrimaryQuery] = useState("");
  const [quickLinkSecondaryQuery, setQuickLinkSecondaryQuery] = useState("");

  const selectedGroup = useMemo(
    () => groups.find((group) => group.id === selectedGroupId) || null,
    [groups, selectedGroupId]
  );
  const selectedTicketSeries = useMemo(
    () => ticketSeries.find((series) => series.id === selectedTicketSeriesId) || null,
    [selectedTicketSeriesId, ticketSeries]
  );

  const currentEmailPayload = useMemo(
    () => ({
      itemId: String(ctx.itemId || "").trim(),
      internetMessageId: String(ctx.internetMessageId || "").trim(),
      conversationId: String(ctx.conversationId || "").trim(),
      subject: String(ctx.subject || "").trim(),
      fromEmail: String(ctx.fromEmail || "").trim(),
      fromName: String(ctx.fromName || "").trim(),
      receivedAtIso: String(ctx.receivedDateTimeIso || "").trim(),
      messageDateIso: String(ctx.receivedDateTimeIso || "").trim(),
      bodyText: String(bodyText || "").trim(),
      bodyHtml: String(bodyHtml || "").trim(),
      attachments: (attachments || []).map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        contentId: attachment.contentId,
      })),
    }),
    [attachments, bodyHtml, bodyText, ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject]
  );

  const currentEmailLinkPayload = useMemo(() => stripEmailPayloadAttachmentContent(currentEmailPayload), [currentEmailPayload]);
  const currentEmailKey = useMemo(() => makeEmailKey(currentEmailLinkPayload), [currentEmailLinkPayload]);
  const favoriteGroupIds = settings?.groupFavoriteIds || [];
  const favoriteGroupSet = useMemo(() => new Set(favoriteGroupIds), [favoriteGroupIds]);
  const groupTicketsEnabled = settings?.groupTicketsEnabled !== false;
  const ticketUi = settings?.groupTicketUi;
  const currentSavableAttachments = useMemo(
    () =>
      (attachments || [])
        .filter((attachment) => String(attachment.name || "").trim() && String(attachment.content || "").trim())
        .filter((attachment) => !(settings?.groupStorage.ignoreInlineAttachments && attachment.isInline)),
    [attachments, settings?.groupStorage.ignoreInlineAttachments]
  );
  const selectedCurrentAttachments = useMemo(() => {
    const selectedSet = new Set(selectedCurrentAttachmentKeys);
    return currentSavableAttachments.filter((attachment) => selectedSet.has(makeAttachmentSelectionKey(attachment)));
  }, [currentSavableAttachments, selectedCurrentAttachmentKeys]);

  useEffect(() => {
    if (!standaloneSettings) return;
    if (view !== "settings" && view !== "labels" && view !== "tickets") {
      setView("settings");
    }
  }, [standaloneSettings, view]);

  useEffect(() => {
    setSelectedCurrentAttachmentKeys(currentSavableAttachments.map((attachment) => makeAttachmentSelectionKey(attachment)));
  }, [currentEmailKey, currentSavableAttachments]);

  async function handleOpenClassificationStudio() {
    const studioSeedPayload = {
      ...currentEmailLinkPayload,
      bodyText: String(bodyText || "").trim(),
      bodyHtml: String(bodyHtml || "").trim(),
      attachments: (attachments || []).map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        contentId: attachment.contentId,
        content: String(attachment.content || "").trim(),
      })),
    };
    const params: Record<string, string> = {};
    if (currentEmailLinkPayload.itemId) params.itemId = currentEmailLinkPayload.itemId;
    if (currentEmailLinkPayload.internetMessageId) params.internetMessageId = currentEmailLinkPayload.internetMessageId;
    if (currentEmailLinkPayload.conversationId) params.conversationId = currentEmailLinkPayload.conversationId;
    if (currentEmailLinkPayload.subject) params.subject = currentEmailLinkPayload.subject;
    if (currentEmailLinkPayload.fromEmail) params.fromEmail = currentEmailLinkPayload.fromEmail;
    if (currentEmailLinkPayload.fromName) params.fromName = currentEmailLinkPayload.fromName;
    if (currentEmailLinkPayload.receivedAtIso) params.receivedAtIso = currentEmailLinkPayload.receivedAtIso;
    try {
      const seedKey = `${GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX}${Date.now()}:${currentEmailLinkPayload.itemId || currentEmailLinkPayload.internetMessageId || currentEmailLinkPayload.conversationId || "email"}`;
      localStorage.setItem(seedKey, JSON.stringify(studioSeedPayload));
      params.seedKey = seedKey;
    } catch {
      // best effort only; keep opening the studio even if the snapshot cache fails
    }
    try {
      if (
        currentEmailLinkPayload.itemId
        || currentEmailLinkPayload.internetMessageId
        || currentEmailLinkPayload.conversationId
        || currentEmailLinkPayload.subject
      ) {
        await registerRelevantEmail(studioSeedPayload);
      }
    } catch {
      // keep opening the studio even if the pre-registration fails
    }
    try {
      await openGroupClassificationStudio(params);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel abrir a janela externa do Classificar.");
    }
  }

  useEffect(() => {
    let cancelled = false;
    setGroupsLoading(true);
    setGroupsError("");
    listLinkGroups("/")
      .then((rows) => {
        if (cancelled) return;
        setGroups(Array.isArray(rows) ? rows : []);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setGroups([]);
        setGroupsError(error?.message || "Nao foi possivel carregar grupos.");
      })
      .finally(() => {
        if (!cancelled) setGroupsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  useEffect(() => {
    setDraft(createDraft(selectedGroup));
  }, [selectedGroup]);

  useEffect(() => {
    if (!selectedGroupId) {
      setGroupEmails([]);
      setSelectedGroupEmailKeys([]);
      return;
    }
    let cancelled = false;
    setGroupEmailsLoading(true);
    getGroupEmails(selectedGroupId)
      .then((rows) => {
        if (cancelled) return;
        setGroupEmails(Array.isArray(rows) ? rows : []);
        setSelectedGroupEmailKeys([]);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setGroupEmails([]);
        setMsg(error?.message || "Nao foi possivel carregar os emails do grupo.");
      })
      .finally(() => {
        if (!cancelled) setGroupEmailsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken, selectedGroupId, setMsg]);

  useEffect(() => {
    let cancelled = false;
    setLibraryLoading(true);
    searchKnownEmails(libraryQuery, { excludeGroupId: selectedGroupId || undefined, limit: 120 })
      .then((rows) => {
        if (cancelled) return;
        setLibraryEmails(Array.isArray(rows) ? rows : []);
        setSelectedLibraryKeys([]);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setLibraryEmails([]);
        setMsg(error?.message || "Nao foi possivel pesquisar emails registados.");
      })
      .finally(() => {
        if (!cancelled) setLibraryLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [libraryQuery, reloadToken, selectedGroupId, setMsg]);

  useEffect(() => {
    if (!groupTicketsEnabled) {
      setTicketSeries([]);
      setSelectedTicketSeriesId("");
      return;
    }
    let cancelled = false;
    setTicketSeriesLoading(true);
    listGroupTicketSeries()
      .then((rows) => {
        if (cancelled) return;
        const nextRows = Array.isArray(rows) ? rows : [];
        setTicketSeries(nextRows);
        setSelectedTicketSeriesId((current) => {
          if (current && nextRows.some((row) => row.id === current)) return current;
          return nextRows.find((row) => row.isActive !== false)?.id || nextRows[0]?.id || "";
        });
      })
      .catch((error: any) => {
        if (cancelled) return;
        setTicketSeries([]);
        setMsg(error?.message || "Nao foi possivel carregar as series de tickets.");
      })
      .finally(() => {
        if (!cancelled) setTicketSeriesLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [groupTicketsEnabled, reloadToken, setMsg]);

  useEffect(() => {
    setTicketSeriesDraft(createTicketSeriesDraft(selectedTicketSeries));
  }, [selectedTicketSeries]);

  useEffect(() => {
    setTicketUiDraft(createTicketUiDraft(settings?.groupTicketUi));
  }, [
    settings?.groupTicketUi?.aiInstructions,
    settings?.groupTicketUi?.autoLinkMode,
    settings?.groupTicketUi?.suggestDraftOnCreate,
    settings?.groupTicketUi?.useAiDrafts,
  ]);

  useEffect(() => {
    if (!groupTicketsEnabled || !(currentEmailLinkPayload.itemId || currentEmailLinkPayload.internetMessageId || currentEmailLinkPayload.conversationId)) {
      setCurrentEmailTickets([]);
      return;
    }
    let cancelled = false;
    searchGroupTickets({
      email: currentEmailLinkPayload,
      limit: 20,
    })
      .then((rows) => {
        if (cancelled) return;
        setCurrentEmailTickets(Array.isArray(rows) ? rows : []);
      })
      .catch((error: any) => {
        if (!cancelled) setMsg(error?.message || "Nao foi possivel carregar os tickets do email atual.");
      });
    return () => {
      cancelled = true;
    };
  }, [currentEmailKey, currentEmailLinkPayload, groupTicketsEnabled, reloadToken, setMsg]);

  useEffect(() => {
    if (!groupTicketsEnabled || !(currentEmailLinkPayload.subject || currentEmailLinkPayload.bodyText || currentEmailLinkPayload.bodyHtml)) {
      setTicketMatches([]);
      return;
    }
    let cancelled = false;
    setTicketDetectionLoading(true);
    detectGroupTicketsForEmail({ email: currentEmailLinkPayload })
      .then(async (matches) => {
        if (cancelled) return;
        const nextMatches = Array.isArray(matches) ? matches : [];
        setTicketMatches(nextMatches);
        setTicketMatchGroupSelection((current) => {
          const next = { ...current };
          for (const match of nextMatches) {
            if (!next[match.ticket.id]?.length) {
              next[match.ticket.id] = (match.proposedGroups || []).map((group) => group.id);
            }
          }
          return next;
        });

        if (ticketUi?.autoLinkMode !== "auto") return;
        const autoMatches = nextMatches.filter((match) => !match.emailLinked);
        if (!autoMatches.length) return;
        for (const match of autoMatches) {
          await linkEmailToGroupTicket(match.ticket.id, {
            email: currentEmailLinkPayload,
            applyGroups: true,
            groupIds: (match.proposedGroups || []).map((group) => group.id),
            membershipKind: "referencia",
          });
        }
        if (!cancelled) {
          setMsg(`${autoMatches.length} ticket(s) detetado(s) e ligado(s) automaticamente ao email atual.`);
          setReloadToken((value) => value + 1);
        }
      })
      .catch((error: any) => {
        if (!cancelled) {
          setTicketMatches([]);
          setMsg(error?.message || "Nao foi possivel detetar tickets no email atual.");
        }
      })
      .finally(() => {
        if (!cancelled) setTicketDetectionLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [currentEmailKey, currentEmailLinkPayload, groupTicketsEnabled, reloadToken, setMsg, ticketUi?.autoLinkMode]);

  useEffect(() => {
    if (!groupTicketsEnabled) {
      setTicketSearchResults([]);
      return;
    }
    let cancelled = false;
    setTicketSearchLoading(true);
    searchGroupTickets({
      q: ticketSearchQuery,
      groupId: selectedGroupId || undefined,
      limit: 20,
    })
      .then((rows) => {
        if (cancelled) return;
        setTicketSearchResults(Array.isArray(rows) ? rows : []);
      })
      .catch((error: any) => {
        if (!cancelled) {
          setTicketSearchResults([]);
          setMsg(error?.message || "Nao foi possivel pesquisar tickets.");
        }
      })
      .finally(() => {
        if (!cancelled) setTicketSearchLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [groupTicketsEnabled, reloadToken, selectedGroupId, setMsg, ticketSearchQuery]);

  useEffect(() => {
    if (view !== "quicklink" || !groupTicketsEnabled) {
      setQuickLinkTicketResults([]);
      setQuickLinkTicketLoading(false);
      return;
    }
    let cancelled = false;
    setQuickLinkTicketLoading(true);
    searchGroupTickets({
      q: quickLinkTicketQuery,
      limit: 12,
    })
      .then((rows) => {
        if (!cancelled) setQuickLinkTicketResults(Array.isArray(rows) ? rows : []);
      })
      .catch((error: any) => {
        if (!cancelled) {
          setQuickLinkTicketResults([]);
          setMsg(error?.message || "Nao foi possivel pesquisar tickets para a ligacao rapida.");
        }
      })
      .finally(() => {
        if (!cancelled) setQuickLinkTicketLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [groupTicketsEnabled, quickLinkTicketQuery, setMsg, view]);

  const labelCatalogEntries = useMemo(
    () => normalizeGroupLabelCatalog(settings?.groupLabelCatalog || []),
    [settings?.groupLabelCatalog]
  );

  const allLabels = useMemo(() => {
    return mergeLabelCatalog(
      getGroupLabelCatalogLabels(labelCatalogEntries),
      groups.flatMap((group) => group.labels || [])
    );
  }, [groups, labelCatalogEntries]);

  const labelsManagerEnabled = settings?.groupLabelsManagerEnabled !== false;

  const labelUsage = useMemo(
    () =>
      allLabels.map((label) => {
        const meta = findGroupLabelCatalogEntry(labelCatalogEntries, label);
        return {
          label,
          count: groups.filter((group) =>
            (group.labels || []).some((entry) => normalizeText(entry) === normalizeText(label))
          ).length,
          categorize: meta?.categorize === true,
          hasStatus: meta?.hasStatus === true,
          status: meta?.status,
        };
      }),
    [allLabels, groups, labelCatalogEntries]
  );

  const selectedManagedLabelEntry = useMemo(
    () => findGroupLabelCatalogEntry(labelCatalogEntries, selectedManagedLabel),
    [labelCatalogEntries, selectedManagedLabel]
  );

  useEffect(() => {
    if (selectedManagedLabel && allLabels.some((label) => normalizeText(label) === normalizeText(selectedManagedLabel))) {
      return;
    }
    setSelectedManagedLabel(allLabels[0] || "");
    setRenameLabelValue(allLabels[0] || "");
  }, [allLabels, selectedManagedLabel]);

  useEffect(() => {
    setRenameLabelValue(selectedManagedLabel || "");
  }, [selectedManagedLabel]);

  useEffect(() => {
    setManagedLabelDraft(createManagedLabelDraft(selectedManagedLabelEntry));
  }, [selectedManagedLabelEntry]);

  const visibleGroups = useMemo(() => {
    const query = normalizeText(groupQuery);
    const filtered = groups
      .filter((group) => {
        if (archiveFilter === "active" && group.isArchived) return false;
        if (archiveFilter === "archived" && !group.isArchived) return false;
        if (statusFilter !== "all" && String(group.status || "em_analise") !== statusFilter) return false;
        if (activeLabelFilters.length) {
          const labels = new Set((group.labels || []).map((entry) => normalizeText(entry)));
          if (!activeLabelFilters.every((entry) => labels.has(normalizeText(entry)))) return false;
        }
        if (!query) return true;
        const haystack = [
          group.name,
          group.description,
          statusLabel(group.status),
          ...(group.labels || []),
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        return haystack.includes(query);
      });

    const sorted = filtered.sort((a, b) => {
      const favoriteDelta = Number(favoriteGroupSet.has(b.id)) - Number(favoriteGroupSet.has(a.id));
      if (favoriteDelta) return favoriteDelta;
      const aUpdated = String(a.updatedAt || a.createdAt || "");
      const bUpdated = String(b.updatedAt || b.createdAt || "");
      if (aUpdated !== bUpdated) return bUpdated.localeCompare(aUpdated);
      return String(a.name || "").localeCompare(String(b.name || ""), "pt-PT");
    });

    const hasManualFilters = statusFilter !== "all" || archiveFilter !== "active" || activeLabelFilters.length > 0;
    if (query || hasManualFilters) return sorted;
    return sorted.slice(0, 8);
  }, [activeLabelFilters, archiveFilter, favoriteGroupSet, groupQuery, groups, statusFilter]);

  const exactGroupMatch = useMemo(
    () => groups.some((group) => normalizeText(group.name) === normalizeText(groupQuery)),
    [groupQuery, groups]
  );

  const visibleGroupEmails = useMemo(() => {
    const query = normalizeText(groupEmailQuery);
    return [...groupEmails]
      .filter((email) => {
        if (!query) return true;
        return [email.subject, email.fromName, email.fromEmail].filter(Boolean).join(" ").toLowerCase().includes(query);
      })
      .sort((a, b) => String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || "")));
  }, [groupEmailQuery, groupEmails]);

  const visibleLibraryEmails = useMemo(() => {
    const rows = [...libraryEmails];
    const hasCurrent = rows.some((row) => makeEmailKey(row) === currentEmailKey);
    if (
      !hasCurrent
      && (currentEmailLinkPayload.itemId || currentEmailLinkPayload.internetMessageId || currentEmailLinkPayload.conversationId)
      && (!selectedGroupId || !groupEmails.some((row) => makeEmailKey(row) === currentEmailKey))
    ) {
      rows.unshift({
        ...currentEmailLinkPayload,
        relatedGroups: [],
        relatedRecords: [],
      });
    }
    return rows;
  }, [currentEmailKey, currentEmailLinkPayload, groupEmails, libraryEmails, selectedGroupId]);

  const selectedLibraryRows = useMemo(
    () => visibleLibraryEmails.filter((row) => selectedLibraryKeys.includes(makeEmailKey(row))),
    [selectedLibraryKeys, visibleLibraryEmails]
  );

  const selectedGroupEmailRows = useMemo(
    () => visibleGroupEmails.filter((row) => selectedGroupEmailKeys.includes(makeEmailKey(row))),
    [selectedGroupEmailKeys, visibleGroupEmails]
  );

  const currentEmailAlreadyLinked = groupEmails.some((email) => makeEmailKey(email) === currentEmailKey || isCurrentContextEmail(email, ctx));
  async function collectCurrentThreadEmails(): Promise<RelevantEmailPayload[]> {
    const collected = new Map<string, RelevantEmailPayload>();
    const register = (email: Partial<RelatedEmailEntry> | RelevantEmailPayload | null | undefined) => {
      if (!email) return;
      const normalized = stripEmailPayloadAttachmentContent(email as RelevantEmailPayload);
      const key = makeEmailKey(normalized as Partial<RelatedEmailEntry>);
      if (!key) return;
      if (currentEmailLinkPayload.conversationId && normalized.conversationId && normalized.conversationId !== currentEmailLinkPayload.conversationId) {
        return;
      }
      collected.set(key, normalized);
    };

    register(currentEmailLinkPayload);

    if (!currentEmailLinkPayload.conversationId && !currentEmailLinkPayload.internetMessageId && !currentEmailLinkPayload.itemId) {
      return Array.from(collected.values());
    }

    try {
      const related = await getRelatedEmailContext(currentEmailLinkPayload);
      for (const email of related.emails || []) register(email);
      if (related.email) register(related.email);
    } catch {
      // best effort; current email still links
    }

    return Array.from(collected.values());
  }

  async function syncCurrentEmailOutlookCategories() {
    const related = await getRelatedEmailContext(currentEmailLinkPayload);
    const knownLabelNames = Array.from(new Set([
      ...(Array.isArray(settings?.groupLabelCatalog) ? settings.groupLabelCatalog.map((entry) => String(entry?.label || "").trim()).filter(Boolean) : []),
      ...(Array.isArray(related?.email?.labels) ? related.email.labels.map((label) => String(label || "").trim()).filter(Boolean) : []),
      ...(Array.isArray(related?.email?.removedInheritedLabels) ? related.email.removedInheritedLabels.map((label) => String(label || "").trim()).filter(Boolean) : []),
      ...(Array.isArray(related?.email?.classificationMeta?.categorizedLabelNames) ? related.email.classificationMeta.categorizedLabelNames.map((label) => String(label || "").trim()).filter(Boolean) : []),
    ]));
    const snapshot = await getManagedOutlookCategorySnapshot(knownLabelNames).catch(() => null);
    await syncOutlookCategorySource(buildOutlookCategorySourceFromRelatedContext({
      email: related?.email || null,
      groups: Array.isArray(related?.groups) ? related.groups : [],
      tickets: Array.isArray(related?.tickets) ? related.tickets : [],
      settings,
      currentOutlookLabelNames: snapshot?.labelNames || [],
    }));
  }

  const rankedGroups = useMemo(
    () =>
      [...groups].sort((a, b) => {
        const favoriteDelta = Number(favoriteGroupSet.has(b.id)) - Number(favoriteGroupSet.has(a.id));
        if (favoriteDelta) return favoriteDelta;
        const aUpdated = String(a.updatedAt || a.createdAt || "");
        const bUpdated = String(b.updatedAt || b.createdAt || "");
        if (aUpdated !== bUpdated) return bUpdated.localeCompare(aUpdated);
        return String(a.name || "").localeCompare(String(b.name || ""), "pt-PT");
      }),
    [favoriteGroupSet, groups]
  );

  const quickLinkTicketCatalog = useMemo(() => {
    const map = new Map<string, GroupTicketEntry>();
    for (const ticket of [
      ...currentEmailTickets,
      ...ticketMatches.map((entry) => entry.ticket),
      ...quickLinkTicketResults,
      ...ticketSearchResults,
    ]) {
      const id = String(ticket?.id || "").trim();
      if (!id) continue;
      map.set(id, ticket);
    }
    return Array.from(map.values());
  }, [currentEmailTickets, quickLinkTicketResults, ticketMatches, ticketSearchResults]);

  const quickLinkSelectedTicket = useMemo(
    () => quickLinkTicketCatalog.find((ticket) => ticket.id === quickLinkDraft.ticketId) || null,
    [quickLinkDraft.ticketId, quickLinkTicketCatalog]
  );

  const quickLinkLabels = useMemo(
    () => parseLabels(quickLinkDraft.labelsText),
    [quickLinkDraft.labelsText]
  );

  const quickLinkPrimaryCandidates = useMemo(() => {
    const query = normalizeText(quickLinkPrimaryQuery);
    return rankedGroups
      .filter((group) => !group.isArchived)
      .filter((group) => {
        if (!query) return true;
        const haystack = [group.name, group.description, ...(group.labels || [])].filter(Boolean).join(" ").toLowerCase();
        return haystack.includes(query);
      })
      .slice(0, 8);
  }, [quickLinkPrimaryQuery, rankedGroups]);

  const quickLinkSecondaryCandidates = useMemo(() => {
    const query = normalizeText(quickLinkSecondaryQuery);
    return rankedGroups
      .filter((group) => !group.isArchived)
      .filter((group) => group.id !== quickLinkDraft.principalGroupId)
      .filter((group) => !quickLinkDraft.secondaryGroupIds.includes(group.id))
      .filter((group) => {
        if (!query) return true;
        const haystack = [group.name, group.description, ...(group.labels || [])].filter(Boolean).join(" ").toLowerCase();
        return haystack.includes(query);
      })
      .slice(0, 8);
  }, [quickLinkDraft.principalGroupId, quickLinkDraft.secondaryGroupIds, quickLinkSecondaryQuery, rankedGroups]);

  async function refreshAll() {
    setReloadToken((value) => value + 1);
  }

  async function persistLabelCatalog(nextCatalog: GroupLabelCatalogEntry[]) {
    const normalizedCatalog = normalizeGroupLabelCatalog(nextCatalog);
    const hasCategorizedLabels = normalizedCatalog.some((entry) => entry.categorize === true);
    await saveSettings({
      groupLabelCatalog: normalizedCatalog,
      ...(hasCategorizedLabels
        ? {
            groupOutlookCategories: {
              enabled: settings?.groupOutlookCategories?.enabled === true,
              includeGroups: settings?.groupOutlookCategories?.includeGroups !== false,
              includeTickets: settings?.groupOutlookCategories?.includeTickets !== false,
              includeStatuses: settings?.groupOutlookCategories?.includeStatuses !== false,
              includeLabels: true,
            },
          }
        : {}),
    });
  }

  async function ensureCatalogLabels(labels: string[]): Promise<string[]> {
    if (!labelsManagerEnabled) return parseLabels(labels);
    const mergedCatalog = normalizeGroupLabelCatalog([
      ...labelCatalogEntries,
      ...labels.map((label) => buildGroupLabelCatalogEntry(label)),
    ]);
    await persistLabelCatalog(mergedCatalog);
    return canonicalizeLabelsWithCatalog(labels, getGroupLabelCatalogLabels(mergedCatalog));
  }

  async function createCatalogLabel(rawValue: string, draft?: Partial<ManagedLabelDraft> | null): Promise<string | null> {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return null;
    }
    const label = String(rawValue || "").trim();
    if (!label) {
      setMsg("Escreve um nome para a nova etiqueta.");
      return null;
    }
    const nextCatalog = normalizeGroupLabelCatalog([
      ...labelCatalogEntries,
      buildGroupLabelCatalogEntry(label, draft),
    ]);
    await persistLabelCatalog(nextCatalog);
    return nextCatalog.find((entry) => normalizeText(entry.label) === normalizeText(label))?.label || label;
  }

  function openQuickLinkView() {
    const suggestedTicket = currentEmailTickets[0] || ticketMatches[0]?.ticket || null;
    setQuickLinkDraft(
      suggestedTicket
        ? createQuickLinkDraftFromTicket(suggestedTicket)
        : createQuickLinkDraft({
          principalGroupId: selectedGroupId || "",
        })
    );
    setQuickLinkTicketQuery("");
    setQuickLinkPrimaryQuery("");
    setQuickLinkSecondaryQuery("");
    setView("quicklink");
  }

  function applyQuickLinkTicket(ticket: GroupTicketEntry) {
    setQuickLinkDraft((current) => {
      const next = createQuickLinkDraftFromTicket(ticket);
      return {
        ...next,
        ticketSeriesId: current.ticketSeriesId || ticket.seriesId || "",
      };
    });
    setQuickLinkTicketQuery(ticket.code || "");
  }

  function clearQuickLinkTicket() {
    setQuickLinkDraft((current) =>
      createQuickLinkDraft({
        ...current,
        ticketMode: "none",
        ticketId: "",
        ticketStatus: "open",
      })
    );
    setQuickLinkTicketQuery("");
  }

  function enableQuickLinkNewTicket() {
    setQuickLinkDraft((current) =>
      createQuickLinkDraft({
        ...current,
        ticketMode: "new",
        ticketId: "",
        ticketSeriesId: current.ticketSeriesId || selectedTicketSeriesId || ticketSeries.find((entry) => entry.isActive !== false)?.id || "",
        ticketStatus: "open",
      })
    );
  }

  function setQuickLinkPrimaryGroup(groupId: string) {
    const normalizedId = String(groupId || "").trim();
    setQuickLinkDraft((current) =>
      createQuickLinkDraft({
        ...current,
        principalGroupId: normalizedId,
        secondaryGroupIds: current.secondaryGroupIds.filter((entry) => entry !== normalizedId),
      })
    );
    setQuickLinkPrimaryQuery("");
  }

  function toggleQuickLinkSecondaryGroup(groupId: string) {
    const normalizedId = String(groupId || "").trim();
    if (!normalizedId) return;
    setQuickLinkDraft((current) => {
      const nextSecondary = current.secondaryGroupIds.includes(normalizedId)
        ? current.secondaryGroupIds.filter((entry) => entry !== normalizedId)
        : [...current.secondaryGroupIds, normalizedId];
      return createQuickLinkDraft({
        ...current,
        secondaryGroupIds: nextSecondary.filter((entry) => entry !== current.principalGroupId),
      });
    });
    setQuickLinkSecondaryQuery("");
  }

  async function createQuickLinkGroupFromQuery(rawValue: string, kind: "principal" | "secondary") {
    const name = String(rawValue || "").trim();
    if (!name) {
      setMsg("Escreve um nome para criar o grupo.");
      return;
    }
    const existing = groups.find((group) => normalizeText(group.name) === normalizeText(name));
    if (existing) {
      if (kind === "principal") setQuickLinkPrimaryGroup(existing.id);
      else toggleQuickLinkSecondaryGroup(existing.id);
      return;
    }
    setBusy(true);
    try {
      const group = await createLinkGroup({
        name,
        description: "",
        labels: [],
        documentsEnabled: true,
        isArchived: false,
      });
      setGroups((current) => [group, ...current.filter((entry) => entry.id !== group.id)]);
      if (kind === "principal") setQuickLinkPrimaryGroup(group.id);
      else toggleQuickLinkSecondaryGroup(group.id);
      setMsg(`Grupo ${group.name} criado.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar o grupo na ligacao rapida.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveQuickLink() {
    const principalGroupId = String(quickLinkDraft.principalGroupId || "").trim();
    const secondaryGroupIds = Array.from(
      new Set(
        quickLinkDraft.secondaryGroupIds
          .map((entry) => String(entry || "").trim())
          .filter((entry) => entry && entry !== principalGroupId)
      )
    );
    const groupIds = [principalGroupId, ...secondaryGroupIds].filter(Boolean);
    const labels = parseLabels(quickLinkDraft.labelsText);

    if (!groupIds.length && quickLinkDraft.ticketMode !== "existing") {
      setMsg("Seleciona pelo menos um grupo ou usa um ticket existente.");
      return;
    }
    if (quickLinkDraft.ticketMode === "new" && !quickLinkDraft.ticketSeriesId) {
      setMsg("Seleciona uma serie para criar o ticket.");
      return;
    }

    setBusy(true);
    try {
      const threadEmails = await collectCurrentThreadEmails();
      if (principalGroupId) {
        for (const email of threadEmails) {
          await addEmailToLinkGroup(principalGroupId, {
            ...email,
            membershipKind: "principal",
          });
        }

        if (labels.length) {
          const principalGroup = groups.find((entry) => entry.id === principalGroupId);
          if (principalGroup) {
            const mergedLabels = await ensureCatalogLabels(mergeLabelCatalog(principalGroup.labels || [], labels));
            await updateLinkGroup(principalGroupId, {
              name: principalGroup.name,
              description: principalGroup.description,
              status: principalGroup.status || "",
              labels: mergedLabels,
              documentsEnabled: principalGroup.documentsEnabled !== false,
              isArchived: principalGroup.isArchived === true,
            });
          }
        }
      }

      for (const groupId of secondaryGroupIds) {
        for (const email of threadEmails) {
          await addEmailToLinkGroup(groupId, {
            ...email,
            membershipKind: "referencia",
          });
        }
      }

      let finalTicket: GroupTicketEntry | null = null;
      if (quickLinkDraft.ticketMode === "existing" && quickLinkDraft.ticketId) {
        finalTicket = await updateGroupTicket(quickLinkDraft.ticketId, {
          labels,
          groupIds,
          status: quickLinkDraft.ticketStatus,
        });
        finalTicket = (await linkEmailToGroupTicket(quickLinkDraft.ticketId, {
          email: currentEmailLinkPayload,
          applyGroups: false,
          groupIds,
        })).ticket;
        for (const email of threadEmails.filter((entry) => makeEmailKey(entry as Partial<RelatedEmailEntry>) !== currentEmailKey)) {
          await linkEmailToGroupTicket(quickLinkDraft.ticketId, {
            email,
            applyGroups: false,
            groupIds,
          });
        }
      } else if (quickLinkDraft.ticketMode === "new" && quickLinkDraft.ticketSeriesId) {
        finalTicket = await createGroupTicket({
          seriesId: quickLinkDraft.ticketSeriesId,
          title: String(currentEmailPayload.subject || "").trim() || "Ticket",
          description: String(currentEmailPayload.bodyText || "").trim().slice(0, 4000),
          labels,
          groupIds,
        });
        if ((quickLinkDraft.ticketStatus || "open") !== "open") {
          finalTicket = await updateGroupTicket(finalTicket.id, {
            labels,
            groupIds,
            status: quickLinkDraft.ticketStatus,
          });
        }
        finalTicket = (await linkEmailToGroupTicket(finalTicket.id, {
          email: currentEmailLinkPayload,
          applyGroups: false,
          groupIds,
        })).ticket;
        for (const email of threadEmails.filter((entry) => makeEmailKey(entry as Partial<RelatedEmailEntry>) !== currentEmailKey)) {
          await linkEmailToGroupTicket(finalTicket!.id, {
            email,
            applyGroups: false,
            groupIds,
          });
        }
      }

      if (principalGroupId) {
        setSelectedGroupId(principalGroupId);
        setActiveGroupForCurrentEmail(principalGroupId);
        setView("detail");
      } else {
        setView("groups");
      }
      await refreshAll();
      if (finalTicket || groupIds.length) {
        await syncCurrentEmailOutlookCategories();
      }
      if (finalTicket && quickLinkDraft.ticketMode === "new" && ticketUi?.suggestDraftOnCreate) {
        await handleOpenTicketReplyDraft(finalTicket);
      }
      setMsg(
        finalTicket
          ? `${threadEmails.length} email(s) da thread ligados com ticket ${finalTicket.code}.`
          : `${threadEmails.length} email(s) da thread ligados a ${groupIds.length} grupo(s).`
      );
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel guardar a ligacao rapida.");
    } finally {
      setBusy(false);
    }
  }

  async function handleCreateGroup() {
    const name = String(groupQuery || "").trim();
    if (!name) {
      setMsg("Escreve um nome para pesquisar ou criar o grupo.");
      return;
    }
    setBusy(true);
    try {
      const group = await createLinkGroup({
        name,
        labels: [],
        documentsEnabled: true,
      });
      setGroupQuery("");
      setSelectedGroupId(group.id);
      setView("detail");
      await refreshAll();
      setMsg("Grupo criado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveGroup() {
    if (!selectedGroup) return;
    const name = String(draft.name || "").trim();
    if (!name) {
      setMsg("O grupo precisa de um nome.");
      return;
    }
    setBusy(true);
    try {
      const labels = await ensureCatalogLabels(parseLabels(draft.labelsText));
      await updateLinkGroup(selectedGroup.id, {
        name,
        description: draft.description,
        status: draft.status || "",
        labels,
        documentsEnabled: draft.documentsEnabled,
        isArchived: draft.isArchived,
      });
      await refreshAll();
      setMsg("Grupo atualizado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleDeleteGroup() {
    if (!selectedGroup) return;
    if (!confirmAction(`Eliminar o grupo "${selectedGroup.name}" e todas as ligacoes?`)) return;
    setBusy(true);
    try {
      await deleteLinkGroup(selectedGroup.id);
      setSelectedGroupId("");
      setView("groups");
      await refreshAll();
      setMsg("Grupo eliminado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel eliminar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleAddCurrentEmail() {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      const threadEmails = await collectCurrentThreadEmails();
      for (const email of threadEmails) {
        await addEmailToLinkGroup(selectedGroup.id, {
          ...email,
          membershipKind: linkKind,
        });
      }
      setActiveGroupForCurrentEmail(selectedGroup.id);
      await refreshAll();
      await syncCurrentEmailOutlookCategories();
      setMsg(
        threadEmails.length > 1
          ? `${threadEmails.length} email(s) da thread associados ao grupo como ${membershipKindLabel(linkKind).toLowerCase()}.`
          : `Email atual associado ao grupo como ${membershipKindLabel(linkKind).toLowerCase()}.`
      );
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel associar o email atual.");
    } finally {
      setBusy(false);
    }
  }

  async function handleAddSelectedLibrary() {
    if (!selectedGroup || !selectedLibraryRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedLibraryRows) {
        await addEmailToLinkGroup(selectedGroup.id, {
          ...stripEmailPayloadAttachmentContent(email),
          membershipKind: linkKind,
        });
      }
      setSelectedLibraryKeys([]);
      await refreshAll();
      setMsg(`${selectedLibraryRows.length} email(s) associados ao grupo como ${membershipKindLabel(linkKind).toLowerCase()}.`);
      setView("detail");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel associar os emails selecionados.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveCurrentAttachmentsToGroup() {
    if (!selectedGroup) return;
    if (selectedGroup.documentsEnabled === false) {
      setMsg("Ativa primeiro os documentos neste grupo.");
      return;
    }
    if (!currentSavableAttachments.length) {
      setMsg("O email atual nao tem anexos guardaveis neste momento.");
      return;
    }
    if (!selectedCurrentAttachments.length) {
      setMsg("Seleciona pelo menos um anexo para guardar no grupo.");
      return;
    }
    setBusy(true);
    try {
      const storageProvider = String(settings?.groupStorage.provider || "cloud").trim();
      const storageBasePath = String(settings?.groupStorage.baseFolderPath || "").trim();
      const docs = selectedCurrentAttachments.map((attachment) => ({
        id: `doc_${globalThis.crypto?.randomUUID?.() || `${Date.now()}_${attachment.id || attachment.name}`}`,
        name: attachment.name,
        contentType: attachment.contentType,
        contentBase64: attachment.content,
        size: attachment.size,
        sourceEmailKey: currentEmailKey,
        sourceItemId: String(ctx.itemId || "").trim(),
        sourceInternetMessageId: String(ctx.internetMessageId || "").trim(),
        sourceConversationId: String(ctx.conversationId || "").trim(),
        sourceEmailSubject: String(ctx.subject || "").trim(),
        storageProvider,
        storageBasePath,
      }));
      const batches = splitDocumentsIntoBatches(docs);
      for (const batch of batches) {
        await saveGroupDocuments(selectedGroup.id, { documents: batch });
      }
      await refreshAll();
      setMsg(
        batches.length > 1
          ? `${docs.length} anexo(s) guardado(s) nos documentos do grupo em ${batches.length} lotes.`
          : `${docs.length} anexo(s) guardado(s) nos documentos do grupo.`
      );
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel guardar os anexos deste email no grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleCreateTicketSeries() {
    const name = String(newTicketSeriesDraft.name || "").trim();
    const prefix = normalizeTicketPrefixInput(newTicketSeriesDraft.prefix);
    if (!name || !prefix) {
      setMsg("Define nome e prefixo para a nova serie.");
      return;
    }
    setBusy(true);
    try {
      const series = await createGroupTicketSeries({
        name,
        prefix,
        replyInstructions: String(newTicketSeriesDraft.replyInstructions || "").trim(),
        yearMode: newTicketSeriesDraft.yearMode,
        separator: newTicketSeriesDraft.separator,
        nextNumber: Math.max(1, Number(newTicketSeriesDraft.nextNumber || 1) || 1),
        padding: Math.max(2, Number(newTicketSeriesDraft.padding || 4) || 4),
        isActive: newTicketSeriesDraft.isActive,
      });
      setTicketSeries((current) =>
        [...current.filter((entry) => entry.id !== series.id), series].sort((a, b) =>
          Number(b.isActive !== false) - Number(a.isActive !== false)
          || String(a.prefix || "").localeCompare(String(b.prefix || ""), "pt-PT")
          || String(a.name || "").localeCompare(String(b.name || ""), "pt-PT")
        )
      );
      setSelectedTicketSeriesId(series.id);
      setNewTicketSeriesDraft({ name: "", prefix: "", replyInstructions: "", yearMode: "none", separator: "-", nextNumber: "1", padding: "4", isActive: true });
      setReloadToken((value) => value + 1);
      setMsg(`Serie ${series.prefix} criada.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar a serie.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveTicketSeries() {
    if (!selectedTicketSeries) return;
    const name = String(ticketSeriesDraft.name || "").trim();
    const prefix = normalizeTicketPrefixInput(ticketSeriesDraft.prefix);
    if (!name || !prefix) {
      setMsg("Define nome e prefixo para a serie.");
      return;
    }
    setBusy(true);
    try {
      const updated = await updateGroupTicketSeries(selectedTicketSeries.id, {
        name,
        prefix,
        replyInstructions: String(ticketSeriesDraft.replyInstructions || "").trim(),
        yearMode: ticketSeriesDraft.yearMode,
        separator: ticketSeriesDraft.separator,
        nextNumber: Math.max(1, Number(ticketSeriesDraft.nextNumber || 1) || 1),
        padding: Math.max(2, Number(ticketSeriesDraft.padding || 4) || 4),
        isActive: ticketSeriesDraft.isActive,
      });
      setTicketSeries((current) =>
        current
          .map((entry) => (entry.id === updated.id ? updated : entry))
          .sort((a, b) =>
            Number(b.isActive !== false) - Number(a.isActive !== false)
            || String(a.prefix || "").localeCompare(String(b.prefix || ""), "pt-PT")
            || String(a.name || "").localeCompare(String(b.name || ""), "pt-PT")
          )
      );
      setTicketSeriesDraft(createTicketSeriesDraft(updated));
      setSelectedTicketSeriesId(updated.id);
      setReloadToken((value) => value + 1);
      setMsg(`Serie ${updated.prefix} atualizada.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar a serie.");
    } finally {
      setBusy(false);
    }
  }

  async function handleResetTicketSeriesCounter() {
    if (!selectedTicketSeries) return;
    setBusy(true);
    try {
      const updated = await updateGroupTicketSeries(selectedTicketSeries.id, {
        nextNumber: 1,
      });
      setTicketSeries((current) => current.map((entry) => (entry.id === updated.id ? updated : entry)));
      setTicketSeriesDraft(createTicketSeriesDraft(updated));
      setSelectedTicketSeriesId(updated.id);
      setReloadToken((value) => value + 1);
      setMsg(`Contador da serie ${updated.prefix} reiniciado.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel reiniciar o contador.");
    } finally {
      setBusy(false);
    }
  }

  async function handleDeleteTicketSeries() {
    if (!selectedTicketSeries) return;
    if (!confirmAction(`Eliminar a serie "${selectedTicketSeries.name}"?`)) return;
    setBusy(true);
    try {
      await deleteGroupTicketSeries(selectedTicketSeries.id);
      setTicketSeries((current) => current.filter((entry) => entry.id !== selectedTicketSeries.id));
      setSelectedTicketSeriesId("");
      setReloadToken((value) => value + 1);
      setMsg("Serie eliminada.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel eliminar a serie.");
    } finally {
      setBusy(false);
    }
  }

  async function handleCreateTicketFromCurrentEmail() {
    if (!groupTicketsEnabled) {
      setMsg("Ativa primeiro os tickets em Settings > Grupos.");
      return;
    }
    if (!selectedTicketSeriesId) {
      setMsg("Seleciona primeiro uma serie.");
      return;
    }
    setBusy(true);
    try {
      const threadEmails = await collectCurrentThreadEmails();
      const ticket = await createGroupTicket({
        seriesId: selectedTicketSeriesId,
        title: String(ctx.subject || "").trim() || `Ticket ${selectedGroup?.name || ""}`.trim() || "Ticket",
        description: String(bodyText || "").trim().slice(0, 2000),
        labels: selectedGroup?.labels || [],
        groupIds: selectedGroup ? [selectedGroup.id] : [],
        email: currentEmailLinkPayload,
        membershipKind: selectedGroup ? linkKind : "referencia",
      });
      for (const email of threadEmails.filter((entry) => makeEmailKey(entry as Partial<RelatedEmailEntry>) !== currentEmailKey)) {
        await linkEmailToGroupTicket(ticket.id, {
          email,
          applyGroups: true,
          groupIds: selectedGroup ? [selectedGroup.id] : [],
          membershipKind: selectedGroup ? linkKind : "referencia",
        });
      }
      await syncCurrentEmailOutlookCategories();
      setReloadToken((value) => value + 1);
      setMsg(
        threadEmails.length > 1
          ? `Ticket ${ticket.code} criado e ligado a ${threadEmails.length} email(s) da thread.`
          : `Ticket ${ticket.code} criado.`
      );
      if (ticketUi?.suggestDraftOnCreate) {
        await handleOpenTicketReplyDraft(ticket);
      }
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar o ticket.");
    } finally {
      setBusy(false);
    }
  }

  async function handleLinkCurrentEmailToTicket(ticket: GroupTicketEntry, groupIdsOverride?: string[]) {
    setBusy(true);
    try {
      const threadEmails = await collectCurrentThreadEmails();
      const ensuredGroupIds = Array.from(
        new Set([
          ...(groupIdsOverride || []),
          ...(selectedGroup ? [selectedGroup.id] : []),
        ].filter(Boolean))
      );
      const result = await linkEmailToGroupTicket(ticket.id, {
        email: currentEmailLinkPayload,
        applyGroups: true,
        groupIds: ensuredGroupIds,
        membershipKind: selectedGroup ? linkKind : "referencia",
      });
      for (const email of threadEmails.filter((entry) => makeEmailKey(entry as Partial<RelatedEmailEntry>) !== currentEmailKey)) {
        await linkEmailToGroupTicket(ticket.id, {
          email,
          applyGroups: true,
          groupIds: ensuredGroupIds,
          membershipKind: selectedGroup ? linkKind : "referencia",
        });
      }
      await syncCurrentEmailOutlookCategories();
      setReloadToken((value) => value + 1);
      const groupSummary = result.appliedGroups.length ? ` ${result.appliedGroups.length} grupo(s) atualizados.` : "";
      setMsg(
        threadEmails.length > 1
          ? `${threadEmails.length} email(s) da thread ligados ao ticket ${ticket.code}.${groupSummary}`
          : `Email atual ligado ao ticket ${ticket.code}.${groupSummary}`
      );
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel ligar o email ao ticket.");
    } finally {
      setBusy(false);
    }
  }

  function toggleTicketMatchGroup(ticketId: string, groupId: string) {
    setTicketMatchGroupSelection((current) => {
      const selected = current[ticketId] || [];
      return {
        ...current,
        [ticketId]: selected.includes(groupId)
          ? selected.filter((entry) => entry !== groupId)
          : [...selected, groupId],
      };
    });
  }

  async function handleConfirmDetectedTicket(match: GroupTicketDetectionMatch) {
    const selectedGroupIds = Array.from(
      new Set([
        ...(ticketMatchGroupSelection[match.ticket.id] || (match.proposedGroups || []).map((group) => group.id)),
        ...(selectedGroup ? [selectedGroup.id] : []),
      ].filter(Boolean))
    );
    await handleLinkCurrentEmailToTicket(match.ticket, selectedGroupIds);
  }

  async function handleOpenTicketDraft(ticket: GroupTicketEntry) {
    const relatedSeries = ticketSeries.find((series) => series.id === ticket.seriesId) || selectedTicketSeries || null;
    const toRecipients = uniqueEmails([String(ctx.fromEmail || "").trim()]);
    const ccRecipients = uniqueEmails((ctx.ccRecipients || []).map((entry: any) => entry?.email));
    const subject = buildTicketEmailSubject(
      String(ctx.subject || "").trim() || `Ticket ${ticket.code}`,
      ticket.code,
      ticketUi?.includeTicketCodeInSubject !== false,
    );
    let body = [
      `<p>Foi aberto o ticket <strong>${ticket.code}</strong>.</p>`,
      `<p>Nas próximas respostas, pedimos que mantenham este número no assunto do email.</p>`,
    ].join("");

    if (ticketUi?.useAiDrafts) {
      try {
        const response = await aiGenerate({
          action: "reply",
          mode: "quality",
          locale: (settings?.replyLanguage || "pt-PT") as any,
          tone: "formal",
          email: {
            subject: String(ctx.subject || "").trim(),
            from: String(ctx.fromEmail || "").trim(),
            to: (ctx.toRecipients || []).map((entry) => entry.email),
            cc: (ctx.ccRecipients || []).map((entry) => entry.email),
            bodyText: String(bodyText || "").trim(),
            bodyScope: "full",
          },
          inputText: [
            `Redige um email de abertura/atualizacao do ticket ${ticket.code}.`,
            "Resume brevemente o tema do email original.",
            `Pede explicitamente que todas as respostas futuras incluam ${ticket.code} no assunto.`,
            String(ticketUi?.aiInstructions || "").trim(),
          ].filter(Boolean).join("\n"),
          knowledge: settings?.aiKnowledge || [],
        } as any);
        if ((response as any)?.ok) {
          body = String((response as any).html || "").trim() || `<p>${String((response as any).text || "").trim().replace(/\n/g, "<br/>")}</p>`;
        }
      } catch {
        // fallback static body
      }
    }

    await displayNewMessageForm({
      toRecipients,
      ccRecipients,
      subject,
      body,
      isHtml: true,
    });
    setMsg(`Draft do ticket ${ticket.code} aberto para envio.`);
  }

  async function handleOpenTicketReplyDraft(ticket: GroupTicketEntry) {
    const relatedSeries = ticketSeries.find((series) => series.id === ticket.seriesId) || selectedTicketSeries || null;
    const replySubject = buildTicketEmailSubject(
      String(ctx.subject || "").trim() || `Ticket ${ticket.code}`,
      ticket.code,
      ticketUi?.includeTicketCodeInSubject !== false,
    );
    let body = [
      `<p>Foi aberto o ticket <strong>${ticket.code}</strong>.</p>`,
      `<p>Agradecemos o seu email. Vamos dar seguimento ao assunto com a maior brevidade possivel.</p>`,
      `<p>Nas proximas respostas, pedimos que mantenham este numero no assunto do email.</p>`,
    ].join("");

    if (ticketUi?.useAiDrafts) {
      try {
        const response = await aiGenerate({
          action: "reply",
          mode: "quality",
          locale: (settings?.replyLanguage || "pt-PT") as any,
          tone: "formal",
          email: {
            subject: String(ctx.subject || "").trim(),
            from: String(ctx.fromEmail || "").trim(),
            to: (ctx.toRecipients || []).map((entry) => entry.email),
            cc: (ctx.ccRecipients || []).map((entry) => entry.email),
            bodyText: String(bodyText || "").trim(),
            bodyScope: "full",
          },
          inputText: [
            `Redige uma resposta ao email original para abertura ou atualizacao do ticket ${ticket.code}.`,
            "A resposta deve manter o contexto da conversa e soar como reply ao email recebido.",
            "Agradece o contacto e confirma a rececao do email.",
            buildSeriesReplyGuidance(relatedSeries, ticket),
            "Resume brevemente o tema do email original sem copiar o email.",
            `Pede explicitamente que todas as respostas futuras incluam ${ticket.code} no assunto.`,
            String(ticketUi?.aiInstructions || "").trim(),
          ].filter(Boolean).join("\n"),
          knowledge: settings?.aiKnowledge || [],
        } as any);
        if ((response as any)?.ok) {
          body = String((response as any).html || "").trim() || `<p>${String((response as any).text || "").trim().replace(/\n/g, "<br/>")}</p>`;
        }
      } catch {
        // fallback static body
      }
    }

    await displayReplyForm(body, true, { replyAll: true });
    if (ticketUi?.includeTicketCodeInSubject !== false && replySubject) {
      try {
        await setSubjectInComposeDraft(replySubject);
      } catch (error) {
        console.warn("[GroupManager] Could not update reply draft subject with ticket code:", error);
      }
    }
    setMsg(`Resposta ao ticket ${ticket.code} aberta em reply all.`);
  }

  async function handleSaveTicketUi() {
    setBusy(true);
    try {
      await saveSettings({
        groupTicketUi: {
          autoLinkMode: ticketUiDraft.autoLinkMode,
          suggestDraftOnCreate: ticketUiDraft.suggestDraftOnCreate,
          useAiDrafts: ticketUiDraft.useAiDrafts,
          includeTicketCodeInSubject: ticketUiDraft.includeTicketCodeInSubject,
          aiInstructions: String(ticketUiDraft.aiInstructions || "").trim(),
        },
      });
      setMsg("Settings locais dos tickets guardados.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel guardar os settings dos tickets.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSetSelectedEmailsKind(nextKind: MembershipKind) {
    if (!selectedGroup || !selectedGroupEmailRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedGroupEmailRows) {
        await addEmailToLinkGroup(selectedGroup.id, {
          ...stripEmailPayloadAttachmentContent(email),
          membershipKind: nextKind,
        });
      }
      await refreshAll();
      setMsg(`${selectedGroupEmailRows.length} email(s) marcados como ${membershipKindLabel(nextKind).toLowerCase()}.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar o tipo de ligacao.");
    } finally {
      setBusy(false);
    }
  }

  async function handleRemoveSelectedEmails() {
    if (!selectedGroup || !selectedGroupEmailRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedGroupEmailRows) {
        await removeEmailFromLinkGroup(selectedGroup.id, {
          ...email,
          emailKey: makeEmailKey(email),
        });
      }
      setSelectedGroupEmailKeys([]);
      await refreshAll();
      setMsg(`${selectedGroupEmailRows.length} email(s) removidos do grupo.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel remover os emails selecionados.");
    } finally {
      setBusy(false);
    }
  }

  function toggleLabelFilter(label: string) {
    setActiveLabelFilters((current) => {
      const key = normalizeText(label);
      return current.some((entry) => normalizeText(entry) === key)
        ? current.filter((entry) => normalizeText(entry) !== key)
        : [...current, label];
    });
  }

  async function toggleFavoriteGroup(groupId: string) {
    const next = favoriteGroupSet.has(groupId)
      ? favoriteGroupIds.filter((id) => id !== groupId)
      : [groupId, ...favoriteGroupIds.filter((id) => id !== groupId)];
    try {
      await saveSettings({ groupFavoriteIds: next });
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar os favoritos.");
    }
  }

  async function handleCreateCatalogLabel() {
    setBusy(true);
    try {
      const canonical = await createCatalogLabel(newCatalogLabel, newCatalogLabelDraft);
      if (!canonical) return;
      setNewCatalogLabel("");
      setNewCatalogLabelDraft(createManagedLabelDraft());
      setSelectedManagedLabel(canonical);
      setRenameLabelValue(canonical);
      setMsg("Etiqueta adicionada ao catalogo.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function handleRenameManagedLabel() {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return;
    }
    const source = String(selectedManagedLabel || "").trim();
    const target = String(renameLabelValue || "").trim();
    if (!source) {
      setMsg("Seleciona primeiro uma etiqueta.");
      return;
    }
    if (!target) {
      setMsg("Escreve o novo nome da etiqueta.");
      return;
    }
    setBusy(true);
    try {
      const sourceEntry = selectedManagedLabelEntry || buildGroupLabelCatalogEntry(source, managedLabelDraft);
      const nextCatalog = normalizeGroupLabelCatalog([
        ...labelCatalogEntries.filter((entry) => normalizeText(entry.label) !== normalizeText(source)),
        { ...sourceEntry, label: target },
      ]);
      const nextCatalogLabels = getGroupLabelCatalogLabels(nextCatalog);
      for (const group of groups) {
        const labels = group.labels || [];
        if (!labels.some((entry) => normalizeText(entry) === normalizeText(source))) continue;
        const nextLabels = canonicalizeLabelsWithCatalog(
          labels.map((entry) => (normalizeText(entry) === normalizeText(source) ? target : entry)),
          nextCatalogLabels
        );
        await updateLinkGroup(group.id, { labels: nextLabels });
      }
      await persistLabelCatalog(nextCatalog);
      setSelectedManagedLabel(target);
      setRenameLabelValue(target);
      await refreshAll();
      setMsg("Etiqueta renomeada em todos os grupos.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel renomear a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function handleDeleteManagedLabel() {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return;
    }
    const label = String(selectedManagedLabel || "").trim();
    if (!label) {
      setMsg("Seleciona primeiro uma etiqueta.");
      return;
    }
    if (!confirmAction(`Eliminar a etiqueta "${label}" de todos os grupos?`)) return;
    setBusy(true);
    try {
      for (const group of groups) {
        const labels = group.labels || [];
        if (!labels.some((entry) => normalizeText(entry) === normalizeText(label))) continue;
        await updateLinkGroup(group.id, {
          labels: labels.filter((entry) => normalizeText(entry) !== normalizeText(label)),
        });
      }
      const nextCatalog = labelCatalogEntries.filter((entry) => normalizeText(entry.label) !== normalizeText(label));
      await persistLabelCatalog(nextCatalog);
      setSelectedManagedLabel("");
      setRenameLabelValue("");
      setManagedLabelDraft(createManagedLabelDraft());
      await refreshAll();
      setMsg("Etiqueta eliminada do catalogo e dos grupos.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel eliminar a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveManagedLabelSettings() {
    const label = String(selectedManagedLabel || "").trim();
    if (!label) {
      setMsg("Seleciona primeiro uma etiqueta.");
      return;
    }
    setBusy(true);
    try {
      const nextCatalog = normalizeGroupLabelCatalog([
        ...labelCatalogEntries.filter((entry) => normalizeText(entry.label) !== normalizeText(label)),
        buildGroupLabelCatalogEntry(label, managedLabelDraft),
      ]);
      await persistLabelCatalog(nextCatalog);
      await refreshAll();
      setMsg("Configuracao da etiqueta guardada.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel guardar a configuracao da etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function openExplorer() {
    if (!selectedGroup) return;
    try {
      await openGroupExplorer({ groupId: selectedGroup.id });
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel abrir o explorer do grupo.");
    }
  }

  const trackTransform = view === "groups" ? "translateX(0%)" : view === "detail" ? "translateX(-33.3333%)" : "translateX(-66.6667%)";
  const configView = view === "settings" || view === "labels" || view === "tickets";

  const headerTitle = standaloneSettings && configView
    ? "Settings dos grupos"
    : view === "settings"
      ? "Settings dos grupos"
      : view === "labels"
        ? "Gestor de etiquetas"
        : view === "tickets"
          ? "Tickets dos grupos"
          : view === "quicklink"
            ? "Ligacao rapida"
            : "Grupos";

  const headerHint = standaloneSettings && configView
    ? "Configuracao do modulo."
    : view === "settings"
      ? "Configuracoes locais."
      : view === "labels"
        ? "Catalogo central."
        : view === "tickets"
          ? "Series, contadores e automacao."
          : view === "quicklink"
            ? "Ligacao simplificada do email atual."
            : "Ligacoes manuais.";

  const standaloneSettingsMenu = standaloneSettings ? (
    <aside style={S.standaloneSettingsSidebar}>
      <button
        type="button"
        style={view === "settings" ? S.standaloneSettingsNavItemOn : S.standaloneSettingsNavItem}
        onClick={() => setView("settings")}
      >
        General
      </button>
      <button
        type="button"
        style={view === "labels" ? S.standaloneSettingsNavItemOn : S.standaloneSettingsNavItem}
        onClick={() => setView("labels")}
      >
        Etiquetas
      </button>
      <button
        type="button"
        style={view === "tickets" ? S.standaloneSettingsNavItemOn : S.standaloneSettingsNavItem}
        onClick={() => setView("tickets")}
      >
        Tickets
      </button>
    </aside>
  ) : null;

  return (
    <div
      style={
        standaloneSettings
          ? { ...S.root, height: "100%", minHeight: 0, gridTemplateRows: "auto minmax(0, 1fr)", overflow: "hidden" }
          : S.root
      }
    >
      <div style={S.header}>
        <div>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.title}>{headerTitle}</div>
          <div style={S.headerHint}>{headerHint}</div>
        </div>
        <div style={S.headerActions}>
          {standaloneSettings ? null : configView ? (
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => setView(view === "labels" || view === "tickets" ? "settings" : "groups")}
              disabled={busy}
            >
              Voltar
            </button>
          ) : (
            <>
              <button type="button" style={S.secondaryBtn} onClick={() => void refreshAll()} disabled={groupsLoading || busy}>
                <Icons.RefreshCw size={12} />
                Atualizar
              </button>
              <button
                type="button"
                style={S.iconGearBtn}
                onClick={() => void openGroupSettings()}
                disabled={busy}
                title="Settings dos grupos"
              >
                <Icons.Settings size={14} />
              </button>
            </>
          )}
        </div>
      </div>

      <div
        style={
          standaloneSettings
            ? { ...S.viewport, height: "100%", minHeight: 0, overflowY: "auto" }
            : S.viewport
        }
      >
        {view === "settings" ? (
          <section style={standaloneSettings ? S.standaloneSettingsSection : S.cleanPanel}>
            {standaloneSettingsMenu}
            <div style={standaloneSettings ? S.standaloneSettingsContent : undefined}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Settings dos grupos</div>
                  <HelpHint text="Area limpa para extras do modulo Grupos. Aqui entras nos ecras de configuracao sem contexto do email aberto." title="Ajuda: Settings dos grupos" />
                </div>
              </div>
            </div>

            <div style={S.settingMenuGrid}>
              <button type="button" style={S.settingEntry} onClick={() => setView("labels")}>
                <div style={S.settingEntryBody}>
                  <div style={S.settingEntryTitle}>Etiquetas</div>
                </div>
                <span style={S.settingEntryMeta}>{labelsManagerEnabled ? "Ativo" : "Desativado"}</span>
              </button>

              <button type="button" style={S.settingEntry} onClick={() => setView("tickets")}>
                <div style={S.settingEntryBody}>
                  <div style={S.settingEntryTitle}>Tickets</div>
                </div>
                <span style={S.settingEntryMeta}>{groupTicketsEnabled ? "Ativo" : "Desativado"}</span>
              </button>

              <div style={S.card}>
                <div style={S.fieldLabel}>Ativacao global</div>
                <div style={S.smallMeta}>A ativacao continua em Settings {" > "} Grupos.</div>
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings gerais
                  </button>
                </div>
              </div>
            </div>
            </div>
          </section>
        ) : view === "labels" ? (
          <section style={standaloneSettings ? S.standaloneSettingsSection : S.cleanPanel}>
            {standaloneSettingsMenu}
            <div style={standaloneSettings ? S.standaloneSettingsContent : undefined}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Gestor de etiquetas</div>
                  <HelpHint text="Catalogo central de etiquetas. Cria, renomeia e remove etiquetas sem depender do email ou do grupo ativo." title="Ajuda: Gestor de etiquetas" />
                </div>
              </div>
            </div>

            {!labelsManagerEnabled ? (
              <div style={S.card}>
                <PanelState
                  compact
                  tone="info"
                  title="Gestor de etiquetas desativado"
                  description="Ativa a funcionalidade em Settings > Grupos para usar o catalogo central."
                />
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings gerais
                  </button>
                </div>
              </div>
            ) : (
              <div style={S.settingsGrid}>
                <div style={S.card}>
                  <div style={S.fieldLabel}>Nova etiqueta</div>
                  <div style={S.inlineRow}>
                    <input
                      style={S.input}
                      value={newCatalogLabel}
                      onChange={(event) => setNewCatalogLabel(event.target.value)}
                      placeholder="Ex.: marca, cliente, ganho"
                    />
                    <button type="button" style={S.primaryBtn} onClick={() => void handleCreateCatalogLabel()} disabled={busy}>
                      <Icons.Plus size={12} />
                      Criar
                    </button>
                  </div>
                  <div style={S.inlineRow}>
                    <label style={S.toggleRow}>
                      <input
                        type="checkbox"
                        checked={newCatalogLabelDraft.categorize}
                        onChange={(event) => setNewCatalogLabelDraft((current) => ({ ...current, categorize: event.target.checked }))}
                      />
                      <span>Categoria Outlook</span>
                    </label>
                    <label style={S.toggleRow}>
                      <input
                        type="checkbox"
                        checked={newCatalogLabelDraft.hasStatus}
                        onChange={(event) =>
                          setNewCatalogLabelDraft((current) => ({
                            ...current,
                            hasStatus: event.target.checked,
                            status: event.target.checked ? current.status || "em_analise" : "em_analise",
                          }))
                        }
                      />
                      <span>Tem estado</span>
                    </label>
                    {newCatalogLabelDraft.hasStatus ? (
                      <select
                        style={S.compactSelect}
                        value={newCatalogLabelDraft.status}
                        onChange={(event) =>
                          setNewCatalogLabelDraft((current) => ({ ...current, status: event.target.value as GroupLabelStatus }))
                        }
                      >
                        {LABEL_STATUS_OPTIONS.map((option) => (
                          <option key={option.value} value={option.value}>{option.label}</option>
                        ))}
                      </select>
                    ) : null}
                  </div>
                  <div style={S.smallMeta}>Catalogo central usado para evitar duplicados e normalizar etiquetas dos grupos.</div>
                </div>

                <div style={S.settingsColumns}>
                  <div style={S.managerList}>
                    <div style={S.fieldLabel}>Etiquetas existentes</div>
                    {!labelUsage.length ? (
                      <PanelState compact tone="info" title="Sem etiquetas" description="Cria a primeira etiqueta ou guarda etiquetas num grupo." />
                    ) : (
                      labelUsage.map((entry) => {
                        const active = normalizeText(entry.label) === normalizeText(selectedManagedLabel);
                        return (
                          <button
                            key={entry.label}
                            type="button"
                            style={active ? S.managerRowActive : S.managerRow}
                            onClick={() => setSelectedManagedLabel(entry.label)}
                          >
                            <div style={{ display: "grid", gap: 2, textAlign: "left" }}>
                              <span>{entry.label}</span>
                              <small style={S.managerMetaRow}>
                                {entry.categorize ? "Cat." : "Sem cat."}
                                {entry.hasStatus ? ` · ${LABEL_STATUS_OPTIONS.find((option) => option.value === entry.status)?.label || "Em analise"}` : " · Sem estado"}
                              </small>
                            </div>
                            <span style={S.managerCount}>{entry.count}</span>
                          </button>
                        );
                      })
                    )}
                  </div>

                  <div style={S.card}>
                    <div style={S.fieldLabel}>Etiqueta selecionada</div>
                    {selectedManagedLabel ? (
                      <>
                        <input
                          style={S.input}
                          value={renameLabelValue}
                          onChange={(event) => setRenameLabelValue(event.target.value)}
                          placeholder="Novo nome da etiqueta"
                        />
                        <div style={S.inlineRow}>
                          <label style={S.toggleRow}>
                            <input
                              type="checkbox"
                              checked={managedLabelDraft.categorize}
                              onChange={(event) =>
                                setManagedLabelDraft((current) => ({ ...current, categorize: event.target.checked }))
                              }
                            />
                            <span>Categoria Outlook</span>
                          </label>
                          <label style={S.toggleRow}>
                            <input
                              type="checkbox"
                              checked={managedLabelDraft.hasStatus}
                              onChange={(event) =>
                                setManagedLabelDraft((current) => ({
                                  ...current,
                                  hasStatus: event.target.checked,
                                  status: event.target.checked ? current.status || "em_analise" : "em_analise",
                                }))
                              }
                            />
                            <span>Tem estado</span>
                          </label>
                          {managedLabelDraft.hasStatus ? (
                            <select
                              style={S.compactSelect}
                              value={managedLabelDraft.status}
                              onChange={(event) =>
                                setManagedLabelDraft((current) => ({ ...current, status: event.target.value as GroupLabelStatus }))
                              }
                            >
                              {LABEL_STATUS_OPTIONS.map((option) => (
                                <option key={option.value} value={option.value}>{option.label}</option>
                              ))}
                            </select>
                          ) : null}
                        </div>
                        <div style={S.smallMeta}>As acoes abaixo atualizam todos os grupos que usam esta etiqueta.</div>
                        <div style={S.inlineRow}>
                          <button type="button" style={S.primaryBtn} onClick={() => void handleSaveManagedLabelSettings()} disabled={busy}>
                            <Icons.Save size={12} />
                            Guardar configuracao
                          </button>
                          <button type="button" style={S.primaryBtn} onClick={() => void handleRenameManagedLabel()} disabled={busy}>
                            <Icons.Save size={12} />
                            Renomear globalmente
                          </button>
                          <button type="button" style={S.dangerBtn} onClick={() => void handleDeleteManagedLabel()} disabled={busy}>
                            <Icons.Trash size={12} />
                            Eliminar globalmente
                          </button>
                        </div>
                      </>
                    ) : (
                      <PanelState compact tone="info" title="Seleciona uma etiqueta" description="Escolhe uma etiqueta da lista para a renomear ou eliminar." />
                    )}
                  </div>
                </div>
              </div>
            )}
            </div>
          </section>
        ) : view === "tickets" ? (
          <section style={standaloneSettings ? S.standaloneSettingsSection : S.cleanPanel}>
            {standaloneSettingsMenu}
            <div style={standaloneSettings ? S.standaloneSettingsContent : undefined}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Tickets dos grupos</div>
                  <HelpHint text="Ecran limpo para series, contadores e automacao local dos tickets. Nao depende do email ou do grupo aberto." title="Ajuda: Tickets dos grupos" />
                </div>
              </div>
            </div>

            {!groupTicketsEnabled ? (
              <div style={S.card}>
                <PanelState
                  compact
                  tone="info"
                  title="Tickets desativados"
                  description="Ativa a funcionalidade em Settings > Grupos para usar series e automatismos de tickets."
                />
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings gerais
                  </button>
                </div>
              </div>
            ) : (
              <div style={S.settingsGrid}>
                <div style={S.card}>
                  <div style={S.sectionTitleRow}>
                    <div style={S.fieldLabel}>Comportamento</div>
                    <HelpHint text="Escolhe se a ligacao aos tickets detetados e automatica ou confirmada. Aqui tambem decides se a app abre um draft e se a IA escreve esse rascunho." title="Ajuda: Comportamento dos tickets" />
                  </div>
                  <div style={S.inlineRow}>
                    <select
                      style={S.compactSelect}
                      value={ticketUiDraft.autoLinkMode}
                      onChange={(event) =>
                        setTicketUiDraft((current) => ({
                          ...current,
                          autoLinkMode: event.target.value === "auto" ? "auto" : "confirm",
                        }))
                      }
                    >
                      <option value="confirm">Confirmar ligacoes</option>
                      <option value="auto">Ligar automaticamente</option>
                    </select>
                    <label style={S.toggleRow}>
                      <input
                        type="checkbox"
                        checked={ticketUiDraft.suggestDraftOnCreate}
                        onChange={(event) =>
                          setTicketUiDraft((current) => ({ ...current, suggestDraftOnCreate: event.target.checked }))
                        }
                      />
                      <span>Sugerir draft ao criar</span>
                    </label>
                    <label style={S.toggleRow}>
                      <input
                        type="checkbox"
                        checked={ticketUiDraft.useAiDrafts}
                        onChange={(event) =>
                          setTicketUiDraft((current) => ({ ...current, useAiDrafts: event.target.checked }))
                        }
                      />
                      <span>Usar IA no draft</span>
                    </label>
                    <label style={S.toggleRow}>
                      <input
                        type="checkbox"
                        checked={ticketUiDraft.includeTicketCodeInSubject}
                        onChange={(event) =>
                          setTicketUiDraft((current) => ({ ...current, includeTicketCodeInSubject: event.target.checked }))
                        }
                      />
                      <span>Incluir ticket no assunto</span>
                    </label>
                  </div>
                  <textarea
                    style={S.textarea}
                    value={ticketUiDraft.aiInstructions}
                    onChange={(event) => setTicketUiDraft((current) => ({ ...current, aiInstructions: event.target.value }))}
                    placeholder="Instrucoes extra para a IA dos tickets"
                  />
                  <div style={S.inlineRow}>
                    <button
                      type="button"
                      style={S.primaryBtn}
                      onClick={() => void handleSaveTicketUi()}
                      disabled={busy || !ticketUiDraftChanged(createTicketUiDraft(settings?.groupTicketUi), ticketUiDraft)}
                    >
                      <Icons.Save size={12} />
                      Guardar settings
                    </button>
                  </div>
                </div>

                <div style={S.settingsColumns}>
                  <div style={S.managerList}>
                    <div style={S.fieldLabel}>Series</div>
                    {ticketSeriesLoading ? (
                      <PanelState compact tone="loading" title="A carregar series" description="A ler o catalogo central de tickets." />
                    ) : !ticketSeries.length ? (
                      <PanelState compact tone="info" title="Sem series" description="Cria a primeira serie para gerar tickets numerados." />
                    ) : (
                      ticketSeries.map((series) => {
                        const active = series.id === selectedTicketSeriesId;
                        return (
                          <button
                            key={series.id}
                            type="button"
                            style={active ? S.managerRowActive : S.managerRow}
                            onClick={() => setSelectedTicketSeriesId(series.id)}
                          >
                            <span>{series.prefix}</span>
                            <span style={S.managerCount}>{series.nextNumber}</span>
                          </button>
                        );
                      })
                    )}
                  </div>

                  <div style={S.settingsGrid}>
                    <div style={S.card}>
                      <div style={S.fieldLabel}>Nova serie</div>
                      <div style={S.inlineRow}>
                        <input
                          style={S.input}
                          value={newTicketSeriesDraft.name}
                          onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, name: event.target.value }))}
                          placeholder="Nome"
                        />
                          <input
                            style={S.input}
                            value={newTicketSeriesDraft.prefix}
                            onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, prefix: normalizeTicketPrefixInput(event.target.value) }))}
                            placeholder="Prefixo"
                          />
                      </div>
                      <div style={S.inlineRow}>
                        <select
                          style={S.compactSelect}
                          value={newTicketSeriesDraft.yearMode}
                          onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, yearMode: event.target.value as TicketSeriesDraft["yearMode"] }))}
                        >
                          {TICKET_YEAR_MODE_OPTIONS.map((option) => (
                            <option key={option.value} value={option.value}>{option.label}</option>
                          ))}
                        </select>
                        <select
                          style={S.compactSelect}
                          value={newTicketSeriesDraft.separator}
                          onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, separator: event.target.value as TicketSeriesDraft["separator"] }))}
                        >
                          {TICKET_SEPARATOR_OPTIONS.map((option) => (
                            <option key={`${option.label}:${option.value}`} value={option.value}>{option.label}</option>
                          ))}
                        </select>
                        <input
                          style={S.input}
                          value={newTicketSeriesDraft.nextNumber}
                          onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, nextNumber: event.target.value }))}
                          placeholder="Proximo numero"
                        />
                        <input
                          style={S.input}
                          value={newTicketSeriesDraft.padding}
                          onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, padding: event.target.value }))}
                          placeholder="Padding"
                        />
                        <label style={S.toggleRow}>
                          <input
                            type="checkbox"
                            checked={newTicketSeriesDraft.isActive}
                            onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, isActive: event.target.checked }))}
                          />
                            <span>Ativa</span>
                          </label>
                      </div>
                      <textarea
                        style={S.textarea}
                        value={newTicketSeriesDraft.replyInstructions}
                        onChange={(event) => setNewTicketSeriesDraft((current) => ({ ...current, replyInstructions: event.target.value }))}
                        placeholder="Instrucoes de resposta para esta serie. Ex.: agradecer o pedido e informar que vamos dar seguimento."
                      />
                      <div style={S.smallMeta}>Preview: {buildTicketSeriesPreview(newTicketSeriesDraft)}</div>
                      <div style={S.inlineRow}>
                        <button type="button" style={S.primaryBtn} onClick={() => void handleCreateTicketSeries()} disabled={busy}>
                          <Icons.Plus size={12} />
                          Criar serie
                        </button>
                      </div>
                    </div>

                    <div style={S.card}>
                      <div style={S.fieldLabel}>Serie selecionada</div>
                      {selectedTicketSeries ? (
                        <>
                          <div style={S.inlineRow}>
                            <input
                              style={S.input}
                              value={ticketSeriesDraft.name}
                              onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, name: event.target.value }))}
                              placeholder="Nome"
                            />
                              <input
                                style={S.input}
                                value={ticketSeriesDraft.prefix}
                                onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, prefix: normalizeTicketPrefixInput(event.target.value) }))}
                                placeholder="Prefixo"
                              />
                          </div>
                          <div style={S.inlineRow}>
                            <select
                              style={S.compactSelect}
                              value={ticketSeriesDraft.yearMode}
                              onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, yearMode: event.target.value as TicketSeriesDraft["yearMode"] }))}
                            >
                              {TICKET_YEAR_MODE_OPTIONS.map((option) => (
                                <option key={option.value} value={option.value}>{option.label}</option>
                              ))}
                            </select>
                            <select
                              style={S.compactSelect}
                              value={ticketSeriesDraft.separator}
                              onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, separator: event.target.value as TicketSeriesDraft["separator"] }))}
                            >
                              {TICKET_SEPARATOR_OPTIONS.map((option) => (
                                <option key={`${option.label}:${option.value}`} value={option.value}>{option.label}</option>
                              ))}
                            </select>
                            <input
                              style={S.input}
                              value={ticketSeriesDraft.nextNumber}
                              onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, nextNumber: event.target.value }))}
                              placeholder="Proximo numero"
                            />
                            <input
                              style={S.input}
                              value={ticketSeriesDraft.padding}
                              onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, padding: event.target.value }))}
                              placeholder="Padding"
                            />
                            <label style={S.toggleRow}>
                              <input
                                type="checkbox"
                                checked={ticketSeriesDraft.isActive}
                                onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, isActive: event.target.checked }))}
                              />
                                <span>Ativa</span>
                              </label>
                          </div>
                          <textarea
                            style={S.textarea}
                            value={ticketSeriesDraft.replyInstructions}
                            onChange={(event) => setTicketSeriesDraft((current) => ({ ...current, replyInstructions: event.target.value }))}
                            placeholder="Instrucoes de resposta para esta serie. Ex.: agradecer a reclamacao e informar que vamos analisar e dar seguimento."
                          />
                          <div style={S.smallMeta}>Preview: {buildTicketSeriesPreview(ticketSeriesDraft)}</div>
                          <div style={S.inlineRow}>
                            <button
                              type="button"
                              style={S.primaryBtn}
                              onClick={() => void handleSaveTicketSeries()}
                              disabled={busy || !ticketSeriesDraftChanged(selectedTicketSeries, ticketSeriesDraft)}
                            >
                              <Icons.Save size={12} />
                              Guardar
                            </button>
                            <button type="button" style={S.secondaryBtn} onClick={() => void handleResetTicketSeriesCounter()} disabled={busy}>
                              <Icons.RefreshCw size={12} />
                              Reiniciar contador
                            </button>
                            <button type="button" style={S.dangerBtn} onClick={() => void handleDeleteTicketSeries()} disabled={busy}>
                              <Icons.Trash size={12} />
                              Eliminar
                            </button>
                          </div>
                        </>
                      ) : (
                        <PanelState compact tone="info" title="Seleciona uma serie" description="Escolhe uma serie para editar o nome, o prefixo ou o contador." />
                      )}
                    </div>
                  </div>
                </div>
              </div>
            )}
            </div>
          </section>
        ) : view === "quicklink" ? (
          <section style={S.cleanPanel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("groups")}>Voltar</button>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Ligacao rapida</div>
                  <HelpHint text="Liga o email atual sem abrir um grupo. O ticket e opcional e, se ja existir, pode preencher logo grupos e etiquetas." title="Ajuda: Ligacao rapida" />
                </div>
              </div>
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Ticket</div>
                <HelpHint text="Opcional. Se o ticket ja existir, a app preenche logo grupos, etiquetas e estado para so confirmares ou ajustares." title="Ajuda: Ticket" />
              </div>

              <div style={S.togglePillBar}>
                <button type="button" style={quickLinkDraft.ticketMode === "none" ? S.togglePillActive : S.togglePill} onClick={() => clearQuickLinkTicket()}>
                  Sem ticket
                </button>
                <button
                  type="button"
                  style={quickLinkDraft.ticketMode === "existing" ? S.togglePillActive : S.togglePill}
                  onClick={() => {
                    const fallback = currentEmailTickets[0] || ticketMatches[0]?.ticket || quickLinkTicketCatalog[0];
                    if (fallback) applyQuickLinkTicket(fallback);
                  }}
                  disabled={!quickLinkTicketCatalog.length}
                >
                  Ticket existente
                </button>
                <button
                  type="button"
                  style={quickLinkDraft.ticketMode === "new" ? S.togglePillActive : S.togglePill}
                  onClick={() => enableQuickLinkNewTicket()}
                  disabled={!groupTicketsEnabled}
                >
                  Novo ticket
                </button>
              </div>

              {groupTicketsEnabled && (ticketMatches.length || currentEmailTickets.length) ? (
                <div style={S.chipWrap}>
                  {ticketMatches.map((match) => (
                    <button key={`match:${match.ticket.id}`} type="button" style={quickLinkDraft.ticketId === match.ticket.id ? S.chipActive : S.chip} onClick={() => applyQuickLinkTicket(match.ticket)}>
                      {match.ticket.code}
                    </button>
                  ))}
                  {currentEmailTickets
                    .filter((ticket) => !ticketMatches.some((match) => match.ticket.id === ticket.id))
                    .map((ticket) => (
                      <button key={`current:${ticket.id}`} type="button" style={quickLinkDraft.ticketId === ticket.id ? S.chipActive : S.chip} onClick={() => applyQuickLinkTicket(ticket)}>
                        {ticket.code}
                      </button>
                    ))}
                </div>
              ) : null}

              {groupTicketsEnabled ? (
                <>
                  <div style={S.labelSearchRow}>
                    <input
                      style={S.input}
                      value={quickLinkTicketQuery}
                      onChange={(event) => {
                        const nextValue = event.target.value;
                        setQuickLinkTicketQuery(nextValue);
                        if (String(nextValue || "").trim()) {
                          setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, ticketMode: current.ticketMode === "new" ? "new" : "existing" }));
                        }
                      }}
                      placeholder="Pesquisar ticket existente"
                    />
                    <button
                      type="button"
                      style={S.iconGhostBtn}
                      title="Pesquisar tickets"
                      onClick={() => setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, ticketMode: current.ticketMode === "new" ? "new" : "existing" }))}
                    >
                      <Icons.Search size={12} />
                    </button>
                  </div>

                  {quickLinkTicketLoading ? (
                    <PanelState compact tone="loading" title="A pesquisar tickets" description="A procurar tickets existentes para preencher o contexto." />
                  ) : quickLinkTicketResults.length ? (
                    <div style={S.cardSoft}>
                      <div style={S.smallMeta}>{quickLinkTicketQuery.trim() ? "Resultados" : "Tickets recentes"}</div>
                      <div style={S.listWrap}>
                        {quickLinkTicketResults.slice(0, 6).map((ticket) => (
                          <button key={ticket.id} type="button" style={quickLinkDraft.ticketId === ticket.id ? S.managerRowActive : S.managerRow} onClick={() => applyQuickLinkTicket(ticket)}>
                            <span>{ticket.code}</span>
                            <span style={S.managerCount}>{ticket.groupIds?.length || 0}</span>
                          </button>
                        ))}
                      </div>
                    </div>
                  ) : quickLinkDraft.ticketMode === "existing" || quickLinkTicketQuery.trim() ? (
                    <PanelState compact tone="info" title="Sem tickets encontrados" description="Nao apareceu nenhum ticket com esse codigo, serie, titulo, estado, grupo ou etiqueta." />
                  ) : null}
                </>
              ) : null}

              {quickLinkDraft.ticketMode === "existing" && quickLinkSelectedTicket ? (
                <div style={S.ticketRow}>
                  <div style={S.ticketRowHead}>
                    <div style={S.ticketCode}>{quickLinkSelectedTicket.code}</div>
                    <div style={S.ticketTitleText}>{quickLinkSelectedTicket.title || "Ticket"}</div>
                  </div>
                  <div style={S.inlineRow}>
                    <select
                      style={S.compactSelect}
                      value={quickLinkDraft.ticketStatus}
                      onChange={(event) => setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, ticketStatus: event.target.value }))}
                    >
                      {TICKET_STATUS_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                    <div style={S.smallMeta}>
                      {quickLinkSelectedTicket.groupIds?.length || 0} grupo(s) · {quickLinkSelectedTicket.labels?.length || 0} etiqueta(s)
                    </div>
                  </div>
                </div>
              ) : null}

              {quickLinkDraft.ticketMode === "new" ? (
                <div style={S.inlineRow}>
                  <select
                    style={S.select}
                    value={quickLinkDraft.ticketSeriesId}
                    onChange={(event) => setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, ticketSeriesId: event.target.value }))}
                  >
                    <option value="">Serie</option>
                    {ticketSeries.filter((entry) => entry.isActive !== false).map((series) => (
                      <option key={series.id} value={series.id}>{series.prefix}</option>
                    ))}
                  </select>
                  <select
                    style={S.compactSelect}
                    value={quickLinkDraft.ticketStatus}
                    onChange={(event) => setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, ticketStatus: event.target.value }))}
                  >
                    {TICKET_STATUS_OPTIONS.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                </div>
              ) : null}

              {quickLinkDraft.ticketMode === "new" && quickLinkDraft.ticketSeriesId ? (
                <div style={S.smallMeta}>
                  Preview: {
                    (() => {
                      const series = ticketSeries.find((entry) => entry.id === quickLinkDraft.ticketSeriesId);
                      return series
                        ? buildTicketSeriesPreview({
                          prefix: series.prefix,
                          yearMode: series.yearMode === "yy" || series.yearMode === "yyyy" ? series.yearMode : "none",
                          separator: series.separator === "/" || series.separator === "_" || series.separator === " " || series.separator === "" ? series.separator : "-",
                          nextNumber: String(series.nextNumber || 1),
                          padding: String(series.padding || 4),
                        })
                        : "Sem preview";
                    })()
                  }
                </div>
              ) : null}
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Grupo principal</div>
                <HelpHint text="O email fica guardado aqui como principal. So pode existir um." title="Ajuda: Grupo principal" />
              </div>
              <div style={S.labelSearchRow}>
                <input
                  style={S.input}
                  value={quickLinkPrimaryQuery}
                  onChange={(event) => setQuickLinkPrimaryQuery(event.target.value)}
                  placeholder="Pesquisar ou criar grupo principal"
                />
                <button
                  type="button"
                  style={!groups.some((group) => normalizeText(group.name) === normalizeText(quickLinkPrimaryQuery)) && quickLinkPrimaryQuery.trim() ? S.primaryBtn : S.iconGhostBtn}
                  onClick={() => void createQuickLinkGroupFromQuery(quickLinkPrimaryQuery, "principal")}
                  disabled={busy || !quickLinkPrimaryQuery.trim()}
                  title="Selecionar ou criar grupo principal"
                >
                  {!groups.some((group) => normalizeText(group.name) === normalizeText(quickLinkPrimaryQuery)) && quickLinkPrimaryQuery.trim() ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
                </button>
              </div>
              {quickLinkDraft.principalGroupId ? (
                <div style={S.selectedLabelRow}>
                  {groups.filter((group) => group.id === quickLinkDraft.principalGroupId).map((group) => (
                    <button key={group.id} type="button" style={S.selectedLabelChip} onClick={() => setQuickLinkPrimaryGroup("")}>
                      <span>{group.name}</span>
                      <span style={S.selectedLabelRemove}>×</span>
                    </button>
                  ))}
                </div>
              ) : null}
              <div style={S.listWrap}>
                {quickLinkPrimaryCandidates.map((group) => (
                  <button key={group.id} type="button" style={quickLinkDraft.principalGroupId === group.id ? S.managerRowActive : S.managerRow} onClick={() => setQuickLinkPrimaryGroup(group.id)}>
                    <span>{group.name}</span>
                    <span style={S.managerCount}>{group.memberCount || 0}</span>
                  </button>
                ))}
              </div>
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Grupos secundarios</div>
                <HelpHint text="Estes grupos ficam com referencia ao email, sem o duplicar como registo principal." title="Ajuda: Grupos secundarios" />
              </div>
              <div style={S.labelSearchRow}>
                <input
                  style={S.input}
                  value={quickLinkSecondaryQuery}
                  onChange={(event) => setQuickLinkSecondaryQuery(event.target.value)}
                  placeholder="Pesquisar ou criar grupo secundario"
                />
                <button
                  type="button"
                  style={!groups.some((group) => normalizeText(group.name) === normalizeText(quickLinkSecondaryQuery)) && quickLinkSecondaryQuery.trim() ? S.primaryBtn : S.iconGhostBtn}
                  onClick={() => void createQuickLinkGroupFromQuery(quickLinkSecondaryQuery, "secondary")}
                  disabled={busy || !quickLinkSecondaryQuery.trim()}
                  title="Selecionar ou criar grupo secundario"
                >
                  {!groups.some((group) => normalizeText(group.name) === normalizeText(quickLinkSecondaryQuery)) && quickLinkSecondaryQuery.trim() ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
                </button>
              </div>
              {quickLinkDraft.secondaryGroupIds.length ? (
                <div style={S.selectedLabelRow}>
                  {groups
                    .filter((group) => quickLinkDraft.secondaryGroupIds.includes(group.id))
                    .map((group) => (
                      <button key={group.id} type="button" style={S.selectedLabelChip} onClick={() => toggleQuickLinkSecondaryGroup(group.id)}>
                        <span>{group.name}</span>
                        <span style={S.selectedLabelRemove}>×</span>
                      </button>
                    ))}
                </div>
              ) : null}
              <div style={S.listWrap}>
                {quickLinkSecondaryCandidates.map((group) => (
                  <button key={group.id} type="button" style={S.managerRow} onClick={() => toggleQuickLinkSecondaryGroup(group.id)}>
                    <span>{group.name}</span>
                    <span style={S.managerCount}>{group.memberCount || 0}</span>
                  </button>
                ))}
              </div>
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Etiquetas</div>
                <HelpHint text="Campo rapido de pesquisa. Se a etiqueta nao existir, podes cria-la logo aqui." title="Ajuda: Etiquetas" />
              </div>
              <LabelPicker
                value={quickLinkDraft.labelsText}
                onChange={(next) => setQuickLinkDraft((current) => createQuickLinkDraft({ ...current, labelsText: next }))}
                catalog={allLabels}
                enabled={labelsManagerEnabled}
                busy={busy}
                placeholder="Pesquisar etiqueta"
                helpText="Pesquisa, escolhe ou cria etiquetas para este contexto."
                onCreateLabel={createCatalogLabel}
              />
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Resumo</div>
                <HelpHint text="Confirma aqui o principal, as referencias, as etiquetas e o ticket antes de gravar." title="Ajuda: Resumo" />
              </div>
              <div style={S.quickLinkSummaryList}>
                <div><b>Principal:</b> {groups.find((group) => group.id === quickLinkDraft.principalGroupId)?.name || "Sem grupo principal"}</div>
                <div><b>Referencias:</b> {groups.filter((group) => quickLinkDraft.secondaryGroupIds.includes(group.id)).map((group) => group.name).join(", ") || "Sem referencias"}</div>
                <div><b>Etiquetas:</b> {quickLinkLabels.join(", ") || "Sem etiquetas"}</div>
                <div><b>Ticket:</b> {quickLinkDraft.ticketMode === "existing" && quickLinkSelectedTicket ? quickLinkSelectedTicket.code : quickLinkDraft.ticketMode === "new" ? "Novo ticket" : "Sem ticket"}</div>
              </div>
              <div style={S.inlineRow}>
                <button type="button" style={S.primaryBtn} onClick={() => void handleSaveQuickLink()} disabled={busy}>
                  <Icons.Save size={12} />
                  Guardar ligacao
                </button>
                <button type="button" style={S.secondaryBtn} onClick={() => setView("groups")} disabled={busy}>
                  Cancelar
                </button>
              </div>
            </div>
          </section>
        ) : (
          <div style={{ ...S.track, transform: trackTransform }}>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Grupos</div>
                  <HelpHint text="Seleciona um grupo para o gerir. Usa os filtros para reduzir a lista e abre o detalhe so quando precisares." title="Ajuda: Grupos" />
                </div>
              </div>
              <div style={S.inlineActions}>
                <button type="button" style={S.primaryBtn} onClick={() => openQuickLinkView()} disabled={busy}>
                  <Icons.Link size={12} />
                  Ligar email
                </button>
                <button type="button" style={S.secondaryBtn} onClick={() => void handleOpenClassificationStudio()} disabled={busy}>
                  <Icons.Target size={12} />
                  Classificar
                </button>
              </div>
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Grupo</div>
                <HelpHint text="Pesquisa um grupo pelo nome. Se nao existir nenhum igual, o mesmo campo permite criar logo esse grupo." title="Ajuda: Grupo" />
              </div>
              <div style={S.labelSearchRow}>
                <input
                  style={S.input}
                  value={groupQuery}
                  onChange={(event) => setGroupQuery(event.target.value)}
                  placeholder="Pesquisar ou criar grupo"
                />
                <button
                  type="button"
                  style={!exactGroupMatch && String(groupQuery || "").trim() ? S.primaryBtn : S.iconGhostBtn}
                  onClick={() => void handleCreateGroup()}
                  disabled={busy || !String(groupQuery || "").trim() || exactGroupMatch}
                  title={!exactGroupMatch && String(groupQuery || "").trim() ? `Criar "${String(groupQuery || "").trim()}"` : "Pesquisar grupos existentes"}
                >
                  {!exactGroupMatch && String(groupQuery || "").trim() ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
                </button>
              </div>
              {!exactGroupMatch && String(groupQuery || "").trim() ? (
                <div style={S.smallMeta}>Criar grupo: <b>{String(groupQuery || "").trim()}</b></div>
              ) : (
                <div style={S.smallMeta}>{favoriteGroupIds.length ? "Favoritos e recentes" : "Recentes"}</div>
              )}
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Filtros</div>
                <HelpHint text="Filtra por texto, estado, arquivo e etiquetas. A explicacao fica aqui para nao sobrecarregar o ecran." title="Ajuda: Filtros" />
              </div>
              <div style={S.inlineRow}>
                <select style={S.select} value={statusFilter} onChange={(event) => setStatusFilter(event.target.value as GroupStatusFilter)}>
                  <option value="all">Todos os estados</option>
                  {STATUS_OPTIONS.map((option) => (
                    <option key={option.value} value={option.value}>{option.label}</option>
                  ))}
                </select>
                <select style={S.select} value={archiveFilter} onChange={(event) => setArchiveFilter(event.target.value as GroupArchiveFilter)}>
                  <option value="active">Ativos</option>
                  <option value="archived">Arquivados</option>
                  <option value="all">Todos</option>
                </select>
              </div>
              {allLabels.length ? (
                <div style={S.chipWrap}>
                  {allLabels.map((label) => {
                    const active = activeLabelFilters.some((entry) => normalizeText(entry) === normalizeText(label));
                    return (
                      <button key={label} type="button" style={active ? S.chipActive : S.chip} onClick={() => toggleLabelFilter(label)}>
                        {label}
                      </button>
                    );
                  })}
                </div>
              ) : null}
            </div>

            {groupsError ? <PanelState compact tone="error" title="Falha a carregar grupos" description={groupsError} /> : null}
            {groupsLoading && !groups.length ? <PanelState compact tone="loading" title="A carregar grupos" description="A preparar a lista central de grupos." /> : null}
            {!groupsLoading && !visibleGroups.length ? <PanelState compact tone="info" title="Sem grupos visiveis" description="Cria um grupo novo ou alarga os filtros." /> : null}

            <div style={S.listWrap}>
              {visibleGroups.map((group) => {
                const selected = group.id === selectedGroupId;
                const favorite = favoriteGroupSet.has(group.id);
                return (
                  <div key={group.id} style={selected ? S.groupItemActive : S.groupItem}>
                    <div style={S.groupItemHead}>
                      <button
                        type="button"
                        style={S.groupMainBtn}
                        onClick={() => {
                          setSelectedGroupId(group.id);
                          setView("detail");
                        }}
                      >
                        <div style={S.groupName}>{group.name}</div>
                        {String(group.status || "").trim() ? (
                          <span style={{ ...S.statusBadge, ...(group.status === "concluido" ? S.statusDone : group.status === "em_progresso" ? S.statusProgress : S.statusAnalysis) }}>
                            {statusLabel(group.status)}
                          </span>
                        ) : null}
                      </button>
                      <button
                        type="button"
                        style={favorite ? S.favoriteBtnActive : S.favoriteBtn}
                        onClick={() => void toggleFavoriteGroup(group.id)}
                        title={favorite ? "Remover dos favoritos" : "Adicionar aos favoritos"}
                      >
                        <Icons.Star size={11} />
                      </button>
                    </div>
                    {group.description ? <div style={S.groupDescription}>{group.description}</div> : null}
                    <div style={S.groupMeta}>
                      <span>{group.memberCount || 0} email(s)</span>
                      {favorite ? <span>Favorito</span> : null}
                      {group.isArchived ? <span>Arquivado</span> : null}
                    </div>
                    {group.labels?.length ? (
                      <div style={S.groupLabels}>
                        {group.labels.slice(0, 2).map((label) => (
                          <span key={label} style={S.labelBadge}>{label}</span>
                        ))}
                      </div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          </section>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("groups")}>Voltar</button>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>{selectedGroup ? selectedGroup.name : "Detalhe do grupo"}</div>
                  <HelpHint text="Fluxo rapido: escolhe Principal ou Referencia e liga o email atual. A biblioteca fica para outros emails ja registados." title="Ajuda: Detalhe do grupo" />
                </div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Seleciona um grupo" description="Escolhe um grupo na lista para o gerir." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.sectionTitleRow}>
                    <div style={S.fieldLabel}>Dados</div>
                    <HelpHint text="Nome, estado, arquivo e etiquetas do grupo." title="Ajuda: Dados do grupo" />
                  </div>
                  <input style={S.input} value={draft.name} onChange={(event) => setDraft((current) => ({ ...current, name: event.target.value }))} placeholder="Nome" />
                  <textarea style={S.textarea} value={draft.description} onChange={(event) => setDraft((current) => ({ ...current, description: event.target.value }))} placeholder="Descricao do grupo" />
                  <div style={S.inlineRow}>
                    <select style={S.select} value={draft.status} onChange={(event) => setDraft((current) => ({ ...current, status: event.target.value as GroupDraft["status"] }))}>
                      {STATUS_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                    <label style={S.toggleRow}>
                      <input type="checkbox" checked={draft.isArchived} onChange={(event) => setDraft((current) => ({ ...current, isArchived: event.target.checked }))} />
                      <span>Arquivado</span>
                    </label>
                  </div>
                  <LabelPicker
                    value={draft.labelsText}
                    onChange={(next) => setDraft((current) => ({ ...current, labelsText: next }))}
                    catalog={allLabels}
                    enabled={labelsManagerEnabled}
                    busy={busy}
                    placeholder="Pesquisar etiqueta"
                    helpText="Pesquisa e adiciona etiquetas ao grupo. Se nao existir nenhuma, o botao cria logo a etiqueta no catalogo."
                    onCreateLabel={createCatalogLabel}
                  />
                  <label style={S.toggleRow}>
                    <input type="checkbox" checked={draft.documentsEnabled} onChange={(event) => setDraft((current) => ({ ...current, documentsEnabled: event.target.checked }))} />
                    <span>Documentos ativos neste grupo</span>
                  </label>
                  <div style={S.inlineRow}>
                    <button type="button" style={S.primaryBtn} onClick={() => void handleSaveGroup()} disabled={busy || !draftChanged(selectedGroup, draft)}>
                      <Icons.Save size={12} />
                      Guardar
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void openExplorer()} disabled={busy}>
                      <Icons.ExternalLink size={12} />
                      Explorer
                    </button>
                    <button type="button" style={S.dangerBtn} onClick={() => void handleDeleteGroup()} disabled={busy}>
                      <Icons.Trash size={12} />
                      Eliminar
                    </button>
                  </div>
                </div>

                <div style={S.card}>
                  <div style={S.sectionRow}>
                    <div>
                      <div style={S.sectionTitleRow}>
                        <div style={S.fieldLabel}>Emails</div>
                        <HelpHint text="Aqui ligas o email atual ao grupo ou mudas a ligacao dos emails ja guardados entre Principal e Referencia." title="Ajuda: Emails do grupo" />
                      </div>
                      <div style={S.smallMeta}>
                        {groupEmails.length} email(s) ligados
                        {currentEmailAlreadyLinked ? " · email atual ligado" : ""}
                      </div>
                    </div>
                    <div style={S.inlineActions}>
                      <div style={S.togglePillBar} title="Tipo de ligacao">
                        {MEMBERSHIP_OPTIONS.map((option) => (
                          <button
                            key={option.value}
                            type="button"
                            style={linkKind === option.value ? S.togglePillActive : S.togglePill}
                            onClick={() => setLinkKind(option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                      <button type="button" style={S.secondaryBtn} onClick={() => void handleAddCurrentEmail()} disabled={busy || currentEmailAlreadyLinked} title="Ligar o email aberto ao grupo selecionado">
                        <Icons.Link size={12} />
                        Ligar atual
                      </button>
                      <button
                        type="button"
                        style={S.secondaryBtn}
                        onClick={() => void handleSaveCurrentAttachmentsToGroup()}
                        disabled={busy || selectedGroup.documentsEnabled === false || !selectedCurrentAttachments.length}
                        title="Guardar os anexos do email atual nos documentos do grupo"
                      >
                        <Icons.Paperclip size={12} />
                        Guardar anexos
                      </button>
                      <button type="button" style={S.secondaryBtn} onClick={() => setView("library")} disabled={busy} title="Abrir biblioteca de emails ja registados">
                        <Icons.Plus size={12} />
                        Biblioteca
                      </button>
                    </div>
                  </div>
                  <div style={S.smallMeta}>
                    {currentSavableAttachments.length
                      ? `${selectedCurrentAttachments.length}/${currentSavableAttachments.length} anexo(s) selecionado(s) no email atual`
                      : "Sem anexos guardaveis no email atual"}
                  </div>
                  {currentSavableAttachments.length ? (
                    <div style={S.chipWrap}>
                      <button
                        type="button"
                        style={selectedCurrentAttachments.length === currentSavableAttachments.length ? S.chipActive : S.chip}
                        onClick={() => setSelectedCurrentAttachmentKeys(currentSavableAttachments.map((attachment) => makeAttachmentSelectionKey(attachment)))}
                      >
                        Todos
                      </button>
                      <button
                        type="button"
                        style={selectedCurrentAttachments.length === 0 ? S.chipActive : S.chip}
                        onClick={() => setSelectedCurrentAttachmentKeys([])}
                      >
                        Nenhum
                      </button>
                      {currentSavableAttachments.map((attachment) => {
                        const attachmentKey = makeAttachmentSelectionKey(attachment);
                        const selected = selectedCurrentAttachmentKeys.includes(attachmentKey);
                        return (
                          <button
                            key={attachmentKey}
                            type="button"
                            style={selected ? S.chipActive : S.chip}
                            title={attachment.name}
                            onClick={() =>
                              setSelectedCurrentAttachmentKeys((current) =>
                                current.includes(attachmentKey)
                                  ? current.filter((entry) => entry !== attachmentKey)
                                  : [...current, attachmentKey]
                              )
                            }
                          >
                            {attachment.name}
                          </button>
                        );
                      })}
                    </div>
                  ) : null}
                  <div style={S.inlineRow}>
                    <input style={S.input} value={groupEmailQuery} onChange={(event) => setGroupEmailQuery(event.target.value)} placeholder="Filtrar emails do grupo" />
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleSetSelectedEmailsKind("principal")} disabled={busy || !selectedGroupEmailRows.length}>
                      Tornar principal
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleSetSelectedEmailsKind("referencia")} disabled={busy || !selectedGroupEmailRows.length}>
                      Tornar referencia
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleRemoveSelectedEmails()} disabled={busy || !selectedGroupEmailRows.length}>
                      Remover selecionados
                    </button>
                  </div>

                  {groupEmailsLoading ? <PanelState compact tone="loading" title="A carregar emails" description="A ler os emails associados a este grupo." /> : null}
                  {!groupEmailsLoading && !visibleGroupEmails.length ? <PanelState compact tone="info" title="Sem emails ligados" description="Usa 'Ligar atual' ou abre a biblioteca." /> : null}

                  <div style={S.emailList}>
                    {visibleGroupEmails.map((email) => {
                      const rowKey = makeEmailKey(email);
                      const selected = selectedGroupEmailKeys.includes(rowKey);
                      const isCurrent = rowKey === currentEmailKey || isCurrentContextEmail(email, ctx);
                      return (
                        <div key={rowKey} style={selected ? S.emailRowActive : S.emailRow}>
                          <label style={S.checkboxCell}>
                            <input
                              type="checkbox"
                              checked={selected}
                              onChange={(event) => {
                                setSelectedGroupEmailKeys((current) => event.target.checked ? [...current, rowKey] : current.filter((entry) => entry !== rowKey));
                              }}
                            />
                          </label>
                          <button type="button" style={S.emailMain} onClick={() => openLinkedOutlookEmail(email as any)}>
                            <div style={S.emailSubject}>
                              {email.subject || "Sem assunto"}
                              <span style={email.membershipKind === "referencia" ? S.refBadge : S.principalBadge}>
                                {membershipKindLabel(email.membershipKind)}
                              </span>
                              {isCurrent ? <span style={S.currentEmailBadge}>Atual</span> : null}
                            </div>
                            <div style={S.emailMeta}>{email.fromName || email.fromEmail || "Sem remetente"} · {formatDate(email.messageDateIso || email.receivedAtIso)}</div>
                          </button>
                          <button type="button" style={S.iconOnlyBtn} onClick={() => openLinkedOutlookEmail(email as any)} title="Abrir email">
                            <Icons.ExternalLink size={12} />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {groupTicketsEnabled ? (
                  <div style={S.card}>
                    <div style={S.sectionRow}>
                      <div>
                        <div style={S.sectionTitleRow}>
                          <div style={S.fieldLabel}>Tickets</div>
                          <HelpHint text="Cria tickets numerados, liga o email atual a tickets existentes e confirma tickets detetados no assunto ou no corpo do email." title="Ajuda: Tickets" />
                        </div>
                        <div style={S.smallMeta}>
                          {currentEmailTickets.length} ticket(s) ligados ao email atual
                          {ticketMatches.length ? ` · ${ticketMatches.length} detecao(oes)` : ""}
                        </div>
                      </div>
                      <div style={S.inlineActions}>
                        <select
                          style={S.compactSelect}
                          value={selectedTicketSeriesId}
                          onChange={(event) => setSelectedTicketSeriesId(event.target.value)}
                        >
                          <option value="">Serie</option>
                          {ticketSeries.filter((entry) => entry.isActive !== false).map((series) => (
                            <option key={series.id} value={series.id}>
                              {series.prefix}
                            </option>
                          ))}
                        </select>
                        <button type="button" style={S.secondaryBtn} onClick={() => void openGroupSettings({ section: "tickets" })} disabled={busy}>
                          <Icons.Settings size={12} />
                          Series
                        </button>
                        <button type="button" style={S.primaryBtn} onClick={() => void handleCreateTicketFromCurrentEmail()} disabled={busy || !selectedTicketSeriesId}>
                          <Icons.Plus size={12} />
                          Criar ticket
                        </button>
                      </div>
                    </div>

                    <div style={S.inlineRow}>
                      <input
                        style={S.input}
                        value={ticketSearchQuery}
                        onChange={(event) => setTicketSearchQuery(event.target.value)}
                        placeholder="Pesquisar ticket por codigo, titulo ou etiqueta"
                      />
                    </div>

                    {ticketDetectionLoading ? (
                      <PanelState compact tone="loading" title="A detetar tickets" description="A verificar se o email atual ja menciona tickets conhecidos." />
                    ) : null}

                    {(ticketUi?.autoLinkMode === "auto" ? false : true) && ticketMatches.length ? (
                      <div style={S.ticketStack}>
                        {ticketMatches.map((match) => {
                          const selectedIds = ticketMatchGroupSelection[match.ticket.id] || (match.proposedGroups || []).map((group) => group.id);
                          return (
                            <div key={match.ticket.id} style={S.ticketMatchCard}>
                              <div style={S.ticketRowHead}>
                                <div style={S.ticketCode}>{match.ticket.code}</div>
                                <div style={S.ticketTitleText}>{match.ticket.title || "Ticket detetado"}</div>
                              </div>
                              <div style={S.ticketMetaRow}>
                                {match.emailLinked ? "Email ja ligado" : "Ticket detetado neste email"}
                              </div>
                              {match.proposedGroups?.length ? (
                                <div style={S.chipWrap}>
                                  {match.proposedGroups.map((group) => {
                                    const active = selectedIds.includes(group.id);
                                    return (
                                      <button
                                        key={`${match.ticket.id}:${group.id}`}
                                        type="button"
                                        style={active ? S.chipActive : S.chip}
                                        onClick={() => toggleTicketMatchGroup(match.ticket.id, group.id)}
                                      >
                                        {group.name}
                                      </button>
                                    );
                                  })}
                                </div>
                              ) : null}
                              <div style={S.inlineRow}>
                                <button
                                  type="button"
                                  style={S.primaryBtn}
                                  onClick={() => void handleConfirmDetectedTicket(match)}
                                  disabled={busy || match.emailLinked}
                                >
                                  <Icons.Link size={12} />
                                  {match.emailLinked ? "Ligado" : "Confirmar ligacao"}
                                </button>
                                <button type="button" style={S.secondaryBtn} onClick={() => void handleOpenTicketReplyDraft(match.ticket)} disabled={busy}>
                                  <Icons.ExternalLink size={12} />
                                  Draft
                                </button>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ) : null}

                    {currentEmailTickets.length ? (
                      <div style={S.ticketStack}>
                        {currentEmailTickets.map((ticket) => (
                          <div key={ticket.id} style={S.ticketRow}>
                            <div style={S.ticketRowHead}>
                              <div style={S.ticketCode}>{ticket.code}</div>
                              <div style={S.ticketTitleText}>{ticket.title || "Ticket"}</div>
                            </div>
                            <div style={S.ticketMetaRow}>
                              {ticket.groupIds?.length || 0} grupo(s) · {ticket.labels?.length || 0} etiqueta(s)
                            </div>
                            <div style={S.inlineRow}>
                              <button type="button" style={S.secondaryBtn} onClick={() => void handleLinkCurrentEmailToTicket(ticket)} disabled={busy}>
                                <Icons.Link size={12} />
                                Reforcar ligacao
                              </button>
                              <button type="button" style={S.secondaryBtn} onClick={() => void handleOpenTicketReplyDraft(ticket)} disabled={busy}>
                                <Icons.ExternalLink size={12} />
                                Draft
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : null}

                    {ticketSearchLoading ? (
                      <PanelState compact tone="loading" title="A pesquisar tickets" description="A procurar tickets existentes para este grupo." />
                    ) : ticketSearchQuery.trim() && ticketSearchResults.length ? (
                      <div style={S.ticketStack}>
                        {ticketSearchResults.slice(0, 8).map((ticket) => (
                          <div key={ticket.id} style={S.ticketRow}>
                            <div style={S.ticketRowHead}>
                              <div style={S.ticketCode}>{ticket.code}</div>
                              <div style={S.ticketTitleText}>{ticket.title || "Ticket"}</div>
                            </div>
                            <div style={S.ticketMetaRow}>
                              {ticket.groupIds?.length || 0} grupo(s) · {ticket.labels?.join(", ") || "sem etiquetas"}
                            </div>
                            <div style={S.inlineRow}>
                              <button type="button" style={S.primaryBtn} onClick={() => void handleLinkCurrentEmailToTicket(ticket)} disabled={busy}>
                                <Icons.Link size={12} />
                                Ligar email
                              </button>
                              <button type="button" style={S.secondaryBtn} onClick={() => void handleOpenTicketReplyDraft(ticket)} disabled={busy}>
                                <Icons.ExternalLink size={12} />
                                Draft
                              </button>
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : !currentEmailTickets.length && !ticketMatches.length ? (
                      <PanelState compact tone="info" title="Sem tickets neste email" description="Cria um ticket novo ou pesquisa um ticket existente para ligar o email atual." />
                    ) : null}
                  </div>
                ) : null}
              </>
            )}
          </section>

          <section style={S.panel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("detail")}>Voltar</button>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Biblioteca</div>
                  <HelpHint text="Usa esta vista apenas quando queres ligar outros emails ja registados. Para o email atual, usa a ligacao rapida no detalhe." title="Ajuda: Biblioteca" />
                </div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Sem grupo selecionado" description="Escolhe primeiro um grupo." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.sectionTitleRow}>
                    <div style={S.fieldLabel}>Pesquisa</div>
                    <HelpHint text="Pesquisa por assunto, remetente, grupo ou registo e liga os emails selecionados ao grupo atual." title="Ajuda: Pesquisa de emails" />
                  </div>
                  <div style={S.inlineRow}>
                    <input style={S.input} value={libraryQuery} onChange={(event) => setLibraryQuery(event.target.value)} placeholder="Assunto, remetente, grupo, registo..." />
                    <div style={S.togglePillBar} title="Tipo de ligacao">
                      {MEMBERSHIP_OPTIONS.map((option) => (
                        <button
                          key={option.value}
                          type="button"
                          style={linkKind === option.value ? S.togglePillActive : S.togglePill}
                          onClick={() => setLinkKind(option.value)}
                        >
                          {option.label}
                        </button>
                      ))}
                    </div>
                    <button type="button" style={S.primaryBtn} onClick={() => void handleAddSelectedLibrary()} disabled={busy || !selectedLibraryRows.length}>
                      <Icons.Link size={12} />
                      Ligar
                    </button>
                  </div>
                  <div style={S.smallMeta}>Grupo alvo: {selectedGroup.name}</div>
                </div>

                {libraryLoading ? <PanelState compact tone="loading" title="A pesquisar emails" description="A varrer os emails ja registados pelo cockpit." /> : null}
                {!libraryLoading && !visibleLibraryEmails.length ? <PanelState compact tone="info" title="Sem resultados" description="Alarga a pesquisa ou usa o email atual." /> : null}

                <div style={S.emailList}>
                  {visibleLibraryEmails.map((email) => {
                    const rowKey = makeEmailKey(email);
                    const selected = selectedLibraryKeys.includes(rowKey);
                    const isCurrent = rowKey === currentEmailKey;
                    return (
                      <div key={rowKey} style={selected ? S.emailRowActive : S.emailRow}>
                        <label style={S.checkboxCell}>
                          <input
                            type="checkbox"
                            checked={selected}
                            onChange={(event) => {
                              setSelectedLibraryKeys((current) => event.target.checked ? [...current, rowKey] : current.filter((entry) => entry !== rowKey));
                            }}
                          />
                        </label>
                        <button type="button" style={S.emailMain} onClick={() => openLinkedOutlookEmail(email as any)}>
                          <div style={S.emailSubject}>
                            {email.subject || "Sem assunto"}
                            {isCurrent ? <span style={S.currentEmailBadge}>Atual</span> : null}
                          </div>
                          <div style={S.emailMeta}>{email.fromName || email.fromEmail || "Sem remetente"} · {formatDate(email.messageDateIso || email.receivedAtIso)}</div>
                          {(email.relatedGroups?.length || email.relatedRecords?.length) ? (
                            <div style={S.emailAssociations}>
                              {(email.relatedGroups || []).slice(0, 3).map((group) => (
                                <span key={`${rowKey}:g:${group.id}`} style={group.relationKind === "referencia" ? S.refBadge : S.labelBadge}>
                                  {group.name || group.id}
                                  {group.relationKind === "referencia" ? " · Ref" : ""}
                                </span>
                              ))}
                              {(email.relatedRecords || []).slice(0, 2).map((record) => (
                                <span key={`${rowKey}:r:${record.model}:${record.recordId}`} style={S.recordBadge}>{record.recordName || `${record.model}#${record.recordId}`}</span>
                              ))}
                            </div>
                          ) : null}
                        </button>
                        <button type="button" style={S.iconOnlyBtn} onClick={() => openLinkedOutlookEmail(email as any)} title="Abrir email">
                          <Icons.ExternalLink size={12} />
                        </button>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </section>
          </div>
        )}
      </div>
    </div>
  );
};

const baseButton: React.CSSProperties = {
  borderRadius: 10,
  border: "1px solid var(--iccc-card-border)",
  padding: "7px 10px",
  fontSize: 11,
  fontWeight: 600,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 6,
};

const S: Record<string, React.CSSProperties> = {
  root: { display: "grid", gap: 8, minHeight: "100%", position: "relative" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: 10, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  headerActions: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  kicker: { fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.05em" },
  title: { fontSize: 15, fontWeight: 700, color: "var(--iccc-text)" },
  headerHint: { fontSize: 11, color: "var(--iccc-text-muted)", marginTop: 1 },
  settingsGrid: { display: "grid", gap: 10 },
  settingsColumns: { display: "grid", gridTemplateColumns: "minmax(210px, 250px) minmax(0, 1fr)", gap: 10, alignItems: "start" },
  managerList: { display: "grid", gap: 6, padding: 10, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", alignContent: "start" },
  managerRow: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "#fff", padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 12, fontWeight: 600 },
  managerRowActive: { width: "100%", borderRadius: 10, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 12, fontWeight: 600 },
  managerMetaRow: { fontSize: 10, lineHeight: 1.3, color: "var(--iccc-text-muted)" },
  managerCount: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 22, height: 22, borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 700 },
  viewport: { overflow: "hidden", borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  standaloneSettingsSection: { display: "grid", gridTemplateColumns: "190px minmax(0, 1fr)", minHeight: "100%", alignItems: "stretch" },
  standaloneSettingsSidebar: { display: "grid", alignContent: "start", gap: 4, padding: 12, borderRight: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.54)", minHeight: 0, overflowY: "auto" },
  standaloneSettingsNavItem: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid transparent", background: "transparent", color: "var(--iccc-text-muted)", padding: "9px 10px", fontSize: 12, fontWeight: 700, cursor: "pointer" },
  standaloneSettingsNavItemOn: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid rgba(37, 99, 235, 0.22)", background: "rgba(219, 234, 254, 0.72)", color: "#1d4ed8", padding: "9px 10px", fontSize: 12, fontWeight: 800, cursor: "pointer" },
  standaloneSettingsContent: { padding: 10, minHeight: 0, overflowY: "auto", display: "grid", alignContent: "start", gap: 10 },
  track: { width: "300%", display: "flex", transition: "transform 0.22s ease" },
  panel: { width: "33.3333%", padding: 10, display: "grid", alignContent: "start", gap: 10, minHeight: "calc(100vh - 220px)", boxSizing: "border-box" },
  cleanPanel: { padding: 10, display: "grid", alignContent: "start", gap: 10, minHeight: "calc(100vh - 220px)", boxSizing: "border-box" },
  settingMenuGrid: { display: "grid", gap: 10, alignContent: "start" },
  settingEntry: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.78)", padding: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer" },
  settingEntryBody: { display: "grid", gap: 0, minWidth: 0 },
  settingEntryTitle: { fontSize: 13, fontWeight: 700, color: "var(--iccc-text)" },
  settingEntryMeta: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" },
  panelHeader: { display: "flex", alignItems: "center", gap: 8 },
  panelTitle: { fontSize: 14, fontWeight: 700, color: "var(--iccc-text)" },
  sectionTitleRow: { display: "inline-flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  backBtn: { ...baseButton, background: "transparent", color: "var(--iccc-text)" },
  primaryBtn: { ...baseButton, background: "linear-gradient(180deg, rgba(96, 165, 250, 0.95) 0%, rgba(37, 99, 235, 0.95) 100%)", color: "#fff", border: "1px solid rgba(37, 99, 235, 0.35)" },
  secondaryBtn: { ...baseButton, background: "rgba(255,255,255,0.78)", color: "var(--iccc-text)" },
  iconGearBtn: { ...baseButton, width: 38, height: 38, padding: 0, background: "rgba(255,255,255,0.78)", color: "var(--iccc-text)" },
  iconGhostBtn: { ...baseButton, width: 34, height: 34, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
  dangerBtn: { ...baseButton, background: "rgba(254, 226, 226, 0.95)", color: "#b91c1c", border: "1px solid rgba(239, 68, 68, 0.25)" },
  card: { display: "grid", gap: 6, padding: 10, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)" },
  cardSoft: { display: "grid", gap: 6, padding: 8, borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.5)" },
  fieldLabel: { fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.05em" },
  input: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)", boxSizing: "border-box" },
  compactSelect: { borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "7px 9px", background: "#fff", fontSize: 11, color: "var(--iccc-text)", minWidth: 120 },
  textarea: { width: "100%", minHeight: 72, borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)", boxSizing: "border-box", resize: "vertical" },
  select: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)" },
  inlineRow: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  toggleRow: { display: "inline-flex", gap: 6, alignItems: "center", fontSize: 11, color: "var(--iccc-text)" },
  chipWrap: { display: "flex", flexWrap: "wrap", gap: 8 },
  chip: { ...baseButton, padding: "4px 8px", fontSize: 10, background: "rgba(255,255,255,0.85)", color: "var(--iccc-text)" },
  chipActive: { ...baseButton, padding: "4px 8px", fontSize: 10, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  listWrap: { display: "grid", gap: 6 },
  groupItem: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", padding: 10, display: "grid", gap: 6, cursor: "pointer" },
  groupItemActive: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 10, display: "grid", gap: 6, cursor: "pointer" },
  groupItemHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  groupMainBtn: { border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, width: "100%", minWidth: 0, cursor: "pointer", textAlign: "left" },
  groupName: { fontSize: 13, fontWeight: 700, color: "var(--iccc-text)" },
  groupDescription: { fontSize: 11, color: "var(--iccc-text-muted)", lineHeight: 1.35, display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" },
  groupMeta: { display: "flex", gap: 8, flexWrap: "wrap", fontSize: 10, color: "var(--iccc-text-muted)" },
  groupLabels: { display: "flex", gap: 6, flexWrap: "wrap" },
  favoriteBtn: { ...baseButton, width: 28, height: 28, padding: 0, background: "rgba(255,255,255,0.88)", color: "var(--iccc-text-muted)" },
  favoriteBtnActive: { ...baseButton, width: 28, height: 28, padding: 0, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  labelBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 10, fontWeight: 600 },
  principalBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  refBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  recordBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 600 },
  statusBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, border: "1px solid transparent", fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" },
  statusAnalysis: { background: "rgba(217, 119, 6, 0.12)", color: "#b45309", borderColor: "rgba(217, 119, 6, 0.2)" },
  statusProgress: { background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", borderColor: "rgba(37, 99, 235, 0.2)" },
  statusDone: { background: "rgba(16, 185, 129, 0.12)", color: "#047857", borderColor: "rgba(16, 185, 129, 0.25)" },
  sectionRow: { display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" },
  smallMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  inlineActions: { display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  emailList: { display: "grid", gap: 6 },
  emailRow: { display: "grid", gridTemplateColumns: "22px minmax(0, 1fr) 30px", gap: 6, alignItems: "center", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.82)", padding: 8 },
  emailRowActive: { display: "grid", gridTemplateColumns: "22px minmax(0, 1fr) 30px", gap: 6, alignItems: "center", borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 8 },
  checkboxCell: { display: "inline-flex", alignItems: "center", justifyContent: "center" },
  emailMain: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 4, cursor: "pointer", minWidth: 0 },
  emailSubject: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" },
  emailMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  emailAssociations: { display: "flex", gap: 6, flexWrap: "wrap" },
  currentEmailBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(16, 185, 129, 0.12)", color: "#047857", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  iconOnlyBtn: { ...baseButton, width: 30, height: 30, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
  ticketStack: { display: "grid", gap: 6 },
  ticketRow: { display: "grid", gap: 5, padding: 8, borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.84)" },
  ticketMatchCard: { display: "grid", gap: 6, padding: 8, borderRadius: 10, border: "1px solid rgba(37, 99, 235, 0.18)", background: "rgba(239, 246, 255, 0.8)" },
  ticketRowHead: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  ticketCode: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(15, 23, 42, 0.08)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 700, letterSpacing: "0.04em" },
  ticketTitleText: { fontSize: 12, fontWeight: 600, color: "var(--iccc-text)" },
  ticketMetaRow: { fontSize: 10, color: "var(--iccc-text-muted)" },
  quickLinkSummaryList: { display: "grid", gap: 4, fontSize: 11, color: "var(--iccc-text)" },
  labelPicker: { display: "grid", gap: 5 },
  labelSearchRow: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) 34px", gap: 6, alignItems: "center" },
  labelSuggestionList: { display: "grid", gap: 4, padding: 6, borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)" },
  labelSuggestion: { border: "none", background: "transparent", textAlign: "left", padding: "5px 6px", borderRadius: 8, fontSize: 11, color: "var(--iccc-text)", cursor: "pointer" },
  selectedLabelRow: { display: "flex", flexWrap: "wrap", gap: 6 },
  selectedLabelChip: { ...baseButton, padding: "3px 8px", fontSize: 10, fontWeight: 600, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8" },
  selectedLabelRemove: { fontSize: 12, lineHeight: 1, opacity: 0.8 },
  togglePillBar: { display: "inline-flex", alignItems: "center", gap: 4, padding: 3, borderRadius: 999, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)" },
  togglePill: { border: "none", background: "transparent", color: "var(--iccc-text-muted)", padding: "5px 8px", borderRadius: 999, fontSize: 10, fontWeight: 700, cursor: "pointer" },
  togglePillActive: { border: "none", background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", padding: "5px 8px", borderRadius: 999, fontSize: 10, fontWeight: 700, cursor: "pointer" },
};

export default GroupManagerCockpit;
