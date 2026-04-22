import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import { addEmailToLinkGroup, createGroupTicket, createLinkGroup, deleteGroupDocument, extractAttachmentTexts, getEmailAttachmentContentBase64, getEmailAttachmentContentUrl, getEmailAttachmentTextContent, getGroupDocumentContentUrl, getGroupDocuments, getGroupEmails, getRelatedEmailContext, listLinkGroups, listGroupTicketSeries, registerRelevantEmail, removeEmailFromLinkGroup, saveGroupDocuments, searchGroupTickets, searchKnownEmails, updateGroupTicket, updateLinkGroup, type GroupDocumentEntry, type GroupTicketEntry, type GroupTicketSeriesEntry, type LinkGroupEntry, type RelatedEmailEntry, type RelevantEmailPayload } from "@/api";
import { clientLog } from "@/logger";
import { beginOutlookCategoryOperation, completeOutlookCategoryOperation, enqueueOutlookCategorySyncRequest, getManagedOutlookCategorySnapshot, OUTLOOK_CATEGORY_SYNC_DEBUG_STORAGE_KEY, requestCockpitHostAction, setOutlookCategoryOperationPhase, waitForOutlookCategorySyncResult } from "@/office";
import { buildOutlookCategoryPlan, buildOutlookCategorySourceFromRelatedContext, getOutlookCategoryPlanSignature, getOutlookCategorySourceSignature } from "@/outlookCategories";
import {
  findGroupLabelCatalogEntry,
  getGroupLabelCatalogLabels,
  getSettings,
  normalizeGroupLabelCatalog,
  type GroupLabelCatalogEntry,
  type GroupLabelStatus,
} from "@/settings";
import { PanelState } from "@/ui/PanelState";
import { applySkin } from "@/ui/skins";
import * as Icons from "@/ui/icons";
import { getStatusDisplayConfig, UNIFIED_STATUS_LEGEND } from "@/statusUtils";
import {
  addReferenceGroupSelection,
  createEmailGroupSelectionState,
  setPrincipalGroupSelection,
  toggleReferenceGroupSelection,
} from "@/modules/crm/groups-v1/contracts";
import {
  clearGroupPreparationSeed,
  readGroupPreparationSeed,
  type GroupPreparationSeed,
} from "@/modules/crm/groups-v1/prepareSession";
import { hydrateIntermediateCaseEmailsToRelatedEntries, mapIntermediateEmailToRelatedEmailEntry } from "@/modules/crm/groups-v1/storage/intermediateCaseAdapters";
import { applyClassificationToIntermediateCase, type IntermediateCaseClassificationDraft } from "@/modules/crm/groups-v1/storage/intermediateCaseClassification";
import { resolveClassificationIntermediateCase } from "@/modules/crm/groups-v1/storage/resolveClassificationIntermediateCase";
import type { IntermediateCase } from "@/modules/crm/groups-v1/storage/intermediateCaseTypes";
import "../../global.css";

import {
  type SectionId, type ScopeMode, type ApplyScopeMode, type ApplyDialogScopeMode, type PreviewMode,
  type ClassificationLayoutMode, type EmailLabelStatus, type DocumentLifecycleState,
  type ClassificationFocus, type TicketEditorMode, type AttachmentPreviewState,
  type LabelDraft, type ReadingSuggestionChip, type GroupContactDraft, type GroupEntityDraft,
  type ClassificationMetaDraft, type StudioParams
} from "./group-classification/types";

import {
  GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX, MENU, LABEL_STATUS_OPTIONS,
  TICKET_STATUS_OPTIONS, DOCUMENT_STATE_OPTIONS, EMPTY_CLASSIFICATION_META
} from "./group-classification/constants";

import {
  readParams, readSeedEmail, buildFallbackEmail, normalizeDocumentLifecycleState,
  formatDocumentLifecycleState, isRejectedDocumentLifecycleState, normalizeStudioAttachment,
  normalizeStudioAttachmentMimeType, inferStudioAttachmentKind, isLikelyDecorativeAttachment,
  isStudioAttachmentHiddenInQuickDocs, formatQuickDocumentMeta, htmlToPlainText,
  makeEmailKey, makeAttachmentKey, getStudioAttachmentRemoteId, isStudioAttachmentHydrated,
  hasHydratedAttachmentCollection, mergeUniqueStrings, mergeUniqueBy,
  scoreStudioAttachment, scoreStudioAttachmentCollection, normalizeClassificationMetaDraft,
  mergeClassificationMetaDrafts, scoreRelatedEmailEntry, mergeRelatedEmailEntries,
  dedupeEmails, buildRelevantEmailPayloadFromRelatedEmail, buildAttachmentStorageOptions,
  normalizeSearchValue, normalizeReferenceCandidate,
  compactReferenceValue, matchReferenceSet, formatDate, buildSnippet,
  buildEmailPreviewText, buildQuickDocumentPreviewText, buildCompactEmailMeta, buildEmailCorpus, isExternalEmail,
  isCurrentContextEmail, detectCaseType, inferCompanyName, normalizeGroupContactDraft,
  normalizeGroupEntityDraft, dedupeGroupContacts, dedupeGroupEntities,
  detectReferences, splitSuggestions
} from "./group-classification/documentUtils";
import {
  buildResolvedStudioApplySelection,
  buildResolvedClassifiedEmailPayload,
  buildResolvedIntermediateCaseClassificationDraft,
  buildResolvedRemoteApplyExecutionPlan,
  buildRemoteApplyFallbackCurrentCategoryEmail,
} from "./group-classification/applyResolution";
import { executeLegacyRemoteApplyForTarget } from "./group-classification/legacyRemoteApply";

import EmailsCard from "./group-classification/components/EmailsCard";
import QuickDocumentsCard from "./group-classification/components/QuickDocumentsCard";
import StatusLegend from "./group-classification/components/StatusLegend";
import ClassificationEditor from "./group-classification/components/ClassificationEditor";
import ApplyDialog from "./group-classification/components/ApplyDialog";
import PreviewPane, { StudioPdfPreview } from "./group-classification/components/PreviewPane";
import ClassificationSummaryTiles from "./group-classification/components/ClassificationSummaryTiles";
import ClassificationEditorHeader from "./group-classification/components/ClassificationEditorHeader";
import { 
  escapeHtml, sanitizeEmailPreviewHtml, buildEmailPreviewHtml, 
  decodeBase64Text, stripDataUrlPrefix, canUseOfficeWebViewer, 
  buildOfficePreviewUrl, dataUrlToUint8Array 
} from "./group-classification/previewUtils";

type CaseGroupEntry = LinkGroupEntry & { relationKind?: string };
type PrepareSeedBootstrapState = {
  key: string;
  seed: GroupPreparationSeed | null;
  status: "idle" | "invalid" | "ready" | "applied" | "skipped";
};

type IntermediateCaseBootstrapState = {
  status: "idle" | "missing" | "ready";
  caseValue: IntermediateCase | null;
  emails: RelatedEmailEntry[];
  lookup: "case_id" | "anchor_email_key" | "none";
  availability: "ready" | "missing_location" | "disabled";
  reason: string;
};

pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

function logClassificationOutlookCategorySync(phase: string, data: any) {
  if (typeof localStorage !== "undefined" && localStorage.getItem(OUTLOOK_CATEGORY_SYNC_DEBUG_STORAGE_KEY)) {
    clientLog("info", `outlook-category-sync:${phase}`, data);
  }
}

function formatEmailLabelStatus(value: string | undefined): string {
  return getStatusDisplayConfig(value).label;
}

function formatGroupStatusLabel(value: string | undefined): string {
  return getStatusDisplayConfig(value).label || "--";
}

function formatTicketStatusLabel(value: string | undefined): string {
  return getStatusDisplayConfig(value).label || "--";
}

function createLabelDraftFromCatalog(
  entry?: Partial<GroupLabelCatalogEntry> | null,
  current?: Partial<LabelDraft> | null,
  explicitStatus?: string,
  explicitCategorize?: boolean
): LabelDraft {
  const normalizedExplicitStatus = String(explicitStatus || "").trim() as EmailLabelStatus | "";
  const hasStatus = current?.hasStatus ?? (normalizedExplicitStatus ? true : entry?.hasStatus === true);
  return {
    categorize: current?.categorize ?? (typeof explicitCategorize === "boolean" ? explicitCategorize : entry?.categorize === true),
    hasStatus,
    status: hasStatus
      ? ((current?.status || normalizedExplicitStatus || entry?.status || "em_analise") as EmailLabelStatus)
      : undefined,
  };
}

function normalizeComparableString(value: unknown): string {
  return String(value || "").trim();
}

function normalizeComparableStringList(values: unknown): string[] {
  if (!Array.isArray(values)) return [];
  return Array.from(new Set(values.map((value) => normalizeComparableString(value)).filter(Boolean)));
}

function normalizeComparableStringMap(values: unknown): Record<string, string> {
  if (!values || typeof values !== "object") return {};
  return Object.fromEntries(
    Object.entries(values as Record<string, unknown>)
      .map(([key, value]) => [normalizeComparableString(key), normalizeComparableString(value)])
      .filter(([key, value]) => key && value)
  );
}

function buildCanonicalLabelDraftsFromEmail(args: {
  email: RelatedEmailEntry | null;
  labels: string[];
  labelCatalogEntries: GroupLabelCatalogEntry[];
}): Record<string, LabelDraft> {
  const labelStates = normalizeComparableStringMap(args.email?.labelStates);
  const categorizedLabelNames = normalizeComparableStringList(args.email?.classificationMeta?.categorizedLabelNames);
  return Object.fromEntries(
    args.labels.map((label) => [
      label,
      createLabelDraftFromCatalog(
        findGroupLabelCatalogEntry(args.labelCatalogEntries, label),
        null,
        labelStates[label],
        categorizedLabelNames.includes(label)
      ),
    ])
  );
}

function replaceEmailsByKey(current: RelatedEmailEntry[], incoming: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const incomingByKey = new Map(
    incoming
      .map((email) => [makeEmailKey(email), email] as const)
      .filter(([key]) => Boolean(key))
  );
  const next: RelatedEmailEntry[] = [];
  const seen = new Set<string>();

  for (const email of current) {
    const key = makeEmailKey(email);
    if (!key) {
      next.push(email);
      continue;
    }
    if (incomingByKey.has(key)) {
      next.push(incomingByKey.get(key)!);
      seen.add(key);
      continue;
    }
    next.push(email);
    seen.add(key);
  }

  for (const [key, email] of incomingByKey.entries()) {
    if (seen.has(key)) continue;
    next.push(email);
  }

  return next;
}
// Handled via imports from documentUtils.ts

function getComparableStringListSignature(values: string[]): string {
  return JSON.stringify(
    Array.from(new Set((values || []).map((value) => String(value || "").trim()).filter(Boolean)))
      .sort((left, right) => left.localeCompare(right, "pt"))
  );
}

function getComparableLabelDraftsSignature(drafts: Record<string, LabelDraft>): string {
  return JSON.stringify(
    Object.keys(drafts || {})
      .sort((left, right) => left.localeCompare(right, "pt"))
      .map((label) => {
        const draft = drafts[label];
        return {
          label,
          categorize: draft?.categorize === true,
          hasStatus: draft?.hasStatus === true,
          status: draft?.hasStatus ? String(draft?.status || "").trim() || undefined : undefined,
        };
      })
  );
}

// Local state comparison helper (derived from extracted types)
function getComparableClassificationMetaSignature(value?: Partial<ClassificationMetaDraft> | null): string {
  const normalized = normalizeClassificationMetaDraft(value);
  return JSON.stringify({
    ...normalized,
    categorizedLabelNames: [...(normalized.categorizedLabelNames || [])].sort((left, right) => left.localeCompare(right, "pt")),
  });
}

// redundant functions removed

// Moved to documentUtils.ts

// helpers moved to previewUtils.ts

// Moved to documentUtils.ts

// Moved to documentUtils.ts

function derivePartnerName(email: RelatedEmailEntry | null): string {
  const fromName = String(email?.fromName || "").trim();
  if (fromName) return fromName;
  const fromEmail = String(email?.fromEmail || "").trim().toLowerCase();
  const domain = fromEmail.includes("@") ? fromEmail.split("@")[1] : "";
  const base = domain.split(".")[0] || "";
  return base ? base.charAt(0).toUpperCase() + base.slice(1) : "";
}

function updateAttachmentStateOnEmail(
  email: RelatedEmailEntry | null,
  attachmentKey: string,
  nextState: DocumentLifecycleState
): RelatedEmailEntry | null {
  if (!email) return email;
  const targetKey = String(attachmentKey || "").trim();
  if (!targetKey || !Array.isArray(email.attachments)) return email;
  let changed = false;
  const nextAttachments = email.attachments.map((attachment) => {
    const currentKey = makeAttachmentKey(attachment || {});
    if (currentKey !== targetKey) return attachment;
    changed = true;
    return {
      ...attachment,
      documentState: nextState,
    };
  });
  if (!changed) return email;
  return {
    ...email,
    attachments: nextAttachments,
  };
}

function updateAttachmentVisibilityOnEmail(
  email: RelatedEmailEntry | null,
  attachmentKey: string,
  isHidden: boolean
): RelatedEmailEntry | null {
  if (!email) return email;
  const targetKey = String(attachmentKey || "").trim();
  if (!targetKey || !Array.isArray(email.attachments)) return email;
  let changed = false;
  const nextAttachments = email.attachments.map((attachment) => {
    const currentKey = makeAttachmentKey(attachment || {});
    if (currentKey !== targetKey) return attachment;
    changed = true;
    return {
      ...attachment,
      isHidden,
    };
  });
  if (!changed) return email;
  return {
    ...email,
    attachments: nextAttachments,
  };
}

// Helpers moved to documentUtils.ts

// Moved to documentUtils.ts

function detectReferencesFocused(text: string): string[] {
  const prepared = String(text || "")
    .replace(/[â€â€‘â€“â€”]/g, "-")
    .replace(/([A-Z0-9])\s*([/-])\s*(?=[A-Z0-9])/gi, "$1$2");
  const rawMatches: string[] = [];
  const patterns = [
    /\b(?:pedido|encomenda|order|po|purchase order|proposta|orcamento|obra|projeto|project|ref(?:erencia)?|doc(?:umento)?|fatura|invoice)\s*(?:n(?:o|Âº|Â°)?\.?\s*)?([A-Z0-9]+(?:[/-][A-Z0-9]+){1,4})\b/gi,
    /\b([A-Z]{0,6}\d{0,6}[A-Z0-9]*(?:[/-][A-Z0-9]+){1,4})\b/g,
    /\b(\d+(?:[/-][A-Z0-9]+){1,4})\b/g,
  ];
  for (const pattern of patterns) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(prepared))) {
      const normalized = normalizeReferenceCandidate(String(match[1] || ""));
      const compact = compactReferenceValue(normalized);
      if (!normalized || normalized.length < 4 || compact.length < 4 || !/\d/.test(compact)) continue;
      rawMatches.push(normalized);
    }
  }
  const ranked = rawMatches
    .reduce<Array<{ display: string; compact: string }>>((acc, value) => {
      const compact = compactReferenceValue(value);
      if (!compact) return acc;
      const existingIndex = acc.findIndex((entry) => entry.compact === compact);
      if (existingIndex >= 0) {
        if (value.length > acc[existingIndex].display.length) acc[existingIndex] = { display: value, compact };
        return acc;
      }
      acc.push({ display: value, compact });
      return acc;
    }, [])
    .sort((a, b) => b.compact.length - a.compact.length || b.display.length - a.display.length || a.display.localeCompare(b.display, "pt"));
  const filtered: Array<{ display: string; compact: string }> = [];
  for (const candidate of ranked) {
    if (filtered.some((entry) => entry.compact.includes(candidate.compact) && entry.compact !== candidate.compact)) continue;
    filtered.push(candidate);
  }
  return filtered.map((entry) => entry.display).slice(0, 8);
}

function classifyDetectedReferences(references: string[], text: string): {
  documents: string[];
  articles: string[];
  others: string[];
} {
  const upperText = String(text || "").toUpperCase();
  const documents: string[] = [];
  const articles: string[] = [];
  const others: string[] = [];
  const pushUnique = (bucket: string[], value: string) => {
    if (!bucket.includes(value)) bucket.push(value);
  };
  for (const reference of references) {
    const normalized = normalizeReferenceCandidate(reference);
    if (!normalized) continue;
    let classification: "documents" | "articles" | "others" = "others";
    const index = upperText.indexOf(normalized);
    const context = index >= 0 ? upperText.slice(Math.max(0, index - 48), Math.min(upperText.length, index + normalized.length + 48)) : "";
    if (/(PEDIDO|ENCOMENDA|ORDER|PROPOSTA|ORCAMENTO|ORÃ‡AMENTO|FATURA|INVOICE|GUIA|OBRA|PROJETO|PROJECT|DOC|DOCUMENTO|REF)/.test(context)) {
      classification = "documents";
    } else if (/(ARTIGO|ITEM|CODIGO|CÃ“DIGO|COD |MODELO|SERIE|SÃ‰RIE|PRODUTO|ACABAMENTO|COR |COLOR|TAMANHO|MEDIDA|DIMENSAO|DIMENSÃƒO)/.test(context)) {
      classification = "articles";
    } else if (/[/-]/.test(normalized)) {
      classification = "documents";
    } else if (/^(?=.*[A-Z])(?=.*\d)[A-Z0-9-]{6,}$/.test(normalized)) {
      classification = "articles";
    }
    if (classification === "documents") pushUnique(documents, normalized);
    else if (classification === "articles") pushUnique(articles, normalized);
    else pushUnique(others, normalized);
  }
  return { documents, articles, others };
}

function scoreReferenceAwareMatch(candidate: string, normalizedText: string, references: string[]): number {
  const normalizedCandidate = normalizeSearchValue(candidate);
  const compactCandidate = compactReferenceValue(candidate);
  let score = 0;
  if (normalizedCandidate && normalizedCandidate.length >= 4 && normalizedText.includes(normalizedCandidate)) {
    score = Math.max(score, 40 + Math.min(normalizedCandidate.length, 24));
  }
  if (compactCandidate && compactCandidate.length >= 4) {
    for (const reference of references) {
      const compactReference = compactReferenceValue(reference);
      if (!compactReference) continue;
      if (compactCandidate === compactReference) score = Math.max(score, 120);
      else if (compactCandidate.includes(compactReference) || compactReference.includes(compactCandidate)) score = Math.max(score, 95);
    }
  }
  return score;
}

function splitSuggestionsFocused(allGroups: LinkGroupEntry[], text: string, references: string[]): LinkGroupEntry[] {
  const normalizedText = normalizeSearchValue(text);
  return allGroups
    .map((group) => {
      if (String(group?.kind || "").trim().toLowerCase() === "conversation") return { group, score: 0 };
      const nameScore = scoreReferenceAwareMatch(String(group.name || ""), normalizedText, references) + 20;
      const labelScore = Math.max(0, ...(group.labels || []).map((label) => scoreReferenceAwareMatch(String(label || ""), normalizedText, references) + 10));
      return { group, score: Math.max(nameScore, labelScore) };
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score || String(a.group.name || "").localeCompare(String(b.group.name || ""), "pt"))
    .map((entry) => entry.group)
    .slice(0, 8);
}

function suggestTicketsFocused(tickets: GroupTicketEntry[], text: string, references: string[]): GroupTicketEntry[] {
  const normalizedText = normalizeSearchValue(text);
  return tickets
    .map((ticket) => {
      const codeScore = scoreReferenceAwareMatch(String(ticket.code || ""), normalizedText, references) + 30;
      const titleScore = scoreReferenceAwareMatch(String(ticket.title || ""), normalizedText, references) + 10;
      const labelScore = Math.max(0, ...(ticket.labels || []).map((label) => scoreReferenceAwareMatch(String(label || ""), normalizedText, references) + 5));
      return { ticket, score: Math.max(codeScore, titleScore, labelScore) };
    })
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score || String(b.ticket.updatedAt || b.ticket.createdAt || "").localeCompare(String(a.ticket.updatedAt || a.ticket.createdAt || "")))
    .map((entry) => entry.ticket)
    .slice(0, 6);
}

function suggestLabelsFocused(labels: string[], text: string, references: string[]): string[] {
  const normalizedText = normalizeSearchValue(text);
  return labels
    .map((label) => ({ label, score: scoreReferenceAwareMatch(label, normalizedText, references) }))
    .filter((entry) => entry.score > 0)
    .sort((a, b) => b.score - a.score || a.label.localeCompare(b.label, "pt"))
    .map((entry) => entry.label)
    .slice(0, 8);
}

function mergeLabels(base: string[], extra: string[]): string[] {
  const seen = new Set<string>();
  return [...base, ...extra].reduce<string[]>((acc, label) => {
    const value = String(label || "").trim();
    const key = value.toLowerCase();
    if (!value || seen.has(key)) return acc;
    seen.add(key);
    acc.push(value);
    return acc;
  }, []);
}

function areStringListsEqual(left: string[], right: string[]): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function formatChipValue(value: string | undefined, fallback = "Sem dados"): string {
  return String(value || "").trim() || fallback;
}

function emailMatchesCurrentContext(email: Partial<RelatedEmailEntry>, ctx: StudioParams | null): boolean {
  if (!ctx) return false;
  const currentItemId = String(ctx.itemId || "").trim();
  const emailItemId = String(email.itemId || "").trim();
  if (currentItemId && emailItemId && currentItemId === emailItemId) return true;
  const currentMessageId = String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  const emailMessageId = String(email.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  if (currentMessageId && emailMessageId && currentMessageId === emailMessageId) return true;
  const currentConversationId = String(ctx.conversationId || "").trim();
  const emailConversationId = String(email.conversationId || "").trim();
  const currentSubject = String(ctx.subject || "").trim().toLowerCase();
  const emailSubject = String(email.subject || "").trim().toLowerCase();
  return Boolean(currentConversationId && emailConversationId && currentConversationId === emailConversationId && currentSubject && currentSubject === emailSubject);
}

function mergeGroupEntryLists(left: LinkGroupEntry[], right: LinkGroupEntry[]): LinkGroupEntry[] {
  return [...left, ...right].reduce<LinkGroupEntry[]>((acc, group) => {
    if (!group?.id || acc.some((entry) => entry.id === group.id)) return acc;
    acc.push(group);
    return acc;
  }, []);
}

function mergeTicketEntryLists(left: GroupTicketEntry[], right: GroupTicketEntry[]): GroupTicketEntry[] {
  return [...left, ...right].reduce<GroupTicketEntry[]>((acc, ticket) => {
    if (!ticket?.id || acc.some((entry) => entry.id === ticket.id)) return acc;
    acc.push(ticket);
    return acc;
  }, []);
}

function getPrepareSeedAttachmentCandidateKeys(email: RelatedEmailEntry | null | undefined, attachment: any): string[] {
  const bareKey = String(makeAttachmentKey(attachment) || "").trim();
  const emailKey = String(makeEmailKey(email || {}) || "").trim();
  return Array.from(new Set([
    bareKey,
    bareKey && emailKey ? `${emailKey}:${bareKey}` : "",
  ].filter(Boolean)));
}

function resolvePrepareSeedAttachmentKey(
  emails: RelatedEmailEntry[],
  seedKeys: string[]
): string[] {
  const normalizedSeedKeys = new Set(seedKeys.map((key) => String(key || "").trim()).filter(Boolean));
  const resolved = new Set<string>();
  for (const email of emails) {
    for (const attachment of email.attachments || []) {
      const candidates = getPrepareSeedAttachmentCandidateKeys(email, attachment);
      if (candidates.some((candidate) => normalizedSeedKeys.has(candidate))) {
        const bareKey = String(makeAttachmentKey(attachment) || "").trim();
        if (bareKey) resolved.add(bareKey);
      }
    }
  }
  return Array.from(resolved);
}

function makeScopedAttachmentKey(
  email: RelatedEmailEntry | null | undefined,
  attachment: any
): string {
  const bareKey = String(makeAttachmentKey(attachment) || "").trim();
  const emailKey = String(makeEmailKey(email || {}) || "").trim();
  if (!bareKey) return "";
  return emailKey ? `${emailKey}:${bareKey}` : bareKey;
}

// Moved to documentUtils.ts

// readSeedEmail & buildFallbackEmail moved to documentUtils.ts

function StudioInner() {
  const params = useMemo(() => readParams(), []);
  const [section, setSection] = useState<SectionId>("emails");
  const [previewMode, setPreviewMode] = useState<PreviewMode>("email");
  const [classificationLayoutMode, setClassificationLayoutMode] = useState<ClassificationLayoutMode>("normal");
  const [scopeMode, setScopeMode] = useState<ScopeMode>("related");
  const [applyScopeMode, setApplyScopeMode] = useState<ApplyScopeMode>("current");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [groupFilterId, setGroupFilterId] = useState("");
  const [ticketFilterId, setTicketFilterId] = useState("");
  const [labelFilterValue, setLabelFilterValue] = useState("");
  const [emailSearch, setEmailSearch] = useState("");
  const [principalSearch, setPrincipalSearch] = useState("");
  const [referenceSearch, setReferenceSearch] = useState("");
  const [classificationLabelInput, setClassificationLabelInput] = useState("");
  const [onlyExternal, setOnlyExternal] = useState(false);
  const [onlyWithAttachments, setOnlyWithAttachments] = useState(false);
  const [allGroups, setAllGroups] = useState<LinkGroupEntry[]>([]);
  const [currentCaseGroups, setCurrentCaseGroups] = useState<CaseGroupEntry[]>([]);
  const [ticketSeries, setTicketSeries] = useState<GroupTicketSeriesEntry[]>([]);
  const [relatedTickets, setRelatedTickets] = useState<GroupTicketEntry[]>([]);
  const [relatedEmails, setRelatedEmails] = useState<RelatedEmailEntry[]>([]);
  const [knownEmails, setKnownEmails] = useState<RelatedEmailEntry[]>([]);
  const [selectedEmailKey, setSelectedEmailKey] = useState("");
  const [selectedTargetEmailKeys, setSelectedTargetEmailKeys] = useState<string[]>([]);
  const [principalGroupId, setPrincipalGroupId] = useState("");
  const [referenceGroupIds, setReferenceGroupIds] = useState<string[]>([]);
  const [selectedSeriesId, setSelectedSeriesId] = useState("");
  const [selectedTicketId, setSelectedTicketId] = useState("");
  const [ticketStatusDraft, setTicketStatusDraft] = useState("");
  const [ticketSearch, setTicketSearch] = useState("");
  const [ticketSearchBusy, setTicketSearchBusy] = useState(false);
  const [ticketSearchResults, setTicketSearchResults] = useState<GroupTicketEntry[]>([]);
  const [labelInput, setLabelInput] = useState("");
  const [labelCatalogReady, setLabelCatalogReady] = useState(false);
  const [labelCatalogEntries, setLabelCatalogEntries] = useState<GroupLabelCatalogEntry[]>([]);
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [labelDrafts, setLabelDrafts] = useState<Record<string, LabelDraft>>({});
  const [classificationMetaDraft, setClassificationMetaDraft] = useState<ClassificationMetaDraft>(EMPTY_CLASSIFICATION_META);
  const [createGroupName, setCreateGroupName] = useState("");
  const [createTicketTitle, setCreateTicketTitle] = useState("");
  const [attachmentPlan, setAttachmentPlan] = useState<Record<string, { analyze: boolean; save: boolean; forward: boolean }>>({});
  const [outlookLabelCategories, setOutlookLabelCategories] = useState<string[]>([]);
  const [attachmentTextMap, setAttachmentTextMap] = useState<Record<string, string>>({});
  const [selectionTouched, setSelectionTouched] = useState({ principal: false, references: false, ticket: false });
  const [actionBusy, setActionBusy] = useState(false);
  const [classificationFocus, setClassificationFocus] = useState<ClassificationFocus>("summary");
  const [applyDialogOpen, setApplyDialogOpen] = useState(false);
  const [applyDialogScopeMode, setApplyDialogScopeMode] = useState<ApplyDialogScopeMode>("current");
  const [applyDialogSection, setApplyDialogSection] = useState<ClassificationFocus>("summary");
  const [applyDialogEmailKeys, setApplyDialogEmailKeys] = useState<string[]>([]);
  const [applyDialogSelectedEmailKeys, setApplyDialogSelectedEmailKeys] = useState<string[]>([]);
  const [applyDialogExpandedEmailKeys, setApplyDialogExpandedEmailKeys] = useState<string[]>([]);
  const [expandedEmailKeys, setExpandedEmailKeys] = useState<string[]>([]);
  const [expandedQuickDocumentKeys, setExpandedQuickDocumentKeys] = useState<string[]>([]);
  const [classificationSuggestionExpanded, setClassificationSuggestionExpanded] = useState<Record<"principal" | "labels", boolean>>({
    principal: false,
    labels: false,
  });
  const [ticketEditorMode, setTicketEditorMode] = useState<TicketEditorMode>("existing");
  const [managedGroupId, setManagedGroupId] = useState("");
  const [managedGroupDescription, setManagedGroupDescription] = useState("");
  const [managedGroupNotes, setManagedGroupNotes] = useState("");
  const [managedGroupContacts, setManagedGroupContacts] = useState<GroupContactDraft[]>([]);
  const [managedGroupEntities, setManagedGroupEntities] = useState<GroupEntityDraft[]>([]);
  const [managedContactSearch, setManagedContactSearch] = useState("");
  const [managedEntitySearch, setManagedEntitySearch] = useState("");
  const [selectedAttachmentPreviewKey, setSelectedAttachmentPreviewKey] = useState("");
  const [selectedAttachmentPreviewRemoteBase64, setSelectedAttachmentPreviewRemoteBase64] = useState("");
  const [selectedAttachmentPreviewRemoteStatus, setSelectedAttachmentPreviewRemoteStatus] = useState<"idle" | "loading" | "ready" | "missing">("idle");
  const [selectedAttachmentPreviewRemoteText, setSelectedAttachmentPreviewRemoteText] = useState("");
  const [showHiddenQuickDocuments, setShowHiddenQuickDocuments] = useState(false);
  const [managedGroupEmails, setManagedGroupEmails] = useState<RelatedEmailEntry[]>([]);
  const [managedGroupDocuments, setManagedGroupDocuments] = useState<GroupDocumentEntry[]>([]);
  const [managedGroupLoading, setManagedGroupLoading] = useState(false);
  const [favoriteGroupIds, setFavoriteGroupIds] = useState<string[]>([]);
  const hydratedEmailKeysRef = useRef<Set<string>>(new Set());
  const ticketSearchRequestSeqRef = useRef(0);
  const selectedEmailRef = useRef<RelatedEmailEntry | null>(null);
  const classificationDraftSnapshotRef = useRef<null | {
    principalGroupId: string;
    principalSearch: string;
    referenceGroupIds: string[];
    referenceSearch: string;
    selectedLabels: string[];
    labelDrafts: Record<string, LabelDraft>;
    classificationMetaDraft: ClassificationMetaDraft;
    selectedTicketId: string;
    selectedSeriesId: string;
    ticketStatusDraft: string;
    ticketSearch: string;
    ticketSearchResults: GroupTicketEntry[];
    createTicketTitle: string;
    selectionTouched: { principal: boolean; references: boolean; ticket: boolean };
  } | null>(null);

  const applyInProgressRef = useRef<Promise<any> | null>(null);
  const lastAppliedSignatureRef = useRef<string | null>(null);
  const prepareSeedHandledKeyRef = useRef("");
  const [prepareSeedBootstrap, setPrepareSeedBootstrap] = useState<PrepareSeedBootstrapState>({
    key: "",
    seed: null,
    status: "idle",
  });
  const [intermediateCaseBootstrap, setIntermediateCaseBootstrap] = useState<IntermediateCaseBootstrapState>({
    status: "idle",
    caseValue: null,
    emails: [],
    lookup: "none",
    availability: "disabled",
    reason: "",
  });

  const currentSeed = useMemo(() => readSeedEmail(params), [params]);
  const fallbackIdentity = useMemo(() => buildFallbackEmail(params), [params]);
  const classificationCase = useMemo(
    () => intermediateCaseBootstrap.status === "ready" ? intermediateCaseBootstrap.caseValue : null,
    [intermediateCaseBootstrap.caseValue, intermediateCaseBootstrap.status]
  );
  const classificationCaseEmails = useMemo(
    () => classificationCase ? dedupeEmails(intermediateCaseBootstrap.emails) : [],
    [classificationCase, intermediateCaseBootstrap.emails]
  );
  const classificationAnchorEmailKey = useMemo(
    () => String(params.anchorEmailKey || classificationCase?.anchorEmailKey || "").trim(),
    [classificationCase?.anchorEmailKey, params.anchorEmailKey]
  );
  const classificationAnchorEmail = useMemo(() => {
    if (!classificationCase) return null;
    const preferredAnchorKey = classificationAnchorEmailKey;
    if (preferredAnchorKey) {
      const byPreferredKey = classificationCaseEmails.find((email) => makeEmailKey(email) === preferredAnchorKey);
      if (byPreferredKey) return byPreferredKey;
    }
    return classificationCaseEmails.find((email) => makeEmailKey(email) === classificationCase.anchorEmailKey) || classificationCaseEmails[0] || null;
  }, [classificationAnchorEmailKey, classificationCase, classificationCaseEmails]);
  const classificationRelatedEmails = useMemo(() => {
    if (!classificationCase) return [];
    const anchorKey = makeEmailKey(classificationAnchorEmail || {}) || classificationAnchorEmailKey;
    return classificationCaseEmails.filter((email) => makeEmailKey(email) !== anchorKey);
  }, [classificationAnchorEmail, classificationAnchorEmailKey, classificationCase, classificationCaseEmails]);
  const classificationContextEmails = useMemo(
    () => classificationCase ? classificationCaseEmails : dedupeEmails(relatedEmails),
    [classificationCase, classificationCaseEmails, relatedEmails]
  );
  const classificationKnownEmails = useMemo(
    () => classificationCase
      ? dedupeEmails([...classificationCaseEmails, ...relatedEmails, ...knownEmails])
      : dedupeEmails([...relatedEmails, ...knownEmails]),
    [classificationCase, classificationCaseEmails, knownEmails, relatedEmails]
  );
  const mergeEmailsIntoClassificationCase = useCallback((incomingEmails: RelatedEmailEntry[]) => {
    setIntermediateCaseBootstrap((current) => {
      if (current.status !== "ready" || !current.caseValue || !incomingEmails.length) return current;
      return {
        ...current,
        emails: dedupeEmails([...current.emails, ...incomingEmails]),
      };
    });
  }, []);
  const mergeEmailIntoClassificationCase = useCallback((incomingEmail: RelatedEmailEntry | null) => {
    if (!incomingEmail) return;
    mergeEmailsIntoClassificationCase([incomingEmail]);
  }, [mergeEmailsIntoClassificationCase]);

  useEffect(() => {
    selectedEmailRef.current = selectedEmail;
  }, [selectedEmail]);

  const rehydrateClassificationEditorFromCaseEmail = useCallback((email: RelatedEmailEntry | null) => {
    if (!email) return;
    const principalGroupId = normalizeComparableString(email.groupId || email.classificationMeta?.principalGroupId);
    const relationGroups = getEmailGroupRelations(email);
    const referenceGroupIds = relationGroups
      .filter((group) => group.id && group.id !== principalGroupId)
      .map((group) => String(group.id || "").trim());
    const normalizedSelection = createEmailGroupSelectionState({
      principalGroupId,
      referenceGroupIds,
    });
    const nextLabels = normalizeComparableStringList(email.labels);
    const nextLabelDrafts = buildCanonicalLabelDraftsFromEmail({
      email,
      labels: nextLabels,
      labelCatalogEntries,
    });
    const nextTicketId = normalizeComparableString((email.classificationMeta as any)?.ticketId);
    const nextPrincipalSearch = "";
    const nextReferenceSearch = "";
    const nextTicketSearch = "";
    const nextTicketSearchResults: GroupTicketEntry[] = [];
    const nextClassificationMetaDraft = normalizeClassificationMetaDraft({
      ...classificationMetaDraft,
      principalGroupId: normalizedSelection.principalGroupId,
      referenceGroupIds: normalizedSelection.referenceGroupIds,
      ticketId: nextTicketId,
      categorizedLabelNames: normalizeComparableStringList(email.classificationMeta?.categorizedLabelNames),
    });
    const nextSelectionTouched = { principal: false, references: false, ticket: false };

    setSelectionTouched(nextSelectionTouched);
    setPrincipalGroupId(normalizedSelection.principalGroupId);
    setPrincipalSearch(nextPrincipalSearch);
    setReferenceGroupIds(normalizedSelection.referenceGroupIds);
    setReferenceSearch(nextReferenceSearch);
    setSelectedLabels(nextLabels);
    setLabelDrafts(nextLabelDrafts);
    setClassificationMetaDraft(nextClassificationMetaDraft);
    setSelectedTicketId(nextTicketId);
    setSelectedSeriesId("");
    setTicketSearch(nextTicketSearch);
    setTicketSearchResults(nextTicketSearchResults);

    classificationDraftSnapshotRef.current = {
      principalGroupId: normalizedSelection.principalGroupId,
      principalSearch: nextPrincipalSearch,
      referenceGroupIds: [...normalizedSelection.referenceGroupIds],
      referenceSearch: nextReferenceSearch,
      selectedLabels: [...nextLabels],
      labelDrafts: structuredClone(nextLabelDrafts),
      classificationMetaDraft: structuredClone(nextClassificationMetaDraft),
      selectedTicketId: nextTicketId,
      selectedSeriesId: "",
      ticketStatusDraft,
      ticketSearch: nextTicketSearch,
      ticketSearchResults: nextTicketSearchResults,
      createTicketTitle,
      selectionTouched: nextSelectionTouched,
    };
  }, [
    classificationMetaDraft,
    createTicketTitle,
    getEmailGroupRelations,
    labelCatalogEntries,
    ticketStatusDraft,
  ]);
  const syncClassificationCaseEmails = useCallback((nextCaseValue: IntermediateCase, options?: {
    preferredSelectedEmailKey?: string;
    preferredTargetEmailKeys?: string[];
    rehydrateSelectedEmail?: boolean;
  }) => {
    const mappedEmails = nextCaseValue.emails.map((email) => mapIntermediateEmailToRelatedEmailEntry(email));
    const mappedEmailKeys = new Set(mappedEmails.map((email) => makeEmailKey(email)).filter(Boolean));
    const preferredSelectedEmailKey = normalizeComparableString(
      options?.preferredSelectedEmailKey || selectedEmailKey || classificationAnchorEmailKey || nextCaseValue.anchorEmailKey
    );
    const nextSelectedEmail = (preferredSelectedEmailKey
      ? mappedEmails.find((email) => makeEmailKey(email) === preferredSelectedEmailKey)
      : null)
      || mappedEmails.find((email) => makeEmailKey(email) === nextCaseValue.anchorEmailKey)
      || mappedEmails[0]
      || null;
    const nextSelectedEmailKey = normalizeComparableString(makeEmailKey(nextSelectedEmail || {}) || nextCaseValue.anchorEmailKey);
    const candidateTargetKeys = Array.isArray(options?.preferredTargetEmailKeys) && options?.preferredTargetEmailKeys.length
      ? options.preferredTargetEmailKeys
      : selectedTargetEmailKeys;
    const nextTargetEmailKeys = Array.from(
      new Set(
        (candidateTargetKeys || [])
          .map((key) => normalizeComparableString(key))
          .filter((key) => key && mappedEmailKeys.has(key))
      )
    );
    setIntermediateCaseBootstrap((current) => {
      if (current.status !== "ready") return current;
      return {
        ...current,
        caseValue: nextCaseValue,
        emails: dedupeEmails(mappedEmails),
      };
    });
    setRelatedEmails((current) => replaceEmailsByKey(current, mappedEmails));
    setKnownEmails((current) => replaceEmailsByKey(current, mappedEmails));
    setSelectedEmailKey(nextSelectedEmailKey);
    setSelectedTargetEmailKeys(
      nextTargetEmailKeys.length
        ? nextTargetEmailKeys
        : (nextSelectedEmailKey ? [nextSelectedEmailKey] : [])
    );
    if (options?.rehydrateSelectedEmail !== false) {
      rehydrateClassificationEditorFromCaseEmail(nextSelectedEmail);
    }
  }, [classificationAnchorEmailKey, rehydrateClassificationEditorFromCaseEmail, selectedEmailKey, selectedTargetEmailKeys]);
  const currentContext = useMemo(() => ({
    conversationId: String(params.conversationId || classificationAnchorEmail?.conversationId || currentSeed?.conversationId || fallbackIdentity?.conversationId || "").trim(),
    internetMessageId: String(params.internetMessageId || classificationAnchorEmail?.internetMessageId || currentSeed?.internetMessageId || fallbackIdentity?.internetMessageId || "").trim(),
    itemId: String(params.itemId || classificationAnchorEmail?.itemId || currentSeed?.itemId || fallbackIdentity?.itemId || "").trim(),
    subject: String(params.subject || classificationAnchorEmail?.subject || currentSeed?.subject || fallbackIdentity?.subject || "").trim(),
    fromEmail: String(params.fromEmail || classificationAnchorEmail?.fromEmail || currentSeed?.fromEmail || fallbackIdentity?.fromEmail || "").trim(),
    fromName: String(params.fromName || classificationAnchorEmail?.fromName || currentSeed?.fromName || fallbackIdentity?.fromName || "").trim(),
    receivedAtIso: String(
      params.receivedAtIso ||
      classificationAnchorEmail?.receivedAtIso ||
      classificationAnchorEmail?.messageDateIso ||
      currentSeed?.receivedAtIso ||
      currentSeed?.messageDateIso ||
      fallbackIdentity?.receivedAtIso ||
      fallbackIdentity?.messageDateIso ||
      ""
    ).trim(),
  }), [classificationAnchorEmail, currentSeed, fallbackIdentity, params]);
  const bootstrapEmailPayload = useMemo<RelevantEmailPayload | null>(() => {
    const base = classificationAnchorEmail || currentSeed || fallbackIdentity;
    if (!base) return null;
    return buildRelevantEmailPayloadFromRelatedEmail({
      ...base,
      itemId: String(currentContext.itemId || base.itemId || "").trim() || undefined,
      internetMessageId: String(currentContext.internetMessageId || base.internetMessageId || "").trim() || undefined,
      conversationId: String(currentContext.conversationId || base.conversationId || "").trim(),
      subject: String(currentContext.subject || base.subject || "").trim() || undefined,
      fromEmail: String(currentContext.fromEmail || base.fromEmail || "").trim() || undefined,
      fromName: String(currentContext.fromName || base.fromName || "").trim() || undefined,
      receivedAtIso: String(currentContext.receivedAtIso || base.receivedAtIso || base.messageDateIso || "").trim() || undefined,
      messageDateIso: String(base.messageDateIso || currentContext.receivedAtIso || base.receivedAtIso || "").trim() || undefined,
    });
  }, [
    classificationAnchorEmail,
    currentContext.conversationId,
    currentContext.fromEmail,
    currentContext.fromName,
    currentContext.internetMessageId,
    currentContext.itemId,
    currentContext.receivedAtIso,
    currentContext.subject,
    currentSeed,
    fallbackIdentity,
  ]);

  useEffect(() => {
    let cancelled = false;
    const caseId = String(params.caseId || "").trim();
    const anchorEmailKey = String(params.anchorEmailKey || "").trim();
    if (!caseId && !anchorEmailKey) {
      setIntermediateCaseBootstrap({
        status: "missing",
        caseValue: null,
        emails: [],
        lookup: "none",
        availability: "disabled",
        reason: "",
      });
      return;
    }

    setIntermediateCaseBootstrap((current) => current.status === "idle" ? current : {
      status: "idle",
      caseValue: null,
      emails: [],
      lookup: "none",
      availability: current.availability,
      reason: current.reason,
    });

    void (async () => {
      const resolved = await resolveClassificationIntermediateCase({
        caseId,
        anchorEmailKey,
      });
      const emails = resolved.caseValue
        ? await hydrateIntermediateCaseEmailsToRelatedEntries({
            caseValue: resolved.caseValue,
            adapter: resolved.storage.adapter,
          })
        : [];
      if (cancelled) return;
      setIntermediateCaseBootstrap({
        status: resolved.caseValue ? "ready" : "missing",
        caseValue: resolved.caseValue,
        emails,
        lookup: resolved.lookup,
        availability: resolved.storage.availability,
        reason: resolved.storage.reason,
      });
    })();

    return () => {
      cancelled = true;
    };
  }, [params.anchorEmailKey, params.caseId]);

  useEffect(() => {
    const key = String(params.prepareSeedKey || "").trim();
    if (!key) {
      setPrepareSeedBootstrap({ key: "", seed: null, status: "idle" });
      return;
    }
    const seed = readGroupPreparationSeed(key);
    if (!seed) {
      clearGroupPreparationSeed(key);
      setPrepareSeedBootstrap({ key, seed: null, status: "invalid" });
      return;
    }
    setPrepareSeedBootstrap({ key, seed, status: "ready" });
  }, [params.prepareSeedKey]);

  useEffect(() => {
    void (async () => {
      try {
        const settings = await getSettings();
        applySkin((settings.skinId || "soft") as any);
        setLabelCatalogEntries(normalizeGroupLabelCatalog(settings.groupLabelCatalog || []));
        setFavoriteGroupIds(Array.isArray((settings as any)?.groupFavoriteIds)
          ? Array.from(new Set((settings as any).groupFavoriteIds.map((entry: any) => String(entry || "").trim()).filter(Boolean)))
          : []);
      } catch {
        applySkin("soft" as any);
        setLabelCatalogEntries([]);
        setFavoriteGroupIds([]);
      } finally {
        setLabelCatalogReady(true);
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    const shouldWaitForCanonicalCase = Boolean(String(params.caseId || "").trim() || String(params.anchorEmailKey || "").trim());
    if (shouldWaitForCanonicalCase && intermediateCaseBootstrap.status === "idle") {
      return () => {
        cancelled = true;
      };
    }
    void (async () => {
      setLoading(true);
      setError("");
      try {
        const payload = {
          conversationId: currentContext.conversationId,
          internetMessageId: currentContext.internetMessageId,
          itemId: currentContext.itemId,
          subject: currentContext.subject,
          fromEmail: currentContext.fromEmail,
          fromName: currentContext.fromName,
          receivedAtIso: currentContext.receivedAtIso,
        };
        const [related, groups, emails, series] = await Promise.all([
          getRelatedEmailContext(payload),
          listLinkGroups(""),
          searchKnownEmails("", { limit: 120 }),
          listGroupTicketSeries(),
        ]);
        if (cancelled) return;
        const mergedGroups = [...groups, ...related.groups].reduce<LinkGroupEntry[]>((acc, group) => {
          if (!group?.id || acc.some((entry) => entry.id === group.id)) return acc;
          acc.push(group);
          return acc;
        }, []);
        const bootstrapContextEmail = bootstrapEmailPayload
          ? ({
              emailKey: makeEmailKey(bootstrapEmailPayload as any),
              itemId: bootstrapEmailPayload.itemId,
              internetMessageId: bootstrapEmailPayload.internetMessageId,
              conversationId: bootstrapEmailPayload.conversationId,
              subject: bootstrapEmailPayload.subject,
              fromEmail: bootstrapEmailPayload.fromEmail,
              fromName: bootstrapEmailPayload.fromName,
              receivedAtIso: bootstrapEmailPayload.receivedAtIso,
              messageDateIso: bootstrapEmailPayload.messageDateIso || bootstrapEmailPayload.receivedAtIso,
              bodyText: bootstrapEmailPayload.bodyText || "",
              bodyHtml: bootstrapEmailPayload.bodyHtml || "",
              attachments: bootstrapEmailPayload.attachments || [],
              relatedGroups: [],
              relatedReasons: [],
              isFallback: true,
            } as RelatedEmailEntry)
          : null;
        const hasCurrentEmailFromServer = Boolean(related.email && isCurrentContextEmail(related.email, currentContext));
        const serverContextualEmails = dedupeEmails([
          ...(related.email ? [related.email] : []),
          ...(related.emails || []),
          ...(!hasCurrentEmailFromServer && bootstrapContextEmail ? [bootstrapContextEmail] : []),
        ]);
        const canonicalContextualEmails = classificationCaseEmails;
        const mergedEmails = dedupeEmails([
          ...canonicalContextualEmails,
          ...(canonicalContextualEmails.length ? serverContextualEmails : []),
          ...(emails || []),
        ]);
        const preferredAnchorEmailKey = classificationAnchorEmail
          ? makeEmailKey(classificationAnchorEmail)
          : classificationAnchorEmailKey;
        setAllGroups(mergedGroups);
        setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
        setTicketSeries(Array.isArray(series) ? series : []);
        setRelatedTickets(Array.isArray(related.tickets) ? related.tickets : []);
        setRelatedEmails(serverContextualEmails);
        setKnownEmails(mergedEmails);
        setSelectedEmailKey((current) => {
          if (current && mergedEmails.some((email) => makeEmailKey(email) === current)) return current;
          if (preferredAnchorEmailKey && mergedEmails.some((email) => makeEmailKey(email) === preferredAnchorEmailKey)) {
            return preferredAnchorEmailKey;
          }
          const currentItem = mergedEmails.find((email) => {
            const itemId = String(email.itemId || "").trim();
            const internetMessageId = String(email.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
            const currentItemId = String(currentContext.itemId || "").trim();
            const currentMessageId = String(currentContext.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
            return (itemId && currentItemId && itemId === currentItemId)
              || (internetMessageId && currentMessageId && internetMessageId === currentMessageId);
          });
          return makeEmailKey(currentItem || mergedEmails[0] || {});
        });
        if (canonicalContextualEmails.length) {
          const sourceHint = intermediateCaseBootstrap.lookup === "case_id"
            ? "caso intermedio"
            : intermediateCaseBootstrap.lookup === "anchor_email_key"
              ? "ancora do caso intermedio"
              : "base intermedia";
          setStatus(`Classificar aberto a partir do ${sourceHint}. O legado fica apenas como fallback nesta ronda.`);
        } else if (mergedEmails.length) {
          setStatus("Janela base pronta. O email atual e os relacionados podem ser analisados aqui.");
        } else if (bootstrapEmailPayload) {
          setStatus("O email atual abriu em modo de leitura. Ainda nao existem relacionados persistidos para mostrar.");
        } else {
          setStatus("Ainda nao encontrÃ¡mos um email persistido para este caso.");
        }
      } catch (fetchError: any) {
        if (!cancelled) setError(String(fetchError?.message || fetchError || "Falha a preparar o studio de classificacao."));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [
    bootstrapEmailPayload,
    classificationAnchorEmail,
    classificationAnchorEmailKey,
    classificationCaseEmails,
    currentContext,
    currentContext.conversationId,
    currentContext.fromEmail,
    currentContext.fromName,
    currentContext.internetMessageId,
    currentContext.itemId,
    currentContext.receivedAtIso,
    currentContext.subject,
    intermediateCaseBootstrap.lookup,
    intermediateCaseBootstrap.status,
    params.anchorEmailKey,
    params.caseId,
  ]);

  useEffect(() => {
    if (loading || prepareSeedBootstrap.status !== "invalid" || !prepareSeedBootstrap.key) return;
    if (prepareSeedHandledKeyRef.current === prepareSeedBootstrap.key) return;
    prepareSeedHandledKeyRef.current = prepareSeedBootstrap.key;
    setStatus("O contexto vindo de Preparar expirou ou estava invalido. O Classificar abriu em modo normal.");
  }, [loading, prepareSeedBootstrap.key, prepareSeedBootstrap.status]);

  const groupMap = useMemo(() => new Map(allGroups.map((group) => [group.id, group])), [allGroups]);
  const businessGroups = useMemo(
    () => allGroups.filter((group) => String(group?.kind || "").trim().toLowerCase() !== "conversation"),
    [allGroups]
  );
  const currentCaseBusinessGroups = useMemo(
    () => currentCaseGroups.filter((group) => String(group?.kind || "").trim().toLowerCase() !== "conversation"),
    [currentCaseGroups]
  );
  const emailPool = useMemo(
    () => (scopeMode === "related" ? classificationContextEmails : classificationKnownEmails),
    [classificationContextEmails, classificationKnownEmails, scopeMode]
  );
  const contextualGroups = useMemo(() => {
    const rows = new Map<string, LinkGroupEntry>();
    for (const email of emailPool) {
      const isCurrentEmail =
        (String(email.itemId || "").trim() && String(email.itemId || "").trim() === String(currentContext.itemId || "").trim())
        || (
          String(email.internetMessageId || "").trim().toLowerCase() &&
          String(email.internetMessageId || "").trim().toLowerCase() === String(currentContext.internetMessageId || "").trim().toLowerCase()
        );
      const groupIds = new Set<string>([
        String(email.groupId || "").trim(),
        ...(email.relatedGroups || []).map((entry) => String(entry.id || "").trim()),
        ...(isCurrentEmail ? currentCaseBusinessGroups.map((group) => String(group.id || "").trim()) : []),
      ].filter(Boolean));
      for (const groupId of groupIds) {
        const group = groupMap.get(groupId);
        if (!group || String(group.kind || "").trim().toLowerCase() === "conversation") continue;
        rows.set(group.id, group);
      }
    }
    return Array.from(rows.values()).sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt"));
  }, [currentCaseBusinessGroups, currentContext.internetMessageId, currentContext.itemId, emailPool, groupMap]);
  const contextualTickets = useMemo(
    () => [...relatedTickets].sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || ""))),
    [relatedTickets]
  );
  const contextualLabels = useMemo(() => {
    const values = mergeLabels(
      contextualGroups.flatMap((group) => group.labels || []),
      contextualTickets.flatMap((ticket) => ticket.labels || [])
    );
    return values.sort((a, b) => a.localeCompare(b, "pt"));
  }, [contextualGroups, contextualTickets]);
  const emailContextMeta = useMemo(() => {
    const map = new Map<string, { groupIds: string[]; labels: string[]; ticketIds: string[] }>();
    for (const email of emailPool) {
      const key = makeEmailKey(email);
      if (!key) continue;
      const isCurrentEmail =
        (String(email.itemId || "").trim() && String(email.itemId || "").trim() === String(currentContext.itemId || "").trim())
        || (
          String(email.internetMessageId || "").trim().toLowerCase() &&
          String(email.internetMessageId || "").trim().toLowerCase() === String(currentContext.internetMessageId || "").trim().toLowerCase()
        );
      const groupIds = Array.from(new Set([
        String(email.groupId || "").trim(),
        ...(email.relatedGroups || []).map((entry) => String(entry.id || "").trim()),
        ...(isCurrentEmail ? currentCaseBusinessGroups.map((group) => String(group.id || "").trim()) : []),
      ].filter(Boolean)));
      const labels = mergeLabels(
        groupIds.flatMap((groupId) => groupMap.get(groupId)?.labels || []),
        contextualTickets
          .filter((ticket) => {
            const ticketGroupIds = new Set<string>([
              ...(ticket.groupIds || []).map((groupId) => String(groupId || "").trim()),
              ...(ticket.groups || []).map((group) => String(group.id || "").trim()),
            ].filter(Boolean));
            const emailKey = String(email.emailKey || "").trim();
            const matchesOrigin = Boolean(emailKey && String(ticket.createdFromEmailKey || "").trim() === emailKey);
            const matchesGroup = ticketGroupIds.size ? groupIds.some((groupId) => ticketGroupIds.has(groupId)) : false;
            return matchesOrigin || matchesGroup;
          })
          .flatMap((ticket) => ticket.labels || [])
      );
      const ticketIds = contextualTickets
        .filter((ticket) => {
          const ticketGroupIds = new Set<string>([
            ...(ticket.groupIds || []).map((groupId) => String(groupId || "").trim()),
            ...(ticket.groups || []).map((group) => String(group.id || "").trim()),
          ].filter(Boolean));
          const emailKey = String(email.emailKey || "").trim();
          const matchesOrigin = Boolean(emailKey && String(ticket.createdFromEmailKey || "").trim() === emailKey);
          const matchesGroup = ticketGroupIds.size ? groupIds.some((groupId) => ticketGroupIds.has(groupId)) : false;
          return matchesOrigin || matchesGroup;
        })
        .map((ticket) => ticket.id);
      const directTicketIds = Array.from(new Set([
        ...ticketIds,
        String((email.classificationMeta as any)?.ticketId || "").trim(),
      ].filter(Boolean)));
      map.set(key, { groupIds, labels, ticketIds: directTicketIds });
    }
    return map;
  }, [contextualTickets, currentCaseBusinessGroups, currentContext.internetMessageId, currentContext.itemId, emailPool, groupMap]);

  const visibleEmails = useMemo(() => {
    const q = String(emailSearch || "").trim().toLowerCase();
    return [...emailPool]
      .sort((a, b) => String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || "")))
      .filter((email) => {
        const meta = emailContextMeta.get(makeEmailKey(email)) || { groupIds: [], labels: [], ticketIds: [] };
        if (onlyExternal && !isExternalEmail(email)) return false;
        if (onlyWithAttachments && !(Array.isArray(email.attachments) && email.attachments.length)) return false;
        if (groupFilterId && !meta.groupIds.includes(groupFilterId)) return false;
        if (ticketFilterId && !meta.ticketIds.includes(ticketFilterId)) return false;
        if (labelFilterValue && !meta.labels.some((label) => String(label || "").trim().toLowerCase() === String(labelFilterValue || "").trim().toLowerCase())) return false;
        if (!q) return true;
        const haystack = [email.subject, email.fromName, email.fromEmail, buildSnippet(email)].join(" ").toLowerCase();
        return haystack.includes(q);
      });
  }, [emailContextMeta, emailPool, emailSearch, groupFilterId, labelFilterValue, onlyExternal, onlyWithAttachments, ticketFilterId]);

  useEffect(() => {
    if (groupFilterId && !contextualGroups.some((group) => group.id === groupFilterId)) setGroupFilterId("");
  }, [contextualGroups, groupFilterId]);

  useEffect(() => {
    if (ticketFilterId && !contextualTickets.some((ticket) => ticket.id === ticketFilterId)) setTicketFilterId("");
  }, [contextualTickets, ticketFilterId]);

  useEffect(() => {
    if (labelFilterValue && !contextualLabels.some((label) => label === labelFilterValue)) setLabelFilterValue("");
  }, [contextualLabels, labelFilterValue]);

  const selectedEmail = useMemo(
    () =>
      visibleEmails.find((email) => makeEmailKey(email) === selectedEmailKey)
      || emailPool.find((email) => makeEmailKey(email) === selectedEmailKey)
      || (classificationAnchorEmail ? emailPool.find((email) => makeEmailKey(email) === makeEmailKey(classificationAnchorEmail)) : null)
      || visibleEmails[0]
      || emailPool[0]
      || null,
    [classificationAnchorEmail, emailPool, selectedEmailKey, visibleEmails]
  );
  const selectedEmailInRelatedContext = useMemo(
    () => Boolean(selectedEmail && classificationContextEmails.some((email) => makeEmailKey(email) === makeEmailKey(selectedEmail))),
    [classificationContextEmails, selectedEmail]
  );

  const selectedEmailIsCurrent = useMemo(() => {
    return isCurrentContextEmail(selectedEmail || {}, currentContext);
  }, [currentContext, selectedEmail]);

  const getEmailGroupRelations = useCallback((email: RelatedEmailEntry | null) => {
    if (!email) return [];
    const list = [
      ...(email.relatedGroups || []),
      ...(email.groupId ? [{ id: email.groupId, name: email.groupName, relationKind: email.membershipKind }] : []),
    ];
    return list.reduce<Array<{ id: string; name?: string; relationKind?: string }>>((acc, row) => {
      if (!row?.id || acc.some((entry) => entry.id === row.id)) return acc;
      const groupKind = String((row as any)?.kind || groupMap.get(row.id)?.kind || "").trim().toLowerCase();
      if (groupKind === "conversation") return acc;
      acc.push(row);
      return acc;
    }, []);
  }, [groupMap]);

  const selectedEmailGroups = useMemo(() => {
    return getEmailGroupRelations(selectedEmail);
  }, [getEmailGroupRelations, selectedEmail]);

  const canonicalGroupSelection = useMemo(() => {
    const principalGroupId = normalizeComparableString(selectedEmail?.groupId || selectedEmail?.classificationMeta?.principalGroupId);
    const relationGroups = getEmailGroupRelations(selectedEmail || null);
    const referenceGroupIds = relationGroups
      .filter((group) => group.id && group.id !== principalGroupId)
      .map((group) => String(group.id || "").trim());
    return createEmailGroupSelectionState({
      principalGroupId,
      referenceGroupIds,
    });
  }, [getEmailGroupRelations, selectedEmail]);
  const effectivePrincipalGroupId = useMemo(
    () => normalizeComparableString(selectionTouched.principal ? principalGroupId : (canonicalGroupSelection.principalGroupId || principalGroupId)),
    [canonicalGroupSelection.principalGroupId, principalGroupId, selectionTouched.principal]
  );
  const effectiveReferenceGroupIds = useMemo(() => {
    if (selectionTouched.references) {
      return createEmailGroupSelectionState({
        principalGroupId: effectivePrincipalGroupId,
        referenceGroupIds,
      }).referenceGroupIds;
    }
    if (canonicalGroupSelection.referenceGroupIds.length || canonicalGroupSelection.principalGroupId || effectivePrincipalGroupId) {
      return createEmailGroupSelectionState({
        principalGroupId: effectivePrincipalGroupId,
        referenceGroupIds: canonicalGroupSelection.referenceGroupIds,
      }).referenceGroupIds;
    }
    return createEmailGroupSelectionState({
      principalGroupId: effectivePrincipalGroupId,
      referenceGroupIds,
    }).referenceGroupIds;
  }, [canonicalGroupSelection.principalGroupId, canonicalGroupSelection.referenceGroupIds, effectivePrincipalGroupId, referenceGroupIds, selectionTouched.references]);

  const principalAnchorGroupId = useMemo(
    () => effectivePrincipalGroupId || selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal")?.id || "",
    [effectivePrincipalGroupId, selectedEmailGroups]
  );

  const selectedTargetEmails = useMemo(
    () => emailPool.filter((email) => selectedTargetEmailKeys.includes(makeEmailKey(email))),
    [emailPool, selectedTargetEmailKeys]
  );
  const caseScopeEmails = useMemo(
    () => classificationCase
      ? classificationCaseEmails
      : dedupeEmails([...(selectedEmail ? [selectedEmail] : []), ...classificationContextEmails]),
    [classificationCase, classificationCaseEmails, classificationContextEmails, selectedEmail]
  );

  useEffect(() => {
    if (loading || prepareSeedBootstrap.status !== "ready" || !prepareSeedBootstrap.seed) return;
    if (prepareSeedHandledKeyRef.current === prepareSeedBootstrap.key) return;

    const bootstrapSeed = prepareSeedBootstrap.seed;
    const availableEmails = classificationKnownEmails;
    const availableEmailKeys = new Set(availableEmails.map((email) => makeEmailKey(email)).filter(Boolean));
    const relatedEmailKeys = new Set(classificationContextEmails.map((email) => makeEmailKey(email)).filter(Boolean));
    const selectedSeedKeys = bootstrapSeed.selectedEmailKeys.filter((key) => availableEmailKeys.has(key));
    const selectedSeedAttachmentKeys = resolvePrepareSeedAttachmentKey(availableEmails, bootstrapSeed.selectedAttachmentKeys);
    const anchorAvailable = Boolean(
      bootstrapSeed.anchorEmailKey
      && availableEmailKeys.has(bootstrapSeed.anchorEmailKey)
    );

    if (!anchorAvailable && bootstrapSeed.selectedEmailKeys.length > 0 && selectedSeedKeys.length === 0) {
      prepareSeedHandledKeyRef.current = prepareSeedBootstrap.key;
      clearGroupPreparationSeed(prepareSeedBootstrap.key);
      setPrepareSeedBootstrap((current) => current.status === "ready" ? { ...current, status: "skipped" } : current);
      setStatus("O contexto vindo de Preparar ja nao corresponde a este conjunto. O Classificar abriu em modo normal.");
      return;
    }

    const effectiveTargetKeys = selectedSeedKeys.length
      ? selectedSeedKeys
      : (anchorAvailable && bootstrapSeed.anchorEmailKey ? [bootstrapSeed.anchorEmailKey] : []);

    if (effectiveTargetKeys.length > 0) {
      const requiresAllScope = effectiveTargetKeys.some((key) => !relatedEmailKeys.has(key));
      if (requiresAllScope) setScopeMode("all");
      setSelectedTargetEmailKeys(effectiveTargetKeys);
      const preferredSelectedEmailKey = effectiveTargetKeys.includes(bootstrapSeed.anchorEmailKey)
        ? bootstrapSeed.anchorEmailKey
        : effectiveTargetKeys[0];
      if (preferredSelectedEmailKey) {
        setSelectedEmailKey((current) => current && effectiveTargetKeys.includes(current) ? current : preferredSelectedEmailKey);
      }
      setApplyScopeMode(effectiveTargetKeys.length > 1 ? "selected" : "current");
    }

    if (bootstrapSeed.workingGroupId) {
      setSelectionTouched((current) => current.principal ? current : { ...current, principal: true });
      setPrincipalGroupId(bootstrapSeed.workingGroupId);
      setReferenceGroupIds((current) => current.filter((groupId) => groupId !== bootstrapSeed.workingGroupId));
    }

    if (bootstrapSeed.filterQuery) {
      setEmailSearch((current) => current || bootstrapSeed.filterQuery);
    }

    if (bootstrapSeed.attachmentMode === "with") {
      setOnlyWithAttachments(true);
    }

    if (selectedSeedAttachmentKeys.length > 0) {
      setAttachmentPlan((current) => {
        let changed = false;
        const next = { ...current };
        selectedSeedAttachmentKeys.forEach((attachmentKey) => {
          const key = String(attachmentKey || "").trim();
          if (!key) return;
          const previous = current[key];
          const nextEntry = {
            analyze: previous?.analyze ?? false,
            save: true,
            forward: previous?.forward ?? false,
          };
          if (
            !previous
            || previous.analyze !== nextEntry.analyze
            || previous.save !== nextEntry.save
            || previous.forward !== nextEntry.forward
          ) {
            next[key] = nextEntry;
            changed = true;
          }
        });
        return changed ? next : current;
      });
    }

    prepareSeedHandledKeyRef.current = prepareSeedBootstrap.key;
    clearGroupPreparationSeed(prepareSeedBootstrap.key);
    setPrepareSeedBootstrap((current) => current.status === "ready" ? { ...current, status: "applied" } : current);

    const bootstrapSummary = [
      effectiveTargetKeys.length > 0 ? `${effectiveTargetKeys.length} email(s)` : "",
      bootstrapSeed.workingGroupId ? "grupo em trabalho" : "",
      selectedSeedAttachmentKeys.length > 0 ? `${selectedSeedAttachmentKeys.length} anexo(s) preparado(s)` : "",
      bootstrapSeed.filterQuery ? `filtro "${bootstrapSeed.filterQuery}"` : "",
    ].filter(Boolean).join(" / ");
    setStatus(
      bootstrapSummary
        ? `Contexto importado de Preparar: ${bootstrapSummary}.`
        : "Contexto de Preparar consumido. O Classificar abriu com bootstrap local."
    );
  }, [classificationContextEmails, classificationKnownEmails, loading, prepareSeedBootstrap]);

  const principalScopeEmails = useMemo(() => {
    if (!principalAnchorGroupId) return [];
    return emailPool.filter((email) =>
      getEmailGroupRelations(email).some(
        (group) => String(group.relationKind || "").toLowerCase() === "principal" && group.id === principalAnchorGroupId
      )
    );
  }, [emailPool, getEmailGroupRelations, principalAnchorGroupId]);
  const defaultApplyTargetEmails = useMemo(
    () => dedupeEmails(
      (
        applyScopeMode === "selected"
          ? selectedTargetEmails
          : applyScopeMode === "principal_group"
            ? principalScopeEmails
            : [selectedEmail].filter(Boolean)
      ) as RelatedEmailEntry[]
    ),
    [applyScopeMode, principalScopeEmails, selectedEmail, selectedTargetEmails]
  );
  const selectedTargetCount = selectedTargetEmails.length;
  const principalScopeCount = principalScopeEmails.length;
  const currentScopeEmail = useMemo(
    () => caseScopeEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || selectedEmail || caseScopeEmails[0] || null,
    [caseScopeEmails, selectedEmail, selectedEmailKey]
  );
  const applyDialogSelectedEmails = useMemo(
    () => caseScopeEmails.filter((email) => applyDialogEmailKeys.includes(makeEmailKey(email))),
    [applyDialogEmailKeys, caseScopeEmails]
  );
  const applyDialogEffectiveEmails = useMemo(() => {
    if (applyDialogScopeMode === "current") {
      return (currentScopeEmail ? [currentScopeEmail] : []) as RelatedEmailEntry[];
    }
    if (applyDialogScopeMode === "case_all") {
      return caseScopeEmails;
    }
    return applyDialogSelectedEmails;
  }, [applyDialogScopeMode, applyDialogSelectedEmails, caseScopeEmails, currentScopeEmail]);
  const normalizedTicketSearch = useMemo(() => String(ticketSearch || "").trim(), [ticketSearch]);

  const selectedEmailTicketIds = useMemo(() => {
    if (!selectedEmail) return [];
    const meta = emailContextMeta.get(makeEmailKey(selectedEmail));
    return Array.isArray(meta?.ticketIds) ? meta.ticketIds.filter(Boolean) : [];
  }, [emailContextMeta, selectedEmail]);

  useEffect(() => {
    if (!selectedEmail) return;
    const principal = selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal");
    const normalizedSelection = createEmailGroupSelectionState({
      principalGroupId: principal?.id,
      referenceGroupIds: selectedEmailGroups
        .filter((group) => String(group.relationKind || "").toLowerCase() !== "principal")
        .map((group) => group.id),
    });
    if (!selectionTouched.principal) {
      setPrincipalGroupId(normalizedSelection.principalGroupId);
    }
    if (!selectionTouched.references) {
      setReferenceGroupIds(normalizedSelection.referenceGroupIds);
    }
  }, [selectedEmail, selectedEmailGroups, selectionTouched.principal, selectionTouched.references]);

  useEffect(() => {
    setReferenceGroupIds((current) => {
      const normalizedSelection = createEmailGroupSelectionState({
        principalGroupId,
        referenceGroupIds: current,
      });
      return getComparableStringListSignature(current) === getComparableStringListSignature(normalizedSelection.referenceGroupIds)
        ? current
        : normalizedSelection.referenceGroupIds;
    });
  }, [principalGroupId]);

  useEffect(() => {
    if (!selectedEmailKey) return;
    setSelectedTargetEmailKeys((current) => {
      const existing = current.filter((key) => emailPool.some((email) => makeEmailKey(email) === key));
      return existing.length ? existing : [selectedEmailKey];
    });
  }, [emailPool, selectedEmailKey]);

  useEffect(() => {
    if (!selectedEmailKey) return;
    setPreviewMode("email");
  }, [selectedEmailKey]);

  useEffect(() => {
    setExpandedEmailKeys((current) => current.filter((key) => visibleEmails.some((email) => makeEmailKey(email) === key)));
  }, [visibleEmails]);

  useEffect(() => {
    setApplyDialogOpen(false);
    setApplyDialogExpandedEmailKeys([]);
    classificationDraftSnapshotRef.current = null;
    if (section === "classification") {
      setClassificationFocus("summary");
      setSection("emails");
    }
  }, [selectedEmailKey]);

  useEffect(() => {
    if (classificationFocus !== "ticket" || ticketEditorMode !== "existing") return;
    if (!normalizedTicketSearch) {
      ticketSearchRequestSeqRef.current += 1;
      setTicketSearchBusy(false);
      setTicketSearchResults([]);
      return;
    }
    const timeoutId = window.setTimeout(() => {
      void handleSearchTickets(normalizedTicketSearch, { silent: true });
    }, 180);
    return () => window.clearTimeout(timeoutId);
  }, [classificationFocus, normalizedTicketSearch, ticketEditorMode]);

  const previewHtml = useMemo(() => buildEmailPreviewHtml(selectedEmail), [selectedEmail]);
  const labelCatalog = useMemo(() => {
    const values = new Set<string>();
    getGroupLabelCatalogLabels(labelCatalogEntries).forEach((label) => values.add(label));
    allGroups.forEach((group) => (group.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    relatedTickets.forEach((ticket) => (ticket.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    selectedLabels.forEach((label) => values.add(label));
    return Array.from(values).sort((a, b) => a.localeCompare(b, "pt"));
  }, [allGroups, labelCatalogEntries, relatedTickets, selectedLabels]);
  const filteredLabelCatalog = useMemo(() => {
    const q = String(labelInput || "").trim().toLowerCase();
    return q ? labelCatalog.filter((label) => label.toLowerCase().includes(q)) : labelCatalog;
  }, [labelCatalog, labelInput]);
  const filteredPrincipalGroups = useMemo(() => {
    const q = normalizeSearchValue(principalSearch);
    const rows = businessGroups.filter((group) => {
      if (!q) return true;
      return normalizeSearchValue(String(group.name || "")).includes(q);
    });
    return rows
      .sort((a, b) => {
        const favoriteDelta = Number(favoriteGroupIds.includes(b.id)) - Number(favoriteGroupIds.includes(a.id));
        if (favoriteDelta) return favoriteDelta;
        return String(a.name || "").localeCompare(String(b.name || ""), "pt");
      })
      .slice(0, 18);
  }, [businessGroups, favoriteGroupIds, principalSearch]);
  const filteredReferenceGroups = useMemo(() => {
    const q = normalizeSearchValue(referenceSearch);
    const rows = businessGroups.filter((group) => {
      if (group.id === effectivePrincipalGroupId) return false;
      if (!q) return true;
      return normalizeSearchValue(String(group.name || "")).includes(q);
    });
    return rows.slice(0, 24);
  }, [businessGroups, effectivePrincipalGroupId, referenceSearch]);
  const filteredClassificationLabels = useMemo(() => {
    const q = String(classificationLabelInput || "").trim().toLowerCase();
    const rows = q ? labelCatalog.filter((label) => label.toLowerCase().includes(q)) : labelCatalog;
    return rows.slice(0, 24);
  }, [classificationLabelInput, labelCatalog]);
  const normalizedClassificationLabelSearch = useMemo(
    () => String(classificationLabelInput || "").trim().toLowerCase(),
    [classificationLabelInput]
  );
  const exactClassificationLabel = useMemo(
    () => normalizedClassificationLabelSearch
      ? labelCatalog.find((label) => label.toLowerCase() === normalizedClassificationLabelSearch) || null
      : null,
    [labelCatalog, normalizedClassificationLabelSearch]
  );
  const classificationLabelCanCreate = useMemo(
    () => Boolean(String(classificationLabelInput || "").trim() && !exactClassificationLabel),
    [classificationLabelInput, exactClassificationLabel]
  );
  const canonicalTicketChoices = useMemo(() => {
    const rows = [...relatedTickets].reduce<GroupTicketEntry[]>((acc, ticket) => {
      if (!ticket?.id || acc.some((entry) => entry.id === ticket.id)) return acc;
      acc.push(ticket);
      return acc;
    }, []);
    return rows.sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));
  }, [relatedTickets]);
  const ticketPickerChoices = useMemo(() => {
    const rows = [...canonicalTicketChoices, ...ticketSearchResults].reduce<GroupTicketEntry[]>((acc, ticket) => {
      if (!ticket?.id || acc.some((entry) => entry.id === ticket.id)) return acc;
      acc.push(ticket);
      return acc;
    }, []);
    return rows.sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));
  }, [canonicalTicketChoices, ticketSearchResults]);
  const canonicalSelectedEmailAttachmentEntries = useMemo(() => {
    if (!selectedEmail) return [] as Array<{ email: RelatedEmailEntry; attachment: any; scopedKey: string }>;
    return (selectedEmail.attachments || [])
      .map((attachment) => normalizeStudioAttachment(attachment))
      .filter((attachment): attachment is NonNullable<typeof attachment> => Boolean(attachment))
      .filter((attachment) => String(attachment.name || "").trim())
      .map((attachment) => ({
        email: selectedEmail,
        attachment,
        scopedKey: makeScopedAttachmentKey(selectedEmail, attachment),
      }))
      .filter((entry) => entry.scopedKey);
  }, [selectedEmail]);
  const selectedEmailAttachments = useMemo(
    () => canonicalSelectedEmailAttachmentEntries.map((entry) => entry.attachment),
    [canonicalSelectedEmailAttachmentEntries]
  );

  const canonicalQuickDocumentAttachments = useMemo(() => {
    const emails = dedupeEmails([selectedEmail, ...classificationContextEmails].filter(Boolean) as RelatedEmailEntry[]);
    const list: Array<{ email: RelatedEmailEntry; attachment: any; scopedKey: string }> = [];
    emails.forEach((email) => {
      const attachments = (email?.attachments || [])
        .map((att) => normalizeStudioAttachment(att))
        .filter((att): att is NonNullable<typeof att> => Boolean(att))
        .filter((att) => String(att.name || "").trim());

      attachments.forEach((attachment) => {
        const scopedKey = makeScopedAttachmentKey(email, attachment);
        if (!scopedKey) return;
        list.push({ email, attachment, scopedKey });
      });
    });
    return list;
  }, [classificationContextEmails, selectedEmail]);

  const quickDocumentAttachments = useMemo(
    () => canonicalQuickDocumentAttachments.filter((entry) => showHiddenQuickDocuments || !isStudioAttachmentHiddenInQuickDocs(entry.attachment)),
    [canonicalQuickDocumentAttachments, showHiddenQuickDocuments]
  );

  const quickDocumentHiddenCount = useMemo(
    () => canonicalQuickDocumentAttachments.filter((entry) => isStudioAttachmentHiddenInQuickDocs(entry.attachment)).length,
    [canonicalQuickDocumentAttachments]
  );
  const canonicalAttachmentPreviewEntries = useMemo(
    () => mergeUniqueBy(
      [...canonicalSelectedEmailAttachmentEntries, ...canonicalQuickDocumentAttachments],
      (entry) => entry.scopedKey
    ),
    [canonicalQuickDocumentAttachments, canonicalSelectedEmailAttachmentEntries]
  );
  useEffect(() => {
    setExpandedQuickDocumentKeys((current) =>
      current.filter((key) => quickDocumentAttachments.some((item) => item.scopedKey === key))
    );
  }, [quickDocumentAttachments]);
  const activeSelectedEmailAttachmentEntries = useMemo(
    () => canonicalSelectedEmailAttachmentEntries.filter((entry) => !isRejectedDocumentLifecycleState((entry.attachment as any)?.documentState)),
    [canonicalSelectedEmailAttachmentEntries]
  );
  const activeSelectedEmailAttachments = useMemo(
    () => activeSelectedEmailAttachmentEntries.map((entry) => entry.attachment),
    [activeSelectedEmailAttachmentEntries]
  );

  useEffect(() => {
    setSelectedAttachmentPreviewKey((current) => {
      if (current && canonicalAttachmentPreviewEntries.some((entry) => entry.scopedKey === current)) return current;
      return "";
    });
  }, [canonicalAttachmentPreviewEntries]);

  useEffect(() => {
    if (prepareSeedBootstrap.status !== "applied" || !prepareSeedBootstrap.seed?.selectedAttachmentKeys.length) return;
    setSelectedAttachmentPreviewKey((current) => {
      if (current && canonicalAttachmentPreviewEntries.some((entry) => entry.scopedKey === current)) return current;
      const seededAttachment = canonicalSelectedEmailAttachmentEntries.find((entry) =>
        getPrepareSeedAttachmentCandidateKeys(entry.email, entry.attachment).some((candidate) =>
          prepareSeedBootstrap.seed?.selectedAttachmentKeys.includes(candidate)
        )
      );
      return seededAttachment?.scopedKey || current;
    });
  }, [canonicalAttachmentPreviewEntries, canonicalSelectedEmailAttachmentEntries, prepareSeedBootstrap]);

  const selectedAttachmentPreview = useMemo(
    () => canonicalAttachmentPreviewEntries.find((entry) => entry.scopedKey === selectedAttachmentPreviewKey)?.attachment || null,
    [canonicalAttachmentPreviewEntries, selectedAttachmentPreviewKey]
  );
  const selectedAttachmentPreviewEmail = useMemo(
    () => canonicalAttachmentPreviewEntries.find((entry) => entry.scopedKey === selectedAttachmentPreviewKey)?.email || null,
    [canonicalAttachmentPreviewEntries, selectedAttachmentPreviewKey]
  );
  const selectedAttachmentDocumentState = useMemo(
    () => normalizeDocumentLifecycleState((selectedAttachmentPreview as any)?.documentState, "ingested"),
    [selectedAttachmentPreview]
  );
  const selectedAttachmentPreviewRemoteId = useMemo(
    () => getStudioAttachmentRemoteId(selectedAttachmentPreview),
    [selectedAttachmentPreview]
  );
  const selectedAttachmentPreviewEmailId = useMemo(
    () => String(selectedAttachmentPreviewEmail?.id || selectedAttachmentPreviewEmail?.emailKey || "").trim(),
    [selectedAttachmentPreviewEmail?.emailKey, selectedAttachmentPreviewEmail?.id]
  );
  const selectedAttachmentPreviewContentUrl = useMemo(() => {
    if (!selectedAttachmentPreviewEmailId || !selectedAttachmentPreviewRemoteId || selectedAttachmentPreview?.hasContent !== true) return "";
    return getEmailAttachmentContentUrl(selectedAttachmentPreviewEmailId, selectedAttachmentPreviewRemoteId);
  }, [selectedAttachmentPreview?.hasContent, selectedAttachmentPreviewEmailId, selectedAttachmentPreviewRemoteId]);

  const selectedAttachmentPreviewMode = useMemo(() => {
    const attachment = selectedAttachmentPreview;
    if (!attachment) return "none";
    const contentType = normalizeStudioAttachmentMimeType(attachment.contentType, attachment.name);
    const name = String(attachment.name || "").toLowerCase();
    if (/^image\//.test(contentType) || /\.(png|jpe?g|gif|webp|bmp|svg)$/.test(name)) return "image";
    if (contentType.includes("pdf") || /\.pdf$/.test(name)) return "pdf";
    if (
      contentType === "application/msword"
      || contentType === "application/vnd.ms-excel"
      || contentType === "application/vnd.ms-powerpoint"
      || contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      || contentType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      || contentType === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      || /\.(docx?|xlsx?|pptx?)$/.test(name)
    ) return "office";
    if (/text|json|xml|csv/.test(contentType) || /\.(txt|csv|json|xml|html?)$/.test(name)) return "text";
    return "unsupported";
  }, [selectedAttachmentPreview]);

  const selectedAttachmentPreviewSrc = useMemo(() => {
    const attachment = selectedAttachmentPreview;
    if (!attachment) return "";
    const contentType = normalizeStudioAttachmentMimeType(attachment.contentType, attachment.name);
    const localContent = String(attachment.content || "").trim() || selectedAttachmentPreviewRemoteBase64;
    if (selectedAttachmentPreviewMode === "image" || selectedAttachmentPreviewMode === "pdf") {
      if (localContent) {
        return `data:${contentType};base64,${localContent}`;
      }
      if (selectedAttachmentPreviewContentUrl) {
        return selectedAttachmentPreviewContentUrl;
      }
    }
    return "";
  }, [selectedAttachmentPreview, selectedAttachmentPreviewMode, selectedAttachmentPreviewContentUrl, selectedAttachmentPreviewRemoteBase64]);
  const selectedAttachmentOfficePreviewUrl = useMemo(
    () => selectedAttachmentPreviewMode === "office" ? buildOfficePreviewUrl(selectedAttachmentPreviewContentUrl) : "",
    [selectedAttachmentPreviewContentUrl, selectedAttachmentPreviewMode]
  );
  useEffect(() => {
    let cancelled = false;
    const localContent = String(selectedAttachmentPreview?.content || "").trim();
    if (
      (selectedAttachmentPreviewMode !== "image" && selectedAttachmentPreviewMode !== "pdf")
      || localContent
      || !selectedAttachmentPreview?.hasContent
      || !selectedAttachmentPreviewEmailId
      || !selectedAttachmentPreviewRemoteId
    ) {
      setSelectedAttachmentPreviewRemoteBase64("");
      setSelectedAttachmentPreviewRemoteStatus(
        localContent
          ? "ready"
          : selectedAttachmentPreview && selectedAttachmentPreviewMode !== "none" && selectedAttachmentPreview?.hasContent !== true
            ? "missing"
            : "idle"
      );
      return () => {
        cancelled = true;
      };
    }

    setSelectedAttachmentPreviewRemoteStatus("loading");
    getEmailAttachmentContentBase64(selectedAttachmentPreviewEmailId, selectedAttachmentPreviewRemoteId)
      .then((result) => {
        if (cancelled) return;
        const base64 = String(result.base64 || "").trim();
        setSelectedAttachmentPreviewRemoteBase64(base64);
        setSelectedAttachmentPreviewRemoteStatus(base64 ? "ready" : "missing");
      })
      .catch(() => {
        if (cancelled) return;
        setSelectedAttachmentPreviewRemoteBase64("");
        setSelectedAttachmentPreviewRemoteStatus("missing");
      });

    return () => {
      cancelled = true;
    };
  }, [
    selectedAttachmentPreview?.content,
    selectedAttachmentPreview?.hasContent,
    selectedAttachmentPreviewEmailId,
    selectedAttachmentPreviewMode,
    selectedAttachmentPreviewRemoteId,
  ]);

  useEffect(() => {
    if (!selectedAttachmentPreviewKey || showHiddenQuickDocuments) return;
    if (canonicalSelectedEmailAttachmentEntries.some((entry) => entry.scopedKey === selectedAttachmentPreviewKey)) return;
    if (quickDocumentAttachments.some((item) => item.scopedKey === selectedAttachmentPreviewKey)) return;
    const nextKey = quickDocumentAttachments[0]?.scopedKey || "";
    setSelectedAttachmentPreviewKey(nextKey);
    if (!nextKey && previewMode === "document") {
      setPreviewMode("email");
    }
  }, [canonicalSelectedEmailAttachmentEntries, previewMode, quickDocumentAttachments, selectedAttachmentPreviewKey, showHiddenQuickDocuments]);

  const selectedAttachmentPreviewText = useMemo(() => {
    if (selectedAttachmentPreviewMode !== "text") return "";
    const localContent = String(selectedAttachmentPreview?.content || "").trim();
    if (localContent) {
      return decodeBase64Text(localContent);
    }
    return selectedAttachmentPreviewRemoteText;
  }, [selectedAttachmentPreview?.content, selectedAttachmentPreviewMode, selectedAttachmentPreviewRemoteText]);
  const selectedAttachmentDocumentPreview = useMemo<AttachmentPreviewState | null>(() => {
    if (!selectedAttachmentPreview) return null;
    if (selectedAttachmentPreviewMode === "office") {
      return selectedAttachmentOfficePreviewUrl ? { kind: "office", url: selectedAttachmentOfficePreviewUrl } : { kind: "unsupported" };
    }
    if (selectedAttachmentPreviewMode === "text") {
      return selectedAttachmentPreviewText ? { kind: "text", text: selectedAttachmentPreviewText } : null;
    }
    if ((selectedAttachmentPreviewMode === "image" || selectedAttachmentPreviewMode === "pdf") && selectedAttachmentPreviewSrc) {
      return { kind: selectedAttachmentPreviewMode, src: selectedAttachmentPreviewSrc };
    }
    if (selectedAttachmentPreviewMode === "unsupported") {
      return { kind: "unsupported" };
    }
    return null;
  }, [
    selectedAttachmentOfficePreviewUrl,
    selectedAttachmentPreview,
    selectedAttachmentPreviewMode,
    selectedAttachmentPreviewSrc,
    selectedAttachmentPreviewText,
  ]);

  useEffect(() => {
    let cancelled = false;
    const localContent = String(selectedAttachmentPreview?.content || "").trim();
    if (
      selectedAttachmentPreviewMode !== "text"
      || localContent
      || !selectedAttachmentPreview?.hasContent
      || !selectedAttachmentPreviewEmailId
      || !selectedAttachmentPreviewRemoteId
    ) {
      setSelectedAttachmentPreviewRemoteText("");
      return () => {
        cancelled = true;
      };
    }
    getEmailAttachmentTextContent(selectedAttachmentPreviewEmailId, selectedAttachmentPreviewRemoteId)
      .then((text) => {
        if (cancelled) return;
        setSelectedAttachmentPreviewRemoteText(String(text || ""));
      })
      .catch(() => {
        if (cancelled) return;
        setSelectedAttachmentPreviewRemoteText("");
      });
    return () => {
      cancelled = true;
    };
  }, [
    selectedAttachmentPreview?.content,
    selectedAttachmentPreview?.hasContent,
    selectedAttachmentPreviewEmailId,
    selectedAttachmentPreviewMode,
    selectedAttachmentPreviewRemoteId,
  ]);

  useEffect(() => {
    setAttachmentPlan((current) => {
      const next = { ...current };
      let changed = false;
      for (const entry of canonicalSelectedEmailAttachmentEntries) {
        const attachment = entry.attachment;
        const key = entry.scopedKey;
        if (!key) continue;
        const contentType = String(attachment.contentType || "").toLowerCase();
        const isDocument = /pdf|image|excel|spreadsheet|word|officedocument|text|csv/.test(contentType) || /\.(pdf|png|jpe?g|xlsx?|docx?|csv|txt)$/i.test(String(attachment.name || ""));
        const previous = current[key];
        const nextEntry = {
          analyze: previous?.analyze ?? (isRejectedDocumentLifecycleState((attachment as any)?.documentState) ? false : isDocument),
          save: previous?.save ?? false,
          forward: previous?.forward ?? false,
        };
        if (
          !previous
          || previous.analyze !== nextEntry.analyze
          || previous.save !== nextEntry.save
          || previous.forward !== nextEntry.forward
        ) {
          next[key] = nextEntry;
          changed = true;
        }
      }
      return changed ? next : current;
    });
  }, [canonicalSelectedEmailAttachmentEntries]);

  useEffect(() => {
    let cancelled = false;
    const extractableFiles = activeSelectedEmailAttachmentEntries
      .map((entry) => ({
        key: makeAttachmentKey(entry.attachment),
        name: String(entry.attachment.name || "").trim(),
        contentType: String(entry.attachment.contentType || "").trim(),
        content: String(entry.attachment.content || "").trim(),
      }))
      .filter((attachment) => {
        if (!attachment.key || !attachment.name || !attachment.content) return false;
        const lowerName = attachment.name.toLowerCase();
        const lowerType = attachment.contentType.toLowerCase();
        return lowerType === "application/pdf"
          || lowerType.startsWith("text/")
          || /json|xml|csv|html|message\/rfc822/.test(lowerType)
          || /\.(pdf|txt|csv|json|xml|html?|eml)$/i.test(lowerName);
      })
      .slice(0, 6);
    if (!extractableFiles.length) {
      setAttachmentTextMap({});
      return () => { cancelled = true; };
    }
    void (async () => {
      try {
        const results = await extractAttachmentTexts(extractableFiles);
        if (cancelled) return;
        const next = results.reduce<Record<string, string>>((acc, entry) => {
          const key = String(entry?.key || "").trim();
          const text = String(entry?.text || "").trim();
          if (key && text) acc[key] = text;
          return acc;
        }, {});
        setAttachmentTextMap(next);
      } catch {
        if (!cancelled) setAttachmentTextMap({});
      }
    })();
    return () => { cancelled = true; };
  }, [activeSelectedEmailAttachmentEntries]);

  const detectionText = useMemo(() => {
    const attachmentNames = activeSelectedEmailAttachments.map((attachment) => attachment.name).join(" ");
    const attachmentTexts = activeSelectedEmailAttachments
      .map((attachment) => attachmentTextMap[makeAttachmentKey(attachment)] || "")
      .filter(Boolean)
      .join("\n\n");
    return [
      selectedEmail?.subject,
      selectedEmail?.fromName,
      selectedEmail?.fromEmail,
      selectedEmail?.bodyText,
      htmlToPlainText(String(selectedEmail?.bodyHtml || "")),
      attachmentNames,
      attachmentTexts,
    ].filter(Boolean).join(" ");
  }, [activeSelectedEmailAttachments, attachmentTextMap, selectedEmail?.bodyHtml, selectedEmail?.bodyText, selectedEmail?.fromEmail, selectedEmail?.fromName, selectedEmail?.subject]);

  const detectedCaseType = useMemo(() => detectCaseType(detectionText), [detectionText]);
  const detectedReferences = useMemo(() => detectReferencesFocused(detectionText), [detectionText]);
  const detectedReferenceBuckets = useMemo(
    () => classifyDetectedReferences(detectedReferences, detectionText),
    [detectedReferences, detectionText]
  );
  const documentReferences = detectedReferenceBuckets.documents.length
    ? detectedReferenceBuckets.documents
    : detectedReferences;
  const articleReferences = detectedReferenceBuckets.articles;
  const analyzedAttachmentNames = useMemo(
    () => activeSelectedEmailAttachments.filter((attachment) => Boolean(attachmentTextMap[makeAttachmentKey(attachment)])).map((attachment) => String(attachment.name || "").trim()).filter(Boolean),
    [activeSelectedEmailAttachments, attachmentTextMap]
  );
  const suggestedExistingGroups = useMemo(() => splitSuggestionsFocused(allGroups, detectionText, documentReferences), [allGroups, detectionText, documentReferences]);
  const suggestedExistingTickets = useMemo(
    () => suggestTicketsFocused(canonicalTicketChoices, detectionText, documentReferences),
    [canonicalTicketChoices, detectionText, documentReferences]
  );
  const suggestedLabelSeeds = useMemo(() => {
    const values = new Set<string>();
    if (detectedCaseType !== "geral") values.add(`tipo:${detectedCaseType}`);
    for (const ref of documentReferences) values.add(ref);
    for (const ref of articleReferences) values.add(`art:${ref}`);
    const partner = derivePartnerName(selectedEmail);
    if (partner) values.add(partner);
    suggestedExistingGroups.forEach((group) => (group.labels || []).forEach((label) => values.add(String(label || "").trim())));
    suggestedExistingTickets.forEach((ticket) => (ticket.labels || []).forEach((label) => values.add(String(label || "").trim())));
    suggestLabelsFocused(contextualLabels, detectionText, documentReferences).forEach((label) => values.add(label));
    return Array.from(values).filter(Boolean).slice(0, 10);
  }, [articleReferences, contextualLabels, detectedCaseType, detectionText, documentReferences, selectedEmail, suggestedExistingGroups, suggestedExistingTickets]);

  const suggestedGroupName = useMemo(() => {
    const partner = derivePartnerName(selectedEmail);
    if (documentReferences.length && partner) return `${partner} / ${documentReferences[0]}`;
    if (documentReferences.length) return documentReferences[0];
    if (partner && detectedCaseType !== "geral") return `${partner} / ${detectedCaseType}`;
    return partner || String(selectedEmail?.subject || "").trim().slice(0, 72);
  }, [detectedCaseType, documentReferences, selectedEmail]);

  const classificationSuggestions = useMemo<ReadingSuggestionChip[]>(() => {
    const entries: ReadingSuggestionChip[] = [];
    const seen = new Set<string>();
    for (const group of suggestedExistingGroups) {
      const id = String(group.id || "").trim();
      const name = String(group.name || "").trim();
      if (!id || !name) continue;
      const key = `group:${id}`;
      if (seen.has(key)) continue;
      seen.add(key);
      entries.push({ key, label: name, kind: "group", value: id });
    }
    for (const ticket of suggestedExistingTickets) {
      const id = String(ticket.id || "").trim();
      const code = String(ticket.code || "").trim();
      if (!id || !code) continue;
      const key = `ticket:${id}`;
      if (seen.has(key)) continue;
      seen.add(key);
      entries.push({ key, label: code, kind: "ticket", value: id });
    }
    for (const label of suggestedLabelSeeds) {
      const value = String(label || "").trim();
      if (!value) continue;
      const key = `label:${value.toLowerCase()}`;
      if (seen.has(key)) continue;
      seen.add(key);
      entries.push({ key, label: value, kind: "label", value });
    }
    return entries;
  }, [suggestedExistingGroups, suggestedExistingTickets, suggestedLabelSeeds]);

  useEffect(() => {
    if (!createGroupName && suggestedGroupName) setCreateGroupName(suggestedGroupName);
  }, [createGroupName, suggestedGroupName]);

  useEffect(() => {
    if (!createTicketTitle) {
      const next = String(selectedEmail?.subject || "").trim() || (suggestedGroupName ? `Caso ${suggestedGroupName}` : "Ticket");
      setCreateTicketTitle(next);
    }
  }, [createTicketTitle, selectedEmail?.subject, suggestedGroupName]);

  const currentEmailPayload = useMemo<RelevantEmailPayload>(() => ({
    itemId: String(selectedEmail?.itemId || currentContext.itemId || "").trim() || undefined,
    internetMessageId: String(selectedEmail?.internetMessageId || currentContext.internetMessageId || "").trim() || undefined,
    conversationId: String(selectedEmail?.conversationId || currentContext.conversationId || "").trim() || undefined,
    subject: String(selectedEmail?.subject || currentContext.subject || "").trim() || undefined,
    fromEmail: String(selectedEmail?.fromEmail || currentContext.fromEmail || "").trim() || undefined,
    fromName: String(selectedEmail?.fromName || currentContext.fromName || "").trim() || undefined,
    receivedAtIso: String(selectedEmail?.receivedAtIso || selectedEmail?.messageDateIso || currentContext.receivedAtIso || "").trim() || undefined,
    messageDateIso: String(selectedEmail?.messageDateIso || selectedEmail?.receivedAtIso || currentContext.receivedAtIso || "").trim() || undefined,
    bodyText: String(selectedEmail?.bodyText || "").trim() || undefined,
    bodyHtml: String(selectedEmail?.bodyHtml || "").trim() || undefined,
    attachments: selectedEmailAttachments.map((attachment) => ({
      key: attachment.key,
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId,
      content: attachment.content,
      storageProvider: (attachment as any).storageProvider,
      storageBasePath: (attachment as any).storageBasePath,
      storagePathHint: (attachment as any).storagePathHint,
      documentState: (attachment as any).documentState,
      hasContent: (attachment as any).hasContent === true || Boolean(String(attachment.content || "").trim()),
      isHidden: typeof (attachment as any).isHidden === "boolean" ? (attachment as any).isHidden : undefined,
    })),
  }), [currentContext.conversationId, currentContext.fromEmail, currentContext.fromName, currentContext.internetMessageId, currentContext.itemId, currentContext.receivedAtIso, currentContext.subject, selectedEmail?.bodyHtml, selectedEmail?.bodyText, selectedEmail?.conversationId, selectedEmail?.fromEmail, selectedEmail?.fromName, selectedEmail?.internetMessageId, selectedEmail?.itemId, selectedEmail?.messageDateIso, selectedEmail?.receivedAtIso, selectedEmail?.subject, selectedEmailAttachments]);

  useEffect(() => {
    const selectedKey = String(selectedEmailKey || "").trim();
    if (!selectedEmail || !selectedKey || loading) return;
    const hasBody = Boolean(String(selectedEmail.bodyText || "").trim() || String(selectedEmail.bodyHtml || "").trim());
    const hasAttachments = Array.isArray(selectedEmail.attachments) && selectedEmail.attachments.length > 0;
    const hasHydratedAttachments = hasHydratedAttachmentCollection(selectedEmail);
    const hasPersistedIdentity = Boolean(String(selectedEmail.id || selectedEmail.emailKey || "").trim());
    const needsHydration = !selectedEmailInRelatedContext || !hasPersistedIdentity || !hasBody || !hasHydratedAttachments;
    if (!needsHydration) return;
    const hydrationSignature = [
      selectedKey,
      hasPersistedIdentity ? "persisted" : "seed",
      hasBody ? "body" : "no-body",
      hasAttachments ? `att:${selectedEmail.attachments?.length || 0}:${hasHydratedAttachments ? "ready" : "pending"}` : "no-att",
    ].join("|");
    if (hydratedEmailKeysRef.current.has(hydrationSignature)) return;

    hydratedEmailKeysRef.current.add(hydrationSignature);
    void refreshSelectedEmailContext(buildRelevantEmailPayloadFromRelatedEmail(selectedEmail) || currentEmailPayload)
      .catch(() => {
        hydratedEmailKeysRef.current.delete(hydrationSignature);
      });
  }, [currentEmailPayload, loading, selectedEmail, selectedEmailInRelatedContext, selectedEmailKey]);

  const similarCases = useMemo(() => {
    if (!selectedEmail) return [];
    const selectedKey = makeEmailKey(selectedEmail);
    const selectedPartner = normalizeSearchValue(`${derivePartnerName(selectedEmail)} ${selectedEmail.fromEmail || ""}`);
    const selectedGroups = new Set(selectedEmailGroups.map((group) => group.id));
    const selectedTickets = new Set(selectedEmailTicketIds);
    const emailUniverse = classificationKnownEmails;
    return emailUniverse
      .filter((email) => makeEmailKey(email) && makeEmailKey(email) !== selectedKey)
      .map((email) => {
        const key = makeEmailKey(email);
        const text = buildEmailCorpus(email);
        const matchedRefs = matchReferenceSet(text, documentReferences);
        const meta = emailContextMeta.get(key) || { groupIds: [], labels: [], ticketIds: [] };
        const candidateGroups = getEmailGroupRelations(email);
        const overlapGroups = meta.groupIds.filter((groupId) => selectedGroups.has(groupId));
        const overlapTickets = meta.ticketIds.filter((ticketId) => selectedTickets.has(ticketId));
        const candidatePartner = normalizeSearchValue(`${derivePartnerName(email)} ${email.fromEmail || ""}`);
        const samePartner = Boolean(selectedPartner && candidatePartner && (candidatePartner.includes(selectedPartner) || selectedPartner.includes(candidatePartner)));
        const sameType = detectCaseType(text) === detectedCaseType && detectedCaseType !== "geral";
        const score =
          matchedRefs.length * 140
          + overlapGroups.length * 36
          + overlapTickets.length * 42
          + (samePartner ? 18 : 0)
          + (sameType ? 8 : 0);
        return {
          email,
          score,
          matchedRefs,
          candidateGroups,
          candidateTickets: contextualTickets.filter((ticket) => meta.ticketIds.includes(ticket.id)).slice(0, 2),
          candidateLabels: meta.labels.slice(0, 3),
        };
      })
      .filter((entry) => entry.score > 0)
      .sort((a, b) => b.score - a.score || String(b.email.messageDateIso || b.email.receivedAtIso || "").localeCompare(String(a.email.messageDateIso || a.email.receivedAtIso || "")))
      .slice(0, 6);
  }, [classificationKnownEmails, detectedCaseType, documentReferences, emailContextMeta, getEmailGroupRelations, selectedEmail, selectedEmailGroups, selectedEmailTicketIds, contextualTickets]);
  const canonicalSelectedTicketId = useMemo(() => {
    const fromClassificationMeta = normalizeComparableString((selectedEmail?.classificationMeta as any)?.ticketId);
    if (fromClassificationMeta) return fromClassificationMeta;
    if (selectedEmailTicketIds.length === 1) return normalizeComparableString(selectedEmailTicketIds[0]);
    return "";
  }, [selectedEmail?.classificationMeta, selectedEmailTicketIds]);
  const resolvedSelectedTicketId = useMemo(
    () => normalizeComparableString(selectionTouched.ticket ? selectedTicketId : (canonicalSelectedTicketId || selectedTicketId)),
    [canonicalSelectedTicketId, selectedTicketId, selectionTouched.ticket]
  );
  const selectedTicket = useMemo(() => {
    if (!resolvedSelectedTicketId) return null;
    return canonicalTicketChoices.find((ticket) => ticket.id === resolvedSelectedTicketId)
      || relatedTickets.find((ticket) => ticket.id === resolvedSelectedTicketId)
      || (selectionTouched.ticket ? ticketSearchResults.find((ticket) => ticket.id === resolvedSelectedTicketId) || null : null);
  }, [canonicalTicketChoices, relatedTickets, resolvedSelectedTicketId, selectionTouched.ticket, ticketSearchResults]);
  const principalGroup = useMemo(() => (effectivePrincipalGroupId ? groupMap.get(effectivePrincipalGroupId) || null : null), [effectivePrincipalGroupId, groupMap]);
  const favoritePrincipalGroups = useMemo(
    () => favoriteGroupIds
      .map((groupId) => businessGroups.find((group) => group.id === groupId) || null)
      .filter(Boolean) as LinkGroupEntry[],
    [businessGroups, favoriteGroupIds]
  );
  const favoriteReferenceGroups = useMemo(
    () => favoritePrincipalGroups.filter((group) => group.id !== effectivePrincipalGroupId).slice(0, 6),
    [effectivePrincipalGroupId, favoritePrincipalGroups]
  );
  const normalizedPrincipalSearch = useMemo(() => normalizeSearchValue(principalSearch), [principalSearch]);
  const normalizedReferenceSearch = useMemo(() => normalizeSearchValue(referenceSearch), [referenceSearch]);
  const exactPrincipalSearchGroup = useMemo(
    () =>
      normalizedPrincipalSearch
        ? businessGroups.find((group) => normalizeSearchValue(String(group.name || "")) === normalizedPrincipalSearch) || null
        : null,
    [businessGroups, normalizedPrincipalSearch]
  );
  const exactReferenceSearchGroup = useMemo(
    () =>
      normalizedReferenceSearch
        ? businessGroups.find((group) =>
          group.id !== effectivePrincipalGroupId
          && normalizeSearchValue(String(group.name || "")) === normalizedReferenceSearch
        ) || null
        : null,
    [businessGroups, effectivePrincipalGroupId, normalizedReferenceSearch]
  );
  const principalSearchResults = useMemo(() => {
    if (!normalizedPrincipalSearch) return [];
    return filteredPrincipalGroups.slice(0, 6);
  }, [filteredPrincipalGroups, normalizedPrincipalSearch]);
  const referenceSearchResults = useMemo(() => {
    if (!normalizedReferenceSearch) return [];
    return filteredReferenceGroups.slice(0, 6);
  }, [filteredReferenceGroups, normalizedReferenceSearch]);
  const principalCanCreate = useMemo(
    () => Boolean(String(principalSearch || "").trim() && !exactPrincipalSearchGroup),
    [exactPrincipalSearchGroup, principalSearch]
  );
  const referenceCanCreate = useMemo(
    () => Boolean(String(referenceSearch || "").trim() && !exactReferenceSearchGroup),
    [exactReferenceSearchGroup, referenceSearch]
  );
  const principalSettingsTargetGroup = useMemo(
    () => exactPrincipalSearchGroup || principalGroup || null,
    [exactPrincipalSearchGroup, principalGroup]
  );
  const referenceGroups = useMemo(
    () => effectiveReferenceGroupIds.map((groupId) => groupMap.get(groupId)).filter(Boolean) as LinkGroupEntry[],
    [effectiveReferenceGroupIds, groupMap]
  );
  const referenceSettingsTargetGroup = useMemo(() => {
    if (exactReferenceSearchGroup) return exactReferenceSearchGroup;
    if (referenceGroups.length === 1) return referenceGroups[0];
    return null;
  }, [exactReferenceSearchGroup, referenceGroups]);
  const manageableGroups = useMemo(() => {
    const rows = new Map<string, LinkGroupEntry>();
    for (const group of contextualGroups) {
      if (!group?.id) continue;
      rows.set(group.id, group);
    }
    if (principalGroup?.id) rows.set(principalGroup.id, principalGroup);
    for (const group of referenceGroups) {
      if (!group?.id) continue;
      rows.set(group.id, group);
    }
    return Array.from(rows.values()).sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt"));
  }, [contextualGroups, principalGroup, referenceGroups]);
  const selectedManagedGroup = useMemo(
    () => (managedGroupId ? manageableGroups.find((group) => group.id === managedGroupId) || null : null),
    [manageableGroups, managedGroupId]
  );
  const caseTitle = useMemo(
    () => principalGroup?.name || selectedManagedGroup?.name || currentCaseBusinessGroups[0]?.name || "Caso sem grupo",
    [currentCaseBusinessGroups, principalGroup?.name, selectedManagedGroup?.name]
  );
  const caseClient = useMemo(
    () => formatChipValue(
      principalGroup?.contacts?.[0]?.company
        || principalGroup?.contacts?.[0]?.name
        || selectedManagedGroup?.contacts?.[0]?.company
        || selectedManagedGroup?.contacts?.[0]?.name
        || selectedEmail?.fromName
        || selectedEmail?.fromEmail,
      "Sem cliente"
    ),
    [
      principalGroup?.contacts,
      selectedManagedGroup?.contacts,
      selectedEmail?.fromEmail,
      selectedEmail?.fromName,
    ]
  );
  const caseBrand = useMemo(
    () => formatChipValue(principalGroup?.entities?.[0]?.name || selectedManagedGroup?.entities?.[0]?.name, "Sem marca"),
    [principalGroup?.entities, selectedManagedGroup?.entities]
  );
  const caseState = useMemo(
    () => formatGroupStatusLabel(principalGroup?.status || selectedManagedGroup?.status || ""),
    [principalGroup?.status, selectedManagedGroup?.status]
  );
  const buildResolvedApplySelectionForTargets = useCallback(
    (targetEmails: RelatedEmailEntry[]) => buildResolvedStudioApplySelection({
      targetEmails,
      principalGroupId: effectivePrincipalGroupId,
      principalGroup,
      referenceGroupIds: effectiveReferenceGroupIds,
      referenceGroups,
      selectedLabels,
      inheritedLabels,
      selectedLabelStates,
      categorizedLabelNames: categorizableLabels,
      selectedTicketId,
      selectedSeriesId,
      selectedTicket,
      ticketStatusDraft,
      classificationMetaDraft,
      existingSelectedEmailGroupIds: selectedEmailGroups.map((group) => String(group.id || "").trim()).filter(Boolean),
      existingSelectedEmailTicketIds: selectedEmailTicketIds,
      existingSelectedEmailLabels: selectedEmailStoredLabels,
      existingSelectedEmailStatus: selectedEmail?.status,
    }),
    [
      categorizableLabels,
      classificationMetaDraft,
      effectivePrincipalGroupId,
      effectiveReferenceGroupIds,
      inheritedLabels,
      principalGroup,
      referenceGroups,
      selectedEmail?.status,
      selectedEmailGroups,
      selectedEmailStoredLabels,
      selectedEmailTicketIds,
      selectedLabelStates,
      selectedLabels,
      selectedSeriesId,
      selectedTicket,
      selectedTicketId,
      ticketStatusDraft,
    ]
  );
  const resolvedApplySelection = useMemo(
    () => buildResolvedApplySelectionForTargets(defaultApplyTargetEmails),
    [buildResolvedApplySelectionForTargets, defaultApplyTargetEmails]
  );
  const canApplyClassification = useMemo(
    () => resolvedApplySelection.hasAnyClassificationValue,
    [resolvedApplySelection.hasAnyClassificationValue]
  );
  const hasPendingClassificationChanges = useMemo(() => {
    const snapshot = classificationDraftSnapshotRef.current;
    if (!snapshot) return false;
    return snapshot.principalGroupId !== principalGroupId
      || getComparableStringListSignature(snapshot.referenceGroupIds) !== getComparableStringListSignature(referenceGroupIds)
      || getComparableStringListSignature(snapshot.selectedLabels) !== getComparableStringListSignature(selectedLabels)
      || getComparableLabelDraftsSignature(snapshot.labelDrafts) !== getComparableLabelDraftsSignature(labelDrafts)
      || getComparableClassificationMetaSignature(snapshot.classificationMetaDraft) !== getComparableClassificationMetaSignature(classificationMetaDraft)
      || snapshot.selectedTicketId !== selectedTicketId
      || snapshot.selectedSeriesId !== selectedSeriesId
      || snapshot.ticketStatusDraft !== ticketStatusDraft
      || snapshot.createTicketTitle !== createTicketTitle;
  }, [
    classificationMetaDraft,
    createTicketTitle,
    labelDrafts,
    principalGroupId,
    referenceGroupIds,
    selectedLabels,
    selectedSeriesId,
    selectedTicketId,
    ticketStatusDraft,
  ]);
  const canApplyFromClassificationEditor = hasPendingClassificationChanges || canApplyClassification;
  const classificationEditorActive = section === "classification" && classificationFocus !== "summary";
  const auxiliaryEditorActive = section === "labels" || section === "filters" || section === "groups";
  const classificationCardTitle = useMemo(() => {
    if (section === "classification") {
      if (classificationFocus === "principal") return "Grupo principal";
      if (classificationFocus === "references") return "Referencias";
      if (classificationFocus === "labels") return "Etiquetas";
      if (classificationFocus === "ticket") return "Ticket";
      return "Resumo";
    }
    if (section === "labels") return "Etiquetas";
    if (section === "filters") return "Filtros";
    if (section === "groups") return "Grupos";
    return "Classificacao";
  }, [classificationFocus, section]);
  const effectiveTicketStatus = useMemo(
    () => String(ticketStatusDraft || selectedTicket?.status || "").trim(),
    [selectedTicket?.status, ticketStatusDraft]
  );
  const classificationSummaryTiles = useMemo(
    () => {
      const ticketCodes = relatedTickets.map((ticket) => String(ticket.code || "").trim()).filter(Boolean);
      const ticketSeriesPrefix = selectedSeriesId ? ticketSeries.find((entry) => entry.id === selectedSeriesId)?.prefix || "" : "";
      const ticketValue = selectedTicket?.code
        || (ticketCodes.length ? ticketCodes.join(", ") : "")
        || (selectedSeriesId ? (ticketSeriesPrefix ? `${ticketSeriesPrefix} (novo)` : "Novo ticket") : "")
        || "--";
      const principalStatusValue = principalGroup?.status ? formatGroupStatusLabel(principalGroup.status) : "";
      const ticketStatusValue = effectiveTicketStatus ? formatTicketStatusLabel(effectiveTicketStatus) : "";
      const referenceSummaryValue = referenceGroups.length ? referenceGroups.map((group) => group.name || group.id).join(", ") : "--";
      return [
        {
          key: "principal" as const,
          title: "Grupo principal",
          value: principalGroup?.name || "Sem grupo principal",
          description: classificationMetaDraft.principalStatusEnabled ? principalStatusValue || "Sem estado ativo" : "Sem estado ativo",
          onClick: () => openClassificationEditor("principal"),
        },
        {
          key: "labels" as const,
          title: "Etiquetas",
          value: selectedLabels.length ? selectedLabels.join(", ") : "Sem etiquetas",
          description: selectedLabels.length ? `${selectedLabels.length} atribuida(s)` : "Sem atribuicoes estruturadas",
          onClick: () => openClassificationEditor("labels"),
        },
        {
          key: "ticket" as const,
          title: "Ticket",
          value: ticketValue,
          description: classificationMetaDraft.ticketStatusEnabled ? ticketStatusValue || "Sem estado ativo" : "Sem seguimento ligado",
          onClick: () => openClassificationEditor("ticket"),
        },
        {
          key: "references" as const,
          title: "Referencias",
          value: referenceSummaryValue,
          description: referenceGroups.length ? `${referenceGroups.length} referencia(s)` : "Disponivel no modo avancado",
          onClick: () => openClassificationEditor("references"),
        },
      ];
    },
    [
      classificationMetaDraft.principalStatusEnabled,
      classificationMetaDraft.ticketStatusEnabled,
      principalGroup?.name,
      principalGroup?.status,
      referenceGroups.length,
      selectedLabels,
      selectedSeriesId,
      selectedTicket?.status,
      selectedTicket?.code,
      ticketStatusDraft,
      ticketSeries,
      relatedTickets,
    ]
  );
  const previewHasDocument = Boolean(selectedAttachmentPreview);

  const managedGroupContactCandidates = useMemo(() => {
    const caseEmails = dedupeEmails([
      ...(selectedEmail ? [selectedEmail] : []),
      ...managedGroupEmails,
      ...classificationContextEmails,
    ]);
    const candidates = caseEmails.flatMap((email) => {
      const company = inferCompanyName(email.fromName, email.fromEmail);
      return [{
        key: String(email.fromEmail || "").trim().toLowerCase() || `${normalizeSearchValue(email.fromName || "")}|${normalizeSearchValue(company)}`,
        name: String(email.fromName || "").trim() || String(email.fromEmail || "").trim(),
        email: String(email.fromEmail || "").trim().toLowerCase() || undefined,
        company: company || undefined,
        source: "email",
      }];
    });
    return dedupeGroupContacts([
      ...(selectedManagedGroup?.contacts || []),
      ...candidates,
    ]);
  }, [classificationContextEmails, managedGroupEmails, selectedEmail, selectedManagedGroup?.contacts]);

  const managedGroupEntityCandidates = useMemo(() => {
    const contactEntities = managedGroupContactCandidates
      .map((contact) => ({
        key: normalizeSearchValue(contact.company || inferCompanyName(contact.name, contact.email)),
        name: String(contact.company || inferCompanyName(contact.name, contact.email) || "").trim(),
        kind: "empresa",
        source: contact.source || "email",
      }))
      .filter((entity) => entity.name);
    const groupEntities = manageableGroups
      .filter((group) => group.id === managedGroupId || selectedEmailGroups.some((entry) => entry.id === group.id))
      .map((group) => ({
        key: normalizeSearchValue(group.name),
        name: group.name,
        kind: "grupo",
        source: "grupo",
      }));
    return dedupeGroupEntities([
      ...(selectedManagedGroup?.entities || []),
      ...contactEntities,
      ...groupEntities,
    ]);
  }, [manageableGroups, managedGroupContactCandidates, managedGroupId, selectedEmailGroups, selectedManagedGroup?.entities]);

  const filteredManagedGroupContacts = useMemo(() => {
    const q = normalizeSearchValue(managedContactSearch);
    if (!q) return managedGroupContactCandidates;
    return managedGroupContactCandidates.filter((contact) =>
      [contact.name, contact.email, contact.company].some((value) => normalizeSearchValue(String(value || "")).includes(q))
    );
  }, [managedContactSearch, managedGroupContactCandidates]);

  const filteredManagedGroupEntities = useMemo(() => {
    const q = normalizeSearchValue(managedEntitySearch);
    if (!q) return managedGroupEntityCandidates;
    return managedGroupEntityCandidates.filter((entity) =>
      [entity.name, entity.kind, entity.source].some((value) => normalizeSearchValue(String(value || "")).includes(q))
    );
  }, [managedEntitySearch, managedGroupEntityCandidates]);
  useEffect(() => {
    setManagedGroupId((current) => {
      if (current && manageableGroups.some((group) => group.id === current)) return current;
      return principalGroupId || manageableGroups[0]?.id || "";
    });
  }, [manageableGroups, principalGroupId]);
  const inheritedLabels = useMemo(
    () =>
      mergeLabels(
        mergeLabels(
          principalGroup?.labels || [],
          referenceGroups.flatMap((group) => group.labels || [])
        ),
        mergeLabels(
          selectedTicket?.labels || [],
          relatedTickets.flatMap((ticket) => ticket.labels || [])
        )
      ),
    [principalGroup?.labels, referenceGroups, relatedTickets, selectedTicket?.labels]
  );
  const selectedEmailStoredLabels = useMemo(
    () => Array.isArray(selectedEmail?.labels) ? selectedEmail.labels.map((label) => String(label || "").trim()).filter(Boolean) : [],
    [selectedEmail?.labels]
  );
  const selectedEmailRemovedInheritedLabels = useMemo(
    () => Array.isArray(selectedEmail?.removedInheritedLabels) ? selectedEmail.removedInheritedLabels.map((label) => String(label || "").trim()).filter(Boolean) : [],
    [selectedEmail?.removedInheritedLabels]
  );
  const selectedEmailLabelStates = useMemo(
    () => selectedEmail?.labelStates && typeof selectedEmail.labelStates === "object"
      ? Object.fromEntries(
          Object.entries(selectedEmail.labelStates)
            .map(([label, status]) => [String(label || "").trim(), String(status || "").trim()])
            .filter(([label, status]) => label && status)
        ) as Record<string, string>
      : {},
    [selectedEmail?.labelStates]
  );
  const selectedEmailCategorizedLabelNames = useMemo(
    () => Array.isArray(selectedEmail?.classificationMeta?.categorizedLabelNames)
      ? selectedEmail.classificationMeta.categorizedLabelNames.map((label) => String(label || "").trim()).filter(Boolean)
      : [],
    [selectedEmail?.classificationMeta?.categorizedLabelNames]
  );
  const summaryLabels = useMemo(
    () => selectedLabels,
    [selectedLabels]
  );
  const categorizableLabels = useMemo(
    () => summaryLabels.filter((label) => labelDrafts[label]?.categorize === true),
    [labelDrafts, summaryLabels]
  );
  const selectedLabelStates = useMemo(() => {
    const entries: Record<string, EmailLabelStatus> = {};
    for (const label of selectedLabels) {
      const draft = labelDrafts[label];
      if (!draft?.hasStatus || !draft.status) continue;
      entries[label] = draft.status;
    }
    return entries;
  }, [labelDrafts, selectedLabels]);
  const selectedLabelStatuses = useMemo(
    () => Array.from(new Set(Object.values(selectedLabelStates).filter(Boolean))),
    [selectedLabelStates]
  );
  const selectedLabelSharedStatus = useMemo(
    () => (selectedLabelStatuses.length === 1 ? selectedLabelStatuses[0] : ""),
    [selectedLabelStatuses]
  );
  const emailStatusSummary = useMemo(
    () => selectedLabelStatuses.length ? selectedLabelStatuses.map((entry) => formatEmailLabelStatus(entry)).join(", ") : "--",
    [selectedLabelStatuses]
  );
  const labelStateSummary = useMemo(
    () => Object.entries(selectedLabelStates).map(([label, status]) => `${label} (${formatEmailLabelStatus(status)})`),
    [selectedLabelStates]
  );
  const referenceGroupSummary = useMemo(
    () => (referenceGroups.length ? referenceGroups.map((group) => group.name || group.id).join(", ") : "--"),
    [referenceGroups]
  );
  const principalGroupStatusLabel = useMemo(
    () => principalGroup?.status ? formatGroupStatusLabel(principalGroup.status) : "",
    [principalGroup?.status]
  );
  const referenceGroupStatusEntries = useMemo(
    () =>
      referenceGroups
        .map((group) => ({
          id: group.id,
          name: group.name || group.id,
          status: formatGroupStatusLabel(group.status),
          hasStatus: Boolean(String(group.status || "").trim()),
        }))
        .filter((entry) => entry.hasStatus),
    [referenceGroups]
  );
  const ticketStatusLabel = useMemo(
    () => effectiveTicketStatus ? formatTicketStatusLabel(effectiveTicketStatus) : "",
    [effectiveTicketStatus]
  );
  const ticketSummary = useMemo(() => {
    if (selectedTicket?.code) return selectedTicket.code;
    if (relatedTickets.length) {
      const codes = relatedTickets.map((ticket) => String(ticket.code || "").trim()).filter(Boolean);
      if (codes.length) return codes.join(", ");
    }
    if (selectedSeriesId) {
      const series = ticketSeries.find((entry) => entry.id === selectedSeriesId);
      return series?.prefix ? `${series.prefix} (novo)` : "Novo ticket";
    }
    return "--";
  }, [relatedTickets, selectedSeriesId, selectedTicket?.code, ticketSeries]);

  useEffect(() => {
    setManagedGroupDescription(String(selectedManagedGroup?.description || "").trim());
    setManagedGroupNotes(String(selectedManagedGroup?.notes || "").trim());
    setManagedGroupContacts(dedupeGroupContacts(selectedManagedGroup?.contacts || []));
    setManagedGroupEntities(dedupeGroupEntities(selectedManagedGroup?.entities || []));
    setManagedContactSearch("");
    setManagedEntitySearch("");
  }, [selectedManagedGroup?.contacts, selectedManagedGroup?.description, selectedManagedGroup?.entities, selectedManagedGroup?.id, selectedManagedGroup?.notes]);

  useEffect(() => {
    let cancelled = false;
    const groupId = String(managedGroupId || "").trim();
    if (!groupId) {
      setManagedGroupEmails([]);
      setManagedGroupDocuments([]);
      return () => { cancelled = true; };
    }
    void (async () => {
      setManagedGroupLoading(true);
      try {
        const [emails, documents] = await Promise.all([
          getGroupEmails(groupId),
          getGroupDocuments(groupId),
        ]);
        if (cancelled) return;
        setManagedGroupEmails(Array.isArray(emails) ? emails : []);
        setManagedGroupDocuments(Array.isArray(documents) ? documents : []);
      } catch (loadError: any) {
        if (!cancelled) setStatus(loadError?.message || "Nao foi possivel carregar o dossier do grupo.");
      } finally {
        if (!cancelled) setManagedGroupLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [managedGroupId]);

  useEffect(() => {
    if (selectionTouched.ticket) return;
    if (canonicalSelectedTicketId) {
      setSelectedTicketId(canonicalSelectedTicketId);
      return;
    }
    if (!selectedEmailTicketIds.length) {
      setSelectedTicketId((current) => (
        current && canonicalTicketChoices.some((ticket) => ticket.id === current)
          ? current
          : ""
      ));
    }
  }, [canonicalSelectedTicketId, canonicalTicketChoices, selectedEmailTicketIds.length, selectionTouched.ticket]);

  useEffect(() => {
    if (!selectedSeriesId || !selectedTicketId) return;
    setSelectedTicketId("");
  }, [selectedSeriesId, selectedTicketId]);

  useEffect(() => {
    if (!selectedTicketId || !selectedSeriesId) return;
    setSelectedSeriesId("");
  }, [selectedTicketId, selectedSeriesId]);

  useEffect(() => {
    if (selectedTicketId) {
      const nextStatus = String(selectedTicket?.status || "").trim();
      setTicketStatusDraft(nextStatus);
      return;
    }
    if (selectedSeriesId) {
      setTicketStatusDraft("");
      return;
    }
    setTicketStatusDraft("");
  }, [selectedSeriesId, selectedTicketId]);

  useEffect(() => {
    const currentSelectedEmail = selectedEmailRef.current;
    if (currentSelectedEmail) {
      rehydrateClassificationEditorFromCaseEmail(currentSelectedEmail);
      return;
    }
    setSelectionTouched({ principal: false, references: false, ticket: false });
    setPrincipalGroupId("");
    setPrincipalSearch("");
    setReferenceGroupIds([]);
    setReferenceSearch("");
    setSelectedLabels([]);
    setLabelDrafts({});
    setClassificationMetaDraft(normalizeClassificationMetaDraft(null));
    setSelectedTicketId("");
    setSelectedSeriesId("");
    setTicketSearch("");
    setTicketSearchResults([]);
  }, [rehydrateClassificationEditorFromCaseEmail, selectedEmailKey]);

  useEffect(() => {
    if (!labelCatalogReady) return;
    if (selectedLabels.length || (!inheritedLabels.length && !selectedEmailStoredLabels.length && !selectedEmailRemovedInheritedLabels.length)) return;
    const visibleInherited = inheritedLabels.filter((label) => !selectedEmailRemovedInheritedLabels.includes(label));
    const seedLabels = mergeLabels(visibleInherited, selectedEmailStoredLabels);
    setSelectedLabels(seedLabels);
    setLabelDrafts((current) => {
      const next = { ...current };
      for (const label of seedLabels) {
        next[label] = createLabelDraftFromCatalog(
          findGroupLabelCatalogEntry(labelCatalogEntries, label),
          current[label],
          selectedEmailLabelStates[label],
          selectedEmailCategorizedLabelNames.includes(label)
        );
      }
      return next;
    });
  }, [inheritedLabels, labelCatalogEntries, labelCatalogReady, selectedEmailCategorizedLabelNames, selectedEmailLabelStates, selectedEmailRemovedInheritedLabels, selectedEmailStoredLabels, selectedLabels.length]);

  useEffect(() => {
    if (!selectedLabels.length) return;
    setLabelDrafts((current) => {
      let changed = false;
      const next = { ...current };
      for (const label of selectedLabels) {
        const resolved = createLabelDraftFromCatalog(
          findGroupLabelCatalogEntry(labelCatalogEntries, label),
          current[label],
          selectedEmailLabelStates[label],
          selectedEmailCategorizedLabelNames.includes(label)
        );
        const previous = current[label];
        if (
          !previous
          || previous.categorize !== resolved.categorize
          || previous.hasStatus !== resolved.hasStatus
          || previous.status !== resolved.status
        ) {
          next[label] = resolved;
          changed = true;
        }
      }
      return changed ? next : current;
    });
  }, [labelCatalogEntries, selectedEmailCategorizedLabelNames, selectedEmailLabelStates, selectedLabels]);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      if (!selectedEmailIsCurrent) {
        if (!cancelled) {
          setOutlookLabelCategories((current) => (current.length ? [] : current));
        }
        return;
      }
      try {
        const snapshot = await getManagedOutlookCategorySnapshot(
          mergeLabels(
            mergeLabels(labelCatalog, selectedEmailStoredLabels),
            selectedEmailRemovedInheritedLabels
          )
        );
        if (cancelled) return;
        const labels = (snapshot?.labelNames || []).map((label) => String(label || "").trim()).filter(Boolean);
        setOutlookLabelCategories((current) => (areStringListsEqual(current, labels) ? current : labels));
      } catch {
        if (!cancelled) {
          setOutlookLabelCategories((current) => (current.length ? [] : current));
        }
      }
    })();
    return () => { cancelled = true; };
  }, [labelCatalog, selectedEmailIsCurrent, selectedEmailRemovedInheritedLabels, selectedEmailStoredLabels]);

  useEffect(() => {
    if (!outlookLabelCategories.length) return;
    setLabelDrafts((current) => {
      let changed = false;
      const next = { ...current };
      for (const label of outlookLabelCategories) {
        const resolved = {
          categorize: true,
          hasStatus: current[label]?.hasStatus ?? false,
          status: current[label]?.status,
        };
        const previous = current[label];
        if (
          !previous
          || previous.categorize !== resolved.categorize
          || previous.hasStatus !== resolved.hasStatus
          || previous.status !== resolved.status
        ) {
          next[label] = resolved;
          changed = true;
        }
      }
      return changed ? next : current;
    });
    setSelectedLabels((current) => {
      const next = mergeLabels(current, outlookLabelCategories);
      return areStringListsEqual(current, next) ? current : next;
    });
  }, [outlookLabelCategories]);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (!closed) window.close();
  }

  async function refreshSelectedEmailContext(targetEmailPayload?: RelevantEmailPayload | null) {
    const lookup = targetEmailPayload || currentEmailPayload;
    const related = await getRelatedEmailContext({
      conversationId: lookup.conversationId,
      internetMessageId: lookup.internetMessageId,
      itemId: lookup.itemId,
      subject: lookup.subject,
      fromEmail: lookup.fromEmail,
      fromName: lookup.fromName,
      receivedAtIso: lookup.receivedAtIso,
    });
    const contextualEmails = dedupeEmails([
      ...(related.email ? [related.email] : []),
      ...(related.emails || []),
    ]);
    setAllGroups((current) => mergeGroupEntryLists(current, related.groups || []));
    setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
    setRelatedTickets((current) => {
      const nextTickets = Array.isArray(related.tickets) ? related.tickets : [];
      const preservedSelectedTicket = resolvedSelectedTicketId
        ? current.find((ticket) => ticket.id === resolvedSelectedTicketId) || null
        : null;
      if (preservedSelectedTicket && !nextTickets.some((ticket) => ticket.id === preservedSelectedTicket.id)) {
        return [preservedSelectedTicket, ...nextTickets];
      }
      return nextTickets;
    });
    setRelatedEmails(contextualEmails);
    setKnownEmails((current) => dedupeEmails([...contextualEmails, ...current]));
    mergeEmailsIntoClassificationCase(contextualEmails);
    return related;
  }

  function toggleTargetEmailKey(emailKey: string) {
    const key = String(emailKey || "").trim();
    if (!key) return;
    setSelectedTargetEmailKeys((current) =>
      current.includes(key)
        ? current.filter((entry) => entry !== key)
        : [...current, key]
    );
  }

  function selectAllVisibleEmails() {
    setSelectedTargetEmailKeys(visibleEmails.map((email) => makeEmailKey(email)).filter(Boolean));
  }

  function clearSelectedTargets() {
    setSelectedTargetEmailKeys(selectedEmailKey ? [selectedEmailKey] : []);
  }

  function toggleReferenceGroup(groupId: string) {
    if (!groupId || groupId === effectivePrincipalGroupId) return;
    setSelectionTouched((current) => ({ ...current, references: true }));
    setReferenceGroupIds((current) =>
      toggleReferenceGroupSelection(
        {
          principalGroupId: effectivePrincipalGroupId,
          referenceGroupIds: selectionTouched.references ? current : effectiveReferenceGroupIds,
        },
        groupId
      ).referenceGroupIds
    );
  }

  function clearPrincipalSelection() {
    setSelectionTouched((current) => ({ ...current, principal: true }));
    setPrincipalGroupId("");
    setPrincipalSearchValue("");
  }

  function setPrincipalSearchValue(value: string) {
    const nextValue = String(value || "");
    setPrincipalSearch(nextValue);
    setCreateGroupName(nextValue);
  }

  function setReferenceSearchValue(value: string) {
    setReferenceSearch(String(value || ""));
  }

  function selectPrincipalGroup(group: LinkGroupEntry | null) {
    if (!group?.id) return;
    const normalizedSelection = setPrincipalGroupSelection(
      {
        principalGroupId: effectivePrincipalGroupId,
        referenceGroupIds: effectiveReferenceGroupIds,
      },
      group.id
    );
    setSelectionTouched((current) => ({ ...current, principal: true }));
    setPrincipalGroupId(normalizedSelection.principalGroupId);
    setReferenceGroupIds(normalizedSelection.referenceGroupIds);
    setPrincipalSearchValue("");
  }

  function toggleExpandedQuickDocumentKey(attachmentKey: string) {
    const key = String(attachmentKey || "").trim();
    if (!key) return;
    setExpandedQuickDocumentKeys((current) =>
      current.includes(key)
        ? current.filter((entry) => entry !== key)
        : [key]
    );
  }

  function toggleFavoritePrincipalGroup(group: LinkGroupEntry) {
    if (!group?.id) return;
    const sameGroup = effectivePrincipalGroupId === group.id;
    if (sameGroup) {
      clearPrincipalSelection();
      return;
    }
    selectPrincipalGroup(group);
  }

  function toggleFavoriteReferenceGroup(group: LinkGroupEntry) {
    if (!group?.id) return;
    const sameGroup = effectiveReferenceGroupIds.includes(group.id);
    if (sameGroup) {
      toggleReferenceGroup(group.id);
      setReferenceSearchValue("");
      return;
    }
    toggleReferenceGroup(group.id);
    setReferenceSearchValue("");
  }

  function openManagedGroupFromPrincipal(group: LinkGroupEntry | null) {
    if (!group?.id) return;
    setManagedGroupId(group.id);
    setSection("groups");
  }

  function clearTicketSelection() {
    setSelectionTouched((current) => ({ ...current, ticket: true }));
    setSelectedTicketId("");
    setSelectedSeriesId("");
  }

  function applySuggestedGroup(groupId: string) {
    if (!groupId) return;
    if (effectivePrincipalGroupId === groupId) {
      clearPrincipalSelection();
      return;
    }
    if (effectiveReferenceGroupIds.includes(groupId)) {
      toggleReferenceGroup(groupId);
      return;
    }
    if (classificationFocus === "references") {
      setSelectionTouched((current) => ({ ...current, references: true }));
      setReferenceGroupIds((current) =>
        addReferenceGroupSelection(
          {
            principalGroupId: effectivePrincipalGroupId,
            referenceGroupIds: selectionTouched.references ? current : effectiveReferenceGroupIds,
          },
          groupId
        ).referenceGroupIds
      );
      setReferenceSearchValue("");
      return;
    }
    if (!effectivePrincipalGroupId || classificationFocus === "principal") {
      setSelectionTouched((current) => ({ ...current, principal: true }));
      setPrincipalGroupId(groupId);
      setPrincipalSearchValue("");
      return;
    }
    if (effectivePrincipalGroupId === groupId) {
      return;
    }
    setSelectionTouched((current) => ({ ...current, references: true }));
    setReferenceGroupIds((current) =>
      addReferenceGroupSelection(
        {
          principalGroupId: effectivePrincipalGroupId,
          referenceGroupIds: selectionTouched.references ? current : effectiveReferenceGroupIds,
        },
        groupId
      ).referenceGroupIds
    );
    setReferenceSearchValue("");
  }

  function applySuggestedTicket(ticketId: string) {
    if (!ticketId) return;
    if (selectedTicketId === ticketId) {
      clearTicketSelection();
      return;
    }
    setSelectionTouched((current) => ({ ...current, ticket: true }));
    setSelectedSeriesId("");
    setSelectedTicketId(ticketId);
  }

  function applySuggestedLabel(label: string) {
    const value = String(label || "").trim();
    if (!value) return;
    if (selectedLabels.includes(value)) {
      removeLabel(value);
      return;
    }
    addLabel(value);
  }

  function resolveSuggestionGroupId(suggestion: ReadingSuggestionChip) {
    if (suggestion.kind === "group") return suggestion.value;
    const normalized = normalizeSearchValue(String(suggestion.label || suggestion.value || ""));
    const match = businessGroups.find((group) => normalizeSearchValue(String(group.name || "")) === normalized);
    return String(match?.id || "").trim();
  }

  function resolveSuggestionTicketId(suggestion: ReadingSuggestionChip) {
    if (suggestion.kind === "ticket") return suggestion.value;
    const normalized = normalizeSearchValue(String(suggestion.label || suggestion.value || ""));
    const match = canonicalTicketChoices.find((ticket) => normalizeSearchValue(String(ticket.code || "")) === normalized);
    return String(match?.id || "").trim();
  }

  function isSuggestionActive(suggestion: ReadingSuggestionChip) {
    if (classificationFocus === "summary") return false;
    if (classificationFocus === "principal") {
      const suggestionText = normalizeSearchValue(String(suggestion.label || suggestion.value || "").trim());
      return Boolean(suggestionText && normalizedPrincipalSearch === suggestionText);
    }
    if (classificationFocus === "references") {
      const suggestionText = normalizeSearchValue(String(suggestion.label || suggestion.value || "").trim());
      return Boolean(suggestionText && normalizedReferenceSearch === suggestionText);
    }
    if (classificationFocus === "ticket") {
      const ticketId = resolveSuggestionTicketId(suggestion);
      return Boolean(ticketId && selectedTicketId === ticketId);
    }
    return selectedLabels.includes(String(suggestion.label || suggestion.value || "").trim());
  }

  function handleSuggestionToggle(suggestion: ReadingSuggestionChip) {
    if (classificationFocus === "summary") return;
    if (classificationFocus === "principal") {
      const suggestionText = String(suggestion.label || suggestion.value || "").trim();
      const normalizedSuggestion = normalizeSearchValue(suggestionText);
      if (!suggestionText) return;
      if (normalizedPrincipalSearch === normalizedSuggestion) {
        setPrincipalSearchValue("");
        return;
      }
      setPrincipalSearchValue(suggestionText);
      return;
    }
    if (classificationFocus === "references") {
      const suggestionText = String(suggestion.label || suggestion.value || "").trim();
      const normalizedSuggestion = normalizeSearchValue(suggestionText);
      if (!suggestionText) return;
      if (normalizedReferenceSearch === normalizedSuggestion) {
        setReferenceSearchValue("");
        return;
      }
      setReferenceSearchValue(suggestionText);
      return;
    }
    if (classificationFocus === "labels") {
      applySuggestedLabel(suggestion.label || suggestion.value);
      return;
    }
    if (classificationFocus === "ticket") {
      const ticketId = resolveSuggestionTicketId(suggestion);
      if (ticketId) applySuggestedTicket(ticketId);
      return;
    }
    const groupId = resolveSuggestionGroupId(suggestion);
    if (groupId) applySuggestedGroup(groupId);
  }

  function addLabel(label: string) {
    const value = String(label || "").trim();
    if (!value) return;
    setSelectedLabels((current) => current.includes(value) ? current : [...current, value]);
    setLabelDrafts((current) => current[value]
      ? current
      : {
          ...current,
          [value]: createLabelDraftFromCatalog(
            findGroupLabelCatalogEntry(labelCatalogEntries, value),
            undefined,
            selectedEmailLabelStates[value],
            selectedEmailCategorizedLabelNames.includes(value)
          ),
        });
    setLabelCatalogEntries((current) => {
      if (current.some((entry) => String(entry?.label || "").trim().toLowerCase() === value.toLowerCase())) {
        return current;
      }
      return [...current, { label: value, categorize: false, hasStatus: false }];
    });
    setLabelInput("");
  }

  function handleClassificationLabelSearchAction() {
    const rawValue = String(classificationLabelInput || "").trim();
    if (!rawValue) return;
    if (classificationLabelCanCreate) {
      addLabel(rawValue);
      return;
    }
    if (exactClassificationLabel) {
      if (selectedLabels.includes(exactClassificationLabel)) {
        removeLabel(exactClassificationLabel);
      } else {
        addLabel(exactClassificationLabel);
      }
    }
  }

  function updateLabelDraft(label: string, patch: Partial<LabelDraft>) {
    setLabelDrafts((current) => {
      const next: LabelDraft = {
        categorize: current[label]?.categorize ?? false,
        hasStatus: current[label]?.hasStatus ?? false,
        status: current[label]?.status,
        ...patch,
      };
      if (next.hasStatus && !next.status) next.status = "em_analise";
      if (!next.hasStatus) next.status = undefined;
      return { ...current, [label]: next };
    });
  }

  function removeLabel(label: string) {
    setSelectedLabels((current) => current.filter((entry) => entry !== label));
  }

  function updateClassificationMeta(patch: Partial<ClassificationMetaDraft>) {
    setClassificationMetaDraft((current) => {
      const next = { ...current, ...patch };
      if (!next.principalStatusEnabled) next.principalStatusCategorize = false;
      if (!next.referenceStatusEnabled) next.referenceStatusCategorize = false;
      if (!next.ticketStatusEnabled) next.ticketStatusCategorize = false;
      return next;
    });
  }

  async function handleCreateGroupAndLink(kind: "principal" | "referencia" = "principal", nameOverride?: string) {
    const name = String(nameOverride || createGroupName || (kind === "principal" ? principalSearch : referenceSearch) || "").trim();
    if (!name) {
      setStatus("Define primeiro o nome do grupo.");
      return;
    }
    setActionBusy(true);
    try {
      const created = await createLinkGroup({
        name,
        labels: selectedLabels,
        documentsEnabled: true,
      });
      await addEmailToLinkGroup(created.id, {
        ...currentEmailPayload,
        membershipKind: kind,
      });
      setAllGroups((current) => current.some((entry) => entry.id === created.id) ? current : [created, ...current]);
      if (kind === "principal") {
        setPrincipalGroupId(created.id);
        setPrincipalSearchValue(created.name);
      } else {
        setReferenceGroupIds((current) =>
          addReferenceGroupSelection(
            {
              principalGroupId,
              referenceGroupIds: current,
            },
            created.id
          ).referenceGroupIds
        );
        setReferenceSearchValue(created.name);
      }
      setManagedGroupId(created.id);
      setCreateGroupName("");
      await refreshSelectedEmailContext();
      setStatus(kind === "principal"
        ? `Grupo "${created.name}" criado e email ligado como principal.`
        : `Grupo "${created.name}" criado e email ligado como referencia.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel criar e ligar o grupo.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleCreateTicketAndLink() {
    if (!selectedSeriesId) {
      setStatus("Escolhe primeiro uma serie de ticket.");
      return;
    }
    setActionBusy(true);
    try {
      const groupIds = [effectivePrincipalGroupId, ...effectiveReferenceGroupIds].filter(Boolean);
      const ticket = await createGroupTicket({
        seriesId: selectedSeriesId,
        title: String(createTicketTitle || selectedEmail?.subject || "Ticket").trim(),
        description: String(selectedEmail?.bodyText || "").trim().slice(0, 4000),
        labels: selectedLabels,
        groupIds,
        email: {
          ...currentEmailPayload,
          labels: selectedLabels.filter((label) => !inheritedLabels.includes(label)),
          removedInheritedLabels: inheritedLabels.filter((label) => !selectedLabels.includes(label)),
          labelStates: selectedLabelStates,
          classificationMeta: classificationMetaDraft,
        },
        membershipKind: effectivePrincipalGroupId ? "principal" : "referencia",
      });
      setRelatedTickets((current) => [ticket, ...current.filter((entry) => entry.id !== ticket.id)]);
      setSelectionTouched((current) => ({ ...current, ticket: true }));
      setSelectedSeriesId("");
      setSelectedTicketId(ticket.id);
      await refreshSelectedEmailContext();
      setStatus(`Ticket ${ticket.code} criado e ligado ao email atual.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel criar o ticket.");
    } finally {
      setActionBusy(false);
    }
  }

  function toggleAttachmentPlan(attachmentKey: string, field: "analyze" | "save" | "forward", checked: boolean) {
    setAttachmentPlan((current) => ({
      ...current,
      [attachmentKey]: {
        analyze: current[attachmentKey]?.analyze ?? false,
        save: current[attachmentKey]?.save ?? false,
        forward: current[attachmentKey]?.forward ?? false,
        [field]: checked,
      },
    }));
  }

  async function handleSetSelectedAttachmentDocumentState(nextState: DocumentLifecycleState) {
    if (!selectedAttachmentPreviewEmail || !selectedAttachmentPreview) {
      setStatus("Escolhe primeiro um anexo para atualizar o estado documental.");
      return;
    }
    const attachmentKey = makeAttachmentKey(selectedAttachmentPreview);
    if (!attachmentKey) {
      setStatus("Nao foi possivel identificar o anexo selecionado.");
      return;
    }
    const updatedEmail = updateAttachmentStateOnEmail(selectedAttachmentPreviewEmail, attachmentKey, nextState);
    if (!updatedEmail) {
      setStatus("Nao foi possivel atualizar o estado documental deste anexo.");
      return;
    }
    setActionBusy(true);
    try {
      setRelatedEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
      setKnownEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
      mergeEmailIntoClassificationCase(updatedEmail);
      setAttachmentPlan((current) => ({
        ...current,
        [selectedAttachmentPreviewKey]: {
          analyze: nextState === "rejected" ? false : (current[selectedAttachmentPreviewKey]?.analyze ?? false),
          save: nextState === "rejected" ? false : (current[selectedAttachmentPreviewKey]?.save ?? false),
          forward: current[selectedAttachmentPreviewKey]?.forward ?? false,
        },
      }));
      const latestSettings = await getSettings().catch(() => null);
      const payload = buildRelevantEmailPayloadFromRelatedEmail(updatedEmail);
      if (payload) {
        await registerRelevantEmail({
          ...payload,
          ...buildAttachmentStorageOptions(latestSettings),
        });
      }
      setStatus(`Estado documental de "${selectedAttachmentPreview.name}" atualizado para ${formatDocumentLifecycleState(nextState)}.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel atualizar o estado documental do anexo.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleSaveSelectedAttachments() {
    if (!effectivePrincipalGroupId) {
      setStatus("Escolhe primeiro um grupo principal para guardar documentos.");
      return;
    }
    const docs = (
      await Promise.all(
        canonicalSelectedEmailAttachmentEntries
          .filter((entry) => attachmentPlan[entry.scopedKey]?.save)
          .map(async (entry) => {
            const attachment = entry.attachment;
            let contentBase64 = String(attachment.content || "").trim();
            const selectedEmailRemoteId = String(selectedEmail?.id || selectedEmail?.emailKey || "").trim();
            if (!contentBase64 && attachment.hasContent && selectedEmailRemoteId) {
              const remoteId = getStudioAttachmentRemoteId(attachment);
              if (remoteId) {
                try {
                  const remote = await getEmailAttachmentContentBase64(selectedEmailRemoteId, remoteId);
                  contentBase64 = String(remote.base64 || "").trim();
                } catch {
                  contentBase64 = "";
                }
              }
            }
            if (!contentBase64) return null;
            return {
              name: attachment.name,
              contentType: attachment.contentType,
              contentBase64,
              size: attachment.size,
              documentState: normalizeDocumentLifecycleState((attachment as any)?.documentState, "accepted"),
              sourceEmailKey: makeEmailKey(selectedEmail || {}),
              sourceItemId: currentEmailPayload.itemId,
              sourceInternetMessageId: currentEmailPayload.internetMessageId,
              sourceConversationId: currentEmailPayload.conversationId,
              sourceEmailSubject: currentEmailPayload.subject,
            };
          })
      )
    ).filter(Boolean);
    if (!docs.length) {
      setStatus("Nao ha anexos com conteudo selecionados para guardar.");
      return;
    }
    setActionBusy(true);
    try {
      const settings = await getSettings().catch(() => null);
      const storageOptions = buildAttachmentStorageOptions(settings);
      const storageProvider = String(storageOptions.attachmentStorageProvider || "cloud").trim();
      const storageBasePath = String(storageOptions.attachmentStorageBasePath || "").trim();
      const safeGroupName = String(principalGroup?.name || principalGroupId || "grupo")
        .trim()
        .replace(/[\\/:*?"<>|]+/g, "_");
      await saveGroupDocuments(principalGroupId, {
        documents: docs.map((doc: any) => ({
          ...doc,
          id: doc.id || doc.key || `new_doc_${Date.now()}_${Math.random().toString(36).slice(2)}`,
          storageProvider,
          storageBasePath,
          storagePathHint: safeGroupName && doc.name
            ? `${safeGroupName}/${String(doc.name || "").trim().replace(/[\\/:*?"<>|]+/g, "_")}`
            : undefined,
        } as GroupDocumentEntry)),
      });
      await refreshSelectedEmailContext();
      setStatus(`${docs.length} anexo(s) guardado(s) nos documentos do grupo principal.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel guardar os anexos no grupo.");
    } finally {
      setActionBusy(false);
    }
  }

  function toggleManagedGroupContact(contact: Partial<GroupContactDraft>) {
    const normalized = normalizeGroupContactDraft(contact);
    if (!normalized) return;
    setManagedGroupContacts((current) =>
      current.some((entry) => entry.key === normalized.key)
        ? current.filter((entry) => entry.key !== normalized.key)
        : dedupeGroupContacts([...current, normalized])
    );
  }

  function toggleManagedGroupEntity(entity: Partial<GroupEntityDraft>) {
    const normalized = normalizeGroupEntityDraft(entity);
    if (!normalized) return;
    setManagedGroupEntities((current) =>
      current.some((entry) => entry.key === normalized.key)
        ? current.filter((entry) => entry.key !== normalized.key)
        : dedupeGroupEntities([...current, normalized])
    );
  }

  async function handleSaveManagedGroupProfile() {
    const groupId = String(managedGroupId || "").trim();
    if (!groupId || !selectedManagedGroup) {
      setStatus("Escolhe primeiro um grupo para atualizar.");
      return;
    }
    setActionBusy(true);
    try {
      const updated = await updateLinkGroup(groupId, {
        name: selectedManagedGroup.name,
        description: managedGroupDescription,
        notes: managedGroupNotes,
        contacts: managedGroupContacts,
        entities: managedGroupEntities,
        documentsEnabled: selectedManagedGroup.documentsEnabled,
        status: selectedManagedGroup.status,
        labels: selectedManagedGroup.labels,
        isArchived: selectedManagedGroup.isArchived,
      });
      setAllGroups((current) => current.map((group) => (group.id === updated.id ? { ...group, ...updated } : group)));
      setCurrentCaseGroups((current) => current.map((group) => (group.id === updated.id ? { ...group, ...updated } : group)));
      setStatus(`Grupo ${updated.name} atualizado com descricao, notas e associacoes.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel atualizar o perfil do grupo.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleRemoveManagedGroupEmail(email: RelatedEmailEntry) {
    const groupId = String(managedGroupId || "").trim();
    if (!groupId) return;
    setActionBusy(true);
    try {
      await removeEmailFromLinkGroup(groupId, {
        ...email,
        emailKey: String(email?.emailKey || "").trim() || undefined,
      });
      setManagedGroupEmails((current) => current.filter((entry) => makeEmailKey(entry) !== makeEmailKey(email)));
      await refreshSelectedEmailContext();
      setStatus("Email removido do grupo.");
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel remover o email do grupo.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleDeleteManagedGroupDocument(document: GroupDocumentEntry) {
    const groupId = String(managedGroupId || "").trim();
    const documentId = String(document?.id || "").trim();
    if (!groupId || !documentId) return;
    setActionBusy(true);
    try {
      await deleteGroupDocument(groupId, documentId);
      setManagedGroupDocuments((current) => current.filter((entry) => String(entry.id || "").trim() !== documentId));
      setStatus("Documento removido do grupo.");
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel remover o documento.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleSearchTickets(queryOverride?: string, options?: { silent?: boolean }) {
    const query = String(queryOverride ?? ticketSearch ?? "").trim();
    const requestSeq = ++ticketSearchRequestSeqRef.current;
    if (!query) {
      setTicketSearchResults([]);
      setTicketSearchBusy(false);
      if (!options?.silent) {
        setStatus("Escreve primeiro parte do codigo ou do titulo para pesquisar tickets.");
      }
      return [];
    }
    setTicketSearchBusy(true);
    try {
      const rows = await searchGroupTickets({
        q: query || undefined,
        limit: 20,
      });
      if (requestSeq !== ticketSearchRequestSeqRef.current) return rows;
      setTicketSearchResults(rows);
      if (!options?.silent) {
        setStatus(rows.length ? `${rows.length} ticket(s) encontrados.` : "Nenhum ticket encontrado para estes filtros.");
      }
      return rows;
    } catch (actionError: any) {
      if (requestSeq === ticketSearchRequestSeqRef.current) {
        setTicketSearchResults([]);
      }
      setStatus(actionError?.message || "Nao foi possivel pesquisar tickets.");
      return [];
    } finally {
      if (requestSeq === ticketSearchRequestSeqRef.current) {
        setTicketSearchBusy(false);
      }
    }
  }

  function getClassificationSignature(targetEmailKeys: string[]) {
    return JSON.stringify({
      principalGroupId,
      referenceGroupIds,
      selectedTicketId,
      selectedSeriesId,
      selectedLabels: [...selectedLabels].sort(),
      labelDrafts: getComparableLabelDraftsSignature(labelDrafts),
      classificationMetaDraft: getComparableClassificationMetaSignature(classificationMetaDraft),
      ticketStatusDraft: String(ticketStatusDraft || "").trim(),
      applyScopeMode,
      targetEmailKeys: [...targetEmailKeys].sort(),
    });
  }

  async function handleApplyClassification(targetEmailsOverride?: RelatedEmailEntry[]): Promise<{ ok: boolean; coreSuccess: boolean; error?: string }> {
    if (applyInProgressRef.current) {
      logClassificationOutlookCategorySync("apply-aborted-busy", { operationId: "" });
      return { ok: false, coreSuccess: false, error: "Ja existe outra classificacao em curso." };
    }

    const targetEmails = dedupeEmails(
      ((targetEmailsOverride && targetEmailsOverride.length)
        ? targetEmailsOverride
        : resolvedApplySelection.targetEmails) as RelatedEmailEntry[]
    );
    const applySelection = (targetEmailsOverride && targetEmailsOverride.length)
      ? buildResolvedApplySelectionForTargets(targetEmails)
      : resolvedApplySelection;
    const targetEmailKeys = applySelection.targetEmailKeys;
    const currentSignature = getClassificationSignature(targetEmailKeys);

    if (lastAppliedSignatureRef.current === currentSignature) {
      logClassificationOutlookCategorySync("apply-noop-dedupe", { signature: currentSignature });
      return { ok: true, coreSuccess: true };
    }

    const applyPromise = (async () => {
      setActionBusy(true);
      setStatus("A iniciar aplicacao...");
      let activeCategoryOperationId = "";
      let activeCategoryRequestId = "";
      let categoryOperationClosed = false;
      let coreSuccess = false;
      let appliedClassificationCase: IntermediateCase | null = null;

      try {
        const effectiveTargetEmails = applySelection.targetEmails.length
          ? applySelection.targetEmails
          : ((selectedEmail ? [selectedEmail] : []) as RelatedEmailEntry[]);
        const preferredSelectedEmailKey = normalizeComparableString(selectedEmailKey || classificationAnchorEmailKey);

        if (!effectiveTargetEmails.length) {
          setStatus("Nao existe nenhum email alvo para atualizar.");
          return { ok: false, coreSuccess: false, error: "Nao existe nenhum email alvo." };
        }
        const remoteApplyPlan = buildResolvedRemoteApplyExecutionPlan({
          targetEmails: effectiveTargetEmails,
          resolvedApplySelection: applySelection,
          currentContext,
          emailContextMeta,
        });
        const currentTargetIdentity = remoteApplyPlan.currentTargetIdentity;
        const includesCurrentTarget = remoteApplyPlan.includesCurrentTarget;

        if (currentTargetIdentity) {
          const openedOperation = beginOutlookCategoryOperation({
            owner: "classification",
            target: currentTargetIdentity,
          });
          if (!openedOperation.ok) {
            const reasonMsg = openedOperation.reason === "locked"
              ? "Ja existe outra classificacao em curso para este email. Aguarda um momento."
              : "Nao foi possivel identificar o email atual para confirmar a classificacao.";
            setStatus(reasonMsg);
            return { ok: false, coreSuccess: false, error: reasonMsg };
          }
          activeCategoryOperationId = openedOperation.operation.operationId;
          setOutlookCategoryOperationPhase(activeCategoryOperationId, "saving");
        }

        const latestSettings = await getSettings().catch(() => null);
        const attachmentStorageOptions = buildAttachmentStorageOptions(latestSettings);
        const currentOutlookTicket = applySelection.selectedTicket;
        const desiredTicketStatus = applySelection.desiredTicketStatus;

        let finalTicket: GroupTicketEntry | null = null;
        if (remoteApplyPlan.shouldCreateTicket) {
          setStatus("A criar ticket Odoo...");
          const baseClassifiedEmailPayload = remoteApplyPlan.targetPlans[0]?.classifiedEmailPayload
            || buildResolvedClassifiedEmailPayload({
              targetEmail: remoteApplyPlan.baseTargetEmail,
              currentContext,
              resolvedApplySelection: applySelection,
            });
          finalTicket = await createGroupTicket({
            seriesId: applySelection.selectedSeriesId,
            title: String(createTicketTitle || remoteApplyPlan.baseTargetEmail?.subject || "Ticket").trim(),
            description: String(remoteApplyPlan.baseTargetEmail?.bodyText || "").trim().slice(0, 4000),
            labels: applySelection.labels,
            groupIds: applySelection.allGroupIds,
            email: baseClassifiedEmailPayload,
            membershipKind: applySelection.targetMembershipKind,
          });
          if (desiredTicketStatus && desiredTicketStatus !== String(finalTicket?.status || "").trim()) {
            finalTicket = await updateGroupTicket(finalTicket.id, { status: desiredTicketStatus });
          }
          setRelatedTickets((current) => [finalTicket as GroupTicketEntry, ...current.filter((entry) => entry.id !== finalTicket?.id)]);
          setSelectedTicketId(finalTicket.id);
        }

        if (remoteApplyPlan.shouldUpdateTicketStatus && desiredTicketStatus !== String(currentOutlookTicket?.status || "").trim()) {
          setStatus("A atualizar estado do ticket...");
          finalTicket = await updateGroupTicket(applySelection.selectedTicketId, { status: desiredTicketStatus });
          setRelatedTickets((current) => [finalTicket as GroupTicketEntry, ...current.filter((entry) => entry.id !== finalTicket?.id)]);
        }

        let emailCounter = 0;
        for (const targetPlan of remoteApplyPlan.targetPlans) {
          emailCounter++;
          setStatus(`A aplicar classificacao (${emailCounter}/${effectiveTargetEmails.length})...`);
          finalTicket = await executeLegacyRemoteApplyForTarget({
            targetPlan,
            resolvedApplySelection: applySelection,
            finalTicket,
            attachmentStorageOptions,
            skipTicketLink: Boolean(finalTicket && targetPlan.targetEmailKey === remoteApplyPlan.baseTargetKey),
          });
        }

        const resolvedCaseTicket = finalTicket || currentOutlookTicket;
        const localClassificationState = String(
          (classificationMetaDraft.ticketStatusEnabled ? desiredTicketStatus || resolvedCaseTicket?.status : "")
          || (classificationMetaDraft.principalStatusEnabled ? applySelection.principalGroup?.status : "")
          || ""
        ).trim();
        const localClassificationDraft: IntermediateCaseClassificationDraft = buildResolvedIntermediateCaseClassificationDraft({
          resolvedApplySelection: applySelection,
          resolvedCaseTicket,
          localClassificationState,
        });

        if (classificationCase) {
          setStatus("A gravar classificacao local no caso...");
          const classificationStorage = await resolveClassificationIntermediateCase({
            caseId: classificationCase.caseId,
            anchorEmailKey: classificationCase.anchorEmailKey,
          });
          const nextClassificationCase = applyClassificationToIntermediateCase({
            caseValue: classificationCase,
            targetEmails: effectiveTargetEmails,
            draft: localClassificationDraft,
          });
          await classificationStorage.storage.repository.writeCase(nextClassificationCase);
          appliedClassificationCase = nextClassificationCase;
          syncClassificationCaseEmails(nextClassificationCase, {
            preferredSelectedEmailKey,
            preferredTargetEmailKeys: targetEmailKeys,
          });
        }

        coreSuccess = true;
        setStatus("A atualizar dados locais...");

        let fallbackCurrentCategoryEmail: RelatedEmailEntry | null = null;
        if (currentTargetIdentity) {
          const currentTargetEmail = effectiveTargetEmails.find((email) => isCurrentContextEmail(email, currentContext))
            || (selectedEmailKey === selectedEmailKey && selectedEmail && isCurrentContextEmail(selectedEmail, currentContext) ? selectedEmail : null);

          fallbackCurrentCategoryEmail = buildRemoteApplyFallbackCurrentCategoryEmail({
            currentTargetEmail,
            currentContext,
            resolvedApplySelection: applySelection,
          });
        }

        setSelectionTouched({ principal: false, references: false, ticket: false });

        if (activeCategoryOperationId) {
          setOutlookCategoryOperationPhase(activeCategoryOperationId, "refreshing");
        }
        
        setStatus("A reidratar emails...");
        const refreshedContext = await refreshSelectedEmailContext().catch(() => null);
        if (appliedClassificationCase) {
          syncClassificationCaseEmails(appliedClassificationCase, {
            preferredSelectedEmailKey,
            preferredTargetEmailKeys: targetEmailKeys,
          });
        }

        if (includesCurrentTarget && currentTargetIdentity) {
          if (activeCategoryOperationId) {
            setOutlookCategoryOperationPhase(activeCategoryOperationId, "rehydrating");
          }
          const refreshedCategoryEmailCandidates = dedupeEmails([
            ...(refreshedContext?.email ? [refreshedContext.email] : []),
            ...(Array.isArray(refreshedContext?.emails) ? refreshedContext.emails : []),
            ...(fallbackCurrentCategoryEmail ? [fallbackCurrentCategoryEmail] : []),
          ]);
          const refreshedCategoryEmail = refreshedCategoryEmailCandidates.find((email) => isCurrentContextEmail(email, currentContext))
            || fallbackCurrentCategoryEmail;

          if (refreshedCategoryEmail) {
            if (activeCategoryOperationId) {
              setOutlookCategoryOperationPhase(activeCategoryOperationId, "planning");
            }
            const refreshedSnapshot = await getManagedOutlookCategorySnapshot(labelCatalog).catch(() => null);
            const refreshedCategorySource = buildOutlookCategorySourceFromRelatedContext({
              email: refreshedCategoryEmail,
              groups: Array.isArray(refreshedContext?.groups) ? refreshedContext.groups : [principalGroup, ...referenceGroups].filter(Boolean) as LinkGroupEntry[],
              tickets: Array.isArray(refreshedContext?.tickets) ? refreshedContext.tickets : [finalTicket, currentOutlookTicket].filter(Boolean) as GroupTicketEntry[],
              settings: latestSettings,
              currentOutlookLabelNames: refreshedSnapshot?.labelNames || [],
            });

            const categoryRequestId = `classification-final:${Date.now()}:${Math.random().toString(36).slice(2)}`;
            const categoryRequestedAtIso = new Date().toISOString();
            const categoryPlan = buildOutlookCategoryPlan(refreshedCategorySource);
            
            activeCategoryRequestId = categoryRequestId;

            logClassificationOutlookCategorySync("final-request", {
              requestId: categoryRequestId,
              operationId: activeCategoryOperationId || undefined,
              target: currentTargetIdentity,
              sourceSignature: getOutlookCategorySourceSignature(refreshedCategorySource),
              planSignature: getOutlookCategoryPlanSignature(categoryPlan),
              desiredCategories: categoryPlan.desiredCategories,
            });

            enqueueOutlookCategorySyncRequest({
              requestId: categoryRequestId,
              operationId: activeCategoryOperationId || undefined,
              createdAtIso: categoryRequestedAtIso,
              reason: "classification-final",
              mode: "source",
              target: currentTargetIdentity,
              source: refreshedCategorySource,
            });

            if (activeCategoryOperationId) {
              setOutlookCategoryOperationPhase(activeCategoryOperationId, "writingOutlook", {
                requestId: categoryRequestId,
              });
            }

            setStatus("A projetar categorias no Outlook...");
            const writerSubmitted = await requestCockpitHostAction({
              type: "sync-managed-categories",
              payload: refreshedCategorySource,
              requestId: categoryRequestId,
              operationId: activeCategoryOperationId || undefined,
              requestedAtIso: categoryRequestedAtIso,
              reason: "classification-final",
              target: currentTargetIdentity,
            }).catch(() => false);

            if (!writerSubmitted) {
              throw new Error("A classificacao foi guardada, mas nao foi possivel submeter a projecao Outlook.");
            }

            if (activeCategoryOperationId) {
              setOutlookCategoryOperationPhase(activeCategoryOperationId, "verifying", {
                requestId: categoryRequestId,
              });
            }

            const writerResult = await waitForOutlookCategorySyncResult(categoryRequestId, {
              timeoutMs: 20_000,
            });

            if (!writerResult) {
              clientLog.warn("[TEMP][outlook-category-apply] writer-timeout", {
                itemId: currentTargetIdentity.itemId,
                requestId: categoryRequestId,
                categoriesRequested: categoryPlan.desiredCategories,
                reason: "waitForOutlookCategorySyncResult timeout",
              });
              if (activeCategoryOperationId) {
                completeOutlookCategoryOperation(activeCategoryOperationId, {
                  result: "timeout",
                  requestId: categoryRequestId,
                  detail: "writer-timeout",
                });
                categoryOperationClosed = true;
              }
              throw new Error("A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias a tempo.");
            }

            if (activeCategoryOperationId) {
              completeOutlookCategoryOperation(activeCategoryOperationId, {
                result: writerResult.result,
                requestId: categoryRequestId,
                detail: writerResult.detail,
              });
              categoryOperationClosed = true;
            }

            if (writerResult.result !== "success" && writerResult.result !== "duplicate") {
              const writerDetail = String(writerResult.detail || "").trim();
              clientLog.warn("[TEMP][outlook-category-apply] writer-degraded-result", {
                itemId: currentTargetIdentity.itemId,
                requestId: categoryRequestId,
                categoriesRequested: categoryPlan.desiredCategories,
                writerResult: writerResult.result,
                detail: writerDetail || undefined,
              });
              throw new Error(
                writerDetail
                  ? `A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias (${writerDetail}).`
                  : "A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias."
              );
            }
          } else if (activeCategoryOperationId) {
            completeOutlookCategoryOperation(activeCategoryOperationId, {
              result: "failed",
              detail: "missing-refreshed-email",
            });
            categoryOperationClosed = true;
            throw new Error("A classificacao foi guardada, mas nao foi possivel rehidratar o email final para projetar as categorias.");
          }
        }

        setStatus(
          effectiveTargetEmails.length > 1
            ? `Classificacao aplicada a ${effectiveTargetEmails.length} emails.`
            : "Classificacao aplicada ao email selecionado."
        );
        lastAppliedSignatureRef.current = currentSignature;
        return { ok: true, coreSuccess: true };
      } catch (actionError: any) {
        if (activeCategoryOperationId && !categoryOperationClosed) {
          completeOutlookCategoryOperation(activeCategoryOperationId, {
            result: "failed",
            detail: String(actionError?.message || "").trim() || undefined,
          });
        }
        const errorMsg = actionError?.message || "Nao foi possivel aplicar a classificacao.";
        setStatus(errorMsg);
        if (coreSuccess) {
          return { ok: true, coreSuccess: true, error: `Guardado com avisos: ${errorMsg}` };
        }
        return { ok: false, coreSuccess: false, error: errorMsg };
      } finally {
        setActionBusy(false);
        applyInProgressRef.current = null;
      }
    })();

    applyInProgressRef.current = applyPromise;
    return applyPromise;
  }

  function handleOpenQuickAttachment(email: RelatedEmailEntry, attachment: NonNullable<ReturnType<typeof normalizeStudioAttachment>>) {
    const key = makeScopedAttachmentKey(email, attachment);
    if (!key) return;
    setSelectedAttachmentPreviewKey(key);
    setPreviewMode("document");
    // Optionally focus the email too? (No, stay on current selected email but show document)
  }

  async function handleSetQuickAttachmentHidden(
    email: RelatedEmailEntry,
    attachment: NonNullable<ReturnType<typeof normalizeStudioAttachment>>,
    nextHidden: boolean
  ) {
    if (!email || !attachment) return;
    const attachmentKey = makeAttachmentKey(attachment);
    if (!attachmentKey) return;
    const updatedEmail = updateAttachmentVisibilityOnEmail(email, attachmentKey, nextHidden);
    if (!updatedEmail) {
      setStatus("Nao foi possivel atualizar a visibilidade deste documento.");
      return;
    }
    setActionBusy(true);
    try {
      setRelatedEmails((current) => current.map((e) => makeEmailKey(e) === makeEmailKey(updatedEmail) ? updatedEmail : e));
      setKnownEmails((current) => current.map((e) => makeEmailKey(e) === makeEmailKey(updatedEmail) ? updatedEmail : e));
      mergeEmailIntoClassificationCase(updatedEmail);
      const latestSettings = await getSettings().catch(() => null);
      const payload = buildRelevantEmailPayloadFromRelatedEmail(updatedEmail);
      if (payload) {
        await registerRelevantEmail({
          ...payload,
          ...buildAttachmentStorageOptions(latestSettings),
        });
      }
      setStatus(
        nextHidden
          ? `Documento "${attachment.name}" ocultado dos documentos rapidos.`
          : `Documento "${attachment.name}" mantido visivel nos documentos rapidos.`
      );
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel atualizar a visibilidade do documento.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handlePreviewReply() {
    if (!selectedEmail) return;
    if (emailMatchesCurrentContext(selectedEmail, currentContext)) {
      const handled = await requestCockpitHostAction({ type: "reply-current" });
      setStatus(handled ? "Formulario de resposta aberto para o email atual." : "Nao foi possivel abrir a resposta.");
      return;
    }
    const opened = await requestCockpitHostAction({ type: "open-email", itemId: selectedEmail.itemId, emailWebLink: selectedEmail.emailWebLink });
    setStatus(opened ? "Email aberto no Outlook. Usa Responder no Outlook para continuar." : "Este email ainda nao tem abertura direta para responder.");
  }

  async function handlePreviewForward() {
    if (!selectedEmail) return;
    if (emailMatchesCurrentContext(selectedEmail, currentContext)) {
      const handled = await requestCockpitHostAction({ type: "forward-current" });
      setStatus(handled ? "Formulario de reencaminhamento aberto para o email atual." : "Nao foi possivel abrir o reencaminhamento.");
      return;
    }
    const opened = await requestCockpitHostAction({ type: "open-email", itemId: selectedEmail.itemId, emailWebLink: selectedEmail.emailWebLink });
    setStatus(opened ? "Email aberto no Outlook. Usa Reencaminhar no Outlook para continuar." : "Este email ainda nao tem abertura direta para reencaminhar.");
  }

  function captureClassificationDraftSnapshot() {
    return {
      principalGroupId,
      principalSearch,
      referenceGroupIds: [...referenceGroupIds],
      referenceSearch,
      selectedLabels: [...selectedLabels],
      labelDrafts: structuredClone(labelDrafts),
      classificationMetaDraft: structuredClone(classificationMetaDraft),
      selectedTicketId,
      selectedSeriesId,
      ticketStatusDraft,
      ticketSearch,
      ticketSearchResults: [...ticketSearchResults],
      createTicketTitle,
      selectionTouched: { ...selectionTouched },
    };
  }

  function restoreClassificationDraftSnapshot() {
    const snapshot = classificationDraftSnapshotRef.current;
    if (!snapshot) return;
    const normalizedSelection = createEmailGroupSelectionState({
      principalGroupId: snapshot.principalGroupId,
      referenceGroupIds: snapshot.referenceGroupIds,
    });
    setPrincipalGroupId(normalizedSelection.principalGroupId);
    setPrincipalSearch(snapshot.principalSearch);
    setReferenceGroupIds(normalizedSelection.referenceGroupIds);
    setReferenceSearch(snapshot.referenceSearch);
    setSelectedLabels([...snapshot.selectedLabels]);
    setLabelDrafts(structuredClone(snapshot.labelDrafts));
    setClassificationMetaDraft(structuredClone(snapshot.classificationMetaDraft));
    setSelectedTicketId(snapshot.selectedTicketId);
    setSelectedSeriesId(snapshot.selectedSeriesId);
    setTicketStatusDraft(snapshot.ticketStatusDraft);
    setTicketSearch(snapshot.ticketSearch);
    setTicketSearchResults([...snapshot.ticketSearchResults]);
    setCreateTicketTitle(snapshot.createTicketTitle);
    setSelectionTouched({ ...snapshot.selectionTouched });
  }

  function clearClassificationDraftSession() {
    classificationDraftSnapshotRef.current = null;
    setClassificationFocus("summary");
    setSection("emails");
    setApplyDialogOpen(false);
    setApplyDialogExpandedEmailKeys([]);
  }

  function openClassificationEditor(nextFocus: ClassificationFocus) {
    if (!classificationDraftSnapshotRef.current) {
      classificationDraftSnapshotRef.current = captureClassificationDraftSnapshot();
    }
    if (nextFocus === "ticket") {
      setTicketEditorMode(selectedSeriesId ? "new" : "existing");
    }
    setSection("classification");
    setClassificationFocus(nextFocus);
  }

  function handleCloseClassificationEditor() {
    restoreClassificationDraftSnapshot();
    clearClassificationDraftSession();
  }

  function getDefaultApplyDialogEmailKeys(mode: ApplyDialogScopeMode): string[] {
    if (mode === "case_all") {
      return caseScopeEmails.map((email) => makeEmailKey(email)).filter(Boolean);
    }
    if (mode === "selected") {
      const selectedKeys = selectedTargetEmailKeys.filter((key) => caseScopeEmails.some((email) => makeEmailKey(email) === key));
      return selectedKeys.length ? selectedKeys : [String(selectedEmailKey || "").trim()].filter(Boolean);
    }
    return [String(selectedEmailKey || "").trim()].filter(Boolean);
  }

  function setApplyDialogScope(mode: ApplyDialogScopeMode) {
    setApplyDialogScopeMode(mode);
    setApplyDialogEmailKeys(getDefaultApplyDialogEmailKeys(mode));
  }

  function openApplyDialog(sectionFocus: ClassificationFocus = classificationFocus) {
    const defaultMode: ApplyDialogScopeMode = selectedTargetEmailKeys.length > 1 ? "selected" : "current";
    setApplyDialogSection(sectionFocus === "summary" ? "summary" : sectionFocus);
    setApplyDialogExpandedEmailKeys([]);
    setApplyDialogOpen(true);
    setApplyDialogScopeMode(defaultMode);
    const keys = getDefaultApplyDialogEmailKeys(defaultMode);
    setApplyDialogEmailKeys(keys);
    setApplyDialogSelectedEmailKeys(keys);
  }

  function toggleApplyDialogEmailKey(emailKey: string) {
    if (!emailKey) return;
    const toggle = (current: string[]) => current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey];
    setApplyDialogEmailKeys(toggle);
    setApplyDialogSelectedEmailKeys(toggle);
  }

  function toggleApplyDialogExpandedEmailKey(emailKey: string) {
    if (!emailKey) return;
    setApplyDialogExpandedEmailKeys((current) => current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]);
  }

  function toggleExpandedEmailKey(emailKey: string) {
    if (!emailKey) return;
    setExpandedEmailKeys((current) => current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]);
  }

  async function handleConfirmApplyDialog() {
    const result = await handleApplyClassification();
    if (result && result.coreSuccess) {
      setApplyDialogOpen(false);
    }
  }

  // legacy renderers removed - handled by modular components

  function renderWorkspace() {
    if (loading) return <PanelState compact tone="loading" title="A preparar a janela" description="A carregar emails, grupos e series para o novo studio." />;
    if (error) return <PanelState compact tone="error" title="Falha a preparar o studio" description={error} />;

    if (section === "emails") {
      if (!selectedEmail) return <PanelState compact tone="info" title="Sem email selecionado" description="Escolhe um email na coluna do meio." />;
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.titleRow}>
              <div>
                <div style={S.cardTitle}>Preview grande</div>
                <div style={S.cardMeta}>{selectedEmail.subject || "(sem assunto)"}</div>
              </div>
              {(selectedEmail.itemId || selectedEmail.emailWebLink) ? (
                <button type="button" style={S.secondaryBtn} onClick={() => void requestCockpitHostAction({ type: "open-email", itemId: selectedEmail.itemId, emailWebLink: selectedEmail.emailWebLink })}>
                  <Icons.ExternalLink size={12} />
                  Abrir no Outlook
                </button>
              ) : null}
            </div>
            <div style={S.metaLine}>
              <span>{selectedEmail.fromName || selectedEmail.fromEmail || "--"}</span>
              <span>{formatDate(selectedEmail.messageDateIso || selectedEmail.receivedAtIso) || "--"}</span>
              <span>{Array.isArray(selectedEmail.attachments) ? `${selectedEmail.attachments.length} anexo(s)` : "Sem anexos"}</span>
            </div>
            {previewHtml ? <div style={S.previewHtml} dangerouslySetInnerHTML={{ __html: previewHtml }} /> : <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado suficiente para preview." />}
          </div>

          <div style={S.card}>
            <div style={S.titleRow}>
              <div>
                <div style={S.cardTitle}>Documentos e imagens</div>
                <div style={S.cardMeta}>Preview simples dos anexos deste email.</div>
              </div>
            </div>
            {canonicalSelectedEmailAttachmentEntries.length ? (
              <div style={S.stackMini}>
                <div style={S.attachmentPickerBar}>
                  {canonicalSelectedEmailAttachmentEntries.map((entry) => {
                    const key = entry.scopedKey;
                    const active = key === selectedAttachmentPreviewKey;
                    return (
                      <button
                        key={key}
                        type="button"
                        style={active ? S.groupChipBtnOn : S.groupChipBtn}
                        onClick={() => setSelectedAttachmentPreviewKey(key)}
                      >
                        {entry.attachment.name}
                      </button>
                    );
                  })}
                </div>
                <div style={S.card}>
                  {selectedAttachmentPreview ? (
                    <>
                      <div style={S.summaryGrid}>
                        <div style={S.summaryRow}><span>Ficheiro</span><strong>{selectedAttachmentPreview.name || "--"}</strong></div>
                        <div style={S.summaryRow}><span>Tipo</span><strong>{selectedAttachmentPreview.contentType || "ficheiro"}</strong></div>
                        <div style={S.summaryRow}><span>Tamanho</span><strong>{selectedAttachmentPreview.size ? `${Math.round(Number(selectedAttachmentPreview.size || 0) / 1024)} KB` : "--"}</strong></div>
                        <div style={S.summaryRow}><span>Estado documental</span><strong>{formatDocumentLifecycleState(selectedAttachmentDocumentState)}</strong></div>
                      </div>
                      <label style={S.field}>
                        <span style={S.label}>Atualizar estado deste anexo</span>
                        <select
                          style={S.select}
                          value={selectedAttachmentDocumentState}
                          onChange={(event) => void handleSetSelectedAttachmentDocumentState(event.target.value as DocumentLifecycleState)}
                          disabled={actionBusy}
                        >
                          {DOCUMENT_STATE_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                        </select>
                      </label>
                      <div style={S.cardMeta}>Se marcares como rejeitado, este anexo deixa de entrar automaticamente em leituras futuras.</div>
                    </>
                  ) : null}
                  {selectedAttachmentPreviewMode === "image" ? (
                    selectedAttachmentPreviewSrc ? (
                      <div style={S.attachmentPreviewWrap}>
                        <img src={selectedAttachmentPreviewSrc} alt={selectedAttachmentPreview?.name || "Imagem"} style={S.attachmentPreviewImage} />
                      </div>
                    ) : (
                      <div style={S.attachmentPreviewEmpty}>
                        {selectedAttachmentPreviewRemoteStatus === "loading"
                          ? "A carregar imagem..."
                          : selectedAttachmentPreview?.hasContent
                            ? "Nao foi possivel carregar o conteudo persistido desta imagem."
                            : "Esta imagem ainda nao foi persistida com conteudo."}
                      </div>
                    )
                  ) : null}
                  {selectedAttachmentPreviewMode === "pdf" ? (
                    selectedAttachmentPreviewSrc ? (
                      <StudioPdfPreview dataUrl={selectedAttachmentPreviewSrc} title={selectedAttachmentPreview?.name || "PDF"} />
                    ) : (
                      <div style={S.attachmentPreviewEmpty}>
                        {selectedAttachmentPreviewRemoteStatus === "loading"
                          ? "A carregar PDF..."
                          : selectedAttachmentPreview?.hasContent
                            ? "Nao foi possivel carregar o conteudo persistido deste PDF."
                            : "Este PDF ainda nao foi persistido com conteudo."}
                      </div>
                    )
                  ) : null}
                  {selectedAttachmentPreviewMode === "text" ? (
                    selectedAttachmentPreviewText ? (
                      <pre style={S.attachmentPreviewText}>{selectedAttachmentPreviewText}</pre>
                    ) : (
                      <div style={S.attachmentPreviewEmpty}>Nao foi possivel ler o conteudo textual deste ficheiro.</div>
                    )
                  ) : null}
                  {selectedAttachmentPreviewMode === "unsupported" ? (
                    <div style={S.attachmentPreviewEmpty}>Preview nao disponivel para este tipo de ficheiro.</div>
                  ) : null}
                  {selectedAttachmentPreviewMode === "none" ? (
                    <div style={S.attachmentPreviewEmpty}>Escolhe um anexo para ver o preview.</div>
                  ) : null}
                </div>
              </div>
            ) : (
              <PanelState compact tone="info" title="Sem anexos disponiveis" description="Este email nao traz anexos guardados para preview." />
            )}
          </div>
        </div>
      );
    }

    if (section === "classification") {
      return (
        <div style={S.stack}>
          <div style={S.cardSticky}>
            <div style={S.classificationHeader}>
              <div>
                <div style={S.cardTitle}>Classificacao</div>
                <div style={S.cardMeta}>Clicar nos chips liga ou desliga a classificacao do email.</div>
              </div>
            </div>
            <div style={S.suggestionDock}>
              <div style={S.suggestionDockMeta}>Sugestoes da leitura. Clica para ligar ou desligar.</div>
              <div style={S.suggestionDockChips}>
                {classificationSuggestions.length ? (
                  classificationSuggestions.map((suggestion) => (
                    <button
                      key={suggestion.key}
                      type="button"
                      style={isSuggestionActive(suggestion) ? S.suggestionDockChipOn : S.suggestionDockChip}
                      onClick={() => handleSuggestionToggle(suggestion)}
                    >
                      {suggestion.label}
                    </button>
                  ))
                ) : (
                  <span style={S.mutedMini}>Sem sugestoes fortes para este email.</span>
                )}
              </div>
            </div>
            <div style={S.classificationFocusBar}>
              <button type="button" style={classificationFocus === "principal" ? S.classificationFocusBtnOn : S.classificationFocusBtn} onClick={() => setClassificationFocus("principal")}>Grupo principal</button>
              <button type="button" style={classificationFocus === "references" ? S.classificationFocusBtnOn : S.classificationFocusBtn} onClick={() => setClassificationFocus("references")}>Referencias</button>
              <button type="button" style={classificationFocus === "labels" ? S.classificationFocusBtnOn : S.classificationFocusBtn} onClick={() => setClassificationFocus("labels")}>Etiquetas</button>
              <button type="button" style={classificationFocus === "ticket" ? S.classificationFocusBtnOn : S.classificationFocusBtn} onClick={() => setClassificationFocus("ticket")}>Ticket</button>
              <button type="button" style={classificationFocus === "summary" ? S.classificationFocusBtnOn : S.classificationFocusBtn} onClick={() => setClassificationFocus("summary")}>Resumo</button>
            </div>
          </div>

          {classificationFocus === "principal" ? (
          <div style={S.classificationSectionCard}>
            <button type="button" style={classificationFocus === "principal" ? S.sectionHeadOn : S.sectionHead} onClick={() => setClassificationFocus("principal")}>
              <span style={S.sectionName}>Grupo principal</span>
              <span style={S.sectionMeta}>Casa principal do email</span>
            </button>
            <div style={S.sectionBody}>
              <div style={S.inlineWrap}>
                {principalGroup ? (
                  <button type="button" style={S.selectedChipOn} onClick={clearPrincipalSelection}>
                    {principalGroup.name}
                  </button>
                ) : (
                  <span style={S.mutedMini}>Sem grupo principal</span>
                )}
              </div>
              <div style={S.stackMini}>
                <div style={S.fieldLineLabel}>Favoritos</div>
                <div style={S.compactRowWrap}>
                  {favoritePrincipalGroups.length ? (
                    favoritePrincipalGroups.slice(0, 6).map((group) => (
                      <button
                        key={group.id}
                        type="button"
                        style={group.id === effectivePrincipalGroupId ? S.miniChipOn : S.miniChip}
                        onClick={() => toggleFavoritePrincipalGroup(group)}
                      >
                        {group.name}
                      </button>
                    ))
                  ) : (
                    <span style={S.mutedMini}>Sem grupos favoritos.</span>
                  )}
                </div>
              </div>
              <div style={S.stackMini}>
                <div style={S.fieldLineLabel}>Pesquisar ou criar</div>
                <div style={S.searchActionRow}>
                  <input
                    style={S.input}
                    value={principalSearch}
                    onChange={(event) => setPrincipalSearchValue(event.target.value)}
                    placeholder="Escreve o nome do grupo..."
                  />
                  <button
                    type="button"
                    style={String(principalSearch || "").trim() ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={() => {
                      if (principalCanCreate) {
                        void handleCreateGroupAndLink("principal", principalSearch);
                        return;
                      }
                      if (exactPrincipalSearchGroup) {
                        selectPrincipalGroup(exactPrincipalSearchGroup);
                      }
                    }}
                    disabled={!String(principalSearch || "").trim()}
                    title={principalCanCreate ? "Criar grupo" : exactPrincipalSearchGroup ? "Selecionar grupo existente" : "Pesquisar grupo"}
                  >
                    {principalCanCreate ? <Icons.Plus size={14} /> : <Icons.Search size={14} />}
                  </button>
                  <button
                    type="button"
                    style={principalSettingsTargetGroup ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={() => openManagedGroupFromPrincipal(principalSettingsTargetGroup)}
                    disabled={!principalSettingsTargetGroup}
                    title={principalSettingsTargetGroup ? "Abrir configuracao do grupo" : "Seleciona ou cria um grupo para abrir a configuracao"}
                  >
                    <Icons.Settings size={14} />
                  </button>
                </div>
                {principalSearchResults.length ? (
                  <div style={S.searchResultList}>
                    {principalSearchResults.map((group) => (
                      <button
                        key={group.id}
                        type="button"
                        style={group.id === effectivePrincipalGroupId ? S.searchResultBtnOn : S.searchResultBtn}
                        onClick={() => {
                          if (group.id === effectivePrincipalGroupId) {
                            clearPrincipalSelection();
                            return;
                          }
                          selectPrincipalGroup(group);
                        }}
                      >
                        <span>{group.name}</span>
                        {group.id === effectivePrincipalGroupId ? <span style={S.resultMiniMeta}>Ligado</span> : null}
                      </button>
                    ))}
                  </div>
                ) : String(principalSearch || "").trim() ? (
                  <div style={S.cardMeta}>
                    {principalCanCreate
                      ? `Ainda nao existe nenhum grupo com este nome. Usa o + para criar "${String(principalSearch || "").trim()}".`
                      : "Grupo exato encontrado. Usa a lupa para o ligar."}
                  </div>
                ) : null}
              </div>
              <div style={S.inlineChecks}>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.principalCategorize}
                    onChange={(event) => updateClassificationMeta({ principalCategorize: event.target.checked })}
                    disabled={!principalGroup}
                  />
                  <span>Grupo em categoria Outlook</span>
                </label>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.principalStatusEnabled}
                    onChange={(event) => updateClassificationMeta({ principalStatusEnabled: event.target.checked })}
                    disabled={!principalGroup?.status}
                  />
                  <span>Estado do grupo</span>
                </label>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.principalStatusCategorize}
                    onChange={(event) => updateClassificationMeta({ principalStatusCategorize: event.target.checked, principalStatusEnabled: event.target.checked ? true : classificationMetaDraft.principalStatusEnabled })}
                    disabled={!principalGroup?.status || !classificationMetaDraft.principalStatusEnabled}
                  />
                  <span>Estado em categoria Outlook</span>
                </label>
              </div>
              <div style={S.cardMeta}>
                {principalGroup?.status ? `Estado atual: ${principalGroupStatusLabel}` : "Sem estado definido neste grupo."}
              </div>
            </div>
          </div>
          ) : null}

          {classificationFocus === "references" ? (
          <div style={S.classificationSectionCard}>
            <button type="button" style={classificationFocus === "references" ? S.sectionHeadOn : S.sectionHead} onClick={() => setClassificationFocus("references")}>
              <span style={S.sectionName}>Referencias</span>
              <span style={S.sectionMeta}>Outros grupos ligados a este email</span>
            </button>
            <div style={S.sectionBody}>
              <div style={S.inlineWrap}>
                {referenceGroups.length ? referenceGroups.map((group) => (
                  <button key={group.id} type="button" style={S.selectedChipOn} onClick={() => toggleReferenceGroup(group.id)}>
                    {group.name}
                  </button>
                )) : <span style={S.mutedMini}>Sem referencias</span>}
              </div>
              <div style={S.stackMini}>
                <div style={S.fieldLineLabel}>Favoritos</div>
                <div style={S.compactRowWrap}>
                  {favoriteReferenceGroups.length ? (
                    favoriteReferenceGroups.map((group) => (
                      <button
                        key={group.id}
                        type="button"
                        style={effectiveReferenceGroupIds.includes(group.id) ? S.miniChipOn : S.miniChip}
                        onClick={() => toggleFavoriteReferenceGroup(group)}
                      >
                        {group.name}
                      </button>
                    ))
                  ) : (
                    <span style={S.mutedMini}>Sem grupos favoritos.</span>
                  )}
                </div>
              </div>
              <div style={S.stackMini}>
                <div style={S.fieldLineLabel}>Pesquisar ou criar</div>
                <div style={S.searchActionRow}>
                  <input
                    style={S.input}
                    value={referenceSearch}
                    onChange={(event) => setReferenceSearchValue(event.target.value)}
                    placeholder="Escreve o nome da referencia..."
                  />
                  <button
                    type="button"
                    style={String(referenceSearch || "").trim() ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={() => {
                      if (referenceCanCreate) {
                        void handleCreateGroupAndLink("referencia", referenceSearch);
                        return;
                      }
                      if (exactReferenceSearchGroup) {
                        toggleReferenceGroup(exactReferenceSearchGroup.id);
                        setReferenceSearchValue("");
                      }
                    }}
                    disabled={!String(referenceSearch || "").trim()}
                    title={referenceCanCreate ? "Criar referencia" : exactReferenceSearchGroup ? "Ligar ou desligar referencia existente" : "Pesquisar referencia"}
                  >
                    {referenceCanCreate ? <Icons.Plus size={14} /> : <Icons.Search size={14} />}
                  </button>
                  <button
                    type="button"
                    style={referenceSettingsTargetGroup ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={() => openManagedGroupFromPrincipal(referenceSettingsTargetGroup)}
                    disabled={!referenceSettingsTargetGroup}
                    title={referenceSettingsTargetGroup ? "Abrir configuracao da referencia" : "Seleciona ou encontra uma referencia para abrir a configuracao"}
                  >
                    <Icons.Settings size={14} />
                  </button>
                </div>
                {referenceSearchResults.length ? (
                  <div style={S.searchResultList}>
                    {referenceSearchResults.map((group) => (
                      <button
                        key={group.id}
                        type="button"
                        style={effectiveReferenceGroupIds.includes(group.id) ? S.searchResultBtnOn : S.searchResultBtn}
                        onClick={() => {
                          toggleReferenceGroup(group.id);
                          setReferenceSearchValue("");
                        }}
                      >
                        <span>{group.name}</span>
                        {effectiveReferenceGroupIds.includes(group.id) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
                      </button>
                    ))}
                  </div>
                ) : String(referenceSearch || "").trim() ? (
                  <div style={S.cardMeta}>
                    {referenceCanCreate
                      ? `Ainda nao existe nenhum grupo com este nome. Usa o + para criar "${String(referenceSearch || "").trim()}".`
                      : "Referencia exata encontrada. Usa a lupa para a ligar ou desligar."}
                  </div>
                ) : null}
              </div>
              <div style={S.inlineChecks}>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.referenceCategorize}
                    onChange={(event) => updateClassificationMeta({ referenceCategorize: event.target.checked })}
                    disabled={!referenceGroups.length}
                  />
                  <span>Referencias em categoria Outlook</span>
                </label>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.referenceStatusEnabled}
                    onChange={(event) => updateClassificationMeta({ referenceStatusEnabled: event.target.checked })}
                    disabled={!referenceGroupStatusEntries.length}
                  />
                  <span>Estado das referencias</span>
                </label>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.referenceStatusCategorize}
                    onChange={(event) => updateClassificationMeta({ referenceStatusCategorize: event.target.checked, referenceStatusEnabled: event.target.checked ? true : classificationMetaDraft.referenceStatusEnabled })}
                    disabled={!referenceGroupStatusEntries.length || !classificationMetaDraft.referenceStatusEnabled}
                  />
                  <span>Estado em categoria Outlook</span>
                </label>
              </div>
              <div style={S.inlineWrap}>
                {referenceGroupStatusEntries.length ? referenceGroupStatusEntries.map((entry) => (
                  <span key={`${entry.id}-status`} style={S.groupChip}>
                    {entry.name}: {entry.status}
                  </span>
                )) : <span style={S.mutedMini}>Sem estado nas referencias atuais.</span>}
              </div>
            </div>
          </div>
          ) : null}

          {classificationFocus === "labels" ? (
          <div style={S.classificationSectionCard}>
            <button type="button" style={classificationFocus === "labels" ? S.sectionHeadOn : S.sectionHead} onClick={() => setClassificationFocus("labels")}>
              <span style={S.sectionName}>Etiquetas</span>
              <span style={S.sectionMeta}>Etiquetas do email, com categoria e estado opcionais</span>
            </button>
            <div style={S.sectionBodyScroll}>
              <div style={S.inlineWrap}>
                {summaryLabels.length ? summaryLabels.map((label) => (
                  <button key={label} type="button" style={S.selectedChipOn} onClick={() => removeLabel(label)}>
                    {label}
                  </button>
                )) : <span style={S.mutedMini}>Sem etiquetas</span>}
              </div>
              <div style={S.stackMini}>
                <div style={S.fieldLineLabel}>Pesquisar ou criar</div>
                <div style={S.compactSearchActionRow}>
                  <input
                    style={S.input}
                    value={classificationLabelInput}
                    onChange={(event) => setClassificationLabelInput(event.target.value)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter") {
                        event.preventDefault();
                        handleClassificationLabelSearchAction();
                      }
                    }}
                    placeholder="Escreve o nome da etiqueta..."
                  />
                  <button
                    type="button"
                    style={String(classificationLabelInput || "").trim() ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={handleClassificationLabelSearchAction}
                    disabled={!String(classificationLabelInput || "").trim()}
                    title={classificationLabelCanCreate ? "Criar etiqueta" : exactClassificationLabel ? "Ligar ou desligar etiqueta existente" : "Pesquisar etiqueta"}
                  >
                    {classificationLabelCanCreate ? <Icons.Plus size={14} /> : <Icons.Search size={14} />}
                  </button>
                </div>
                {filteredClassificationLabels.length && String(classificationLabelInput || "").trim() ? (
                  <div style={S.searchResultList}>
                    {filteredClassificationLabels.map((label) => (
                      <button
                        key={label}
                        type="button"
                        style={selectedLabels.includes(label) ? S.searchResultBtnOn : S.searchResultBtn}
                        onClick={() => {
                          if (selectedLabels.includes(label)) {
                            removeLabel(label);
                          } else {
                            addLabel(label);
                          }
                          setClassificationLabelInput(label);
                        }}
                      >
                        <span>{label}</span>
                        {selectedLabels.includes(label) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
                      </button>
                    ))}
                  </div>
                ) : String(classificationLabelInput || "").trim() ? (
                  <div style={S.cardMeta}>
                    {classificationLabelCanCreate
                      ? `Ainda nao existe nenhuma etiqueta com este nome. Usa o + para criar "${String(classificationLabelInput || "").trim()}".`
                      : "Etiqueta exata encontrada. Usa a lupa para a ligar ou desligar."}
                  </div>
                ) : null}
              </div>
              {selectedLabels.length ? (
                <div style={S.labelGrid}>
                  {selectedLabels.map((label) => {
                    const draft = labelDrafts[label] || { categorize: false, hasStatus: false };
                    return (
                      <div key={label} style={S.labelRowCompact}>
                        <div style={S.labelHead}>
                          <strong>{label}</strong>
                          <button type="button" style={S.linkBtn} onClick={() => removeLabel(label)}>Off</button>
                        </div>
                        <div style={S.inlineChecks}>
                          <label style={S.check}><input type="checkbox" checked={draft.categorize} onChange={(event) => updateLabelDraft(label, { categorize: event.target.checked })} /><span>Categoria</span></label>
                          <label style={S.check}><input type="checkbox" checked={draft.hasStatus} onChange={(event) => updateLabelDraft(label, { hasStatus: event.target.checked, status: event.target.checked ? (draft.status || "em_analise") : undefined })} /><span>Estado</span></label>
                        </div>
                        {draft.hasStatus ? (
                          <select style={S.select} value={draft.status || "em_analise"} onChange={(event) => updateLabelDraft(label, { status: event.target.value as EmailLabelStatus, hasStatus: true })}>
                            {LABEL_STATUS_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                          </select>
                        ) : null}
                      </div>
                    );
                  })}
                </div>
              ) : null}
            </div>
          </div>
          ) : null}

          {classificationFocus === "ticket" ? (
          <div style={S.classificationSectionCard}>
            <button type="button" style={classificationFocus === "ticket" ? S.sectionHeadOn : S.sectionHead} onClick={() => setClassificationFocus("ticket")}>
              <span style={S.sectionName}>Ticket</span>
              <span style={S.sectionMeta}>Escolher ticket existente ou criar novo</span>
            </button>
            <div style={S.sectionBody}>
              <div style={S.inlineWrap}>
                {selectedTicket ? (
                  <button type="button" style={S.selectedChipOn} onClick={clearTicketSelection}>
                    {selectedTicket.code}
                  </button>
                ) : selectedSeriesId ? (
                  <button type="button" style={S.selectedChipPending} onClick={clearTicketSelection}>
                    {ticketSummary}
                  </button>
                ) : (
                  <span style={S.mutedMini}>Sem ticket</span>
                )}
              </div>
              <div style={S.sectionControls}>
                <input style={S.input} value={ticketSearch} onChange={(event) => setTicketSearch(event.target.value)} placeholder="Pesquisar por codigo, titulo ou etiqueta..." />
                <button type="button" style={S.secondaryBtn} onClick={() => void handleSearchTickets()} disabled={actionBusy}>
                  <Icons.Search size={12} />
                  Pesquisar
                </button>
              </div>
              <div style={S.chips}>
                {(ticketSearchResults.length ? ticketSearchResults : ticketPickerChoices.slice(0, 12)).map((ticket) => (
                  <button
                    key={ticket.id}
                    type="button"
                    style={ticket.id === selectedTicketId ? S.groupChipBtnOn : S.groupChipBtn}
                    onClick={() => {
                      setSelectionTouched((current) => ({ ...current, ticket: true }));
                      setSelectedTicketId(ticket.id === selectedTicketId ? "" : ticket.id);
                      if (ticket.id !== selectedTicketId) setSelectedSeriesId("");
                    }}
                  >
                    {ticket.code}
                  </button>
                ))}
              </div>
              <div style={S.grid2}>
                <select style={S.select} value={selectedSeriesId} onChange={(event) => { const nextValue = event.target.value; setSelectionTouched((current) => ({ ...current, ticket: true })); setSelectedSeriesId(nextValue); if (nextValue) setSelectedTicketId(""); }}>
                  <option value="">Sem novo ticket</option>
                  {ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} Â· {series.name}</option>)}
                </select>
                <input style={S.input} value={createTicketTitle} onChange={(event) => setCreateTicketTitle(event.target.value)} placeholder="Titulo do ticket" />
              </div>
              <div style={S.inline}>
                <button type="button" style={S.secondaryBtn} onClick={() => void handleCreateTicketAndLink()} disabled={actionBusy || !selectedSeriesId}>
                  <Icons.Plus size={12} />
                  Criar ticket
                </button>
              </div>
              <div style={S.grid2}>
                <select
                  style={S.select}
                  value={ticketStatusDraft}
                  onChange={(event) => {
                    setSelectionTouched((current) => ({ ...current, ticket: true }));
                    setTicketStatusDraft(event.target.value);
                  }}
                  disabled={!selectedTicketId && !selectedSeriesId}
                >
                  {TICKET_STATUS_OPTIONS.map((option) => (
                    <option key={option.value || "empty"} value={option.value}>{option.label}</option>
                  ))}
                </select>
                <div style={S.cardMeta}>
                  {effectiveTicketStatus ? `Estado preparado: ${ticketStatusLabel}` : "Sem estado definido neste ticket."}
                </div>
              </div>
              <div style={S.inlineChecks}>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.ticketStatusEnabled}
                    onChange={(event) => updateClassificationMeta({ ticketStatusEnabled: event.target.checked })}
                    disabled={!effectiveTicketStatus}
                  />
                  <span>Estado do ticket</span>
                </label>
                <label style={S.check}>
                  <input
                    type="checkbox"
                    checked={classificationMetaDraft.ticketStatusCategorize}
                    onChange={(event) => updateClassificationMeta({ ticketStatusCategorize: event.target.checked, ticketStatusEnabled: event.target.checked ? true : classificationMetaDraft.ticketStatusEnabled })}
                    disabled={!effectiveTicketStatus || !classificationMetaDraft.ticketStatusEnabled}
                  />
                  <span>Estado em categoria Outlook</span>
                </label>
              </div>
              <div style={S.cardMeta}>
                {effectiveTicketStatus ? `Estado atual: ${ticketStatusLabel}` : "Sem estado definido neste ticket."}
              </div>
            </div>
          </div>
          ) : null}

          {classificationFocus === "summary" ? (
          <div style={S.classificationSectionCard}>
            <div style={S.sectionHeadStatic}>
              <span style={S.sectionName}>Resumo e gravacao</span>
              <span style={S.sectionMeta}>Revisao final do que vai ser aplicado</span>
            </div>
            <div style={S.sectionBodyScroll}>
              <div style={S.subTitle}>Ambito de aplicacao</div>
              <select style={S.select} value={applyScopeMode} onChange={(event) => setApplyScopeMode(event.target.value as ApplyScopeMode)}>
                <option value="current">So email atual</option>
                <option value="selected">Emails selecionados ({selectedTargetCount})</option>
                <option value="principal_group">Mesmo grupo principal ({principalScopeCount})</option>
              </select>
              <div style={S.summaryGrid}>
                <div style={S.summaryRow}><span>Email atual</span><strong>{selectedEmail?.subject || "--"}</strong></div>
                <div style={S.summaryRow}><span>Selecionados manualmente</span><strong>{selectedTargetCount}</strong></div>
                <div style={S.summaryRow}><span>No mesmo grupo principal</span><strong>{principalScopeCount}</strong></div>
              </div>
              <div style={S.cardMeta}>
                Em modo multiplo, aplicamos a classificacao atual exatamente aos emails escolhidos.
              </div>
              <div style={S.subTitle}>Atualizar email</div>
              <div style={S.summaryGrid}>
                <div style={S.summaryRow}><span>Grupo principal</span><strong>{principalGroup?.name || principalGroupId || "--"}</strong></div>
                <div style={S.summaryRow}><span>Estado grupo</span><strong>{classificationMetaDraft.principalStatusEnabled ? principalGroupStatusLabel || "--" : "--"}</strong></div>
                <div style={S.summaryRow}><span>Referencias</span><strong>{referenceGroupSummary}</strong></div>
                <div style={S.summaryRow}><span>Estado referencias</span><strong>{classificationMetaDraft.referenceStatusEnabled ? (referenceGroupStatusEntries.length ? referenceGroupStatusEntries.map((entry) => entry.status).join(", ") : "--") : "--"}</strong></div>
                <div style={S.summaryRow}><span>Ticket</span><strong>{ticketSummary}</strong></div>
                <div style={S.summaryRow}><span>Estado ticket</span><strong>{classificationMetaDraft.ticketStatusEnabled ? ticketStatusLabel || "--" : "--"}</strong></div>
                <div style={S.summaryRow}><span>Etiquetas</span><strong>{summaryLabels.length ? summaryLabels.join(", ") : "--"}</strong></div>
                <div style={S.summaryRow}><span>Estado por etiquetas</span><strong>{emailStatusSummary}</strong></div>
              </div>
              <div style={S.summaryActionBar}>
                <button type="button" style={S.primaryBtn} onClick={() => void handleApplyClassification()} disabled={actionBusy || (!effectivePrincipalGroupId && !effectiveReferenceGroupIds.length && !selectedTicketId && !selectedSeriesId && !selectedEmailGroups.length && !selectedEmailTicketIds.length && !selectedLabels.length && !selectedEmailStoredLabels.length && !String(selectedEmail?.status || "").trim())}>
                  <Icons.Save size={12} />
                  Gravar / atualizar
                </button>
                <span style={S.cardMeta}>Mantemos a logica atual de gravacao enquanto fechamos a nova estrutura.</span>
              </div>
            </div>
          </div>
          ) : null}
        </div>
      );
    }

    if (section === "labels") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas estruturadas</div>
            <div style={S.cardMeta}>Aqui podes manter o email so com etiquetas, com ou sem estado, mesmo sem grupo principal nem ticket.</div>
            <div style={S.inline}>
              <input style={S.input} value={labelInput} onChange={(event) => setLabelInput(event.target.value)} placeholder="Pesquisar ou criar etiqueta" />
              <button type="button" style={S.secondaryBtn} onClick={() => addLabel(labelInput)} disabled={!String(labelInput || "").trim()}><Icons.Plus size={12} />Adicionar</button>
            </div>
            {filteredLabelCatalog.length ? <div style={S.chips}>{filteredLabelCatalog.slice(0, 24).map((label) => <button key={label} type="button" style={selectedLabels.includes(label) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => addLabel(label)}>{label}</button>)}</div> : null}
            {outlookLabelCategories.length ? <div style={S.cardMeta}>Ja categorizadas no Outlook: {outlookLabelCategories.join(", ")}</div> : null}
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas selecionadas</div>
            {selectedLabels.length ? selectedLabels.map((label) => {
              const draft = labelDrafts[label] || { categorize: false, hasStatus: false };
              return (
                <div key={label} style={S.labelRow}>
                  <div style={S.labelHead}><strong>{label}</strong><button type="button" style={S.linkBtn} onClick={() => removeLabel(label)}>Remover</button></div>
                  <label style={S.check}><input type="checkbox" checked={draft.categorize} onChange={(event) => updateLabelDraft(label, { categorize: event.target.checked })} /><span>Virar categoria Outlook</span></label>
                  <label style={S.check}><input type="checkbox" checked={draft.hasStatus} onChange={(event) => updateLabelDraft(label, { hasStatus: event.target.checked, status: event.target.checked ? (draft.status || "em_analise") : undefined })} /><span>Tem estado associado</span></label>
                  {draft.hasStatus ? (
                    <label style={S.field}>
                      <span style={S.label}>Estado desta etiqueta</span>
                      <select style={S.select} value={draft.status || "em_analise"} onChange={(event) => updateLabelDraft(label, { status: event.target.value as EmailLabelStatus, hasStatus: true })}>
                        {LABEL_STATUS_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                      </select>
                    </label>
                  ) : null}
                </div>
              );
            }) : <PanelState compact tone="info" title="Sem etiquetas ainda" description="Vai adicionando etiquetas para testar esta estrutura nova." />}
          </div>
        </div>
      );
    }

    if (section === "filters") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Filtros da janela</div>
            <div style={S.grid2}>
              <label style={S.field}><span style={S.label}>Fonte da lista</span><select style={S.select} value={scopeMode} onChange={(event) => setScopeMode(event.target.value as ScopeMode)}><option value="related">So emails relacionados</option><option value="all">Todos os emails conhecidos</option></select></label>
              <label style={S.field}><span style={S.label}>Filtrar por grupo</span><select style={S.select} value={groupFilterId} onChange={(event) => setGroupFilterId(event.target.value)}><option value="">Sem filtro</option>{contextualGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}</select></label>
              <label style={S.field}><span style={S.label}>Filtrar por ticket</span><select style={S.select} value={ticketFilterId} onChange={(event) => setTicketFilterId(event.target.value)}><option value="">Sem filtro</option>{contextualTickets.map((ticket) => <option key={ticket.id} value={ticket.id}>{ticket.code} Â· {ticket.title}</option>)}</select></label>
              <label style={S.field}><span style={S.label}>Filtrar por etiqueta</span><select style={S.select} value={labelFilterValue} onChange={(event) => setLabelFilterValue(event.target.value)}><option value="">Sem filtro</option>{contextualLabels.map((label) => <option key={label} value={label}>{label}</option>)}</select></label>
            </div>
            <div style={S.inlineChecks}>
              <label style={S.check}><input type="checkbox" checked={onlyExternal} onChange={(event) => setOnlyExternal(event.target.checked)} /><span>So emails externos</span></label>
              <label style={S.check}><input type="checkbox" checked={onlyWithAttachments} onChange={(event) => setOnlyWithAttachments(event.target.checked)} /><span>So emails com anexos</span></label>
            </div>
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Resultado atual</div>
            <div style={S.summaryRow}><span>Emails visiveis</span><strong>{visibleEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Emails relacionados</span><strong>{classificationCase ? classificationRelatedEmails.length : classificationContextEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Total conhecido</span><strong>{classificationKnownEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Grupos neste conjunto</span><strong>{contextualGroups.length}</strong></div>
            <div style={S.summaryRow}><span>Tickets neste conjunto</span><strong>{contextualTickets.length}</strong></div>
            <div style={S.summaryRow}><span>Etiquetas neste conjunto</span><strong>{contextualLabels.length}</strong></div>
          </div>
        </div>
      );
    }

    if (section === "groups") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.titleRow}>
              <div>
                <div style={S.cardTitle}>Dossier do grupo</div>
                <div style={S.cardMeta}>Descricao, notas, emails, documentos e associacoes do grupo.</div>
              </div>
              <button type="button" style={S.secondaryBtn} onClick={() => { setSection("classification"); setClassificationFocus("principal"); }}>
                Voltar a Classificacao
              </button>
            </div>
            <div style={S.grid2}>
              <label style={S.field}>
                <span style={S.label}>Grupo a gerir</span>
                <select style={S.select} value={managedGroupId} onChange={(event) => setManagedGroupId(event.target.value)}>
                  <option value="">Escolher grupo...</option>
                  {manageableGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}
                </select>
              </label>
              <div style={S.summaryGrid}>
                <div style={S.summaryRow}><span>Grupo principal atual</span><strong>{principalGroup?.name || "--"}</strong></div>
                <div style={S.summaryRow}><span>Referencias atuais</span><strong>{referenceGroupSummary}</strong></div>
              </div>
            </div>
          </div>

          <div style={S.grid2Wide}>
            <div style={S.card}>
              <div style={S.cardTitle}>Descricao e notas</div>
              <div style={S.cardMeta}>Aqui mantemos o contexto base do grupo: descricao curta e notas operacionais relevantes.</div>
              <label style={S.field}>
                <span style={S.label}>Descricao do grupo</span>
                <textarea
                  style={S.textarea}
                  value={managedGroupDescription}
                  onChange={(event) => setManagedGroupDescription(event.target.value)}
                  placeholder={selectedManagedGroup ? "Descreve o objetivo deste grupo..." : "Escolhe primeiro um grupo"}
                  disabled={!selectedManagedGroup}
                />
              </label>
              <label style={S.field}>
                <span style={S.label}>Notas importantes</span>
                <textarea
                  style={{ ...S.textarea, minHeight: 110 }}
                  value={managedGroupNotes}
                  onChange={(event) => setManagedGroupNotes(event.target.value)}
                  placeholder={selectedManagedGroup ? "Notas operacionais, alertas e contexto util deste grupo..." : "Escolhe primeiro um grupo"}
                  disabled={!selectedManagedGroup}
                />
              </label>
              <div style={S.inline}>
                <button type="button" style={S.primaryBtn} onClick={() => void handleSaveManagedGroupProfile()} disabled={actionBusy || !selectedManagedGroup}>
                  <Icons.Save size={12} />
                  Guardar grupo
                </button>
              </div>
            </div>

            <div style={S.card}>
              <div style={S.cardTitle}>Pessoas e entidades</div>
              <div style={S.cardMeta}>Ligacoes do grupo a contactos e entidades reais do proprio caso. Para ja usamos contactos dos emails e do grupo; Outlook e Odoo entram depois.</div>
              <div style={S.summaryGrid}>
                <div style={S.summaryRow}><span>Contactos ligados</span><strong>{managedGroupContacts.length}</strong></div>
                <div style={S.summaryRow}><span>Entidades ligadas</span><strong>{managedGroupEntities.length}</strong></div>
              </div>
              <div style={S.grid2}>
                <div style={S.field}>
                  <span style={S.label}>Contactos do caso</span>
                  <input
                    style={S.input}
                    value={managedContactSearch}
                    onChange={(event) => setManagedContactSearch(event.target.value)}
                    placeholder={selectedManagedGroup ? "Pesquisar nome, email ou empresa..." : "Escolhe primeiro um grupo"}
                    disabled={!selectedManagedGroup}
                  />
                  <div style={S.inlineWrap}>
                    {managedGroupContacts.length ? managedGroupContacts.map((contact) => (
                      <button key={contact.key} type="button" style={S.selectedChipOn} onClick={() => toggleManagedGroupContact(contact)} disabled={!selectedManagedGroup}>
                        {contact.name}{contact.company ? ` Â· ${contact.company}` : ""}{contact.email ? ` Â· ${contact.email}` : ""}
                      </button>
                    )) : <span style={S.mutedMini}>Sem contactos associados.</span>}
                  </div>
                  <div style={S.chips}>
                    {selectedManagedGroup ? filteredManagedGroupContacts.slice(0, 18).map((contact) => {
                      const active = managedGroupContacts.some((entry) => entry.key === contact.key);
                      return (
                        <button key={contact.key} type="button" style={active ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleManagedGroupContact(contact)}>
                          {contact.name}{contact.company ? ` Â· ${contact.company}` : ""}{contact.email ? ` Â· ${contact.email}` : ""}
                        </button>
                      );
                    }) : null}
                  </div>
                </div>

                <div style={S.field}>
                  <span style={S.label}>Entidades do caso</span>
                  <input
                    style={S.input}
                    value={managedEntitySearch}
                    onChange={(event) => setManagedEntitySearch(event.target.value)}
                    placeholder={selectedManagedGroup ? "Pesquisar empresa, grupo ou origem..." : "Escolhe primeiro um grupo"}
                    disabled={!selectedManagedGroup}
                  />
                  <div style={S.inlineWrap}>
                    {managedGroupEntities.length ? managedGroupEntities.map((entity) => (
                      <button key={entity.key} type="button" style={S.selectedChipOn} onClick={() => toggleManagedGroupEntity(entity)} disabled={!selectedManagedGroup}>
                        {entity.name}{entity.kind ? ` Â· ${entity.kind}` : ""}
                      </button>
                    )) : <span style={S.mutedMini}>Sem entidades associadas.</span>}
                  </div>
                  <div style={S.chips}>
                    {selectedManagedGroup ? filteredManagedGroupEntities.slice(0, 18).map((entity) => {
                      const active = managedGroupEntities.some((entry) => entry.key === entity.key);
                      return (
                        <button key={entity.key} type="button" style={active ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleManagedGroupEntity(entity)}>
                          {entity.name}{entity.kind ? ` Â· ${entity.kind}` : ""}
                        </button>
                      );
                    }) : null}
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div style={S.grid2Wide}>
            <div style={S.card}>
              <div style={S.cardTitle}>Emails do grupo</div>
              <div style={S.cardMeta}>Lista real dos emails ligados ao grupo selecionado.</div>
              {managedGroupLoading ? (
                <PanelState compact tone="loading" title="A carregar emails do grupo" description="A preparar o dossier selecionado." />
              ) : !selectedManagedGroup ? (
                <PanelState compact tone="info" title="Escolhe um grupo" description="Seleciona primeiro o grupo que queres gerir." />
              ) : !managedGroupEmails.length ? (
                <PanelState compact tone="info" title="Sem emails ligados" description="Este grupo ainda nao tem emails ligados." />
              ) : (
                <div style={S.itemList}>
                  {managedGroupEmails.map((email) => (
                    <div key={makeEmailKey(email)} style={S.itemRow}>
                      <div style={S.itemMeta}>
                        <strong>{email.subject || "(sem assunto)"}</strong>
                        <small>{email.fromName || email.fromEmail || "--"} Â· {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</small>
                      </div>
                      <div style={S.inline}>
                        {(email.itemId || email.emailWebLink) ? (
                          <button type="button" style={S.secondaryBtn} onClick={() => void requestCockpitHostAction({ type: "open-email", itemId: email.itemId, emailWebLink: email.emailWebLink })}>
                            <Icons.ExternalLink size={12} />
                            Abrir
                          </button>
                        ) : null}
                        <button type="button" style={S.secondaryBtn} onClick={() => void handleRemoveManagedGroupEmail(email)} disabled={actionBusy}>
                          <Icons.Trash size={12} />
                          Remover
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={S.card}>
              <div style={S.cardTitle}>Documentos do grupo</div>
              <div style={S.cardMeta}>Documentos guardados neste grupo, com abertura e remocao.</div>
              {managedGroupLoading ? (
                <PanelState compact tone="loading" title="A carregar documentos" description="A abrir o dossier documental do grupo." />
              ) : !selectedManagedGroup ? (
                <PanelState compact tone="info" title="Escolhe um grupo" description="Seleciona primeiro o grupo que queres gerir." />
              ) : !managedGroupDocuments.length ? (
                <PanelState compact tone="info" title="Sem documentos guardados" description="Este grupo ainda nao tem documentos guardados." />
              ) : (
                <div style={S.itemList}>
                  {managedGroupDocuments.map((document) => (
                    <div key={document.id} style={S.itemRow}>
                      <div style={S.itemMeta}>
                        <strong>{document.name || "Documento"}</strong>
                        <small>{document.contentType || "ficheiro"}{document.size ? ` Â· ${Math.round(Number(document.size || 0) / 1024)} KB` : ""}</small>
                      </div>
                      <div style={S.inline}>
                        <a style={S.secondaryBtn} href={getGroupDocumentContentUrl(selectedManagedGroup.id, document.id)} target="_blank" rel="noreferrer">
                          <Icons.ExternalLink size={12} />
                          Abrir
                        </a>
                        <button type="button" style={S.secondaryBtn} onClick={() => void handleDeleteManagedGroupDocument(document)} disabled={actionBusy}>
                          <Icons.Trash size={12} />
                          Remover
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      );
    }

    return (
      <div style={S.stack}>
        <div style={S.card}>
          <div style={S.cardTitle}>Resumo vivo</div>
          <div style={S.cardMeta}>Espelho do estado atual. Os chips tambem servem para desligar antes de gravar.</div>
          <div style={S.summaryGrid}>
            <div style={S.summaryRow}><span>Email selecionado</span><strong>{selectedEmail?.subject || "--"}</strong></div>
            <div style={S.summaryRow}><span>Anexos</span><strong>{selectedEmailAttachments.length}</strong></div>
          </div>
        </div>

        <div style={S.sectionCard}>
          <div style={S.sectionHeadStatic}>
            <span style={S.sectionName}>Grupo principal</span>
            <span style={S.sectionMeta}>Casa principal atual do email</span>
          </div>
          <div style={S.sectionBody}>
            <div style={S.inlineWrap}>
              {principalGroup ? (
                <button type="button" style={S.selectedChipOn} onClick={clearPrincipalSelection}>
                  {principalGroup.name}
                </button>
              ) : (
                <span style={S.mutedMini}>Sem grupo principal</span>
              )}
            </div>
          </div>
        </div>

        <div style={S.sectionCard}>
          <div style={S.sectionHeadStatic}>
            <span style={S.sectionName}>Referencias</span>
            <span style={S.sectionMeta}>Grupos adicionais ligados ao email</span>
          </div>
          <div style={S.sectionBody}>
            <div style={S.inlineWrap}>
              {referenceGroups.length ? referenceGroups.map((group) => (
                <button key={group.id} type="button" style={S.selectedChipOn} onClick={() => toggleReferenceGroup(group.id)}>
                  {group.name}
                </button>
              )) : <span style={S.mutedMini}>Sem referencias</span>}
            </div>
          </div>
        </div>

        <div style={S.sectionCard}>
          <div style={S.sectionHeadStatic}>
            <span style={S.sectionName}>Ticket</span>
            <span style={S.sectionMeta}>Ticket ou novo ticket preparado</span>
          </div>
          <div style={S.sectionBody}>
            <div style={S.inlineWrap}>
              {selectedTicket ? (
                <button type="button" style={S.selectedChipOn} onClick={clearTicketSelection}>
                  {selectedTicket.code}
                </button>
              ) : selectedSeriesId ? (
                <button type="button" style={S.selectedChipPending} onClick={clearTicketSelection}>
                  {ticketSummary}
                </button>
              ) : (
                <span style={S.mutedMini}>Sem ticket</span>
              )}
            </div>
          </div>
        </div>

        <div style={S.sectionCard}>
          <div style={S.sectionHeadStatic}>
            <span style={S.sectionName}>Etiquetas</span>
            <span style={S.sectionMeta}>Etiquetas finais do email, com estado opcional</span>
          </div>
          <div style={S.sectionBody}>
            <div style={S.inlineWrap}>
              {summaryLabels.length ? summaryLabels.map((label) => (
                <button key={label} type="button" style={S.selectedChipOn} onClick={() => removeLabel(label)}>
                  {label}
                </button>
              )) : <span style={S.mutedMini}>Sem etiquetas</span>}
            </div>
            {labelStateSummary.length ? (
              <div style={S.stackMini}>
                <div style={S.cardMeta}>Etiquetas com estado</div>
                <div style={S.inlineWrap}>
                  {summaryLabels
                    .filter((label) => labelDrafts[label]?.hasStatus && labelDrafts[label]?.status)
                    .map((label) => (
                      <button
                        key={`${label}-status`}
                        type="button"
                        style={S.selectedChipPending}
                        onClick={() => updateLabelDraft(label, { hasStatus: false, status: undefined })}
                      >
                        {label}: {formatEmailLabelStatus(labelDrafts[label]?.status)}
                      </button>
                    ))}
                </div>
              </div>
            ) : null}
          </div>
        </div>

        <div style={S.card}>
          <div style={S.cardTitle}>Gravar / atualizar</div>
          <div style={S.cardMeta}>Quando estiver tudo certo, gravamos o estado atual do email selecionado.</div>
          <div style={S.summaryGrid}>
            <div style={S.summaryRow}>
              <span>Ambito</span>
              <strong>
                {applyScopeMode === "current"
                  ? "So email atual"
                  : applyScopeMode === "selected"
                    ? `Emails selecionados (${selectedTargetCount})`
                    : `Mesmo grupo principal (${principalScopeCount})`}
              </strong>
            </div>
          </div>
          <div style={S.inline}>
            <button
              type="button"
              style={S.primaryBtn}
              onClick={() => void handleApplyClassification()}
              disabled={
                actionBusy ||
                (!effectivePrincipalGroupId &&
                  !effectiveReferenceGroupIds.length &&
                  !selectedTicketId &&
                  !selectedSeriesId &&
                  !selectedEmailGroups.length &&
                  !selectedEmailTicketIds.length &&
                  !selectedLabels.length &&
                  !selectedEmailStoredLabels.length &&
                  !String(selectedEmail?.status || "").trim())
              }
            >
              <Icons.Save size={12} />
              Gravar / atualizar
            </button>
            <button type="button" style={S.secondaryBtn} onClick={() => setSection("classification")}>
              Voltar a Classificacao
            </button>
          </div>
        </div>
      </div>
    );
  }

  const dashboardStyle = classificationEditorActive ? S.dashboardFocus : S.dashboard;
  const topCardsGridStyle = classificationEditorActive ? S.topCardsGridFocus : S.topCardsGrid;
  const emailsCardStyle = classificationEditorActive ? S.focusEmailsCard : S.topCard;
  const quickDocumentsCardStyle = classificationEditorActive ? S.focusQuickDocumentsCard : S.topCard;
  const classificationCardStyle = classificationEditorActive ? S.focusClassificationCard : S.topCardWide;
  const previewShellStyle = classificationEditorActive ? S.focusPreviewShell : S.previewShellLarge;

  return (
    <div style={S.root} data-testid="studio-root">
      <div style={S.header}>
        <div style={S.headerMain}>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.mainTitle}>Studio de classificacao</div>
          <div style={S.mainMeta}>Janela nova e isolada para desenhar a futura atribuicao completa de grupos, tickets, etiquetas e filtros.</div>
          <div style={S.caseTitleRow}>
            <div style={S.caseTitle}>{caseTitle}</div>
            <div style={S.caseChips}>
              <span style={S.caseChip}>Cliente: {caseClient}</span>
              <span style={S.caseChip}>Marca: {caseBrand}</span>
              <span style={S.caseChip}>Estado: {caseState}</span>
            </div>
          </div>
        </div>
        <div style={S.headerActions}>
          <button type="button" style={S.secondaryBtn} onClick={() => setSection("groups")} disabled={!manageableGroups.length}>Renomear</button>
          <button type="button" style={S.secondaryBtn} onClick={() => setStatus("Fluxo de fundir preparado para a fase seguinte.")} disabled={!manageableGroups.length}>Fundir</button>
          <button data-testid="main-save-button" type="button" style={S.primaryBtn} onClick={() => openApplyDialog(classificationEditorActive ? classificationFocus : "summary")} disabled={actionBusy || !(classificationEditorActive ? canApplyFromClassificationEditor : canApplyClassification)}>
            <Icons.Save size={12} />
            Guardar
          </button>
          <button type="button" style={S.secondaryBtn} onClick={handleClose}>Fechar</button>
        </div>
      </div>

      <div style={S.context}>
        <div><div style={S.kicker}>Email atual</div><div style={S.contextTitle}>{selectedEmail?.subject || currentContext.subject || "(sem assunto)"}</div></div>
        <div style={S.badges}><span style={S.badge}>{selectedEmailAttachments.length} anexo(s)</span><span style={S.badge}>{relatedTickets.length} ticket(s)</span><span style={S.badge}>{relatedEmails.length} relacionados</span></div>
      </div>

      {status ? <div style={S.notice}>{status}</div> : null}

      <div style={dashboardStyle}>
        <div style={topCardsGridStyle}>

          <EmailsCard
            style={emailsCardStyle}
            loading={loading}
            visibleEmails={visibleEmails}
            selectedEmail={selectedEmail}
            emailSearch={emailSearch}
            setEmailSearch={setEmailSearch}
            selectAllVisibleEmails={selectAllVisibleEmails}
            clearSelectedTargets={clearSelectedTargets}
            selectedTargetCount={selectedTargetCount}
            selectedTargetEmailKeys={selectedTargetEmailKeys}
            toggleTargetEmailKey={toggleTargetEmailKey}
            expandedEmailKeys={expandedEmailKeys}
            toggleExpandedEmailKey={toggleExpandedEmailKey}
            setSelectedEmailKey={setSelectedEmailKey}
          />

          <QuickDocumentsCard
            style={quickDocumentsCardStyle}
            quickDocumentAttachments={quickDocumentAttachments}
            selectedAttachmentPreviewKey={selectedAttachmentPreviewKey}
            previewMode={previewMode}
            expandedQuickDocumentKeys={expandedQuickDocumentKeys}
            quickDocumentHiddenCount={quickDocumentHiddenCount}
            showHiddenQuickDocuments={showHiddenQuickDocuments}
            setShowHiddenQuickDocuments={setShowHiddenQuickDocuments}
            handleOpenQuickAttachment={handleOpenQuickAttachment}
            handleSetQuickAttachmentHidden={handleSetQuickAttachmentHidden}
            toggleExpandedQuickDocumentKey={toggleExpandedQuickDocumentKey}
            actionBusy={actionBusy}
          />

          <section style={classificationCardStyle}>
            <div style={S.sectionHeaderCompact}>
              <div>
                <div style={S.sectionTitle}>Classificacao</div>
                <div style={S.sectionSubtitle}>{classificationEditorActive ? "Editor aberto" : "Resumo do que esta atribuido"}</div>
              </div>
              <div style={S.segmentedControl}>
                <button data-testid="mode-normal-button" type="button" style={classificationLayoutMode === "normal" ? S.segmentBtnActive : S.segmentBtn} onClick={() => setClassificationLayoutMode("normal")}>Normal</button>
                <button data-testid="mode-advanced-button" type="button" style={classificationLayoutMode === "advanced" ? S.segmentBtnActive : S.segmentBtn} onClick={() => setClassificationLayoutMode("advanced")}>Avancado</button>
              </div>
            </div>
            {auxiliaryEditorActive ? (
              <div style={S.classificationEditorShell}>
                <div style={S.classificationEditorHeader}>
                  <button
                    type="button"
                    style={S.secondaryBtn}
                    onClick={() => {
                      setSection("emails");
                      setClassificationFocus("summary");
                    }}
                  >
                    Voltar
                  </button>
                  <div>
                    <div style={S.cardTitle}>{classificationCardTitle}</div>
                    <div style={S.cardMeta}>Editor contextual dentro do card Classificacao.</div>
                  </div>
                </div>
                <div style={S.classificationEditorBody}>{renderWorkspace()}</div>
              </div>
            ) : !classificationEditorActive ? (
              <ClassificationSummaryTiles
                tiles={classificationSummaryTiles}
                classificationLayoutMode={classificationLayoutMode}
                style={S.classificationSummary}
              />
            ) : (
              <div style={S.classificationEditorShell}>
                <ClassificationEditorHeader
                  classificationFocus={classificationFocus}
                  classificationLayoutMode={classificationLayoutMode}
                  onBack={handleCloseClassificationEditor}
                  onApply={openApplyDialog}
                  canApply={canApplyFromClassificationEditor}
                  actionBusy={actionBusy}
                />
                <div style={S.classificationEditorBody}>
                  <ClassificationEditor
                    data-testid="classification-editor"
                    classificationFocus={classificationFocus}
                    classificationLayoutMode={classificationLayoutMode}
                    classificationSuggestionExpanded={classificationSuggestionExpanded}
                    setClassificationSuggestionExpanded={setClassificationSuggestionExpanded}
                    suggestedExistingGroups={suggestedExistingGroups}
                    principalGroupId={principalGroupId}
                    clearPrincipalSelection={clearPrincipalSelection}
                    selectPrincipalGroup={selectPrincipalGroup}
                    principalSearch={principalSearch}
                    setPrincipalSearchValue={setPrincipalSearchValue}
                    principalCanCreate={principalCanCreate}
                    handleCreateGroupAndLink={handleCreateGroupAndLink}
                    exactPrincipalSearchGroup={exactPrincipalSearchGroup}
                    principalSearchResults={principalSearchResults}
                    principalGroup={principalGroup}
                    classificationMetaDraft={classificationMetaDraft}
                    updateClassificationMeta={updateClassificationMeta}
                    suggestedLabelSeeds={suggestedLabelSeeds}
                    selectedLabels={selectedLabels}
                    applySuggestedLabel={applySuggestedLabel}
                    classificationLabelInput={classificationLabelInput}
                    setClassificationLabelInput={setClassificationLabelInput}
                    handleClassificationLabelSearchAction={handleClassificationLabelSearchAction}
                    classificationLabelCanCreate={classificationLabelCanCreate}
                    filteredClassificationLabels={filteredClassificationLabels}
                    removeLabel={removeLabel}
                    addLabel={addLabel}
                    selectedLabelSharedStatus={selectedLabelSharedStatus}
                    updateLabelDraft={updateLabelDraft}
                    labelDrafts={labelDrafts}
                    LABEL_STATUS_OPTIONS={LABEL_STATUS_OPTIONS}
                    normalizedTicketSearch={normalizedTicketSearch}
                    ticketSearchResults={ticketSearchResults}
                    availableTicketChoices={ticketPickerChoices}
                    selectedTicket={selectedTicket}
                    selectedSeriesId={selectedSeriesId}
                    ticketEditorMode={ticketEditorMode}
                    setTicketEditorMode={setTicketEditorMode}
                    ticketSearch={ticketSearch}
                    setTicketSearch={setTicketSearch}
                    handleSearchTickets={handleSearchTickets}
                    ticketSearchBusy={ticketSearchBusy}
                    setSelectedSeriesId={setSelectedSeriesId}
                    setSelectionTouched={setSelectionTouched}
                    ticketSeries={ticketSeries}
                    createTicketTitle={createTicketTitle}
                    setCreateTicketTitle={setCreateTicketTitle}
                    ticketStatusDraft={ticketStatusDraft}
                    setTicketStatusDraft={setTicketStatusDraft}
                    TICKET_STATUS_OPTIONS={TICKET_STATUS_OPTIONS}
                    effectiveTicketStatus={effectiveTicketStatus}
                    ticketStatusLabel={ticketStatusLabel}
                    selectedTicketId={selectedTicketId}
                    applySuggestedTicket={applySuggestedTicket}
                    clearTicketSelection={clearTicketSelection}
                    referenceGroups={referenceGroups}
                    toggleReferenceGroup={toggleReferenceGroup}
                    referenceSearch={referenceSearch}
                    setReferenceSearchValue={setReferenceSearchValue}
                    referenceCanCreate={referenceCanCreate}
                    exactReferenceSearchGroup={exactReferenceSearchGroup}
                    referenceSearchResults={referenceSearchResults}
                    referenceGroupIds={referenceGroupIds}
                    actionBusy={actionBusy}
                  />
                </div>
              </div>
            )}
          </section>
        </div>

        <PreviewPane
          data-testid="preview-pane"
          previewShellStyle={previewShellStyle}
          previewMode={previewMode}
          setPreviewMode={setPreviewMode}
          previewHtml={previewHtml}
          previewHasDocument={previewHasDocument}
          selectedEmail={selectedEmail}
          selectedAttachmentPreview={selectedAttachmentPreview}
          selectedAttachmentDocumentPreview={selectedAttachmentDocumentPreview}
          selectedAttachmentPreviewRemoteStatus={selectedAttachmentPreviewRemoteStatus}
          selectedAttachmentPreviewMode={selectedAttachmentPreviewMode}
          handlePreviewReply={handlePreviewReply}
          handlePreviewForward={handlePreviewForward}
        />
      </div>
      <ApplyDialog
        data-testid="apply-dialog"
        isOpen={applyDialogOpen}
        onClose={() => setApplyDialogOpen(false)}
        section={applyDialogSection}
        scopeMode={applyDialogScopeMode}
        setScopeMode={setApplyDialogScope}
        currentScopeEmail={currentScopeEmail}
        caseScopeEmails={caseScopeEmails}
        selectedEmailKeys={applyDialogSelectedEmailKeys}
        setSelectedEmailKeys={setApplyDialogEmailKeys}
        expandedEmailKeys={applyDialogExpandedEmailKeys}
        toggleExpandedEmailKey={toggleApplyDialogExpandedEmailKey}
        toggleEmailKey={toggleApplyDialogEmailKey}
        status={status}
        actionBusy={actionBusy}
        handleConfirm={() => void handleConfirmApplyDialog()}
      />
    </div>
  );
}

export default function GroupClassificationStudioApp(): JSX.Element {
  return <StudioInner />;
}

const S: Record<string, React.CSSProperties> = {
  root: { height: "100vh", boxSizing: "border-box", padding: 12, display: "grid", gridTemplateRows: "auto auto auto auto minmax(0,1fr)", gap: 8, background: "linear-gradient(180deg, rgba(248,250,252,0.96) 0%, rgba(241,245,249,0.94) 100%)", color: "var(--iccc-text)", fontFamily: "var(--iccc-font)", overflow: "hidden" },
  header: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10, padding: "8px 10px", borderRadius: 14, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.88)", boxShadow: "0 10px 24px rgba(15,23,42,0.04)" },
  headerMain: { display: "grid", gap: 4, minWidth: 0 },
  headerActions: { display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 5, flexWrap: "wrap" },
  kicker: { fontSize: 10, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  mainTitle: { fontSize: 16, fontWeight: 650, color: "var(--iccc-text)" },
  mainMeta: { fontSize: 10.5, lineHeight: 1.3, color: "var(--iccc-muted)", maxWidth: 720 },
  caseTitleRow: { display: "grid", gap: 5 },
  caseTitle: { fontSize: 13, fontWeight: 650, color: "var(--iccc-text)" },
  caseChips: { display: "flex", gap: 5, flexWrap: "wrap" },
  caseChip: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(148,163,184,0.12)", color: "rgba(15,23,42,0.8)", fontSize: 9.5, fontWeight: 600 },
  primaryBtn: { height: 30, padding: "0 11px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "linear-gradient(180deg,#3b82f6 0%, #2563eb 100%)", color: "#fff", fontSize: 10.5, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", boxShadow: "0 4px 10px rgba(37,99,235,0.14)" },
  secondaryBtn: { height: 28, padding: "0 10px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)", fontSize: 10.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer" },
  context: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, padding: "7px 10px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.8)" },
  contextTitle: { fontSize: 12, fontWeight: 600, color: "var(--iccc-text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 780 },
  badges: { display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap", justifyContent: "flex-end" },
  badge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(30,64,175,0.08)", color: "#1d4ed8", fontSize: 9.5, fontWeight: 600 },
  notice: { padding: "7px 9px", borderRadius: 10, border: "1px solid #bfdbfe", background: "#eff6ff", color: "#1d4ed8", fontSize: 10.5, lineHeight: 1.35 },
  dashboard: { minHeight: 0, display: "grid", gridTemplateRows: "minmax(0,0.84fr) minmax(0,1.46fr)", gap: 8, overflow: "hidden" },
  topCardsGrid: { minHeight: 0, display: "grid", gridTemplateColumns: "minmax(0,1.04fr) minmax(0,0.88fr) minmax(0,1.16fr)", gap: 8, transition: "grid-template-columns 180ms ease" },
  dashboardFocus: { minHeight: 0, display: "grid", gridTemplateColumns: "minmax(240px,0.98fr) minmax(210px,0.76fr) minmax(520px,1.86fr)", gridTemplateRows: "minmax(0,1fr) minmax(0,0.88fr)", gap: 8, overflow: "hidden" },
  topCardsGridFocus: { display: "contents" },
  topCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.9)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 5, overflow: "hidden", transition: "transform 180ms ease, width 180ms ease, box-shadow 180ms ease" },
  topCardWide: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.9)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 5, overflow: "hidden", transition: "transform 180ms ease, width 180ms ease, box-shadow 180ms ease" },
  focusEmailsCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.92)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 5, overflow: "hidden", gridColumn: "1", gridRow: "1" },
  focusQuickDocumentsCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.88)", boxShadow: "0 6px 18px rgba(15,23,42,0.025)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 5, overflow: "hidden", gridColumn: "2", gridRow: "1" },
  focusClassificationCard: { minHeight: 0, borderRadius: 14, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(255,255,255,0.97)", boxShadow: "0 18px 36px rgba(37,99,235,0.08)", padding: 10, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 8, overflow: "hidden", gridColumn: "3", gridRow: "1 / span 2" },
  topCardScroll: { minHeight: 0, display: "grid", gap: 3, alignContent: "start", overflowY: "auto", paddingRight: 1 },
  sectionHeaderCompact: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 },
  sectionTitle: { fontSize: 9.5, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.1em", color: "rgba(15,23,42,0.82)" },
  sectionSubtitle: { fontSize: 9.5, color: "var(--iccc-muted)" },
  shell: { minHeight: 0, display: "grid", gridTemplateColumns: "220px 320px minmax(0,1fr)", gap: 12 },
  sidebar: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gap: 8, alignContent: "start", overflowY: "auto" },
  menu: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(148,163,184,0.2)", background: "rgba(255,255,255,0.78)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  menuOn: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.9)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  listCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 10, overflow: "hidden" },
  colTitle: { fontSize: 17, fontWeight: 800, color: "var(--iccc-text)" },
  emailTools: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 6, flexWrap: "wrap" },
  emailControlsRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", alignItems: "center", gap: 6 },
  emailToolsInline: { display: "inline-flex", alignItems: "center", justifyContent: "flex-end", gap: 8, flexWrap: "wrap" },
  input: { width: "100%", height: 30, boxSizing: "border-box", borderRadius: 9, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(248,250,252,0.92)", padding: "0 9px", fontSize: 11, color: "var(--iccc-text)", outline: "none" },
  textarea: { width: "100%", minHeight: 120, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "10px 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none", resize: "vertical" },
  select: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  listBody: { minHeight: 0, display: "grid", gap: 6, alignContent: "start", overflowY: "auto", paddingRight: 2 },
  email: { width: "100%", height: 30, boxSizing: "border-box", textAlign: "left", borderRadius: 8, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.78)", padding: "0 7px", display: "flex", alignItems: "center", gap: 6, cursor: "pointer", overflow: "hidden" },
  emailOn: { width: "100%", height: 30, boxSizing: "border-box", textAlign: "left", borderRadius: 8, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.96)", padding: "0 7px", display: "flex", alignItems: "center", gap: 6, cursor: "pointer", overflow: "hidden" },
  emailTop: { display: "flex", alignItems: "center", gap: 6, minWidth: 0, flex: 1, overflow: "hidden" },
  emailPick: { display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", alignItems: "center", gap: 6, minWidth: 0, cursor: "pointer" },
  emailTopRight: { display: "flex", alignItems: "center", justifyContent: "flex-end", gap: 5, minWidth: 0 },
  emailSubject: { fontSize: 10.25, fontWeight: 550, lineHeight: 1.15, color: "var(--iccc-text)", minWidth: 0, textAlign: "left", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  emailMeta: { fontSize: 8.75, color: "var(--iccc-muted)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 108 },
  emailSnippet: { maxHeight: 84, overflowY: "auto", padding: "6px 8px", borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.86)", color: "var(--iccc-text-soft, #334155)", fontSize: 9.4, lineHeight: 1.34, whiteSpace: "pre-wrap" },
  counter: { minWidth: 14, height: 14, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(15,23,42,0.06)", color: "var(--iccc-text)", fontSize: 8.4, fontWeight: 700 },
  quickDocList: { display: "grid", gap: 4, alignContent: "start" },
  quickDocLineMain: { display: "flex", alignItems: "center", minWidth: 0, gap: 6, flex: 1, overflow: "hidden" },
  quickDocRowHiddenTone: { opacity: 0.82 },
  quickDocStateBadge: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 44, height: 18, padding: "0 7px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.9)", color: "#64748b", fontSize: 8.75, fontWeight: 700 },
  quickDocActionBtn: { height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.92)", color: "#475569", fontSize: 9.25, fontWeight: 700, cursor: "pointer" },
  quickDocActionBtnOn: { height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(239,246,255,0.88)", color: "#1d4ed8", fontSize: 9.25, fontWeight: 700, cursor: "pointer" },
  quietToggleBtn: { height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.88)", color: "#64748b", fontSize: 9.25, fontWeight: 700, cursor: "pointer" },
  quietToggleBtnOn: { height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(239,246,255,0.88)", color: "#1d4ed8", fontSize: 9.25, fontWeight: 700, cursor: "pointer" },
  inlineActionBtn: { height: 24, padding: "0 9px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(239,246,255,0.88)", color: "#1d4ed8", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  workCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, overflow: "hidden" },
  stack: { height: "100%", minHeight: 0, display: "grid", gap: 10, alignContent: "start", overflowY: "auto", paddingRight: 2 },
  card: { borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.74)", padding: 12, display: "grid", gap: 10 },
  cardSticky: { position: "sticky", top: 0, zIndex: 4, borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.97)", padding: 12, display: "grid", gap: 10, boxShadow: "0 8px 24px rgba(15,23,42,0.06)" },
  segmentedControl: { display: "inline-flex", alignItems: "center", borderRadius: 999, border: "1px solid rgba(37,99,235,0.16)", overflow: "hidden", background: "rgba(239,246,255,0.66)" },
  segmentBtn: { height: 24, padding: "0 9px", border: "none", background: "transparent", color: "#475569", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  segmentBtnActive: { height: 24, padding: "0 9px", border: "none", background: "rgba(37,99,235,0.14)", color: "#1d4ed8", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  classificationSummary: { minHeight: 0, display: "grid", gap: 6, alignContent: "start", overflowY: "auto", paddingRight: 1 },
  legendRow: { display: "flex", flexWrap: "wrap", gap: 6, marginTop: 4 },
  legendChip: { padding: "3px 8px", borderRadius: 6, fontSize: 9.5, fontWeight: 700, border: "1px solid" },
  advancedHintBox: { display: "flex", flexWrap: "wrap", gap: 8 },
  advancedHintChip: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(239,246,255,0.72)", color: "#1d4ed8", fontSize: 9.5, fontWeight: 700 },
  classificationExtraGrid: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  classificationFooter: { display: "flex", justifyContent: "flex-start", paddingTop: 4 },
  classificationEditorShell: { minHeight: 0, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 8, overflow: "hidden" },
  classificationEditorHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap" },
  classificationEditorBody: { minHeight: 0, overflow: "auto", paddingRight: 2 },
  editorHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10, flexWrap: "wrap" },
  editorHeaderMeta: { display: "grid", gap: 3 },
  editorHeaderTitle: { fontSize: 13.5, fontWeight: 650, color: "var(--iccc-text)" },
  editorHeaderActions: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  editorModeText: { fontSize: 10, color: "var(--iccc-muted)" },
  editorPanelStack: { display: "grid", gap: 10, alignContent: "start" },
  editorModeKicker: { fontSize: 11, fontWeight: 700, letterSpacing: "0.12em", textTransform: "uppercase", color: "#1d4ed8" },
  editorLead: { fontSize: 11, lineHeight: 1.4, color: "var(--iccc-text-soft, #334155)" },
  editorBlock: { display: "grid", gap: 8, padding: 10, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.84)" },
  editorBlockHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  editorBlockTitle: { fontSize: 10.5, fontWeight: 700, color: "var(--iccc-text)" },
  editorValueStrong: { fontSize: 12.5, fontWeight: 600, color: "var(--iccc-text)" },
  editorExpandableOpen: { borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.82)", padding: "7px 9px" },
  editorExpandableScroll: { display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 96, overflowY: "auto", alignContent: "flex-start" },
  editorExpandableHint: { fontSize: 9.5, lineHeight: 1.35, color: "var(--iccc-muted)" },
  chipGridCompact: { display: "flex", flexWrap: "wrap", gap: 6 },
  editorOptionGrid: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  editorOptionStackLoose: { display: "grid", gap: 12 },
  editorLegendWrap: { paddingTop: 2 },
  editorAdvancedFieldGrid: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  compactCheck: { display: "flex", alignItems: "center", gap: 8, fontSize: 10.5, color: "var(--iccc-text)" },
  compactCheckBoxField: { minHeight: 34, display: "flex", alignItems: "center", gap: 8, padding: "0 10px", borderRadius: 10, border: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.88)", fontSize: 10.5, color: "var(--iccc-text)" },
  searchInlineRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center" },
  searchResultListCompact: { display: "grid", gap: 6, maxHeight: 172, overflowY: "auto", paddingRight: 1 },
  chevronBtn: { width: 20, height: 20, borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.88)", color: "#475569", fontSize: 11, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  editorSplitRow: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  editorModeBtn: { minHeight: 42, padding: "0 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.86)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 600, textAlign: "left", cursor: "pointer" },
  editorModeBtnOn: { minHeight: 42, padding: "0 12px", borderRadius: 12, border: "1px solid rgba(37,99,235,0.22)", background: "rgba(219,234,254,0.9)", color: "#1d4ed8", fontSize: 11, fontWeight: 700, textAlign: "left", cursor: "pointer" },
  previewShellLarge: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.92)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 6, overflow: "hidden", transition: "width 180ms ease, max-width 180ms ease, transform 180ms ease, grid-column 180ms ease" },
  focusPreviewShell: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.92)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 6, overflow: "hidden", gridColumn: "1 / span 2", gridRow: "2", width: "100%", maxWidth: "100%", minWidth: 0, justifySelf: "stretch" },
  previewToolbar: { display: "flex", alignItems: "center", gap: 5, flexWrap: "wrap", paddingBottom: 1, borderBottom: "1px solid rgba(148,163,184,0.1)" },
  previewTab: { height: 24, padding: "0 9px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  previewTabOn: { height: 24, padding: "0 9px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(219,234,254,0.9)", color: "#1d4ed8", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  previewBody: { minHeight: 0, overflow: "auto", paddingRight: 1, display: "grid", gap: 6, alignContent: "start" },
  previewPlaceholder: { minHeight: 400, borderRadius: 12, border: "1px dashed rgba(148,163,184,0.24)", background: "rgba(248,250,252,0.82)", display: "grid", alignContent: "center", justifyItems: "start", gap: 8, padding: 18 },
  sectionCard: { borderRadius: 16, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.78)", overflow: "hidden", display: "grid" },
  classificationSectionCard: { borderRadius: 16, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.78)", overflow: "hidden", display: "grid", scrollMarginTop: 168 },
  sectionHead: { width: "100%", border: "none", borderBottom: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.58)", color: "var(--iccc-text)", padding: "10px 14px", display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, cursor: "pointer" },
  sectionHeadOn: { width: "100%", border: "none", borderBottom: "1px solid rgba(37,99,235,0.18)", background: "rgba(239,246,255,0.9)", color: "#1d4ed8", padding: "10px 14px", display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, cursor: "pointer" },
  sectionHeadStatic: { borderBottom: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.58)", color: "var(--iccc-text)", padding: "10px 14px", display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12 },
  sectionName: { fontSize: 13, fontWeight: 700 },
  sectionMeta: { fontSize: 10, color: "var(--iccc-muted)" },
  sectionBody: { padding: 12, display: "grid", gap: 10 },
  sectionBodyScroll: { padding: 12, display: "grid", gap: 10, maxHeight: "min(52vh, 520px)", overflowY: "auto", paddingRight: 8 },
  stackMini: { display: "grid", gap: 6 },
  fieldLineLabel: { fontSize: 10, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  compactRowWrap: { display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  sectionControls: { display: "grid", gridTemplateColumns: "minmax(0,1fr) 260px", gap: 10 },
  compactCreateRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center" },
  compactSearchActionRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) 34px", gap: 8, alignItems: "center" },
  searchActionRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) 34px 34px", gap: 8, alignItems: "center" },
  iconActionBtn: { width: 34, height: 34, borderRadius: 10, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.92)", color: "#1d4ed8", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  iconActionBtnDisabled: { width: 34, height: 34, borderRadius: 10, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.78)", color: "rgba(100,116,139,0.55)", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "not-allowed" },
  inlineWrap: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  selectedChipOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 12, fontWeight: 700, padding: "7px 11px", cursor: "pointer" },
  selectedChipPending: { borderRadius: 999, border: "1px solid rgba(245,158,11,0.24)", background: "rgba(254,243,199,0.92)", color: "#b45309", fontSize: 12, fontWeight: 700, padding: "7px 11px", cursor: "pointer" },
  miniChip: { borderRadius: 999, border: "1px solid rgba(148,163,184,0.22)", background: "rgba(255,255,255,0.94)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 600, padding: "5px 9px", cursor: "pointer" },
  miniChipOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 11, fontWeight: 700, padding: "5px 9px", cursor: "pointer" },
  mutedMini: { fontSize: 12, color: "var(--iccc-muted)" },
  classificationHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap" },
  suggestionDock: { marginTop: 10, display: "grid", gap: 8, padding: "10px 12px", borderRadius: 14, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(248,250,252,0.9)" },
  suggestionDockMeta: { fontSize: 11, color: "var(--iccc-muted)" },
  suggestionDockChips: { display: "flex", flexWrap: "wrap", gap: 6 },
  suggestionDockChip: { borderRadius: 999, border: "1px solid rgba(148,163,184,0.24)", background: "rgba(255,255,255,0.98)", color: "var(--iccc-muted)", fontSize: 10, fontWeight: 700, padding: "4px 8px", cursor: "pointer" },
  suggestionDockChipOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 10, fontWeight: 700, padding: "4px 8px", cursor: "pointer" },
  classificationFocusBar: { display: "grid", gridTemplateColumns: "repeat(5,minmax(0,1fr))", gap: 0, borderRadius: 12, overflow: "hidden", border: "1px solid rgba(37,99,235,0.24)", background: "rgba(239,246,255,0.75)" },
  classificationFocusBtn: { height: 30, border: "none", borderRight: "1px solid rgba(37,99,235,0.24)", background: "transparent", color: "#475569", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  classificationFocusBtnOn: { height: 30, border: "none", borderRight: "1px solid rgba(37,99,235,0.24)", background: "rgba(37,99,235,0.16)", color: "#1d4ed8", fontSize: 11, fontWeight: 800, cursor: "pointer" },
  titleRow: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  cardTitle: { fontSize: 12, fontWeight: 650, color: "var(--iccc-text)" },
  cardMeta: { fontSize: 9.5, lineHeight: 1.25, color: "var(--iccc-muted)" },
  metaLine: { display: "flex", gap: 12, flexWrap: "wrap", fontSize: 11, color: "var(--iccc-muted)" },
  chips: { display: "flex", flexWrap: "wrap", gap: 8 },
  groupChip: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, background: "rgba(29,78,216,0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  groupChipBtn: { borderRadius: 999, border: "1px solid rgba(148,163,184,0.24)", background: "rgba(255,255,255,0.92)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  groupChipBtnOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  searchResultList: { display: "grid", gap: 6 },
  searchResultBtn: { width: "100%", borderRadius: 10, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 600, padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", textAlign: "left" },
  searchResultBtnOn: { width: "100%", borderRadius: 10, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 12, fontWeight: 700, padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", textAlign: "left" },
  resultMiniMeta: { fontSize: 10, fontWeight: 700, color: "inherit", opacity: 0.85 },
  preview: { width: "100%", minHeight: 520, borderRadius: 14, overflow: "hidden", border: "1px solid rgba(148,163,184,0.24)", background: "#fff" },
  previewHtml: { width: "100%", minHeight: 540, height: "100%", overflow: "auto", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", boxShadow: "inset 0 1px 0 rgba(255,255,255,0.45)" },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 },
  grid2Wide: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 },
  field: { display: "grid", gap: 6 },
  label: { fontSize: 11, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  subTitle: { fontSize: 12, fontWeight: 800, color: "var(--iccc-text)" },
  inline: { display: "flex", alignItems: "center", gap: 8 },
  labelRow: { borderRadius: 14, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", padding: 12, display: "grid", gap: 8 },
  labelRowCompact: { borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.7)", padding: 10, display: "grid", gap: 8 },
  labelGrid: { display: "grid", gap: 8 },
  labelHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  linkBtn: { border: "none", background: "transparent", color: "#2563eb", fontSize: 12, fontWeight: 700, cursor: "pointer", padding: 0 },
  check: { display: "inline-flex", alignItems: "center", gap: 8, fontSize: 12, color: "var(--iccc-text)" },
  inlineChecks: { display: "flex", gap: 16, flexWrap: "wrap" },
  attachmentPickerBar: { display: "flex", flexWrap: "wrap", gap: 8 },
  documentPreviewShell: { minHeight: 0, height: "100%", display: "grid", alignContent: "stretch", gap: 6 },
  documentPreviewFrame: { borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", overflow: "hidden", background: "#f8fafc", minHeight: 0, height: "100%", boxShadow: "inset 0 1px 0 rgba(255,255,255,0.45)" },
  documentPreviewIframe: { width: "100%", height: "100%", minHeight: 540, border: "none", display: "block", background: "#fff" },
  attachmentPreviewWrap: { borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#f8fafc", overflow: "hidden", minHeight: 0, height: "100%", boxShadow: "inset 0 1px 0 rgba(255,255,255,0.45)" },
  attachmentPreviewImage: { width: "100%", height: "100%", minHeight: 540, objectFit: "contain", display: "block", background: "#fff" },
  attachmentPdfPreviewShell: { display: "grid", height: "100%", minHeight: 0, background: "#f8fafc", borderRadius: 12, overflow: "hidden", border: "1px solid rgba(15, 23, 42, 0.08)", boxShadow: "inset 0 1px 0 rgba(255,255,255,0.45)" },
  attachmentPdfPreviewLoading: { display: "grid", placeItems: "center", minHeight: 220, padding: 18, color: "var(--iccc-muted)", fontSize: 10.5 },
  attachmentPdfPreviewCanvasHost: { overflow: "auto", padding: 12, display: "grid", justifyItems: "center", alignContent: "start", gap: 12, minHeight: 0, height: "100%" },
  attachmentPreviewText: { margin: 0, padding: 12, background: "#f8fafc", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", fontFamily: "Consolas, monospace", fontSize: 10.5, lineHeight: 1.42, whiteSpace: "pre-wrap", height: "100%", overflow: "auto", boxSizing: "border-box" },
  attachmentPreviewEmpty: { padding: "14px 12px", borderRadius: 12, border: "1px dashed rgba(148,163,184,0.24)", background: "rgba(248,250,252,0.82)", color: "var(--iccc-muted)", fontSize: 10.5 },
  attachList: { display: "grid", gap: 10 },
  attachRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 12, alignItems: "center", padding: "10px 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)" },
  attachMeta: { display: "grid", gap: 3, minWidth: 0, color: "var(--iccc-text)" },
  attachChecks: { display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "flex-end" },
  itemList: { display: "grid", gap: 10 },
  itemRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 12, alignItems: "center", padding: "10px 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)" },
  itemMeta: { display: "grid", gap: 4, minWidth: 0, color: "var(--iccc-text)" },
  similarMainBtn: { border: "none", background: "transparent", padding: 0, margin: 0, textAlign: "left", display: "grid", minWidth: 0, cursor: "pointer" },
  summaryRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "9px 11px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", fontSize: 12, color: "var(--iccc-text)" },
  summaryGrid: { display: "grid", gap: 8 },
  summaryActionBar: { position: "sticky", bottom: -12, display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap", paddingTop: 12, paddingBottom: 4, background: "linear-gradient(180deg, rgba(255,255,255,0) 0%, rgba(255,255,255,0.96) 16%, rgba(255,255,255,0.98) 100%)" },
  note: { padding: "12px 14px", borderRadius: 14, border: "1px solid rgba(191,219,254,0.8)", background: "#eff6ff", color: "#1d4ed8", fontSize: 13, lineHeight: 1.5 },
  modalBackdrop: { position: "fixed", inset: 0, background: "rgba(15,23,42,0.18)", display: "grid", placeItems: "center", padding: 20, zIndex: 60 },
  modalSheet: { width: "min(860px, 100%)", maxHeight: "min(84vh, 920px)", overflow: "hidden", borderRadius: 18, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.98)", boxShadow: "0 24px 60px rgba(15,23,42,0.18)", display: "grid", gridTemplateRows: "auto auto minmax(0,1fr) auto", gap: 12, padding: 16 },
  modalHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  modalTitle: { fontSize: 14, fontWeight: 650, color: "var(--iccc-text)" },
  modalScopeRow: { display: "grid", gridTemplateColumns: "repeat(3,minmax(0,1fr))", gap: 8 },
  scopeChip: { minHeight: 38, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)", fontSize: 10.75, fontWeight: 600, cursor: "pointer" },
  scopeChipOn: { minHeight: 38, borderRadius: 12, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 10.75, fontWeight: 700, cursor: "pointer" },
  modalBlock: { minHeight: 0, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 8 },
  modalBlockHeader: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 },
  applyEmailList: { minHeight: 0, overflowY: "auto", display: "grid", gap: 8, paddingRight: 2 },
  applyEmailRow: { borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(248,250,252,0.84)", padding: "9px 10px", display: "grid", gap: 8 },
  applyEmailRowOn: { borderRadius: 12, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.9)", padding: "9px 10px", display: "grid", gap: 8 },
  applySingleEmailSummary: { borderRadius: 12, border: "1px solid rgba(37,99,235,0.16)", background: "rgba(239,246,255,0.76)", padding: "10px 11px", display: "grid", gap: 6 },
  applyEmailSummaryHead: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center" },
  applyEmailSummaryTitle: { fontSize: 11.25, fontWeight: 600, color: "var(--iccc-text)", lineHeight: 1.2, minWidth: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  applyEmailSummaryMeta: { fontSize: 9.5, color: "var(--iccc-muted)", lineHeight: 1.2 },
  applyEmailRowTop: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center" },
  applyEmailMain: { display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", alignItems: "center", gap: 8, minWidth: 0 },
  applyEmailRowTail: { display: "flex", alignItems: "center", gap: 6, minWidth: 0, justifyContent: "flex-end" },
  applyEmailSubject: { fontSize: 10.75, fontWeight: 600, color: "var(--iccc-text)", lineHeight: 1.2, display: "block", minWidth: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  applyEmailMeta: { fontSize: 9.25, color: "var(--iccc-muted)", lineHeight: 1.2, display: "block", maxWidth: 160, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  applyScopeBadge: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 56, height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.92)", color: "#64748b", fontSize: 9, fontWeight: 700 },
  applyScopeBadgeOn: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 56, height: 22, padding: "0 8px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(239,246,255,0.92)", color: "#1d4ed8", fontSize: 9, fontWeight: 700 },
  applyEmailPreview: { maxHeight: 92, overflowY: "auto", padding: "7px 9px", borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(255,255,255,0.9)", color: "var(--iccc-text-soft, #334155)", fontSize: 10, lineHeight: 1.4, whiteSpace: "pre-wrap" },
  modalFooter: { display: "flex", justifyContent: "flex-end", gap: 8, flexWrap: "wrap" },
};

