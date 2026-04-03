import React, { useEffect, useMemo, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import { addEmailToLinkGroup, createGroupTicket, createLinkGroup, deleteGroupDocument, extractAttachmentTexts, getEmailAttachmentContentBase64, getEmailAttachmentContentUrl, getEmailAttachmentTextContent, getGroupDocumentContentUrl, getGroupDocuments, getGroupEmails, getRelatedEmailContext, linkEmailToGroupTicket, listLinkGroups, listGroupTicketSeries, registerRelevantEmail, removeEmailFromLinkGroup, saveGroupDocuments, searchGroupTickets, searchKnownEmails, unlinkEmailFromGroupTicket, updateGroupTicket, updateLinkGroup, type GroupDocumentEntry, type GroupTicketEntry, type GroupTicketSeriesEntry, type LinkGroupEntry, type RelatedEmailEntry, type RelevantEmailPayload } from "@/api";
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
import "../../global.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

type SectionId = "emails" | "classification" | "labels" | "filters" | "groups";
type ScopeMode = "related" | "all";
type ApplyScopeMode = "current" | "selected" | "principal_group";
type ApplyDialogScopeMode = "current" | "selected" | "case_all";
type PreviewMode = "email" | "document" | "reply" | "forward";
type ClassificationLayoutMode = "normal" | "advanced";
type EmailLabelStatus = GroupLabelStatus;
type DocumentLifecycleState = "ingested" | "processed" | "accepted" | "rejected" | "reread_requested";
type ClassificationFocus = "principal" | "references" | "labels" | "ticket" | "summary";
type TicketEditorMode = "existing" | "new";
type AttachmentPreviewState =
  | { kind: "image"; src: string }
  | { kind: "pdf"; src: string }
  | { kind: "office"; url: string }
  | { kind: "text"; text: string }
  | { kind: "unsupported" };
type LabelDraft = { categorize: boolean; hasStatus: boolean; status?: EmailLabelStatus };
type ReadingSuggestionChip = { key: string; label: string; kind: "group" | "ticket" | "label"; value: string };
type GroupContactDraft = { key: string; name: string; email?: string; company?: string; source?: string };
type GroupEntityDraft = { key: string; name: string; kind?: string; source?: string };
type ClassificationMetaDraft = {
  principalCategorize: boolean;
  principalStatusEnabled: boolean;
  principalStatusCategorize: boolean;
  referenceCategorize: boolean;
  referenceStatusEnabled: boolean;
  referenceStatusCategorize: boolean;
  ticketStatusEnabled: boolean;
  ticketStatusCategorize: boolean;
  categorizedLabelNames?: string[];
};
type CaseGroupEntry = LinkGroupEntry & { relationKind?: string };
type StudioParams = {
  conversationId?: string;
  internetMessageId?: string;
  itemId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  receivedAtIso?: string;
  seedKey?: string;
};

const GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX = "iccc_group_classification_seed_v1:";

function isOutlookCategorySyncDebugEnabled(): boolean {
  try {
    return window.localStorage?.getItem(OUTLOOK_CATEGORY_SYNC_DEBUG_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

function logClassificationOutlookCategorySync(event: string, data?: any) {
  if (!isOutlookCategorySyncDebugEnabled()) return;
  clientLog.debug(`[outlook-category-sync] classification ${event}`, data);
}

const MENU: Array<{ id: SectionId; label: string; icon: React.ReactNode; help: string }> = [
  { id: "emails", label: "Emails", icon: <Icons.MessageSquare size={15} />, help: "Lista e preview base do caso." },
  { id: "classification", label: "Classificacao", icon: <Icons.Target size={15} />, help: "Grupo principal, referencias e ticket." },
  { id: "labels", label: "Etiquetas", icon: <Icons.Star size={15} />, help: "Etiquetas e futuras categorias Outlook." },
  { id: "filters", label: "Filtros", icon: <Icons.Search size={15} />, help: "Reducao da lista e testes de vista." },
  { id: "groups", label: "Grupos", icon: <Icons.Building size={15} />, help: "Gestao do grupo como dossier." },
];

const LABEL_STATUS_OPTIONS: Array<{ value: EmailLabelStatus; label: string }> = [
  { value: "em_analise", label: "Em analise" },
  { value: "em_progresso", label: "Em progresso" },
  { value: "concluido", label: "Concluido" },
];

const TICKET_STATUS_OPTIONS: Array<{ value: string; label: string }> = [
  { value: "", label: "Sem estado" },
  { value: "open", label: "Aberto" },
  { value: "em_analise", label: "Em analise" },
  { value: "em_progresso", label: "Em progresso" },
  { value: "concluido", label: "Concluido" },
  { value: "closed", label: "Fechado" },
];

const DOCUMENT_STATE_OPTIONS: Array<{ value: DocumentLifecycleState; label: string }> = [
  { value: "ingested", label: "Recebido" },
  { value: "processed", label: "Processado" },
  { value: "accepted", label: "Aceite" },
  { value: "rejected", label: "Rejeitado" },
  { value: "reread_requested", label: "Reler" },
];

const OUTLOOK_CATEGORY_COLOR_LEGEND = [
  { key: "blue", label: "Azul = Em analise", style: { borderColor: "rgba(59,130,246,0.34)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8" } },
  { key: "amber", label: "Amarelo = Aguarda", style: { borderColor: "rgba(245,158,11,0.3)", background: "rgba(254,243,199,0.95)", color: "#b45309" } },
  { key: "green", label: "Verde = Concluido", style: { borderColor: "rgba(34,197,94,0.28)", background: "rgba(220,252,231,0.95)", color: "#15803d" } },
  { key: "red", label: "Vermelho = Bloqueado", style: { borderColor: "rgba(239,68,68,0.26)", background: "rgba(254,226,226,0.95)", color: "#b91c1c" } },
];

const EMPTY_CLASSIFICATION_META: ClassificationMetaDraft = {
  principalCategorize: true,
  principalStatusEnabled: false,
  principalStatusCategorize: false,
  referenceCategorize: true,
  referenceStatusEnabled: false,
  referenceStatusCategorize: false,
  ticketStatusEnabled: false,
  ticketStatusCategorize: false,
};

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

function formatEmailLabelStatus(value: string | undefined): string {
  if (value === "concluido") return "Concluido";
  if (value === "em_progresso") return "Em progresso";
  return "Em analise";
}

function formatGroupStatusLabel(value: string | undefined): string {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "concluido") return "Concluido";
  if (normalized === "em_progresso") return "Em progresso";
  if (normalized === "em_analise") return "Em analise";
  return String(value || "").trim() || "--";
}

function formatTicketStatusLabel(value: string | undefined): string {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "open") return "Aberto";
  if (normalized === "closed") return "Fechado";
  return formatGroupStatusLabel(value);
}

function normalizeDocumentLifecycleState(value: string | undefined, fallback: DocumentLifecycleState = "ingested"): DocumentLifecycleState {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "processed" || normalized === "accepted" || normalized === "rejected" || normalized === "reread_requested") {
    return normalized;
  }
  if (normalized === "ingested") return "ingested";
  return fallback;
}

function formatDocumentLifecycleState(value: string | undefined): string {
  const normalized = normalizeDocumentLifecycleState(value);
  const match = DOCUMENT_STATE_OPTIONS.find((entry) => entry.value === normalized);
  return match?.label || "Recebido";
}

function isRejectedDocumentLifecycleState(value: string | undefined): boolean {
  return normalizeDocumentLifecycleState(value) === "rejected";
}

function inferStudioAttachmentKind(
  attachment: ReturnType<typeof normalizeStudioAttachment>
): AttachmentPreviewState["kind"] | "none" {
  if (!attachment) return "none";
  const name = String(attachment.name || "").toLowerCase();
  const type = normalizeStudioAttachmentMimeType(attachment.contentType, attachment.name);
  if (type.startsWith("image/") || /\.(png|jpe?g|gif|bmp|webp|svg)$/.test(name)) return "image";
  if (type.includes("pdf") || /\.pdf$/.test(name)) return "pdf";
  if (
    type === "application/msword"
    || type === "application/vnd.ms-excel"
    || type === "application/vnd.ms-powerpoint"
    || type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    || type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    || type === "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    || /\.(docx?|xlsx?|pptx?)$/.test(name)
  ) {
    return "office";
  }
  if (type.startsWith("text/") || type.includes("json") || type.includes("xml") || type.includes("csv") || /\.(txt|md|json|xml|csv|log|ya?ml)$/.test(name)) return "text";
  return "unsupported";
}

function isLikelyDecorativeAttachment(
  attachment: ReturnType<typeof normalizeStudioAttachment>
): boolean {
  if (!attachment) return false;
  const type = normalizeStudioAttachmentMimeType(attachment.contentType, attachment.name);
  const name = String(attachment.name || "").trim().toLowerCase();
  const size = Number(attachment.size || 0) || 0;
  const hasContentId = Boolean(String(attachment.contentId || "").trim());
  const signatureLikeName = /\b(signature|assinatura|logo|smime|favicon)\b/.test(name);
  const genericInlineImageName = /^image\d+\.(png|jpe?g|gif|bmp|webp)$/i.test(name);
  if (attachment.isHidden === true) return true;
  if (attachment.isHidden === false) return false;
  if (attachment.isInline) return true;
  if (!/^image\//.test(type)) return false;
  if (signatureLikeName && (hasContentId || size <= 48 * 1024)) return true;
  if (genericInlineImageName && hasContentId) return true;
  return Boolean(hasContentId && size > 0 && size <= 48 * 1024);
}

function isStudioAttachmentHiddenInQuickDocs(
  attachment: ReturnType<typeof normalizeStudioAttachment>
): boolean {
  if (!attachment) return false;
  if (attachment.isHidden === true) return true;
  if (attachment.isHidden === false) return false;
  return isLikelyDecorativeAttachment(attachment);
}

function formatQuickDocumentMeta(
  attachment: ReturnType<typeof normalizeStudioAttachment>
): string {
  if (!attachment) return "";
  const name = String(attachment.name || "").trim();
  const ext = name.includes(".") ? name.split(".").pop() || "" : "";
  const kind = inferStudioAttachmentKind(attachment);
  let typeLabel = "";
  if (kind === "pdf") typeLabel = "PDF";
  else if (kind === "office") typeLabel = (ext || "office").toUpperCase();
  else if (kind === "image") typeLabel = (ext || "imagem").toUpperCase();
  else if (kind === "text") typeLabel = "Texto";
  else typeLabel = (ext || "ficheiro").toUpperCase();
  const size = Number(attachment.size || 0) || 0;
  const sizeLabel = size > 0 ? `${Math.max(1, Math.round(size / 1024))} KB` : "";
  return [typeLabel, sizeLabel].filter(Boolean).join(" · ");
}

function normalizeClassificationMetaDraft(
  value?: Partial<ClassificationMetaDraft> | null
): ClassificationMetaDraft {
  return {
    principalCategorize: value?.principalCategorize !== false,
    principalStatusEnabled: value?.principalStatusEnabled === true,
    principalStatusCategorize: value?.principalStatusCategorize === true,
    referenceCategorize: value?.referenceCategorize !== false,
    referenceStatusEnabled: value?.referenceStatusEnabled === true,
    referenceStatusCategorize: value?.referenceStatusCategorize === true,
    ticketStatusEnabled: value?.ticketStatusEnabled === true,
    ticketStatusCategorize: value?.ticketStatusCategorize === true,
    categorizedLabelNames: Array.isArray(value?.categorizedLabelNames)
      ? value.categorizedLabelNames.map((label) => String(label || "").trim()).filter(Boolean)
      : [],
  };
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  const emailKey = String(email?.emailKey || "").trim();
  const emailId = String(email?.id || "").trim();
  const itemId = String(email?.itemId || "").trim();
  const internetMessageId = String(email?.internetMessageId || "").trim().replace(/[<>\s]/g, "").toLowerCase();
  const isSynthetic = (value: string) => /^email_[0-9a-f-]+$/i.test(value);

  return String(
    itemId
    || internetMessageId
    || (emailKey && !isSynthetic(emailKey) ? emailKey : "")
    || (emailId && !isSynthetic(emailId) ? emailId : "")
    || emailKey
    || emailId
    || [
      String(email?.conversationId || "").trim(),
      String(email?.subject || "").trim().toLowerCase(),
      String(email?.fromEmail || "").trim().toLowerCase(),
      String(email?.messageDateIso || email?.receivedAtIso || "").trim(),
    ].filter(Boolean).join("|")
  );
}

function mergeUniqueStrings(values: string[]): string[] {
  const seen = new Set<string>();
  const next: string[] = [];
  for (const value of values || []) {
    const normalized = String(value || "").trim();
    if (!normalized || seen.has(normalized)) continue;
    seen.add(normalized);
    next.push(normalized);
  }
  return next;
}

function mergeUniqueBy<T>(values: T[], getKey: (value: T) => string): T[] {
  const seen = new Set<string>();
  const next: T[] = [];
  for (const value of values || []) {
    const key = String(getKey(value) || "").trim();
    if (!key || seen.has(key)) continue;
    seen.add(key);
    next.push(value);
  }
  return next;
}

function scoreStudioAttachment(attachment: any): number {
  const normalized = normalizeStudioAttachment(attachment);
  if (!normalized?.name) return 0;
  let score = 10;
  if (String(normalized.content || "").trim()) score += 40;
  if (normalized.hasContent === true) score += 25;
  if (String(normalized.key || "").trim()) score += 18;
  if (String(normalized.id || "").trim()) score += 10;
  if (String(normalized.contentId || "").trim()) score += 6;
  if (String((attachment as any)?.storagePathHint || "").trim()) score += 12;
  if (String((attachment as any)?.storageBasePath || "").trim()) score += 8;
  return score;
}

function scoreStudioAttachmentCollection(attachments: any[]): number {
  return (Array.isArray(attachments) ? attachments : []).reduce((total, attachment) => total + scoreStudioAttachment(attachment), 0);
}

function mergeClassificationMetaDrafts(
  current?: ClassificationMetaDraft | null,
  incoming?: ClassificationMetaDraft | null
): ClassificationMetaDraft | undefined {
  if (!current && !incoming) return undefined;
  const fallback = normalizeClassificationMetaDraft(current);
  const preferred = normalizeClassificationMetaDraft(incoming);
  return {
    ...fallback,
    ...preferred,
    categorizedLabelNames: mergeUniqueStrings([
      ...(fallback.categorizedLabelNames || []),
      ...(preferred.categorizedLabelNames || []),
    ]),
  };
}

function scoreRelatedEmailEntry(email: RelatedEmailEntry | null | undefined): number {
  if (!email) return 0;
  return Number(Boolean(String(email.emailKey || email.id || email.itemId || email.internetMessageId || "").trim())) * 40
    + Number(Boolean(String(email.bodyText || "").trim() || String(email.bodyHtml || "").trim())) * 60
    + Number(Boolean(String(email.status || "").trim())) * 8
    + scoreStudioAttachmentCollection(Array.isArray(email.attachments) ? email.attachments : [])
    + ((Array.isArray(email.labels) ? email.labels.length : 0) * 2)
    + ((Array.isArray(email.relatedGroups) ? email.relatedGroups.length : 0) * 2)
    + ((Array.isArray(email.relatedRecords) ? email.relatedRecords.length : 0) * 2);
}

function mergeRelatedEmailEntries(current: RelatedEmailEntry, incoming: RelatedEmailEntry): RelatedEmailEntry {
  const preferred = scoreRelatedEmailEntry(incoming) >= scoreRelatedEmailEntry(current) ? incoming : current;
  const fallback = preferred === incoming ? current : incoming;
  const preferredAttachments = Array.isArray(preferred.attachments) ? preferred.attachments : [];
  const fallbackAttachments = Array.isArray(fallback.attachments) ? fallback.attachments : [];
  return {
    ...fallback,
    ...preferred,
    emailKey: String(preferred.emailKey || fallback.emailKey || "").trim() || undefined,
    id: String(preferred.id || fallback.id || "").trim() || undefined,
    itemId: String(preferred.itemId || fallback.itemId || "").trim() || undefined,
    internetMessageId: String(preferred.internetMessageId || fallback.internetMessageId || "").trim() || undefined,
    conversationId: String(preferred.conversationId || fallback.conversationId || "").trim() || undefined,
    subject: String(preferred.subject || fallback.subject || "").trim(),
    fromEmail: String(preferred.fromEmail || fallback.fromEmail || "").trim(),
    fromName: String(preferred.fromName || fallback.fromName || "").trim(),
    receivedAtIso: String(preferred.receivedAtIso || fallback.receivedAtIso || preferred.messageDateIso || fallback.messageDateIso || "").trim() || undefined,
    messageDateIso: String(preferred.messageDateIso || fallback.messageDateIso || preferred.receivedAtIso || fallback.receivedAtIso || "").trim() || undefined,
    bodyText: String(preferred.bodyText || fallback.bodyText || "").trim(),
    bodyHtml: String(preferred.bodyHtml || fallback.bodyHtml || "").trim(),
    status: String(preferred.status || fallback.status || "").trim() || undefined,
    labels: mergeUniqueStrings([...(fallback.labels || []), ...(preferred.labels || [])]),
    removedInheritedLabels: mergeUniqueStrings([...(fallback.removedInheritedLabels || []), ...(preferred.removedInheritedLabels || [])]),
    labelStates: {
      ...(fallback.labelStates || {}),
      ...(preferred.labelStates || {}),
    },
    classificationMeta: mergeClassificationMetaDrafts(fallback.classificationMeta, preferred.classificationMeta),
    attachments: scoreStudioAttachmentCollection(preferredAttachments) >= scoreStudioAttachmentCollection(fallbackAttachments)
      ? preferredAttachments
      : fallbackAttachments,
    relatedGroups: mergeUniqueBy(
      [
        ...(preferred.relatedGroups || []),
        ...(fallback.relatedGroups || []),
      ],
      (group) => String(group?.id || "").trim()
    ),
    relatedRecords: mergeUniqueBy(
      [
        ...(preferred.relatedRecords || []),
        ...(fallback.relatedRecords || []),
      ],
      (record) => `${String(record?.model || "").trim()}:${String(record?.recordId || "").trim()}`
    ),
    relatedReasons: mergeUniqueBy(
      [
        ...(preferred.relatedReasons || []),
        ...(fallback.relatedReasons || []),
      ],
      (reason) => JSON.stringify(reason || {})
    ),
  };
}

function dedupeEmails(emails: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const seen = new Map<string, RelatedEmailEntry>();
  for (const email of emails || []) {
    const key = makeEmailKey(email);
    if (!key) continue;
    const current = seen.get(key);
    seen.set(key, current ? mergeRelatedEmailEntries(current, email) : email);
  }
  return Array.from(seen.values());
}

function htmlToPlainText(html: string): string {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;|&#039;/gi, "'")
    .replace(/&quot;/gi, "\"")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function escapeHtml(value: string): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function sanitizeEmailPreviewHtml(html: string): string {
  const raw = String(html || "")
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<\?xml[\s\S]*?\?>/gi, " ")
    .replace(/<\/?(xml|o:[^>\s]+|v:[^>\s]+)[^>]*>/gi, " ")
    .trim();
  if (!raw) return "";

  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(raw, "text/html");
    doc.querySelectorAll("script, noscript, iframe, object, embed, form, link[rel='stylesheet'], meta[http-equiv], base, svg").forEach((node) => node.remove());
    doc.querySelectorAll<HTMLElement>("*").forEach((element) => {
      Array.from(element.attributes).forEach((attribute) => {
        const name = String(attribute.name || "").toLowerCase();
        const value = String(attribute.value || "").trim();
        if (!name) return;
        if (name.startsWith("on")) {
          element.removeAttribute(attribute.name);
          return;
        }
        if (name === "style" && /url\s*\(/i.test(value)) {
          element.removeAttribute(attribute.name);
          return;
        }
        if (!["src", "href", "poster", "background", "data"].includes(name)) return;
        if (/^(cid|javascript|vbscript|file|ms-appx|about):/i.test(value)) {
          if (element.tagName === "IMG") {
            const fallbackLabel = element.getAttribute("alt") || element.getAttribute("title") || "Imagem inline indisponivel neste preview.";
            element.setAttribute("alt", fallbackLabel);
          }
          element.removeAttribute(attribute.name);
        }
      });
    });
    return String(doc.body?.innerHTML || "")
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .trim();
  } catch {
    return raw
      .replace(/<!--[\s\S]*?-->/g, " ")
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
      .replace(/<svg[\s\S]*?<\/svg>/gi, "")
      .replace(/\s(on\w+)=(".*?"|'.*?'|[^\s>]+)/gi, "")
      .replace(/\s(style)=(".*?url\s*\(.*?\).*?"|'.*?url\s*\(.*?\).*?'|[^\s>]+)/gi, "")
      .replace(/\s(src|href|poster|background|data)=("cid:[^"]*"|'cid:[^']*'|cid:[^\s>]+)/gi, "");
  }
}

function buildEmailPreviewHtml(email: RelatedEmailEntry | null): string {
  const html = String(email?.bodyHtml || "").trim();
  if (html) {
    const sanitizedHtml = sanitizeEmailPreviewHtml(html);
    if (sanitizedHtml) {
      return `<div style="padding:18px;color:#172b4d;font:14px/1.5 'Segoe UI',sans-serif;word-break:break-word">${sanitizedHtml}</div>`;
    }
  }
  const text = String(email?.bodyText || "").trim();
  if (!text) return "";
  return `<pre style="margin:0;padding:18px;color:#172b4d;background:#fff;font:14px/1.55 'Segoe UI',sans-serif;white-space:pre-wrap;word-break:break-word">${escapeHtml(text)}</pre>`;
}

function decodeBase64Text(content: string): string {
  try {
    const binary = globalThis.atob(String(content || "").trim());
    const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
    return new TextDecoder("utf-8").decode(bytes);
  } catch {
    return "";
  }
}

function stripDataUrlPrefix(value: string): string {
  const raw = String(value || "").trim();
  const separatorIndex = raw.indexOf(",");
  if (raw.startsWith("data:") && separatorIndex >= 0) return raw.slice(separatorIndex + 1);
  return raw;
}

function normalizeStudioAttachmentMimeType(value: string | undefined, name: string | undefined): string {
  const raw = String(value || "").trim().toLowerCase();
  const fileName = String(name || "").trim().toLowerCase();
  if (raw === "application/x-pdf" || (!raw && /\.pdf$/.test(fileName))) return "application/pdf";
  if (raw === "image/jpg") return "image/jpeg";
  return raw || "application/octet-stream";
}

function canUseOfficeWebViewer(): boolean {
  try {
    const url = new URL(window.location.origin);
    const hostname = String(url.hostname || "").trim().toLowerCase();
    return Boolean(
      /^https?:$/i.test(url.protocol)
      && hostname
      && hostname !== "localhost"
      && hostname !== "127.0.0.1"
    );
  } catch {
    return false;
  }
}

function buildOfficePreviewUrl(sourceUrl: string): string {
  const normalizedSourceUrl = String(sourceUrl || "").trim();
  if (!normalizedSourceUrl || !canUseOfficeWebViewer()) return "";
  return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(normalizedSourceUrl)}`;
}

function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = stripDataUrlPrefix(dataUrl);
  const binary = globalThis.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function StudioPdfPreview({ dataUrl, title }: { dataUrl: string; title: string }) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");

  useEffect(() => {
    let cancelled = false;
    const host = hostRef.current;
    if (!host || !dataUrl) {
      setStatus("error");
      return;
    }

    host.innerHTML = "";
    setStatus("loading");

    (async () => {
      try {
        const loadingTask = pdfjsLib.getDocument({ data: dataUrlToUint8Array(dataUrl) });
        const pdf = await loadingTask.promise;
        if (cancelled) {
          void loadingTask.destroy();
          return;
        }

        const nextPageCount = Number(pdf.numPages || 0);

        for (let pageNumber = 1; pageNumber <= nextPageCount; pageNumber += 1) {
          if (cancelled) break;
          const page = await pdf.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1.15 });
          const canvas = document.createElement("canvas");
          canvas.style.display = "block";
          canvas.style.width = "100%";
          canvas.style.maxWidth = `${Math.ceil(viewport.width)}px`;
          canvas.style.height = "auto";
          canvas.style.margin = pageNumber === nextPageCount ? "0 auto" : "0 auto 12px auto";
          canvas.style.background = "#fff";
          canvas.style.borderRadius = "8px";
          canvas.style.boxShadow = "0 6px 16px rgba(15,23,42,0.08)";
          const context = canvas.getContext("2d", { alpha: false });
          if (!context) continue;
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          host.appendChild(canvas);
          await page.render({ canvasContext: context, viewport }).promise;
        }

        if (!cancelled) setStatus("ready");
      } catch (error) {
        console.warn("[classification-studio] pdf preview failed", error);
        if (!cancelled) setStatus("error");
      }
    })();

    return () => {
      cancelled = true;
      if (hostRef.current) hostRef.current.innerHTML = "";
    };
  }, [dataUrl]);

  if (status === "error") {
    return <div style={S.attachmentPreviewEmpty}>Este PDF foi detetado, mas nao foi possivel renderiza-lo dentro do add-in.</div>;
  }

  return (
    <div style={S.attachmentPdfPreviewShell} aria-label={title}>
      {status === "loading" ? (
        <div style={S.attachmentPdfPreviewLoading}>A carregar PDF...</div>
      ) : null}
      <div
        ref={hostRef}
        style={{
          ...S.attachmentPdfPreviewCanvasHost,
          display: status === "loading" ? "none" : S.attachmentPdfPreviewCanvasHost.display,
        }}
      />
    </div>
  );
}

function buildSnippet(email: RelatedEmailEntry): string {
  const source = String(email.bodyText || "").trim() || htmlToPlainText(String(email.bodyHtml || ""));
  return source.length > 180 ? `${source.slice(0, 177).trim()}...` : source;
}

function buildEmailCorpus(email: RelatedEmailEntry): string {
  return [
    email.subject,
    email.fromName,
    email.fromEmail,
    email.bodyText,
    htmlToPlainText(String(email.bodyHtml || "")),
    ...(email.attachments || []).map((attachment) => attachment.name),
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" ");
}

function matchReferenceSet(text: string, references: string[]): string[] {
  const compactHaystack = String(text || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
  return references.filter((reference) => {
    const compactReference = compactReferenceValue(reference);
    return Boolean(compactReference && compactHaystack.includes(compactReference));
  });
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

function isExternalEmail(email: RelatedEmailEntry): boolean {
  const from = String(email.fromEmail || "").toLowerCase();
  return from ? !from.endsWith("@divitek.pt") : true;
}

function isCurrentContextEmail(email: Partial<RelatedEmailEntry>, currentContext: Partial<StudioParams>) {
  const emailItemId = String(email?.itemId || "").trim();
  const contextItemId = String(currentContext?.itemId || "").trim();
  if (emailItemId && contextItemId && emailItemId === contextItemId) return true;
  const emailMessageId = String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  const contextMessageId = String(currentContext?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  return Boolean(emailMessageId && contextMessageId && emailMessageId === contextMessageId);
}

function makeAttachmentKey(attachment: { key?: string; id?: string; name?: string; contentId?: string }): string {
  return String(attachment.key || attachment.id || attachment.contentId || attachment.name || "").trim();
}

function normalizeStudioAttachment(attachment: any) {
  if (!attachment || typeof attachment !== "object") return null;
  return {
    key: String(attachment.key || "").trim() || undefined,
    id: String(attachment.id || "").trim() || undefined,
    name: String(attachment.name || "").trim(),
    contentType: String(attachment.contentType || "application/octet-stream"),
    content: String(attachment.content || ""),
    size: attachment.size,
    isInline: attachment.isInline,
    contentId: String(attachment.contentId || "").trim() || undefined,
    documentState: normalizeDocumentLifecycleState(attachment.documentState, "ingested"),
    storageProvider: String(attachment.storageProvider || "").trim() || undefined,
    storageBasePath: String(attachment.storageBasePath || "").trim() || undefined,
    storagePathHint: String(attachment.storagePathHint || "").trim() || undefined,
    hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
    isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : undefined,
  };
}

function getStudioAttachmentRemoteId(attachment: any): string {
  const normalized = normalizeStudioAttachment(attachment);
  if (!normalized) return "";
  return String(normalized.key || normalized.id || normalized.contentId || normalized.name || "").trim();
}

function isStudioAttachmentHydrated(attachment: any): boolean {
  const normalized = normalizeStudioAttachment(attachment);
  if (!normalized?.name) return false;
  if (String(normalized.content || "").trim()) return true;
  if (normalized.hasContent !== true) return false;
  return Boolean(getStudioAttachmentRemoteId(normalized));
}

function hasHydratedAttachmentCollection(email: RelatedEmailEntry | null): boolean {
  const attachments = Array.isArray(email?.attachments) ? email.attachments : [];
  if (!attachments.length) return false;
  return attachments.every((attachment) => isStudioAttachmentHydrated(attachment));
}

function buildRelevantEmailPayloadFromRelatedEmail(email: RelatedEmailEntry | null): RelevantEmailPayload | null {
  if (!email) return null;
  const itemId = String(email.itemId || "").trim();
  const internetMessageId = String(email.internetMessageId || "").trim();
  const conversationId = String(email.conversationId || "").trim();
  const subject = String(email.subject || "").trim();
  const fromEmail = String(email.fromEmail || "").trim();
  if (!(itemId || internetMessageId || conversationId || subject || fromEmail)) return null;
  const attachments = Array.isArray(email.attachments)
    ? email.attachments
        .map((attachment) => normalizeStudioAttachment(attachment))
        .filter((attachment): attachment is NonNullable<ReturnType<typeof normalizeStudioAttachment>> => Boolean(attachment))
        .map((attachment) => ({
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
        }))
    : [];
  return {
    itemId: itemId || undefined,
    internetMessageId: internetMessageId || undefined,
    conversationId: conversationId || undefined,
    subject: subject || undefined,
    fromEmail: fromEmail || undefined,
    fromName: String(email.fromName || "").trim() || undefined,
    receivedAtIso: String(email.receivedAtIso || email.messageDateIso || "").trim() || undefined,
    messageDateIso: String(email.messageDateIso || email.receivedAtIso || "").trim() || undefined,
    bodyText: String(email.bodyText || "").trim() || undefined,
    bodyHtml: String(email.bodyHtml || "").trim() || undefined,
    ...(attachments.length ? { attachments } : {}),
  };
}

function buildAttachmentStorageOptions(settings?: any): Pick<RelevantEmailPayload, "attachmentStorageProvider" | "attachmentStorageBasePath"> {
  return {
    attachmentStorageProvider: settings?.groupStorage?.provider || "cloud",
    attachmentStorageBasePath: settings?.groupStorage?.baseFolderPath || "",
  };
}

async function persistRelatedEmailsToServer(emails: RelatedEmailEntry[], settings?: any): Promise<void> {
  const storageOptions = buildAttachmentStorageOptions(settings);
  const payloads = dedupeEmails(emails)
    .map((email) => buildRelevantEmailPayloadFromRelatedEmail(email))
    .filter(Boolean) as RelevantEmailPayload[];
  if (!payloads.length) return;
  await Promise.allSettled(
    payloads.map((payload) => registerRelevantEmail({
      ...payload,
      ...storageOptions,
    }))
  );
}

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

function detectCaseType(text: string): string {
  const value = text.toLowerCase();
  if (/(reclam|inciden|nao conforme|defeito)/.test(value)) return "reclamacao";
  if (/(pedido|encomenda|order|po\b|purchase order|material listo)/.test(value)) return "pedido/encomenda";
  if (/(proposta|orcamento|quote|quotation)/.test(value)) return "proposta";
  if (/(projeto|project|obra|worksite)/.test(value)) return "projeto";
  return "geral";
}

function inferCompanyName(fromName: string | undefined, fromEmail: string | undefined): string {
  const rawName = String(fromName || "").trim();
  if (rawName.includes("|")) {
    const parts = rawName.split("|").map((part) => part.trim()).filter(Boolean);
    if (parts.length >= 2) return parts[parts.length - 1];
  }
  const email = String(fromEmail || "").trim().toLowerCase();
  if (!email.includes("@")) return "";
  const domain = email.split("@")[1] || "";
  const base = domain.split(".")[0] || "";
  if (!base) return "";
  return base
    .split(/[-_]+/g)
    .filter(Boolean)
    .map((chunk) => chunk.charAt(0).toUpperCase() + chunk.slice(1))
    .join(" ");
}

function normalizeGroupContactDraft(value: Partial<GroupContactDraft> | null | undefined): GroupContactDraft | null {
  const name = String(value?.name || "").trim();
  const email = String(value?.email || "").trim().toLowerCase();
  const company = String(value?.company || "").trim();
  const source = String(value?.source || "").trim() || "email";
  const key = String(value?.key || email || `${normalizeSearchValue(name)}|${normalizeSearchValue(company)}`).trim();
  if (!key || (!name && !email)) return null;
  return {
    key,
    name: name || email,
    email: email || undefined,
    company: company || undefined,
    source,
  };
}

function normalizeGroupEntityDraft(value: Partial<GroupEntityDraft> | null | undefined): GroupEntityDraft | null {
  const name = String(value?.name || "").trim();
  const kind = String(value?.kind || "").trim() || "empresa";
  const source = String(value?.source || "").trim() || "email";
  const key = String(value?.key || normalizeSearchValue(name)).trim();
  if (!key || !name) return null;
  return {
    key,
    name,
    kind,
    source,
  };
}

function dedupeGroupContacts(rows: Array<Partial<GroupContactDraft> | null | undefined>): GroupContactDraft[] {
  const seen = new Set<string>();
  const out: GroupContactDraft[] = [];
  rows.forEach((row) => {
    const normalized = normalizeGroupContactDraft(row);
    if (!normalized || seen.has(normalized.key)) return;
    seen.add(normalized.key);
    out.push(normalized);
  });
  return out;
}

function dedupeGroupEntities(rows: Array<Partial<GroupEntityDraft> | null | undefined>): GroupEntityDraft[] {
  const seen = new Set<string>();
  const out: GroupEntityDraft[] = [];
  rows.forEach((row) => {
    const normalized = normalizeGroupEntityDraft(row);
    if (!normalized || seen.has(normalized.key)) return;
    seen.add(normalized.key);
    out.push(normalized);
  });
  return out;
}

function detectReferences(text: string): string[] {
  const refs = new Set<string>();
  const patterns = [
    /\b(?:pedido|encomenda|order|po|proposta|orcamento|obra|projeto|project)\s*(?:n[.oº°]*)?\s*([A-Z]{0,6}[-/]?\d{2,}[A-Z0-9/-]*)/gi,
    /\b([A-Z]{2,6}[-/]\d{2,})\b/g,
    /\b(\d{3,}[A-Z0-9/-]{0,10})\b/g,
  ];
  for (const pattern of patterns) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(text))) {
      const value = String(match[1] || "").trim();
      if (value && value.length >= 4) refs.add(value);
    }
  }
  return Array.from(refs).slice(0, 6);
}

function splitSuggestions(allGroups: LinkGroupEntry[], text: string): LinkGroupEntry[] {
  const value = text.toLowerCase();
  return allGroups.filter((group) => {
    if (String(group?.kind || "").trim().toLowerCase() === "conversation") return false;
    const name = String(group.name || "").trim().toLowerCase();
    if (!name || name.length < 4) return false;
    return value.includes(name);
  }).slice(0, 8);
}

function normalizeSearchValue(value: string): string {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

function normalizeReferenceCandidate(value: string): string {
  return String(value || "")
    .replace(/[‐‑–—]/g, "-")
    .replace(/\s*([/-])\s*/g, "$1")
    .replace(/^[^A-Z0-9]+|[^A-Z0-9]+$/gi, "")
    .replace(/[.,;:)\]]+$/g, "")
    .trim()
    .toUpperCase();
}

function compactReferenceValue(value: string): string {
  return normalizeReferenceCandidate(value).replace(/[\s/-]+/g, "");
}

function detectReferencesFocused(text: string): string[] {
  const prepared = String(text || "")
    .replace(/[‐‑–—]/g, "-")
    .replace(/([A-Z0-9])\s*([/-])\s*(?=[A-Z0-9])/gi, "$1$2");
  const rawMatches: string[] = [];
  const patterns = [
    /\b(?:pedido|encomenda|order|po|purchase order|proposta|orcamento|obra|projeto|project|ref(?:erencia)?|doc(?:umento)?|fatura|invoice)\s*(?:n(?:o|º|°)?\.?\s*)?([A-Z0-9]+(?:[/-][A-Z0-9]+){1,4})\b/gi,
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
    if (/(PEDIDO|ENCOMENDA|ORDER|PROPOSTA|ORCAMENTO|ORÇAMENTO|FATURA|INVOICE|GUIA|OBRA|PROJETO|PROJECT|DOC|DOCUMENTO|REF)/.test(context)) {
      classification = "documents";
    } else if (/(ARTIGO|ITEM|CODIGO|CÓDIGO|COD |MODELO|SERIE|SÉRIE|PRODUTO|ACABAMENTO|COR |COLOR|TAMANHO|MEDIDA|DIMENSAO|DIMENSÃO)/.test(context)) {
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

function readParams(): StudioParams {
  const params = new URLSearchParams(window.location.search);
  return {
    conversationId: String(params.get("conversationId") || "").trim() || undefined,
    internetMessageId: String(params.get("internetMessageId") || "").trim() || undefined,
    itemId: String(params.get("itemId") || "").trim() || undefined,
    subject: String(params.get("subject") || "").trim() || undefined,
    fromEmail: String(params.get("fromEmail") || "").trim() || undefined,
    fromName: String(params.get("fromName") || "").trim() || undefined,
    receivedAtIso: String(params.get("receivedAtIso") || "").trim() || undefined,
    seedKey: String(params.get("seedKey") || "").trim() || undefined,
  };
}

function readSeedEmail(params: StudioParams): RelatedEmailEntry | null {
  const key = String(params.seedKey || "").trim();
  if (!key || !key.startsWith(GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX)) return null;
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed: any = JSON.parse(raw);
    const itemId = String(parsed?.itemId || "").trim();
    const internetMessageId = String(parsed?.internetMessageId || "").trim();
    const conversationId = String(parsed?.conversationId || "").trim();
    const subject = String(parsed?.subject || "").trim();
    const fromEmail = String(parsed?.fromEmail || "").trim();
    const fromName = String(parsed?.fromName || "").trim();
    const receivedAtIso = String(parsed?.receivedAtIso || parsed?.messageDateIso || "").trim();
    if (!(itemId || internetMessageId || conversationId || subject || fromEmail)) return null;
    return {
      emailKey: itemId || internetMessageId || `${conversationId}|${subject || fromEmail}`,
      itemId: itemId || undefined,
      internetMessageId: internetMessageId || undefined,
      conversationId: conversationId || undefined,
      subject: subject || "(sem assunto)",
      fromEmail: fromEmail || undefined,
      fromName: fromName || undefined,
      receivedAtIso: receivedAtIso || undefined,
      messageDateIso: receivedAtIso || undefined,
      bodyText: String(parsed?.bodyText || "").trim(),
      bodyHtml: String(parsed?.bodyHtml || "").trim(),
      attachments: Array.isArray(parsed?.attachments)
        ? parsed.attachments
          .map((attachment: any) => ({
            id: String(attachment?.id || "").trim() || undefined,
            name: String(attachment?.name || "").trim(),
            contentType: String(attachment?.contentType || "application/octet-stream").trim(),
            size: Number(attachment?.size || 0) || undefined,
            isInline: Boolean(attachment?.isInline),
            contentId: String(attachment?.contentId || "").trim() || undefined,
            content: String(attachment?.content || "").trim(),
          }))
          .filter((attachment: any) => attachment.name)
        : [],
      relatedGroups: [],
      relatedReasons: [],
    };
  } catch {
    return null;
  }
}

function buildFallbackEmail(params: StudioParams): RelatedEmailEntry | null {
  const itemId = String(params.itemId || "").trim();
  const internetMessageId = String(params.internetMessageId || "").trim();
  const conversationId = String(params.conversationId || "").trim();
  const subject = String(params.subject || "").trim();
  const fromEmail = String(params.fromEmail || "").trim();
  const fromName = String(params.fromName || "").trim();
  const receivedAtIso = String(params.receivedAtIso || "").trim();
  if (!(itemId || internetMessageId || conversationId || subject || fromEmail)) return null;
  return {
    emailKey: itemId || internetMessageId || `${conversationId}|${subject || fromEmail}`,
    itemId: itemId || undefined,
    internetMessageId: internetMessageId || undefined,
    conversationId: conversationId || undefined,
    subject: subject || "(sem assunto)",
    fromEmail: fromEmail || undefined,
    fromName: fromName || undefined,
    receivedAtIso: receivedAtIso || undefined,
    messageDateIso: receivedAtIso || undefined,
    bodyText: "",
    bodyHtml: "",
    attachments: [],
    relatedGroups: [],
    relatedReasons: [],
  };
}

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
  const [applyDialogExpandedEmailKeys, setApplyDialogExpandedEmailKeys] = useState<string[]>([]);
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

  const currentSeed = useMemo(() => readSeedEmail(params), [params]);
  const fallbackIdentity = useMemo(() => buildFallbackEmail(params), [params]);
  const currentContext = useMemo(() => ({
    conversationId: String(params.conversationId || currentSeed?.conversationId || fallbackIdentity?.conversationId || "").trim(),
    internetMessageId: String(params.internetMessageId || currentSeed?.internetMessageId || fallbackIdentity?.internetMessageId || "").trim(),
    itemId: String(params.itemId || currentSeed?.itemId || fallbackIdentity?.itemId || "").trim(),
    subject: String(params.subject || currentSeed?.subject || fallbackIdentity?.subject || "").trim(),
    fromEmail: String(params.fromEmail || currentSeed?.fromEmail || fallbackIdentity?.fromEmail || "").trim(),
    fromName: String(params.fromName || currentSeed?.fromName || fallbackIdentity?.fromName || "").trim(),
    receivedAtIso: String(
      params.receivedAtIso ||
      currentSeed?.receivedAtIso ||
      currentSeed?.messageDateIso ||
      fallbackIdentity?.receivedAtIso ||
      fallbackIdentity?.messageDateIso ||
      ""
    ).trim(),
  }), [currentSeed, fallbackIdentity, params]);
  const bootstrapEmailPayload = useMemo<RelevantEmailPayload | null>(() => {
    const base = currentSeed || fallbackIdentity;
    if (!base) return null;
    return buildRelevantEmailPayloadFromRelatedEmail({
      ...base,
      itemId: String(currentContext.itemId || base.itemId || "").trim() || undefined,
      internetMessageId: String(currentContext.internetMessageId || base.internetMessageId || "").trim() || undefined,
      conversationId: String(currentContext.conversationId || base.conversationId || "").trim() || undefined,
      subject: String(currentContext.subject || base.subject || "").trim() || undefined,
      fromEmail: String(currentContext.fromEmail || base.fromEmail || "").trim() || undefined,
      fromName: String(currentContext.fromName || base.fromName || "").trim() || undefined,
      receivedAtIso: String(currentContext.receivedAtIso || base.receivedAtIso || base.messageDateIso || "").trim() || undefined,
      messageDateIso: String(base.messageDateIso || currentContext.receivedAtIso || base.receivedAtIso || "").trim() || undefined,
    });
  }, [
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
    void (async () => {
      try {
        const settings = await getSettings();
        applySkin(settings.skinId || "soft");
        setLabelCatalogEntries(normalizeGroupLabelCatalog(settings.groupLabelCatalog || []));
        setFavoriteGroupIds(Array.isArray((settings as any)?.groupFavoriteIds)
          ? Array.from(new Set((settings as any).groupFavoriteIds.map((entry: any) => String(entry || "").trim()).filter(Boolean)))
          : []);
      } catch {
        applySkin("soft");
        setLabelCatalogEntries([]);
        setFavoriteGroupIds([]);
      } finally {
        setLabelCatalogReady(true);
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoading(true);
      setError("");
      try {
        const latestSettings = await getSettings().catch(() => null);
        if (bootstrapEmailPayload) {
          await registerRelevantEmail({
            ...bootstrapEmailPayload,
            attachmentStorageProvider: latestSettings?.groupStorage?.provider || "cloud",
            attachmentStorageBasePath: latestSettings?.groupStorage?.baseFolderPath || "",
          }).catch(() => null);
        }
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
        const contextualEmails = dedupeEmails([
          ...(related.email ? [related.email] : []),
          ...(related.emails || []),
        ]);
        await persistRelatedEmailsToServer(contextualEmails, latestSettings);
        const mergedEmails = dedupeEmails([...contextualEmails, ...(emails || [])]);
        setAllGroups(mergedGroups);
        setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
        setTicketSeries(Array.isArray(series) ? series : []);
        setRelatedTickets(Array.isArray(related.tickets) ? related.tickets : []);
        setRelatedEmails(contextualEmails);
        setKnownEmails(mergedEmails);
        setSelectedEmailKey((current) => {
          if (current && mergedEmails.some((email) => makeEmailKey(email) === current)) return current;
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
        if (mergedEmails.length) {
          setStatus("Janela base pronta. O email atual e os relacionados persistidos ja podem ser analisados aqui.");
        } else if (bootstrapEmailPayload) {
          setStatus("O email atual foi enviado para o servidor, mas ainda nao existem relacionados persistidos para mostrar.");
        } else {
          setStatus("Ainda nao encontrámos um email persistido para este caso.");
        }
      } catch (fetchError: any) {
        if (!cancelled) setError(String(fetchError?.message || fetchError || "Falha a preparar o studio de classificacao."));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [bootstrapEmailPayload, currentContext.conversationId, currentContext.fromEmail, currentContext.fromName, currentContext.internetMessageId, currentContext.itemId, currentContext.receivedAtIso, currentContext.subject]);

  const groupMap = useMemo(() => new Map(allGroups.map((group) => [group.id, group])), [allGroups]);
  const businessGroups = useMemo(
    () => allGroups.filter((group) => String(group?.kind || "").trim().toLowerCase() !== "conversation"),
    [allGroups]
  );
  const currentCaseBusinessGroups = useMemo(
    () => currentCaseGroups.filter((group) => String(group?.kind || "").trim().toLowerCase() !== "conversation"),
    [currentCaseGroups]
  );
  const emailPool = useMemo(() => (scopeMode === "related" ? dedupeEmails(relatedEmails) : dedupeEmails([...relatedEmails, ...knownEmails])), [knownEmails, relatedEmails, scopeMode]);
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
      map.set(key, { groupIds, labels, ticketIds });
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
    () => visibleEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || emailPool.find((email) => makeEmailKey(email) === selectedEmailKey) || visibleEmails[0] || emailPool[0] || null,
    [emailPool, selectedEmailKey, visibleEmails]
  );
  const selectedEmailInRelatedContext = useMemo(
    () => Boolean(selectedEmail && relatedEmails.some((email) => makeEmailKey(email) === makeEmailKey(selectedEmail))),
    [relatedEmails, selectedEmail]
  );

  const selectedEmailIsCurrent = useMemo(() => {
    return isCurrentContextEmail(selectedEmail || {}, currentContext);
  }, [currentContext, selectedEmail]);

  function getEmailGroupRelations(email: RelatedEmailEntry | null) {
    if (!email) return [];
    const fallbackCurrentGroups = isCurrentContextEmail(email, currentContext)
      ? currentCaseBusinessGroups.map((group) => ({
          id: group.id,
          name: group.name,
          relationKind: group.relationKind,
          kind: group.kind,
        }))
      : [];
    const list = [
      ...(email.relatedGroups || []),
      ...(email.groupId ? [{ id: email.groupId, name: email.groupName, relationKind: email.membershipKind }] : []),
      ...fallbackCurrentGroups,
    ];
    return list.reduce<Array<{ id: string; name?: string; relationKind?: string }>>((acc, row) => {
      if (!row?.id || acc.some((entry) => entry.id === row.id)) return acc;
      const groupKind = String((row as any)?.kind || groupMap.get(row.id)?.kind || "").trim().toLowerCase();
      if (groupKind === "conversation") return acc;
      acc.push(row);
      return acc;
    }, []);
  }

  const selectedEmailGroups = useMemo(() => {
    return getEmailGroupRelations(selectedEmail);
  }, [selectedEmail, currentCaseBusinessGroups, currentContext, groupMap]);

  const principalAnchorGroupId = useMemo(
    () => principalGroupId || selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal")?.id || "",
    [principalGroupId, selectedEmailGroups]
  );

  const selectedTargetEmails = useMemo(
    () => emailPool.filter((email) => selectedTargetEmailKeys.includes(makeEmailKey(email))),
    [emailPool, selectedTargetEmailKeys]
  );
  const caseScopeEmails = useMemo(
    () => dedupeEmails([...(selectedEmail ? [selectedEmail] : []), ...relatedEmails]),
    [relatedEmails, selectedEmail]
  );

  const principalScopeEmails = useMemo(() => {
    if (!principalAnchorGroupId) return [];
    return emailPool.filter((email) =>
      getEmailGroupRelations(email).some(
        (group) => String(group.relationKind || "").toLowerCase() === "principal" && group.id === principalAnchorGroupId
      )
    );
  }, [emailPool, principalAnchorGroupId, currentCaseBusinessGroups, currentContext, groupMap]);
  const selectedTargetCount = selectedTargetEmails.length;
  const principalScopeCount = principalScopeEmails.length;
  const applyDialogSelectedEmails = useMemo(
    () => caseScopeEmails.filter((email) => applyDialogEmailKeys.includes(makeEmailKey(email))),
    [applyDialogEmailKeys, caseScopeEmails]
  );

  const selectedEmailTicketIds = useMemo(() => {
    if (!selectedEmail) return [];
    const meta = emailContextMeta.get(makeEmailKey(selectedEmail));
    return Array.isArray(meta?.ticketIds) ? meta.ticketIds.filter(Boolean) : [];
  }, [emailContextMeta, selectedEmail]);

  useEffect(() => {
    if (!selectedEmail) return;
    if (!selectionTouched.principal) {
      const principal = selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal");
      setPrincipalGroupId(principal?.id || "");
    }
    if (!selectionTouched.references) {
      setReferenceGroupIds(selectedEmailGroups.filter((group) => String(group.relationKind || "").toLowerCase() !== "principal").map((group) => group.id));
    }
  }, [selectedEmail, selectedEmailGroups, selectionTouched.principal, selectionTouched.references]);

  useEffect(() => {
    if (!principalGroupId) return;
    setReferenceGroupIds((current) => current.filter((groupId) => groupId !== principalGroupId));
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
    setApplyDialogOpen(false);
    setApplyDialogExpandedEmailKeys([]);
    classificationDraftSnapshotRef.current = null;
    if (section === "classification") {
      setClassificationFocus("summary");
      setSection("emails");
    }
  }, [selectedEmailKey]);

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
      if (group.id === principalGroupId) return false;
      if (!q) return true;
      return normalizeSearchValue(String(group.name || "")).includes(q);
    });
    return rows.slice(0, 24);
  }, [businessGroups, principalGroupId, referenceSearch]);
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
  const availableTicketChoices = useMemo(() => {
    const rows = [...relatedTickets, ...ticketSearchResults].reduce<GroupTicketEntry[]>((acc, ticket) => {
      if (!ticket?.id || acc.some((entry) => entry.id === ticket.id)) return acc;
      acc.push(ticket);
      return acc;
    }, []);
    return rows.sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));
  }, [relatedTickets, ticketSearchResults]);

  const selectedEmailAttachments = useMemo(() => {
    return (selectedEmail?.attachments || [])
      .map((attachment) => normalizeStudioAttachment(attachment))
      .filter((attachment): attachment is NonNullable<typeof attachment> => Boolean(attachment))
      .filter((attachment) => String(attachment.name || "").trim());
  }, [selectedEmail?.attachments]);
  const quickDocumentAttachments = useMemo(
    () => selectedEmailAttachments.filter((attachment) => showHiddenQuickDocuments || !isStudioAttachmentHiddenInQuickDocs(attachment)),
    [selectedEmailAttachments, showHiddenQuickDocuments]
  );
  const quickDocumentHiddenCount = useMemo(
    () => selectedEmailAttachments.filter((attachment) => isStudioAttachmentHiddenInQuickDocs(attachment)).length,
    [selectedEmailAttachments]
  );
  const activeSelectedEmailAttachments = useMemo(
    () => selectedEmailAttachments.filter((attachment) => !isRejectedDocumentLifecycleState((attachment as any)?.documentState)),
    [selectedEmailAttachments]
  );

  useEffect(() => {
    setSelectedAttachmentPreviewKey((current) => {
      if (current && selectedEmailAttachments.some((attachment) => makeAttachmentKey(attachment) === current)) return current;
      return "";
    });
  }, [selectedEmailAttachments]);

  const selectedAttachmentPreview = useMemo(
    () => selectedEmailAttachments.find((attachment) => makeAttachmentKey(attachment) === selectedAttachmentPreviewKey) || null,
    [selectedAttachmentPreviewKey, selectedEmailAttachments]
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
    () => String(selectedEmail?.id || selectedEmail?.emailKey || "").trim(),
    [selectedEmail?.emailKey, selectedEmail?.id]
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
    if (quickDocumentAttachments.some((attachment) => makeAttachmentKey(attachment) === selectedAttachmentPreviewKey)) return;
    const nextKey = quickDocumentAttachments[0] ? makeAttachmentKey(quickDocumentAttachments[0]) : "";
    setSelectedAttachmentPreviewKey(nextKey);
    if (!nextKey && previewMode === "document") {
      setPreviewMode("email");
    }
  }, [previewMode, quickDocumentAttachments, selectedAttachmentPreviewKey, showHiddenQuickDocuments]);

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
      for (const attachment of selectedEmailAttachments) {
        const key = makeAttachmentKey(attachment);
        if (!key || next[key]) continue;
        const contentType = String(attachment.contentType || "").toLowerCase();
        const isDocument = /pdf|image|excel|spreadsheet|word|officedocument|text|csv/.test(contentType) || /\.(pdf|png|jpe?g|xlsx?|docx?|csv|txt)$/i.test(String(attachment.name || ""));
        next[key] = {
          analyze: isRejectedDocumentLifecycleState((attachment as any)?.documentState) ? false : isDocument,
          save: false,
          forward: false,
        };
      }
      return next;
    });
  }, [selectedEmailAttachments]);

  useEffect(() => {
    let cancelled = false;
    const extractableFiles = activeSelectedEmailAttachments
      .map((attachment) => ({
        key: makeAttachmentKey(attachment),
        name: String(attachment.name || "").trim(),
        contentType: String(attachment.contentType || "").trim(),
        content: String(attachment.content || "").trim(),
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
  }, [activeSelectedEmailAttachments]);

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
    () => suggestTicketsFocused(availableTicketChoices, detectionText, documentReferences),
    [availableTicketChoices, detectionText, documentReferences]
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
    const emailUniverse = dedupeEmails([...relatedEmails, ...knownEmails]);
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
  }, [detectedCaseType, documentReferences, emailContextMeta, getEmailGroupRelations, knownEmails, relatedEmails, selectedEmail, selectedEmailGroups, selectedEmailTicketIds, contextualTickets]);
  const selectedTicket = useMemo(() => availableTicketChoices.find((ticket) => ticket.id === selectedTicketId) || relatedTickets.find((ticket) => ticket.id === selectedTicketId) || null, [availableTicketChoices, relatedTickets, selectedTicketId]);
  const principalGroup = useMemo(() => (principalGroupId ? groupMap.get(principalGroupId) || null : null), [groupMap, principalGroupId]);
  const favoritePrincipalGroups = useMemo(
    () => favoriteGroupIds
      .map((groupId) => businessGroups.find((group) => group.id === groupId) || null)
      .filter(Boolean) as LinkGroupEntry[],
    [businessGroups, favoriteGroupIds]
  );
  const favoriteReferenceGroups = useMemo(
    () => favoritePrincipalGroups.filter((group) => group.id !== principalGroupId).slice(0, 6),
    [favoritePrincipalGroups, principalGroupId]
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
          group.id !== principalGroupId
          && normalizeSearchValue(String(group.name || "")) === normalizedReferenceSearch
        ) || null
        : null,
    [businessGroups, normalizedReferenceSearch, principalGroupId]
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
    () => referenceGroupIds.map((groupId) => groupMap.get(groupId)).filter(Boolean) as LinkGroupEntry[],
    [groupMap, referenceGroupIds]
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
  const canApplyClassification = useMemo(
    () => Boolean(
      principalGroupId
      || referenceGroupIds.length
      || selectedTicketId
      || selectedSeriesId
      || selectedEmailGroups.length
      || selectedEmailTicketIds.length
      || selectedLabels.length
      || (selectedEmail?.labels || []).length
      || String(selectedEmail?.status || "").trim()
    ),
    [
      principalGroupId,
      referenceGroupIds.length,
      selectedTicketId,
      selectedSeriesId,
      selectedEmailGroups.length,
      selectedEmailTicketIds.length,
      selectedLabels.length,
      selectedEmail?.labels,
      selectedEmail?.status,
    ]
  );
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
      ...relatedEmails,
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
  }, [managedGroupEmails, relatedEmails, selectedEmail, selectedManagedGroup?.contacts]);

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
  const emailOwnedLabels = useMemo(
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
    if (selectedEmailTicketIds.length === 1) {
      setSelectedTicketId(selectedEmailTicketIds[0]);
      return;
    }
    if (!selectedEmailTicketIds.length) {
      setSelectedTicketId((current) => (
        current && availableTicketChoices.some((ticket) => ticket.id === current)
          ? current
          : ""
      ));
    }
  }, [availableTicketChoices, selectedEmailTicketIds, selectionTouched.ticket]);

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
    setSelectionTouched({ principal: false, references: false, ticket: false });
    setSelectedLabels([]);
    setLabelDrafts({});
    setClassificationMetaDraft(normalizeClassificationMetaDraft(selectedEmail?.classificationMeta));
  }, [selectedEmailKey]);

  useEffect(() => {
    if (!labelCatalogReady) return;
    if (selectedLabels.length || (!inheritedLabels.length && !emailOwnedLabels.length && !selectedEmailRemovedInheritedLabels.length)) return;
    const visibleInherited = inheritedLabels.filter((label) => !selectedEmailRemovedInheritedLabels.includes(label));
    const seedLabels = mergeLabels(visibleInherited, emailOwnedLabels);
    setSelectedLabels(seedLabels);
    setLabelDrafts((current) => {
      const next = { ...current };
      for (const label of seedLabels) {
        next[label] = createLabelDraftFromCatalog(
          findGroupLabelCatalogEntry(labelCatalogEntries, label),
          current[label],
          selectedEmailLabelStates[label],
          selectedEmailCategorizedLabelNames.length
            ? selectedEmailCategorizedLabelNames.includes(label)
            : undefined
        );
      }
      return next;
    });
  }, [emailOwnedLabels, inheritedLabels, labelCatalogEntries, labelCatalogReady, selectedEmailCategorizedLabelNames, selectedEmailLabelStates, selectedEmailRemovedInheritedLabels, selectedLabels.length]);

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
          selectedEmailCategorizedLabelNames.length
            ? selectedEmailCategorizedLabelNames.includes(label)
            : undefined
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
            mergeLabels(labelCatalog, selectedEmail?.labels || []),
            selectedEmail?.removedInheritedLabels || []
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
  }, [selectedEmailIsCurrent, selectedEmailKey]);

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
    const latestSettings = await getSettings().catch(() => null);
    await persistRelatedEmailsToServer(contextualEmails, latestSettings);
    setAllGroups((current) => mergeGroupEntryLists(current, related.groups || []));
    setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
    setRelatedTickets((current) => {
      const nextTickets = Array.isArray(related.tickets) ? related.tickets : [];
      const preservedSelectedTicket = selectedTicketId
        ? current.find((ticket) => ticket.id === selectedTicketId) || null
        : null;
      if (preservedSelectedTicket && !nextTickets.some((ticket) => ticket.id === preservedSelectedTicket.id)) {
        return [preservedSelectedTicket, ...nextTickets];
      }
      return nextTickets;
    });
    setRelatedEmails(contextualEmails);
    setKnownEmails((current) => dedupeEmails([...contextualEmails, ...current]));
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
    setSelectionTouched((current) => ({ ...current, references: true }));
    setReferenceGroupIds((current) => current.includes(groupId) ? current.filter((entry) => entry !== groupId) : [...current, groupId]);
  }

  function clearPrincipalSelection() {
    setSelectionTouched((current) => ({ ...current, principal: true }));
    setPrincipalGroupId("");
  }

  function setPrincipalSearchValue(value: string) {
    const nextValue = String(value || "").trim();
    setPrincipalSearch(nextValue);
    setCreateGroupName(nextValue);
  }

  function setReferenceSearchValue(value: string) {
    setReferenceSearch(String(value || "").trim());
  }

  function selectPrincipalGroup(group: LinkGroupEntry | null) {
    if (!group?.id) return;
    setSelectionTouched((current) => ({ ...current, principal: true }));
    setPrincipalGroupId(group.id);
    setPrincipalSearchValue(group.name);
  }

  function toggleFavoritePrincipalGroup(group: LinkGroupEntry) {
    if (!group?.id) return;
    const sameGroup = principalGroupId === group.id;
    if (sameGroup) {
      clearPrincipalSelection();
      if (normalizeSearchValue(principalSearch) === normalizeSearchValue(group.name)) {
        setPrincipalSearchValue("");
      }
      return;
    }
    selectPrincipalGroup(group);
  }

  function toggleFavoriteReferenceGroup(group: LinkGroupEntry) {
    if (!group?.id) return;
    const sameGroup = referenceGroupIds.includes(group.id);
    setReferenceSearchValue(group.name);
    if (sameGroup) {
      toggleReferenceGroup(group.id);
      if (normalizeSearchValue(referenceSearch) === normalizeSearchValue(group.name)) {
        setReferenceSearchValue("");
      }
      return;
    }
    toggleReferenceGroup(group.id);
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
    if (principalGroupId === groupId) {
      clearPrincipalSelection();
      return;
    }
    if (referenceGroupIds.includes(groupId)) {
      toggleReferenceGroup(groupId);
      return;
    }
    if (classificationFocus === "references") {
      setSelectionTouched((current) => ({ ...current, references: true }));
      setReferenceGroupIds((current) => current.includes(groupId) ? current : [...current, groupId]);
      return;
    }
    if (!principalGroupId || classificationFocus === "principal") {
      setSelectionTouched((current) => ({ ...current, principal: true }));
      setPrincipalGroupId(groupId);
      return;
    }
    if (principalGroupId === groupId) {
      return;
    }
    setSelectionTouched((current) => ({ ...current, references: true }));
    setReferenceGroupIds((current) => current.includes(groupId) ? current : [...current, groupId]);
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
    const match = availableTicketChoices.find((ticket) => normalizeSearchValue(String(ticket.code || "")) === normalized);
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
            selectedEmailCategorizedLabelNames.length
              ? selectedEmailCategorizedLabelNames.includes(value)
              : undefined
          ),
        });
    setLabelInput("");
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
        setReferenceGroupIds((current) => current.includes(created.id) ? current : [...current, created.id]);
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
      const groupIds = [principalGroupId, ...referenceGroupIds].filter(Boolean);
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
        membershipKind: principalGroupId ? "principal" : "referencia",
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
    if (!selectedEmail || !selectedAttachmentPreview) {
      setStatus("Escolhe primeiro um anexo para atualizar o estado documental.");
      return;
    }
    const attachmentKey = makeAttachmentKey(selectedAttachmentPreview);
    if (!attachmentKey) {
      setStatus("Nao foi possivel identificar o anexo selecionado.");
      return;
    }
    const updatedEmail = updateAttachmentStateOnEmail(selectedEmail, attachmentKey, nextState);
    if (!updatedEmail) {
      setStatus("Nao foi possivel atualizar o estado documental deste anexo.");
      return;
    }
    setActionBusy(true);
    try {
      setRelatedEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
      setKnownEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
      setAttachmentPlan((current) => ({
        ...current,
        [attachmentKey]: {
          analyze: nextState === "rejected" ? false : (current[attachmentKey]?.analyze ?? false),
          save: nextState === "rejected" ? false : (current[attachmentKey]?.save ?? false),
          forward: current[attachmentKey]?.forward ?? false,
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
    if (!principalGroupId) {
      setStatus("Escolhe primeiro um grupo principal para guardar documentos.");
      return;
    }
    const docs = (
      await Promise.all(
        selectedEmailAttachments
          .filter((attachment) => attachmentPlan[makeAttachmentKey(attachment)]?.save)
          .map(async (attachment) => {
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
      const storageProvider = String(settings?.groupStorage?.provider || "cloud").trim();
      const storageBasePath = String(settings?.groupStorage?.baseFolderPath || "").trim();
      const safeGroupName = String(principalGroup?.name || principalGroupId || "grupo")
        .trim()
        .replace(/[\\/:*?"<>|]+/g, "_");
      await saveGroupDocuments(principalGroupId, {
        documents: docs.map((doc) => ({
          ...doc,
          storageProvider,
          storageBasePath,
          storagePathHint: safeGroupName && doc.name
            ? `${safeGroupName}/${String(doc.name || "").trim().replace(/[\\/:*?"<>|]+/g, "_")}`
            : undefined,
        })),
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

  async function handleSearchTickets() {
    setActionBusy(true);
    try {
      const rows = await searchGroupTickets({
        q: String(ticketSearch || "").trim() || undefined,
        limit: 20,
      });
      setTicketSearchResults(rows);
      setStatus(rows.length ? `${rows.length} ticket(s) encontrados.` : "Nenhum ticket encontrado para estes filtros.");
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel pesquisar tickets.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleApplyClassification(targetEmailsOverride?: RelatedEmailEntry[]) {
    setActionBusy(true);
    let activeCategoryOperationId = "";
    let activeCategoryRequestId = "";
    let categoryOperationClosed = false;
    try {
      const targetEmails = dedupeEmails(
        (targetEmailsOverride && targetEmailsOverride.length
          ? targetEmailsOverride
          : (
            applyScopeMode === "selected"
              ? selectedTargetEmails
              : applyScopeMode === "principal_group"
                ? principalScopeEmails
                : [selectedEmail].filter(Boolean)
          )) as RelatedEmailEntry[]
      );
      const effectiveTargetEmails = targetEmails.length
        ? targetEmails
        : ((selectedEmail ? [selectedEmail] : []) as RelatedEmailEntry[]);
      if (!effectiveTargetEmails.length) {
        setStatus("Nao existe nenhum email alvo para atualizar.");
        return false;
      }
      const includesCurrentTarget = selectedEmailIsCurrent || effectiveTargetEmails.some((email) => isCurrentContextEmail(email, currentContext));
      const currentTargetIdentity = includesCurrentTarget
        ? {
            itemId: String(currentContext.itemId || "").trim() || undefined,
            internetMessageId: String(currentContext.internetMessageId || "").trim() || undefined,
            conversationId: String(currentContext.conversationId || "").trim() || undefined,
          }
        : null;
      if (currentTargetIdentity) {
        const openedOperation = beginOutlookCategoryOperation({
          owner: "classification",
          target: currentTargetIdentity,
        });
        if (!openedOperation.ok) {
          setStatus(
            openedOperation.reason === "locked"
              ? "Ja existe outra classificacao em curso para este email. Aguarda um momento."
              : "Nao foi possivel identificar o email atual para confirmar a classificacao."
          );
          return false;
        }
        activeCategoryOperationId = openedOperation.operation.operationId;
        setOutlookCategoryOperationPhase(activeCategoryOperationId, "saving");
      }

      const principalGroup = principalGroupId ? groupMap.get(principalGroupId) || null : null;
      const referenceGroups = referenceGroupIds.map((groupId) => groupMap.get(groupId)).filter(Boolean) as LinkGroupEntry[];
      const allGroupIds = [principalGroupId, ...referenceGroupIds].filter(Boolean);
      const emailLabelStatus = selectedLabelStatuses[0] || "";
      const removedInheritedLabels = inheritedLabels.filter((label) => !selectedLabels.includes(label));
      const emailOwnedSelectedLabels = selectedLabels.filter((label) => !inheritedLabels.includes(label));
      const latestSettings = await getSettings().catch(() => null);
      const currentOutlookTicket = selectedTicketId
        ? (availableTicketChoices.find((ticket) => ticket.id === selectedTicketId)
          || relatedTickets.find((ticket) => ticket.id === selectedTicketId)
          || null)
        : null;
      const desiredTicketStatus = String(ticketStatusDraft || "").trim();

      let finalTicket: GroupTicketEntry | null = null;
      const buildTargetPayload = (targetEmail: RelatedEmailEntry): RelevantEmailPayload => {
        const targetIsCurrent = isCurrentContextEmail(targetEmail, currentContext);
        const targetAttachments = (targetEmail.attachments || []).map((attachment) => ({
          key: attachment.key,
          id: attachment.id,
          name: attachment.name,
          contentType: String(attachment.contentType || "application/octet-stream"),
          content: String(attachment.content || ""),
          size: attachment.size,
          isInline: attachment.isInline,
          contentId: attachment.contentId,
          storageProvider: (attachment as any).storageProvider,
          storageBasePath: (attachment as any).storageBasePath,
          storagePathHint: (attachment as any).storagePathHint,
          documentState: normalizeDocumentLifecycleState((attachment as any)?.documentState, "ingested"),
          hasContent: (attachment as any)?.hasContent === true || Boolean(String(attachment.content || "").trim()),
          isHidden: typeof (attachment as any)?.isHidden === "boolean" ? (attachment as any).isHidden : undefined,
        }));
        return {
          itemId: String(targetEmail?.itemId || (targetIsCurrent ? currentContext.itemId : "") || "").trim() || undefined,
          internetMessageId: String(targetEmail?.internetMessageId || (targetIsCurrent ? currentContext.internetMessageId : "") || "").trim() || undefined,
          conversationId: String(targetEmail?.conversationId || (targetIsCurrent ? currentContext.conversationId : "") || "").trim() || undefined,
          subject: String(targetEmail?.subject || (targetIsCurrent ? currentContext.subject : "") || "").trim() || undefined,
          fromEmail: String(targetEmail?.fromEmail || (targetIsCurrent ? currentContext.fromEmail : "") || "").trim() || undefined,
          fromName: String(targetEmail?.fromName || (targetIsCurrent ? currentContext.fromName : "") || "").trim() || undefined,
          receivedAtIso: String(targetEmail?.receivedAtIso || targetEmail?.messageDateIso || (targetIsCurrent ? currentContext.receivedAtIso : "") || "").trim() || undefined,
          messageDateIso: String(targetEmail?.messageDateIso || targetEmail?.receivedAtIso || (targetIsCurrent ? currentContext.receivedAtIso : "") || "").trim() || undefined,
          bodyText: String(targetEmail?.bodyText || "").trim() || undefined,
          bodyHtml: String(targetEmail?.bodyHtml || "").trim() || undefined,
          attachments: targetAttachments.map((attachment) => ({
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
            isHidden: typeof (attachment as any)?.isHidden === "boolean" ? (attachment as any).isHidden : undefined,
          })),
        };
      };

      const buildClassifiedEmailPayload = (targetEmail: RelatedEmailEntry): RelevantEmailPayload => ({
        ...buildTargetPayload(targetEmail),
        status: emailLabelStatus,
        labels: emailOwnedSelectedLabels,
        removedInheritedLabels,
        labelStates: selectedLabelStates,
        classificationMeta: {
          ...classificationMetaDraft,
          categorizedLabelNames: categorizableLabels,
        },
      });

      const baseTargetEmail = effectiveTargetEmails[0];
      const baseTargetKey = makeEmailKey(baseTargetEmail);
      if (!selectedTicketId && selectedSeriesId) {
        const baseClassifiedEmailPayload = buildClassifiedEmailPayload(baseTargetEmail);
        finalTicket = await createGroupTicket({
          seriesId: selectedSeriesId,
          title: String(createTicketTitle || baseTargetEmail?.subject || "Ticket").trim(),
          description: String(baseTargetEmail?.bodyText || "").trim().slice(0, 4000),
          labels: selectedLabels,
          groupIds: allGroupIds,
          email: baseClassifiedEmailPayload,
          membershipKind: principalGroupId ? "principal" : "referencia",
        });
        if (desiredTicketStatus && desiredTicketStatus !== String(finalTicket?.status || "").trim()) {
          finalTicket = await updateGroupTicket(finalTicket.id, { status: desiredTicketStatus });
        }
        setRelatedTickets((current) => [finalTicket as GroupTicketEntry, ...current.filter((entry) => entry.id !== finalTicket?.id)]);
        setSelectedTicketId(finalTicket.id);
      }

      if (selectedTicketId && desiredTicketStatus !== String(currentOutlookTicket?.status || "").trim()) {
        finalTicket = await updateGroupTicket(selectedTicketId, { status: desiredTicketStatus });
        setRelatedTickets((current) => [finalTicket as GroupTicketEntry, ...current.filter((entry) => entry.id !== finalTicket?.id)]);
      }

      for (const targetEmail of effectiveTargetEmails) {
        const targetEmailKey = makeEmailKey(targetEmail);
        const targetEmailPayload = buildTargetPayload(targetEmail);
        const classifiedEmailPayload = {
          ...targetEmailPayload,
          status: emailLabelStatus,
          labels: emailOwnedSelectedLabels,
          removedInheritedLabels,
          labelStates: selectedLabelStates,
          classificationMeta: {
            ...classificationMetaDraft,
            categorizedLabelNames: categorizableLabels,
          },
        };
        const targetGroups = getEmailGroupRelations(targetEmail);
        const currentGroupIds = targetGroups.map((group) => String(group.id || "").trim()).filter(Boolean);
        const groupsToRemove = currentGroupIds.filter((groupId) => !allGroupIds.includes(groupId));
        const ticketIdsToRemove = ((emailContextMeta.get(targetEmailKey)?.ticketIds || []) as string[]).filter((ticketId) => ticketId !== selectedTicketId && ticketId !== finalTicket?.id);

        for (const groupId of groupsToRemove) {
          await removeEmailFromLinkGroup(groupId, {
            ...targetEmailPayload,
            emailKey: String(targetEmail?.emailKey || "").trim() || undefined,
          });
        }

        if (principalGroupId) {
          await addEmailToLinkGroup(principalGroupId, {
            ...classifiedEmailPayload,
            membershipKind: "principal",
          });
        }
        for (const groupId of referenceGroupIds) {
          await addEmailToLinkGroup(groupId, {
            ...classifiedEmailPayload,
            membershipKind: "referencia",
          });
        }

        for (const ticketId of ticketIdsToRemove) {
          await unlinkEmailFromGroupTicket(ticketId, {
            email: targetEmailPayload,
            emailKey: String(targetEmail?.emailKey || "").trim() || undefined,
          });
        }

        await registerRelevantEmail({
          ...classifiedEmailPayload,
          attachmentStorageProvider: latestSettings?.groupStorage?.provider || "cloud",
          attachmentStorageBasePath: latestSettings?.groupStorage?.baseFolderPath || "",
        });

        const targetTicketId = finalTicket?.id || selectedTicketId;
        if (targetTicketId && !(finalTicket && targetEmailKey === baseTargetKey)) {
          const linked = await linkEmailToGroupTicket(targetTicketId, {
            email: classifiedEmailPayload,
            applyGroups: allGroupIds.length > 0,
            groupIds: allGroupIds,
            membershipKind: principalGroupId ? "principal" : "referencia",
          });
          finalTicket = linked.ticket;
        }
      }

      let fallbackCurrentCategoryEmail: RelatedEmailEntry | null = null;
      if (currentTargetIdentity) {
        const currentTargetEmail = effectiveTargetEmails.find((email) => isCurrentContextEmail(email, currentContext))
          || (selectedEmailIsCurrent ? selectedEmail : null);
        fallbackCurrentCategoryEmail = currentTargetEmail
          ? {
              ...currentTargetEmail,
              itemId: String(currentContext.itemId || currentTargetEmail.itemId || "").trim() || undefined,
              internetMessageId: String(currentContext.internetMessageId || currentTargetEmail.internetMessageId || "").trim() || undefined,
              conversationId: String(currentContext.conversationId || currentTargetEmail.conversationId || "").trim(),
              subject: String(currentTargetEmail.subject || currentContext.subject || "").trim() || undefined,
              fromEmail: String(currentTargetEmail.fromEmail || currentContext.fromEmail || "").trim() || undefined,
              fromName: String(currentTargetEmail.fromName || currentContext.fromName || "").trim() || undefined,
              receivedAtIso: String(currentTargetEmail.receivedAtIso || currentTargetEmail.messageDateIso || currentContext.receivedAtIso || "").trim() || undefined,
              messageDateIso: String(currentTargetEmail.messageDateIso || currentTargetEmail.receivedAtIso || currentContext.receivedAtIso || "").trim() || undefined,
              status: emailLabelStatus,
              labels: emailOwnedSelectedLabels,
              removedInheritedLabels,
              labelStates: selectedLabelStates,
              classificationMeta: {
                ...classificationMetaDraft,
                categorizedLabelNames: categorizableLabels,
              },
              relatedGroups: [
                ...(principalGroup?.id ? [{
                  id: principalGroup.id,
                  name: principalGroup.name,
                  kind: principalGroup.kind,
                  relationKind: "principal",
                }] : []),
                ...referenceGroups.map((group) => ({
                  id: group.id,
                  name: group.name,
                  kind: group.kind,
                  relationKind: "referencia" as const,
                })),
              ],
            }
          : null;
      }

      setSelectionTouched({ principal: false, references: false, ticket: false });
      if (activeCategoryOperationId) {
        setOutlookCategoryOperationPhase(activeCategoryOperationId, "refreshing");
      }
      const refreshedContext = await refreshSelectedEmailContext();
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
            throw new Error("A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias.");
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
      return true;
    } catch (actionError: any) {
      if (activeCategoryOperationId && !categoryOperationClosed) {
        completeOutlookCategoryOperation(activeCategoryOperationId, {
          result: "failed",
          requestId: activeCategoryRequestId || undefined,
          detail: String(actionError?.message || "").trim() || undefined,
        });
      }
      setStatus(actionError?.message || "Nao foi possivel aplicar a classificacao.");
      return false;
    } finally {
      setActionBusy(false);
    }
  }

  function handleOpenQuickAttachment(attachment: NonNullable<ReturnType<typeof normalizeStudioAttachment>>) {
    const key = makeAttachmentKey(attachment);
    if (!key) return;
    setSelectedAttachmentPreviewKey(key);
    setPreviewMode("document");
  }

  async function handleSetQuickAttachmentHidden(
    attachment: NonNullable<ReturnType<typeof normalizeStudioAttachment>>,
    nextHidden: boolean
  ) {
    if (!selectedEmail || !attachment) return;
    const attachmentKey = makeAttachmentKey(attachment);
    if (!attachmentKey) return;
    const updatedEmail = updateAttachmentVisibilityOnEmail(selectedEmail, attachmentKey, nextHidden);
    if (!updatedEmail) {
      setStatus("Nao foi possivel atualizar a visibilidade deste documento.");
      return;
    }
    setActionBusy(true);
    try {
      setRelatedEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
      setKnownEmails((current) => current.map((email) => makeEmailKey(email) === makeEmailKey(updatedEmail) ? updatedEmail : email));
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
    setPrincipalGroupId(snapshot.principalGroupId);
    setPrincipalSearch(snapshot.principalSearch);
    setReferenceGroupIds([...snapshot.referenceGroupIds]);
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
    setApplyDialogEmailKeys(getDefaultApplyDialogEmailKeys(defaultMode));
  }

  function toggleApplyDialogEmailKey(emailKey: string) {
    if (!emailKey) return;
    setApplyDialogEmailKeys((current) => current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]);
  }

  function toggleApplyDialogExpandedEmailKey(emailKey: string) {
    if (!emailKey) return;
    setApplyDialogExpandedEmailKeys((current) => current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]);
  }

  async function handleConfirmApplyDialog() {
    const selectedEmails = applyDialogSelectedEmails.length
      ? applyDialogSelectedEmails
      : ((selectedEmail ? [selectedEmail] : []) as RelatedEmailEntry[]);
    const applied = await handleApplyClassification(selectedEmails);
    if (!applied) return;
    clearClassificationDraftSession();
  }

  function renderOutlookColorLegend() {
    return (
      <div style={S.legendRow}>
        {OUTLOOK_CATEGORY_COLOR_LEGEND.map((entry) => (
          <span key={entry.key} style={{ ...S.legendChip, ...entry.style }}>{entry.label}</span>
        ))}
      </div>
    );
  }

  function renderSuggestionTray(
    kind: "principal" | "labels",
    title: string,
    chips: Array<{ key: string; label: string; active?: boolean; onClick: () => void }>,
    helper: string
  ) {
    const visible = chips.slice(0, 3);
    const hidden = chips.slice(3);
    const expanded = classificationSuggestionExpanded[kind];
    return (
      <div style={S.editorBlock}>
        <div style={S.editorBlockHeader}>
          <div style={S.editorBlockTitle}>{title}</div>
          {hidden.length ? (
            <button
              type="button"
              style={S.chevronBtn}
              onClick={() => setClassificationSuggestionExpanded((current) => ({ ...current, [kind]: !current[kind] }))}
            >
              {expanded ? "⌃" : "⌄"}
            </button>
          ) : null}
        </div>
        <div style={S.chipGridCompact}>
          {visible.length ? visible.map((chip) => (
            <button key={chip.key} type="button" style={chip.active ? S.miniChipOn : S.miniChip} onClick={chip.onClick}>
              {chip.label}
            </button>
          )) : <span style={S.mutedMini}>Sem sugestoes fortes nesta leitura.</span>}
        </div>
        {hidden.length ? (
          <div style={expanded ? S.editorExpandableOpen : S.editorExpandableClosed}>
            <div style={expanded ? S.editorExpandableScroll : S.editorExpandableHint}>
              {expanded
                ? hidden.map((chip) => (
                  <button key={chip.key} type="button" style={chip.active ? S.miniChipOn : S.miniChip} onClick={chip.onClick}>
                    {chip.label}
                  </button>
                ))
                : helper}
            </div>
          </div>
        ) : null}
      </div>
    );
  }

  function renderClassificationEditorHeader() {
    const focusTitle = classificationFocus === "principal"
      ? "Grupo principal"
      : classificationFocus === "labels"
        ? "Etiquetas"
        : classificationFocus === "ticket"
          ? "Ticket"
          : "Referencias";
    return (
      <div style={S.editorHeader}>
        <div style={S.editorHeaderMeta}>
          <div style={S.sectionTitle}>Classificacao</div>
          <div style={S.editorHeaderTitle}>{focusTitle}</div>
          <div style={S.editorModeText}>{classificationLayoutMode === "advanced" ? "Modo avancado" : "Modo normal"}</div>
        </div>
        <div style={S.editorHeaderActions}>
          <button type="button" style={S.secondaryBtn} onClick={handleCloseClassificationEditor}>Voltar</button>
          <button type="button" style={S.primaryBtn} onClick={() => openApplyDialog(classificationFocus)} disabled={actionBusy || !canApplyClassification}>
            <Icons.Save size={12} />
            Aplicar
          </button>
        </div>
      </div>
    );
  }

  function renderPrincipalEditor() {
    const suggestionChips = suggestedExistingGroups.map((group) => ({
      key: group.id,
      label: group.name || group.id,
      active: group.id === principalGroupId,
      onClick: () => {
        if (group.id === principalGroupId) clearPrincipalSelection();
        else selectPrincipalGroup(group);
      },
    }));
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Grupo principal</div>
        <div style={S.editorLead}>Escolhe ou ajusta o dossier principal do email.</div>
        {renderSuggestionTray("principal", "Sugestoes", suggestionChips, "Ao expandir, aparecem as restantes sugestoes com barra de scroll vertical se forem muitas.")}
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Pesquisar ou criar</div>
          <div style={S.searchInlineRow}>
            <input
              style={S.input}
              value={principalSearch}
              onChange={(event) => setPrincipalSearchValue(event.target.value)}
              placeholder="Escreve o nome do grupo..."
            />
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => {
                if (exactPrincipalSearchGroup) {
                  selectPrincipalGroup(exactPrincipalSearchGroup);
                  return;
                }
                if (principalCanCreate) {
                  setStatus("Este grupo sera criado quando aplicares a classificacao.");
                }
              }}
            >
              Pesquisar
            </button>
          </div>
          {principalSearchResults.length ? (
            <div style={S.searchResultListCompact}>
              {principalSearchResults.map((group) => (
                <button
                  key={group.id}
                  type="button"
                  style={group.id === principalGroupId ? S.searchResultBtnOn : S.searchResultBtn}
                  onClick={() => selectPrincipalGroup(group)}
                >
                  <span>{group.name}</span>
                  {group.id === principalGroupId ? <span style={S.resultMiniMeta}>Selecionado</span> : null}
                </button>
              ))}
            </div>
          ) : null}
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Selecionado</div>
          <div style={S.editorValueStrong}>{principalGroup?.name || (principalCanCreate ? principalSearch || "--" : "--")}</div>
        </div>
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorOptionGrid}>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.principalCategorize} onChange={(event) => updateClassificationMeta({ principalCategorize: event.target.checked })} /> Grupo em categoria Outlook</label>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.principalStatusCategorize} onChange={(event) => updateClassificationMeta({ principalStatusEnabled: event.target.checked, principalStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
            </div>
            {renderOutlookColorLegend()}
          </div>
        ) : null}
      </div>
    );
  }

  function renderLabelsEditor() {
    const suggestionChips = suggestedLabelSeeds.map((label) => ({
      key: label,
      label,
      active: selectedLabels.includes(label),
      onClick: () => applySuggestedLabel(label),
    }));
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Etiquetas</div>
        <div style={S.editorLead}>Liga ou desliga apenas as etiquetas relevantes.</div>
        {renderSuggestionTray("labels", "Sugestoes da leitura", suggestionChips, "Ao expandir, aparecem as restantes sugestoes com barra de scroll vertical se forem muitas.")}
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Selecionadas</div>
          <div style={S.chipGridCompact}>
            {selectedLabels.length ? selectedLabels.map((label) => (
              <button key={label} type="button" style={S.groupChipBtnOn} onClick={() => removeLabel(label)}>{label}</button>
            )) : <span style={S.mutedMini}>Sem etiquetas selecionadas.</span>}
          </div>
        </div>
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Estado da etiqueta</span>
                <select
                  style={S.select}
                  value={selectedLabelSharedStatus}
                  onChange={(event) => {
                    const nextValue = String(event.target.value || "").trim() as EmailLabelStatus | "";
                    selectedLabels.forEach((label) => updateLabelDraft(label, {
                      hasStatus: Boolean(nextValue),
                      status: nextValue || undefined,
                    }));
                  }}
                >
                  <option value="">Sem estado</option>
                  {LABEL_STATUS_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                </select>
              </label>
              <label style={S.compactCheckBoxField}><input type="checkbox" checked={selectedLabels.some((label) => labelDrafts[label]?.categorize === true)} onChange={(event) => selectedLabels.forEach((label) => updateLabelDraft(label, { categorize: event.target.checked }))} /> Etiqueta em categoria Outlook</label>
            </div>
            <label style={S.compactCheckBoxField}><input type="checkbox" checked={selectedLabels.some((label) => labelDrafts[label]?.hasStatus === true)} onChange={(event) => selectedLabels.forEach((label) => updateLabelDraft(label, { hasStatus: event.target.checked, status: event.target.checked ? (labelDrafts[label]?.status || "em_analise") : undefined }))} /> Refletir estado pela cor da categoria</label>
            {renderOutlookColorLegend()}
          </div>
        ) : null}
      </div>
    );
  }

  function renderTicketEditor() {
    const activeList = ticketSearchResults.length ? ticketSearchResults : availableTicketChoices.slice(0, 8);
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Ticket</div>
        <div style={S.editorLead}>Liga um ticket so se houver seguimento operacional.</div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Estado atual</div>
          <div style={S.editorValueStrong}>{selectedTicket?.code || (selectedSeriesId ? "Novo ticket preparado" : "Sem ticket ligado")}</div>
        </div>
        <div style={S.editorSplitRow}>
          <button type="button" style={ticketEditorMode === "existing" ? S.editorModeBtnOn : S.editorModeBtn} onClick={() => setTicketEditorMode("existing")}>Ligar ticket existente</button>
          <button type="button" style={ticketEditorMode === "new" ? S.editorModeBtnOn : S.editorModeBtn} onClick={() => setTicketEditorMode("new")}>Criar novo ticket</button>
        </div>
        {ticketEditorMode === "existing" ? (
          <div style={S.editorBlock}>
            <div style={S.searchInlineRow}>
              <input style={S.input} value={ticketSearch} onChange={(event) => setTicketSearch(event.target.value)} placeholder="Pesquisar ticket por codigo..." />
              <button type="button" style={S.secondaryBtn} onClick={() => void handleSearchTickets()}>Procurar</button>
            </div>
            <div style={S.searchResultListCompact}>
              {activeList.length ? activeList.map((ticket) => (
                <button key={ticket.id} type="button" style={ticket.id === selectedTicketId ? S.searchResultBtnOn : S.searchResultBtn} onClick={() => applySuggestedTicket(ticket.id)}>
                  <span>{ticket.code || ticket.title || "Ticket"}</span>
                  {ticket.id === selectedTicketId ? <span style={S.resultMiniMeta}>Ligado</span> : null}
                </button>
              )) : <span style={S.mutedMini}>Sem tickets disponiveis para ligar.</span>}
            </div>
          </div>
        ) : (
          <div style={S.editorBlock}>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Serie</span>
                <select style={S.select} value={selectedSeriesId} onChange={(event) => { setSelectedSeriesId(event.target.value); setSelectionTouched((current) => ({ ...current, ticket: true })); }}>
                  <option value="">Escolher serie...</option>
                  {ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}
                </select>
              </label>
              <label style={S.field}>
                <span style={S.cardMeta}>Titulo</span>
                <input style={S.input} value={createTicketTitle} onChange={(event) => setCreateTicketTitle(event.target.value)} placeholder="Titulo do novo ticket" />
              </label>
            </div>
          </div>
        )}
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Estado do ticket</span>
                <select style={S.select} value={ticketStatusDraft} onChange={(event) => setTicketStatusDraft(event.target.value)}>
                  {TICKET_STATUS_OPTIONS.map((option) => <option key={option.value || "none"} value={option.value}>{option.label}</option>)}
                </select>
              </label>
              <label style={S.compactCheckBoxField}><input type="checkbox" checked={Boolean(selectedTicketId || selectedSeriesId)} onChange={(event) => { if (!event.target.checked) clearTicketSelection(); }} /> Ticket em categoria Outlook</label>
            </div>
            <label style={S.compactCheckBoxField}><input type="checkbox" checked={classificationMetaDraft.ticketStatusCategorize} onChange={(event) => updateClassificationMeta({ ticketStatusEnabled: event.target.checked, ticketStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
            {renderOutlookColorLegend()}
          </div>
        ) : null}
      </div>
    );
  }

  function renderReferencesEditor() {
    if (classificationLayoutMode !== "advanced") {
      return (
        <div style={S.editorPanelStack}>
          <div style={S.editorModeKicker}>Referencias</div>
          <div style={S.editorLead}>As referencias so aparecem no modo avancado.</div>
        </div>
      );
    }
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Referencias</div>
        <div style={S.editorLead}>Liga este caso a outros dossiers apenas quando houver ligacao estrutural real.</div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Ligadas</div>
          <div style={S.chipGridCompact}>
            {referenceGroups.length ? referenceGroups.map((group) => (
              <button key={group.id} type="button" style={S.groupChipBtnOn} onClick={() => toggleReferenceGroup(group.id)}>
                {group.name || group.id}
              </button>
            )) : <span style={S.mutedMini}>Sem referencias ligadas.</span>}
          </div>
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Pesquisar outro dossier</div>
          <div style={S.searchInlineRow}>
            <input style={S.input} value={referenceSearch} onChange={(event) => setReferenceSearchValue(event.target.value)} placeholder="Escreve para pesquisar..." />
            <button type="button" style={S.secondaryBtn} onClick={() => { if (exactReferenceSearchGroup) toggleReferenceGroup(exactReferenceSearchGroup.id); }}>Procurar</button>
          </div>
          {referenceSearchResults.length ? (
            <div style={S.searchResultListCompact}>
              {referenceSearchResults.map((group) => (
                <button key={group.id} type="button" style={referenceGroupIds.includes(group.id) ? S.searchResultBtnOn : S.searchResultBtn} onClick={() => toggleReferenceGroup(group.id)}>
                  <span>{group.name}</span>
                  {referenceGroupIds.includes(group.id) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
                </button>
              ))}
            </div>
          ) : null}
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Opcoes avancadas</div>
          <div style={S.editorOptionGrid}>
            <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.referenceCategorize} onChange={(event) => updateClassificationMeta({ referenceCategorize: event.target.checked })} /> Referencia em categoria Outlook</label>
            <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.referenceStatusCategorize} onChange={(event) => updateClassificationMeta({ referenceStatusEnabled: event.target.checked, referenceStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
          </div>
          {renderOutlookColorLegend()}
        </div>
      </div>
    );
  }

  function renderClassificationEditorContent() {
    if (classificationFocus === "principal") return renderPrincipalEditor();
    if (classificationFocus === "labels") return renderLabelsEditor();
    if (classificationFocus === "ticket") return renderTicketEditor();
    return renderReferencesEditor();
  }

  function renderApplyDialog() {
    if (!applyDialogOpen) return null;
    const sectionLabel = applyDialogSection === "principal"
      ? "Grupo principal"
      : applyDialogSection === "labels"
        ? "Etiquetas"
        : applyDialogSection === "ticket"
          ? "Ticket"
          : applyDialogSection === "references"
            ? "Referencias"
            : "Classificacao";
    return (
      <div style={S.modalBackdrop}>
        <div style={S.modalSheet}>
          <div style={S.modalHeader}>
            <div>
              <div style={S.kicker}>Aplicar alteracoes</div>
              <div style={S.modalTitle}>{sectionLabel}</div>
            </div>
            <button type="button" style={S.secondaryBtn} onClick={() => setApplyDialogOpen(false)}>Cancelar</button>
          </div>
          <div style={S.modalScopeRow}>
            <button type="button" style={applyDialogScopeMode === "current" ? S.scopeChipOn : S.scopeChip} onClick={() => setApplyDialogScope("current")}>So este email</button>
            <button type="button" style={applyDialogScopeMode === "selected" ? S.scopeChipOn : S.scopeChip} onClick={() => setApplyDialogScope("selected")}>Emails selecionados</button>
            <button type="button" style={applyDialogScopeMode === "case_all" ? S.scopeChipOn : S.scopeChip} onClick={() => setApplyDialogScope("case_all")}>Todos os emails do caso</button>
          </div>
          <div style={S.modalBlock}>
            <div style={S.modalBlockHeader}>
              <div style={S.editorBlockTitle}>Escolher emails</div>
              <button type="button" style={S.linkBtn} onClick={() => setApplyDialogEmailKeys(caseScopeEmails.map((email) => makeEmailKey(email)).filter(Boolean))}>Selecionar todos</button>
            </div>
            <div style={S.applyEmailList}>
              {caseScopeEmails.map((email) => {
                const emailKey = makeEmailKey(email);
                const expanded = applyDialogExpandedEmailKeys.includes(emailKey);
                const checked = applyDialogEmailKeys.includes(emailKey);
                const previewText = String(email.bodyText || htmlToPlainText(email.bodyHtml || "") || buildSnippet(email) || "").trim();
                return (
                  <div key={emailKey} style={checked ? S.applyEmailRowOn : S.applyEmailRow}>
                    <div style={S.applyEmailRowTop}>
                      <label style={S.applyEmailMain}>
                        <input type="checkbox" checked={checked} onChange={() => toggleApplyDialogEmailKey(emailKey)} />
                        <span style={S.applyEmailSubject}>{email.subject || "(sem assunto)"}</span>
                        <span style={S.applyEmailMeta}>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</span>
                      </label>
                      <button type="button" style={S.chevronBtn} onClick={() => toggleApplyDialogExpandedEmailKey(emailKey)}>{expanded ? "⌃" : "⌄"}</button>
                    </div>
                    {expanded ? (
                      <div style={S.applyEmailPreview}>{previewText || "Sem preview resumido para este email."}</div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          </div>
          <div style={S.modalFooter}>
            <button type="button" style={S.secondaryBtn} onClick={() => setApplyDialogOpen(false)}>Cancelar</button>
            <button type="button" style={S.primaryBtn} onClick={() => void handleConfirmApplyDialog()} disabled={actionBusy || !applyDialogSelectedEmails.length}>
              Confirmar aplicacao
            </button>
          </div>
        </div>
      </div>
    );
  }

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
            {selectedEmailAttachments.length ? (
              <div style={S.stackMini}>
                <div style={S.attachmentPickerBar}>
                  {selectedEmailAttachments.map((attachment) => {
                    const key = makeAttachmentKey(attachment);
                    const active = key === selectedAttachmentPreviewKey;
                    return (
                      <button
                        key={key}
                        type="button"
                        style={active ? S.groupChipBtnOn : S.groupChipBtn}
                        onClick={() => setSelectedAttachmentPreviewKey(key)}
                      >
                        {attachment.name}
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
                        style={group.id === principalGroupId ? S.miniChipOn : S.miniChip}
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
                        style={group.id === principalGroupId ? S.searchResultBtnOn : S.searchResultBtn}
                        onClick={() => {
                          if (group.id === principalGroupId) {
                            clearPrincipalSelection();
                            return;
                          }
                          selectPrincipalGroup(group);
                        }}
                      >
                        <span>{group.name}</span>
                        {group.id === principalGroupId ? <span style={S.resultMiniMeta}>Ligado</span> : null}
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
                        style={referenceGroupIds.includes(group.id) ? S.miniChipOn : S.miniChip}
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
                        setReferenceSearchValue(exactReferenceSearchGroup.name);
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
                        style={referenceGroupIds.includes(group.id) ? S.searchResultBtnOn : S.searchResultBtn}
                        onClick={() => {
                          toggleReferenceGroup(group.id);
                          setReferenceSearchValue(group.name);
                        }}
                      >
                        <span>{group.name}</span>
                        {referenceGroupIds.includes(group.id) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
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
                    placeholder="Escreve o nome da etiqueta..."
                  />
                  <button
                    type="button"
                    style={String(classificationLabelInput || "").trim() ? S.iconActionBtn : S.iconActionBtnDisabled}
                    onClick={() => {
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
                    }}
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
                {(ticketSearchResults.length ? ticketSearchResults : availableTicketChoices.slice(0, 12)).map((ticket) => (
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
                  {ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}
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
                <button type="button" style={S.primaryBtn} onClick={() => void handleApplyClassification()} disabled={actionBusy || (!principalGroupId && !referenceGroupIds.length && !selectedTicketId && !selectedSeriesId && !selectedEmailGroups.length && !selectedEmailTicketIds.length && !selectedLabels.length && !(selectedEmail?.labels || []).length && !String(selectedEmail?.status || "").trim())}>
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
              <label style={S.field}><span style={S.label}>Filtrar por ticket</span><select style={S.select} value={ticketFilterId} onChange={(event) => setTicketFilterId(event.target.value)}><option value="">Sem filtro</option>{contextualTickets.map((ticket) => <option key={ticket.id} value={ticket.id}>{ticket.code} · {ticket.title}</option>)}</select></label>
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
            <div style={S.summaryRow}><span>Emails relacionados</span><strong>{relatedEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Total conhecido</span><strong>{dedupeEmails([...relatedEmails, ...knownEmails]).length}</strong></div>
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
                        {contact.name}{contact.company ? ` · ${contact.company}` : ""}{contact.email ? ` · ${contact.email}` : ""}
                      </button>
                    )) : <span style={S.mutedMini}>Sem contactos associados.</span>}
                  </div>
                  <div style={S.chips}>
                    {selectedManagedGroup ? filteredManagedGroupContacts.slice(0, 18).map((contact) => {
                      const active = managedGroupContacts.some((entry) => entry.key === contact.key);
                      return (
                        <button key={contact.key} type="button" style={active ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleManagedGroupContact(contact)}>
                          {contact.name}{contact.company ? ` · ${contact.company}` : ""}{contact.email ? ` · ${contact.email}` : ""}
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
                        {entity.name}{entity.kind ? ` · ${entity.kind}` : ""}
                      </button>
                    )) : <span style={S.mutedMini}>Sem entidades associadas.</span>}
                  </div>
                  <div style={S.chips}>
                    {selectedManagedGroup ? filteredManagedGroupEntities.slice(0, 18).map((entity) => {
                      const active = managedGroupEntities.some((entry) => entry.key === entity.key);
                      return (
                        <button key={entity.key} type="button" style={active ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleManagedGroupEntity(entity)}>
                          {entity.name}{entity.kind ? ` · ${entity.kind}` : ""}
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
                        <small>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</small>
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
                        <small>{document.contentType || "ficheiro"}{document.size ? ` · ${Math.round(Number(document.size || 0) / 1024)} KB` : ""}</small>
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
                (!principalGroupId &&
                  !referenceGroupIds.length &&
                  !selectedTicketId &&
                  !selectedSeriesId &&
                  !selectedEmailGroups.length &&
                  !selectedEmailTicketIds.length &&
                  !selectedLabels.length &&
                  !(selectedEmail?.labels || []).length &&
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
    <div style={S.root}>
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
          <button type="button" style={S.primaryBtn} onClick={() => openApplyDialog(classificationEditorActive ? classificationFocus : "summary")} disabled={actionBusy || !canApplyClassification}>
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

          <section style={emailsCardStyle}>
            <div style={S.sectionHeaderCompact}>
              <div>
                <div style={S.sectionTitle}>Emails</div>
                <div style={S.sectionSubtitle}>Exploracao do caso</div>
              </div>
              <span style={S.cardMeta}>Selecionados: {selectedTargetCount}</span>
            </div>
            <div style={S.emailTools}>
              <button type="button" style={S.linkBtn} onClick={selectAllVisibleEmails}>Todos visiveis</button>
              <button type="button" style={S.linkBtn} onClick={clearSelectedTargets}>Limpar</button>
            </div>
            <input style={S.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Pesquisar por assunto, remetente ou texto..." />
            <div style={S.topCardScroll}>
              {loading ? <PanelState compact tone="loading" title="A carregar emails" description="A preparar a lista desta nova janela." /> : null}
              {!loading && !visibleEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Ajusta os filtros ou muda a fonte da lista." /> : null}
              {!loading && visibleEmails.map((email) => (
                <button key={makeEmailKey(email)} type="button" style={makeEmailKey(email) === makeEmailKey(selectedEmail || {}) ? S.emailOn : S.email} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                  <div style={S.emailTop}>
                    <label style={S.emailPick} onClick={(event) => event.stopPropagation()}>
                      <input
                        type="checkbox"
                        checked={selectedTargetEmailKeys.includes(makeEmailKey(email))}
                        onChange={() => toggleTargetEmailKey(makeEmailKey(email))}
                      />
                      <span style={S.emailSubject}>{email.subject || "(sem assunto)"}</span>
                    </label>
                    {Array.isArray(email.attachments) && email.attachments.length ? <span style={S.counter}>{email.attachments.length}</span> : null}
                  </div>
                  <div style={S.emailMeta}>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</div>
                  <div style={S.emailSnippet}>{buildSnippet(email) || "Sem preview curto disponivel."}</div>
                </button>
              ))}
            </div>
          </section>

          <section style={quickDocumentsCardStyle}>
            <div style={S.sectionHeaderCompact}>
              <div>
                <div style={S.sectionTitle}>Documentos rapidos</div>
                <div style={S.sectionSubtitle}>Do email selecionado</div>
              </div>
              {quickDocumentHiddenCount ? (
                <button
                  type="button"
                  style={showHiddenQuickDocuments ? S.quietToggleBtnOn : S.quietToggleBtn}
                  onClick={() => setShowHiddenQuickDocuments((current) => !current)}
                >
                  {showHiddenQuickDocuments ? `Ocultar ocultos (${quickDocumentHiddenCount})` : `Mostrar ocultos (${quickDocumentHiddenCount})`}
                </button>
              ) : null}
            </div>
            <div style={S.topCardScroll}>
              {!quickDocumentAttachments.length ? (
                <PanelState compact tone="info" title="Sem documentos rapidos" description="Este email ainda nao tem anexos persistidos para abrir aqui." />
              ) : (
                <div style={S.quickDocList}>
                  {quickDocumentAttachments.map((attachment) => {
                    const key = makeAttachmentKey(attachment);
                    const active = key === selectedAttachmentPreviewKey && previewMode === "document";
                    const hidden = isStudioAttachmentHiddenInQuickDocs(attachment);
                    const rowStyle = active
                      ? hidden
                        ? { ...S.quickDocRowOn, ...S.quickDocRowHidden }
                        : S.quickDocRowOn
                      : hidden
                        ? { ...S.quickDocRow, ...S.quickDocRowHidden }
                        : S.quickDocRow;
                    return (
                      <div
                        key={key}
                        style={rowStyle}
                        role="button"
                        tabIndex={0}
                        onClick={() => handleOpenQuickAttachment(attachment)}
                        onKeyDown={(event) => {
                          if (event.key === "Enter" || event.key === " ") {
                            event.preventDefault();
                            handleOpenQuickAttachment(attachment);
                          }
                        }}
                      >
                        <div style={S.quickDocMain}>
                          <span style={S.quickDocTitle}>{attachment.name || "Anexo"}</span>
                          <span style={S.quickDocMeta}>{formatQuickDocumentMeta(attachment)}</span>
                        </div>
                        <div style={S.quickDocActions}>
                          {hidden ? <span style={S.quickDocHint}>oculto</span> : null}
                          <button
                            type="button"
                            style={hidden ? S.quickDocActionBtnOn : S.quickDocActionBtn}
                            onClick={(event) => {
                              event.stopPropagation();
                              void handleSetQuickAttachmentHidden(attachment, hidden ? false : true);
                            }}
                            disabled={actionBusy}
                          >
                            {hidden ? "Mostrar" : "Ocultar"}
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </section>

          <section style={classificationCardStyle}>
            <div style={S.sectionHeaderCompact}>
              <div>
                <div style={S.sectionTitle}>Classificacao</div>
                <div style={S.sectionSubtitle}>{classificationEditorActive ? "Editor aberto" : "Resumo do que esta atribuido"}</div>
              </div>
              <div style={S.segmentedControl}>
                <button type="button" style={classificationLayoutMode === "normal" ? S.segmentBtnActive : S.segmentBtn} onClick={() => setClassificationLayoutMode("normal")}>Normal</button>
                <button type="button" style={classificationLayoutMode === "advanced" ? S.segmentBtnActive : S.segmentBtn} onClick={() => setClassificationLayoutMode("advanced")}>Avancado</button>
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
              <div style={S.classificationSummary}>
                {classificationSummaryTiles
                  .filter((item) => classificationLayoutMode === "advanced" || item.key !== "references")
                  .map((item) => (
                    <button key={item.key} type="button" style={S.classificationTile} onClick={item.onClick}>
                      <span style={S.classificationTileLabel}>{item.title}</span>
                      <span style={S.classificationTileValue}>{item.value}</span>
                      <span style={S.classificationTileMeta}>{item.description}</span>
                    </button>
                  ))}
                <div style={S.classificationModeHint}>
                  {classificationLayoutMode === "normal"
                    ? "Modo normal: grupo principal, etiquetas e ticket."
                    : "Modo avancado: inclui referencias e opcoes finas."}
                </div>
              </div>
            ) : (
              <div style={S.classificationEditorShell}>
                {renderClassificationEditorHeader()}
                <div style={S.classificationEditorBody}>{renderClassificationEditorContent()}</div>
              </div>
            )}
          </section>
        </div>

        <section style={previewShellStyle}>
            <div style={S.previewToolbar}>
              <button type="button" style={previewMode === "email" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("email")} disabled={!previewHtml}>Email</button>
              <button type="button" style={previewMode === "document" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("document")} disabled={!previewHasDocument}>Documento</button>
              <button type="button" style={previewMode === "reply" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("reply")} disabled={!selectedEmail}>Responder</button>
              <button type="button" style={previewMode === "forward" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("forward")} disabled={!selectedEmail}>Reencaminhar</button>
            </div>
            <div style={S.previewBody}>
              {previewMode === "email" ? (
                previewHtml ? (
                  <div style={S.previewHtml} dangerouslySetInnerHTML={{ __html: previewHtml }} />
                ) : (
                  <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado suficiente para preview." />
                )
              ) : null}
              {previewMode === "document" ? (
                selectedAttachmentPreview ? (
                  <div style={S.documentPreviewShell}>
                    {selectedAttachmentDocumentPreview?.kind === "image" ? (
                      <div style={S.documentPreviewFrame}>
                        <img src={selectedAttachmentDocumentPreview.src} alt={selectedAttachmentPreview?.name || "Imagem"} style={S.attachmentPreviewImage} />
                      </div>
                    ) : null}
                    {selectedAttachmentDocumentPreview?.kind === "pdf" ? (
                      <div style={S.documentPreviewFrame}>
                        {selectedAttachmentDocumentPreview.src.startsWith("data:")
                          ? <StudioPdfPreview dataUrl={selectedAttachmentDocumentPreview.src} title={selectedAttachmentPreview?.name || "PDF"} />
                          : <iframe title={selectedAttachmentPreview?.name || "PDF"} src={selectedAttachmentDocumentPreview.src} style={S.documentPreviewIframe} />}
                      </div>
                    ) : null}
                    {selectedAttachmentDocumentPreview?.kind === "office" ? (
                      <div style={S.documentPreviewFrame}>
                        <iframe title={selectedAttachmentPreview?.name || "Documento"} src={selectedAttachmentDocumentPreview.url} style={S.documentPreviewIframe} />
                      </div>
                    ) : null}
                    {selectedAttachmentDocumentPreview?.kind === "text" ? (
                      <pre style={S.attachmentPreviewText}>{selectedAttachmentDocumentPreview.text}</pre>
                    ) : null}
                    {!selectedAttachmentDocumentPreview && selectedAttachmentPreviewRemoteStatus === "loading" ? (
                      <PanelState compact tone="loading" title="A carregar documento" description="A preparar o preview do documento selecionado." />
                    ) : null}
                    {selectedAttachmentDocumentPreview?.kind === "unsupported" ? (
                      <PanelState compact tone="info" title="Preview nao disponivel" description="Este documento pode exigir download ou URL publica para preview." />
                    ) : null}
                    {!selectedAttachmentDocumentPreview && selectedAttachmentPreviewRemoteStatus !== "loading" && selectedAttachmentPreviewMode !== "none" ? (
                      <PanelState compact tone="info" title="Preview nao disponivel" description="Nao foi possivel abrir este documento com a mesma base de preview da aba Grupos." />
                    ) : null}
                    {selectedAttachmentPreviewMode === "none" ? (
                      <PanelState compact tone="info" title="Escolhe um documento" description="Seleciona um documento rapido para abrir o preview." />
                    ) : null}
                  </div>
                ) : (
                  <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe primeiro um documento rapido para abrir o preview." />
                )
              ) : null}
              {previewMode === "reply" ? (
                <div style={S.previewPlaceholder}>
                  <div style={S.cardTitle}>Responder</div>
                  <div style={S.cardMeta}>Estrutura pronta para editor, IA e selecao de anexos numa fase seguinte.</div>
                  <button type="button" style={S.primaryBtn} onClick={() => void handlePreviewReply()} disabled={!selectedEmail}>
                    <Icons.MessageSquare size={12} />
                    Abrir resposta
                  </button>
                </div>
              ) : null}
              {previewMode === "forward" ? (
                <div style={S.previewPlaceholder}>
                  <div style={S.cardTitle}>Reencaminhar</div>
                  <div style={S.cardMeta}>Estrutura pronta para editor, IA e composicao de envio numa fase seguinte.</div>
                  <button type="button" style={S.primaryBtn} onClick={() => void handlePreviewForward()} disabled={!selectedEmail}>
                    <Icons.ExternalLink size={12} />
                    Abrir reencaminhamento
                  </button>
                </div>
              ) : null}
            </div>
        </section>
      </div>
      {renderApplyDialog()}
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
  topCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.9)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 6, overflow: "hidden", transition: "transform 180ms ease, width 180ms ease, box-shadow 180ms ease" },
  topCardWide: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.9)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 6, overflow: "hidden", transition: "transform 180ms ease, width 180ms ease, box-shadow 180ms ease" },
  focusEmailsCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.92)", boxShadow: "0 8px 20px rgba(15,23,42,0.03)", padding: 8, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 6, overflow: "hidden", gridColumn: "1", gridRow: "1" },
  focusQuickDocumentsCard: { minHeight: 0, borderRadius: 12, border: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.88)", boxShadow: "0 6px 18px rgba(15,23,42,0.025)", padding: 8, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 6, overflow: "hidden", gridColumn: "2", gridRow: "1" },
  focusClassificationCard: { minHeight: 0, borderRadius: 14, border: "1px solid rgba(37,99,235,0.18)", background: "rgba(255,255,255,0.97)", boxShadow: "0 18px 36px rgba(37,99,235,0.08)", padding: 10, display: "grid", gridTemplateRows: "auto minmax(0,1fr)", gap: 8, overflow: "hidden", gridColumn: "3", gridRow: "1 / span 2" },
  topCardScroll: { minHeight: 0, display: "grid", gap: 5, overflowY: "auto", paddingRight: 1 },
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
  input: { width: "100%", height: 30, boxSizing: "border-box", borderRadius: 9, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(248,250,252,0.92)", padding: "0 9px", fontSize: 11, color: "var(--iccc-text)", outline: "none" },
  textarea: { width: "100%", minHeight: 120, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "10px 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none", resize: "vertical" },
  select: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  listBody: { minHeight: 0, display: "grid", gap: 8, overflowY: "auto", paddingRight: 2 },
  email: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.78)", padding: "7px 8px", display: "grid", gap: 3, cursor: "pointer" },
  emailOn: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.96)", padding: "7px 8px", display: "grid", gap: 3, cursor: "pointer" },
  emailTop: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 },
  emailPick: { display: "flex", alignItems: "flex-start", gap: 8, minWidth: 0, cursor: "pointer" },
  emailSubject: { fontSize: 10.75, fontWeight: 550, lineHeight: 1.25, color: "var(--iccc-text)", minWidth: 0, textAlign: "left" },
  emailMeta: { fontSize: 9.25, color: "var(--iccc-muted)" },
  emailSnippet: { fontSize: 10, lineHeight: 1.3, color: "var(--iccc-text-soft, #334155)", maxHeight: 28, overflow: "hidden", opacity: 0.88 },
  counter: { minWidth: 16, height: 16, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(15,23,42,0.06)", color: "var(--iccc-text)", fontSize: 9.25, fontWeight: 700 },
  quickDocList: { display: "grid", gap: 5 },
  quickDocRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center", borderRadius: 10, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.78)", padding: "7px 8px" },
  quickDocRowOn: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center", borderRadius: 10, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.96)", padding: "7px 8px" },
  quickDocRowHidden: { opacity: 0.8 },
  quickDocMain: { display: "grid", gap: 3, minWidth: 0 },
  quickDocTitle: { fontSize: 10.5, fontWeight: 550, color: "var(--iccc-text)", lineHeight: 1.2, minWidth: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  quickDocMeta: { fontSize: 9.25, color: "var(--iccc-muted)", lineHeight: 1.15 },
  quickDocActions: { display: "flex", alignItems: "center", gap: 6, justifyContent: "flex-end" },
  quickDocHint: { fontSize: 9, fontWeight: 700, color: "var(--iccc-muted)" },
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
  classificationTile: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.76)", padding: "8px 9px", display: "grid", gap: 2, cursor: "pointer" },
  classificationTileMuted: { width: "100%", textAlign: "left", borderRadius: 10, border: "1px solid rgba(191,219,254,0.52)", background: "rgba(239,246,255,0.62)", padding: "8px 9px", display: "grid", gap: 2, cursor: "pointer" },
  classificationTileLabel: { fontSize: 8.5, fontWeight: 700, letterSpacing: "0.09em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  classificationTileValue: { fontSize: 11.25, fontWeight: 550, color: "var(--iccc-text)", lineHeight: 1.2 },
  classificationTileMeta: { fontSize: 9.5, lineHeight: 1.25, color: "var(--iccc-muted)" },
  classificationModeHint: { padding: "7px 9px", borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.82)", color: "var(--iccc-muted)", fontSize: 9.75, lineHeight: 1.35 },
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
  editorExpandableClosed: { borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.82)", padding: "7px 9px" },
  editorExpandableOpen: { borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.82)", padding: "7px 9px" },
  editorExpandableScroll: { display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 96, overflowY: "auto", alignContent: "flex-start" },
  editorExpandableHint: { fontSize: 9.5, lineHeight: 1.35, color: "var(--iccc-muted)" },
  chipGridCompact: { display: "flex", flexWrap: "wrap", gap: 6 },
  editorOptionGrid: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  editorAdvancedFieldGrid: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 8 },
  compactCheck: { display: "flex", alignItems: "center", gap: 8, fontSize: 10.5, color: "var(--iccc-text)" },
  compactCheckBoxField: { minHeight: 34, display: "flex", alignItems: "center", gap: 8, padding: "0 10px", borderRadius: 10, border: "1px solid rgba(148,163,184,0.14)", background: "rgba(255,255,255,0.88)", fontSize: 10.5, color: "var(--iccc-text)" },
  searchInlineRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "center" },
  searchResultListCompact: { display: "grid", gap: 6, maxHeight: 172, overflowY: "auto", paddingRight: 1 },
  chevronBtn: { width: 24, height: 24, borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.88)", color: "#475569", fontSize: 12, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  legendRow: { display: "flex", flexWrap: "wrap", gap: 6 },
  legendChip: { display: "inline-flex", alignItems: "center", padding: "3px 8px", borderRadius: 999, border: "1px solid transparent", fontSize: 9.5, fontWeight: 700 },
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
  applyEmailRowTop: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 8, alignItems: "start" },
  applyEmailMain: { display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", alignItems: "start", gap: 8, minWidth: 0 },
  applyEmailSubject: { fontSize: 10.75, fontWeight: 600, color: "var(--iccc-text)", lineHeight: 1.25, display: "block", minWidth: 0 },
  applyEmailMeta: { fontSize: 9.5, color: "var(--iccc-muted)", lineHeight: 1.2, display: "block", marginTop: 2 },
  applyEmailPreview: { maxHeight: 92, overflowY: "auto", padding: "7px 9px", borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(255,255,255,0.9)", color: "var(--iccc-text-soft, #334155)", fontSize: 10, lineHeight: 1.4, whiteSpace: "pre-wrap" },
  modalFooter: { display: "flex", justifyContent: "flex-end", gap: 8, flexWrap: "wrap" },
};

