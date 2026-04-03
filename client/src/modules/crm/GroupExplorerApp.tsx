import React, { useEffect, useMemo, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import {
  getEmailAttachmentContentBase64,
  getEmailAttachmentTextContent,
  getGroupDocumentTextContent,
  getGroupDocumentContentUrl,
  getGroupDocuments,
  getGroupEmails,
  getRelatedEmailContext,
  listLinkGroups,
  removeEmailFromLinkGroup,
  type GroupDocumentEntry,
  type GroupTicketEntry,
  type LinkGroupEntry,
  type RelatedEmailEntry,
} from "@/api";
import { getAttachments, getSelectedMessageContext, openGroupClassificationStudio, openGroupSettings, openLinkedOutlookEmail, requestCockpitHostAction, type OutlookAttachment, type OutlookMessageContext } from "@/office";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import "../../global.css";

pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

type EmailSortMode = "date_desc" | "date_asc" | "subject_asc" | "subject_desc";
type EmailAttachmentFilter = "all" | "with" | "without";
type DocumentFilterMode = "all" | "selected_email";
type PreviewMode = "email" | "document" | "reply" | "forward";
type ClassificationMode = "normal" | "advanced";
type ClassificationEditor = "summary" | "principal" | "labels" | "ticket" | "references";
type PreviewState =
  | { kind: "image"; src: string }
  | { kind: "pdf"; src: string }
  | { kind: "office"; url: string }
  | { kind: "text"; text: string }
  | { kind: "unsupported" };
type EmailAttachmentEntry = NonNullable<RelatedEmailEntry["attachments"]>[number];
type QuickDocumentEntry = {
  key: string;
  title: string;
  meta: string;
  attachmentCount?: number;
  kind: "attachment" | "document";
  attachment?: EmailAttachmentEntry;
  document?: GroupDocumentEntry;
};
type ClassificationContextState = {
  email: RelatedEmailEntry | null;
  emails: RelatedEmailEntry[];
  groups: LinkGroupEntry[];
  tickets: GroupTicketEntry[];
};

function dataUrlToUint8Array(dataUrl: string): Uint8Array {
  const base64 = stripDataUrlPrefix(dataUrl);
  const binary = globalThis.atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function PdfPreview({ dataUrl, title }: { dataUrl: string; title: string }) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");
  const [pageCount, setPageCount] = useState(0);

  useEffect(() => {
    let cancelled = false;
    const host = hostRef.current;
    if (!host || !dataUrl) {
      setStatus("error");
      return;
    }

    host.innerHTML = "";
    setStatus("loading");
    setPageCount(0);

    (async () => {
      try {
        const loadingTask = pdfjsLib.getDocument({ data: dataUrlToUint8Array(dataUrl) });
        const pdf = await loadingTask.promise;
        if (cancelled) {
          void loadingTask.destroy();
          return;
        }

        const nextPageCount = Number(pdf.numPages || 0);
        setPageCount(nextPageCount);

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
        console.warn("[group-explorer] pdf preview failed", error);
        if (!cancelled) setStatus("error");
      }
    })();

    return () => {
      cancelled = true;
      if (hostRef.current) hostRef.current.innerHTML = "";
    };
  }, [dataUrl]);

  if (status === "error") {
    return <PanelState compact tone="info" title="Preview PDF indisponivel" description="Este PDF foi detetado, mas nao foi possivel renderiza-lo dentro do add-in." />;
  }

  return (
    <div style={styles.pdfPreviewShell} aria-label={title}>
      {status === "loading" ? <div style={styles.pdfPreviewMeta}>A carregar PDF...</div> : null}
      {status === "ready" && pageCount > 0 ? <div style={styles.pdfPreviewMeta}>{pageCount} pagina(s)</div> : null}
      <div ref={hostRef} style={styles.pdfPreviewCanvasHost} />
    </div>
  );
}

function closeExplorer() {
  try { (window as any).Office?.context?.ui?.messageParent?.("close"); } catch {}
  try { window.close(); } catch {}
  try { window.location.assign(window.location.origin); } catch {}
}

function normalizeProvider(value: string | undefined): "cloud" | "local" | "onedrive" {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "local" || normalized === "onedrive") return normalized;
  return "cloud";
}

function providerLabel(value: string | undefined): string {
  const provider = normalizeProvider(value);
  if (provider === "onedrive") return "OneDrive / SharePoint";
  if (provider === "local") return "Pasta local do utilizador";
  return "Cockpit Cloud";
}

function readExplorerParams() {
  const params = new URLSearchParams(window.location.search);
  return {
    groupId: String(params.get("groupId") || "").trim(),
    emailKey: String(params.get("emailKey") || "").trim(),
    documentId: String(params.get("documentId") || "").trim(),
  };
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function formatBytes(value: number | undefined): string {
  const size = Number(value || 0);
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function normalizeMessageId(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function normalizeCidValue(value: string | undefined): string {
  return String(value || "")
    .trim()
    .replace(/^cid:/i, "")
    .replace(/[<>\s]/g, "")
    .toLowerCase();
}

function stripDataUrlPrefix(value: string | undefined): string {
  return String(value || "").trim().replace(/^data:[^,]+,/, "");
}

function normalizeDocumentMimeType(value: string | undefined, name: string | undefined): string {
  const raw = String(value || "").trim().toLowerCase();
  const fileName = String(name || "").trim().toLowerCase();
  if (raw === "application/x-pdf" || (!raw && /\.pdf$/.test(fileName))) return "application/pdf";
  if (raw === "image/jpg") return "image/jpeg";
  return raw || "application/octet-stream";
}

function normalizeAttachmentMimeType(value: string | undefined, name: string | undefined): string {
  return normalizeDocumentMimeType(value, name);
}

function buildAttachmentDataUrl(attachment: Partial<EmailAttachmentEntry & OutlookAttachment>): string {
  const base64 = stripDataUrlPrefix(attachment?.content);
  if (!base64) return "";
  return `data:${normalizeAttachmentMimeType(attachment?.contentType, attachment?.name)};base64,${base64}`;
}

function attachmentMatchesCid(attachment: Partial<EmailAttachmentEntry & OutlookAttachment>, cid: string): boolean {
  const normalizedCid = normalizeCidValue(cid);
  if (!normalizedCid) return false;
  const normalizedContentId = normalizeCidValue((attachment as any)?.contentId);
  if (normalizedContentId && normalizedContentId === normalizedCid) return true;
  const normalizedName = String(attachment?.name || "").trim().toLowerCase();
  if (!normalizedName) return false;
  return normalizedCid === normalizedName || normalizedCid.startsWith(`${normalizedName}@`) || normalizedCid.includes(normalizedName);
}

function mergeEmailAttachments(
  persistedAttachments: EmailAttachmentEntry[] | undefined,
  liveAttachments: OutlookAttachment[] | undefined
): EmailAttachmentEntry[] {
  const merged = new Map<string, EmailAttachmentEntry>();
  const addAttachment = (attachment: Partial<EmailAttachmentEntry & OutlookAttachment>) => {
    const name = String(attachment?.name || "").trim();
    if (!name) return;
    const key = [
      String((attachment as any)?.id || "").trim(),
      normalizeCidValue((attachment as any)?.contentId),
      name.toLowerCase(),
      String(attachment?.contentType || "").trim().toLowerCase(),
    ].join("|");
    const current = merged.get(key);
    const next: EmailAttachmentEntry = {
      id: String((attachment as any)?.id || current?.id || "").trim() || undefined,
      name,
      contentType: String(attachment?.contentType || current?.contentType || "").trim(),
      size: Number(attachment?.size || current?.size || 0) || undefined,
      isInline: Boolean((attachment as any)?.isInline ?? current?.isInline),
      contentId: String((attachment as any)?.contentId || current?.contentId || "").trim() || undefined,
      content: String(attachment?.content || current?.content || "").trim(),
    };
    merged.set(key, next);
  };
  (persistedAttachments || []).forEach(addAttachment);
  (liveAttachments || []).forEach(addAttachment);
  return Array.from(merged.values());
}

function emailMatchesCurrentContext(email: Partial<RelatedEmailEntry>, ctx: OutlookMessageContext | null): boolean {
  if (!ctx) return false;
  const currentItemId = String(ctx.itemId || "").trim();
  const emailItemId = String(email.itemId || "").trim();
  if (currentItemId && emailItemId && currentItemId === emailItemId) return true;

  const currentMessageId = normalizeMessageId(ctx.internetMessageId);
  const emailMessageId = normalizeMessageId(email.internetMessageId);
  if (currentMessageId && emailMessageId && currentMessageId === emailMessageId) return true;

  const currentConversationId = String(ctx.conversationId || "").trim();
  const emailConversationId = String(email.conversationId || "").trim();
  const currentSubject = String(ctx.subject || "").trim().toLowerCase();
  const emailSubject = String(email.subject || "").trim().toLowerCase();
  return Boolean(currentConversationId && emailConversationId && currentConversationId === emailConversationId && currentSubject && currentSubject === emailSubject);
}

function rewriteEmailHtmlInlineImages(
  html: string,
  attachments: EmailAttachmentEntry[] | OutlookAttachment[]
): string {
  if (!html) return "";
  return html
    .replace(/\b(src|background)=(["'])cid:([^"'<>]+)\2/gi, (match, attr, quote, cid) => {
      const attachment = (attachments || []).find((entry) => attachmentMatchesCid(entry, cid));
      const dataUrl = attachment ? buildAttachmentDataUrl(attachment) : "";
      if (!dataUrl) return `data-iccc-missing-${attr}=${quote}${cid}${quote}`;
      return `${attr}=${quote}${dataUrl}${quote}`;
    })
    .replace(/url\((["']?)cid:([^)"']+)\1\)/gi, (match, quote, cid) => {
      const attachment = (attachments || []).find((entry) => attachmentMatchesCid(entry, cid));
      const dataUrl = attachment ? buildAttachmentDataUrl(attachment) : "";
      return dataUrl ? `url(${quote}${dataUrl}${quote})` : "none";
    });
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function makeDocumentKey(document: Partial<GroupDocumentEntry>): string {
  return String(document?.id || document?.storagePathHint || document?.name || "");
}

function makeAttachmentKey(attachment: Partial<EmailAttachmentEntry>): string {
  return String(attachment?.key || attachment?.id || attachment?.contentId || attachment?.name || "");
}

function formatStatusLabel(value: string | undefined): string {
  const normalized = String(value || "").trim();
  if (!normalized) return "Sem estado";
  return normalized
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (match) => match.toUpperCase());
}

function formatChipValue(value: string | undefined, fallback = "Sem dados"): string {
  return String(value || "").trim() || fallback;
}

function emailHasAttachments(email: RelatedEmailEntry): boolean {
  return Array.isArray(email.attachments) && email.attachments.length > 0;
}

function getEmailTimestamp(email: RelatedEmailEntry): number {
  const parsed = new Date(String(email.messageDateIso || email.receivedAtIso || email.sentAtIso || "").trim()).getTime();
  return Number.isFinite(parsed) ? parsed : 0;
}

function inferDocumentKind(document: GroupDocumentEntry): "image" | "pdf" | "office" | "text" | "unsupported" {
  const name = String(document.name || "").toLowerCase();
  const type = normalizeDocumentMimeType(document.contentType, document.name);
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
  if (!stripDataUrlPrefix(document.contentBase64)) return "unsupported";
  return "unsupported";
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

function buildOfficePreviewUrl(groupId: string, document: GroupDocumentEntry | null): string {
  if (!groupId || !document?.id || !canUseOfficeWebViewer()) return "";
  const sourceUrl = getGroupDocumentContentUrl(groupId, document.id);
  return `https://view.officeapps.live.com/op/embed.aspx?src=${encodeURIComponent(sourceUrl)}`;
}

function buildDocumentPreview(groupId: string, document: GroupDocumentEntry | null, textPreview?: string): PreviewState | null {
  if (!document) return null;
  const kind = inferDocumentKind(document);
  if (kind === "office") {
    const url = buildOfficePreviewUrl(groupId, document);
    return url ? { kind, url } : { kind: "unsupported" };
  }
  if (kind === "text") {
    if (typeof textPreview === "string" && textPreview.trim()) return { kind, text: textPreview };
    if (!document.contentBase64) return null;
    try { return { kind, text: globalThis.atob(stripDataUrlPrefix(document.contentBase64)) }; } catch { return { kind: "unsupported" }; }
  }
  if (!document.contentBase64) {
    if (!groupId || !document.id || document.hasContent === false) return null;
    const src = getGroupDocumentContentUrl(groupId, document.id);
    if (kind === "image") return { kind, src };
    if (kind === "pdf") return { kind, src };
    return null;
  }
  const base64 = stripDataUrlPrefix(document.contentBase64);
  if (!base64) return null;
  const src = `data:${normalizeDocumentMimeType(document.contentType, document.name)};base64,${base64}`;
  if (kind === "image") return { kind, src };
  if (kind === "pdf") return { kind, src };
  return { kind: "unsupported" };
}

function inferAttachmentKind(attachment: EmailAttachmentEntry): "image" | "pdf" | "office" | "text" | "unsupported" {
  return inferDocumentKind({
    id: makeAttachmentKey(attachment),
    name: attachment.name,
    contentType: attachment.contentType,
    contentBase64: attachment.content,
  } as GroupDocumentEntry);
}

function buildAttachmentPreview(
  attachment: EmailAttachmentEntry | null,
  hydratedContent?: { base64: string; contentType: string; fileName: string } | null,
  textPreview?: string
): PreviewState | null {
  if (!attachment) return null;
  const kind = inferAttachmentKind(attachment);
  if (kind === "text") {
    if (typeof textPreview === "string" && textPreview.trim()) return { kind, text: textPreview };
    const content = stripDataUrlPrefix(attachment.content);
    if (!content) return null;
    try {
      return { kind, text: globalThis.atob(content) };
    } catch {
      return { kind: "unsupported" };
    }
  }

  const base64 = stripDataUrlPrefix(hydratedContent?.base64 || attachment.content);
  if (!base64) return null;
  const src = `data:${normalizeAttachmentMimeType(hydratedContent?.contentType || attachment.contentType, hydratedContent?.fileName || attachment.name)};base64,${base64}`;
  if (kind === "image") return { kind, src };
  if (kind === "pdf") return { kind, src };
  if (kind === "office") return { kind: "unsupported" };
  return { kind: "unsupported" };
}

function buildClassificationStudioParams(email: RelatedEmailEntry | null): Record<string, string> {
  if (!email) return {};
  const params: Record<string, string> = {};
  if (email.itemId) params.itemId = String(email.itemId);
  if (email.internetMessageId) params.internetMessageId = String(email.internetMessageId);
  if (email.conversationId) params.conversationId = String(email.conversationId);
  if (email.subject) params.subject = String(email.subject);
  if (email.fromEmail) params.fromEmail = String(email.fromEmail);
  if (email.fromName) params.fromName = String(email.fromName);
  if (email.receivedAtIso || email.messageDateIso) params.receivedAtIso = String(email.receivedAtIso || email.messageDateIso);
  return params;
}

function resolvePrimaryGroup(
  email: RelatedEmailEntry | null,
  groups: LinkGroupEntry[],
  selectedGroup: LinkGroupEntry | null
): LinkGroupEntry | null {
  if (!email) return selectedGroup;
  const principalRelated = (email.relatedGroups || []).find((group) => String(group.relationKind || "").toLowerCase() === "principal");
  const principalId = String(principalRelated?.id || email.groupId || "").trim();
  if (principalId) {
    const match = groups.find((group) => group.id === principalId);
    if (match) return match;
  }
  if (selectedGroup && (!principalId || selectedGroup.id === principalId)) return selectedGroup;
  return null;
}

function resolveReferenceGroups(
  email: RelatedEmailEntry | null,
  groups: LinkGroupEntry[],
  principalGroupId: string | undefined
): LinkGroupEntry[] {
  if (!email) return [];
  const refs = (email.relatedGroups || []).filter((group) => String(group.relationKind || "").toLowerCase() !== "principal");
  const resolved = refs
    .map((entry) => groups.find((group) => group.id === entry.id) || ({ id: entry.id || "", name: entry.name || entry.id || "Referencia" } as LinkGroupEntry))
    .filter((entry) => entry?.id && entry.id !== principalGroupId);
  return resolved.filter((entry, index, all) => all.findIndex((candidate) => candidate.id === entry.id) === index);
}

function resolveTicketForEmail(email: RelatedEmailEntry | null, tickets: GroupTicketEntry[], principalGroupId: string | undefined): GroupTicketEntry | null {
  if (!email || !tickets.length) return null;
  const emailKey = String(email.emailKey || email.id || "").trim();
  return (
    tickets.find((ticket) => emailKey && String(ticket.createdFromEmailKey || "").trim() === emailKey)
    || tickets.find((ticket) => Boolean(ticket.emailLinked))
    || tickets.find((ticket) => principalGroupId && Array.isArray(ticket.groupIds) && ticket.groupIds.includes(principalGroupId))
    || tickets[0]
    || null
  );
}

function sanitizeEmailHtml(html: string | undefined): string {
  const raw = String(html || "").trim();
  if (!raw) return "";
  return raw
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+=(["']).*?\1/gi, "")
    .replace(/\sjavascript:/gi, " ");
}

function buildEmailPreviewHtml(
  email: RelatedEmailEntry | null,
  attachments: EmailAttachmentEntry[] | OutlookAttachment[] = []
): string {
  const safeHtml = rewriteEmailHtmlInlineImages(sanitizeEmailHtml(email?.bodyHtml), attachments);
  if (safeHtml) {
    return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      :root { color-scheme: light; }
      html, body { margin: 0; padding: 0; background: #ffffff; color: #172b4d; font: 14px/1.45 'Segoe UI', sans-serif; }
      body { padding: 16px; }
      img { max-width: 100%; height: auto; }
      table { max-width: 100%; }
      pre { white-space: pre-wrap; word-break: break-word; }
      blockquote { margin-left: 0; padding-left: 12px; border-left: 3px solid #dbeafe; color: #42526e; }
    </style>
  </head>
  <body>${safeHtml}</body>
</html>`;
  }

  const safeText = String(email?.bodyText || "").trim();
  if (!safeText) return "";
  const escaped = safeText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      html, body { margin: 0; padding: 0; background: #ffffff; color: #172b4d; font: 14px/1.5 'Segoe UI', sans-serif; }
      body { padding: 16px; }
      pre { white-space: pre-wrap; word-break: break-word; font: inherit; margin: 0; }
    </style>
  </head>
  <body><pre>${escaped}</pre></body>
</html>`;
}

function matchesSelectedEmail(document: GroupDocumentEntry, email: RelatedEmailEntry | null): boolean {
  if (!email) return false;
  const keys = new Set([
    String(email.id || "").trim(),
    String(email.itemId || "").trim(),
    normalizeMessageId(email.internetMessageId),
    String(email.conversationId || "").trim(),
    String(email.subject || "").trim(),
  ].filter(Boolean));
  return Boolean(
    (document.sourceEmailKey && keys.has(String(document.sourceEmailKey).trim()))
    || (document.sourceItemId && keys.has(String(document.sourceItemId).trim()))
    || (document.sourceInternetMessageId && keys.has(normalizeMessageId(document.sourceInternetMessageId)))
    || (document.sourceConversationId && keys.has(String(document.sourceConversationId).trim()))
    || (document.sourceEmailSubject && keys.has(String(document.sourceEmailSubject).trim()))
  );
}

function buildEmailHoverText(email: RelatedEmailEntry): string {
  return [
    email.subject ? `Assunto: ${email.subject}` : "",
    email.fromName || email.fromEmail ? `De: ${email.fromName || email.fromEmail}` : "",
    formatDate(email.messageDateIso || email.receivedAtIso) ? `Data: ${formatDate(email.messageDateIso || email.receivedAtIso)}` : "",
    Array.isArray(email.attachments) ? `Anexos: ${email.attachments.length}` : "",
  ].filter(Boolean).join("\n");
}

export default function GroupExplorerApp(): JSX.Element {
  const initial = useMemo(() => readExplorerParams(), []);
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState(initial.groupId);
  const [selectedEmailKey, setSelectedEmailKey] = useState(initial.emailKey);
  const [selectedDocumentId, setSelectedDocumentId] = useState(initial.documentId);
  const [selectedQuickDocumentKey, setSelectedQuickDocumentKey] = useState("");
  const [groupEmails, setGroupEmails] = useState<RelatedEmailEntry[]>([]);
  const [groupDocuments, setGroupDocuments] = useState<GroupDocumentEntry[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);
  const [loadingEmails, setLoadingEmails] = useState(false);
  const [loadingDocuments, setLoadingDocuments] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [emailSearch, setEmailSearch] = useState("");
  const [emailAttachmentFilter, setEmailAttachmentFilter] = useState<EmailAttachmentFilter>("all");
  const [emailSort, setEmailSort] = useState<EmailSortMode>("date_desc");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [documentFilterMode, setDocumentFilterMode] = useState<DocumentFilterMode>("all");
  const [documentSearch, setDocumentSearch] = useState("");
  const [previewMode, setPreviewMode] = useState<PreviewMode>("email");
  const [classificationMode, setClassificationMode] = useState<ClassificationMode>("normal");
  const [classificationEditor, setClassificationEditor] = useState<ClassificationEditor>("summary");
  const [classificationContext, setClassificationContext] = useState<ClassificationContextState | null>(null);
  const [classificationLoading, setClassificationLoading] = useState(false);
  const [liveCurrentContext, setLiveCurrentContext] = useState<OutlookMessageContext | null>(null);
  const [liveCurrentAttachments, setLiveCurrentAttachments] = useState<OutlookAttachment[]>([]);
  const [documentTextPreviewById, setDocumentTextPreviewById] = useState<Record<string, string>>({});
  const [attachmentTextPreviewByKey, setAttachmentTextPreviewByKey] = useState<Record<string, string>>({});
  const [attachmentPreviewDataByKey, setAttachmentPreviewDataByKey] = useState<Record<string, { base64: string; contentType: string; fileName: string }>>({});

  useEffect(() => {
    (async () => {
      try {
        const settings = await getSettings();
        if (settings.skinId) applySkin(settings.skinId);
      } catch {}
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const [ctx, attachments] = await Promise.all([
          getSelectedMessageContext().catch(() => ({} as OutlookMessageContext)),
          getAttachments().catch(() => [] as OutlookAttachment[]),
        ]);
        if (cancelled) return;
        setLiveCurrentContext(ctx || null);
        setLiveCurrentAttachments(Array.isArray(attachments) ? attachments : []);
      } catch {
        if (cancelled) return;
        setLiveCurrentContext(null);
        setLiveCurrentAttachments([]);
      }
    })();
    return () => { cancelled = true; };
  }, [selectedEmailKey, selectedGroupId]);

  useEffect(() => {
    let cancelled = false;
    setLoadingGroups(true);
    setError(null);
    listLinkGroups("/")
      .then((nextGroups) => {
        if (cancelled) return;
        setGroups(nextGroups);
        setSelectedGroupId((current) => {
          if (current && nextGroups.some((group) => group.id === current)) return current;
          if (initial.groupId && nextGroups.some((group) => group.id === initial.groupId)) return initial.groupId;
          return nextGroups[0]?.id || "";
        });
      })
      .catch((nextError: any) => { if (!cancelled) setError(nextError?.message || "Nao foi possivel carregar os grupos."); })
      .finally(() => { if (!cancelled) setLoadingGroups(false); });
    return () => { cancelled = true; };
  }, [initial.groupId]);

  useEffect(() => {
    if (!selectedGroupId) {
      setGroupEmails([]);
      setGroupDocuments([]);
      return;
    }
    let cancelled = false;
    setLoadingEmails(true);
    setLoadingDocuments(true);
    setError(null);
    Promise.all([getGroupEmails(selectedGroupId), getGroupDocuments(selectedGroupId)])
      .then(([emails, documents]) => {
        if (cancelled) return;
        setGroupEmails(emails);
        setGroupDocuments(documents);
        setSelectedEmailKey((current) => current && emails.some((email) => makeEmailKey(email) === current) ? current : (emails[0] ? makeEmailKey(emails[0]) : ""));
        setSelectedDocumentId((current) => {
          if (current && documents.some((document) => makeDocumentKey(document) === current)) return current;
          if (initial.documentId && documents.some((document) => makeDocumentKey(document) === initial.documentId)) return initial.documentId;
          return "";
        });
      })
      .catch((nextError: any) => { if (!cancelled) setError(nextError?.message || "Nao foi possivel carregar os dados do grupo."); })
      .finally(() => {
        if (cancelled) return;
        setLoadingEmails(false);
        setLoadingDocuments(false);
      });
    return () => { cancelled = true; };
  }, [initial.documentId, initial.emailKey, selectedGroupId]);

  const selectedGroup = useMemo(() => groups.find((group) => group.id === selectedGroupId) || null, [groups, selectedGroupId]);

  const filteredEmails = useMemo(() => {
    const query = String(emailSearch || "").trim().toLowerCase();
    const fromTime = dateFrom ? new Date(`${dateFrom}T00:00:00`).getTime() : 0;
    const toTime = dateTo ? new Date(`${dateTo}T23:59:59`).getTime() : 0;
    const next = groupEmails.filter((email) => {
      if (query) {
        const haystack = [email.subject, email.fromName, email.fromEmail].map((value) => String(value || "").toLowerCase()).join(" ");
        if (!haystack.includes(query)) return false;
      }
      if (emailAttachmentFilter === "with" && !emailHasAttachments(email)) return false;
      if (emailAttachmentFilter === "without" && emailHasAttachments(email)) return false;
      if (fromTime || toTime) {
        const timestamp = getEmailTimestamp(email);
        if (fromTime && (!timestamp || timestamp < fromTime)) return false;
        if (toTime && (!timestamp || timestamp > toTime)) return false;
      }
      return true;
    });
    next.sort((a, b) => {
      if (emailSort === "date_asc") return getEmailTimestamp(a) - getEmailTimestamp(b);
      if (emailSort === "subject_asc") return String(a.subject || "").localeCompare(String(b.subject || ""), "pt");
      if (emailSort === "subject_desc") return String(b.subject || "").localeCompare(String(a.subject || ""), "pt");
      return getEmailTimestamp(b) - getEmailTimestamp(a);
    });
    return next;
  }, [dateFrom, dateTo, emailAttachmentFilter, emailSearch, emailSort, groupEmails]);

  const selectedEmail = useMemo(
    () => filteredEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || groupEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || filteredEmails[0] || groupEmails[0] || null,
    [filteredEmails, groupEmails, selectedEmailKey]
  );

  const selectedEmailDocuments = useMemo(() => {
    const query = String(documentSearch || "").trim().toLowerCase();
    return groupDocuments.filter((document) => {
      if (!matchesSelectedEmail(document, selectedEmail)) return false;
      if (documentFilterMode === "selected_email" && !matchesSelectedEmail(document, selectedEmail)) return false;
      if (!query) return true;
      const haystack = [document.name, document.sourceEmailSubject, document.contentType].map((value) => String(value || "").toLowerCase()).join(" ");
      return haystack.includes(query);
    });
  }, [documentFilterMode, documentSearch, groupDocuments, selectedEmail]);

  const selectedEmailAttachments = useMemo(() => {
    const persisted = Array.isArray(selectedEmail?.attachments) ? selectedEmail.attachments : [];
    if (!selectedEmail || !emailMatchesCurrentContext(selectedEmail, liveCurrentContext)) {
      return mergeEmailAttachments(persisted, []);
    }
    return mergeEmailAttachments(persisted, liveCurrentAttachments);
  }, [liveCurrentAttachments, liveCurrentContext, selectedEmail]);

  const quickDocuments = useMemo<QuickDocumentEntry[]>(() => {
    const attachmentEntries = selectedEmailAttachments.map((attachment) => ({
      key: `attachment:${makeAttachmentKey(attachment)}`,
      title: attachment.name,
      meta: [normalizeAttachmentMimeType(attachment.contentType, attachment.name) || "Anexo", formatBytes(attachment.size)].filter(Boolean).join(" - "),
      kind: "attachment" as const,
      attachment,
    }));
    const documentEntries = selectedEmailDocuments.map((document) => ({
      key: `document:${makeDocumentKey(document)}`,
      title: document.name,
      meta: [normalizeDocumentMimeType(document.contentType, document.name) || "Documento", formatBytes(document.size)].filter(Boolean).join(" - "),
      kind: "document" as const,
      document,
    }));
    return [...attachmentEntries, ...documentEntries];
  }, [selectedEmailAttachments, selectedEmailDocuments]);

  const selectedQuickDocument = useMemo(
    () => quickDocuments.find((entry) => entry.key === selectedQuickDocumentKey) || quickDocuments[0] || null,
    [quickDocuments, selectedQuickDocumentKey]
  );

  const selectedAttachment = useMemo(
    () => (selectedQuickDocument?.kind === "attachment" ? selectedQuickDocument.attachment || null : null),
    [selectedQuickDocument]
  );

  const selectedDocument = useMemo(() => {
    if (selectedQuickDocument?.kind === "document") return selectedQuickDocument.document || null;
    return selectedEmailDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || null;
  }, [selectedDocumentId, selectedEmailDocuments, selectedQuickDocument]);

  const selectedDocumentTextPreview = useMemo(
    () => (selectedDocument?.id ? documentTextPreviewById[selectedDocument.id] : ""),
    [documentTextPreviewById, selectedDocument?.id]
  );
  const selectedDocumentPreview = useMemo(
    () => buildDocumentPreview(selectedGroupId, selectedDocument, selectedDocumentTextPreview),
    [selectedDocument, selectedDocumentTextPreview, selectedGroupId]
  );
  const selectedAttachmentKey = useMemo(() => (selectedAttachment ? makeAttachmentKey(selectedAttachment) : ""), [selectedAttachment]);
  const selectedAttachmentTextPreview = useMemo(
    () => (selectedAttachmentKey ? attachmentTextPreviewByKey[selectedAttachmentKey] : ""),
    [attachmentTextPreviewByKey, selectedAttachmentKey]
  );
  const selectedAttachmentPreview = useMemo(
    () => buildAttachmentPreview(selectedAttachment, selectedAttachmentKey ? attachmentPreviewDataByKey[selectedAttachmentKey] : null, selectedAttachmentTextPreview),
    [attachmentPreviewDataByKey, selectedAttachment, selectedAttachmentKey, selectedAttachmentTextPreview]
  );
  const selectedEmailPreviewHtml = useMemo(
    () => buildEmailPreviewHtml(selectedEmail, selectedEmailAttachments),
    [selectedEmail, selectedEmailAttachments]
  );
  const selectedEmailHasPreview = Boolean(String(selectedEmail?.bodyHtml || "").trim() || String(selectedEmail?.bodyText || "").trim());
  const selectedProvider = useMemo(() => providerLabel(selectedDocument?.storageProvider || groupDocuments[0]?.storageProvider), [groupDocuments, selectedDocument]);
  const selectedEmailContextPayload = useMemo(
    () => ({
      itemId: String(selectedEmail?.itemId || "").trim(),
      internetMessageId: String(selectedEmail?.internetMessageId || "").trim(),
      conversationId: String(selectedEmail?.conversationId || "").trim(),
      subject: String(selectedEmail?.subject || "").trim(),
      fromEmail: String(selectedEmail?.fromEmail || "").trim(),
      receivedAtIso: String(selectedEmail?.receivedAtIso || selectedEmail?.messageDateIso || "").trim(),
    }),
    [selectedEmail]
  );

  useEffect(() => {
    if (!filteredEmails.some((email) => makeEmailKey(email) === selectedEmailKey)) {
      setSelectedEmailKey(filteredEmails[0] ? makeEmailKey(filteredEmails[0]) : "");
    }
  }, [filteredEmails, selectedEmailKey]);

  useEffect(() => {
    if (!quickDocuments.some((entry) => entry.key === selectedQuickDocumentKey)) {
      setSelectedQuickDocumentKey(quickDocuments[0]?.key || "");
    }
  }, [quickDocuments, selectedQuickDocumentKey]);

  useEffect(() => {
    setClassificationEditor("summary");
    setPreviewMode("email");
  }, [selectedEmailKey]);

  useEffect(() => {
    const documentId = String(selectedDocument?.id || "").trim();
    if (!selectedGroupId || !documentId || !selectedDocument) return;
    if (inferDocumentKind(selectedDocument) !== "text") return;
    if (selectedDocument.contentBase64 || documentTextPreviewById[documentId]) return;
    if (selectedDocument.hasContent === false) return;

    let cancelled = false;
    void getGroupDocumentTextContent(selectedGroupId, documentId)
      .then((text) => {
        if (cancelled) return;
        setDocumentTextPreviewById((current) => (
          current[documentId] === text ? current : { ...current, [documentId]: text }
        ));
      })
      .catch(() => {
        if (cancelled) return;
        setDocumentTextPreviewById((current) => (
          Object.prototype.hasOwnProperty.call(current, documentId)
            ? current
            : { ...current, [documentId]: "" }
        ));
      });

    return () => {
      cancelled = true;
    };
  }, [documentTextPreviewById, selectedDocument, selectedGroupId]);

  useEffect(() => {
    const attachment = selectedAttachment;
    const emailId = String(selectedEmail?.id || "").trim();
    const attachmentKey = selectedAttachmentKey;
    if (!attachment || !emailId || !attachmentKey) return;
    if (attachment.content) return;
    if ((attachment as any).hasContent === false) return;

    const kind = inferAttachmentKind(attachment);
    if (kind === "text") {
      if (Object.prototype.hasOwnProperty.call(attachmentTextPreviewByKey, attachmentKey)) return;
      let cancelled = false;
      void getEmailAttachmentTextContent(emailId, attachmentKey)
        .then((text) => {
          if (cancelled) return;
          setAttachmentTextPreviewByKey((current) => current[attachmentKey] === text ? current : { ...current, [attachmentKey]: text });
        })
        .catch(() => {
          if (cancelled) return;
          setAttachmentTextPreviewByKey((current) => Object.prototype.hasOwnProperty.call(current, attachmentKey) ? current : { ...current, [attachmentKey]: "" });
        });
      return () => {
        cancelled = true;
      };
    }

    if (attachmentPreviewDataByKey[attachmentKey]) return;
    let cancelled = false;
    void getEmailAttachmentContentBase64(emailId, attachmentKey)
      .then((payload) => {
        if (cancelled) return;
        setAttachmentPreviewDataByKey((current) => current[attachmentKey]?.base64 === payload.base64 ? current : { ...current, [attachmentKey]: payload });
      })
      .catch(() => {
        if (cancelled) return;
      });
    return () => {
      cancelled = true;
    };
  }, [attachmentPreviewDataByKey, attachmentTextPreviewByKey, selectedAttachment, selectedAttachmentKey, selectedEmail?.id]);

  useEffect(() => {
    const hasIdentity = Object.values(selectedEmailContextPayload).some(Boolean);
    if (!hasIdentity) {
      setClassificationContext(null);
      return;
    }
    let cancelled = false;
    setClassificationLoading(true);
    void getRelatedEmailContext(selectedEmailContextPayload)
      .then((nextContext) => {
        if (cancelled) return;
        setClassificationContext(nextContext);
      })
      .catch(() => {
        if (cancelled) return;
        setClassificationContext(null);
      })
      .finally(() => {
        if (!cancelled) setClassificationLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [selectedEmailContextPayload]);

  const classificationEmail = useMemo(
    () => classificationContext?.email || selectedEmail,
    [classificationContext?.email, selectedEmail]
  );
  const classificationGroups = useMemo(() => {
    const merged = new Map<string, LinkGroupEntry>();
    [selectedGroup, ...groups, ...(classificationContext?.groups || [])].filter(Boolean).forEach((group) => {
      if (group?.id) merged.set(group.id, group);
    });
    return Array.from(merged.values());
  }, [classificationContext?.groups, groups, selectedGroup]);
  const classificationPrincipalGroup = useMemo(
    () => resolvePrimaryGroup(classificationEmail, classificationGroups, selectedGroup),
    [classificationEmail, classificationGroups, selectedGroup]
  );
  const classificationReferenceGroups = useMemo(
    () => resolveReferenceGroups(classificationEmail, classificationGroups, classificationPrincipalGroup?.id),
    [classificationEmail, classificationGroups, classificationPrincipalGroup?.id]
  );
  const classificationTicket = useMemo(
    () => resolveTicketForEmail(classificationEmail, classificationContext?.tickets || [], classificationPrincipalGroup?.id),
    [classificationContext?.tickets, classificationEmail, classificationPrincipalGroup?.id]
  );
  const classificationLabels = useMemo(
    () => Array.from(new Set((classificationEmail?.labels || []).map((label) => String(label || "").trim()).filter(Boolean))).sort((left, right) => left.localeCompare(right, "pt")),
    [classificationEmail?.labels]
  );
  const classificationAdvancedSummary = useMemo(() => {
    const meta = classificationEmail?.classificationMeta;
    const values: string[] = [];
    if (meta?.principalCategorize) values.push("Grupo principal em categoria");
    if (meta?.principalStatusEnabled) values.push("Estado do grupo ativo");
    if (meta?.referenceCategorize) values.push("Referencias em categoria");
    if (meta?.referenceStatusEnabled) values.push("Estado de referencias ativo");
    if (meta?.ticketStatusEnabled) values.push("Estado de ticket ativo");
    if (Array.isArray(meta?.categorizedLabelNames) && meta.categorizedLabelNames.length) values.push(`${meta.categorizedLabelNames.length} etiqueta(s) categorizada(s)`);
    return values;
  }, [classificationEmail?.classificationMeta]);

  async function refreshCurrentGroup() {
    if (!selectedGroupId) return;
    setLoadingEmails(true);
    setLoadingDocuments(true);
    setError(null);
    try {
      const [emails, documents] = await Promise.all([getGroupEmails(selectedGroupId), getGroupDocuments(selectedGroupId)]);
      setGroupEmails(emails);
      setGroupDocuments(documents);
      setNotice("Explorador atualizado.");
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel atualizar o explorador.");
    } finally {
      setLoadingEmails(false);
      setLoadingDocuments(false);
    }
  }

  async function handleOpenGroupSettings() {
    if (!selectedGroupId) return;
    try {
      await openGroupSettings({ groupId: selectedGroupId });
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel abrir as definicoes do grupo.");
    }
  }

  async function handleOpenClassificationStudio() {
    if (!selectedEmail) {
      setNotice("Seleciona um email para abrir o editor completo.");
      return;
    }
    try {
      await openGroupClassificationStudio(buildClassificationStudioParams(selectedEmail));
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel abrir o editor completo do caso.");
    }
  }

  function handleOpenQuickDocument(entry: QuickDocumentEntry) {
    setSelectedQuickDocumentKey(entry.key);
    if (entry.kind === "document" && entry.document) {
      setSelectedDocumentId(makeDocumentKey(entry.document));
    }
    setPreviewMode("document");
  }

  async function handleOpenEmail(email: RelatedEmailEntry) {
    const opened = await openLinkedOutlookEmail({ itemId: email.itemId, emailWebLink: email.emailWebLink });
    if (!opened) setNotice("Este email ainda nao tem abertura direta disponivel.");
  }

  async function handleReplyEmail(email: RelatedEmailEntry | null) {
    if (!email) return;
    if (emailMatchesCurrentContext(email, liveCurrentContext)) {
      const handled = await requestCockpitHostAction({ type: "reply-current" });
      if (handled) setNotice("Formulario de resposta aberto para o email atual.");
      else setError("Nao foi possivel abrir a resposta.");
      return;
    }

    const opened = await requestCockpitHostAction({ type: "open-email", itemId: email.itemId, emailWebLink: email.emailWebLink });
    if (opened) {
      setNotice("Email aberto no Outlook. Usa Responder no Outlook para continuar.");
    } else {
      setNotice("Este email ainda nao tem abertura direta para responder.");
    }
  }

  async function handleForwardEmail(email: RelatedEmailEntry | null) {
    if (!email) return;
    if (emailMatchesCurrentContext(email, liveCurrentContext)) {
      const handled = await requestCockpitHostAction({ type: "forward-current" });
      if (handled) setNotice("Formulario de reencaminhamento aberto para o email atual.");
      else setError("Nao foi possivel abrir o reencaminhamento.");
      return;
    }

    const opened = await requestCockpitHostAction({ type: "open-email", itemId: email.itemId, emailWebLink: email.emailWebLink });
    if (opened) {
      setNotice("Email aberto no Outlook. Usa Reencaminhar no Outlook para continuar.");
    } else {
      setNotice("Este email ainda nao tem abertura direta para reencaminhar.");
    }
  }

  async function handleRemoveEmail(email: RelatedEmailEntry) {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      const persistentEmailKey = String(email.id || "").startsWith("email_") ? undefined : String(email.id || "").trim() || undefined;
      await removeEmailFromLinkGroup(selectedGroup.id, {
        emailKey: persistentEmailKey,
        itemId: email.itemId,
        internetMessageId: email.internetMessageId,
        conversationId: email.conversationId,
        subject: email.subject,
        fromEmail: email.fromEmail,
        receivedAtIso: email.receivedAtIso || email.messageDateIso,
      });
      await refreshCurrentGroup();
      setNotice("Email removido do grupo.");
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel remover o email do grupo.");
    } finally {
      setBusy(false);
    }
  }

  const caseClient = formatChipValue(selectedGroup?.contacts?.[0]?.company || selectedGroup?.contacts?.[0]?.name, "Sem cliente");
  const caseBrand = formatChipValue(selectedGroup?.entities?.[0]?.name, "Sem marca");
  const caseState = formatStatusLabel(selectedGroup?.status);
  const activePreviewState = selectedQuickDocument?.kind === "attachment" ? selectedAttachmentPreview : selectedDocumentPreview;
  const activePreviewTitle = selectedQuickDocument?.title || selectedDocument?.name || "Documento";
  const previewHasDocument = Boolean(selectedQuickDocument);
  const classificationItems = [
    {
      key: "principal" as const,
      title: "Grupo principal",
      value: classificationPrincipalGroup?.name || "Sem grupo principal",
      description: classificationEmail?.classificationMeta?.principalStatusEnabled ? formatStatusLabel(classificationPrincipalGroup?.status) : "Sem estado ativo",
    },
    {
      key: "labels" as const,
      title: "Etiquetas",
      value: classificationLabels.length ? classificationLabels.join(", ") : "Sem etiquetas",
      description: classificationLabels.length ? `${classificationLabels.length} atribuida(s)` : "Sem atribuicoes estruturadas",
    },
    {
      key: "ticket" as const,
      title: "Ticket",
      value: classificationTicket ? `${classificationTicket.code} - ${classificationTicket.title}` : "Sem ticket",
      description: classificationTicket ? formatStatusLabel(classificationTicket.status) : "Sem seguimento ligado",
    },
    {
      key: "references" as const,
      title: "Referencias",
      value: classificationReferenceGroups.length ? classificationReferenceGroups.map((entry) => entry.name || entry.id).join(", ") : "Sem referencias",
      description: classificationReferenceGroups.length ? `${classificationReferenceGroups.length} referencia(s)` : "Disponivel no modo avancado",
    },
  ];

  return (
    <div style={styles.root}>
      <header style={styles.headerShell}>
        <div style={styles.headerIdentity}>
          <div style={styles.eyebrow}>Explorador de caso</div>
          <div style={styles.caseTitleRow}>
            <div style={styles.title}>{selectedGroup?.name || "Grupo"}</div>
            <div style={styles.caseChips}>
              <span style={styles.caseChip}>Cliente: {caseClient}</span>
              <span style={styles.caseChip}>Marca: {caseBrand}</span>
              <span style={styles.caseChip}>Estado: {caseState}</span>
            </div>
          </div>
          <div style={styles.headerSelectorRow}>
            <div style={{ ...styles.selectWrap, minWidth: 220 }}>
              <select style={styles.select} value={selectedGroupId} onChange={(event) => { setSelectedGroupId(event.target.value); setSelectedEmailKey(""); setSelectedDocumentId(""); setSelectedQuickDocumentKey(""); setNotice(null); }}>
                {groups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}
              </select>
            </div>
            <button type="button" style={styles.ghostBtn} onClick={() => void handleOpenGroupSettings()} disabled={!selectedGroupId}>Renomear</button>
            <button type="button" style={styles.ghostBtn} onClick={() => setNotice("Fluxo de fusao preparado para fase seguinte.")} disabled={!selectedGroupId}>Fundir</button>
            <button type="button" style={styles.primaryBtn} onClick={() => void handleOpenClassificationStudio()} disabled={!selectedEmail}>Guardar</button>
            <button type="button" style={styles.iconBtn} onClick={() => void refreshCurrentGroup()} disabled={loadingGroups || loadingEmails || loadingDocuments} title="Atualizar"><Icons.RefreshCw size={12} /></button>
            <button type="button" style={styles.closeBtn} onClick={closeExplorer}>Fechar</button>
          </div>
        </div>
        <div style={styles.headerMetrics}>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Emails</span><span style={styles.metricValue}>{groupEmails.length}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Rapidos</span><span style={styles.metricValue}>{quickDocuments.length}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Docs guardados</span><span style={styles.metricValue}>{groupDocuments.length}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Provider</span><span style={styles.metricValue}>{selectedProvider}</span></div>
        </div>
      </header>

      <div style={styles.statusStack}>
        {error ? <PanelState compact tone="error" title="Falha no explorador" description={error} /> : null}
        {notice ? <PanelState compact tone="info" title="Explorador" description={notice} /> : null}
        {loadingGroups ? <PanelState compact tone="loading" title="A carregar grupos" description="Estamos a preparar o explorador documental." /> : null}
        {!loadingGroups && !selectedGroup ? <PanelState compact tone="info" title="Sem grupos" description="Ainda nao existem grupos manuais disponiveis para este explorador." /> : null}
      </div>
      {selectedGroup ? (
        <div style={styles.explorerBody}>
          <section style={styles.topCardsGrid}>
            <section style={styles.panelCompact}>
              <div style={styles.sectionHeaderCompact}>
                <div>
                  <div style={styles.sectionTitle}>Emails</div>
                  <div style={styles.sectionSubtitle}>Exploracao do caso</div>
                </div>
              </div>
              <label style={styles.filterFieldWide}>
                <span style={styles.filterLabel}>Pesquisar</span>
                <input style={styles.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Assunto, contacto ou email..." />
              </label>
              <div style={styles.listShellEmails}>
                {loadingEmails && !groupEmails.length ? <PanelState compact tone="loading" title="A carregar emails" description="A listar os emails do grupo." /> : null}
                {!loadingEmails && !filteredEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Nao ha emails a corresponder aos filtros atuais." /> : null}
                {filteredEmails.map((email) => {
                  const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                  const canOpen = Boolean(email.itemId || email.emailWebLink);
                  const attachmentCount = emailHasAttachments(email) ? (email.attachments?.length || 0) : 0;
                  return (
                    <div key={makeEmailKey(email)} style={active ? styles.cardLineActive : styles.cardLine}>
                      <button
                        type="button"
                        style={styles.cardMainExpanded}
                        onClick={() => {
                          setSelectedEmailKey(makeEmailKey(email));
                          setPreviewMode("email");
                        }}
                        title={buildEmailHoverText(email)}
                      >
                        <div style={styles.cardTitle}>{email.subject || "(sem assunto)"}</div>
                        <div style={styles.cardMetaLine}>{[email.fromName || email.fromEmail || "Sem remetente", formatDate(email.messageDateIso || email.receivedAtIso) || "Sem data"].filter(Boolean).join(" - ")}</div>
                        <div style={styles.cardBadgeRow}>
                          <span style={styles.metaTag}>{attachmentCount} anexo(s)</span>
                        </div>
                      </button>
                      <div style={styles.cardActions}>
                        <button type="button" style={styles.iconBtn} onClick={() => void handleOpenEmail(email)} disabled={!canOpen} title={canOpen ? "Abrir email" : "Sem abertura direta"}><Icons.MessageSquare size={10} /></button>
                        <button type="button" style={styles.iconBtnDanger} onClick={() => void handleRemoveEmail(email)} disabled={busy} title="Remover do grupo"><Icons.Trash size={10} /></button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>

            <section style={styles.panelCompact}>
              <div style={styles.sectionHeaderCompact}>
                <div>
                  <div style={styles.sectionTitle}>Documentos rapidos</div>
                  <div style={styles.sectionSubtitle}>Do email selecionado</div>
                </div>
              </div>
              <div style={styles.listShellDocumentsCompact}>
                {loadingDocuments && !quickDocuments.length ? <PanelState compact tone="loading" title="A carregar documentos" description="A preparar os documentos do email selecionado." /> : null}
                {!loadingDocuments && !quickDocuments.length ? <PanelState compact tone="info" title="Sem documentos rapidos" description="Este email ainda nao tem anexos ou documentos associados para abrir aqui." /> : null}
                {quickDocuments.map((entry) => {
                  const active = entry.key === selectedQuickDocument?.key;
                  return (
                    <div key={entry.key} style={active ? styles.cardLineActive : styles.cardLine}>
                      <div style={styles.cardMainExpanded}>
                        <div style={styles.cardTitle}>{entry.title}</div>
                        <div style={styles.cardMetaLine}>{entry.meta || (entry.kind === "attachment" ? "Anexo" : "Documento guardado")}</div>
                      </div>
                      <button type="button" style={styles.inlineActionBtn} onClick={() => handleOpenQuickDocument(entry)}>Abrir</button>
                    </div>
                  );
                })}
              </div>
            </section>

            <section style={styles.panelCompact}>
              <div style={styles.sectionHeaderCompact}>
                <div>
                  <div style={styles.sectionTitle}>Classificacao</div>
                  <div style={styles.sectionSubtitle}>Compacta e contextual</div>
                </div>
                <div style={styles.segmentedControl}>
                  <button type="button" style={classificationMode === "normal" ? styles.segmentBtnActive : styles.segmentBtn} onClick={() => setClassificationMode("normal")}>Normal</button>
                  <button type="button" style={classificationMode === "advanced" ? styles.segmentBtnActive : styles.segmentBtn} onClick={() => setClassificationMode("advanced")}>Avancado</button>
                </div>
              </div>
              {classificationEditor === "summary" ? (
                <div style={styles.classificationSummary}>
                  {classificationLoading ? <PanelState compact tone="loading" title="A preparar classificacao" description="A carregar o estado final do email selecionado." /> : null}
                  {!classificationLoading ? classificationItems.filter((item) => classificationMode === "advanced" || item.key !== "references").map((item) => (
                    <button key={item.key} type="button" style={styles.classificationTile} onClick={() => setClassificationEditor(item.key)}>
                      <span style={styles.classificationTileLabel}>{item.title}</span>
                      <strong style={styles.classificationTileValue}>{item.value}</strong>
                      <span style={styles.classificationTileMeta}>{item.description}</span>
                    </button>
                  )) : null}
                  {classificationMode === "advanced" && classificationAdvancedSummary.length ? (
                    <div style={styles.advancedHintBox}>
                      {classificationAdvancedSummary.map((item) => <span key={item} style={styles.advancedHintChip}>{item}</span>)}
                    </div>
                  ) : null}
                </div>
              ) : (
                <div style={styles.classificationEditor}>
                  <button type="button" style={styles.backBtn} onClick={() => setClassificationEditor("summary")}>Voltar</button>
                  <div style={styles.classificationEditorTitle}>
                    {classificationEditor === "principal" ? "Grupo principal" : classificationEditor === "labels" ? "Etiquetas" : classificationEditor === "ticket" ? "Ticket" : "Referencias"}
                  </div>
                  {classificationEditor === "principal" ? <div style={styles.editorLead}>{classificationPrincipalGroup?.name || "Sem grupo principal ligado"}</div> : null}
                  {classificationEditor === "labels" ? <div style={styles.editorLead}>{classificationLabels.length ? classificationLabels.join(", ") : "Sem etiquetas"}</div> : null}
                  {classificationEditor === "ticket" ? <div style={styles.editorLead}>{classificationTicket ? `${classificationTicket.code} - ${classificationTicket.title}` : "Sem ticket ligado"}</div> : null}
                  {classificationEditor === "references" ? <div style={styles.editorLead}>{classificationReferenceGroups.length ? `${classificationReferenceGroups.length} referencia(s)` : "Sem referencias ligadas"}</div> : null}
                  <div style={styles.editorBodyText}>
                    {classificationEditor === "principal" ? "Mostramos o grupo principal atual e deixamos o editor completo preparado para a fase seguinte." : null}
                    {classificationEditor === "labels" ? "Modo compacto: aqui vemos apenas as etiquetas atribuidas no caso atual." : null}
                    {classificationEditor === "ticket" ? `Estado atual: ${classificationTicket ? formatStatusLabel(classificationTicket.status) : "Nao definido"}` : null}
                    {classificationEditor === "references" ? "As referencias e opcoes adicionais continuam acessiveis no modo avancado e no editor completo." : null}
                  </div>
                  {classificationEditor === "references" && classificationReferenceGroups.length ? (
                    <div style={styles.inlineChips}>
                      {classificationReferenceGroups.map((group) => <span key={group.id} style={styles.metaTag}>{group.name || group.id}</span>)}
                    </div>
                  ) : null}
                  <button type="button" style={styles.primaryBtnWide} onClick={() => void handleOpenClassificationStudio()} disabled={!selectedEmail}>
                    Abrir editor completo
                  </button>
                </div>
              )}
            </section>
          </section>

          <section style={styles.previewPanel}>
            <div style={styles.sectionHeaderCompact}>
              <div style={styles.sectionTitle}>Preview</div>
              <div style={styles.previewTabs}>
                <button type="button" style={previewMode === "email" ? styles.previewTabActive : styles.previewTab} onClick={() => setPreviewMode("email")} disabled={!selectedEmailHasPreview}>Email</button>
                <button type="button" style={previewMode === "document" ? styles.previewTabActive : styles.previewTab} onClick={() => setPreviewMode("document")} disabled={!previewHasDocument}>Documento</button>
                <button type="button" style={previewMode === "reply" ? styles.previewTabActive : styles.previewTab} onClick={() => setPreviewMode("reply")} disabled={!selectedEmail}>Responder</button>
                <button type="button" style={previewMode === "forward" ? styles.previewTabActive : styles.previewTab} onClick={() => setPreviewMode("forward")} disabled={!selectedEmail}>Reencaminhar</button>
              </div>
            </div>
            {previewMode === "email" ? (
              selectedEmailHasPreview ? (
                <div style={styles.previewFrame}>
                  <iframe
                    title={selectedEmail?.subject || "Preview do email"}
                    srcDoc={selectedEmailPreviewHtml}
                    sandbox=""
                    style={styles.previewIframe}
                  />
                </div>
              ) : (
                <PanelState compact tone="info" title="Preview do email indisponivel" description="Este email ainda nao tem corpo HTML ou texto guardado para preview." />
              )
            ) : previewMode === "document" && !previewHasDocument ? (
              <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe um documento rapido para abrir o preview." />
            ) : previewMode === "document" && activePreviewState?.kind === "image" ? (
              <div style={styles.previewFrame}><img src={activePreviewState.src} alt={activePreviewTitle} style={styles.previewImage} /></div>
            ) : previewMode === "document" && activePreviewState?.kind === "pdf" ? (
              <div style={styles.previewFrame}>
                {activePreviewState.src.startsWith("data:")
                  ? <PdfPreview title={activePreviewTitle} dataUrl={activePreviewState.src} />
                  : <iframe title={activePreviewTitle} src={activePreviewState.src} style={styles.previewIframe} />}
              </div>
            ) : previewMode === "document" && activePreviewState?.kind === "office" ? (
              <div style={styles.previewFrame}>
                <iframe
                  title={activePreviewTitle}
                  src={activePreviewState.url}
                  style={styles.previewIframe}
                />
              </div>
            ) : previewMode === "document" && activePreviewState?.kind === "text" ? (
              <pre style={styles.previewText}>{activePreviewState.text}</pre>
            ) : previewMode === "reply" ? (
              <div style={styles.replyComposerShell}>
                <div style={styles.replyComposerLead}>Resposta preparada para o email selecionado</div>
                <div style={styles.replyComposerMeta}>Estrutura pronta para futura integracao com editor, IA e selecao de anexos.</div>
                <div style={styles.replyComposerActions}>
                  <button type="button" style={styles.primaryBtnWide} onClick={() => void handleReplyEmail(selectedEmail)}>Abrir resposta no Outlook</button>
                </div>
              </div>
            ) : previewMode === "forward" ? (
              <div style={styles.replyComposerShell}>
                <div style={styles.replyComposerLead}>Reencaminhamento preparado para o email selecionado</div>
                <div style={styles.replyComposerMeta}>A vista fica pronta para receber o editor e a configuracao de anexos numa fase seguinte.</div>
                <div style={styles.replyComposerActions}>
                  <button type="button" style={styles.primaryBtnWide} onClick={() => void handleForwardEmail(selectedEmail)}>Abrir reencaminhamento no Outlook</button>
                </div>
              </div>
            ) : (
              <PanelState compact tone="info" title="Preview nao disponivel" description="Este documento pode ser descarregado ou anexado. Alguns formatos Office podem exigir URL publica para preview." />
            )}
          </section>
        </div>
      ) : null}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: {
    height: "calc(100vh - 24px)",
    background: "var(--iccc-bg, #edf2f7)",
    color: "var(--iccc-text, #172b4d)",
    fontFamily: "var(--iccc-font, 'Segoe UI', sans-serif)",
    padding: 12,
    display: "grid",
    gridTemplateRows: "auto auto minmax(0, 1fr)",
    gap: 10,
    overflow: "hidden",
    boxSizing: "border-box",
  },
  statusStack: { display: "grid", gap: 8, minHeight: 0, alignContent: "start" },
  explorerBody: { display: "grid", gridTemplateRows: "minmax(310px, 0.85fr) minmax(320px, 1.15fr)", gap: 10, minHeight: 0, overflow: "hidden" },
  headerShell: { display: "grid", gridTemplateColumns: "minmax(0, 1.35fr) minmax(260px, 0.7fr)", gap: 10, padding: 14, borderRadius: 16, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.94)", boxShadow: "0 12px 30px rgba(15, 23, 42, 0.08)", alignItems: "stretch" },
  headerIdentity: { display: "grid", gap: 10, minWidth: 0, alignContent: "space-between" },
  eyebrow: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#5b6b83" },
  caseTitleRow: { display: "grid", gap: 8, minWidth: 0 },
  title: { fontSize: 22, fontWeight: 800, color: "#0f172a", lineHeight: 1.1 },
  caseChips: { display: "flex", gap: 8, flexWrap: "wrap" },
  caseChip: { fontSize: 11, fontWeight: 700, color: "#334155", background: "rgba(15, 23, 42, 0.05)", border: "1px solid rgba(15, 23, 42, 0.08)", borderRadius: 999, padding: "4px 10px" },
  headerSelectorRow: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  headerMetrics: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 8, alignContent: "stretch", minHeight: 0 },
  metricMini: { display: "grid", gap: 2, padding: 10, borderRadius: 12, background: "rgba(15, 23, 42, 0.03)", minWidth: 0 },
  metricLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.06em", color: "#6b7280" },
  metricValue: { fontSize: 12, fontWeight: 700, color: "#0f172a", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  selectWrap: { minWidth: 0 },
  select: { width: "100%", borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "rgba(248,250,252,0.95)", color: "#172b4d", padding: "9px 14px", fontSize: 12, fontWeight: 600, outline: "none" },
  ghostBtn: { borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#fff", color: "#0f172a", padding: "8px 12px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  primaryBtn: { borderRadius: 999, border: "none", background: "#1d4ed8", color: "#fff", padding: "8px 14px", fontSize: 11, fontWeight: 800, cursor: "pointer" },
  closeBtn: { borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", color: "#0f172a", padding: "8px 14px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  topCardsGrid: { display: "grid", gridTemplateColumns: "minmax(0, 1.2fr) minmax(280px, 0.9fr) minmax(300px, 1fr)", gap: 10, alignItems: "stretch", minHeight: 0, overflow: "hidden" },
  panelCompact: { display: "grid", gap: 10, padding: 12, borderRadius: 16, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.94)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)", minHeight: 0, overflow: "hidden", gridTemplateRows: "auto auto minmax(0, 1fr)" },
  previewPanel: { display: "grid", gap: 10, padding: 12, borderRadius: 16, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.94)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)", minHeight: 0, overflow: "hidden", gridTemplateRows: "auto minmax(0, 1fr)" },
  sectionHeaderCompact: { display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" },
  sectionTitle: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#0f172a" },
  sectionSubtitle: { fontSize: 12, color: "#64748b" },
  previewTabs: { display: "inline-flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  previewTab: { borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#fff", color: "#42526E", padding: "6px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  previewTabActive: { borderRadius: 999, border: "1px solid rgba(37, 99, 235, 0.22)", background: "#dbeafe", color: "#1d4ed8", padding: "6px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  filterFieldWide: { display: "grid", gap: 3, minWidth: 0 },
  filterLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.05em", color: "#64748b" },
  input: { width: "100%", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "9px 10px", fontSize: 12, outline: "none", minWidth: 0 },
  listShellEmails: { display: "grid", gap: 6, overflowY: "auto", minHeight: 0, paddingRight: 4, alignContent: "start" },
  listShellDocumentsCompact: { display: "grid", gap: 6, overflowY: "auto", minHeight: 0, paddingRight: 4, alignContent: "start" },
  cardLine: { display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "center", padding: 10, borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff" },
  cardLineActive: { display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "center", padding: 10, borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.35)", background: "rgba(219, 234, 254, 0.55)" },
  cardMainExpanded: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 4, minWidth: 0, cursor: "pointer" },
  cardTitle: { fontSize: 12, fontWeight: 700, color: "#172b4d", lineHeight: 1.28, wordBreak: "break-word" },
  cardMetaLine: { fontSize: 11, color: "#64748b", lineHeight: 1.35, wordBreak: "break-word" },
  cardBadgeRow: { display: "flex", flexWrap: "wrap", gap: 4 },
  cardActions: { display: "inline-flex", gap: 4, alignItems: "center" },
  inlineActionBtn: { borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#fff", color: "#0f172a", padding: "6px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  segmentedControl: { display: "inline-flex", gap: 4, padding: 3, borderRadius: 999, background: "rgba(15, 23, 42, 0.06)" },
  segmentBtn: { borderRadius: 999, border: "none", background: "transparent", color: "#64748b", padding: "6px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  segmentBtnActive: { borderRadius: 999, border: "none", background: "#fff", color: "#0f172a", padding: "6px 10px", fontSize: 11, fontWeight: 800, boxShadow: "0 2px 8px rgba(15,23,42,0.08)", cursor: "pointer" },
  classificationSummary: { display: "grid", gap: 8, minHeight: 0, overflowY: "auto", alignContent: "start" },
  classificationTile: { border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", borderRadius: 12, padding: 10, textAlign: "left", display: "grid", gap: 4, cursor: "pointer" },
  classificationTileLabel: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.05em", color: "#64748b" },
  classificationTileValue: { fontSize: 13, color: "#0f172a", lineHeight: 1.3 },
  classificationTileMeta: { fontSize: 11, color: "#64748b" },
  advancedHintBox: { display: "flex", flexWrap: "wrap", gap: 6, paddingTop: 4 },
  advancedHintChip: { fontSize: 10, fontWeight: 700, color: "#1d4ed8", background: "rgba(219, 234, 254, 0.9)", borderRadius: 999, padding: "4px 8px" },
  classificationEditor: { display: "grid", gridTemplateRows: "auto auto auto 1fr auto", gap: 10, minHeight: 0 },
  backBtn: { justifySelf: "start", borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#fff", color: "#0f172a", padding: "6px 10px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  classificationEditorTitle: { fontSize: 13, fontWeight: 800, color: "#0f172a" },
  editorLead: { fontSize: 14, fontWeight: 700, color: "#172b4d", lineHeight: 1.35 },
  editorBodyText: { fontSize: 12, color: "#64748b", lineHeight: 1.45 },
  inlineChips: { display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" },
  emptyInlineText: { fontSize: 11, color: "#94a3b8" },
  primaryBtnWide: { borderRadius: 12, border: "none", background: "#1d4ed8", color: "#fff", padding: "10px 14px", fontSize: 12, fontWeight: 800, cursor: "pointer", width: "100%" },
  metaTag: { fontSize: 10, color: "#42526E", background: "#FFFFFF", borderRadius: 999, padding: "3px 8px", border: "1px solid rgba(15, 23, 42, 0.08)" },
  iconBtn: { width: 22, height: 22, borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", color: "#1d4ed8", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  iconBtnDanger: { width: 22, height: 22, borderRadius: 999, border: "1px solid rgba(239, 68, 68, 0.18)", background: "rgba(254, 226, 226, 0.9)", color: "#b91c1c", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  previewFrame: { borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", overflow: "hidden", background: "#f8fafc", minHeight: 0, height: "100%" },
  previewImage: { width: "100%", height: "100%", minHeight: 0, objectFit: "contain", display: "block", background: "#fff" },
  previewIframe: { width: "100%", height: "100%", border: "none", display: "block", background: "#fff" },
  pdfPreviewShell: { display: "grid", gridTemplateRows: "auto minmax(0, 1fr)", height: "100%", minHeight: 0, background: "#f8fafc" },
  pdfPreviewMeta: { padding: "8px 12px", fontSize: 10, fontWeight: 700, color: "#64748b", borderBottom: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.8)" },
  pdfPreviewCanvasHost: { overflow: "auto", padding: 12, display: "grid", justifyItems: "center", alignContent: "start", gap: 12, minHeight: 0 },
  previewText: { margin: 0, padding: 12, background: "#f8fafc", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", fontFamily: "Consolas, monospace", fontSize: 11, lineHeight: 1.45, whiteSpace: "pre-wrap", height: "100%", overflow: "auto", boxSizing: "border-box" },
  replyComposerShell: { display: "grid", placeItems: "center", alignContent: "center", justifyItems: "center", gap: 12, minHeight: 0, background: "linear-gradient(180deg, rgba(248,250,252,0.9), rgba(241,245,249,0.95))", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", padding: 24, textAlign: "center" },
  replyComposerLead: { fontSize: 18, fontWeight: 800, color: "#0f172a" },
  replyComposerMeta: { fontSize: 13, color: "#64748b", maxWidth: 560, lineHeight: 1.5 },
  replyComposerActions: { display: "flex", gap: 10, flexWrap: "wrap", justifyContent: "center" },
};
