import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  deleteGroupDocument,
  getGroupDocuments,
  getGroupEmails,
  listLinkGroups,
  removeEmailFromLinkGroup,
  type GroupDocumentEntry,
  type LinkGroupEntry,
  type RelatedEmailEntry,
} from "@/api";
import { addBase64AttachmentToCompose, getAttachments, getSelectedMessageContext, openLinkedOutlookEmail, type OutlookAttachment, type OutlookMessageContext } from "@/office";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import "../../global.css";

type EmailSortMode = "date_desc" | "date_asc" | "subject_asc" | "subject_desc";
type EmailAttachmentFilter = "all" | "with" | "without";
type DocumentFilterMode = "all" | "selected_email";
type PreviewMode = "email" | "document";
type PreviewState =
  | { kind: "image"; dataUrl: string }
  | { kind: "pdf"; dataUrl: string }
  | { kind: "text"; text: string }
  | { kind: "unsupported" };
type EmailAttachmentEntry = NonNullable<RelatedEmailEntry["attachments"]>[number];

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

function emailHasAttachments(email: RelatedEmailEntry): boolean {
  return Array.isArray(email.attachments) && email.attachments.length > 0;
}

function getEmailTimestamp(email: RelatedEmailEntry): number {
  const parsed = new Date(String(email.messageDateIso || email.receivedAtIso || email.sentAtIso || "").trim()).getTime();
  return Number.isFinite(parsed) ? parsed : 0;
}

function inferDocumentKind(document: GroupDocumentEntry): "image" | "pdf" | "text" | "unsupported" {
  const name = String(document.name || "").toLowerCase();
  const type = normalizeDocumentMimeType(document.contentType, document.name);
  if (!stripDataUrlPrefix(document.contentBase64)) return "unsupported";
  if (type.startsWith("image/") || /\.(png|jpe?g|gif|bmp|webp|svg)$/.test(name)) return "image";
  if (type.includes("pdf") || /\.pdf$/.test(name)) return "pdf";
  if (type.startsWith("text/") || type.includes("json") || type.includes("xml") || type.includes("csv") || /\.(txt|md|json|xml|csv|log|ya?ml)$/.test(name)) return "text";
  return "unsupported";
}

function buildDocumentPreview(document: GroupDocumentEntry | null): PreviewState | null {
  if (!document?.contentBase64) return null;
  const base64 = stripDataUrlPrefix(document.contentBase64);
  if (!base64) return null;
  const dataUrl = `data:${normalizeDocumentMimeType(document.contentType, document.name)};base64,${base64}`;
  const kind = inferDocumentKind(document);
  if (kind === "image") return { kind, dataUrl };
  if (kind === "pdf") return { kind, dataUrl };
  if (kind === "text") {
    try { return { kind, text: globalThis.atob(base64) }; } catch { return { kind: "unsupported" }; }
  }
  return { kind: "unsupported" };
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

function buildDocumentHoverText(document: GroupDocumentEntry): string {
  return [
    document.name ? `Documento: ${document.name}` : "",
    normalizeDocumentMimeType(document.contentType, document.name) ? `Tipo: ${normalizeDocumentMimeType(document.contentType, document.name)}` : "",
    formatBytes(document.size) ? `Tamanho: ${formatBytes(document.size)}` : "",
    document.sourceEmailSubject ? `Email: ${document.sourceEmailSubject}` : "",
  ].filter(Boolean).join("\n");
}

export default function GroupExplorerApp(): JSX.Element {
  const initial = useMemo(() => readExplorerParams(), []);
  const downloadAnchorRef = useRef<HTMLAnchorElement | null>(null);
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState(initial.groupId);
  const [selectedEmailKey, setSelectedEmailKey] = useState(initial.emailKey);
  const [selectedDocumentId, setSelectedDocumentId] = useState(initial.documentId);
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
  const [previewMode, setPreviewMode] = useState<PreviewMode>("document");
  const [liveCurrentContext, setLiveCurrentContext] = useState<OutlookMessageContext | null>(null);
  const [liveCurrentAttachments, setLiveCurrentAttachments] = useState<OutlookAttachment[]>([]);

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

  const filteredDocuments = useMemo(() => {
    const query = String(documentSearch || "").trim().toLowerCase();
    return groupDocuments.filter((document) => {
      if (documentFilterMode === "selected_email" && !matchesSelectedEmail(document, selectedEmail)) return false;
      if (!query) return true;
      const haystack = [document.name, document.sourceEmailSubject, document.contentType].map((value) => String(value || "").toLowerCase()).join(" ");
      return haystack.includes(query);
    });
  }, [documentFilterMode, documentSearch, groupDocuments, selectedEmail]);

  const selectedDocument = useMemo(
    () => filteredDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || groupDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || null,
    [filteredDocuments, groupDocuments, selectedDocumentId]
  );

  const selectedEmailAttachments = useMemo(() => {
    const persisted = Array.isArray(selectedEmail?.attachments) ? selectedEmail.attachments : [];
    if (!selectedEmail || !emailMatchesCurrentContext(selectedEmail, liveCurrentContext)) {
      return mergeEmailAttachments(persisted, []);
    }
    return mergeEmailAttachments(persisted, liveCurrentAttachments);
  }, [liveCurrentAttachments, liveCurrentContext, selectedEmail]);

  const selectedDocumentPreview = useMemo(() => buildDocumentPreview(selectedDocument), [selectedDocument]);
  const selectedEmailPreviewHtml = useMemo(
    () => buildEmailPreviewHtml(selectedEmail, selectedEmailAttachments),
    [selectedEmail, selectedEmailAttachments]
  );
  const selectedEmailHasPreview = Boolean(String(selectedEmail?.bodyHtml || "").trim() || String(selectedEmail?.bodyText || "").trim());
  const selectedProvider = useMemo(() => providerLabel(selectedDocument?.storageProvider || groupDocuments[0]?.storageProvider), [groupDocuments, selectedDocument]);

  useEffect(() => {
    if (!filteredEmails.some((email) => makeEmailKey(email) === selectedEmailKey)) {
      setSelectedEmailKey(filteredEmails[0] ? makeEmailKey(filteredEmails[0]) : "");
    }
  }, [filteredEmails, selectedEmailKey]);

  useEffect(() => {
    if (!filteredDocuments.some((document) => makeDocumentKey(document) === selectedDocumentId)) {
      setSelectedDocumentId("");
    }
  }, [filteredDocuments, selectedDocumentId]);

  useEffect(() => {
    if (selectedDocument) {
      setPreviewMode("document");
      return;
    }
    if (selectedEmailHasPreview) {
      setPreviewMode("email");
    }
  }, [selectedDocument, selectedEmailHasPreview, selectedEmailKey]);

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

  async function handleOpenEmail(email: RelatedEmailEntry) {
    const opened = await openLinkedOutlookEmail({ itemId: email.itemId, emailWebLink: email.emailWebLink });
    if (!opened) setNotice("Este email ainda nao tem abertura direta disponivel.");
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

function handleDownloadDocument(document: GroupDocumentEntry) {
  const base64 = stripDataUrlPrefix(document.contentBase64);
  if (!base64) {
    setNotice("Este documento nao tem conteudo disponivel para download.");
    return;
    }
    const bytes = globalThis.atob(base64);
    const buffer = new Array(bytes.length);
    for (let index = 0; index < bytes.length; index += 1) buffer[index] = bytes.charCodeAt(index);
  const blob = new Blob([new Uint8Array(buffer)], { type: normalizeDocumentMimeType(document.contentType, document.name) });
    const url = URL.createObjectURL(blob);
    const anchor = downloadAnchorRef.current || globalThis.document.createElement("a");
    downloadAnchorRef.current = anchor;
    anchor.href = url;
    anchor.download = document.name || "documento";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  async function handleAttachDocument(document: GroupDocumentEntry) {
    try {
      await addBase64AttachmentToCompose(document.name || "documento", stripDataUrlPrefix(document.contentBase64));
      setNotice(`Documento "${document.name}" anexado ao email em edicao.`);
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel anexar o documento.");
    }
  }

  async function handleDeleteDocument(document: GroupDocumentEntry) {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      await deleteGroupDocument(selectedGroup.id, document.id);
      await refreshCurrentGroup();
      setNotice(`Documento "${document.name}" removido.`);
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel remover o documento.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={styles.root}>
      <header style={styles.headerShell}>
        <div style={styles.headerIdentity}>
          <div style={styles.eyebrow}>Explorador documental</div>
          <div style={styles.title}>Grupos</div>
          <div style={styles.headerSelectorRow}>
            <div style={styles.selectWrap}>
              <select style={styles.select} value={selectedGroupId} onChange={(event) => { setSelectedGroupId(event.target.value); setSelectedEmailKey(""); setSelectedDocumentId(""); setNotice(null); }}>
                {groups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}
              </select>
            </div>
            <button type="button" style={styles.iconBtn} onClick={() => void refreshCurrentGroup()} disabled={loadingGroups || loadingEmails || loadingDocuments} title="Atualizar"><Icons.RefreshCw size={12} /></button>
            <button type="button" style={styles.closeBtn} onClick={closeExplorer}>Fechar</button>
          </div>
        </div>
        <div style={styles.headerMetrics}>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Grupo</span><span style={styles.metricValue}>{selectedGroup?.name || "-"}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Provider</span><span style={styles.metricValue}>{selectedProvider}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Docs</span><span style={styles.metricValue}>{groupDocuments.length}</span></div>
          <div style={styles.metricMini}><span style={styles.metricLabel}>Emails</span><span style={styles.metricValue}>{groupEmails.length}</span></div>
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
          <section style={styles.columnsGrid}>
            <section style={styles.panel}>
              <div style={styles.sectionHeaderCompact}>
                <div style={styles.sectionTitle}>Emails</div>
              </div>
              <div style={styles.filterGridEmails}>
                <label style={styles.filterFieldWide}>
                  <span style={styles.filterLabel}>Pesquisar</span>
                  <input style={styles.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Assunto, contacto ou email..." />
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Anexos</span>
                  <select style={styles.compactSelect} value={emailAttachmentFilter} onChange={(event) => setEmailAttachmentFilter(event.target.value as EmailAttachmentFilter)}>
                    <option value="all">Todos</option>
                    <option value="with">Com</option>
                    <option value="without">Sem</option>
                  </select>
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Ordenar</span>
                  <select style={styles.compactSelect} value={emailSort} onChange={(event) => setEmailSort(event.target.value as EmailSortMode)}>
                    <option value="date_desc">Recentes</option>
                    <option value="date_asc">Antigos</option>
                    <option value="subject_asc">A-Z</option>
                    <option value="subject_desc">Z-A</option>
                  </select>
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>De</span>
                  <input style={styles.compactInput} type="date" value={dateFrom} onChange={(event) => setDateFrom(event.target.value)} />
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Ate</span>
                  <input style={styles.compactInput} type="date" value={dateTo} onChange={(event) => setDateTo(event.target.value)} />
                </label>
              </div>
              <div style={styles.listShellEmails}>
                {loadingEmails && !groupEmails.length ? <PanelState compact tone="loading" title="A carregar emails" description="A listar os emails do grupo." /> : null}
                {!loadingEmails && !filteredEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Nao ha emails a corresponder aos filtros atuais." /> : null}
                {filteredEmails.map((email) => {
                  const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                  const canOpen = Boolean(email.itemId || email.emailWebLink);
                  const attachmentCount = emailHasAttachments(email) ? (email.attachments?.length || 0) : 0;
                  return (
                    <div key={makeEmailKey(email)} style={active ? styles.cardActive : styles.card}>
                      <button
                        type="button"
                        style={styles.cardMain}
                        onClick={() => setSelectedEmailKey(makeEmailKey(email))}
                        title={buildEmailHoverText(email)}
                      >
                        <div style={styles.cardTitle}>{email.subject || "(sem assunto)"}</div>
                        <div style={styles.cardBadgeRow}>
                          <span style={styles.metaTag} title={formatDate(email.messageDateIso || email.receivedAtIso) || "Sem data"}>{formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</span>
                          <span style={styles.metaTag} title={attachmentCount ? `${attachmentCount} anexo(s)` : "Sem anexos"}>{attachmentCount || 0}</span>
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

            <section style={styles.panel}>
              <div style={styles.sectionHeaderCompact}>
                <div style={styles.sectionTitle}>Documentos</div>
              </div>
              <div style={styles.filterGridDocuments}>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Filtro</span>
                  <select style={styles.compactSelect} value={documentFilterMode} onChange={(event) => setDocumentFilterMode(event.target.value as DocumentFilterMode)}>
                    <option value="all">Todos</option>
                    <option value="selected_email">Email ativo</option>
                  </select>
                </label>
                <label style={styles.filterFieldWide}>
                  <span style={styles.filterLabel}>Pesquisar</span>
                  <input style={styles.input} value={documentSearch} onChange={(event) => setDocumentSearch(event.target.value)} placeholder="Nome, tipo ou assunto..." />
                </label>
              </div>
              <div style={styles.listShellDocuments}>
                {loadingDocuments && !groupDocuments.length ? <PanelState compact tone="loading" title="A carregar documentos" description="A listar os documentos guardados." /> : null}
                {!loadingDocuments && !filteredDocuments.length ? <PanelState compact tone="info" title="Sem documentos visiveis" description={documentFilterMode === "selected_email" ? "Nao ha documentos associados ao email atualmente selecionado." : "Este grupo ainda nao tem documentos guardados visiveis neste filtro."} /> : null}
                {filteredDocuments.map((document) => {
                  const active = makeDocumentKey(document) === makeDocumentKey(selectedDocument || {});
                  return (
                    <div key={makeDocumentKey(document)} style={active ? styles.cardActive : styles.card}>
                      <button
                        type="button"
                        style={styles.cardMain}
                        onClick={() => setSelectedDocumentId(makeDocumentKey(document))}
                        title={buildDocumentHoverText(document)}
                      >
                        <div style={styles.cardTitle}>{document.name}</div>
                        <div style={styles.cardMetaCompact}>
                          <span>{formatBytes(document.size) || document.contentType || "Documento"}</span>
                        </div>
                      </button>
                      <div style={styles.cardActions}>
                        <button type="button" style={styles.iconBtn} onClick={() => handleDownloadDocument(document)} disabled={!document.contentBase64} title="Download"><Icons.Download size={10} /></button>
                        <button type="button" style={styles.iconBtn} onClick={() => void handleAttachDocument(document)} disabled={!document.contentBase64} title="Anexar ao email"><Icons.Upload size={10} /></button>
                        <button type="button" style={styles.iconBtnDanger} onClick={() => void handleDeleteDocument(document)} disabled={busy} title="Apagar"><Icons.Trash size={10} /></button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>
          </section>

          <section style={styles.previewPanel}>
            <div style={styles.sectionHeaderCompact}>
              <div style={styles.sectionTitle}>Preview</div>
              <div style={styles.sectionActions}>
                <div style={styles.previewTabs}>
                  <button
                    type="button"
                    style={previewMode === "email" ? styles.previewTabActive : styles.previewTab}
                    onClick={() => setPreviewMode("email")}
                    disabled={!selectedEmailHasPreview}
                    title={selectedEmailHasPreview ? "Ver preview do email selecionado" : "Este email ainda nao tem corpo guardado"}
                  >
                    Email
                  </button>
                  <button
                    type="button"
                    style={previewMode === "document" ? styles.previewTabActive : styles.previewTab}
                    onClick={() => setPreviewMode("document")}
                    disabled={!selectedDocument}
                    title={selectedDocument ? "Ver preview do documento selecionado" : "Seleciona um documento"}
                  >
                    Documento
                  </button>
                </div>
                {selectedDocument ? (
                  <>
                  <button type="button" style={styles.iconBtn} onClick={() => handleDownloadDocument(selectedDocument)} disabled={!selectedDocument.contentBase64} title="Download"><Icons.Download size={10} /></button>
                  <button type="button" style={styles.iconBtn} onClick={() => void handleAttachDocument(selectedDocument)} disabled={!selectedDocument.contentBase64} title="Anexar ao email"><Icons.Upload size={10} /></button>
                  </>
                ) : null}
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
            ) : !selectedDocument ? (
              <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe um documento para abrir o preview." />
            ) : selectedDocumentPreview?.kind === "image" ? (
              <div style={styles.previewFrame}><img src={selectedDocumentPreview.dataUrl} alt={selectedDocument.name} style={styles.previewImage} /></div>
            ) : selectedDocumentPreview?.kind === "pdf" ? (
              <div style={styles.previewFrame}><iframe title={selectedDocument.name} src={selectedDocumentPreview.dataUrl} style={styles.previewIframe} /></div>
            ) : selectedDocumentPreview?.kind === "text" ? (
              <pre style={styles.previewText}>{selectedDocumentPreview.text}</pre>
            ) : (
              <PanelState compact tone="info" title="Preview nao disponivel" description="Este documento pode ser descarregado ou anexado, mas ainda nao tem preview interno para este formato." />
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
  explorerBody: { display: "grid", gridTemplateRows: "minmax(280px, 0.9fr) minmax(300px, 1.1fr)", gap: 10, minHeight: 0, overflow: "hidden" },
  headerShell: { display: "grid", gridTemplateColumns: "minmax(240px, 1.25fr) minmax(300px, 0.95fr)", gap: 8, padding: 10, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.9)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.08)", minHeight: 116, maxHeight: 116, overflow: "hidden", alignItems: "stretch" },
  headerIdentity: { display: "grid", gap: 6, minWidth: 0, alignContent: "space-between" },
  eyebrow: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#5b6b83" },
  title: { fontSize: 16, fontWeight: 700, color: "#0f172a" },
  headerSelectorRow: { display: "grid", gridTemplateColumns: "1fr auto auto", gap: 6, alignItems: "center" },
  headerMetrics: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 6, alignContent: "stretch", minHeight: 0 },
  metricMini: { display: "grid", gap: 1, padding: 8, borderRadius: 10, background: "rgba(15, 23, 42, 0.03)", minWidth: 0 },
  metricLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.06em", color: "#6b7280" },
  metricValue: { fontSize: 11, fontWeight: 600, color: "#0f172a", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  selectWrap: { minWidth: 0 },
  select: { width: "100%", borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "rgba(248,250,252,0.95)", color: "#172b4d", padding: "8px 12px", fontSize: 11, fontWeight: 600, outline: "none" },
  closeBtn: { borderRadius: 999, border: "none", background: "#1d4ed8", color: "#fff", padding: "7px 13px", fontSize: 11, fontWeight: 700, cursor: "pointer" },
  columnsGrid: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) minmax(0, 1fr)", gap: 10, alignItems: "stretch", minHeight: 0, overflow: "hidden" },
  panel: { display: "grid", gap: 8, padding: 10, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.92)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)", minHeight: 0, height: "100%", gridTemplateRows: "auto auto minmax(0, 1fr)", overflow: "hidden" },
  previewPanel: { display: "grid", gap: 8, padding: 10, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.92)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)", minHeight: 0, height: "100%", overflow: "hidden", gridTemplateRows: "auto minmax(0, 1fr)" },
  sectionHeaderCompact: { display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center" },
  sectionTitle: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#0f172a" },
  sectionActions: { display: "inline-flex", gap: 4, alignItems: "center" },
  previewTabs: { display: "inline-flex", gap: 4, alignItems: "center", marginRight: 4 },
  previewTab: { borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#fff", color: "#42526E", padding: "4px 8px", fontSize: 10, fontWeight: 700, cursor: "pointer" },
  previewTabActive: { borderRadius: 999, border: "1px solid rgba(37, 99, 235, 0.22)", background: "#dbeafe", color: "#1d4ed8", padding: "4px 8px", fontSize: 10, fontWeight: 700, cursor: "pointer" },
  filterGridEmails: { display: "grid", gridTemplateColumns: "minmax(0, 1.4fr) repeat(4, minmax(0, 0.7fr))", gap: 6, alignItems: "end" },
  filterGridDocuments: { display: "grid", gridTemplateColumns: "minmax(120px, 0.6fr) minmax(0, 1.4fr)", gap: 6, alignItems: "end" },
  filterField: { display: "grid", gap: 3, minWidth: 0 },
  filterFieldWide: { display: "grid", gap: 3, minWidth: 0 },
  filterLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.05em", color: "#64748b" },
  input: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "7px 9px", fontSize: 11, outline: "none", minWidth: 0 },
  compactInput: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "7px 9px", fontSize: 11, outline: "none", minWidth: 0 },
  compactSelect: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "7px 9px", fontSize: 11, outline: "none", minWidth: 0 },
  listShellTall: { display: "grid", gap: 6, overflowY: "auto", minHeight: 0, paddingRight: 4 },
  listShellEmails: { display: "grid", gap: 6, overflowY: "auto", minHeight: 0, paddingRight: 4, alignContent: "start" },
  listShellDocuments: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 6, overflowY: "auto", minHeight: 0, paddingRight: 4, alignContent: "start" },
  card: { display: "grid", gridTemplateColumns: "1fr auto", gap: 6, alignItems: "center", padding: 7, borderRadius: 11, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff" },
  cardActive: { display: "grid", gridTemplateColumns: "1fr auto", gap: 6, alignItems: "center", padding: 7, borderRadius: 11, border: "1px solid rgba(37, 99, 235, 0.35)", background: "rgba(219, 234, 254, 0.65)" },
  cardMain: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 3, minWidth: 0, cursor: "pointer" },
  cardTitle: { fontSize: 10.5, fontWeight: 600, color: "#172b4d", lineHeight: 1.22, wordBreak: "break-word" },
  cardMeta: { display: "none" },
  cardMetaCompact: { display: "flex", flexWrap: "wrap", gap: 6, fontSize: 9.5, color: "#6b778c", lineHeight: 1.2 },
  cardBadgeRow: { display: "flex", flexWrap: "wrap", gap: 4 },
  cardActions: { display: "inline-flex", gap: 4, alignItems: "center" },
  metaTag: { fontSize: 9, color: "#42526E", background: "#FFFFFF", borderRadius: 999, padding: "1px 6px", border: "1px solid rgba(15, 23, 42, 0.08)" },
  iconBtn: { width: 22, height: 22, borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", color: "#1d4ed8", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  iconBtnDanger: { width: 22, height: 22, borderRadius: 999, border: "1px solid rgba(239, 68, 68, 0.18)", background: "rgba(254, 226, 226, 0.9)", color: "#b91c1c", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  previewFrame: { borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", overflow: "hidden", background: "#f8fafc", minHeight: 0, height: "100%" },
  previewImage: { width: "100%", height: "100%", minHeight: 0, objectFit: "contain", display: "block", background: "#fff" },
  previewIframe: { width: "100%", height: "100%", border: "none", display: "block", background: "#fff" },
  previewText: { margin: 0, padding: 12, background: "#f8fafc", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", fontFamily: "Consolas, monospace", fontSize: 11, lineHeight: 1.45, whiteSpace: "pre-wrap", height: "100%", overflow: "auto", boxSizing: "border-box" },
};
