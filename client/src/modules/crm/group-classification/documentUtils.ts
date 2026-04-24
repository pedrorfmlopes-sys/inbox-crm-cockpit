import {
  EmailRecipientEntry, RelatedEmailEntry, RelevantEmailPayload, LinkGroupEntry, GroupTicketEntry
} from "@/api";
import {
  ClassificationMetaDraft, DocumentLifecycleState, LabelDraft,
  GroupContactDraft, GroupEntityDraft, StudioParams
} from "./types";
import { GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX, EMPTY_CLASSIFICATION_META } from "./constants";
import { getGroupAttachmentStorageOptions } from "@/modules/crm/groups-v1/storage/resolveStorageMode";
import {
  buildGroupsTabAttachmentStorageOptions,
  resolveGroupsTabAttachmentDecision,
} from "@/modules/crm/groups-v1/settings/groupsTabRuntime";

// StudioParams definition moved to types.ts

export function readParams(): StudioParams {
  const urlParams = new URLSearchParams(window.location.search);
  return {
    seedKey: urlParams.get("seedKey") || undefined,
    prepareSeedKey: urlParams.get("prepareSeedKey") || undefined,
    caseId: urlParams.get("caseId") || undefined,
    anchorEmailKey: urlParams.get("anchorEmailKey") || undefined,
    itemId: urlParams.get("itemId") || undefined,
    internetMessageId: urlParams.get("internetMessageId") || undefined,
    conversationId: urlParams.get("conversationId") || undefined,
    subject: urlParams.get("subject") || undefined,
    fromEmail: urlParams.get("fromEmail") || undefined,
    fromName: urlParams.get("fromName") || undefined,
    receivedAtIso: urlParams.get("receivedAtIso") || undefined,
  };
}

export function readSeedEmail(params: StudioParams): RelatedEmailEntry | null {
  const key = String(params.seedKey || "").trim();
  if (!key || !key.startsWith(GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX)) return null;
  try {
    const raw = typeof localStorage !== "undefined" ? localStorage.getItem(key) : null;
    if (!raw) return null;
    const parsed: any = JSON.parse(raw);
    const itemId = String(parsed?.itemId || "").trim();
    const internetMessageId = String(parsed?.internetMessageId || "").trim();
    const conversationId = String(parsed?.conversationId || "").trim();
    const subject = String(parsed?.subject || "").trim();
    const fromEmail = String(parsed?.fromEmail || "").trim();
    const fromName = String(parsed?.fromName || "").trim();
    const emailWebLink = String(parsed?.emailWebLink || "").trim();
    const receivedAtIso = String(parsed?.receivedAtIso || parsed?.messageDateIso || "").trim();
    const sentAtIso = String(parsed?.sentAtIso || "").trim();
    const toRecipients = Array.isArray(parsed?.toRecipients)
      ? parsed.toRecipients
          .map((recipient: any) => ({
            email: String(recipient?.email || "").trim().toLowerCase(),
            name: String(recipient?.name || "").trim() || undefined,
          }))
          .filter((recipient: { email: string }) => recipient.email)
      : [];
    const ccRecipients = Array.isArray(parsed?.ccRecipients)
      ? parsed.ccRecipients
          .map((recipient: any) => ({
            email: String(recipient?.email || "").trim().toLowerCase(),
            name: String(recipient?.name || "").trim() || undefined,
          }))
          .filter((recipient: { email: string }) => recipient.email)
      : [];
    if (!(itemId || internetMessageId || conversationId || subject || fromEmail)) return null;
    return ({
      emailKey: itemId || internetMessageId || `${conversationId}|${subject || fromEmail}`,
      itemId: itemId || undefined,
      internetMessageId: internetMessageId || undefined,
      conversationId: conversationId || undefined,
      subject: subject || undefined,
      fromEmail: fromEmail || undefined,
      fromName: fromName || undefined,
      emailWebLink: emailWebLink || undefined,
      receivedAtIso: receivedAtIso || undefined,
      messageDateIso: receivedAtIso || undefined,
      sentAtIso: sentAtIso || undefined,
      toRecipients,
      ccRecipients,
      bodyHtml: String(parsed?.bodyHtml || "").trim(),
      attachments: Array.isArray(parsed?.attachments)
        ? parsed.attachments
          .map((attachment: any) => ({
            id: String(attachment?.id || "").trim() || undefined,
            name: String(attachment?.name || "").trim(),
            contentType: String(attachment?.contentType || "").trim(),
            size: Number(attachment?.size || 0),
            isInline: attachment?.isInline === true,
            contentId: String(attachment?.contentId || "").trim() || undefined,
          }))
        : [],
      relatedGroups: [],
      relatedReasons: [],
    } as any);
  } catch {
    return null;
  }
}

export function buildFallbackEmail(params: StudioParams): RelatedEmailEntry {
  const itemId = String(params.itemId || "").trim();
  const internetMessageId = String(params.internetMessageId || "").trim();
  const conversationId = String(params.conversationId || "").trim();
  const subject = String(params.subject || "").trim();
  const fromEmail = String(params.fromEmail || "").trim();
  const fromName = String(params.fromName || "").trim();
  const receivedAtIso = String(params.receivedAtIso || "").trim();

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
    isFallback: true,
    relatedGroups: [],
    relatedReasons: [],
  } as any;
}

export function normalizeDocumentLifecycleState(value: string | undefined, fallback: DocumentLifecycleState = "ingested"): DocumentLifecycleState {
  const token = String(value || "").trim().toLowerCase();
  if (token === "ingested") return "ingested";
  if (token === "processed") return "processed";
  if (token === "accepted") return "accepted";
  if (token === "rejected") return "rejected";
  if (token === "reread_requested") return "reread_requested";
  return fallback;
}

export function formatDocumentLifecycleState(value: string | undefined): string {
  const token = normalizeDocumentLifecycleState(value);
  if (token === "ingested") return "Ingerido";
  if (token === "processed") return "Processado";
  if (token === "accepted") return "Aceite";
  if (token === "rejected") return "Rejeitado";
  if (token === "reread_requested") return "Re-leitura";
  return String(value || "--");
}

export function isRejectedDocumentLifecycleState(value: string | undefined): boolean {
  return normalizeDocumentLifecycleState(value) === "rejected";
}

export function normalizeStudioAttachment(attachment: any) {
  if (!attachment) return null;
  const key = String(attachment.key || attachment.id || attachment.contentId || attachment.name || "").trim();
  if (!key) return null;
  return {
    ...attachment,
    key,
    name: String(attachment.name || "").trim() || "Anexo",
    contentType: String(attachment.contentType || "application/octet-stream").trim().toLowerCase(),
    size: Number(attachment.size || 0) || 0,
    isInline: Boolean(attachment.isInline || attachment.contentId),
    contentId: String(attachment.contentId || "").trim() || undefined,
    id: String(attachment.id || "").trim() || undefined,
    documentState: normalizeDocumentLifecycleState(attachment.documentState),
    storagePathHint: String(attachment.storagePathHint || "").trim() || undefined,
  };
}

export function normalizeStudioAttachmentMimeType(value: string | undefined, name: string | undefined): string {
  const ct = String(value || "").trim().toLowerCase();
  if (ct) return ct;
  const n = String(name || "").trim().toLowerCase();
  if (n.endsWith(".pdf")) return "application/pdf";
  if (n.endsWith(".png")) return "image/png";
  if (n.endsWith(".jpg") || n.endsWith(".jpeg")) return "image/jpeg";
  return "application/octet-stream";
}

export function inferStudioAttachmentKind(
  contentType: string | undefined,
  name: string | undefined
): "image" | "pdf" | "office" | "text" | "unsupported" {
  const ct = String(contentType || "").trim().toLowerCase();
  const n = String(name || "").trim().toLowerCase();
  if (ct.startsWith("image/")) return "image";
  if (ct === "application/pdf" || n.endsWith(".pdf")) return "pdf";
  if (
    ct.includes("officedocument") ||
    ct.includes("ms-excel") ||
    ct.includes("ms-word") ||
    ct.includes("ms-powerpoint") ||
    n.endsWith(".docx") || n.endsWith(".xlsx") || n.endsWith(".pptx")
  ) {
    return "office";
  }
  if (ct.startsWith("text/")) return "text";
  return "unsupported";
}

export function isLikelyDecorativeAttachment(
  attachment: { name?: string; size?: number; isInline?: boolean; contentType?: string } | null
): boolean {
  if (!attachment) return false;
  const name = String(attachment.name || "").toLowerCase();
  const isInline = Boolean(attachment.isInline);
  const size = Number(attachment.size || 0);
  const ct = String(attachment.contentType || "").toLowerCase();
  if (name.includes("image001") || name.includes("image002") || name.includes("image003")) return true;
  if (isInline && size < 15000) return true;
  if (ct.includes("image/") && size < 5000) return true;
  return false;
}

export function isStudioAttachmentHiddenInQuickDocs(
  attachment: any,
  attachmentTextMap: Record<string, string> = {}
): boolean {
  if (!attachment) return true;
  if (attachment.isHidden === true) return true;
  if (isLikelyDecorativeAttachment(attachment)) return true;
  const kind = inferStudioAttachmentKind(attachment.contentType, attachment.name);
  if (kind === "unsupported") return true;
  const key = attachment.key || attachment.id || attachment.contentId || attachment.name;
  if (kind === "text" && !String(attachment.content || attachmentTextMap[key] || "").trim()) return true;
  return false;
}

export function formatQuickDocumentMeta(
  attachment: { size?: number; contentType?: string; documentState?: string } | null
): string {
  if (!attachment) return "--";
  const sizeKb = Math.ceil((attachment.size || 0) / 1024);
  const sizeText = sizeKb > 1024 ? `${(sizeKb / 1024).toFixed(1)} MB` : `${sizeKb} KB`;
  const stateText = formatDocumentLifecycleState(attachment.documentState);
  return `${sizeText} | ${stateText}`;
}

export function buildQuickDocumentPreviewText(
  attachment: { name?: string; contentType?: string; content?: string } | null,
  attachmentTextMap: Record<string, string> = {}
): string {
  if (!attachment) return "";
  const key = (attachment as any).key || (attachment as any).id || (attachment as any).contentId || attachment.name || "";
  const text = String(attachment.content || attachmentTextMap[key] || "").trim();
  if (!text) return "";
  const lines = text.split("\n").map((l) => l.trim()).filter(Boolean);
  return lines.slice(0, 3).join(" | ");
}

export function htmlToPlainText(html: string): string {
  if (!html) return "";
  let text = html
    .replace(/<style([\s\S]*?)<\/style>/gi, "")
    .replace(/<script([\s\S]*?)<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  const entities: Record<string, string> = {
    "&nbsp;": " ",
    "&amp;": "&",
    "&lt;": "<",
    "&gt;": ">",
    "&quot;": "\"",
    "&#39;": "'",
  };
  Object.keys(entities).forEach((key) => {
    text = text.replace(new RegExp(key, "g"), entities[key]);
  });
  return text;
}

export function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  const emailKey = String(email?.emailKey || "").trim();
  const emailId = String(email?.id || "").trim();
  const itemId = String(email?.itemId || "").trim();
  const internetMessageId = String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
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

export function makeAttachmentKey(attachment: { key?: string; id?: string; name?: string; contentId?: string }): string {
  return String(attachment?.key || attachment?.id || attachment?.contentId || attachment?.name || "");
}

export function getStudioAttachmentRemoteId(attachment: any): string {
  return String(attachment?.id || attachment?.contentId || attachment?.key || "");
}

export function isStudioAttachmentHydrated(attachment: any): boolean {
  if (!attachment) return false;
  return Boolean(attachment.contentBase64 || attachment.content || attachment.hasContent);
}

export function isStudioAttachmentHydratedInCollection(attachments: any[], key: string): boolean {
  const found = attachments.find((a) => makeAttachmentKey(a) === key);
  return isStudioAttachmentHydrated(found);
}

export function hasHydratedAttachmentCollection(email: RelatedEmailEntry | null): boolean {
  if (!email || !Array.isArray(email.attachments)) return false;
  return email.attachments.some((a) => isStudioAttachmentHydrated(a));
}

export function mergeUniqueStrings(values: string[]): string[] {
  return Array.from(new Set((values || []).map((v) => String(v || "").trim()).filter(Boolean)));
}

export function mergeUniqueBy<T>(values: T[], getKey: (value: T) => string): T[] {
  const seen = new Set<string>();
  const result: T[] = [];
  for (const val of values) {
    const key = getKey(val);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(val);
  }
  return result;
}

export function scoreStudioAttachment(attachment: any): number {
  if (!attachment) return 0;
  let score = 0;
  const name = String(attachment.name || "").toLowerCase();
  const kind = inferStudioAttachmentKind(attachment.contentType, attachment.name);
  if (kind === "pdf") score += 50;
  if (kind === "office") score += 40;
  if (kind === "image" && !isLikelyDecorativeAttachment(attachment)) score += 20;
  if (name.includes("fatura") || name.includes("invoice") || name.includes("proforma")) score += 30;
  if (name.includes("pedido") || name.includes("order")) score += 25;
  return score;
}

export function scoreStudioAttachmentCollection(attachments: any[]): number {
  if (!Array.isArray(attachments) || !attachments.length) return 0;
  return attachments.reduce((acc, a) => acc + scoreStudioAttachment(a), 0);
}

export function normalizeClassificationMetaDraft(
  value?: Partial<ClassificationMetaDraft> | null
): ClassificationMetaDraft {
  const base = value || {};
  return {
    ...EMPTY_CLASSIFICATION_META,
    ...base,
    referenceGroupIds: mergeUniqueStrings(base.referenceGroupIds || []),
    labelStates: base.labelStates || {},
    principalGroupState: String(base.principalGroupState || "").trim(),
    referenceGroupStates:
      base.referenceGroupStates && typeof base.referenceGroupStates === "object"
        ? Object.fromEntries(
            Object.entries(base.referenceGroupStates)
              .map(([groupId, state]) => [String(groupId || "").trim(), String(state || "").trim()])
              .filter(([groupId, state]) => groupId && state)
          )
        : {},
  };
}

export function mergeClassificationMetaDrafts(
  left: Partial<ClassificationMetaDraft>,
  right: Partial<ClassificationMetaDraft>
): ClassificationMetaDraft {
  const l = normalizeClassificationMetaDraft(left);
  const r = normalizeClassificationMetaDraft(right);
  return {
    principalGroupId: r.principalGroupId || l.principalGroupId,
    principalGroupState: r.principalGroupState || l.principalGroupState,
    ticketId: r.ticketId || l.ticketId,
    referenceGroupIds: Array.from(new Set([...l.referenceGroupIds, ...r.referenceGroupIds])),
    labelStates: { ...l.labelStates, ...r.labelStates },
    referenceGroupStates: { ...l.referenceGroupStates, ...r.referenceGroupStates },
  };
}

export function scoreRelatedEmailEntry(email: RelatedEmailEntry | null | undefined): number {
  if (!email) return 0;
  let score = 0;
  if (email.bodyText && email.bodyText.length > 50) score += 10;
  if (Array.isArray(email.attachments)) {
    score += scoreStudioAttachmentCollection(email.attachments);
  }
  if (!(email as any).isFallback) score += 5;
  return score;
}

export function mergeRelatedEmailEntries(current: RelatedEmailEntry, incoming: RelatedEmailEntry): RelatedEmailEntry {
  const currentScore = scoreRelatedEmailEntry(current);
  const incomingScore = scoreRelatedEmailEntry(incoming);
  const winner = incomingScore >= currentScore ? incoming : current;
  const loser = winner === incoming ? current : incoming;

  const mergedAttachments = mergeUniqueBy(
    [...(winner.attachments || []), ...(loser.attachments || [])].map((a) => normalizeStudioAttachment(a)).filter(Boolean) as any[],
    (a) => a.key
  );

  return {
    ...winner,
    attachments: mergedAttachments,
    relatedGroups: mergeUniqueBy([...(winner.relatedGroups || []), ...(loser.relatedGroups || [])], (g) => g.id),
    relatedReasons: Array.from(new Set([...(winner.relatedReasons || []), ...(loser.relatedReasons || [])])),
  };
}

export function dedupeEmails(emails: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const seen = new Map<string, RelatedEmailEntry>();
  for (const email of emails) {
    const key = makeEmailKey(email);
    if (!key) continue;
    const current = seen.get(key);
    seen.set(key, current ? mergeRelatedEmailEntries(current, email) : email);
  }
  return Array.from(seen.values());
}

export function buildRelevantEmailPayloadFromRelatedEmail(email: RelatedEmailEntry | null, settings?: any): RelevantEmailPayload | null {
  if (!email) return null;
  const normalizeRecipients = (values: Array<EmailRecipientEntry | null | undefined> | undefined): EmailRecipientEntry[] =>
    Array.isArray(values)
      ? values
          .map((recipient) => ({
            email: String(recipient?.email || "").trim().toLowerCase(),
            name: String(recipient?.name || "").trim() || undefined,
          }))
          .filter((recipient) => recipient.email)
      : [];
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
        .map((attachment) => {
          const decision = resolveGroupsTabAttachmentDecision({
            key: attachment.key,
            name: attachment.name,
            size: attachment.size,
            isInline: attachment.isInline,
            hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
          }, settings || null);
          if (decision.storageDecision === "skip_inline") return null;
          const content = decision.includeBinaryInPayload ? attachment.content : undefined;
          const hasContent = decision.includeBinaryInPayload
            ? attachment.hasContent === true || Boolean(String(attachment.content || "").trim())
            : false;
          if (!decision.includeMetadataOnServer && !content) return null;
          return {
            key: attachment.key,
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            size: attachment.size,
            isInline: attachment.isInline,
            contentId: attachment.contentId,
            content,
            storageProvider: attachment.storageProvider || decision.storageProvider,
            storageBasePath: attachment.storageBasePath || decision.storageBasePath,
            storagePathHint: attachment.storagePathHint || decision.storagePathHint,
            documentState: attachment.documentState,
            hasContent,
            isHidden: typeof attachment.isHidden === "boolean" ? attachment.isHidden : undefined,
          };
        })
        .filter(Boolean)
    : [];

  return {
    itemId: itemId || undefined,
    internetMessageId: internetMessageId || undefined,
    conversationId: conversationId || undefined,
    subject: subject || undefined,
    fromEmail: fromEmail || undefined,
    fromName: email.fromName || undefined,
    emailWebLink: String(email.emailWebLink || "").trim() || undefined,
    receivedAtIso: String(email.receivedAtIso || email.messageDateIso || "").trim() || undefined,
    sentAtIso: String(email.sentAtIso || "").trim() || undefined,
    messageDateIso: String(email.messageDateIso || email.receivedAtIso || "").trim() || undefined,
    toRecipients: normalizeRecipients(email.toRecipients),
    ccRecipients: normalizeRecipients(email.ccRecipients),
    bodyText: email.bodyText || "",
    bodyHtml: email.bodyHtml || "",
    replaceAttachments: false,
    attachments,
  };
}

export function buildAttachmentStorageOptions(settings?: any): Pick<RelevantEmailPayload, "attachmentStorageProvider" | "attachmentStorageBasePath"> {
  const tabAware = buildGroupsTabAttachmentStorageOptions(settings || null);
  if (String(tabAware.attachmentStorageBasePath || "").trim()) {
    return tabAware;
  }
  return getGroupAttachmentStorageOptions(settings?.groups?.storage || null);
}

export function normalizeSearchValue(value: string): string {
  return String(value || "").trim().toLowerCase();
}

export function normalizeReferenceCandidate(value: string): string {
  return String(value || "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
}

export function compactReferenceValue(value: string): string {
  return String(value || "").trim().toUpperCase();
}

export function matchReferenceSet(text: string, references: string[]): string[] {
  const normText = normalizeReferenceCandidate(text);
  if (!normText) return [];
  return references.filter((ref) => {
    const normRef = normalizeReferenceCandidate(ref);
    return normRef && normText.includes(normRef);
  });
}

export function formatDate(value: string | undefined): string {
  if (!value) return "--";
  try {
    return new Date(value).toLocaleDateString("pt-PT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  } catch {
    return value;
  }
}

export function buildSnippet(email: RelatedEmailEntry): string {
  return htmlToPlainText(email.bodyHtml || email.bodyText || "").slice(0, 150) + "...";
}

export function buildEmailPreviewText(email: RelatedEmailEntry): string {
  return htmlToPlainText(email.bodyHtml || email.bodyText || "").slice(0, 200);
}

export function buildCompactEmailMeta(email: RelatedEmailEntry): string {
  const date = formatDate(email.receivedAtIso || email.messageDateIso);
  const from = email.fromName || email.fromEmail || "Desconhecido";
  return `${from} | ${date}`;
}

export function buildEmailCorpus(email: RelatedEmailEntry): string {
  return [
    email.subject,
    email.fromName,
    email.fromEmail,
    htmlToPlainText(email.bodyHtml || email.bodyText || "")
  ].filter(Boolean).join(" ");
}

export function isExternalEmail(email: RelatedEmailEntry): boolean {
  const from = String(email.fromEmail || "").toLowerCase();
  return Boolean(from && !from.includes("nicolazzi.it") && !from.includes("inboxcockpit.com"));
}

export function isCurrentContextEmail(email: Partial<RelatedEmailEntry>, currentContext: Partial<StudioParams>) {
  if (!email) return false;
  const ek = makeEmailKey(email);
  const ck = makeEmailKey({
    itemId: currentContext.itemId,
    internetMessageId: currentContext.internetMessageId,
    conversationId: currentContext.conversationId,
    subject: currentContext.subject,
    fromEmail: currentContext.fromEmail,
    receivedAtIso: currentContext.receivedAtIso
  } as any);
  return ek === ck;
}

export function detectCaseType(text: string): string {
  const t = text.toLowerCase();
  if (t.includes("fatura") || t.includes("invoice")) return "Fatura";
  if (t.includes("pedido") || t.includes("order")) return "Pedido";
  if (t.includes("proforma")) return "Proforma";
  return "Geral";
}

export function inferCompanyName(fromName: string | undefined, fromEmail: string | undefined): string {
  if (fromName && fromName.includes(" - ")) return fromName.split(" - ")[1].trim();
  const domain = String(fromEmail || "").split("@")[1];
  if (domain && !["gmail.com", "outlook.com", "hotmail.com"].includes(domain)) {
    return domain.split(".")[0].charAt(0).toUpperCase() + domain.split(".")[0].slice(1);
  }
  return "";
}

export function normalizeGroupContactDraft(value: Partial<GroupContactDraft> | null | undefined): GroupContactDraft | null {
  if (!value || !value.email) return null;
  return {
    key: String(value.email || "").trim().toLowerCase(),
    name: String(value.name || "").trim(),
    email: String(value.email || "").trim().toLowerCase(),
    role: String(value.role || "").trim(),
    isPrincipal: value.isPrincipal === true,
  };
}

export function normalizeGroupEntityDraft(value: Partial<GroupEntityDraft> | null | undefined): GroupEntityDraft | null {
  if (!value || (!value.id && !value.name)) return null;
  return {
    key: String(value.id || value.name || "").trim(),
    id: value.id,
    name: String(value.name || "").trim(),
    role: String(value.role || "").trim(),
    isPrincipal: value.isPrincipal === true,
  };
}

export function dedupeGroupContacts(rows: Array<Partial<GroupContactDraft> | null | undefined>): GroupContactDraft[] {
  const seen = new Set<string>();
  const result: GroupContactDraft[] = [];
  for (const row of rows) {
    const norm = normalizeGroupContactDraft(row);
    if (norm && !seen.has(norm.key)) {
      seen.add(norm.key);
      result.push(norm);
    }
  }
  return result;
}

export function dedupeGroupEntities(rows: Array<Partial<GroupEntityDraft> | null | undefined>): GroupEntityDraft[] {
  const seen = new Set<string>();
  const result: GroupEntityDraft[] = [];
  for (const row of rows) {
    const norm = normalizeGroupEntityDraft(row);
    if (norm && !seen.has(norm.key)) {
      seen.add(norm.key);
      result.push(norm);
    }
  }
  return result;
}

export function detectReferences(text: string): string[] {
  const matches = text.match(/[A-Z0-9]{5,}/g) || [];
  return Array.from(new Set(matches));
}

export function splitSuggestions(allGroups: any[], text: string): any[] {
  return allGroups.filter((g) => text.includes(g.name));
}
