import { getLinksByRecord, getRelatedEmailContext, type GroupTicketEntry, type LinkEntry, type LinkGroupEntry, type RelatedEmailEntry } from "@/api";
import type { OutlookAttachment, OutlookMessageContext } from "@/office";

type BuildAiContextBundleInput = {
  ctx: OutlookMessageContext;
  bodyText: string;
  bodyHtml?: string;
  links?: LinkEntry[];
  attachments?: OutlookAttachment[];
};

export type AiContextBundle = {
  cacheKey: string;
  promptContext: string;
  briefingContext: string;
  groups: LinkGroupEntry[];
  tickets: GroupTicketEntry[];
  relatedEmails: RelatedEmailEntry[];
  linkedRecordEmails: LinkEntry[];
  linkedRecords: Array<{ model: string; recordId: number; recordName: string }>;
};

function normalizeText(value: unknown): string {
  return String(value || "").replace(/\r/g, "").trim();
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
    .replace(/&quot;/gi, '"')
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function truncate(text: string, max = 400): string {
  const normalized = normalizeText(text);
  if (!normalized) return "";
  if (normalized.length <= max) return normalized;
  return `${normalized.slice(0, Math.max(0, max - 3)).trim()}...`;
}

function formatDateLabel(value: unknown): string {
  const raw = normalizeText(value);
  if (!raw) return "";
  const match = raw.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : raw.slice(0, 16);
}

function makeEmailIdentity(entry: Partial<RelatedEmailEntry | LinkEntry>): string {
  const itemId = normalizeText((entry as any)?.itemId);
  const internetMessageId = normalizeText((entry as any)?.internetMessageId).toLowerCase().replace(/[<>\s]/g, "");
  const conversationId = normalizeText((entry as any)?.conversationId);
  const subject = normalizeText((entry as any)?.subject).toLowerCase();
  const date = normalizeText((entry as any)?.messageDateIso || (entry as any)?.receivedAtIso || (entry as any)?.sentAtIso);
  return [itemId, internetMessageId, conversationId, subject, date].join("|");
}

function dedupeRelatedEmails(entries: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const seen = new Set<string>();
  return (entries || []).filter((entry) => {
    const key = makeEmailIdentity(entry) || normalizeText(entry.emailKey);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function dedupeLinkEntries(entries: LinkEntry[]): LinkEntry[] {
  const seen = new Set<string>();
  return (entries || []).filter((entry) => {
    const key = makeEmailIdentity(entry);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function dedupeLinkedRecords(entries: LinkEntry[]): Array<{ model: string; recordId: number; recordName: string }> {
  const seen = new Set<string>();
  return (entries || [])
    .map((entry) => {
      const model = normalizeText(entry.model);
      const recordId = Number(entry.recordId || entry.resId || 0);
      const recordName = normalizeText(entry.recordName || entry.name || entry.title);
      if (!model || !recordId) return null;
      return { model, recordId, recordName };
    })
    .filter((entry): entry is { model: string; recordId: number; recordName: string } => Boolean(entry))
    .filter((entry) => {
      const key = `${entry.model}:${entry.recordId}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function summarizeReasons(email: RelatedEmailEntry): string {
  const parts: string[] = [];
  for (const reason of Array.isArray(email.relatedReasons) ? email.relatedReasons : []) {
    if (reason?.kind === "entity") {
      const label = [normalizeText(reason.recordName), normalizeText(reason.model)].filter(Boolean).join(" | ");
      if (label) parts.push(`registo ${label}`);
    } else if ((reason?.kind === "group" || reason?.kind === "conversation") && reason.groupName) {
      parts.push(reason.kind === "conversation" ? `thread ${reason.groupName}` : `grupo ${reason.groupName}`);
    }
  }
  return Array.from(new Set(parts)).slice(0, 3).join("; ");
}

function scoreEmail(entry: RelatedEmailEntry): number {
  const reasons = Array.isArray(entry.relatedReasons) ? entry.relatedReasons : [];
  let score = 0;
  if (reasons.some((reason) => reason?.kind === "conversation")) score += 100;
  if (reasons.some((reason) => reason?.kind === "entity")) score += 60;
  if (reasons.some((reason) => reason?.kind === "group")) score += 40;
  if (Array.isArray(entry.relatedRecords) && entry.relatedRecords.length) score += 20;
  if (Array.isArray(entry.relatedGroups) && entry.relatedGroups.length) score += 10;
  return score;
}

function sortRelatedEmails(entries: RelatedEmailEntry[]): RelatedEmailEntry[] {
  return [...(entries || [])].sort((a, b) => {
    const scoreDiff = scoreEmail(b) - scoreEmail(a);
    if (scoreDiff) return scoreDiff;
    const dateA = normalizeText(a.messageDateIso || a.receivedAtIso || a.sentAtIso);
    const dateB = normalizeText(b.messageDateIso || b.receivedAtIso || b.sentAtIso);
    return dateB.localeCompare(dateA);
  });
}

function buildEmailSection(entries: RelatedEmailEntry[], limit: number, excerptLength: number): string {
  const rows = sortRelatedEmails(entries).slice(0, limit);
  if (!rows.length) return "Sem emails adicionais relacionados guardados.";
  return rows
    .map((email, index) => {
      const headerParts = [
        formatDateLabel(email.messageDateIso || email.receivedAtIso || email.sentAtIso),
        normalizeText(email.fromName || email.fromEmail),
        normalizeText(email.subject || "(sem assunto)"),
      ].filter(Boolean);
      const reason = summarizeReasons(email);
      const excerptSource = normalizeText(email.bodyText) || htmlToPlainText(normalizeText(email.bodyHtml));
      const lines = [
        `${index + 1}. ${headerParts.join(" | ")}`.trim(),
      ];
      if (reason) lines.push(`   Motivo da relevância: ${reason}`);
      if (excerptSource) lines.push(`   Excerto: ${truncate(excerptSource, excerptLength)}`);
      return lines.join("\n");
    })
    .join("\n");
}

function buildLinkedRecordEmailSection(entries: LinkEntry[], limit: number): string {
  const rows = dedupeLinkEntries(entries).slice(0, limit);
  if (!rows.length) return "Sem outras conversas ligadas aos mesmos registos.";
  return rows
    .map((entry, index) => {
      const recordName = normalizeText(entry.recordName || entry.name || entry.title);
      const model = normalizeText(entry.model);
      const headerParts = [
        formatDateLabel(entry.messageDateIso || entry.receivedAtIso || entry.sentAtIso || entry.linkedAt),
        normalizeText(entry.fromName || entry.fromEmail),
        normalizeText(entry.subject || "(sem assunto)"),
      ].filter(Boolean);
      return `${index + 1}. ${headerParts.join(" | ")}\n   Registo ligado: ${[recordName, model].filter(Boolean).join(" | ")}`;
    })
    .join("\n");
}

function buildGroupSection(groups: LinkGroupEntry[]): string {
  const rows = (groups || []).slice(0, 10);
  if (!rows.length) return "Sem grupos ligados.";
  return rows.map((group) => {
    const bits = [
      normalizeText(group.name),
      normalizeText(group.status),
      Array.isArray(group.labels) && group.labels.length ? `etiquetas: ${group.labels.join(", ")}` : "",
    ].filter(Boolean);
    return `- ${bits.join(" | ")}`;
  }).join("\n");
}

function buildTicketSection(tickets: GroupTicketEntry[]): string {
  const rows = (tickets || []).slice(0, 10);
  if (!rows.length) return "Sem tickets ligados.";
  return rows.map((ticket) => {
    const bits = [
      normalizeText(ticket.code),
      normalizeText(ticket.title),
      normalizeText(ticket.status),
      Array.isArray(ticket.labels) && ticket.labels.length ? `etiquetas: ${ticket.labels.join(", ")}` : "",
    ].filter(Boolean);
    return `- ${bits.join(" | ")}`;
  }).join("\n");
}

function buildLinkedRecordSection(records: Array<{ model: string; recordId: number; recordName: string }>): string {
  const rows = (records || []).slice(0, 10);
  if (!rows.length) return "Sem registos Odoo/CRM ligados.";
  return rows
    .map((record) => `- ${[record.recordName, `${record.model}#${record.recordId}`].filter(Boolean).join(" | ")}`)
    .join("\n");
}

function buildAttachmentSection(attachments: OutlookAttachment[]): string {
  const rows = (attachments || []).filter((entry) => normalizeText(entry?.name)).slice(0, 12);
  if (!rows.length) return "Sem anexos relevantes no email atual.";
  return rows
    .map((entry) => `- ${normalizeText(entry.name)}${entry.size ? ` (${Math.round(Number(entry.size) / 1024)} KB)` : ""}`)
    .join("\n");
}

function buildCurrentEmailSection(input: BuildAiContextBundleInput): string {
  const effectiveBody = normalizeText(input.bodyText) || htmlToPlainText(normalizeText(input.bodyHtml));
  return [
    `Assunto: ${normalizeText(input.ctx.subject || "(sem assunto)")}`,
    `De: ${normalizeText(input.ctx.fromName || input.ctx.fromEmail)}`,
    `Para: ${(input.ctx.toRecipients || []).map((entry) => normalizeText(entry.email)).filter(Boolean).join("; ") || "--"}`,
    `Cc: ${(input.ctx.ccRecipients || []).map((entry) => normalizeText(entry.email)).filter(Boolean).join("; ") || "--"}`,
    `Data: ${formatDateLabel(input.ctx.receivedDateTimeIso) || "--"}`,
    effectiveBody ? `Texto atual: ${truncate(effectiveBody, 1200)}` : "Texto atual: --",
  ].join("\n");
}

function buildContextText(
  input: BuildAiContextBundleInput,
  data: {
    groups: LinkGroupEntry[];
    tickets: GroupTicketEntry[];
    relatedEmails: RelatedEmailEntry[];
    linkedRecordEmails: LinkEntry[];
    linkedRecords: Array<{ model: string; recordId: number; recordName: string }>;
  },
  options: { emailLimit: number; excerptLength: number; linkedRecordEmailLimit: number }
): string {
  const sections = [
    `EMAIL ATUAL\n${buildCurrentEmailSection(input)}`,
    `REGISTOS LIGADOS NO ODOO/CRM\n${buildLinkedRecordSection(data.linkedRecords)}`,
    `GRUPOS LIGADOS\n${buildGroupSection(data.groups)}`,
    `TICKETS LIGADOS\n${buildTicketSection(data.tickets)}`,
    `ANEXOS DO EMAIL ATUAL\n${buildAttachmentSection(input.attachments || [])}`,
    `EMAILS RELACIONADOS AO MESMO TEMA\n${buildEmailSection(data.relatedEmails, options.emailLimit, options.excerptLength)}`,
    `OUTRAS CONVERSAS LIGADAS AOS MESMOS REGISTOS\n${buildLinkedRecordEmailSection(data.linkedRecordEmails, options.linkedRecordEmailLimit)}`,
  ];
  return sections.join("\n\n").trim();
}

export async function buildAiContextBundle(input: BuildAiContextBundleInput): Promise<AiContextBundle> {
  const payload = {
    itemId: normalizeText(input.ctx.itemId),
    internetMessageId: normalizeText(input.ctx.internetMessageId),
    conversationId: normalizeText(input.ctx.conversationId),
    subject: normalizeText(input.ctx.subject),
    fromEmail: normalizeText(input.ctx.fromEmail),
    fromName: normalizeText(input.ctx.fromName),
    receivedAtIso: normalizeText(input.ctx.receivedDateTimeIso),
    messageDateIso: normalizeText(input.ctx.receivedDateTimeIso),
    bodyText: normalizeText(input.bodyText),
    bodyHtml: normalizeText(input.bodyHtml),
  };

  const related = await getRelatedEmailContext(payload);
  const linkedRecords = dedupeLinkedRecords(input.links || []);
  const relatedEmailIds = new Set(dedupeRelatedEmails(related.emails || []).map((entry) => makeEmailIdentity(entry)));
  const currentEmailIdentity = makeEmailIdentity({
    itemId: input.ctx.itemId,
    internetMessageId: input.ctx.internetMessageId,
    conversationId: input.ctx.conversationId,
    subject: input.ctx.subject,
    messageDateIso: input.ctx.receivedDateTimeIso,
    receivedAtIso: input.ctx.receivedDateTimeIso,
  });
  const recordEmailRowsNested = await Promise.all(
    linkedRecords.slice(0, 6).map(async (record) => {
      try {
        return await getLinksByRecord(record.model, record.recordId);
      } catch {
        return [];
      }
    })
  );
  const linkedRecordEmails = dedupeLinkEntries(recordEmailRowsNested.flat().filter((entry) => {
    const identity = makeEmailIdentity(entry);
    return Boolean(identity) && identity !== currentEmailIdentity && !relatedEmailIds.has(identity);
  }));

  const relatedEmails = dedupeRelatedEmails(related.emails || []);
  const cacheKey = [
    normalizeText(input.ctx.conversationId),
    normalizeText(input.ctx.itemId),
    normalizeText(input.ctx.internetMessageId),
    ...linkedRecords.map((record) => `${record.model}:${record.recordId}`),
    ...(related.groups || []).map((group) => String(group.id || "")),
    ...(related.tickets || []).map((ticket) => String(ticket.id || ticket.code || "")),
    ...relatedEmails.slice(0, 12).map((entry) => makeEmailIdentity(entry)),
  ].filter(Boolean).join("|");

  return {
    cacheKey,
    groups: Array.isArray(related.groups) ? related.groups : [],
    tickets: Array.isArray(related.tickets) ? related.tickets : [],
    relatedEmails,
    linkedRecordEmails,
    linkedRecords,
    promptContext: buildContextText(input, {
      groups: Array.isArray(related.groups) ? related.groups : [],
      tickets: Array.isArray(related.tickets) ? related.tickets : [],
      relatedEmails,
      linkedRecordEmails,
      linkedRecords,
    }, {
      emailLimit: 8,
      excerptLength: 360,
      linkedRecordEmailLimit: 8,
    }),
    briefingContext: buildContextText(input, {
      groups: Array.isArray(related.groups) ? related.groups : [],
      tickets: Array.isArray(related.tickets) ? related.tickets : [],
      relatedEmails,
      linkedRecordEmails,
      linkedRecords,
    }, {
      emailLimit: 12,
      excerptLength: 520,
      linkedRecordEmailLimit: 12,
    }),
  };
}
