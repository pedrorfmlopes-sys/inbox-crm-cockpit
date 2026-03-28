import React, { useEffect, useMemo, useState } from "react";
import { CockpitProvider, useCockpit } from "@/components/shell/CockpitProvider";
import { addEmailToLinkGroup, createGroupTicket, createLinkGroup, getRelatedEmailContext, linkEmailToGroupTicket, listLinkGroups, listGroupTicketSeries, saveGroupDocuments, searchGroupTickets, searchKnownEmails, updateLinkGroup, type GroupTicketEntry, type GroupTicketSeriesEntry, type LinkGroupEntry, type RelatedEmailEntry, type RelevantEmailPayload } from "@/api";
import { requestCockpitHostAction, syncManagedOutlookCategories } from "@/office";
import { getSettings } from "@/settings";
import { PanelState } from "@/ui/PanelState";
import { applySkin } from "@/ui/skins";
import * as Icons from "@/ui/icons";
import "../../global.css";

type SectionId = "emails" | "classification" | "labels" | "filters" | "summary";
type ScopeMode = "related" | "all";
type LabelDraft = { categorize: boolean; hasStatus: boolean };
type CaseGroupEntry = LinkGroupEntry & { relationKind?: string };
type StudioParams = {
  conversationId?: string;
  internetMessageId?: string;
  itemId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  receivedAtIso?: string;
};

const MENU: Array<{ id: SectionId; label: string; icon: React.ReactNode; help: string }> = [
  { id: "emails", label: "Emails", icon: <Icons.MessageSquare size={15} />, help: "Lista e preview base do caso." },
  { id: "classification", label: "Classificacao", icon: <Icons.Target size={15} />, help: "Grupo principal, referencias e ticket." },
  { id: "labels", label: "Etiquetas", icon: <Icons.Star size={15} />, help: "Etiquetas e futuras categorias Outlook." },
  { id: "filters", label: "Filtros", icon: <Icons.Search size={15} />, help: "Reducao da lista e testes de vista." },
  { id: "summary", label: "Resumo", icon: <Icons.Clipboard size={15} />, help: "Fotografia do que esta preparado." },
];

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.emailKey || email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function dedupeEmails(emails: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const seen = new Set<string>();
  return emails.filter((email) => {
    const key = makeEmailKey(email);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
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

function buildEmailPreviewHtml(email: RelatedEmailEntry | null): string {
  const html = String(email?.bodyHtml || "").trim();
  if (html) {
    return `<!doctype html><html><head><meta charset="utf-8" /><style>html,body{margin:0;padding:0;background:#fff;color:#172b4d;font:14px/1.5 'Segoe UI',sans-serif}body{padding:18px}img{max-width:100%;height:auto}table{max-width:100%}blockquote{margin-left:0;padding-left:12px;border-left:3px solid #dbeafe;color:#475569}pre{white-space:pre-wrap;word-break:break-word}</style></head><body>${html}</body></html>`;
  }
  const text = String(email?.bodyText || "").trim();
  if (!text) return "";
  return `<!doctype html><html><head><meta charset="utf-8" /><style>html,body{margin:0;padding:0;background:#fff;color:#172b4d;font:14px/1.55 'Segoe UI',sans-serif}body{padding:18px}pre{margin:0;white-space:pre-wrap;word-break:break-word;font:inherit}</style></head><body><pre>${escapeHtml(text)}</pre></body></html>`;
}

function buildSnippet(email: RelatedEmailEntry): string {
  const source = String(email.bodyText || "").trim() || htmlToPlainText(String(email.bodyHtml || ""));
  return source.length > 180 ? `${source.slice(0, 177).trim()}...` : source;
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

function makeAttachmentKey(attachment: { id?: string; name?: string; contentId?: string }): string {
  return String(attachment.id || attachment.contentId || attachment.name || "").trim();
}

function derivePartnerName(email: RelatedEmailEntry | null): string {
  const fromName = String(email?.fromName || "").trim();
  if (fromName) return fromName;
  const fromEmail = String(email?.fromEmail || "").trim().toLowerCase();
  const domain = fromEmail.includes("@") ? fromEmail.split("@")[1] : "";
  const base = domain.split(".")[0] || "";
  return base ? base.charAt(0).toUpperCase() + base.slice(1) : "";
}

function detectCaseType(text: string): string {
  const value = text.toLowerCase();
  if (/(reclam|inciden|nao conforme|defeito)/.test(value)) return "reclamacao";
  if (/(pedido|encomenda|order|po\b|purchase order|material listo)/.test(value)) return "pedido/encomenda";
  if (/(proposta|orcamento|quote|quotation)/.test(value)) return "proposta";
  if (/(projeto|project|obra|worksite)/.test(value)) return "projeto";
  return "geral";
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
  };
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
  const { ctx, attachments } = useCockpit();
  const params = useMemo(() => readParams(), []);
  const [section, setSection] = useState<SectionId>("emails");
  const [scopeMode, setScopeMode] = useState<ScopeMode>("related");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [groupFilterId, setGroupFilterId] = useState("");
  const [emailSearch, setEmailSearch] = useState("");
  const [onlyExternal, setOnlyExternal] = useState(false);
  const [onlyWithAttachments, setOnlyWithAttachments] = useState(false);
  const [allGroups, setAllGroups] = useState<LinkGroupEntry[]>([]);
  const [currentCaseGroups, setCurrentCaseGroups] = useState<CaseGroupEntry[]>([]);
  const [ticketSeries, setTicketSeries] = useState<GroupTicketSeriesEntry[]>([]);
  const [relatedTickets, setRelatedTickets] = useState<GroupTicketEntry[]>([]);
  const [relatedEmails, setRelatedEmails] = useState<RelatedEmailEntry[]>([]);
  const [knownEmails, setKnownEmails] = useState<RelatedEmailEntry[]>([]);
  const [selectedEmailKey, setSelectedEmailKey] = useState("");
  const [principalGroupId, setPrincipalGroupId] = useState("");
  const [referenceGroupIds, setReferenceGroupIds] = useState<string[]>([]);
  const [selectedSeriesId, setSelectedSeriesId] = useState("");
  const [selectedTicketId, setSelectedTicketId] = useState("");
  const [ticketSearch, setTicketSearch] = useState("");
  const [ticketSearchResults, setTicketSearchResults] = useState<GroupTicketEntry[]>([]);
  const [labelInput, setLabelInput] = useState("");
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [labelDrafts, setLabelDrafts] = useState<Record<string, LabelDraft>>({});
  const [createGroupName, setCreateGroupName] = useState("");
  const [createTicketTitle, setCreateTicketTitle] = useState("");
  const [attachmentPlan, setAttachmentPlan] = useState<Record<string, { analyze: boolean; save: boolean; forward: boolean }>>({});
  const [actionBusy, setActionBusy] = useState(false);

  const currentSeed = useMemo(() => buildFallbackEmail(params), [params]);
  const currentContext = useMemo(() => ({
    conversationId: String(ctx.conversationId || params.conversationId || "").trim(),
    internetMessageId: String(ctx.internetMessageId || params.internetMessageId || "").trim(),
    itemId: String(ctx.itemId || params.itemId || "").trim(),
    subject: String(ctx.subject || params.subject || "").trim(),
    fromEmail: String(ctx.fromEmail || params.fromEmail || "").trim(),
    fromName: String(ctx.fromName || params.fromName || "").trim(),
    receivedAtIso: String(ctx.receivedDateTimeIso || params.receivedAtIso || "").trim(),
  }), [ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject, params]);

  useEffect(() => {
    void (async () => {
      try {
        const settings = await getSettings();
        applySkin(settings.skinId || "soft");
      } catch {
        applySkin("soft");
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
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
        const contextualEmails = dedupeEmails([
          ...(related.email ? [related.email] : []),
          ...(related.emails || []),
          ...(currentSeed ? [currentSeed] : []),
        ]);
        const mergedEmails = dedupeEmails([...contextualEmails, ...(emails || [])]);
        setAllGroups(mergedGroups);
        setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
        setTicketSeries(Array.isArray(series) ? series : []);
        setRelatedTickets(Array.isArray(related.tickets) ? related.tickets : []);
        setRelatedEmails(contextualEmails);
        setKnownEmails(Array.isArray(emails) ? emails : []);
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
        setStatus(mergedEmails.length
          ? "Janela base pronta. O email atual e os relacionados ja podem ser analisados aqui."
          : "Ainda nao encontrámos emails relacionados. Esta janela vai usar o email atual como ponto de partida.");
      } catch (fetchError: any) {
        if (!cancelled) setError(String(fetchError?.message || fetchError || "Falha a preparar o studio de classificacao."));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [currentContext.conversationId, currentContext.fromEmail, currentContext.fromName, currentContext.internetMessageId, currentContext.itemId, currentContext.receivedAtIso, currentContext.subject, currentSeed]);

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

  const visibleEmails = useMemo(() => {
    const q = String(emailSearch || "").trim().toLowerCase();
    return [...emailPool]
      .sort((a, b) => String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || "")))
      .filter((email) => {
        if (onlyExternal && !isExternalEmail(email)) return false;
        if (onlyWithAttachments && !(Array.isArray(email.attachments) && email.attachments.length)) return false;
        if (groupFilterId) {
          const relatedGroupIds = new Set([email.groupId, ...(email.relatedGroups || []).map((entry) => entry.id)].filter(Boolean));
          if (!relatedGroupIds.has(groupFilterId)) return false;
        }
        if (!q) return true;
        const haystack = [email.subject, email.fromName, email.fromEmail, buildSnippet(email)].join(" ").toLowerCase();
        return haystack.includes(q);
      });
  }, [emailPool, emailSearch, groupFilterId, onlyExternal, onlyWithAttachments]);

  const selectedEmail = useMemo(
    () => visibleEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || emailPool.find((email) => makeEmailKey(email) === selectedEmailKey) || visibleEmails[0] || emailPool[0] || null,
    [emailPool, selectedEmailKey, visibleEmails]
  );

  const selectedEmailGroups = useMemo(() => {
    if (!selectedEmail) return [];
    const fallbackCurrentGroups = selectedEmailIsCurrent ? currentCaseBusinessGroups.map((group) => ({
      id: group.id,
      name: group.name,
      relationKind: group.relationKind,
      kind: group.kind,
    })) : [];
    const list = [
      ...(selectedEmail.relatedGroups || []),
      ...(selectedEmail.groupId ? [{ id: selectedEmail.groupId, name: selectedEmail.groupName, relationKind: selectedEmail.membershipKind }] : []),
      ...fallbackCurrentGroups,
    ];
    return list.reduce<Array<{ id: string; name?: string; relationKind?: string }>>((acc, row) => {
      if (!row?.id || acc.some((entry) => entry.id === row.id)) return acc;
      const groupKind = String((row as any)?.kind || groupMap.get(row.id)?.kind || "").trim().toLowerCase();
      if (groupKind === "conversation") return acc;
      acc.push(row);
      return acc;
    }, []);
  }, [currentCaseBusinessGroups, groupMap, selectedEmail, selectedEmailIsCurrent]);

  useEffect(() => {
    if (!selectedEmail) return;
    setPrincipalGroupId((current) => {
      if (current) return current;
      const principal = selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal");
      return principal?.id || "";
    });
    setReferenceGroupIds((current) => {
      if (current.length) return current;
      return selectedEmailGroups.filter((group) => String(group.relationKind || "").toLowerCase() !== "principal").map((group) => group.id);
    });
  }, [selectedEmail, selectedEmailGroups]);

  useEffect(() => {
    if (!principalGroupId) return;
    setReferenceGroupIds((current) => current.filter((groupId) => groupId !== principalGroupId));
  }, [principalGroupId]);

  const previewHtml = useMemo(() => buildEmailPreviewHtml(selectedEmail), [selectedEmail]);
  const labelCatalog = useMemo(() => {
    const values = new Set<string>();
    allGroups.forEach((group) => (group.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    relatedTickets.forEach((ticket) => (ticket.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    selectedLabels.forEach((label) => values.add(label));
    return Array.from(values).sort((a, b) => a.localeCompare(b, "pt"));
  }, [allGroups, relatedTickets, selectedLabels]);
  const filteredLabelCatalog = useMemo(() => {
    const q = String(labelInput || "").trim().toLowerCase();
    return q ? labelCatalog.filter((label) => label.toLowerCase().includes(q)) : labelCatalog;
  }, [labelCatalog, labelInput]);
  const availableTicketChoices = useMemo(() => {
    const rows = [...relatedTickets, ...ticketSearchResults].reduce<GroupTicketEntry[]>((acc, ticket) => {
      if (!ticket?.id || acc.some((entry) => entry.id === ticket.id)) return acc;
      acc.push(ticket);
      return acc;
    }, []);
    return rows.sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")));
  }, [relatedTickets, ticketSearchResults]);
  const selectedEmailIsCurrent = useMemo(() => {
    const selectedItemId = String(selectedEmail?.itemId || "").trim();
    const currentItemId = String(ctx.itemId || "").trim();
    if (selectedItemId && currentItemId && selectedItemId === currentItemId) return true;
    const selectedMessageId = String(selectedEmail?.internetMessageId || "").trim().toLowerCase();
    const currentMessageId = String(ctx.internetMessageId || "").trim().toLowerCase();
    return Boolean(selectedMessageId && currentMessageId && selectedMessageId === currentMessageId);
  }, [ctx.internetMessageId, ctx.itemId, selectedEmail?.internetMessageId, selectedEmail?.itemId]);

  const selectedEmailAttachments = useMemo(() => {
    const source = selectedEmailIsCurrent
      ? attachments.map((attachment) => ({ ...attachment }))
      : (selectedEmail?.attachments || []).map((attachment) => ({
          id: attachment.id,
          name: attachment.name,
          contentType: String(attachment.contentType || "application/octet-stream"),
          content: String(attachment.content || ""),
          size: attachment.size,
          isInline: attachment.isInline,
          contentId: attachment.contentId,
        }));
    return source.filter((attachment) => String(attachment.name || "").trim());
  }, [attachments, selectedEmail?.attachments, selectedEmailIsCurrent]);

  useEffect(() => {
    setAttachmentPlan((current) => {
      const next = { ...current };
      for (const attachment of selectedEmailAttachments) {
        const key = makeAttachmentKey(attachment);
        if (!key || next[key]) continue;
        const contentType = String(attachment.contentType || "").toLowerCase();
        const isDocument = /pdf|image|excel|spreadsheet|word|officedocument|text|csv/.test(contentType) || /\.(pdf|png|jpe?g|xlsx?|docx?|csv|txt)$/i.test(String(attachment.name || ""));
        next[key] = { analyze: isDocument, save: false, forward: false };
      }
      return next;
    });
  }, [selectedEmailAttachments]);

  const detectionText = useMemo(() => {
    const attachmentNames = selectedEmailAttachments.map((attachment) => attachment.name).join(" ");
    return [
      selectedEmail?.subject,
      selectedEmail?.fromName,
      selectedEmail?.fromEmail,
      selectedEmail?.bodyText,
      htmlToPlainText(String(selectedEmail?.bodyHtml || "")),
      attachmentNames,
    ].filter(Boolean).join(" ");
  }, [selectedEmail?.bodyHtml, selectedEmail?.bodyText, selectedEmail?.fromEmail, selectedEmail?.fromName, selectedEmail?.subject, selectedEmailAttachments]);

  const detectedCaseType = useMemo(() => detectCaseType(detectionText), [detectionText]);
  const detectedReferences = useMemo(() => detectReferences(detectionText), [detectionText]);
  const suggestedExistingGroups = useMemo(() => splitSuggestions(allGroups, detectionText), [allGroups, detectionText]);
  const suggestedLabelSeeds = useMemo(() => {
    const values = new Set<string>();
    if (detectedCaseType !== "geral") values.add(`tipo:${detectedCaseType}`);
    for (const ref of detectedReferences) values.add(ref);
    const partner = derivePartnerName(selectedEmail);
    if (partner) values.add(partner);
    return Array.from(values).slice(0, 8);
  }, [detectedCaseType, detectedReferences, selectedEmail]);

  const suggestedGroupName = useMemo(() => {
    const partner = derivePartnerName(selectedEmail);
    if (detectedReferences.length && partner) return `${partner} / ${detectedReferences[0]}`;
    if (detectedReferences.length) return detectedReferences[0];
    if (partner && detectedCaseType !== "geral") return `${partner} / ${detectedCaseType}`;
    return partner || String(selectedEmail?.subject || "").trim().slice(0, 72);
  }, [detectedCaseType, detectedReferences, selectedEmail]);

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
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId,
      content: attachment.content,
    })),
  }), [currentContext.conversationId, currentContext.fromEmail, currentContext.fromName, currentContext.internetMessageId, currentContext.itemId, currentContext.receivedAtIso, currentContext.subject, selectedEmail?.bodyHtml, selectedEmail?.bodyText, selectedEmail?.conversationId, selectedEmail?.fromEmail, selectedEmail?.fromName, selectedEmail?.internetMessageId, selectedEmail?.itemId, selectedEmail?.messageDateIso, selectedEmail?.receivedAtIso, selectedEmail?.subject, selectedEmailAttachments]);
  const selectedTicket = useMemo(() => availableTicketChoices.find((ticket) => ticket.id === selectedTicketId) || relatedTickets.find((ticket) => ticket.id === selectedTicketId) || null, [availableTicketChoices, relatedTickets, selectedTicketId]);

  useEffect(() => {
    setSelectedTicketId((current) => {
      if (current && availableTicketChoices.some((ticket) => ticket.id === current)) return current;
      if (relatedTickets.length === 1) return relatedTickets[0].id;
      return current || "";
    });
  }, [availableTicketChoices, relatedTickets]);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (!closed) window.close();
  }

  async function refreshSelectedEmailContext() {
    const related = await getRelatedEmailContext({
      conversationId: currentEmailPayload.conversationId,
      internetMessageId: currentEmailPayload.internetMessageId,
      itemId: currentEmailPayload.itemId,
      subject: currentEmailPayload.subject,
      fromEmail: currentEmailPayload.fromEmail,
      fromName: currentEmailPayload.fromName,
      receivedAtIso: currentEmailPayload.receivedAtIso,
    });
    const nextGroups = [...allGroups, ...(related.groups || [])].reduce<LinkGroupEntry[]>((acc, group) => {
      if (!group?.id || acc.some((entry) => entry.id === group.id)) return acc;
      acc.push(group);
      return acc;
    }, []);
    const contextualEmails = dedupeEmails([
      ...(related.email ? [related.email] : []),
      ...(related.emails || []),
      ...(currentSeed ? [currentSeed] : []),
    ]);
    setAllGroups(nextGroups);
    setCurrentCaseGroups(Array.isArray(related.groups) ? related.groups as CaseGroupEntry[] : []);
    setRelatedTickets(Array.isArray(related.tickets) ? related.tickets : []);
    setRelatedEmails(contextualEmails);
  }

  function toggleReferenceGroup(groupId: string) {
    setReferenceGroupIds((current) => current.includes(groupId) ? current.filter((entry) => entry !== groupId) : [...current, groupId]);
  }

  function addLabel(label: string) {
    const value = String(label || "").trim();
    if (!value) return;
    setSelectedLabels((current) => current.includes(value) ? current : [...current, value]);
    setLabelDrafts((current) => current[value] ? current : { ...current, [value]: { categorize: false, hasStatus: false } });
    setLabelInput("");
  }

  function updateLabelDraft(label: string, patch: Partial<LabelDraft>) {
    setLabelDrafts((current) => ({ ...current, [label]: { categorize: current[label]?.categorize ?? false, hasStatus: current[label]?.hasStatus ?? false, ...patch } }));
  }

  function removeLabel(label: string) {
    setSelectedLabels((current) => current.filter((entry) => entry !== label));
  }

  async function handleCreateGroupAndLink() {
    const name = String(createGroupName || "").trim();
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
        membershipKind: "principal",
      });
      setAllGroups((current) => current.some((entry) => entry.id === created.id) ? current : [created, ...current]);
      setPrincipalGroupId(created.id);
      await refreshSelectedEmailContext();
      setStatus(`Grupo "${created.name}" criado e email ligado como principal.`);
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
        email: currentEmailPayload,
        membershipKind: principalGroupId ? "principal" : "referencia",
      });
      setRelatedTickets((current) => [ticket, ...current.filter((entry) => entry.id !== ticket.id)]);
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

  async function handleSaveSelectedAttachments() {
    if (!principalGroupId) {
      setStatus("Escolhe primeiro um grupo principal para guardar documentos.");
      return;
    }
    const docs = selectedEmailAttachments
      .filter((attachment) => attachmentPlan[makeAttachmentKey(attachment)]?.save)
      .filter((attachment) => String(attachment.content || "").trim())
      .map((attachment) => ({
        name: attachment.name,
        contentType: attachment.contentType,
        contentBase64: attachment.content,
        size: attachment.size,
        sourceEmailKey: makeEmailKey(selectedEmail || {}),
        sourceItemId: currentEmailPayload.itemId,
        sourceInternetMessageId: currentEmailPayload.internetMessageId,
        sourceConversationId: currentEmailPayload.conversationId,
        sourceEmailSubject: currentEmailPayload.subject,
      }));
    if (!docs.length) {
      setStatus("Nao ha anexos com conteudo selecionados para guardar.");
      return;
    }
    setActionBusy(true);
    try {
      await saveGroupDocuments(principalGroupId, { documents: docs });
      await refreshSelectedEmailContext();
      setStatus(`${docs.length} anexo(s) guardado(s) nos documentos do grupo principal.`);
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel guardar os anexos no grupo.");
    } finally {
      setActionBusy(false);
    }
  }

  async function handleSearchTickets() {
    setActionBusy(true);
    try {
      const rows = await searchGroupTickets({
        q: String(ticketSearch || "").trim() || undefined,
        groupId: principalGroupId || undefined,
        email: currentEmailPayload,
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

  async function handleApplyClassification() {
    setActionBusy(true);
    try {
      const principalGroup = principalGroupId ? groupMap.get(principalGroupId) || null : null;
      const referenceGroups = referenceGroupIds.map((groupId) => groupMap.get(groupId)).filter(Boolean) as LinkGroupEntry[];
      const allGroupIds = [principalGroupId, ...referenceGroupIds].filter(Boolean);

      if (principalGroupId) {
        await addEmailToLinkGroup(principalGroupId, {
          ...currentEmailPayload,
          membershipKind: "principal",
        });
      }
      for (const groupId of referenceGroupIds) {
        await addEmailToLinkGroup(groupId, {
          ...currentEmailPayload,
          membershipKind: "referencia",
        });
      }

      if (principalGroup && selectedLabels.length) {
        await updateLinkGroup(principalGroup.id, {
          name: principalGroup.name,
          description: principalGroup.description,
          documentsEnabled: principalGroup.documentsEnabled,
          status: principalGroup.status,
          isArchived: principalGroup.isArchived,
          labels: mergeLabels(principalGroup.labels || [], selectedLabels),
        });
      }

      let finalTicket: GroupTicketEntry | null = null;
      if (selectedTicketId) {
        const linked = await linkEmailToGroupTicket(selectedTicketId, {
          email: currentEmailPayload,
          applyGroups: allGroupIds.length > 0,
          groupIds: allGroupIds,
          membershipKind: principalGroupId ? "principal" : "referencia",
        });
        finalTicket = linked.ticket;
      } else if (selectedSeriesId) {
        finalTicket = await createGroupTicket({
          seriesId: selectedSeriesId,
          title: String(createTicketTitle || selectedEmail?.subject || "Ticket").trim(),
          description: String(selectedEmail?.bodyText || "").trim().slice(0, 4000),
          labels: selectedLabels,
          groupIds: allGroupIds,
          email: currentEmailPayload,
          membershipKind: principalGroupId ? "principal" : "referencia",
        });
        setRelatedTickets((current) => [finalTicket as GroupTicketEntry, ...current.filter((entry) => entry.id !== finalTicket?.id)]);
        setSelectedTicketId(finalTicket.id);
      }

      if (selectedEmailIsCurrent) {
        await syncManagedOutlookCategories({
          groupNames: principalGroup ? [principalGroup.name] : [],
          ticketCodes: finalTicket?.code ? [finalTicket.code] : [],
          statuses: finalTicket?.status ? [finalTicket.status] : [],
        }).catch(() => undefined);
      }

      await refreshSelectedEmailContext();
      setStatus("Classificacao aplicada ao email selecionado.");
    } catch (actionError: any) {
      setStatus(actionError?.message || "Nao foi possivel aplicar a classificacao.");
    } finally {
      setActionBusy(false);
    }
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
              <span>{detectedCaseType}</span>
            </div>
            {selectedEmailGroups.length ? <div style={S.chips}>{selectedEmailGroups.map((group) => <span key={group.id} style={S.groupChip}>{group.name || groupMap.get(group.id)?.name || group.id}</span>)}</div> : null}
            {previewHtml ? <iframe title={selectedEmail.subject || "Preview"} srcDoc={previewHtml} style={S.preview} sandbox="" /> : <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado suficiente para preview." />}
          </div>

          <div style={S.grid2Wide}>
            <div style={S.card}>
              <div style={S.cardTitle}>Deteccoes e sugestoes</div>
              <div style={S.cardMeta}>Leitura inicial com base em assunto, corpo, nomes de anexos e contexto ja guardado.</div>
              <div style={S.summaryRow}><span>Tipo detetado</span><strong>{detectedCaseType}</strong></div>
              <div style={S.summaryRow}><span>Parceiro detetado</span><strong>{derivePartnerName(selectedEmail) || "--"}</strong></div>
              <div style={S.summaryRow}><span>Referencias detetadas</span><strong>{detectedReferences.length ? detectedReferences.join(", ") : "--"}</strong></div>
              <div style={S.summaryRow}><span>Sugestao de grupo</span><strong>{suggestedGroupName || "--"}</strong></div>
              {suggestedExistingGroups.length ? (
                <>
                  <div style={S.subTitle}>Grupos sugeridos</div>
                  <div style={S.chips}>
                    {suggestedExistingGroups.map((group) => (
                      <button key={group.id} type="button" style={group.id === principalGroupId ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => setPrincipalGroupId(group.id)}>
                        {group.name}
                      </button>
                    ))}
                  </div>
                </>
              ) : null}
              {suggestedLabelSeeds.length ? (
                <>
                  <div style={S.subTitle}>Etiquetas sugeridas</div>
                  <div style={S.chips}>
                    {suggestedLabelSeeds.map((label) => (
                      <button key={label} type="button" style={selectedLabels.includes(label) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => addLabel(label)}>
                        {label}
                      </button>
                    ))}
                  </div>
                </>
              ) : null}
            </div>

            <div style={S.card}>
              <div style={S.cardTitle}>Criacao rapida</div>
              <div style={S.cardMeta}>Ja comecamos aqui a criar e ligar grupos ou tickets sem sair da janela.</div>
              <label style={S.field}>
                <span style={S.label}>Novo grupo</span>
                <div style={S.inline}>
                  <input style={S.input} value={createGroupName} onChange={(event) => setCreateGroupName(event.target.value)} placeholder="Nome do grupo" />
                  <button type="button" style={S.secondaryBtn} onClick={() => void handleCreateGroupAndLink()} disabled={actionBusy || !String(createGroupName || "").trim()}>
                    <Icons.Plus size={12} />
                    Criar grupo
                  </button>
                </div>
              </label>
              <label style={S.field}>
                <span style={S.label}>Novo ticket</span>
                <input style={S.input} value={createTicketTitle} onChange={(event) => setCreateTicketTitle(event.target.value)} placeholder="Titulo do ticket" />
              </label>
              <label style={S.field}>
                <span style={S.label}>Serie de ticket</span>
                <div style={S.inline}>
                  <select style={S.select} value={selectedSeriesId} onChange={(event) => setSelectedSeriesId(event.target.value)}>
                    <option value="">Sem ticket/caso</option>
                    {ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}
                  </select>
                  <button type="button" style={S.secondaryBtn} onClick={() => void handleCreateTicketAndLink()} disabled={actionBusy || !selectedSeriesId}>
                    <Icons.Plus size={12} />
                    Criar ticket
                  </button>
                </div>
              </label>
            </div>
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>Anexos deste email</div>
            <div style={S.cardMeta}>Cada anexo pode ser marcado para analisar, guardar no grupo principal ou reenviar mais tarde.</div>
            {selectedEmailAttachments.length ? (
              <>
                <div style={S.attachList}>
                  {selectedEmailAttachments.map((attachment) => {
                    const key = makeAttachmentKey(attachment);
                    const plan = attachmentPlan[key] || { analyze: false, save: false, forward: false };
                    const hasContent = Boolean(String(attachment.content || "").trim());
                    return (
                      <div key={key} style={S.attachRow}>
                        <div style={S.attachMeta}>
                          <strong>{attachment.name}</strong>
                          <small>{attachment.contentType || "ficheiro"}{attachment.size ? ` · ${Math.round(Number(attachment.size || 0) / 1024)} KB` : ""}{hasContent ? "" : " · sem conteudo guardado"}</small>
                        </div>
                        <div style={S.attachChecks}>
                          <label style={S.check}><input type="checkbox" checked={plan.analyze} onChange={(event) => toggleAttachmentPlan(key, "analyze", event.target.checked)} /><span>Analisar</span></label>
                          <label style={S.check}><input type="checkbox" checked={plan.save} onChange={(event) => toggleAttachmentPlan(key, "save", event.target.checked)} disabled={!hasContent} /><span>Guardar</span></label>
                          <label style={S.check}><input type="checkbox" checked={plan.forward} onChange={(event) => toggleAttachmentPlan(key, "forward", event.target.checked)} /><span>Reenviar</span></label>
                        </div>
                      </div>
                    );
                  })}
                </div>
                <div style={S.inline}>
                  <button type="button" style={S.secondaryBtn} onClick={() => void handleSaveSelectedAttachments()} disabled={actionBusy || !principalGroupId}>
                    <Icons.Save size={12} />
                    Guardar no grupo principal
                  </button>
                  <span style={S.cardMeta}>Necessita de grupo principal selecionado e de anexos com conteudo disponivel.</span>
                </div>
              </>
            ) : (
              <PanelState compact tone="info" title="Sem anexos disponiveis" description="Este email nao traz anexos guardados ou ainda nao temos o conteudo disponivel nesta janela." />
            )}
          </div>
        </div>
      );
    }

    if (section === "classification") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Classificacao base</div>
            <div style={S.cardMeta}>Agora ja aplica ao sistema real: grupo principal, referencias e ticket do email selecionado.</div>
            <div style={S.grid2}>
              <label style={S.field}><span style={S.label}>Grupo principal</span><select style={S.select} value={principalGroupId} onChange={(event) => setPrincipalGroupId(event.target.value)}><option value="">Sem grupo principal</option>{businessGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}</select></label>
              <label style={S.field}><span style={S.label}>Ticket existente</span><select style={S.select} value={selectedTicketId} onChange={(event) => setSelectedTicketId(event.target.value)}><option value="">Sem ticket existente</option>{availableTicketChoices.map((ticket) => <option key={ticket.id} value={ticket.id}>{ticket.code} · {ticket.title}</option>)}</select></label>
            </div>
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>Pesquisa de tickets</div>
            <div style={S.inline}>
              <input style={S.input} value={ticketSearch} onChange={(event) => setTicketSearch(event.target.value)} placeholder="Pesquisar por codigo, titulo ou etiqueta" />
              <button type="button" style={S.secondaryBtn} onClick={() => void handleSearchTickets()} disabled={actionBusy}>
                <Icons.Search size={12} />
                Pesquisar
              </button>
            </div>
            <div style={S.grid2}>
              <label style={S.field}><span style={S.label}>Serie para novo ticket</span><select style={S.select} value={selectedSeriesId} onChange={(event) => setSelectedSeriesId(event.target.value)}><option value="">Sem novo ticket</option>{ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}</select></label>
              <label style={S.field}><span style={S.label}>Titulo do ticket</span><input style={S.input} value={createTicketTitle} onChange={(event) => setCreateTicketTitle(event.target.value)} placeholder="Titulo do ticket" /></label>
            </div>
            {selectedTicket ? <div style={S.summaryRow}><span>Ticket selecionado</span><strong>{selectedTicket.code} · {selectedTicket.title}</strong></div> : null}
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>Grupos referencia</div>
            <div style={S.chips}>{businessGroups.filter((group) => group.id !== principalGroupId).map((group) => <button key={group.id} type="button" style={referenceGroupIds.includes(group.id) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleReferenceGroup(group.id)}>{group.name}</button>)}</div>
          </div>

          <div style={S.card}>
            <div style={S.cardTitle}>Aplicar ao email selecionado</div>
            <div style={S.summaryRow}><span>Grupo principal</span><strong>{principalGroupId ? groupMap.get(principalGroupId)?.name || principalGroupId : "--"}</strong></div>
            <div style={S.summaryRow}><span>Grupos referencia</span><strong>{referenceGroupIds.length}</strong></div>
            <div style={S.summaryRow}><span>Ticket</span><strong>{selectedTicket ? selectedTicket.code : (selectedSeriesId ? "Novo ticket a criar" : "--")}</strong></div>
            <div style={S.summaryRow}><span>Etiquetas selecionadas</span><strong>{selectedLabels.length}</strong></div>
            <div style={S.inline}>
              <button type="button" style={S.primaryBtn} onClick={() => void handleApplyClassification()} disabled={actionBusy || (!principalGroupId && !referenceGroupIds.length && !selectedTicketId && !selectedSeriesId)}>
                <Icons.Save size={12} />
                Aplicar classificacao
              </button>
              <span style={S.cardMeta}>No email atual, tambem tenta atualizar as categorias Outlook geridas.</span>
            </div>
          </div>
        </div>
      );
    }

    if (section === "labels") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas estruturadas</div>
            <div style={S.inline}>
              <input style={S.input} value={labelInput} onChange={(event) => setLabelInput(event.target.value)} placeholder="Pesquisar ou criar etiqueta" />
              <button type="button" style={S.secondaryBtn} onClick={() => addLabel(labelInput)} disabled={!String(labelInput || "").trim()}><Icons.Plus size={12} />Adicionar</button>
            </div>
            {filteredLabelCatalog.length ? <div style={S.chips}>{filteredLabelCatalog.slice(0, 24).map((label) => <button key={label} type="button" style={selectedLabels.includes(label) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => addLabel(label)}>{label}</button>)}</div> : null}
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas selecionadas</div>
            {selectedLabels.length ? selectedLabels.map((label) => {
              const draft = labelDrafts[label] || { categorize: false, hasStatus: false };
              return (
                <div key={label} style={S.labelRow}>
                  <div style={S.labelHead}><strong>{label}</strong><button type="button" style={S.linkBtn} onClick={() => removeLabel(label)}>Remover</button></div>
                  <label style={S.check}><input type="checkbox" checked={draft.categorize} onChange={(event) => updateLabelDraft(label, { categorize: event.target.checked })} /><span>Virar categoria Outlook</span></label>
                  <label style={S.check}><input type="checkbox" checked={draft.hasStatus} onChange={(event) => updateLabelDraft(label, { hasStatus: event.target.checked })} /><span>Tem estado associado</span></label>
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
              <label style={S.field}><span style={S.label}>Filtrar por grupo</span><select style={S.select} value={groupFilterId} onChange={(event) => setGroupFilterId(event.target.value)}><option value="">Sem filtro</option>{businessGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}</select></label>
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
            <div style={S.summaryRow}><span>Tickets do caso</span><strong>{relatedTickets.length}</strong></div>
          </div>
        </div>
      );
    }

    return (
      <div style={S.stack}>
        <div style={S.card}>
          <div style={S.cardTitle}>Resumo da estrutura</div>
          <div style={S.summaryRow}><span>Email selecionado</span><strong>{selectedEmail?.subject || "--"}</strong></div>
          <div style={S.summaryRow}><span>Grupo principal</span><strong>{principalGroupId ? groupMap.get(principalGroupId)?.name || principalGroupId : "--"}</strong></div>
          <div style={S.summaryRow}><span>Grupos referencia</span><strong>{referenceGroupIds.length}</strong></div>
          <div style={S.summaryRow}><span>Serie de ticket</span><strong>{selectedSeriesId ? ticketSeries.find((entry) => entry.id === selectedSeriesId)?.prefix || selectedSeriesId : "--"}</strong></div>
          <div style={S.summaryRow}><span>Etiquetas</span><strong>{selectedLabels.length}</strong></div>
          <div style={S.summaryRow}><span>Anexos do email atual</span><strong>{selectedEmailAttachments.length}</strong></div>
        </div>
        <div style={S.note}>Janela nova criada sem alterar o fluxo atual dos grupos. O proximo passo sera ligar estas escolhas ao sistema real de classificacao e categorias.</div>
      </div>
    );
  }

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.mainTitle}>Studio de classificacao</div>
          <div style={S.mainMeta}>Janela nova e isolada para desenhar a futura atribuicao completa de grupos, tickets, etiquetas e filtros.</div>
        </div>
        <button type="button" style={S.secondaryBtn} onClick={handleClose}>Fechar</button>
      </div>

      <div style={S.context}>
        <div><div style={S.kicker}>Email atual</div><div style={S.contextTitle}>{selectedEmail?.subject || currentContext.subject || "(sem assunto)"}</div></div>
        <div style={S.badges}><span style={S.badge}>{selectedEmailAttachments.length} anexo(s)</span><span style={S.badge}>{relatedTickets.length} ticket(s)</span><span style={S.badge}>{relatedEmails.length} relacionados</span></div>
      </div>

      {status ? <div style={S.notice}>{status}</div> : null}

      <div style={S.shell}>
        <aside style={S.sidebar}>
          {MENU.map((item) => (
            <button key={item.id} type="button" style={section === item.id ? S.menuOn : S.menu} onClick={() => setSection(item.id)}>
              <span>{item.icon}</span>
              <span style={{ display: "grid", gap: 2, textAlign: "left" }}><strong>{item.label}</strong><small>{item.help}</small></span>
            </button>
          ))}
        </aside>

        <section style={S.listCol}>
          <div style={S.colTitle}>Emails</div>
          <input style={S.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Pesquisar por assunto, remetente ou texto..." />
          <div style={S.listBody}>
            {loading ? <PanelState compact tone="loading" title="A carregar emails" description="A preparar a lista desta nova janela." /> : null}
            {!loading && !visibleEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Ajusta os filtros ou muda a fonte da lista." /> : null}
            {!loading && visibleEmails.map((email) => (
              <button key={makeEmailKey(email)} type="button" style={makeEmailKey(email) === makeEmailKey(selectedEmail || {}) ? S.emailOn : S.email} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                <div style={S.emailTop}><strong>{email.subject || "(sem assunto)"}</strong>{Array.isArray(email.attachments) && email.attachments.length ? <span style={S.counter}>{email.attachments.length}</span> : null}</div>
                <div style={S.emailMeta}>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</div>
                <div style={S.emailSnippet}>{buildSnippet(email) || "Sem preview curto disponivel."}</div>
              </button>
            ))}
          </div>
        </section>

        <main style={S.workCol}>{renderWorkspace()}</main>
      </div>
    </div>
  );
}

export default function GroupClassificationStudioApp(): JSX.Element {
  return <CockpitProvider><StudioInner /></CockpitProvider>;
}

const S: Record<string, React.CSSProperties> = {
  root: { height: "100vh", boxSizing: "border-box", padding: 18, display: "grid", gridTemplateRows: "auto auto auto minmax(0,1fr)", gap: 12, background: "var(--iccc-bg)", color: "var(--iccc-text)", fontFamily: "var(--iccc-font)", overflow: "hidden" },
  header: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, padding: "14px 16px", borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)" },
  kicker: { fontSize: 10, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  mainTitle: { fontSize: 24, fontWeight: 800, color: "var(--iccc-text)" },
  mainMeta: { fontSize: 13, lineHeight: 1.45, color: "var(--iccc-muted)", maxWidth: 820 },
  primaryBtn: { height: 36, padding: "0 14px", borderRadius: 12, border: "1px solid rgba(37,99,235,0.2)", background: "linear-gradient(180deg,#3b82f6 0%, #2563eb 100%)", color: "#fff", fontSize: 12, fontWeight: 800, display: "inline-flex", alignItems: "center", gap: 8, cursor: "pointer", boxShadow: "0 8px 18px rgba(37,99,235,0.25)" },
  secondaryBtn: { height: 34, padding: "0 12px", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 8, cursor: "pointer" },
  context: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "12px 14px", borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.8)" },
  contextTitle: { fontSize: 15, fontWeight: 700, color: "var(--iccc-text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 780 },
  badges: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  badge: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, background: "rgba(30,64,175,0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  notice: { padding: "10px 12px", borderRadius: 12, border: "1px solid #bfdbfe", background: "#eff6ff", color: "#1d4ed8", fontSize: 12, lineHeight: 1.45 },
  shell: { minHeight: 0, display: "grid", gridTemplateColumns: "220px 320px minmax(0,1fr)", gap: 12 },
  sidebar: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gap: 8, alignContent: "start", overflowY: "auto" },
  menu: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(148,163,184,0.2)", background: "rgba(255,255,255,0.78)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  menuOn: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.9)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  listCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 10, overflow: "hidden" },
  colTitle: { fontSize: 17, fontWeight: 800, color: "var(--iccc-text)" },
  input: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  select: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  listBody: { minHeight: 0, display: "grid", gap: 8, overflowY: "auto", paddingRight: 2 },
  email: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(148,163,184,0.2)", background: "rgba(255,255,255,0.78)", padding: "10px 12px", display: "grid", gap: 6, cursor: "pointer" },
  emailOn: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", padding: "10px 12px", display: "grid", gap: 6, cursor: "pointer" },
  emailTop: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 },
  emailMeta: { fontSize: 11, color: "var(--iccc-muted)" },
  emailSnippet: { fontSize: 12, lineHeight: 1.45, color: "var(--iccc-text-soft, #334155)" },
  counter: { minWidth: 22, height: 22, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(15,23,42,0.06)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 700 },
  workCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, overflow: "hidden" },
  stack: { height: "100%", minHeight: 0, display: "grid", gap: 12, alignContent: "start", overflowY: "auto", paddingRight: 2 },
  card: { borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.74)", padding: 14, display: "grid", gap: 12 },
  titleRow: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  cardTitle: { fontSize: 16, fontWeight: 800, color: "var(--iccc-text)" },
  cardMeta: { fontSize: 12, lineHeight: 1.45, color: "var(--iccc-muted)" },
  metaLine: { display: "flex", gap: 12, flexWrap: "wrap", fontSize: 11, color: "var(--iccc-muted)" },
  chips: { display: "flex", flexWrap: "wrap", gap: 8 },
  groupChip: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, background: "rgba(29,78,216,0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  groupChipBtn: { borderRadius: 999, border: "1px solid rgba(148,163,184,0.24)", background: "rgba(255,255,255,0.92)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  groupChipBtnOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  preview: { width: "100%", minHeight: 520, borderRadius: 14, overflow: "hidden", border: "1px solid rgba(148,163,184,0.24)", background: "#fff" },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 },
  grid2Wide: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 },
  field: { display: "grid", gap: 6 },
  label: { fontSize: 11, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  subTitle: { fontSize: 12, fontWeight: 800, color: "var(--iccc-text)" },
  inline: { display: "flex", alignItems: "center", gap: 8 },
  labelRow: { borderRadius: 14, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", padding: 12, display: "grid", gap: 8 },
  labelHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  linkBtn: { border: "none", background: "transparent", color: "#2563eb", fontSize: 12, fontWeight: 700, cursor: "pointer", padding: 0 },
  check: { display: "inline-flex", alignItems: "center", gap: 8, fontSize: 12, color: "var(--iccc-text)" },
  inlineChecks: { display: "flex", gap: 16, flexWrap: "wrap" },
  attachList: { display: "grid", gap: 10 },
  attachRow: { display: "grid", gridTemplateColumns: "minmax(0,1fr) auto", gap: 12, alignItems: "center", padding: "10px 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)" },
  attachMeta: { display: "grid", gap: 3, minWidth: 0, color: "var(--iccc-text)" },
  attachChecks: { display: "flex", gap: 12, flexWrap: "wrap", justifyContent: "flex-end" },
  summaryRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "10px 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", fontSize: 13, color: "var(--iccc-text)" },
  note: { padding: "12px 14px", borderRadius: 14, border: "1px solid rgba(191,219,254,0.8)", background: "#eff6ff", color: "#1d4ed8", fontSize: 13, lineHeight: 1.5 },
};
