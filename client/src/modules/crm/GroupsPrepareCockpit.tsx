import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  createLinkGroup,
  getRelatedEmailContext,
  listLinkGroups,
  registerRelevantEmail,
  searchGroupTickets,
  searchKnownEmails,
  type GroupTicketEntry,
  type LinkGroupEntry,
  type RelatedEmailEntry,
  type RelevantEmailPayload,
} from "@/api";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX } from "@/modules/crm/group-classification/constants";
import {
  GROUP_CHANGE_CONTRACT,
  buildGroupChangeRequest,
} from "@/modules/crm/groups-v1/contracts";
import {
  buildGroupsPrepareSessionSnapshot,
  DEFAULT_GROUPS_PREPARE_SESSION_STATE,
  GROUPS_PREPARE_CLASSIFY_PARAM,
  GROUPS_PREPARE_SESSION_SAVE_DEBOUNCE_MS,
  buildGroupPreparationSeed,
  getGroupsPrepareSessionSignature,
  readGroupsPrepareSession,
  type GroupsPrepareAttachmentMode,
  type GroupsPrepareGroupMode,
  type GroupsPrepareSessionSaveReason,
  type GroupsPrepareSessionState,
  type GroupsPrepareSubview,
  writeGroupPreparationSeed,
  writeGroupsPrepareSession,
} from "@/modules/crm/groups-v1/prepareSession";
import { openGroupClassificationStudio } from "@/office";
import { HelpHint } from "@/ui/HelpHint";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type PrepareAttachmentRow = {
  key: string;
  emailKey: string;
  emailSubject: string;
  emailDateIso: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  hasContent: boolean;
  documentState?: string;
  source: "anchor" | "known";
};

type EmailKeyCandidate = Partial<RelatedEmailEntry | RelevantEmailPayload> & {
  emailKey?: string;
  messageDateIso?: string;
  receivedAtIso?: string;
};

const STATUS_LABELS: Record<string, string> = {
  em_analise: "Em analise",
  em_progresso: "Em progresso",
  concluido: "Concluido",
  respondido: "Respondido",
  confirmado: "Confirmado",
  arquivado: "Arquivado",
  cancelado: "Cancelado",
};

function normalizeText(value: string | undefined): string {
  return String(value || "").trim().toLowerCase();
}

function getErrorMessage(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message.trim()) return error.message;
  return fallback;
}

function normalizeMessageId(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function makeEmailKey(email: Partial<RelatedEmailEntry | RelevantEmailPayload>): string {
  const candidate = email as EmailKeyCandidate;
  return (
    String(candidate.emailKey || "").trim()
    || String(email?.itemId || "").trim()
    || normalizeMessageId(email?.internetMessageId)
    || [
      String(email?.conversationId || "").trim(),
      String(email?.subject || "").trim().toLowerCase(),
      String(email?.fromEmail || "").trim().toLowerCase(),
      String(candidate.messageDateIso || candidate.receivedAtIso || "").trim(),
    ].join("|")
  );
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function formatBytes(value: number | undefined): string {
  const size = Number(value || 0);
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function estimateBase64Size(base64: string | undefined): number {
  const raw = String(base64 || "").trim().replace(/^data:[^,]+,/, "");
  if (!raw) return 0;
  const padding = raw.endsWith("==") ? 2 : raw.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((raw.length * 3) / 4) - padding);
}

function makeAttachmentSelectionKey(attachment: {
  id?: string;
  key?: string;
  name?: string;
  size?: number;
  contentId?: string;
  contentType?: string;
}): string {
  return (
    String(attachment?.key || attachment?.id || "").trim()
    || [
      String(attachment?.name || "").trim(),
      String(attachment?.size || "").trim(),
      String(attachment?.contentId || "").trim(),
      String(attachment?.contentType || "").trim(),
    ].join("|")
  );
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

function emailMatchesCurrentContext(
  email: Partial<RelatedEmailEntry>,
  ctx: {
    itemId?: string;
    internetMessageId?: string;
    conversationId?: string;
    subject?: string;
    fromEmail?: string;
    receivedDateTimeIso?: string;
  }
): boolean {
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
  const currentFrom = String(ctx.fromEmail || "").trim().toLowerCase();
  const emailFrom = String(email.fromEmail || "").trim().toLowerCase();
  const currentDate = String(ctx.receivedDateTimeIso || "").trim();
  const emailDate = String(email.messageDateIso || email.receivedAtIso || "").trim();

  return Boolean(
    currentConversationId
    && emailConversationId
    && currentConversationId === emailConversationId
    && currentSubject
    && emailSubject
    && currentSubject === emailSubject
    && (!currentFrom || !emailFrom || currentFrom === emailFrom)
    && (!currentDate || !emailDate || currentDate === emailDate)
  );
}

function isRejectedAttachmentState(value: string | undefined): boolean {
  return String(value || "").trim().toLowerCase() === "rejected";
}

function sameStringArray(left: string[], right: string[]): boolean {
  if (left.length !== right.length) return false;
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

function dedupeEmails(rows: Array<RelatedEmailEntry | null | undefined>): RelatedEmailEntry[] {
  const map = new Map<string, RelatedEmailEntry>();
  for (const row of rows) {
    if (!row) continue;
    const key = makeEmailKey(row);
    if (!key) continue;
    if (!map.has(key)) {
      map.set(key, row);
      continue;
    }
    const current = map.get(key)!;
    map.set(key, {
      ...current,
      ...row,
      attachments: Array.isArray(row.attachments) && row.attachments.length ? row.attachments : current.attachments,
      labels: Array.isArray(row.labels) && row.labels.length ? row.labels : current.labels,
      relatedGroups: Array.isArray(row.relatedGroups) && row.relatedGroups.length ? row.relatedGroups : current.relatedGroups,
    });
  }
  return Array.from(map.values()).sort((a, b) =>
    String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || ""))
  );
}

function statusLabel(value: string | undefined): string {
  const token = normalizeText(value);
  if (!token) return "";
  return STATUS_LABELS[token] || String(value || "").trim();
}

function extractPrincipalGroup(email: Partial<RelatedEmailEntry>): { id: string; name: string } | null {
  const principal = (email.relatedGroups || []).find((group) => normalizeText(group.relationKind) !== "referencia");
  if (principal?.id) {
    return { id: String(principal.id).trim(), name: String(principal.name || principal.id).trim() };
  }
  const groupId = String(email.groupId || "").trim();
  return groupId ? { id: groupId, name: String(email.groupName || groupId).trim() } : null;
}

function extractReferenceGroups(email: Partial<RelatedEmailEntry>): Array<{ id: string; name: string }> {
  return (email.relatedGroups || [])
    .filter((group) => normalizeText(group.relationKind) === "referencia")
    .map((group) => ({ id: String(group.id || "").trim(), name: String(group.name || group.id || "").trim() }))
    .filter((group) => group.id);
}

function buildRelevantEmailPayloadFromEmail(email: RelatedEmailEntry): RelevantEmailPayload {
  return {
    itemId: String(email.itemId || "").trim() || undefined,
    internetMessageId: String(email.internetMessageId || "").trim() || undefined,
    conversationId: String(email.conversationId || "").trim() || undefined,
    subject: String(email.subject || "").trim() || undefined,
    fromEmail: String(email.fromEmail || "").trim() || undefined,
    fromName: String(email.fromName || "").trim() || undefined,
    receivedAtIso: String(email.messageDateIso || email.receivedAtIso || "").trim() || undefined,
    messageDateIso: String(email.messageDateIso || email.receivedAtIso || "").trim() || undefined,
    bodyText: String(email.bodyText || "").trim() || undefined,
    bodyHtml: String(email.bodyHtml || "").trim() || undefined,
    attachments: Array.isArray(email.attachments)
      ? email.attachments.map((attachment) => ({
          key: attachment.key,
          id: attachment.id,
          name: attachment.name,
          contentType: attachment.contentType,
          size: attachment.size,
          isInline: attachment.isInline,
          contentId: attachment.contentId,
          content: String(attachment.content || "").trim() || undefined,
          hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
          documentState: attachment.documentState,
        }))
      : [],
  };
}

function CompactToggle({
  label,
  active,
  onClick,
}: {
  label: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      style={S.compactToggleButton}
      onClick={onClick}
      aria-pressed={active}
      aria-label={`${label}: ${active ? "ativo" : "inativo"}`}
    >
      <span style={S.compactToggleLabel}>{label}</span>
      <span style={active ? S.compactToggleTrackOn : S.compactToggleTrackOff}>
        <span style={S.compactToggleThumb} />
      </span>
    </button>
  );
}

export const GroupsPrepareCockpit: React.FC = () => {
  const { ctx, bodyText, bodyHtml, attachments, settings, setMsg, activeGroupSelection, setActiveGroupForCurrentEmail } = useCockpit();
  const [mode] = useState<"prepare" | "explore">("prepare");
  const [subview, setSubview] = useState<GroupsPrepareSubview>("list");
  const [showGroupPanel, setShowGroupPanel] = useState(false);
  const [showFiltersPanel, setShowFiltersPanel] = useState(false);
  const [workingGroupId, setWorkingGroupId] = useState("");
  const [workingGroupQuery, setWorkingGroupQuery] = useState("");
  const [filterQuery, setFilterQuery] = useState("");
  const [attachmentMode, setAttachmentMode] = useState<GroupsPrepareAttachmentMode>("all");
  const [groupMode, setGroupMode] = useState<GroupsPrepareGroupMode>("all");
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [groupsLoading, setGroupsLoading] = useState(false);
  const [groupsError, setGroupsError] = useState("");
  const [contextEmails, setContextEmails] = useState<RelatedEmailEntry[]>([]);
  const [contextTickets, setContextTickets] = useState<GroupTicketEntry[]>([]);
  const [contextLoading, setContextLoading] = useState(false);
  const [knownEmails, setKnownEmails] = useState<RelatedEmailEntry[]>([]);
  const [knownEmailsLoading, setKnownEmailsLoading] = useState(false);
  const [selectedEmailKeys, setSelectedEmailKeys] = useState<string[]>([]);
  const [expandedEmailKeys, setExpandedEmailKeys] = useState<string[]>([]);
  const [selectedAttachmentKeys, setSelectedAttachmentKeys] = useState<string[]>([]);
  const [persistedCurrentEmail, setPersistedCurrentEmail] = useState<RelatedEmailEntry | null>(null);
  const [emailTicketMap, setEmailTicketMap] = useState<Record<string, GroupTicketEntry[]>>({});
  const [busy, setBusy] = useState(false);
  const [sessionReady, setSessionReady] = useState(false);
  const [sessionScopeKey, setSessionScopeKey] = useState("");

  const sessionSnapshot = useMemo<GroupsPrepareSessionState>(() => buildGroupsPrepareSessionSnapshot({
    subview,
    showGroupPanel,
    showFiltersPanel,
    workingGroupId,
    workingGroupQuery,
    filterQuery,
    attachmentMode,
    groupMode,
    selectedEmailKeys,
    expandedEmailKeys,
    selectedAttachmentKeys,
  }), [
    attachmentMode,
    expandedEmailKeys,
    filterQuery,
    groupMode,
    selectedAttachmentKeys,
    selectedEmailKeys,
    showFiltersPanel,
    showGroupPanel,
    subview,
    workingGroupId,
    workingGroupQuery,
  ]);
  const sessionSignature = useMemo(
    () => getGroupsPrepareSessionSignature(sessionSnapshot),
    [sessionSnapshot]
  );
  const renderedSessionRef = useRef<{
    emailKey: string;
    snapshot: GroupsPrepareSessionState;
    signature: string;
  }>({
    emailKey: "",
    snapshot: { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE },
    signature: getGroupsPrepareSessionSignature(DEFAULT_GROUPS_PREPARE_SESSION_STATE),
  });
  const lastPersistedSessionRef = useRef<{ emailKey: string; signature: string }>({
    emailKey: "",
    signature: getGroupsPrepareSessionSignature(DEFAULT_GROUPS_PREPARE_SESSION_STATE),
  });

  const currentEmailBootstrapPayload = useMemo<RelevantEmailPayload>(() => ({
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
      size: attachment.size || estimateBase64Size(attachment.content),
      isInline: attachment.isInline,
      contentId: attachment.contentId,
      content: String(attachment.content || "").trim(),
      hasContent: Boolean(String(attachment.content || "").trim()),
    })),
  }), [attachments, bodyHtml, bodyText, ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject]);

  const currentEmailBootstrapLinkPayload = useMemo(
    () => stripEmailPayloadAttachmentContent(currentEmailBootstrapPayload),
    [currentEmailBootstrapPayload]
  );

  const currentEmailPayload = useMemo<RelevantEmailPayload>(() => {
    if (!persistedCurrentEmail) return currentEmailBootstrapPayload;
    return {
      itemId: String(persistedCurrentEmail.itemId || currentEmailBootstrapPayload.itemId || "").trim(),
      internetMessageId: String(persistedCurrentEmail.internetMessageId || currentEmailBootstrapPayload.internetMessageId || "").trim(),
      conversationId: String(persistedCurrentEmail.conversationId || currentEmailBootstrapPayload.conversationId || "").trim(),
      subject: String(persistedCurrentEmail.subject || currentEmailBootstrapPayload.subject || "").trim(),
      fromEmail: String(persistedCurrentEmail.fromEmail || currentEmailBootstrapPayload.fromEmail || "").trim(),
      fromName: String(persistedCurrentEmail.fromName || currentEmailBootstrapPayload.fromName || "").trim(),
      receivedAtIso: String(persistedCurrentEmail.messageDateIso || persistedCurrentEmail.receivedAtIso || currentEmailBootstrapPayload.receivedAtIso || "").trim(),
      messageDateIso: String(persistedCurrentEmail.messageDateIso || persistedCurrentEmail.receivedAtIso || currentEmailBootstrapPayload.messageDateIso || "").trim(),
      bodyText: String(persistedCurrentEmail.bodyText || currentEmailBootstrapPayload.bodyText || "").trim(),
      bodyHtml: String(persistedCurrentEmail.bodyHtml || currentEmailBootstrapPayload.bodyHtml || "").trim(),
      attachments: Array.isArray(persistedCurrentEmail.attachments)
        ? persistedCurrentEmail.attachments.map((attachment) => ({
            key: attachment.key,
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            size: attachment.size,
            isInline: attachment.isInline,
            contentId: attachment.contentId,
            content: String(attachment.content || "").trim(),
            hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
            documentState: attachment.documentState,
          }))
        : [],
    };
  }, [currentEmailBootstrapPayload, persistedCurrentEmail]);

  const currentEmailLinkPayload = useMemo(
    () => stripEmailPayloadAttachmentContent(currentEmailPayload),
    [currentEmailPayload]
  );

  const currentEmailKey = useMemo(
    () => makeEmailKey(currentEmailLinkPayload),
    [currentEmailLinkPayload]
  );

  const currentEmailEntry = useMemo<RelatedEmailEntry>(() => ({
    emailKey: currentEmailKey,
    itemId: currentEmailLinkPayload.itemId,
    internetMessageId: currentEmailLinkPayload.internetMessageId,
    conversationId: currentEmailLinkPayload.conversationId,
    subject: String(currentEmailPayload.subject || "").trim(),
    fromEmail: String(currentEmailPayload.fromEmail || "").trim() || undefined,
    fromName: String(currentEmailPayload.fromName || "").trim() || undefined,
    receivedAtIso: String(currentEmailPayload.receivedAtIso || currentEmailPayload.messageDateIso || "").trim() || undefined,
    messageDateIso: String(currentEmailPayload.messageDateIso || currentEmailPayload.receivedAtIso || "").trim() || undefined,
    bodyText: String(currentEmailPayload.bodyText || "").trim(),
    bodyHtml: String(currentEmailPayload.bodyHtml || "").trim(),
    status: persistedCurrentEmail?.status,
    labels: persistedCurrentEmail?.labels || [],
    relatedGroups: persistedCurrentEmail?.relatedGroups || [],
    relatedReasons: persistedCurrentEmail?.relatedReasons || [],
    attachments: (currentEmailPayload.attachments || []).map((attachment) => ({
      key: attachment.key,
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size,
      isInline: attachment.isInline,
      contentId: attachment.contentId,
      content: attachment.content,
      hasContent: attachment.hasContent,
      documentState: attachment.documentState,
    })),
  } as RelatedEmailEntry), [currentEmailKey, currentEmailLinkPayload, currentEmailPayload, persistedCurrentEmail]);

  const hasCurrentIdentity = Boolean(
    currentEmailBootstrapLinkPayload.itemId
    || currentEmailBootstrapLinkPayload.internetMessageId
    || currentEmailBootstrapLinkPayload.conversationId
    || currentEmailBootstrapLinkPayload.subject
  );

  const flushSession = useCallback((
    reason: GroupsPrepareSessionSaveReason,
    options?: {
      force?: boolean;
      emailKey?: string | null;
      snapshot?: GroupsPrepareSessionState;
      signature?: string;
    }
  ): boolean => {
    const emailKey = String(options?.emailKey ?? renderedSessionRef.current.emailKey ?? "").trim();
    if (!emailKey) return false;
    const snapshot = buildGroupsPrepareSessionSnapshot(options?.snapshot ?? renderedSessionRef.current.snapshot);
    const signature = String(options?.signature || getGroupsPrepareSessionSignature(snapshot));
    const lastPersisted = lastPersistedSessionRef.current;
    if (!options?.force && lastPersisted.emailKey === emailKey && lastPersisted.signature === signature) {
      return true;
    }
    const saved = writeGroupsPrepareSession(emailKey, snapshot, { reason });
    if (saved) {
      lastPersistedSessionRef.current = { emailKey, signature };
    }
    return saved;
  }, []);

  useEffect(() => {
    setPersistedCurrentEmail(null);
  }, [currentEmailKey]);

  useEffect(() => {
    if (!hasCurrentIdentity) {
      setPersistedCurrentEmail(null);
      return;
    }
    let cancelled = false;

    const resolveCurrentEmail = async () => {
      const loadRelated = async () =>
        getRelatedEmailContext(currentEmailBootstrapLinkPayload).catch(() => null as Awaited<ReturnType<typeof getRelatedEmailContext>> | null);

      const pickBestEmail = (response: Awaited<ReturnType<typeof getRelatedEmailContext>> | null): RelatedEmailEntry | null => {
        const rows = [
          response?.email,
          ...((response?.emails || []).filter(Boolean) as RelatedEmailEntry[]),
        ].filter(Boolean) as RelatedEmailEntry[];
        return rows.find((email) => makeEmailKey(email) === currentEmailKey || emailMatchesCurrentContext(email, ctx)) || rows[0] || null;
      };

      let response = await loadRelated();
      let email = pickBestEmail(response);
      const needsRegistration = !email || (
        Array.isArray(currentEmailBootstrapPayload.attachments)
        && currentEmailBootstrapPayload.attachments.length > 0
        && !(Array.isArray(email?.attachments) && email.attachments.length > 0)
      );

      if (needsRegistration) {
        await registerRelevantEmail({
          ...currentEmailBootstrapPayload,
          attachmentStorageProvider: settings?.groupStorage?.provider || "cloud",
          attachmentStorageBasePath: settings?.groupStorage?.baseFolderPath || "",
        }).catch(() => null);
        response = await loadRelated();
        email = pickBestEmail(response);
      }

      if (!cancelled) {
        setPersistedCurrentEmail(email || null);
      }
    };

    void resolveCurrentEmail();

    return () => {
      cancelled = true;
    };
  }, [
    ctx,
    currentEmailBootstrapLinkPayload,
    currentEmailBootstrapPayload,
    currentEmailKey,
    hasCurrentIdentity,
    settings?.groupStorage?.baseFolderPath,
    settings?.groupStorage?.provider,
  ]);

  useEffect(() => {
    let cancelled = false;
    setGroupsLoading(true);
    setGroupsError("");
    listLinkGroups("/")
      .then((rows) => {
        if (!cancelled) setGroups(Array.isArray(rows) ? rows : []);
      })
      .catch((error: unknown) => {
        if (cancelled) return;
        setGroups([]);
        setGroupsError(getErrorMessage(error, "Nao foi possivel carregar grupos."));
      })
      .finally(() => {
        if (!cancelled) setGroupsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!hasCurrentIdentity) {
      setContextEmails([]);
      setContextTickets([]);
      return;
    }
    let cancelled = false;
    setContextLoading(true);
    getRelatedEmailContext(currentEmailLinkPayload)
      .then((response) => {
        if (cancelled) return;
        setContextEmails(dedupeEmails([
          response?.email || null,
          ...((response?.emails || []).filter(Boolean) as RelatedEmailEntry[]),
        ]));
        setContextTickets(Array.isArray(response?.tickets) ? response.tickets : []);
      })
      .catch((error: unknown) => {
        if (!cancelled) {
          setContextEmails([]);
          setContextTickets([]);
          setMsg(getErrorMessage(error, "Nao foi possivel carregar o contexto do email ancora."));
        }
      })
      .finally(() => {
        if (!cancelled) setContextLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [currentEmailKey, currentEmailLinkPayload, hasCurrentIdentity, setMsg]);

  useEffect(() => {
    const query = String(filterQuery || "").trim();
    if (!showFiltersPanel || !query) {
      setKnownEmails([]);
      setKnownEmailsLoading(false);
      return;
    }
    let cancelled = false;
    const timer = window.setTimeout(() => {
      setKnownEmailsLoading(true);
      searchKnownEmails(query, { limit: 36 })
        .then((rows) => {
          if (!cancelled) setKnownEmails(Array.isArray(rows) ? rows : []);
        })
        .catch((error: unknown) => {
          if (!cancelled) {
            setKnownEmails([]);
            setMsg(getErrorMessage(error, "Nao foi possivel pesquisar emails conhecidos para Preparar."));
          }
        })
        .finally(() => {
          if (!cancelled) setKnownEmailsLoading(false);
        });
    }, 220);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [filterQuery, setMsg, showFiltersPanel]);

  useEffect(() => {
    const previousSession = renderedSessionRef.current;
    if (
      previousSession.emailKey
      && previousSession.emailKey !== currentEmailKey
      && previousSession.emailKey === sessionScopeKey
    ) {
      flushSession("before_context_change", {
        emailKey: previousSession.emailKey,
        snapshot: previousSession.snapshot,
        signature: previousSession.signature,
      });
    }
  }, [currentEmailKey, flushSession, sessionScopeKey]);

  useEffect(() => {
    setSessionReady(false);
    setSessionScopeKey("");
    const sessionKey = String(currentEmailKey || "").trim();
    const sessionState = sessionKey
      ? buildGroupsPrepareSessionSnapshot(readGroupsPrepareSession(sessionKey))
      : { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
    const sessionStateSignature = getGroupsPrepareSessionSignature(sessionState);
    setSubview(sessionState.subview);
    setShowGroupPanel(sessionState.showGroupPanel);
    setShowFiltersPanel(sessionState.showFiltersPanel);
    setWorkingGroupId(sessionState.workingGroupId);
    setWorkingGroupQuery(sessionState.workingGroupQuery);
    setFilterQuery(sessionState.filterQuery);
    setAttachmentMode(sessionState.attachmentMode);
    setGroupMode(sessionState.groupMode);
    setSelectedEmailKeys(sessionState.selectedEmailKeys);
    setExpandedEmailKeys(sessionState.expandedEmailKeys);
    setSelectedAttachmentKeys(sessionState.selectedAttachmentKeys);
    lastPersistedSessionRef.current = { emailKey: sessionKey, signature: sessionStateSignature };
    renderedSessionRef.current = { emailKey: sessionKey, snapshot: sessionState, signature: sessionStateSignature };
    setSessionScopeKey(sessionKey);
    setSessionReady(Boolean(sessionKey));
  }, [currentEmailKey]);

  const preferredWorkingGroupId = useMemo(() => {
    const providerGroupId =
      activeGroupSelection.emailKey === currentEmailKey
        ? String(activeGroupSelection.groupId || "").trim()
        : "";
    if (providerGroupId && groups.some((group) => group.id === providerGroupId)) {
      return providerGroupId;
    }
    const principalGroup = extractPrincipalGroup(currentEmailEntry);
    if (principalGroup && groups.some((group) => group.id === principalGroup.id)) {
      return principalGroup.id;
    }
    return "";
  }, [activeGroupSelection, currentEmailEntry, currentEmailKey, groups]);

  useEffect(() => {
    if (!sessionReady || workingGroupId || !preferredWorkingGroupId) return;
    setWorkingGroupId(preferredWorkingGroupId);
    const preferredGroup = groups.find((group) => group.id === preferredWorkingGroupId) || null;
    if (preferredGroup) setWorkingGroupQuery(preferredGroup.name);
  }, [groups, preferredWorkingGroupId, sessionReady, workingGroupId]);

  useEffect(() => {
    renderedSessionRef.current = {
      emailKey: String(sessionScopeKey || currentEmailKey || "").trim(),
      snapshot: sessionSnapshot,
      signature: sessionSignature,
    };
  }, [currentEmailKey, sessionScopeKey, sessionSignature, sessionSnapshot]);

  useEffect(() => {
    if (!sessionReady || !currentEmailKey || sessionScopeKey !== currentEmailKey) return;
    const lastPersisted = lastPersistedSessionRef.current;
    if (lastPersisted.emailKey === currentEmailKey && lastPersisted.signature === sessionSignature) {
      return;
    }
    const timer = window.setTimeout(() => {
      flushSession("debounced", {
        emailKey: currentEmailKey,
        snapshot: sessionSnapshot,
        signature: sessionSignature,
      });
    }, GROUPS_PREPARE_SESSION_SAVE_DEBOUNCE_MS);
    return () => window.clearTimeout(timer);
  }, [
    currentEmailKey,
    flushSession,
    sessionReady,
    sessionScopeKey,
    sessionSignature,
    sessionSnapshot,
  ]);

  useEffect(() => {
    const handleSaveBeforeExit = () => {
      const currentSession = renderedSessionRef.current;
      if (!currentSession.emailKey) return;
      flushSession("before_exit", {
        emailKey: currentSession.emailKey,
        snapshot: currentSession.snapshot,
        signature: currentSession.signature,
      });
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === "hidden") {
        handleSaveBeforeExit();
      }
    };
    window.addEventListener("pagehide", handleSaveBeforeExit);
    window.addEventListener("beforeunload", handleSaveBeforeExit);
    document.addEventListener("visibilitychange", handleVisibilityChange);
    return () => {
      handleSaveBeforeExit();
      window.removeEventListener("pagehide", handleSaveBeforeExit);
      window.removeEventListener("beforeunload", handleSaveBeforeExit);
      document.removeEventListener("visibilitychange", handleVisibilityChange);
    };
  }, [flushSession]);

  useEffect(() => {
    if (!sessionReady) return;
    setActiveGroupForCurrentEmail(workingGroupId || null);
  }, [sessionReady, setActiveGroupForCurrentEmail, workingGroupId]);

  const mergedCandidateEmails = useMemo(
    () => dedupeEmails([currentEmailEntry, ...contextEmails, ...knownEmails]),
    [contextEmails, currentEmailEntry, knownEmails]
  );

  const visibleEmails = useMemo(() => {
    const query = normalizeText(filterQuery);
    return mergedCandidateEmails.filter((email) => {
      if (showFiltersPanel && query) {
        const haystack = [
          email.subject,
          email.fromName,
          email.fromEmail,
          ...(email.labels || []),
          ...(email.relatedGroups || []).map((group) => group.name || group.id),
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        if (!haystack.includes(query)) return false;
      }

      const attachmentCount = Array.isArray(email.attachments) ? email.attachments.length : 0;
      if (attachmentMode === "with" && attachmentCount === 0) return false;
      if (attachmentMode === "without" && attachmentCount > 0) return false;

      const principalGroup = extractPrincipalGroup(email);
      if (groupMode === "with_group" && !principalGroup) return false;
      if (groupMode === "without_group" && principalGroup) return false;

      return true;
    });
  }, [attachmentMode, filterQuery, groupMode, mergedCandidateEmails, showFiltersPanel]);

  useEffect(() => {
    if (!sessionReady) return;
    const validKeys = new Set(mergedCandidateEmails.map((email) => makeEmailKey(email)).filter(Boolean));
    const defaultSelection = currentEmailKey && validKeys.has(currentEmailKey)
      ? [currentEmailKey]
      : mergedCandidateEmails[0]
        ? [makeEmailKey(mergedCandidateEmails[0])]
        : [];
    setSelectedEmailKeys((current) => {
      const next = current.filter((key) => validKeys.has(key));
      if (next.length) return sameStringArray(next, current) ? current : next;
      return sameStringArray(defaultSelection, current) ? current : defaultSelection;
    });
    setExpandedEmailKeys((current) => {
      const next = current.filter((key) => validKeys.has(key));
      const fallback = mergedCandidateEmails.slice(0, 2).map((email) => makeEmailKey(email)).filter(Boolean);
      const resolved = next.length ? next : fallback;
      return sameStringArray(resolved, current) ? current : resolved;
    });
  }, [currentEmailKey, mergedCandidateEmails, sessionReady]);

  const selectedEmails = useMemo(
    () => mergedCandidateEmails.filter((email) => selectedEmailKeys.includes(makeEmailKey(email))),
    [mergedCandidateEmails, selectedEmailKeys]
  );

  const attachmentRows = useMemo<PrepareAttachmentRow[]>(() => {
    const rows: PrepareAttachmentRow[] = [];
    const ignoreInline = settings?.groupStorage?.ignoreInlineAttachments === true;

    for (const email of selectedEmails) {
      const emailKey = makeEmailKey(email);
      const sourceAttachments = emailKey === currentEmailKey
        ? currentEmailPayload.attachments || []
        : email.attachments || [];
      for (const attachment of sourceAttachments) {
        const attachmentKey = makeAttachmentSelectionKey(attachment);
        if (!attachmentKey || !String(attachment.name || "").trim()) continue;
        if (isRejectedAttachmentState(attachment.documentState)) continue;
        if (ignoreInline && attachment.isInline) continue;
        rows.push({
          key: `${emailKey}:${attachmentKey}`,
          emailKey,
          emailSubject: String(email.subject || "(sem assunto)").trim(),
          emailDateIso: String(email.messageDateIso || email.receivedAtIso || "").trim(),
          name: String(attachment.name || "").trim(),
          contentType: attachment.contentType,
          size: Number(attachment.size || 0) || estimateBase64Size(attachment.content),
          isInline: attachment.isInline === true,
          hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
          documentState: attachment.documentState,
          source: emailKey === currentEmailKey ? "anchor" : "known",
        });
      }
    }

    return rows.sort((a, b) => {
      if (a.source !== b.source) return a.source === "anchor" ? -1 : 1;
      if (a.emailDateIso !== b.emailDateIso) return b.emailDateIso.localeCompare(a.emailDateIso);
      return a.name.localeCompare(b.name, "pt-PT");
    });
  }, [currentEmailKey, currentEmailPayload.attachments, selectedEmails, settings?.groupStorage?.ignoreInlineAttachments]);

  useEffect(() => {
    if (!sessionReady) return;
    const validKeys = new Set(attachmentRows.map((attachment) => attachment.key));
    const defaultSelection = attachmentRows.map((attachment) => attachment.key);
    setSelectedAttachmentKeys((current) => {
      const next = current.filter((key) => validKeys.has(key));
      const resolved = next.length ? next : defaultSelection;
      return sameStringArray(resolved, current) ? current : resolved;
    });
  }, [attachmentRows, sessionReady]);

  const selectedAttachments = useMemo(
    () => attachmentRows.filter((attachment) => selectedAttachmentKeys.includes(attachment.key)),
    [attachmentRows, selectedAttachmentKeys]
  );

  const selectedAttachmentCountByEmail = useMemo(() => {
    const map = new Map<string, number>();
    for (const attachment of selectedAttachments) {
      map.set(attachment.emailKey, (map.get(attachment.emailKey) || 0) + 1);
    }
    return map;
  }, [selectedAttachments]);

  const workingGroup = useMemo(
    () => groups.find((group) => group.id === workingGroupId) || null,
    [groups, workingGroupId]
  );

  const exactWorkingGroupMatch = useMemo(
    () => groups.find((group) => normalizeText(group.name) === normalizeText(workingGroupQuery)) || null,
    [groups, workingGroupQuery]
  );

  const workingGroupCandidates = useMemo(() => {
    const query = normalizeText(workingGroupQuery);
    return [...groups]
      .filter((group) => !group.isArchived)
      .filter((group) => {
        if (!query) return true;
        return [group.name, group.description, ...(group.labels || [])]
          .filter(Boolean)
          .join(" ")
          .toLowerCase()
          .includes(query);
      })
      .sort((a, b) => {
        const aUpdated = String(a.updatedAt || a.createdAt || "");
        const bUpdated = String(b.updatedAt || b.createdAt || "");
        if (aUpdated !== bUpdated) return bUpdated.localeCompare(aUpdated);
        return String(a.name || "").localeCompare(String(b.name || ""), "pt-PT");
      })
      .slice(0, 6);
  }, [groups, workingGroupQuery]);

  const emailGroupChangeRequests = useMemo(
    () => selectedEmails
      .map((email) =>
        buildGroupChangeRequest({
          emailKey: makeEmailKey(email),
          previousPrincipalGroupId: extractPrincipalGroup(email)?.id || null,
          nextPrincipalGroupId: workingGroupId || null,
          keepPreviousGroupAsReference: false,
        })
      )
      .filter(Boolean),
    [selectedEmails, workingGroupId]
  );

  const activeFilterSummary = useMemo(() => {
    const summary: string[] = [];
    if (showFiltersPanel && String(filterQuery || "").trim()) summary.push(`Pesquisa: ${String(filterQuery || "").trim()}`);
    if (attachmentMode === "with") summary.push("So com anexos");
    if (attachmentMode === "without") summary.push("So sem anexos");
    if (groupMode === "with_group") summary.push("So com grupo");
    if (groupMode === "without_group") summary.push("So sem grupo");
    return summary;
  }, [attachmentMode, filterQuery, groupMode, showFiltersPanel]);

  const expandedEmailRows = useMemo(
    () => visibleEmails.filter((email) => expandedEmailKeys.includes(makeEmailKey(email))),
    [expandedEmailKeys, visibleEmails]
  );

  useEffect(() => {
    const targets = expandedEmailRows.filter((email) => {
      const key = makeEmailKey(email);
      return key && emailTicketMap[key] === undefined;
    });
    if (!targets.length) return;

    let cancelled = false;
    void Promise.all(
      targets.map(async (email) => {
        const key = makeEmailKey(email);
        try {
          const rows = await searchGroupTickets({
            email: buildRelevantEmailPayloadFromEmail(email),
            limit: 4,
          });
          return { key, tickets: Array.isArray(rows) ? rows : [] };
        } catch {
          return { key, tickets: [] as GroupTicketEntry[] };
        }
      })
    ).then((results) => {
      if (cancelled) return;
      setEmailTicketMap((current) => {
        const next = { ...current };
        for (const result of results) next[result.key] = result.tickets;
        return next;
      });
    });

    return () => {
      cancelled = true;
    };
  }, [emailTicketMap, expandedEmailRows]);

  async function handleCreateWorkingGroup() {
    const name = String(workingGroupQuery || "").trim();
    if (!name) {
      setMsg("Escreve um nome de grupo para o preparar.");
      return;
    }
    if (exactWorkingGroupMatch) {
      setWorkingGroupId(exactWorkingGroupMatch.id);
      setWorkingGroupQuery(exactWorkingGroupMatch.name);
      return;
    }
    setBusy(true);
    try {
      const created = await createLinkGroup({ name, documentsEnabled: true });
      setGroups((current) => dedupeGroupEntries([created, ...current]));
      setWorkingGroupId(created.id);
      setWorkingGroupQuery(created.name);
      setShowGroupPanel(true);
      setMsg(`Grupo preparado: ${created.name}`);
    } catch (error: unknown) {
      setMsg(getErrorMessage(error, "Nao foi possivel criar o grupo em trabalho."));
    } finally {
      setBusy(false);
    }
  }

  async function handleOpenClassificationFromPrepare() {
    flushSession("before_classify", {
      emailKey: currentEmailKey,
      snapshot: sessionSnapshot,
      signature: sessionSignature,
    });
    const params: Record<string, string> = {};
    if (currentEmailLinkPayload.itemId) params.itemId = currentEmailLinkPayload.itemId;
    if (currentEmailLinkPayload.internetMessageId) params.internetMessageId = currentEmailLinkPayload.internetMessageId;
    if (currentEmailLinkPayload.conversationId) params.conversationId = currentEmailLinkPayload.conversationId;
    if (currentEmailLinkPayload.subject) params.subject = String(currentEmailLinkPayload.subject);
    if (currentEmailLinkPayload.fromEmail) params.fromEmail = String(currentEmailLinkPayload.fromEmail);
    if (currentEmailLinkPayload.fromName) params.fromName = String(currentEmailLinkPayload.fromName);
    if (currentEmailLinkPayload.receivedAtIso) params.receivedAtIso = String(currentEmailLinkPayload.receivedAtIso);

    try {
      const seedKey = `${GROUP_CLASSIFICATION_SEED_STORAGE_PREFIX}${Date.now()}:${currentEmailKey || "email"}`;
      localStorage.setItem(seedKey, JSON.stringify({
        ...currentEmailLinkPayload,
        bodyText: String(currentEmailPayload.bodyText || "").trim(),
        bodyHtml: String(currentEmailPayload.bodyHtml || "").trim(),
        attachments: (currentEmailPayload.attachments || []).map((attachment) => ({
          key: attachment.key,
          id: attachment.id,
          name: attachment.name,
          contentType: attachment.contentType,
          size: attachment.size,
          isInline: attachment.isInline,
          contentId: attachment.contentId,
          content: String(attachment.content || "").trim(),
          documentState: attachment.documentState,
          hasContent: attachment.hasContent === true || Boolean(String(attachment.content || "").trim()),
        })),
      }));
      params.seedKey = seedKey;
    } catch {
      // keep opening even if local seed write fails
    }

    const prepareSeedKey = writeGroupPreparationSeed(
      buildGroupPreparationSeed({
        anchorEmailKey: currentEmailKey,
        selectedEmailKeys,
        selectedAttachmentKeys,
        workingGroupId,
        filterQuery,
        attachmentMode,
        groupMode,
      })
    );
    if (prepareSeedKey) {
      params[GROUPS_PREPARE_CLASSIFY_PARAM] = prepareSeedKey;
    }

    try {
      await openGroupClassificationStudio(params);
    } catch (error: unknown) {
      setMsg(getErrorMessage(error, "Nao foi possivel abrir o Classificar."));
    }
  }

  function handleManualSessionSave() {
    if (!currentEmailKey) {
      setMsg("Sem email ancora para guardar sessao.");
      return;
    }
    const saved = flushSession("manual", {
      force: true,
      emailKey: currentEmailKey,
      snapshot: sessionSnapshot,
      signature: sessionSignature,
    });
    setMsg(saved ? "Sessao de Preparar guardada localmente." : "Nao foi possivel guardar a sessao local.");
  }

  function handleSubviewChange(nextSubview: GroupsPrepareSubview) {
    if (nextSubview === subview) return;
    flushSession("before_subview_change", {
      emailKey: currentEmailKey,
      snapshot: sessionSnapshot,
      signature: sessionSignature,
    });
    setSubview(nextSubview);
  }

  function toggleEmailSelection(emailKey: string) {
    setSelectedEmailKeys((current) =>
      current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]
    );
  }

  function toggleEmailExpanded(emailKey: string) {
    setExpandedEmailKeys((current) =>
      current.includes(emailKey) ? current.filter((entry) => entry !== emailKey) : [...current, emailKey]
    );
  }

  function toggleAttachmentSelection(attachmentKey: string) {
    setSelectedAttachmentKeys((current) =>
      current.includes(attachmentKey) ? current.filter((entry) => entry !== attachmentKey) : [...current, attachmentKey]
    );
  }

  const summaryMetricCards = [
    { label: "Emails", value: String(selectedEmails.length), meta: selectedEmails.length === 1 ? "1 selecionado" : `${selectedEmails.length} selecionados` },
    { label: "Anexos", value: String(selectedAttachments.length), meta: selectedAttachments.filter((attachment) => attachment.hasContent).length ? `${selectedAttachments.filter((attachment) => attachment.hasContent).length} com conteudo local` : "Sem upload remoto nesta fase" },
    { label: "Filtros", value: String(activeFilterSummary.length), meta: activeFilterSummary.length ? activeFilterSummary.join(" / ") : "Sem filtros ativos" },
  ];

  if (!hasCurrentIdentity) {
    return (
      <div style={S.root}>
        <div style={S.header}>
          <div>
            <div style={S.kicker}>Grupos v1</div>
            <div style={S.title}>Preparar</div>
          </div>
        </div>
        <PanelState
          tone="info"
          title="Sem email ancora disponivel"
          description="Abre um email no Outlook para preparar o conjunto de trabalho desta aba."
        />
      </div>
    );
  }

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div style={S.headerMain}>
          <div style={S.kicker}>Gestor de grupos</div>
          <div style={S.title}>Grupos</div>
          <div style={S.headerHint}>Preparar monta o conjunto de trabalho. Classificar continua a fechar.</div>
        </div>
        <HelpHint
          title="Ajuda: Preparar"
          text="Esta ronda implementa apenas a entrada de Preparar. Explorar no add-in, Explorador de Grupos e Gestor do Grupo ficam separados."
        />
      </div>

      <div style={S.segmentBar}>
        <button type="button" style={mode === "prepare" ? S.segmentActive : S.segment} disabled>
          Preparar
        </button>
        <button type="button" style={S.segmentDisabled} disabled title="Explorar entra numa fase seguinte.">
          Explorar
        </button>
      </div>

      <div style={S.anchorCard}>
        <div style={S.anchorCopy}>
          <div style={S.sectionTitleRow}>
            <div style={S.fieldLabel}>Email ancora</div>
            <HelpHint
              title="Ajuda: Email ancora"
              text="O email aberto no Outlook fixa o contexto de Preparar. Esta sub-aba nao reusa essa ancora em Explorar."
            />
          </div>
          <div style={S.anchorSubject}>{currentEmailEntry.subject || "(sem assunto)"}</div>
          <div style={S.anchorInfoChips}>
            <span style={S.mutedBadge}>{currentEmailEntry.fromName || currentEmailEntry.fromEmail || "Sem remetente"}</span>
            {formatDate(currentEmailEntry.messageDateIso || currentEmailEntry.receivedAtIso)
              ? <span style={S.mutedBadge}>{formatDate(currentEmailEntry.messageDateIso || currentEmailEntry.receivedAtIso)}</span>
              : null}
            {statusLabel(currentEmailEntry.status) ? <span style={S.statusBadge}>{statusLabel(currentEmailEntry.status)}</span> : null}
            {extractPrincipalGroup(currentEmailEntry) ? <span style={S.labelBadge}>Grupo: {extractPrincipalGroup(currentEmailEntry)?.name}</span> : null}
            <span style={S.anchorBadge}>{Array.isArray(currentEmailEntry.attachments) ? currentEmailEntry.attachments.length : 0} anexo(s)</span>
          </div>
        </div>
        <div style={S.anchorActions}>
          <CompactToggle label="Grupo" active={showGroupPanel} onClick={() => setShowGroupPanel((value) => !value)} />
          <CompactToggle label="Filtros" active={showFiltersPanel} onClick={() => setShowFiltersPanel((value) => !value)} />
        </div>
      </div>

      {showGroupPanel ? (
        <div style={S.panelCard}>
          <div style={S.sectionTitleRow}>
            <div style={S.fieldLabel}>Grupo em trabalho</div>
            <HelpHint
              title="Ajuda: Grupo em trabalho"
              text="Aqui escolhes ou crias o grupo que esta a preparar o trabalho. Ainda nao ha edicao rica nem alteracao final dos emails."
            />
          </div>
          <div style={S.compactRow}>
            <input style={{ ...S.input, flex: "1 1 220px" }} value={workingGroupQuery} onChange={(event) => setWorkingGroupQuery(event.target.value)} placeholder="Pesquisar grupo existente" />
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => {
                if (exactWorkingGroupMatch) {
                  setWorkingGroupId(exactWorkingGroupMatch.id);
                  setWorkingGroupQuery(exactWorkingGroupMatch.name);
                }
              }}
              disabled={!exactWorkingGroupMatch}
            >
              <Icons.Search size={12} />
              Pesquisar
            </button>
            <button type="button" style={S.primaryBtn} onClick={() => void handleCreateWorkingGroup()} disabled={busy || !String(workingGroupQuery || "").trim() || Boolean(exactWorkingGroupMatch)}>
              <Icons.Plus size={12} />
              Criar
            </button>
          </div>
          {workingGroup ? (
            <div style={S.selectedGroupCard}>
              <div style={S.selectedGroupMain}>
                <div style={S.selectedGroupTitle}>{workingGroup.name}</div>
                <div style={S.smallMeta}>{workingGroup.memberCount || 0} email(s) / {workingGroup.documentsEnabled === false ? "documentos off" : "documentos on"}</div>
              </div>
              <button type="button" style={S.iconGhostBtn} onClick={() => setWorkingGroupId("")} title="Limpar grupo em trabalho">
                <Icons.RefreshCw size={12} />
              </button>
            </div>
          ) : null}
          {groupsError ? <PanelState compact tone="error" title="Falha a carregar grupos" description={groupsError} /> : null}
          {groupsLoading ? <PanelState compact tone="loading" title="A carregar grupos" description="A preparar os grupos existentes para selecao." /> : null}
          {!groupsLoading && workingGroupCandidates.length ? (
            <div style={S.listWrap}>
              {workingGroupCandidates.map((group) => (
                <button key={group.id} type="button" style={group.id === workingGroupId ? S.listRowActive : S.listRow} onClick={() => { setWorkingGroupId(group.id); setWorkingGroupQuery(group.name); }}>
                  <span>{group.name}</span>
                  <span style={S.countBadge}>{group.memberCount || 0}</span>
                </button>
              ))}
            </div>
          ) : null}
          {emailGroupChangeRequests.length ? (
            <div style={S.warningBox}>
              {emailGroupChangeRequests.length} email(s) ja tem grupo principal diferente. A mudanca continua explicita e sera confirmada no Classificar; o grupo antigo pode seguir como referencia so nesse email.
            </div>
          ) : (
            <div style={S.smallMeta}>{workingGroup ? "Grupo em trabalho definido sem aplicar classificacao final." : "Ainda sem grupo em trabalho."}</div>
          )}
        </div>
      ) : null}

      {showFiltersPanel ? (
        <div style={S.panelCard}>
          <div style={S.sectionTitleRow}>
            <div style={S.fieldLabel}>Filtros de pesquisa</div>
            <HelpHint
              title="Ajuda: Filtros"
              text="Os filtros so ajudam a trazer emails para o conjunto de trabalho. Nao fecham classificacao final nem substituem o Classificar."
            />
          </div>
          <div style={S.filterGrid}>
            <input style={S.input} value={filterQuery} onChange={(event) => setFilterQuery(event.target.value)} placeholder="Assunto, remetente, grupo ou etiqueta" />
            <select style={S.select} value={attachmentMode} onChange={(event) => setAttachmentMode(event.target.value as GroupsPrepareAttachmentMode)}>
              <option value="all">Todos os anexos</option>
              <option value="with">So com anexos</option>
              <option value="without">So sem anexos</option>
            </select>
            <select style={S.select} value={groupMode} onChange={(event) => setGroupMode(event.target.value as GroupsPrepareGroupMode)}>
              <option value="all">Todos os grupos</option>
              <option value="with_group">So com grupo</option>
              <option value="without_group">So sem grupo</option>
            </select>
          </div>
          <div style={S.smallMeta}>
            {knownEmailsLoading
              ? "A pesquisar emails registados para alargar o conjunto."
              : activeFilterSummary.length
                ? activeFilterSummary.join(" / ")
                : "Sem filtros ativos."}
          </div>
        </div>
      ) : null}

      <div style={S.segmentBar}>
        <button type="button" style={subview === "list" ? S.segmentActive : S.segment} onClick={() => handleSubviewChange("list")}>Lista</button>
        <button type="button" style={subview === "attachments" ? S.segmentActive : S.segment} onClick={() => handleSubviewChange("attachments")}>Anexos</button>
        <button type="button" style={subview === "summary" ? S.segmentActive : S.segment} onClick={() => handleSubviewChange("summary")}>Resumo</button>
      </div>

      {subview === "list" ? (
        <div style={S.viewStack}>
          <div style={S.inlineMetaRow}>
            <span style={S.smallMeta}>{selectedEmails.length} email(s) no conjunto de trabalho</span>
            <button type="button" style={S.tinyBtn} onClick={() => setSelectedEmailKeys(visibleEmails.map((email) => makeEmailKey(email)))}>Todos os visiveis</button>
            <button type="button" style={S.tinyBtn} onClick={() => setSelectedEmailKeys(currentEmailKey ? [currentEmailKey] : [])}>So ancora</button>
          </div>
          {contextLoading ? <PanelState compact tone="loading" title="A carregar emails" description="A montar o conjunto base a partir do email ancora." /> : null}
          {!contextLoading && !visibleEmails.length ? (
            <PanelState compact tone="info" title="Sem emails visiveis" description="Liga o painel de filtros ou alarga a pesquisa para trazer emails para o conjunto." />
          ) : (
            <div style={S.emailList}>
              {visibleEmails.map((email) => {
                const emailKey = makeEmailKey(email);
                const expanded = expandedEmailKeys.includes(emailKey);
                const selected = selectedEmailKeys.includes(emailKey);
                const principalGroup = extractPrincipalGroup(email);
                const referenceGroups = extractReferenceGroups(email);
                const tickets = emailKey === currentEmailKey ? contextTickets : (emailTicketMap[emailKey] || []);
                const attachmentCount = Array.isArray(email.attachments) ? email.attachments.length : 0;
                const pendingChange = workingGroupId ? buildGroupChangeRequest({
                  emailKey,
                  previousPrincipalGroupId: principalGroup?.id || null,
                  nextPrincipalGroupId: workingGroupId || null,
                  keepPreviousGroupAsReference: false,
                }) : null;

                return (
                  <div key={emailKey} style={expanded ? S.emailCardExpanded : S.emailCard}>
                    <div style={S.emailCardHead}>
                      <label style={S.checkboxCell}>
                        <input type="checkbox" checked={selected} onChange={() => toggleEmailSelection(emailKey)} />
                      </label>
                      <button type="button" style={S.emailCardMain} onClick={() => toggleEmailExpanded(emailKey)}>
                        <div style={S.emailCardCopy}>
                          <div style={S.emailSubject}>
                            {email.subject || "(sem assunto)"}
                            {emailKey === currentEmailKey ? <span style={S.anchorBadge}>Ancora</span> : null}
                          </div>
                          <div style={S.emailMeta}>
                            {email.fromName || email.fromEmail || "Sem remetente"}
                            {formatDate(email.messageDateIso || email.receivedAtIso) ? ` / ${formatDate(email.messageDateIso || email.receivedAtIso)}` : ""}
                            {statusLabel(email.status) ? ` / ${statusLabel(email.status)}` : ""}
                          </div>
                        </div>
                        <div style={S.emailHeadBadges}>
                          {attachmentCount ? <span style={S.countBadge}>{attachmentCount}</span> : null}
                          {expanded ? <Icons.ArrowUp size={12} /> : <Icons.ArrowDown size={12} />}
                        </div>
                      </button>
                    </div>

                    {expanded ? (
                      <>
                        {(email.labels || []).length ? (
                          <div style={S.detailBadgeStack}>
                            {(email.labels || []).slice(0, 6).map((label) => (
                              <span key={`${emailKey}:${label}`} style={S.labelBadge}>{label}</span>
                            ))}
                          </div>
                        ) : null}
                        <div style={S.detailBadgeStack}>
                          <span style={principalGroup ? S.primaryBadge : S.mutedBadge}>
                            Grupo: {principalGroup?.name || "Sem grupo principal"}
                          </span>
                          <span style={referenceGroups.length ? S.labelBadge : S.mutedBadge}>
                            Ref: {referenceGroups.map((group) => group.name).join(", ") || "Sem referencias"}
                          </span>
                          <span style={tickets[0]?.code ? S.readyBadge : S.mutedBadge}>
                            Ticket: {tickets[0]?.code || (emailKey === currentEmailKey ? "Sem ticket no contexto atual" : "Sem ticket conhecido")}
                          </span>
                          <span style={S.mutedBadge}>
                            Anexos: {attachmentCount} / {selectedAttachmentCountByEmail.get(emailKey) || 0}
                          </span>
                        </div>
                        {pendingChange ? (
                          <div style={S.warningSubtle}>
                            Se fores para "{workingGroup?.name || workingGroupId}", o Classificar tem de confirmar esta mudanca explicitamente e pode manter o grupo antigo como referencia no scope {GROUP_CHANGE_CONTRACT.scope}.
                          </div>
                        ) : null}
                      </>
                    ) : null}
                  </div>
                );
              })}
            </div>
          )}
        </div>
      ) : null}

      {subview === "attachments" ? (
        <div style={S.viewStack}>
          <div style={S.panelCard}>
            <div style={S.sectionTitleRow}>
              <div style={S.fieldLabel}>Gestor de anexos</div>
              <HelpHint title="Ajuda: Anexos" text="Aqui so escolhes e marcas anexos para o conjunto preparado. Nao existe upload remoto imediato nesta fase." />
            </div>
            <div style={S.inlineMetaRow}>
              <span style={S.smallMeta}>{selectedAttachments.length}/{attachmentRows.length} anexo(s) preparado(s)</span>
              <button type="button" style={S.tinyBtn} onClick={() => setSelectedAttachmentKeys(attachmentRows.map((attachment) => attachment.key))}>Todos</button>
              <button type="button" style={S.tinyBtn} onClick={() => setSelectedAttachmentKeys([])}>Nenhum</button>
            </div>
            {!selectedEmails.length ? (
              <PanelState compact tone="info" title="Sem emails selecionados" description="Escolhe primeiro emails na Lista para preparar anexos." />
            ) : !attachmentRows.length ? (
              <PanelState compact tone="info" title="Sem anexos disponiveis" description="Os emails selecionados nao expõem anexos utilizaveis para esta fase." />
            ) : (
              <div style={S.attachmentList}>
                {attachmentRows.map((attachment) => {
                  const selected = selectedAttachmentKeys.includes(attachment.key);
                  return (
                    <label key={attachment.key} style={selected ? S.attachmentRowActive : S.attachmentRow}>
                      <input type="checkbox" checked={selected} onChange={() => toggleAttachmentSelection(attachment.key)} />
                      <div style={S.attachmentCopy}>
                        <div style={S.attachmentName}>{attachment.name}</div>
                        <div style={S.emailMeta}>
                          {attachment.emailSubject}
                          {attachment.emailDateIso ? ` / ${formatDate(attachment.emailDateIso)}` : ""}
                          {formatBytes(attachment.size) ? ` / ${formatBytes(attachment.size)}` : ""}
                        </div>
                      </div>
                      <div style={S.badgeWrap}>
                        <span style={attachment.hasContent ? S.readyBadge : S.mutedBadge}>{attachment.hasContent ? "Local" : "Metadados"}</span>
                        {attachment.isInline ? <span style={S.mutedBadge}>Inline</span> : null}
                        {selected ? <span style={S.selectedBadge}>Marcado</span> : null}
                      </div>
                    </label>
                  );
                })}
              </div>
            )}
          </div>
        </div>
      ) : null}

      {subview === "summary" ? (
        <div style={S.viewStack}>
          <div style={S.metricGrid}>
            {summaryMetricCards.map((card) => (
              <div key={card.label} style={S.metricCard}>
                <div style={S.metricLabel}>{card.label}</div>
                <div style={S.metricValue}>{card.value}</div>
                <div style={S.metricMeta}>{card.meta}</div>
              </div>
            ))}
          </div>

          <div style={S.panelCard}>
            <div style={S.sectionTitleRow}>
              <div style={S.fieldLabel}>Resumo antes de abrir no Classificar</div>
              <HelpHint title="Ajuda: Resumo" text="Este resumo confirma selecao, grupo em trabalho, anexos preparados e filtros ativos. O fecho final continua no Classificar." />
            </div>
            <div style={S.summaryList}>
              <div><b>Email ancora:</b> {currentEmailEntry.subject || "(sem assunto)"}</div>
              <div><b>Grupo em trabalho:</b> {workingGroup?.name || "Sem grupo em trabalho"}</div>
              <div><b>Anexos preparados:</b> {selectedAttachments.length || 0}</div>
              <div><b>Filtros ativos:</b> {activeFilterSummary.join(", ") || "Sem filtros ativos"}</div>
              <div><b>Mudancas a confirmar:</b> {emailGroupChangeRequests.length ? `${emailGroupChangeRequests.length} email(s) com mudanca explicita pendente` : "Sem mudancas de grupo pendentes"}</div>
            </div>
            <div style={S.badgeWrap}>
              {workingGroup ? <span style={S.primaryBadge}>{workingGroup.name}</span> : null}
              {selectedAttachments.length ? <span style={S.selectedBadge}>Anexos preparados</span> : null}
              {emailGroupChangeRequests.length ? <span style={S.warningBadge}>Mudanca explicita</span> : null}
              {activeFilterSummary.length ? <span style={S.labelBadge}>Filtros ativos</span> : null}
            </div>
          </div>
        </div>
      ) : null}

      <div style={S.footerBar}>
        <div style={S.footerMain}>
          <div style={S.footerStats}>
            <span style={S.footerStat}>{selectedEmails.length} email(s)</span>
            <span style={S.footerStat}>{selectedAttachments.length} anexo(s)</span>
          </div>
          <div style={S.footerCopy}>
            Sessao local separada da persistencia remota.
          </div>
        </div>
        <div style={S.inlineActions}>
          <button type="button" style={S.secondaryBtn} onClick={handleManualSessionSave}>
            <Icons.Save size={12} />
            Guardar
          </button>
          <button type="button" style={S.primaryBtn} onClick={() => void handleOpenClassificationFromPrepare()}>
            <Icons.Target size={12} />
            Classificar
          </button>
        </div>
      </div>
    </div>
  );
};

function dedupeGroupEntries(rows: LinkGroupEntry[]): LinkGroupEntry[] {
  const map = new Map<string, LinkGroupEntry>();
  for (const row of rows) {
    if (!row?.id) continue;
    map.set(row.id, row);
  }
  return Array.from(map.values()).sort((a, b) => {
    const aUpdated = String(a.updatedAt || a.createdAt || "");
    const bUpdated = String(b.updatedAt || b.createdAt || "");
    if (aUpdated !== bUpdated) return bUpdated.localeCompare(aUpdated);
    return String(a.name || "").localeCompare(String(b.name || ""), "pt-PT");
  });
}

const baseButton: React.CSSProperties = {
  borderRadius: 999,
  border: "1px solid var(--iccc-card-border)",
  padding: "6px 10px",
  fontSize: 10,
  fontWeight: 700,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 5,
  lineHeight: 1.1,
};

const S: Record<string, React.CSSProperties> = {
  root: { display: "grid", gap: 6, alignContent: "start" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8, padding: 8, borderRadius: 18, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  headerMain: { display: "grid", gap: 1, minWidth: 0 },
  kicker: { fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--iccc-text-muted)" },
  title: { fontSize: 14, fontWeight: 800, color: "var(--iccc-text)" },
  headerHint: { marginTop: 1, fontSize: 10, lineHeight: 1.35, color: "var(--iccc-text-muted)" },
  segmentBar: { display: "flex", gap: 3, padding: 3, borderRadius: 999, border: "1px solid var(--iccc-card-border)", background: "rgba(241,245,249,0.92)", width: "100%", boxSizing: "border-box" },
  segment: { flex: "1 1 0", border: "none", background: "transparent", color: "var(--iccc-text-muted)", padding: "6px 8px", borderRadius: 999, fontSize: 10, fontWeight: 600, cursor: "pointer" },
  segmentActive: { flex: "1 1 0", border: "1px solid rgba(148,163,184,0.18)", background: "#fff", color: "var(--iccc-text)", padding: "6px 8px", borderRadius: 999, fontSize: 10, fontWeight: 700, cursor: "pointer", boxShadow: "0 1px 2px rgba(15,23,42,0.06)" },
  segmentDisabled: { flex: "1 1 0", border: "none", background: "transparent", color: "rgba(100,116,139,0.68)", padding: "6px 8px", borderRadius: 999, fontSize: 10, fontWeight: 600, cursor: "not-allowed" },
  anchorCard: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 8, alignItems: "start", padding: 9, borderRadius: 18, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.88)" },
  anchorCopy: { display: "grid", gap: 4, minWidth: 0 },
  anchorSubject: { fontSize: 12.5, fontWeight: 800, color: "var(--iccc-text)", wordBreak: "break-word", lineHeight: 1.25 },
  anchorMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  anchorActions: { display: "flex", gap: 8, alignItems: "center", justifyContent: "flex-end", flexWrap: "wrap" },
  anchorInfoChips: { display: "flex", gap: 4, flexWrap: "wrap" },
  compactToggleButton: { border: "none", background: "transparent", padding: 0, display: "inline-flex", alignItems: "center", gap: 5, color: "var(--iccc-text-muted)", fontSize: 9.5, fontWeight: 700, cursor: "pointer" },
  compactToggleLabel: { lineHeight: 1 },
  compactToggleTrackOff: { width: 18, height: 11, borderRadius: 999, background: "#fb7185", display: "inline-flex", alignItems: "center", justifyContent: "flex-start", padding: 1, boxSizing: "border-box", flexShrink: 0 },
  compactToggleTrackOn: { width: 18, height: 11, borderRadius: 999, background: "#22c55e", display: "inline-flex", alignItems: "center", justifyContent: "flex-end", padding: 1, boxSizing: "border-box", flexShrink: 0 },
  compactToggleThumb: { width: 7, height: 7, borderRadius: 999, background: "#fff", boxShadow: "0 0 0 1px rgba(15,23,42,0.05), 0 1px 1px rgba(15,23,42,0.18)" },
  panelCard: { display: "grid", gap: 6, padding: 8, borderRadius: 18, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.8)" },
  sectionTitleRow: { display: "inline-flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  fieldLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--iccc-text-muted)" },
  compactRow: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  filterGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 6 },
  input: { width: "100%", borderRadius: 999, border: "1px solid var(--iccc-card-border)", padding: "7px 10px", background: "#fff", fontSize: 11, color: "var(--iccc-text)", boxSizing: "border-box" },
  select: { width: "100%", borderRadius: 999, border: "1px solid var(--iccc-card-border)", padding: "7px 10px", background: "#fff", fontSize: 11, color: "var(--iccc-text)" },
  primaryBtn: { ...baseButton, background: "linear-gradient(180deg, rgba(96,165,250,0.95) 0%, rgba(37,99,235,0.95) 100%)", color: "#fff", border: "1px solid rgba(37,99,235,0.35)" },
  secondaryBtn: { ...baseButton, background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)" },
  iconGhostBtn: { ...baseButton, width: 28, height: 28, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
  selectedGroupCard: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 8, alignItems: "center", borderRadius: 14, border: "1px solid rgba(37,99,235,0.16)", background: "rgba(239,246,255,0.82)", padding: 7 },
  selectedGroupMain: { display: "grid", gap: 2, minWidth: 0 },
  selectedGroupTitle: { fontSize: 11, fontWeight: 800, color: "var(--iccc-text)" },
  listWrap: { display: "grid", gap: 6 },
  listRow: { width: "100%", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "#fff", padding: "7px 9px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "var(--iccc-text)", fontSize: 11, fontWeight: 700 },
  listRowActive: { width: "100%", borderRadius: 12, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.78)", padding: "7px 9px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "#1d4ed8", fontSize: 11, fontWeight: 800 },
  countBadge: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 18, height: 18, borderRadius: 999, background: "rgba(15,23,42,0.06)", color: "var(--iccc-text)", fontSize: 9, fontWeight: 800 },
  warningBox: { padding: "7px 9px", borderRadius: 12, border: "1px solid rgba(245,158,11,0.22)", background: "rgba(255,247,237,0.9)", color: "#9a3412", fontSize: 10, lineHeight: 1.4 },
  smallMeta: { fontSize: 9.5, color: "var(--iccc-text-muted)", lineHeight: 1.35 },
  viewStack: { display: "grid", gap: 6 },
  inlineMetaRow: { display: "flex", flexWrap: "wrap", gap: 6, alignItems: "center" },
  tinyBtn: { ...baseButton, padding: "3px 8px", fontSize: 9.5, background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)" },
  emailList: { display: "grid", gap: 6 },
  emailCard: { display: "grid", gap: 0, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)", overflow: "hidden" },
  emailCardExpanded: { display: "grid", gap: 5, borderRadius: 14, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(255,255,255,0.96)", paddingBottom: 6, overflow: "hidden" },
  emailCardHead: { display: "grid", gridTemplateColumns: "24px minmax(0, 1fr)", gap: 4, alignItems: "start", padding: "7px 8px" },
  checkboxCell: { display: "inline-flex", alignItems: "center", justifyContent: "center", paddingTop: 2 },
  emailCardMain: { border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8, cursor: "pointer", textAlign: "left", minWidth: 0 },
  emailCardCopy: { display: "grid", gap: 2, minWidth: 0 },
  emailSubject: { display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap", fontSize: 11.5, fontWeight: 800, color: "var(--iccc-text)", lineHeight: 1.25 },
  emailMeta: { fontSize: 9.5, color: "var(--iccc-text-muted)", lineHeight: 1.3 },
  emailHeadBadges: { display: "inline-flex", alignItems: "center", gap: 4, color: "var(--iccc-text-muted)", paddingTop: 1 },
  badgeWrap: { display: "flex", gap: 4, flexWrap: "wrap" },
  detailBadgeStack: { display: "flex", flexWrap: "wrap", gap: 4, padding: "0 8px 0 28px" },
  anchorBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(37,99,235,0.08)", color: "#1d4ed8", fontSize: 8.5, fontWeight: 700 },
  statusBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(249,115,22,0.12)", color: "#c2410c", fontSize: 8.5, fontWeight: 700 },
  primaryBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(59,130,246,0.08)", color: "#2563eb", fontSize: 8.5, fontWeight: 700 },
  mutedBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(148,163,184,0.14)", color: "var(--iccc-text-muted)", fontSize: 8.5, fontWeight: 700 },
  selectedBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(30,64,175,0.1)", color: "#1d4ed8", fontSize: 8.5, fontWeight: 700 },
  warningBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(245,158,11,0.12)", color: "#b45309", fontSize: 8.5, fontWeight: 700 },
  labelBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(34,197,94,0.12)", color: "#15803d", fontSize: 8.5, fontWeight: 700 },
  readyBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(16,185,129,0.12)", color: "#047857", fontSize: 8.5, fontWeight: 700 },
  detailGrid: { display: "grid", gap: 5, padding: "0 12px" },
  detailRow: { display: "grid", gridTemplateColumns: "88px minmax(0, 1fr)", gap: 8, alignItems: "start" },
  detailLabel: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--iccc-text-muted)" },
  detailValue: { fontSize: 11, color: "var(--iccc-text)" },
  detailValuePrimary: { fontSize: 11, color: "#1d4ed8", fontWeight: 700 },
  warningSubtle: { margin: "0 8px 0 28px", padding: "6px 8px", borderRadius: 12, background: "rgba(255,247,237,0.86)", color: "#9a3412", fontSize: 9.5, lineHeight: 1.35 },
  attachmentList: { display: "grid", gap: 6 },
  attachmentRow: { display: "grid", gridTemplateColumns: "18px minmax(0, 1fr) auto", gap: 8, alignItems: "center", padding: "7px 9px", borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.9)" },
  attachmentRowActive: { display: "grid", gridTemplateColumns: "18px minmax(0, 1fr) auto", gap: 8, alignItems: "center", padding: "7px 9px", borderRadius: 14, border: "1px solid rgba(37,99,235,0.2)", background: "rgba(239,246,255,0.84)" },
  attachmentCopy: { display: "grid", gap: 2, minWidth: 0 },
  attachmentName: { fontSize: 11, fontWeight: 800, color: "var(--iccc-text)", wordBreak: "break-word" },
  metricGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(86px, 1fr))", gap: 6 },
  metricCard: { display: "grid", gap: 3, padding: 9, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.88)" },
  metricLabel: { fontSize: 9, fontWeight: 800, textTransform: "uppercase", color: "var(--iccc-text-muted)" },
  metricValue: { fontSize: 18, fontWeight: 800, color: "var(--iccc-text)" },
  metricMeta: { fontSize: 9.5, color: "var(--iccc-text-muted)", lineHeight: 1.3 },
  summaryList: { display: "grid", gap: 4, fontSize: 10.5, color: "var(--iccc-text)" },
  footerBar: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 8, alignItems: "center", padding: 8, borderRadius: 18, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.88)" },
  footerMain: { display: "grid", gap: 4, minWidth: 0 },
  footerStats: { display: "flex", gap: 4, flexWrap: "wrap" },
  footerStat: { display: "inline-flex", alignItems: "center", padding: "2px 7px", borderRadius: 999, background: "rgba(15,23,42,0.06)", color: "var(--iccc-text-muted)", fontSize: 8.5, fontWeight: 700 },
  footerCopy: { fontSize: 9, color: "var(--iccc-text-muted)", lineHeight: 1.3 },
  inlineActions: { display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" },
};

export default GroupsPrepareCockpit;
