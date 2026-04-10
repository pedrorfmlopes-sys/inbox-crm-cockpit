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
  hasGroupsPrepareSession,
  readGroupsPrepareSession,
  type GroupsPrepareAttachmentMode,
  type GroupsPrepareGroupMode,
  type GroupsPrepareSessionSaveReason,
  type GroupsPrepareSessionState,
  type GroupsPrepareSubview,
  writeGroupPreparationSeed,
  writeGroupsPrepareSession,
} from "@/modules/crm/groups-v1/prepareSession";
import { buildPrepareWorksetManifest } from "@/modules/crm/groups-v1/storage/buildPrepareWorksetManifest";
import { loadPrimaryGroupWorkset } from "@/modules/crm/groups-v1/storage/loadWorkset";
import { resolveGroupStorageRuntime, getGroupAttachmentStorageOptions } from "@/modules/crm/groups-v1/storage/resolveStorageMode";
import { savePrimaryGroupWorkset } from "@/modules/crm/groups-v1/storage/saveWorkset";
import { getGroupWorksetManifestSignature } from "@/modules/crm/groups-v1/storage/worksetManifest";
import { openGroupClassificationStudio } from "@/office";
import { getStatusDisplayConfig } from "@/statusUtils";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import type { GroupWorksetManifest } from "@/modules/crm/groups-v1/storage/types";

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
      style={{ ...S.compactToggleButton, ...(active ? S.compactToggleButtonOn : S.compactToggleButtonOff) }}
      onClick={onClick}
      aria-pressed={active}
      aria-label={`${label}: ${active ? "ativo" : "inativo"}`}
    >
      <span style={{ ...S.compactToggleLabel, ...(active ? S.compactToggleLabelOn : S.compactToggleLabelOff) }}>{label}</span>
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
  const runtime = useMemo(
    () => resolveGroupStorageRuntime(settings),
    [settings]
  );
  const canPersistWorkset = Boolean(settings) && (runtime.mode === "supabase" || runtime.mode === "hybrid");
  const hasStoredSessionRef = useRef(false);
  const persistedWorksetRef = useRef<GroupWorksetManifest | null>(null);
  const preferredGroupAppliedForEmailRef = useRef("");
  const renderedWorksetRef = useRef<{ worksetKey: string; manifest: GroupWorksetManifest | null; signature: string }>({
    worksetKey: "",
    manifest: null,
    signature: "",
  });
  const lastPersistedWorksetRef = useRef<{ worksetKey: string; signature: string }>({
    worksetKey: "",
    signature: "",
  });
  const hydratedWorksetScopeRef = useRef("");

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

  const flushWorkset = useCallback(async (
    reason: GroupsPrepareSessionSaveReason,
    options?: {
      force?: boolean;
      keepalive?: boolean;
      reportError?: boolean;
      manifest?: GroupWorksetManifest | null;
      signature?: string;
    }
  ): Promise<boolean> => {
    if (!canPersistWorkset) return false;
    const manifest = options?.manifest ?? renderedWorksetRef.current.manifest;
    const signature = String(options?.signature || (manifest ? getGroupWorksetManifestSignature(manifest) : ""));
    const worksetKey = String(manifest?.worksetKey || renderedWorksetRef.current.worksetKey || "").trim();
    if (!manifest || !worksetKey || !signature) return false;
    const lastPersisted = lastPersistedWorksetRef.current;
    if (!options?.force && lastPersisted.worksetKey === worksetKey && lastPersisted.signature === signature) {
      return true;
    }
    try {
      const saved = await savePrimaryGroupWorkset({
        runtime,
        manifest,
        current: persistedWorksetRef.current,
        keepalive: options?.keepalive === true || reason === "before_exit",
      });
      if (saved) {
        persistedWorksetRef.current = saved;
        const savedSignature = getGroupWorksetManifestSignature(saved);
        lastPersistedWorksetRef.current = { worksetKey: saved.worksetKey, signature: savedSignature };
        renderedWorksetRef.current = { worksetKey: saved.worksetKey, manifest: saved, signature: savedSignature };
        return true;
      }
      return false;
    } catch (error: unknown) {
      if (options?.reportError) {
        setMsg(getErrorMessage(error, `Nao foi possivel guardar o workset principal (${reason}).`));
      } else {
        console.warn("[GroupsPrepareCockpit] Workset save skipped:", error);
      }
      return false;
    }
  }, [canPersistWorkset, runtime, setMsg]);

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
          ...getGroupAttachmentStorageOptions(settings),
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
    settings,
    settings?.groupStorage?.baseFolderPath,
    settings?.groupStorage?.mode,
    settings?.groupStorage?.provider,
  ]);

  useEffect(() => {
    let cancelled = false;
    const query = String(workingGroupQuery || "").trim();
    if (!showGroupPanel || query.length < 2) {
      setGroups([]);
      setGroupsLoading(false);
      setGroupsError("");
      return () => {
        cancelled = true;
      };
    }
    const timer = window.setTimeout(() => {
      setGroupsLoading(true);
      setGroupsError("");
      listLinkGroups(query)
        .then((rows) => {
          if (!cancelled) setGroups(dedupeGroupEntries(Array.isArray(rows) ? rows : []).slice(0, 8));
        })
        .catch((error: unknown) => {
          if (cancelled) return;
          setGroups([]);
          setGroupsError(getErrorMessage(error, "Nao foi possivel pesquisar grupos."));
        })
        .finally(() => {
          if (!cancelled) setGroupsLoading(false);
        });
    }, 180);
    return () => {
      cancelled = true;
      window.clearTimeout(timer);
    };
  }, [showGroupPanel, workingGroupQuery]);

  useEffect(() => {
    const selectedGroup = groups.find((group) => group.id === workingGroupId);
    if (workingGroupId && selectedGroup && workingGroupQuery !== selectedGroup.name) {
      setWorkingGroupQuery(selectedGroup.name);
    }
  }, [groups, workingGroupId, workingGroupQuery]);

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
    const previousWorkset = renderedWorksetRef.current;
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
    if (
      previousWorkset.worksetKey
      && previousSession.emailKey
      && previousSession.emailKey !== currentEmailKey
      && previousSession.emailKey === sessionScopeKey
    ) {
      void flushWorkset("before_context_change", {
        force: true,
        manifest: previousWorkset.manifest,
        signature: previousWorkset.signature,
      });
    }
  }, [currentEmailKey, flushSession, flushWorkset, sessionScopeKey]);

  useEffect(() => {
    setSessionReady(false);
    setSessionScopeKey("");
    const sessionKey = String(currentEmailKey || "").trim();
    const hasStoredSession = hasGroupsPrepareSession(sessionKey);
    hasStoredSessionRef.current = hasStoredSession;
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
    persistedWorksetRef.current = null;
    renderedWorksetRef.current = { worksetKey: "", manifest: null, signature: "" };
    lastPersistedWorksetRef.current = { worksetKey: "", signature: "" };
    hydratedWorksetScopeRef.current = "";
    preferredGroupAppliedForEmailRef.current = "";
    setSessionScopeKey(sessionKey);
    setSessionReady(Boolean(sessionKey));
  }, [currentEmailKey]);

  const currentPrincipalGroup = useMemo(
    () => extractPrincipalGroup(currentEmailEntry),
    [currentEmailEntry]
  );

  const preferredWorkingGroupId = useMemo(() => {
    const providerGroupId =
      activeGroupSelection.emailKey === currentEmailKey
        ? String(activeGroupSelection.groupId || "").trim()
        : "";
    if (providerGroupId) return providerGroupId;
    return currentPrincipalGroup?.id || "";
  }, [activeGroupSelection, currentEmailKey, currentPrincipalGroup]);

  useEffect(() => {
    if (!sessionReady || workingGroupId || !preferredWorkingGroupId) return;
    if (preferredGroupAppliedForEmailRef.current === currentEmailKey) return;
    preferredGroupAppliedForEmailRef.current = currentEmailKey;
    setWorkingGroupId(preferredWorkingGroupId);
    const preferredGroup = groups.find((group) => group.id === preferredWorkingGroupId)
      || (currentPrincipalGroup?.id === preferredWorkingGroupId ? currentPrincipalGroup : null);
    if (preferredGroup) setWorkingGroupQuery(preferredGroup.name);
  }, [currentEmailKey, currentPrincipalGroup, groups, preferredWorkingGroupId, sessionReady, workingGroupId]);

  useEffect(() => {
    renderedSessionRef.current = {
      emailKey: String(sessionScopeKey || currentEmailKey || "").trim(),
      snapshot: sessionSnapshot,
      signature: sessionSignature,
    };
  }, [currentEmailKey, sessionScopeKey, sessionSignature, sessionSnapshot]);

  useEffect(() => {
    if (!canPersistWorkset || !sessionReady || !currentEmailKey || hydratedWorksetScopeRef.current === currentEmailKey) return;
    let cancelled = false;
    void loadPrimaryGroupWorkset({
      anchorEmailKey: currentEmailKey,
      runtime,
    })
      .then((manifest) => {
        if (cancelled || !manifest) return;
        persistedWorksetRef.current = manifest;
        const persistedSignature = getGroupWorksetManifestSignature(manifest);
        lastPersistedWorksetRef.current = {
          worksetKey: manifest.worksetKey,
          signature: persistedSignature,
        };
        if (hasStoredSessionRef.current) {
          hydratedWorksetScopeRef.current = currentEmailKey;
          return;
        }
        hydratedWorksetScopeRef.current = currentEmailKey;
        setWorkingGroupId((current) => current || manifest.workingGroupId || "");
        setWorkingGroupQuery((current) => current || manifest.workingGroupName || "");
        setFilterQuery((current) => current || String(manifest.filters?.query || "").trim());
        setAttachmentMode(manifest.filters?.attachmentMode || "all");
        setGroupMode(manifest.filters?.groupMode || "all");
        if (manifest.workingGroupId || manifest.workingGroupName) setShowGroupPanel(true);
        if (manifest.filters?.query || manifest.filters?.attachmentMode !== "all" || manifest.filters?.groupMode !== "all") {
          setShowFiltersPanel(true);
        }
        if (Array.isArray(manifest.includedEmailKeys) && manifest.includedEmailKeys.length) {
          setSelectedEmailKeys(manifest.includedEmailKeys);
        }
        const selectedAttachments = (manifest.attachments || [])
          .filter((attachment) => attachment.selection === "selected")
          .map((attachment) => attachment.key);
        if (selectedAttachments.length) {
          setSelectedAttachmentKeys(selectedAttachments);
        }
      })
      .catch((error: unknown) => {
        if (!cancelled) {
          console.warn("[GroupsPrepareCockpit] Persisted workset load skipped:", error);
        }
      });
    return () => {
      cancelled = true;
    };
  }, [canPersistWorkset, currentEmailKey, runtime, sessionReady]);

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
      void flushWorkset("before_exit", {
        force: true,
        keepalive: true,
        manifest: renderedWorksetRef.current.manifest,
        signature: renderedWorksetRef.current.signature,
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
  }, [flushSession, flushWorkset]);

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

  const workingGroup = useMemo<LinkGroupEntry | null>(() => {
    const found = groups.find((group) => group.id === workingGroupId) || null;
    if (found) return found;
    if (currentPrincipalGroup?.id === workingGroupId) {
      return { id: currentPrincipalGroup.id, name: currentPrincipalGroup.name, kind: "custom" };
    }
    if (workingGroupId) {
      return { id: workingGroupId, name: String(workingGroupQuery || "Grupo selecionado").trim(), kind: "custom" };
    }
    return null;
  }, [currentPrincipalGroup, groups, workingGroupId, workingGroupQuery]);

  const exactWorkingGroupMatch = useMemo(
    () => groups.find((group) => normalizeText(group.name) === normalizeText(workingGroupQuery))
      || (workingGroup && normalizeText(workingGroup.name) === normalizeText(workingGroupQuery) ? workingGroup : null)
      || null,
    [groups, workingGroup, workingGroupQuery]
  );

  const workingGroupCandidates = useMemo(() => {
    const query = normalizeText(workingGroupQuery);
    if (query.length < 2) return [];
    return [...groups]
      .filter((group) => !group.isArchived)
      .filter((group) => {
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

  const worksetManifest = useMemo(
    () => !canPersistWorkset ? null : buildPrepareWorksetManifest({
      anchorEmailKey: currentEmailKey,
      settings: runtime.settings,
      runtime,
      selectedEmailKeys,
      selectedAttachmentKeys,
      attachmentRows,
      workingGroupId,
      workingGroupName: String(workingGroup?.name || workingGroupQuery || "").trim() || undefined,
      filterQuery,
      attachmentMode,
      groupMode,
      previous: persistedWorksetRef.current,
    }),
    [
      attachmentMode,
      attachmentRows,
      canPersistWorkset,
      currentEmailKey,
      filterQuery,
      groupMode,
      runtime,
      selectedAttachmentKeys,
      selectedEmailKeys,
      workingGroup,
      workingGroupId,
      workingGroupQuery,
    ]
  );

  const worksetSignature = useMemo(
    () => getGroupWorksetManifestSignature(worksetManifest),
    [worksetManifest]
  );

  useEffect(() => {
    renderedWorksetRef.current = {
      worksetKey: String(worksetManifest?.worksetKey || "").trim(),
      manifest: worksetManifest,
      signature: worksetSignature,
    };
  }, [worksetManifest, worksetSignature]);

  useEffect(() => {
    if (!sessionReady || !currentEmailKey || !worksetManifest) return;
    const lastPersisted = lastPersistedWorksetRef.current;
    if (lastPersisted.worksetKey === worksetManifest.worksetKey && lastPersisted.signature === worksetSignature) {
      return;
    }
    const timer = window.setTimeout(() => {
      void flushWorkset("debounced", {
        manifest: worksetManifest,
        signature: worksetSignature,
      });
    }, 2200);
    return () => window.clearTimeout(timer);
  }, [currentEmailKey, flushWorkset, sessionReady, worksetManifest, worksetSignature]);

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
    await flushWorkset("before_classify", {
      force: true,
      reportError: true,
      manifest: worksetManifest,
      signature: worksetSignature,
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
    void flushWorkset("manual", {
      force: true,
      reportError: false,
      manifest: worksetManifest,
      signature: worksetSignature,
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
  const storageModeChip = useMemo(() => {
    if (runtime.mode === "hybrid") return { label: "Hibrido", style: S.hybridBadge };
    if (runtime.mode === "supabase") return { label: "Remoto", style: S.remoteBadge };
    return { label: "Local", style: S.localBadge };
  }, [runtime.mode]);
  const worksetStateChip = useMemo(() => {
    if (!canPersistWorkset) return { label: "Sessao", style: S.sessionBadge };
    if (lastPersistedWorksetRef.current.worksetKey) return { label: "Persistido", style: S.persistedBadge };
    if (worksetManifest) return { label: "Pendente", style: S.pendingBadge };
    return { label: "Draft", style: S.draftBadge };
  }, [canPersistWorkset, worksetManifest]);
  const anchorStatusConfig = getStatusDisplayConfig(currentEmailEntry.status);

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
          <div style={S.title}>Groups</div>
        </div>
        <button type="button" style={S.headerToolBtn} disabled title="Pesquisa entra noutra fase.">
          <Icons.Search size={11} />
        </button>
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
        <div style={S.anchorLead}>
          <div style={S.anchorIcon}>
            <Icons.MessageSquare size={12} />
          </div>
          <div style={S.anchorCopy}>
            <div style={S.fieldLabel}>Email ancora</div>
            <div style={S.anchorSubject}>{currentEmailEntry.subject || "(sem assunto)"}</div>
            <div style={S.anchorInfoChips}>
              <span style={S.mutedBadge}>{currentEmailEntry.fromName || currentEmailEntry.fromEmail || "Sem remetente"}</span>
              {formatDate(currentEmailEntry.messageDateIso || currentEmailEntry.receivedAtIso)
                ? <span style={S.mutedBadge}>{formatDate(currentEmailEntry.messageDateIso || currentEmailEntry.receivedAtIso)}</span>
                : null}
              <span style={storageModeChip.style}>{storageModeChip.label}</span>
              <span style={worksetStateChip.style}>{worksetStateChip.label}</span>
              <span style={{ ...S.statusBadge, ...anchorStatusConfig.style }}>{anchorStatusConfig.label}</span>
            </div>
          </div>
        </div>
        <div style={S.anchorActions}>
          <CompactToggle label="Grupo" active={showGroupPanel} onClick={() => setShowGroupPanel((value) => !value)} />
          <CompactToggle label="Filtros" active={showFiltersPanel} onClick={() => setShowFiltersPanel((value) => !value)} />
        </div>
      </div>

      {showGroupPanel ? (
        <div style={S.panelCard}>
          <div style={S.fieldLabel}>Grupo em trabalho</div>
          <div style={S.currentGroupLine}>
            {currentPrincipalGroup
              ? <>Grupo atual neste email: <strong>{currentPrincipalGroup.name}</strong></>
              : "Este email ainda nao tem grupo principal."}
          </div>
          <div style={S.searchBox}>
            <div style={S.searchInputWrap}>
              <Icons.Search size={11} />
              <input
                style={S.searchInput}
                value={workingGroupQuery}
                onChange={(event) => {
                  setWorkingGroupQuery(event.target.value);
                  setWorkingGroupId("");
                }}
                placeholder="Pesquisar grupo existente"
              />
            </div>
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
              <button type="button" style={S.iconGhostBtn} onClick={() => { setWorkingGroupId(""); setWorkingGroupQuery(""); }} title="Limpar grupo em trabalho">
                <Icons.RefreshCw size={12} />
              </button>
            </div>
          ) : null}
          {groupsError ? <PanelState compact tone="error" title="Falha a carregar grupos" description={groupsError} /> : null}
          {groupsLoading ? <PanelState compact tone="loading" title="A procurar grupos" description="Escreve para ver sugestoes compactas." /> : null}
          {!groupsLoading && workingGroupCandidates.length ? (
            <div style={S.suggestionDropdown}>
              {workingGroupCandidates.map((group) => (
                <button key={group.id} type="button" style={group.id === workingGroupId ? S.suggestionRowActive : S.suggestionRow} onClick={() => { setWorkingGroupId(group.id); setWorkingGroupQuery(group.name); }}>
                  <span>{group.name}</span>
                  <span style={S.countBadge}>{group.memberCount || 0}</span>
                </button>
              ))}
            </div>
          ) : !groupsLoading && normalizeText(workingGroupQuery).length >= 2 && !exactWorkingGroupMatch ? (
            <div style={S.smallMeta}>Sem sugestoes para esta pesquisa. Podes criar o grupo se fizer sentido.</div>
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
          <div style={S.fieldLabel}>Filtros de pesquisa</div>
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
                const emailStatusConfig = getStatusDisplayConfig(email.status);
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
                            <span style={emailKey === currentEmailKey ? S.subjectDotActive : S.subjectDot} />
                            <span style={S.emailSubjectText}>{email.subject || "(sem assunto)"}</span>
                            {emailKey === currentEmailKey ? <span style={S.anchorBadge}>Ancora</span> : null}
                            <span style={{ ...S.statusBadge, ...emailStatusConfig.style }}>{emailStatusConfig.label}</span>
                          </div>
                          <div style={S.emailMeta}>
                            {email.fromName || email.fromEmail || "Sem remetente"}
                            {formatDate(email.messageDateIso || email.receivedAtIso) ? ` · ${formatDate(email.messageDateIso || email.receivedAtIso)}` : ""}
                          </div>
                        </div>
                        <div style={S.emailHeadBadges}>
                          {selected ? <span style={worksetStateChip.style}>{worksetStateChip.label}</span> : null}
                          {emailKey === currentEmailKey ? <span style={storageModeChip.style}>{storageModeChip.label}</span> : null}
                          {attachmentCount ? <span style={S.countBadge}>{attachmentCount}</span> : null}
                          {expanded ? <Icons.ArrowUp size={12} /> : <Icons.ArrowDown size={12} />}
                        </div>
                      </button>
                    </div>

                    {expanded ? (
                      <>
                        {(email.labels || []).length ? (
                          <div style={S.detailBadgeStack}>
                            {(email.labels || []).slice(0, 3).map((label) => (
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
                          <span style={storageModeChip.style}>
                            Storage: {storageModeChip.label}
                          </span>
                          <span style={selected ? worksetStateChip.style : S.mutedBadge}>
                            Workset: {selected ? worksetStateChip.label : "Fora da selecao"}
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
            <div style={S.fieldLabel}>Anexos</div>
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
            <div style={S.fieldLabel}>Resumo antes de abrir no Classificar</div>
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
          <div style={S.footerCopy}>{canPersistWorkset ? `${worksetStateChip.label} / ${storageModeChip.label}` : "Sessao local."}</div>
        </div>
        <div style={S.inlineActions}>
          <button type="button" style={{ ...S.secondaryBtn, minWidth: 92 }} onClick={handleManualSessionSave}>
            <Icons.Save size={12} />
            Guardar
          </button>
          <button type="button" style={{ ...S.primaryBtn, minWidth: 102 }} onClick={() => void handleOpenClassificationFromPrepare()}>
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
  borderRadius: 12,
  border: "1px solid var(--iccc-card-border)",
  padding: "5px 9px",
  fontSize: 9,
  fontWeight: 700,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 4,
  lineHeight: 1.1,
};

const S: Record<string, React.CSSProperties> = {
  root: { display: "grid", gap: 4, alignContent: "start" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8, padding: "6px 7px", borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)" },
  headerMain: { display: "grid", gap: 1, minWidth: 0 },
  kicker: { fontSize: 8, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: "var(--iccc-text-muted)" },
  title: { fontSize: 12.5, fontWeight: 650, color: "#243244" },
  headerToolBtn: { ...baseButton, width: 24, height: 24, padding: 0, borderRadius: 999, background: "#fff", color: "var(--iccc-text-muted)", cursor: "not-allowed" },
  segmentBar: { display: "flex", gap: 2, padding: 2, borderRadius: 999, border: "1px solid rgba(148,163,184,0.22)", background: "rgba(241,245,249,0.92)", width: "100%", boxSizing: "border-box" },
  segment: { flex: "1 1 0", border: "none", background: "transparent", color: "var(--iccc-text-muted)", padding: "4px 7px", borderRadius: 999, fontSize: 9, fontWeight: 600, cursor: "pointer" },
  segmentActive: { flex: "1 1 0", border: "1px solid rgba(148,163,184,0.18)", background: "#fff", color: "var(--iccc-text)", padding: "4px 7px", borderRadius: 999, fontSize: 9, fontWeight: 700, cursor: "pointer", boxShadow: "0 1px 2px rgba(15,23,42,0.06)" },
  segmentDisabled: { flex: "1 1 0", border: "none", background: "transparent", color: "rgba(100,116,139,0.68)", padding: "4px 7px", borderRadius: 999, fontSize: 9, fontWeight: 600, cursor: "not-allowed" },
  anchorCard: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 6, alignItems: "start", padding: 7, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.9)" },
  anchorLead: { display: "flex", gap: 6, alignItems: "flex-start", minWidth: 0 },
  anchorIcon: { width: 22, height: 22, borderRadius: 8, background: "#27425f", color: "#fff", display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 1 },
  anchorCopy: { display: "grid", gap: 2, minWidth: 0 },
  anchorSubject: { fontSize: 11, fontWeight: 650, color: "#26364a", wordBreak: "break-word", lineHeight: 1.18 },
  anchorMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  anchorActions: { display: "flex", gap: 6, alignItems: "center", justifyContent: "flex-end", flexWrap: "nowrap", paddingTop: 1 },
  anchorInfoChips: { display: "flex", gap: 3, flexWrap: "wrap" },
  compactToggleButton: { border: "none", background: "transparent", padding: 0, display: "inline-flex", alignItems: "center", gap: 4, color: "var(--iccc-text-muted)", fontSize: 8, fontWeight: 700, cursor: "pointer" },
  compactToggleButtonOn: { color: "#15803d" },
  compactToggleButtonOff: { color: "#b91c1c" },
  compactToggleLabel: { lineHeight: 1 },
  compactToggleLabelOn: { color: "#15803d" },
  compactToggleLabelOff: { color: "#b91c1c" },
  compactToggleTrackOff: { width: 15, height: 9, borderRadius: 999, background: "rgba(239,68,68,0.72)", display: "inline-flex", alignItems: "center", justifyContent: "flex-start", padding: 1, boxSizing: "border-box", flexShrink: 0 },
  compactToggleTrackOn: { width: 15, height: 9, borderRadius: 999, background: "rgba(34,197,94,0.78)", display: "inline-flex", alignItems: "center", justifyContent: "flex-end", padding: 1, boxSizing: "border-box", flexShrink: 0 },
  compactToggleThumb: { width: 5, height: 5, borderRadius: 999, background: "#fff" },
  panelCard: { display: "grid", gap: 5, padding: 6, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.82)" },
  sectionTitleRow: { display: "inline-flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  fieldLabel: { fontSize: 8, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", color: "#64748b" },
  compactRow: { display: "flex", gap: 5, alignItems: "center", flexWrap: "wrap" },
  currentGroupLine: { fontSize: 8.7, color: "#526173", lineHeight: 1.25 },
  searchBox: { display: "flex", gap: 5, alignItems: "center" },
  searchInputWrap: { flex: "1 1 0", minWidth: 0, display: "flex", alignItems: "center", gap: 4, borderRadius: 11, border: "1px solid rgba(148,163,184,0.32)", background: "rgba(255,255,255,0.94)", color: "#64748b", padding: "4px 7px", boxSizing: "border-box" },
  searchInput: { width: "100%", minWidth: 0, border: "none", outline: "none", background: "transparent", fontSize: 9.5, color: "#26364a", padding: 0 },
  filterGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(132px, 1fr))", gap: 5 },
  input: { width: "100%", borderRadius: 11, border: "1px solid rgba(148,163,184,0.32)", padding: "5px 8px", background: "rgba(255,255,255,0.94)", fontSize: 9.5, color: "#26364a", boxSizing: "border-box" },
  select: { width: "100%", borderRadius: 11, border: "1px solid rgba(148,163,184,0.32)", padding: "5px 8px", background: "rgba(255,255,255,0.94)", fontSize: 9.5, color: "#26364a" },
  primaryBtn: { ...baseButton, background: "#27425f", color: "#fff", border: "1px solid #27425f" },
  secondaryBtn: { ...baseButton, background: "rgba(255,255,255,0.88)", color: "#334155" },
  iconGhostBtn: { ...baseButton, width: 22, height: 22, padding: 0, background: "rgba(255,255,255,0.9)", color: "#526173" },
  selectedGroupCard: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 6, alignItems: "center", borderRadius: 11, border: "1px solid rgba(39,66,95,0.16)", background: "rgba(248,250,252,0.95)", padding: 6 },
  selectedGroupMain: { display: "grid", gap: 2, minWidth: 0 },
  selectedGroupTitle: { fontSize: 10, fontWeight: 700, color: "var(--iccc-text)" },
  listWrap: { display: "grid", gap: 4 },
  listRow: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "#fff", padding: "6px 8px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "var(--iccc-text)", fontSize: 10, fontWeight: 700 },
  listRowActive: { width: "100%", borderRadius: 10, border: "1px solid rgba(15,23,42,0.18)", background: "rgba(248,250,252,0.95)", padding: "6px 8px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "var(--iccc-text)", fontSize: 10, fontWeight: 700 },
  suggestionDropdown: { display: "grid", gap: 2, borderRadius: 11, border: "1px solid rgba(148,163,184,0.24)", background: "rgba(255,255,255,0.96)", padding: 3, boxShadow: "0 8px 20px rgba(15,23,42,0.08)" },
  suggestionRow: { width: "100%", border: "none", borderRadius: 8, background: "transparent", padding: "4px 6px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "#26364a", fontSize: 9.5, fontWeight: 600, textAlign: "left" },
  suggestionRowActive: { width: "100%", border: "none", borderRadius: 8, background: "rgba(219,234,254,0.55)", padding: "4px 6px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, cursor: "pointer", color: "#1e3a5f", fontSize: 9.5, fontWeight: 650, textAlign: "left" },
  countBadge: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 14, height: 14, borderRadius: 999, background: "rgba(71,85,105,0.08)", color: "#526173", fontSize: 7.5, fontWeight: 700 },
  warningBox: { padding: "6px 8px", borderRadius: 10, border: "1px solid rgba(245,158,11,0.16)", background: "rgba(255,247,237,0.8)", color: "#9a3412", fontSize: 9, lineHeight: 1.35 },
  smallMeta: { fontSize: 8.5, color: "#526173", lineHeight: 1.25 },
  viewStack: { display: "grid", gap: 5 },
  inlineMetaRow: { display: "flex", flexWrap: "wrap", gap: 5, alignItems: "center" },
  tinyBtn: { ...baseButton, padding: "2px 7px", fontSize: 8, background: "rgba(255,255,255,0.88)", color: "var(--iccc-text-muted)" },
  emailList: { display: "grid", gap: 5 },
  emailCard: { display: "grid", gap: 0, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.94)", overflow: "hidden" },
  emailCardExpanded: { display: "grid", gap: 4, borderRadius: 12, border: "1px solid rgba(39,66,95,0.18)", background: "rgba(255,255,255,0.98)", paddingBottom: 4, overflow: "hidden" },
  emailCardHead: { display: "grid", gridTemplateColumns: "18px minmax(0, 1fr)", gap: 4, alignItems: "start", padding: "5px 6px" },
  checkboxCell: { display: "inline-flex", alignItems: "center", justifyContent: "center", paddingTop: 1 },
  emailCardMain: { border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 6, cursor: "pointer", textAlign: "left", minWidth: 0 },
  emailCardCopy: { display: "grid", gap: 1, minWidth: 0 },
  emailSubject: { display: "flex", alignItems: "center", gap: 4, flexWrap: "wrap", color: "var(--iccc-text)", lineHeight: 1.15 },
  subjectDot: { width: 5, height: 5, borderRadius: 999, background: "rgba(100,116,139,0.42)", flexShrink: 0 },
  subjectDotActive: { width: 5, height: 5, borderRadius: 999, background: "#27425f", flexShrink: 0 },
  emailSubjectText: { fontSize: 10.5, fontWeight: 650, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", color: "#26364a" },
  emailMeta: { fontSize: 8.5, color: "#526173", lineHeight: 1.2 },
  emailHeadBadges: { display: "inline-flex", alignItems: "center", gap: 3, color: "var(--iccc-text-muted)", paddingTop: 1 },
  badgeWrap: { display: "flex", gap: 4, flexWrap: "wrap" },
  detailBadgeStack: { display: "flex", flexWrap: "wrap", gap: 3, padding: "0 6px 0 20px" },
  anchorBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(71,85,105,0.08)", color: "#526173", fontSize: 7.5, fontWeight: 700 },
  statusBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, border: "1px solid transparent", background: "rgba(249,115,22,0.1)", color: "#c2410c", fontSize: 7.5, fontWeight: 700 },
  primaryBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(39,66,95,0.08)", color: "#26364a", fontSize: 7.5, fontWeight: 700 },
  mutedBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(148,163,184,0.12)", color: "#526173", fontSize: 7.5, fontWeight: 700 },
  selectedBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(39,66,95,0.09)", color: "#26364a", fontSize: 7.5, fontWeight: 700 },
  warningBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(245,158,11,0.1)", color: "#b45309", fontSize: 7.5, fontWeight: 700 },
  labelBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(34,197,94,0.09)", color: "#15803d", fontSize: 7.5, fontWeight: 700 },
  readyBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(16,185,129,0.1)", color: "#047857", fontSize: 7.5, fontWeight: 700 },
  localBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(220,252,231,0.78)", color: "#15803d", border: "1px solid rgba(34,197,94,0.18)", fontSize: 7.5, fontWeight: 700 },
  remoteBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(219,234,254,0.82)", color: "#1d4ed8", border: "1px solid rgba(59,130,246,0.2)", fontSize: 7.5, fontWeight: 700 },
  hybridBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(237,233,254,0.82)", color: "#6d28d9", border: "1px solid rgba(124,58,237,0.18)", fontSize: 7.5, fontWeight: 700 },
  sessionBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(241,245,249,0.86)", color: "#475569", border: "1px solid rgba(148,163,184,0.18)", fontSize: 7.5, fontWeight: 700 },
  draftBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(255,247,237,0.82)", color: "#c2410c", border: "1px solid rgba(249,115,22,0.18)", fontSize: 7.5, fontWeight: 700 },
  pendingBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(254,243,199,0.82)", color: "#b45309", border: "1px solid rgba(245,158,11,0.18)", fontSize: 7.5, fontWeight: 700 },
  persistedBadge: { display: "inline-flex", alignItems: "center", padding: "1px 5px", borderRadius: 999, background: "rgba(220,252,231,0.88)", color: "#15803d", border: "1px solid rgba(34,197,94,0.2)", fontSize: 7.5, fontWeight: 700 },
  detailGrid: { display: "grid", gap: 5, padding: "0 12px" },
  detailRow: { display: "grid", gridTemplateColumns: "88px minmax(0, 1fr)", gap: 8, alignItems: "start" },
  detailLabel: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.04em", color: "var(--iccc-text-muted)" },
  detailValue: { fontSize: 11, color: "var(--iccc-text)" },
  detailValuePrimary: { fontSize: 11, color: "#1d4ed8", fontWeight: 700 },
  warningSubtle: { margin: "0 6px 0 20px", padding: "4px 6px", borderRadius: 10, background: "rgba(255,247,237,0.78)", color: "#9a3412", fontSize: 8.5, lineHeight: 1.25 },
  attachmentList: { display: "grid", gap: 5 },
  attachmentRow: { display: "grid", gridTemplateColumns: "16px minmax(0, 1fr) auto", gap: 7, alignItems: "center", padding: "6px 8px", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)" },
  attachmentRowActive: { display: "grid", gridTemplateColumns: "16px minmax(0, 1fr) auto", gap: 7, alignItems: "center", padding: "6px 8px", borderRadius: 12, border: "1px solid rgba(15,23,42,0.14)", background: "rgba(248,250,252,0.95)" },
  attachmentCopy: { display: "grid", gap: 2, minWidth: 0 },
  attachmentName: { fontSize: 10, fontWeight: 700, color: "var(--iccc-text)", wordBreak: "break-word" },
  metricGrid: { display: "grid", gridTemplateColumns: "repeat(3, minmax(0, 1fr))", gap: 5 },
  metricCard: { display: "grid", gap: 2, padding: 6, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.9)" },
  metricLabel: { fontSize: 8, fontWeight: 800, textTransform: "uppercase", color: "var(--iccc-text-muted)" },
  metricValue: { fontSize: 14, fontWeight: 700, color: "var(--iccc-text)" },
  metricMeta: { fontSize: 8, color: "var(--iccc-text-muted)", lineHeight: 1.2 },
  summaryList: { display: "grid", gap: 3, fontSize: 9.5, color: "var(--iccc-text)" },
  footerBar: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: 6, alignItems: "center", padding: 6, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.9)" },
  footerMain: { display: "grid", gap: 2, minWidth: 0 },
  footerStats: { display: "flex", gap: 3, flexWrap: "wrap" },
  footerStat: { display: "inline-flex", alignItems: "center", padding: "1px 6px", borderRadius: 999, background: "rgba(15,23,42,0.06)", color: "var(--iccc-text-muted)", fontSize: 7.5, fontWeight: 700 },
  footerCopy: { fontSize: 8, color: "var(--iccc-text-muted)", lineHeight: 1.1 },
  inlineActions: { display: "flex", gap: 4, flexWrap: "nowrap", justifyContent: "flex-end" },
};

export default GroupsPrepareCockpit;
