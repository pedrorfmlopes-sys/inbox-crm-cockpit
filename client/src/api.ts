// Goal: keep UI stable even if server endpoints evolve.
import { getSettings, saveSettings, type CockpitSettingsV1 } from "./settings";

let _sessionToken: string | null = null;
let _sessionBootstrapPromise: Promise<string | null> | null = null;
const SESSION_BOOTSTRAP_TIMEOUT_MS = 10000;
const API_REQUEST_TIMEOUT_MS = 30000;

export function setApiSessionToken(token: string | null) {
  _sessionToken = token;
}

function isOdooApiPath(path: string): boolean {
  return String(path || "").startsWith("/api/odoo");
}

function shouldRetryOdooWithSavedSettings(message: string): boolean {
  return /ODOO_CONFIG_MISSING|Odoo configuration incomplete|Sess[aã]o expirada|Session expired|HTTP 401/i.test(message);
}

function getJsonErrorMessage(body: any): string | null {
  if (body == null || typeof body !== "object") return null;
  if (body.ok === false || body.error || body.message || body.details) {
    return String(body.details || body.message || body.error || "Unknown Odoo error");
  }
  return null;
}

async function bootstrapSessionFromSavedSettings(forceRefresh = false): Promise<string | null> {
  if (!forceRefresh && _sessionToken) return _sessionToken;
  if (_sessionBootstrapPromise) return _sessionBootstrapPromise;

  _sessionBootstrapPromise = (async () => {
    const settings = await getSettings();
    const savedToken = String(settings.odooSessionToken || "").trim();
    const url = String(settings.odooUrl || "").trim();
    const db = String(settings.odooDb || "").trim();
    const login = String(settings.odooLogin || "").trim();
    const password = String(settings.odooPassword || "").trim();

    if (!forceRefresh && savedToken) {
      _sessionToken = savedToken;
      return savedToken;
    }

    if (!url || !db || !login || !password) {
      return savedToken || null;
    }

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), SESSION_BOOTSTRAP_TIMEOUT_MS);

    try {
      const res = await fetch("/api/auth/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ url, db, login, password }),
        signal: controller.signal,
      });

      const ct = (res.headers.get("content-type") || "").toLowerCase();
      const body = ct.includes("application/json") ? await res.json() : await res.text();

      if (!res.ok || !body?.ok || !body?.token) {
        const msg =
          typeof body === "string"
            ? body
            : body?.message || body?.error || JSON.stringify(body);
        throw new Error(`HTTP ${res.status}: ${msg}`);
      }

      const token = String(body.token || "").trim();
      _sessionToken = token || null;
      if (token && token !== savedToken) {
        await saveSettings({ odooSessionToken: token });
      }
      return token || null;
    } catch (error: any) {
      if (error?.name === "AbortError") {
        throw new Error("Odoo session bootstrap timed out");
      }
      throw error;
    } finally {
      clearTimeout(timeoutId);
    }
  })().finally(() => {
    _sessionBootstrapPromise = null;
  });

  return _sessionBootstrapPromise;
}

export type OdooMeta = {
  ok: boolean;
  baseUrl?: string;     // preferred
  webBaseUrl?: string;  // compat (algum código ainda usa isto)
  url?: string;         // compat (outro código ainda usa isto)
  db?: string;
  user?: { id: number; name: string; login?: string } | null;
  models?: string[];
};

export type OdooMetaResponse = { ok: boolean; meta: OdooMeta };

const isMock = () => typeof window !== "undefined" && !!(window as any).__ICCC_MOCK__;
const getMockData = (type: string, model: string) => (window as any).__ICCC_MOCK_DATA?.[type]?.[model] || null;

export type LinkEntry = {
  id?: string;
  conversationId: string;
  model: string;
  itemId?: string;
  emailWebLink?: string;
  messageDateIso?: string;
  sentAtIso?: string;
  receivedAtIso?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  linkedAt?: string;
  internetMessageId?: string;

  // preferred
  recordId?: number;
  recordName?: string;

  // compat aliases
  resId?: number;
  name?: string;

  // display helpers
  title?: string;
  url?: string;

  createdAt?: string;
  updatedAt?: string;
};

export type RelevantEmailPayload = {
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  emailWebLink?: string;
  messageDateIso?: string;
  sentAtIso?: string;
  receivedAtIso?: string;
  bodyText?: string;
  bodyHtml?: string;
  status?: "em_analise" | "em_progresso" | "concluido" | string;
  labels?: string[];
  removedInheritedLabels?: string[];
  labelStates?: Record<string, "em_analise" | "em_progresso" | "concluido" | string>;
  classificationMeta?: {
    principalStatusEnabled?: boolean;
    principalStatusCategorize?: boolean;
    referenceStatusEnabled?: boolean;
    referenceStatusCategorize?: boolean;
    ticketStatusEnabled?: boolean;
    ticketStatusCategorize?: boolean;
    categorizedLabelNames?: string[];
  };
  attachmentStorageProvider?: "cloud" | "local" | "onedrive" | string;
  attachmentStorageBasePath?: string;
  membershipKind?: "principal" | "referencia" | string;
  attachments?: Array<{
    key?: string;
    id?: string;
    name: string;
    contentType?: string;
    size?: number;
    isInline?: boolean;
    contentId?: string;
    content?: string;
    storageProvider?: "cloud" | "local" | "onedrive" | string;
    storageBasePath?: string;
    storagePathHint?: string;
    documentState?: "ingested" | "processed" | "accepted" | "rejected" | "reread_requested" | string;
    hasContent?: boolean;
  }>;
};

export type AttachmentTextExtractionEntry = {
  key: string;
  name: string;
  contentType?: string;
  text: string;
};

export type LinkGroupEntry = {
  id: string;
  kind: "custom" | "conversation" | string;
  name: string;
  description?: string;
  notes?: string;
  contacts?: Array<{
    key: string;
    name: string;
    email?: string;
    company?: string;
    source?: string;
  }>;
  entities?: Array<{
    key: string;
    name: string;
    kind?: string;
    source?: string;
  }>;
  conversationId?: string;
  status?: "em_analise" | "em_progresso" | "concluido" | string;
  labels?: string[];
  isArchived?: boolean;
  archivedAt?: string;
  memberCount?: number;
  documentCount?: number;
  documentsEnabled?: boolean;
  createdAt?: string;
  updatedAt?: string;
};

export type GroupDocumentEntry = {
  id: string;
  name: string;
  contentType?: string;
  size?: number;
  contentBase64?: string;
  hasContent?: boolean;
  documentState?: "ingested" | "processed" | "accepted" | "rejected" | "reread_requested" | string;
  sourceEmailKey?: string;
  sourceItemId?: string;
  sourceInternetMessageId?: string;
  sourceConversationId?: string;
  sourceEmailSubject?: string;
  storageProvider?: string;
  storageBasePath?: string;
  storagePathHint?: string;
  createdAt?: string;
  updatedAt?: string;
};

export type GroupAttachmentFlagEntry = {
  attachmentKey: string;
  emailKey?: string;
  attachmentName?: string;
  contentType?: string;
  size?: number;
  disposition?: string;
  createdAt?: string;
  updatedAt?: string;
};

export type GroupTicketSeriesEntry = {
  id: string;
  name: string;
  prefix: string;
  replyInstructions?: string;
  yearMode?: "none" | "yy" | "yyyy";
  separator?: "-" | "/" | "_" | " " | "";
  nextNumber: number;
  padding: number;
  isActive?: boolean;
  usageCount?: number;
  createdAt?: string;
  updatedAt?: string;
};

export type GroupTicketEntry = {
  id: string;
  seriesId: string;
  seriesName?: string;
  prefix?: string;
  yearMode?: "none" | "yy" | "yyyy";
  separator?: "-" | "/" | "_" | " " | "";
  yearValue?: string;
  code: string;
  sequenceNumber: number;
  title: string;
  description?: string;
  status?: "open" | "closed" | string;
  labels?: string[];
  groupIds?: string[];
  groups?: LinkGroupEntry[];
  emailCount?: number;
  emailLinked?: boolean;
  createdFromEmailKey?: string;
  createdAt?: string;
  updatedAt?: string;
};

export type GroupTicketDetectionMatch = {
  matchedCode: string;
  ticket: GroupTicketEntry;
  emailLinked: boolean;
  proposedGroups: LinkGroupEntry[];
};

export type RelatedReason =
  | {
    kind: "entity";
    model: string;
    recordId: number;
    recordName?: string;
  }
  | {
    kind: "group" | "conversation";
    groupId: string;
    groupName?: string;
    conversationId?: string;
  };

export type RelatedEmailEntry = Omit<LinkEntry, "model" | "recordId" | "recordName" | "resId" | "name" | "title"> & {
  emailKey?: string;
  status?: string;
  labels?: string[];
  removedInheritedLabels?: string[];
  labelStates?: Record<string, string>;
  classificationMeta?: {
    principalStatusEnabled?: boolean;
    principalStatusCategorize?: boolean;
    referenceStatusEnabled?: boolean;
    referenceStatusCategorize?: boolean;
    ticketStatusEnabled?: boolean;
    ticketStatusCategorize?: boolean;
    categorizedLabelNames?: string[];
  };
  membershipKind?: "principal" | "referencia" | string;
  relatedRecords?: Array<{ model: string; recordId: number; recordName: string }>;
  relatedGroups?: Array<{ id: string; name?: string; kind?: string; relationKind?: "principal" | "referencia" | string }>;
  relatedReasons?: RelatedReason[];
  groupId?: string;
  groupName?: string;
  bodyText?: string;
  bodyHtml?: string;
  attachments?: Array<{
    key?: string;
    id?: string;
    name: string;
    contentType?: string;
    size?: number;
    isInline?: boolean;
    contentId?: string;
    content?: string;
    storageProvider?: "cloud" | "local" | "onedrive" | string;
    storageBasePath?: string;
    storagePathHint?: string;
    documentState?: "ingested" | "processed" | "accepted" | "rejected" | "reread_requested" | string;
    hasContent?: boolean;
  }>;
};

export type LinkPayload = {
  conversationId: string;
  model: string;

  // preferred
  recordId: number;
  recordName?: string;

  emailSubject?: string;
  emailFrom?: string;
  emailWebLink?: string;
  internetMessageId?: string;
  itemId?: string;
  receivedAtIso?: string;
  bodyHtml?: string;
  bodyText?: string;
  postToChatter?: boolean;
  attachmentIds?: number[];

  // compat aliases
  resId?: number;
  name?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  url?: string;
};

export type AiGenerateResponse =
  | {
    ok: true;
    html?: string;
    text: string;
    data?: any;
    suggestedRecipients?: { to: string[]; cc: string[] };
    suggestedSubject?: string;
  }
  | { ok: false; error: string };

export type AuthResponse = { ok: true; token: string; meta: OdooMeta } | { ok: false; message: string };

type Json = any;

async function requestJSON<T = Json>(path: string, init?: RequestInit, allowOdooRetry = true): Promise<T> {
  if (isOdooApiPath(path) && !_sessionToken) {
    try {
      await bootstrapSessionFromSavedSettings(false);
    } catch {
      // Surface the original backend/auth error from the real request.
    }
  }

  const controller = new AbortController();
  const timeoutMessage = `Pedido excedeu ${Math.round(API_REQUEST_TIMEOUT_MS / 1000)}s: ${String(path || "")}`;
  const id = setTimeout(() => controller.abort(timeoutMessage), API_REQUEST_TIMEOUT_MS);

  try {
    const res = await fetch(path, {
      ...init,
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        ...(_sessionToken ? { "Authorization": `Session ${_sessionToken}` } : {}),
        ...(init?.headers || {}),
      },
    });

    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const body = ct.includes("application/json") ? await res.json() : await res.text();
    const bodyErrorMessage = typeof body === "string" ? null : getJsonErrorMessage(body);

    if (!res.ok) {
      const msg =
        typeof body === "string"
          ? body
          : body?.error || body?.message || JSON.stringify(body);

      if (allowOdooRetry && isOdooApiPath(path) && shouldRetryOdooWithSavedSettings(msg)) {
        const renewedToken = await bootstrapSessionFromSavedSettings(true).catch(() => null);
        if (renewedToken) {
          return await requestJSON<T>(path, init, false);
        }
      }
      throw new Error(`HTTP ${res.status}: ${msg}`);
    }

    if (bodyErrorMessage) {
      if (allowOdooRetry && isOdooApiPath(path) && shouldRetryOdooWithSavedSettings(bodyErrorMessage)) {
        const renewedToken = await bootstrapSessionFromSavedSettings(true).catch(() => null);
        if (renewedToken) {
          return await requestJSON<T>(path, init, false);
        }
      }
      throw new Error(bodyErrorMessage);
    }

    return body as T;
  } catch (error: any) {
    const aborted = controller.signal.aborted || error?.name === "AbortError";
    if (aborted) {
      const reason =
        typeof controller.signal.reason === "string" && controller.signal.reason.trim()
          ? controller.signal.reason.trim()
          : timeoutMessage;
      throw new Error(reason);
    }
    throw error;
  } finally {
    clearTimeout(id);
  }
}

// -------- Auth --------
export async function login(credentials: any): Promise<AuthResponse> {
  return await requestJSON(`/api/auth/login`, {
    method: "POST",
    body: JSON.stringify(credentials),
  });
}

export async function checkAuth(): Promise<{ ok: boolean; meta?: OdooMeta }> {
  return await requestJSON(`/api/auth/check`);
}

// -------- Odoo meta / ping --------
export async function getOdooMeta(): Promise<OdooMeta> {
  const r: any = await requestJSON(`/api/odoo/meta`);
  const m = (r?.meta ?? r) as OdooMeta;
  if (m) {
    // Standardize URL property
    m.baseUrl = m.baseUrl || m.webBaseUrl || m.url;
  }
  return m;
}

export function getOdooAutoLoginUrl(_token: string | null, redirect: string = "/web", odooBaseUrl?: string): string {
  let targetPath = redirect;
  if (odooBaseUrl && redirect.startsWith(odooBaseUrl)) {
    targetPath = redirect.substring(odooBaseUrl.length);
  }
  if (!targetPath.startsWith("/")) targetPath = "/" + targetPath;

  if (!odooBaseUrl) {
    return targetPath;
  }

  const base = String(odooBaseUrl).replace(/\/+$/, "");
  const redirectUrl = new URL(targetPath, base + "/");
  const db = redirectUrl.searchParams.get("db");
  const loginUrl = new URL(base + "/web/login");

  if (db) {
    loginUrl.searchParams.set("db", db);
  }
  loginUrl.searchParams.set("redirect", redirectUrl.pathname + redirectUrl.search + redirectUrl.hash);

  return loginUrl.toString();
}

export async function odooPing(): Promise<{ ok: boolean }> {
  return await requestJSON(`/api/odoo/ping`);
}

export type Crm2LayoutValidationCheck = {
  key: string;
  label: string;
  kind: "field" | "tab";
  configuredName: string;
  status: "ok" | "warning" | "error";
  message: string;
  details?: string;
  actualType?: string;
  expectedTypes?: string[];
  recommendedType?: string;
  presentInFormView?: boolean;
};

export type Crm2LayoutValidationResult = {
  ok: boolean;
  target?: "project" | "lead" | "task" | "ticket";
  mode: "description_only" | "structured_project";
  model: string;
  ready: boolean;
  summary: {
    ok: number;
    warning: number;
    error: number;
  };
  checks: Crm2LayoutValidationCheck[];
  formView?: {
    available: boolean;
    tabTitles?: string[];
    error?: string;
  };
};

export async function validateCrm2OdooLayout(
  layout: CockpitSettingsV1["crm2OdooLayout"],
  target: "project" | "lead" | "task" | "ticket" = "project",
): Promise<Crm2LayoutValidationResult> {
  return await requestJSON(`/api/odoo/layout/validate`, {
    method: "POST",
    body: JSON.stringify({ layout, target }),
  });
}

// -------- Links --------
function normalizeLinkEntry(link: any): LinkEntry {
  return {
    ...link,
    resId: link?.resId ?? link?.recordId,
    recordId: link?.recordId ?? link?.resId,
    name: link?.name ?? link?.recordName ?? link?.title,
    title: link?.title ?? link?.recordName ?? link?.name ?? link?.subject ?? link?.model,
    url: link?.url ?? link?.emailWebLink,
  };
}

function normalizeRelatedEmailEntry(entry: any): RelatedEmailEntry {
  const normalizedLabelStates = entry?.labelStates && typeof entry.labelStates === "object"
    ? Object.fromEntries(
        Object.entries(entry.labelStates)
          .map(([label, value]) => [String(label || "").trim(), String(value || "").trim()])
          .filter(([label, value]) => label && value)
      )
    : {};
  const normalizedClassificationMeta = entry?.classificationMeta && typeof entry.classificationMeta === "object"
    ? {
        principalStatusEnabled: entry.classificationMeta.principalStatusEnabled === true,
        principalStatusCategorize: entry.classificationMeta.principalStatusCategorize === true,
        referenceStatusEnabled: entry.classificationMeta.referenceStatusEnabled === true,
        referenceStatusCategorize: entry.classificationMeta.referenceStatusCategorize === true,
        ticketStatusEnabled: entry.classificationMeta.ticketStatusEnabled === true,
        ticketStatusCategorize: entry.classificationMeta.ticketStatusCategorize === true,
        categorizedLabelNames: Array.isArray(entry.classificationMeta.categorizedLabelNames)
          ? entry.classificationMeta.categorizedLabelNames.map((label: any) => String(label || "").trim()).filter(Boolean)
          : undefined,
      }
    : undefined;
  return {
    ...normalizeLinkEntry(entry),
    emailKey: String(entry?.emailKey || "").trim(),
    status: String(entry?.status || "").trim() || undefined,
    labels: Array.isArray(entry?.labels) ? entry.labels.map((label: any) => String(label || "").trim()).filter(Boolean) : [],
    removedInheritedLabels: Array.isArray(entry?.removedInheritedLabels) ? entry.removedInheritedLabels.map((label: any) => String(label || "").trim()).filter(Boolean) : [],
    labelStates: normalizedLabelStates,
    classificationMeta: normalizedClassificationMeta,
    membershipKind: String(entry?.membershipKind || "").trim() || undefined,
    bodyText: String(entry?.bodyText || "").trim(),
    bodyHtml: String(entry?.bodyHtml || "").trim(),
    relatedRecords: Array.isArray(entry?.relatedRecords)
      ? entry.relatedRecords.map((record: any) => ({
        model: String(record?.model || "").trim(),
        recordId: Number(record?.recordId || 0),
        recordName: String(record?.recordName || "").trim(),
      })).filter((record: any) => record.model && record.recordId)
      : [],
    relatedGroups: Array.isArray(entry?.relatedGroups)
      ? entry.relatedGroups.map((group: any) => ({
        id: String(group?.id || "").trim(),
        name: String(group?.name || "").trim(),
        kind: String(group?.kind || "").trim(),
        relationKind: String(group?.relationKind || "").trim() || undefined,
      })).filter((group: any) => group.id)
      : [],
    attachments: Array.isArray(entry?.attachments)
      ? entry.attachments
        .map((attachment: any) => ({
          key: String(attachment?.key || "").trim() || undefined,
          id: String(attachment?.id || "").trim() || undefined,
          name: String(attachment?.name || "").trim(),
          contentType: String(attachment?.contentType || "").trim(),
          size: Number(attachment?.size || 0) || undefined,
          isInline: Boolean(attachment?.isInline),
          contentId: String(attachment?.contentId || "").trim() || undefined,
          content: String(attachment?.content || "").trim(),
          storageProvider: String(attachment?.storageProvider || "").trim() || undefined,
          storageBasePath: String(attachment?.storageBasePath || "").trim() || undefined,
          storagePathHint: String(attachment?.storagePathHint || "").trim() || undefined,
          documentState: String(attachment?.documentState || "").trim() || undefined,
          hasContent: attachment?.hasContent === true || Boolean(String(attachment?.content || "").trim()),
        }))
        .filter((attachment: any) => attachment.name)
      : [],
    relatedReasons: Array.isArray(entry?.relatedReasons) ? entry.relatedReasons : [],
  };
}

export async function getLinks(conversationId?: string, internetMessageId?: string, itemId?: string): Promise<LinkEntry[]> {
  const params = new URLSearchParams();
  const normalizedConversationId = String(conversationId || "").trim();
  const normalizedInternetMessageId = String(internetMessageId || "")
    .trim()
    .toLowerCase()
    .replace(/[<>\s]/g, "");
  const normalizedItemId = String(itemId || "").trim();
  const lookupKey = normalizedConversationId || normalizedInternetMessageId
    ? `${normalizedConversationId}||${normalizedInternetMessageId}`
    : "";
  if (lookupKey) params.set("conversationId", lookupKey);
  if (normalizedInternetMessageId) params.set("internetMessageId", normalizedInternetMessageId);
  if (normalizedItemId) params.set("itemId", normalizedItemId);
  const r: any = await requestJSON(`/api/links?${params.toString()}`);
  const links: LinkEntry[] = r?.links ?? r ?? [];
  return (Array.isArray(links) ? links : []).map(normalizeLinkEntry);
}

export async function getLinksByRecord(model: string, recordId: number): Promise<LinkEntry[]> {
  const params = new URLSearchParams();
  params.set("model", String(model || "").trim());
  params.set("recordId", String(Number(recordId || 0)));
  const r: any = await requestJSON(`/api/links/by-record?${params.toString()}`);
  const links: LinkEntry[] = r?.links ?? r ?? [];
  return (Array.isArray(links) ? links : []).map(normalizeLinkEntry);
}

export async function linkEmailToRecord(payload: LinkPayload): Promise<{ ok: boolean; link?: LinkEntry }> {
  try {
    return await requestJSON(`/api/links/link`, { method: "POST", body: JSON.stringify(payload) });
  } catch {
    // fallback for older servers
    return await requestJSON(`/api/odoo/link-email`, { method: "POST", body: JSON.stringify(payload) });
  }
}

export async function registerRelevantEmail(payload: RelevantEmailPayload): Promise<RelevantEmailPayload & { id?: string; groups?: LinkGroupEntry[] }> {
  const response: any = await requestJSON(`/api/links/email`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return response?.email ?? response ?? {};
}

export async function getRelatedEmailContext(payload: RelevantEmailPayload): Promise<{
  email: RelatedEmailEntry | null;
  emails: RelatedEmailEntry[];
  groups: LinkGroupEntry[];
  tickets: GroupTicketEntry[];
}> {
  const params = new URLSearchParams();
  if (payload.conversationId) params.set("conversationId", String(payload.conversationId).trim());
  if (payload.internetMessageId) params.set("internetMessageId", String(payload.internetMessageId).trim());
  if (payload.itemId) params.set("itemId", String(payload.itemId).trim());
  if (payload.subject) params.set("subject", String(payload.subject).trim());
  if (payload.fromEmail) params.set("fromEmail", String(payload.fromEmail).trim());
  if (payload.receivedAtIso) params.set("receivedAtIso", String(payload.receivedAtIso).trim());
  const response: any = await requestJSON(`/api/links/related?${params.toString()}`);
  return {
    email: response?.email ? normalizeRelatedEmailEntry(response.email) : null,
    emails: Array.isArray(response?.emails) ? response.emails.map(normalizeRelatedEmailEntry) : [],
    groups: Array.isArray(response?.groups) ? response.groups : [],
    tickets: Array.isArray(response?.tickets) ? response.tickets : [],
  };
}

export async function listLinkGroups(query = ""): Promise<LinkGroupEntry[]> {
  const params = new URLSearchParams();
  if (String(query || "").trim()) params.set("q", String(query || "").trim());
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/groups?${params.toString()}`);
  return Array.isArray(response?.groups) ? response.groups : [];
}

export async function searchKnownEmails(query = "", options?: { excludeGroupId?: string; limit?: number }): Promise<RelatedEmailEntry[]> {
  const params = new URLSearchParams();
  if (String(query || "").trim()) params.set("q", String(query || "").trim());
  if (String(options?.excludeGroupId || "").trim()) params.set("excludeGroupId", String(options?.excludeGroupId || "").trim());
  if (Number(options?.limit || 0) > 0) params.set("limit", String(Number(options?.limit)));
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/emails?${params.toString()}`);
  return Array.isArray(response?.emails) ? response.emails.map(normalizeRelatedEmailEntry) : [];
}

export async function createLinkGroup(payload: {
  name: string;
  description?: string;
  notes?: string;
  contacts?: Array<{ key?: string; name: string; email?: string; company?: string; source?: string }>;
  entities?: Array<{ key?: string; name: string; kind?: string; source?: string }>;
  documentsEnabled?: boolean;
  status?: string;
  labels?: string[];
  isArchived?: boolean;
}): Promise<LinkGroupEntry> {
  const response: any = await requestJSON(`/api/links/groups`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return response?.group ?? response;
}

export async function updateLinkGroup(
  groupId: string,
  payload: {
    name?: string;
    description?: string;
    notes?: string;
    contacts?: Array<{ key?: string; name: string; email?: string; company?: string; source?: string }>;
    entities?: Array<{ key?: string; name: string; kind?: string; source?: string }>;
    documentsEnabled?: boolean;
    status?: string;
    labels?: string[];
    isArchived?: boolean;
    archivedAt?: string;
  }
): Promise<LinkGroupEntry> {
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}`, {
    method: "PATCH",
    body: JSON.stringify(payload),
  });
  return response?.group ?? response;
}

export async function deleteLinkGroup(groupId: string): Promise<{ ok: boolean }> {
  await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}`, {
    method: "DELETE",
  });
  return { ok: true };
}

export async function addEmailToLinkGroup(groupId: string, payload: RelevantEmailPayload): Promise<{ group: LinkGroupEntry; email: RelatedEmailEntry | null }> {
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/emails`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return {
    group: response?.group ?? response,
    email: response?.email ? normalizeRelatedEmailEntry(response.email) : null,
  };
}

export async function removeEmailFromLinkGroup(groupId: string, payload: RelevantEmailPayload & { emailKey?: string }): Promise<{ ok: boolean }> {
  await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/emails`, {
    method: "DELETE",
    body: JSON.stringify(payload),
  });
  return { ok: true };
}

export async function getGroupEmails(groupId: string): Promise<RelatedEmailEntry[]> {
  const params = new URLSearchParams();
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/emails?${params.toString()}`);
  return Array.isArray(response?.emails) ? response.emails.map(normalizeRelatedEmailEntry) : [];
}

export async function getGroupDocuments(groupId: string): Promise<GroupDocumentEntry[]> {
  const params = new URLSearchParams();
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/documents?${params.toString()}`);
  return Array.isArray(response?.documents) ? response.documents : [];
}

export function getGroupDocumentContentUrl(groupId: string, documentId: string, options?: { download?: boolean }): string {
  const normalizedGroupId = encodeURIComponent(String(groupId || "").trim());
  const normalizedDocumentId = encodeURIComponent(String(documentId || "").trim());
  const url = new URL(
    `/api/links/groups/${normalizedGroupId}/documents/${normalizedDocumentId}/content`,
    window.location.origin
  );
  if (options?.download) url.searchParams.set("download", "1");
  return url.toString();
}

export function getEmailAttachmentContentUrl(emailId: string, attachmentKey: string, options?: { download?: boolean }): string {
  const normalizedEmailId = encodeURIComponent(String(emailId || "").trim());
  const normalizedAttachmentKey = encodeURIComponent(String(attachmentKey || "").trim());
  const url = new URL(
    `/api/links/emails/${normalizedEmailId}/attachments/${normalizedAttachmentKey}/content`,
    window.location.origin
  );
  if (options?.download) url.searchParams.set("download", "1");
  return url.toString();
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.subarray(offset, Math.min(bytes.length, offset + chunkSize));
    binary += String.fromCharCode(...chunk);
  }
  return globalThis.btoa(binary);
}

export async function getGroupDocumentContentBase64(
  groupId: string,
  documentId: string
): Promise<{ base64: string; contentType: string; fileName: string }> {
  const response = await fetch(getGroupDocumentContentUrl(groupId, documentId), {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Nao foi possivel carregar o documento (${response.status}).`);
  }
  const fileName = String(
    response.headers.get("x-file-name")
    || response.headers.get("content-disposition")
    || ""
  ).trim();
  const contentType = String(response.headers.get("content-type") || "application/octet-stream").trim();
  const bytes = new Uint8Array(await response.arrayBuffer());
  return {
    base64: uint8ArrayToBase64(bytes),
    contentType,
    fileName,
  };
}

export async function getGroupDocumentTextContent(groupId: string, documentId: string): Promise<string> {
  const response = await fetch(getGroupDocumentContentUrl(groupId, documentId), {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Nao foi possivel ler o documento (${response.status}).`);
  }
  return await response.text();
}

export async function getEmailAttachmentContentBase64(
  emailId: string,
  attachmentKey: string
): Promise<{ base64: string; contentType: string; fileName: string }> {
  const response = await fetch(getEmailAttachmentContentUrl(emailId, attachmentKey), {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Nao foi possivel carregar o anexo (${response.status}).`);
  }
  const fileName = String(
    response.headers.get("x-file-name")
    || response.headers.get("content-disposition")
    || ""
  ).trim();
  const contentType = String(response.headers.get("content-type") || "application/octet-stream").trim();
  const bytes = new Uint8Array(await response.arrayBuffer());
  return {
    base64: uint8ArrayToBase64(bytes),
    contentType,
    fileName,
  };
}

export async function getEmailAttachmentTextContent(emailId: string, attachmentKey: string): Promise<string> {
  const response = await fetch(getEmailAttachmentContentUrl(emailId, attachmentKey), {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error(`Nao foi possivel ler o anexo (${response.status}).`);
  }
  return await response.text();
}

export async function getGroupAttachmentFlags(groupId: string): Promise<GroupAttachmentFlagEntry[]> {
  const params = new URLSearchParams();
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/attachment-flags?${params.toString()}`);
  return Array.isArray(response?.flags) ? response.flags : [];
}

export async function saveGroupAttachmentFlags(
  groupId: string,
  payload: { entries: GroupAttachmentFlagEntry[] }
): Promise<{ ok: boolean; flags: GroupAttachmentFlagEntry[] }> {
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/attachment-flags`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return {
    ok: Boolean(response?.ok),
    flags: Array.isArray(response?.flags) ? response.flags : [],
  };
}

export async function saveGroupDocuments(
  groupId: string,
  payload: {
    documents: GroupDocumentEntry[];
  }
): Promise<{ ok: boolean; group?: LinkGroupEntry; documents: GroupDocumentEntry[] }> {
  const response: any = await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/documents`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return {
    ok: Boolean(response?.ok),
    group: response?.group,
    documents: Array.isArray(response?.documents) ? response.documents : [],
  };
}

export async function deleteGroupDocument(groupId: string, documentId: string): Promise<{ ok: boolean }> {
  await requestJSON(`/api/links/groups/${encodeURIComponent(String(groupId || "").trim())}/documents/${encodeURIComponent(String(documentId || "").trim())}`, {
    method: "DELETE",
  });
  return { ok: true };
}

export async function listGroupTicketSeries(): Promise<GroupTicketSeriesEntry[]> {
  const params = new URLSearchParams();
  params.set("_ts", String(Date.now()));
  const response: any = await requestJSON(`/api/links/group-ticket-series?${params.toString()}`);
  return Array.isArray(response?.series) ? response.series : [];
}

export async function createGroupTicketSeries(payload: {
  name: string;
  prefix: string;
  replyInstructions?: string;
  yearMode?: "none" | "yy" | "yyyy";
  separator?: "-" | "/" | "_" | " " | "";
  nextNumber?: number;
  padding?: number;
  isActive?: boolean;
}): Promise<GroupTicketSeriesEntry> {
  const response: any = await requestJSON(`/api/links/group-ticket-series`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return response?.series ?? response;
}

export async function updateGroupTicketSeries(
  seriesId: string,
  payload: {
    name?: string;
    prefix?: string;
    replyInstructions?: string;
    yearMode?: "none" | "yy" | "yyyy";
    separator?: "-" | "/" | "_" | " " | "";
    nextNumber?: number;
    padding?: number;
    isActive?: boolean;
  }
): Promise<GroupTicketSeriesEntry> {
  const response: any = await requestJSON(`/api/links/group-ticket-series/${encodeURIComponent(String(seriesId || "").trim())}`, {
    method: "PATCH",
    body: JSON.stringify(payload),
  });
  return response?.series ?? response;
}

export async function deleteGroupTicketSeries(seriesId: string): Promise<{ ok: boolean }> {
  await requestJSON(`/api/links/group-ticket-series/${encodeURIComponent(String(seriesId || "").trim())}`, {
    method: "DELETE",
  });
  return { ok: true };
}

export async function searchGroupTickets(payload: {
  q?: string;
  groupId?: string;
  email?: RelevantEmailPayload;
  limit?: number;
}): Promise<GroupTicketEntry[]> {
  const response: any = await requestJSON(`/api/links/group-tickets/search`, {
    method: "POST",
    body: JSON.stringify(payload || {}),
  });
  return Array.isArray(response?.tickets) ? response.tickets : [];
}

export async function createGroupTicket(payload: {
  seriesId: string;
  title: string;
  description?: string;
  labels?: string[];
  groupIds?: string[];
  email?: RelevantEmailPayload;
  membershipKind?: "principal" | "referencia" | string;
}): Promise<GroupTicketEntry> {
  const response: any = await requestJSON(`/api/links/group-tickets`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return response?.ticket ?? response;
}

export async function updateGroupTicket(
  ticketId: string,
  payload: { title?: string; description?: string; labels?: string[]; groupIds?: string[]; status?: string }
): Promise<GroupTicketEntry> {
  const response: any = await requestJSON(`/api/links/group-tickets/${encodeURIComponent(String(ticketId || "").trim())}`, {
    method: "PATCH",
    body: JSON.stringify(payload),
  });
  return response?.ticket ?? response;
}

export async function linkEmailToGroupTicket(
  ticketId: string,
  payload: {
    email: RelevantEmailPayload;
    applyGroups?: boolean;
    groupIds?: string[];
    membershipKind?: "principal" | "referencia" | string;
  }
): Promise<{ ok: boolean; ticket: GroupTicketEntry; appliedGroups: LinkGroupEntry[]; email?: RelatedEmailEntry | null }> {
  const response: any = await requestJSON(`/api/links/group-tickets/${encodeURIComponent(String(ticketId || "").trim())}/email`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return {
    ok: Boolean(response?.ok),
    ticket: response?.ticket,
    appliedGroups: Array.isArray(response?.appliedGroups) ? response.appliedGroups : [],
    email: response?.email ? normalizeRelatedEmailEntry(response.email) : null,
  };
}

export async function unlinkEmailFromGroupTicket(
  ticketId: string,
  payload: {
    email?: RelevantEmailPayload;
    emailKey?: string;
  }
): Promise<{ ok: boolean; removed: boolean; ticket?: GroupTicketEntry | null; emailKey?: string }> {
  const response: any = await requestJSON(`/api/links/group-tickets/${encodeURIComponent(String(ticketId || "").trim())}/email`, {
    method: "DELETE",
    body: JSON.stringify(payload),
  });
  return {
    ok: Boolean(response?.ok),
    removed: Boolean(response?.removed),
    ticket: response?.ticket ?? null,
    emailKey: String(response?.emailKey || "").trim() || undefined,
  };
}

export async function extractAttachmentTexts(
  files: Array<{ key: string; name: string; contentType?: string; content?: string }>
): Promise<AttachmentTextExtractionEntry[]> {
  const response: any = await requestJSON("/api/links/attachments/extract-text", {
    method: "POST",
    body: JSON.stringify({ files }),
  });
  return Array.isArray(response?.results)
    ? response.results.map((entry: any) => ({
        key: String(entry?.key || "").trim(),
        name: String(entry?.name || "").trim(),
        contentType: String(entry?.contentType || "").trim() || undefined,
        text: String(entry?.text || ""),
      })).filter((entry: AttachmentTextExtractionEntry) => entry.key)
    : [];
}

export async function detectGroupTicketsForEmail(payload: {
  email: RelevantEmailPayload;
}): Promise<GroupTicketDetectionMatch[]> {
  const response: any = await requestJSON(`/api/links/group-tickets/detect`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return Array.isArray(response?.matches) ? response.matches : [];
}


export type PartnerLite = {
  id: number;
  name?: string;
  email?: string;
  company_type?: "person" | "company" | string;
  parent_id?: [number, string] | number | null;
  function?: string;
  phone?: string;
  mobile?: string;
};

export type PartnerRelationItem = {
  model: string;
  recordId: number;
  title: string;
  meta?: string;
  secondary?: string;
};

export type PartnerRelationSection = {
  key: string;
  label: string;
  model: string;
  total: number;
  items: PartnerRelationItem[];
};

export type PartnerRelationsResponse = {
  ok: boolean;
  partner?: PartnerLite | null;
  total?: number;
  relations?: PartnerRelationSection[];
};

export async function getPartnerByEmail(email: string): Promise<PartnerLite | null> {
  const q = encodeURIComponent(String(email || "").trim());
  const r: any = await requestJSON(`/api/odoo/partners/by-email?email=${q}`);
  return r?.partner ?? null;
}

export async function getPartnerRelations(partnerId: number): Promise<{ partner: PartnerLite | null; total: number; relations: PartnerRelationSection[] }> {
  const id = Number(partnerId || 0);
  if (!id) return { partner: null, total: 0, relations: [] };
  const r: PartnerRelationsResponse = await requestJSON(`/api/odoo/partners/${encodeURIComponent(String(id))}/relations`);
  return {
    partner: r?.partner ?? null,
    total: Number(r?.total || 0),
    relations: Array.isArray(r?.relations) ? r.relations : [],
  };
}

export async function createOrUpdatePartner(payload: {
  mode: "create" | "update";
  targetPartnerId?: number;
  data: {
    name?: string;
    email?: string;
    company_type?: "person" | "company";
    parent_id?: number | null;
    function?: string;
    phone?: string;
    mobile?: string;
  };
}): Promise<any> {
  return await requestJSON(`/api/odoo/partners/create-or-update`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export async function searchCompanies(q: string): Promise<any[]> {
  const query = encodeURIComponent(String(q || "").trim());
  const r: any = await requestJSON(`/api/odoo/companies/search?q=${query}`);
  if (r?.warning === "odoo_unavailable") {
    throw new Error(String(r?.message || "Odoo indisponivel"));
  }
  // Merge-conflict resolution: keep compatibility with both payload shapes
  // and always return a safe array.
  const results = r?.results ?? r?.companies ?? [];
  return Array.isArray(results) ? results : [];
}

// -------- Odoo generic helpers --------
export async function readOdoo(model: string, ids: number[] | number, fields: string[]): Promise<any[]> {
  const idList = Array.isArray(ids) ? ids : [ids];
  if (isMock()) {
    console.log(`[Mock] Reading ${model}`, idList);
    const mock = getMockData("read", model);
    return idList.map(id => mock?.[id] || { id, name: `Mock ${model} ${id}` });
  }
  try {
    const r: any = await requestJSON(`/api/odoo/read`, {
      method: "POST",
      body: JSON.stringify({ model, ids: idList, fields }),
    });
    return r?.records ?? r?.result ?? r ?? [];
  } catch {
    // fallback to domain search (if read endpoint absent)
    const r2: any = await requestJSON(`/api/odoo/search-domain`, {
      method: "POST",
      body: JSON.stringify({ model, domain: [["id", "in", idList]], fields, limit: idList.length }),
    });
    return r2?.records ?? r2?.result ?? r2 ?? [];
  }
}

// searchOdoo: supports both old (model, query, limit) and new (args object)
export async function searchOdoo(
  modelOrArgs:
    | string
    | { model: string; domain: any[]; fields?: string[]; limit?: number; order?: string },
  query?: string,
  limit?: number
): Promise<any[]> {
  if (typeof modelOrArgs === "string") {
    const model = modelOrArgs;
    const q = (query ?? "").trim();
    const lim = limit ?? 20;

    // if server implements free-text search, use it; else fallback to name ilike
    try {
      const r: any = await requestJSON(`/api/odoo/search`, {
        method: "POST",
        body: JSON.stringify({ model, query: q, limit: lim }),
      });
      return r?.records ?? r?.result ?? r ?? [];
    } catch {
      const domain = q ? [["name", "ilike", q]] : [];
      const r2: any = await requestJSON(`/api/odoo/search-domain`, {
        method: "POST",
        body: JSON.stringify({ model, domain, fields: ["id", "name"], limit: lim }),
      });
      return r2?.records ?? r2?.result ?? r2 ?? [];
    }
  }

  // args object
  try {
    const r: any = await requestJSON(`/api/odoo/search`, { method: "POST", body: JSON.stringify(modelOrArgs) });
    return r?.records ?? r?.result ?? r ?? [];
  } catch {
    const r2: any = await requestJSON(`/api/odoo/search-domain`, { method: "POST", body: JSON.stringify(modelOrArgs) });
    return r2?.records ?? r2?.result ?? r2 ?? [];
  }
}

// searchOdooDomain: supports both signatures
export async function searchOdooDomain(args: { model: string; domain: any[]; fields?: string[]; limit?: number; order?: string }): Promise<any[]>;
export async function searchOdooDomain(model: string, domain: any[], fields?: string[], limit?: number): Promise<any[]>;
export async function searchOdooDomain(
  a: any,
  b?: any,
  c?: any,
  d?: any
): Promise<any[]> {
  const payload =
    typeof a === "string"
      ? { model: a, domain: b ?? [], fields: c, limit: d }
      : a;

  const r: any = isMock() ? { records: getMockData("search_domain", payload.model) || [] } : await requestJSON(`/api/odoo/search-domain`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return r?.records ?? r?.result ?? r ?? [];
}

export async function callOdoo(payload: { model: string; method: string; args: any[]; kwargs?: Record<string, any> }): Promise<any> {
  const response: any = await requestJSON(`/api/odoo/call`, { method: "POST", body: JSON.stringify(payload) });
  return response?.result ?? response;
}

export type OdooFieldMeta = {
  name: string;
  string?: string;
  type?: string;
  relation?: string;
  selection?: Array<[string, string]>;
};

function normalizeFieldLabel(value: string): string {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

export async function findOdooFieldByLabel(model: string, label: string): Promise<OdooFieldMeta | null> {
  const target = normalizeFieldLabel(label);
  if (!target) return null;

  const result: Record<string, any> = await callOdoo({
    model,
    method: "fields_get",
    args: [],
    kwargs: { attributes: ["string", "type", "relation", "selection"] },
  });

  for (const [name, meta] of Object.entries(result || {})) {
    if (normalizeFieldLabel(String(meta?.string || "")) !== target) continue;
    return {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };
  }

  return null;
}

export async function getOdooFieldMeta(model: string, fieldName: string): Promise<OdooFieldMeta | null> {
  const normalizedFieldName = String(fieldName || "").trim();
  if (!normalizedFieldName) return null;

  const result: Record<string, any> = await callOdoo({
    model,
    method: "fields_get",
    args: [[normalizedFieldName]],
    kwargs: { attributes: ["string", "type", "relation", "selection"] },
  });

  const meta = result?.[normalizedFieldName];
  if (!meta) return null;

  return {
    name: normalizedFieldName,
    string: meta?.string,
    type: meta?.type,
    relation: meta?.relation,
    selection: Array.isArray(meta?.selection) ? meta.selection : [],
  };
}

export async function getLeadTypeFieldMeta(): Promise<OdooFieldMeta | null> {
  const meta = await getOdooFieldMeta("crm.lead", "x_studio_tipo_de_lead");
  if (!meta || meta.type !== "selection") return null;
  return {
    ...meta,
    selection: Array.isArray(meta.selection) ? meta.selection : [],
  };
}

export async function findOdooField(
  model: string,
  options: {
    labels?: string[];
    nameCandidates?: string[];
    namePatterns?: RegExp[];
    preferredTypes?: string[];
  }
): Promise<OdooFieldMeta | null> {
  const result: Record<string, any> = await callOdoo({
    model,
    method: "fields_get",
    args: [],
    kwargs: { attributes: ["string", "type", "relation", "selection"] },
  });

  const labelTargets = (options.labels || []).map(normalizeFieldLabel).filter(Boolean);
  const candidateNames = new Set((options.nameCandidates || []).map((name) => String(name || "").trim()).filter(Boolean));
  const normalizedCandidateNames = new Set(
    (options.nameCandidates || []).map((name) => normalizeFieldLabel(String(name || ""))).filter(Boolean)
  );
  const preferredTypes = new Set(options.preferredTypes || []);

  for (const [name, meta] of Object.entries(result || {})) {
    const normalizedName = normalizeFieldLabel(name);
    if (!candidateNames.has(name) && !normalizedCandidateNames.has(normalizedName)) continue;
    return {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };
  }

  let best: { score: number; field: OdooFieldMeta } | null = null;

  for (const [name, meta] of Object.entries(result || {})) {
    const normalizedLabel = normalizeFieldLabel(String(meta?.string || ""));
    const normalizedName = normalizeFieldLabel(name);
    let score = -1;

    if (candidateNames.has(name) || normalizedCandidateNames.has(normalizedName)) score = 110;
    else if (labelTargets.includes(normalizedLabel)) score = 100;
    else if (labelTargets.some((target) => normalizedLabel.includes(target) || target.includes(normalizedLabel))) score = 80;
    else if (labelTargets.some((target) => target.split(/\s+/).every((token) => normalizedLabel.includes(token)))) score = 70;
    else if ((options.namePatterns || []).some((pattern) => pattern.test(name))) score = 60;

    if (score < 0) continue;
    if (preferredTypes.has(String(meta?.type || ""))) score += 5;
    if (labelTargets.some((target) => target.split(/\s+/).every((token) => normalizedName.includes(token)))) score += 2;
    if (String(meta?.type || "") === "selection" && Array.isArray(meta?.selection) && meta.selection.length) score += 20;
    if (String(meta?.type || "") === "many2one" && meta?.relation) score += 10;

    const field: OdooFieldMeta = {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };

    if (!best || score > best.score) {
      best = { score, field };
    }
  }

  return best?.field || null;
}

// createOdoo: return number (DialogApp expects number)
export async function createOdoo(model: string, values: Record<string, any>): Promise<number> {
  // prefer dedicated endpoint
  if (isMock()) {
    const id = Math.floor(Math.random() * 10000);
    console.log(`[Mock] Created ${model} with ID ${id}`, values);
    return id;
  }
  try {
    const r: any = await requestJSON(`/api/odoo/create`, { method: "POST", body: JSON.stringify({ model, values }) });
    const id = r?.id ?? r?.result ?? r;
    return Number(id);
  } catch {
    const r2: any = await callOdoo({ model, method: "create", args: [values] });
    const id = r2?.id ?? r2?.result ?? r2;
    return Number(id);
  }
}

export async function writeOdoo(model: string, ids: number[] | number, values: Record<string, any>): Promise<boolean> {
  const idList = Array.isArray(ids) ? ids : [ids];
  // try direct endpoint if it exists
  try {
    const r: any = await requestJSON(`/api/odoo/write`, { method: "POST", body: JSON.stringify({ model, id: idList[0], ids: idList, values }) });
    return Boolean(r?.ok ?? r?.result ?? r ?? true);
  } catch {
    const r2: any = await callOdoo({ model, method: "write", args: [idList, values] });
    return Boolean(r2?.ok ?? r2?.result ?? r2 ?? true);
  }
}

// -------- AI --------
export async function aiSelftest(customModels?: any): Promise<{ ok: boolean; openai: { ok: boolean; error?: string }; gemini: { ok: boolean; error?: string }; error?: string }> {
  return await requestJSON(`/api/ai/selftest`, { method: "POST", body: JSON.stringify({ customModels }) });
}

export async function aiGenerate(payload: any, customModels?: any): Promise<AiGenerateResponse> {
  const mergedPayload = { ...payload, customModels: customModels ?? payload.customModels };
  return await requestJSON(`/api/ai/generate`, { method: "POST", body: JSON.stringify(mergedPayload) });
}

export async function aiExtractAnchors(emailBody: string, customModels?: any, emailContext?: any): Promise<{ ok: boolean; anchors: any }> {
  return await requestJSON(`/api/ai/extract-anchors`, { method: "POST", body: JSON.stringify({ emailBody, emailContext, customModels }) });
}

export async function aiGenerateBriefing(
  context: string,
  history: any[] = [],
  customModels?: any,
  conversationId?: string,
  cacheKey?: string,
): Promise<{ ok: boolean; summary: string }> {
  return await requestJSON(`/api/ai/briefing`, {
    method: "POST",
    body: JSON.stringify({ context, history, customModels, conversationId, cacheKey }),
  });
}

export async function aiVoiceCommand(commandText: string, context: any): Promise<{ ok: boolean; actions: string[] }> {
  return await requestJSON(`/api/ai/voice-command`, { method: "POST", body: JSON.stringify({ commandText, context }) });
}

export async function aiListModels(): Promise<{ ok: boolean; openai: string[]; gemini: string[] }> {
  return await requestJSON(`/api/ai/list-models`);
}

export type InvoiceStudioUploadFile = {
  name: string;
  type?: string;
  content: string;
};

export type InvoiceStudioUploadPayload = {
  baseUrl: string;
  email: string;
  password: string;
  project?: string;
  batchId?: string;
  metadata?: Record<string, any>;
  files: InvoiceStudioUploadFile[];
};

export type InvoiceStudioUploadResult = {
  ok: boolean;
  batchId: string;
  count: number;
  project?: string;
  status?: string;
  upload?: any;
};

export type InvoiceStudioBatchStatusResult = {
  ok: boolean;
  batchId: string;
  project?: string;
  progress?: {
    total?: number;
    done?: number;
    errors?: number;
    status?: string;
  };
  rows?: Array<Record<string, any>>;
};

export async function uploadToInvoiceStudio(payload: InvoiceStudioUploadPayload): Promise<InvoiceStudioUploadResult> {
  return await requestJSON(`/api/invoice-studio/upload`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export async function getInvoiceStudioBatchStatus(payload: {
  baseUrl: string;
  email: string;
  password: string;
  project?: string;
  batchId: string;
}): Promise<InvoiceStudioBatchStatusResult> {
  return await requestJSON(`/api/invoice-studio/status`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

// -------- Learning --------
export async function logLearningInteraction(log: any): Promise<{ ok: boolean }> {
  return await requestJSON(`/api/learning/log`, {
    method: "POST",
    body: JSON.stringify(log),
  });
}

export async function getLearningProfile(userId: string = "global"): Promise<{ ok: boolean; profile: any }> {
  return await requestJSON(`/api/learning/profile?userId=${encodeURIComponent(userId)}`);
}
