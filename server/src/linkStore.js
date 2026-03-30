import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import { createOptionalPgStore } from "./optionalPg.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

const PRIMARY_DATA_DIR = path.resolve(__dirname, "../data");
const PRIMARY_FILE_PATH = path.join(PRIMARY_DATA_DIR, "links.json");
const LEGACY_FILE_PATH = path.join(process.cwd(), "server", "data", "links.json");
const STORE_VERSION = 2;
const CUSTOM_GROUP_KIND = "custom";
const DEFAULT_GROUP_STATUS = "em_analise";
const DEFAULT_GROUP_MEMBERSHIP_KIND = "principal";
const DEFAULT_GROUP_TICKET_STATUS = "open";
const DEFAULT_GROUP_TICKET_YEAR_MODE = "none";
const DEFAULT_GROUP_TICKET_SEPARATOR = "-";

const db = createOptionalPgStore("linkStore");
let customGroupDbInitPromise = null;

function hasDurablePersistence() {
  return typeof db.isConfigured === "function" && db.isConfigured();
}

function durablePersistenceError(action, error) {
  const detail = normalizeString(error?.message);
  const suffix = detail ? ` Detalhe tecnico: ${detail}` : "";
  return new Error(
    `Nao foi possivel ${action} porque a persistencia segura em PostgreSQL esta indisponivel. Nenhuma alteracao foi guardada.${suffix}`
  );
}

async function requireDurablePersistence(action, options = {}) {
  if (!hasDurablePersistence()) return false;
  if (!db.isEnabled()) {
    throw durablePersistenceError(action);
  }
  try {
    await ensureCustomGroupDb();
    if (typeof options.syncStore === "function") {
      await options.syncStore();
    }
    return true;
  } catch (error) {
    throw durablePersistenceError(action, error);
  }
}

function writeCacheStore(store) {
  try {
    writeStore(store);
  } catch (error) {
    if (hasDurablePersistence()) {
      console.warn("[linkStore] Local cache write failed after durable persistence success:", error?.message || error);
      return;
    }
    throw error;
  }
}

function nowIso() {
  return new Date().toISOString();
}

function normalizeString(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "" : value.toISOString();
  }
  return String(value || "").trim();
}

function normalizeMessageId(value) {
  return normalizeString(value)
    .toLowerCase()
    .replace(/[<>\s]/g, "");
}

function normalizeModel(value) {
  return normalizeString(value);
}

function normalizeRecordId(value) {
  return Number(value || 0);
}

function normalizeGroupStatus(value) {
  const normalized = normalizeString(value).toLowerCase().replace(/\s+/g, "_");
  if (normalized === "concluido" || normalized === "concluido." || normalized === "completed" || normalized === "done") {
    return "concluido";
  }
  if (normalized === "em_progresso" || normalized === "progresso" || normalized === "in_progress" || normalized === "progress") {
    return "em_progresso";
  }
  return DEFAULT_GROUP_STATUS;
}

function normalizeGroupLabels(value) {
  const rawItems = Array.isArray(value)
    ? value
    : String(value || "")
      .split(/[,\n;]/g)
      .map((entry) => entry.trim());
  const seen = new Set();
  const labels = [];
  for (const item of rawItems) {
    const label = normalizeString(item);
    if (!label) continue;
    const key = label.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    labels.push(label);
  }
  return labels.sort((a, b) => a.localeCompare(b, "pt"));
}

function parseGroupLabelsJson(value) {
  if (Array.isArray(value)) return normalizeGroupLabels(value);
  const raw = normalizeString(value);
  if (!raw) return [];
  try {
    return normalizeGroupLabels(JSON.parse(raw));
  } catch {
    return normalizeGroupLabels(raw);
  }
}

function normalizeGroupMembershipKind(value) {
  const normalized = normalizeString(value).toLowerCase().replace(/\s+/g, "_");
  if (normalized === "referencia" || normalized === "reference" || normalized === "linked" || normalized === "link") {
    return "referencia";
  }
  return DEFAULT_GROUP_MEMBERSHIP_KIND;
}

function normalizeGroupMembershipMeta(value = {}) {
  if (typeof value === "string") {
    return {
      kind: normalizeGroupMembershipKind(value),
      linkedAt: "",
      updatedAt: "",
    };
  }
  return {
    kind: normalizeGroupMembershipKind(value?.kind),
    linkedAt: normalizeString(value?.linkedAt),
    updatedAt: normalizeString(value?.updatedAt),
  };
}

function normalizePositiveInt(value, fallback = 1, min = 1, max = 999999) {
  const parsed = Number(value || 0);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.max(min, Math.min(max, Math.trunc(parsed)));
}

function normalizeTicketStatus(value) {
  const normalized = normalizeString(value).toLowerCase().replace(/\s+/g, "_");
  return normalized === "closed" || normalized === "fechado" || normalized === "concluido"
    ? "closed"
    : DEFAULT_GROUP_TICKET_STATUS;
}

function normalizeTicketPrefix(value) {
  return normalizeString(value)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "");
}

function normalizeTicketYearMode(value) {
  const normalized = normalizeString(value).toLowerCase();
  if (normalized === "yy" || normalized === "2" || normalized === "short") return "yy";
  if (normalized === "yyyy" || normalized === "4" || normalized === "long") return "yyyy";
  return DEFAULT_GROUP_TICKET_YEAR_MODE;
}

function normalizeTicketSeparator(value) {
  const normalized = String(value ?? DEFAULT_GROUP_TICKET_SEPARATOR);
  return normalized === "-" || normalized === "/" || normalized === "_" || normalized === " " || normalized === ""
    ? normalized
    : DEFAULT_GROUP_TICKET_SEPARATOR;
}

function normalizeGroupIds(value) {
  const items = Array.isArray(value) ? value : [value];
  return Array.from(
    new Set(
      items
        .map((entry) => normalizeString(entry))
        .filter(Boolean)
    )
  );
}

function stripHtmlForTicketLookup(html) {
  return String(html || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function getTicketYearValue(value, yearMode = DEFAULT_GROUP_TICKET_YEAR_MODE) {
  const normalizedMode = normalizeTicketYearMode(yearMode);
  if (normalizedMode === "none") return "";
  const date = value ? new Date(value) : new Date();
  if (Number.isNaN(date.getTime())) return "";
  const fullYear = String(date.getUTCFullYear());
  return normalizedMode === "yy" ? fullYear.slice(-2) : fullYear;
}

function buildTicketCode(prefix, sequenceNumber, padding = 4, options = {}) {
  const safePrefix = normalizeTicketPrefix(prefix);
  const safeNumber = normalizePositiveInt(sequenceNumber, 1);
  const safePadding = normalizePositiveInt(padding, 4, 2, 8);
  const yearMode = normalizeTicketYearMode(options?.yearMode);
  const separator = normalizeTicketSeparator(options?.separator);
  const yearValue = normalizeString(options?.yearValue) || getTicketYearValue(options?.dateValue, yearMode);
  if (!safePrefix) return "";
  const parts = [safePrefix];
  if (yearMode !== "none" && yearValue) parts.push(yearValue);
  parts.push(String(safeNumber).padStart(safePadding, "0"));
  return parts.join(separator);
}

function splitLookupKey(conversationId, internetMessageId = "") {
  const rawConversationId = normalizeString(conversationId);
  if (rawConversationId.includes("||")) {
    const [cid, imid] = rawConversationId.split("||");
    return {
      conversationId: normalizeString(cid),
      internetMessageId: normalizeMessageId(internetMessageId || imid),
    };
  }
  return {
    conversationId: rawConversationId,
    internetMessageId: normalizeMessageId(internetMessageId),
  };
}

function makeEntityKey(model, recordId) {
  const normalizedModel = normalizeModel(model);
  const normalizedRecordId = normalizeRecordId(recordId);
  return normalizedModel && normalizedRecordId ? `${normalizedModel}:${normalizedRecordId}` : "";
}

function makeEmailFingerprint(entry) {
  const subject = normalizeString(entry?.subject).toLowerCase();
  const fromEmail = normalizeString(entry?.fromEmail).toLowerCase();
  const messageDateIso = normalizeString(
    entry?.messageDateIso || entry?.receivedAtIso || entry?.sentAtIso || entry?.linkedAt
  );
  const conversationId = normalizeString(entry?.conversationId);
  const parts = [conversationId, subject, fromEmail, messageDateIso].filter(Boolean);
  return parts.length >= 3 ? parts.join("|") : "";
}

function createEmptyStore() {
  return {
    version: STORE_VERSION,
    emails: {},
    indexes: {
      itemIds: {},
      internetMessageIds: {},
      fingerprints: {},
      conversations: {},
    },
    groups: {},
    groupMembers: {},
    groupMemberLinks: {},
    emailGroups: {},
    conversationGroups: {},
    entityLinks: {},
    emailEntityLinks: {},
    groupDocuments: {},
    groupAttachmentFlags: {},
    groupTicketSeries: {},
    groupTickets: {},
    groupTicketEmails: {},
    groupEmailTickets: {},
  };
}

function uniqueFilePaths() {
  return Array.from(new Set([PRIMARY_FILE_PATH, LEGACY_FILE_PATH].filter(Boolean)));
}

function ensurePrimaryDir() {
  if (!fs.existsSync(PRIMARY_DATA_DIR)) fs.mkdirSync(PRIMARY_DATA_DIR, { recursive: true });
}

function readRawFile(filePath) {
  try {
    if (!fs.existsSync(filePath)) return null;
    const raw = fs.readFileSync(filePath, "utf-8");
    return JSON.parse(raw || "{}");
  } catch {
    return null;
  }
}

function writeStore(store) {
  ensurePrimaryDir();
  fs.writeFileSync(PRIMARY_FILE_PATH, JSON.stringify(store, null, 2), "utf-8");
}

function mergeEntryData(current, incoming) {
  const next = { ...current };
  for (const [key, value] of Object.entries(incoming || {})) {
    if (value === undefined || value === null || value === "") continue;
    if (Array.isArray(value)) {
      if (!Array.isArray(next[key]) || next[key].length === 0) {
        next[key] = value;
      }
      continue;
    }
    if (!next[key]) next[key] = value;
  }
  if (String(incoming?.linkedAt || "") > String(current?.linkedAt || "")) {
    next.linkedAt = incoming.linkedAt;
  }
  if (String(incoming?.messageDateIso || incoming?.receivedAtIso || "") > String(current?.messageDateIso || current?.receivedAtIso || "")) {
    next.messageDateIso = incoming.messageDateIso || incoming.receivedAtIso;
    next.receivedAtIso = incoming.receivedAtIso || incoming.messageDateIso;
  }
  return next;
}

function dedupeRecordLinks(entries) {
  const seen = new Map();
  for (const entry of entries || []) {
    const key = makeEntityKey(entry?.model, entry?.recordId ?? entry?.resId);
    if (!key) continue;
    const current = seen.get(key);
    seen.set(key, current ? mergeEntryData(current, entry) : entry);
  }
  return Array.from(seen.values()).sort((a, b) => String(b.linkedAt || "").localeCompare(String(a.linkedAt || "")));
}

function makeEmailLookupKey(entry) {
  return normalizeString(entry?.itemId)
    || normalizeMessageId(entry?.internetMessageId)
    || makeEmailFingerprint(entry)
    || normalizeString(entry?.conversationId)
    || [
      normalizeString(entry?.subject).toLowerCase(),
      normalizeString(entry?.fromEmail).toLowerCase(),
      normalizeString(entry?.messageDateIso || entry?.receivedAtIso || entry?.linkedAt),
    ].join("|");
}

function dedupeEmailLinks(entries) {
  const seen = new Map();
  for (const entry of entries || []) {
    const key = makeEmailLookupKey(entry);
    if (!key) continue;
    const current = seen.get(key);
    seen.set(key, current ? mergeEntryData(current, entry) : entry);
  }
  return Array.from(seen.values()).sort((a, b) =>
    String(b.messageDateIso || b.receivedAtIso || b.linkedAt || "").localeCompare(
      String(a.messageDateIso || a.receivedAtIso || a.linkedAt || "")
    )
  );
}

function parseLegacyStorageKey(key) {
  const raw = normalizeString(key);
  if (raw.startsWith("conversationId:")) {
    return { conversationId: raw.slice("conversationId:".length).trim(), internetMessageId: "" };
  }
  if (raw.startsWith("internetMessageId:")) {
    return { conversationId: "", internetMessageId: normalizeMessageId(raw.slice("internetMessageId:".length)) };
  }
  return { conversationId: "", internetMessageId: "" };
}

function normalizeEmailInput(input) {
  const attachments = normalizeAttachments(input?.attachments);
  return {
    itemId: normalizeString(input?.itemId),
    internetMessageId: normalizeMessageId(input?.internetMessageId),
    conversationId: normalizeString(input?.conversationId),
    subject: normalizeString(input?.subject),
    fromEmail: normalizeString(input?.fromEmail),
    fromName: normalizeString(input?.fromName),
    emailWebLink: normalizeString(input?.emailWebLink || input?.url),
    receivedAtIso: normalizeString(input?.receivedAtIso),
    sentAtIso: normalizeString(input?.sentAtIso),
    messageDateIso: normalizeString(input?.messageDateIso || input?.receivedAtIso || input?.sentAtIso || input?.linkedAt),
    linkedAt: normalizeString(input?.linkedAt),
    bodyText: normalizeString(input?.bodyText),
    bodyHtml: normalizeString(input?.bodyHtml),
    status: Object.prototype.hasOwnProperty.call(input || {}, "status")
      ? (normalizeString(input?.status) ? normalizeGroupStatus(input?.status) : "")
      : undefined,
    labels: Object.prototype.hasOwnProperty.call(input || {}, "labels")
      ? normalizeGroupLabels(input?.labels)
      : undefined,
    labelStates: Object.prototype.hasOwnProperty.call(input || {}, "labelStates")
      ? normalizeEmailLabelStates(input?.labelStates)
      : undefined,
    ...(attachments.length ? { attachments } : {}),
  };
}

function normalizeEmailLabelStates(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return {};
  const next = {};
  for (const [label, status] of Object.entries(value || {})) {
    const normalizedLabel = normalizeString(label);
    const rawStatus = normalizeString(status);
    if (!normalizedLabel || !rawStatus) continue;
    next[normalizedLabel] = normalizeGroupStatus(rawStatus);
  }
  return next;
}

function normalizeAttachments(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((attachment) => ({
      id: normalizeString(attachment?.id),
      name: normalizeString(attachment?.name),
      contentType: normalizeString(attachment?.contentType),
      size: Number(attachment?.size || 0) || undefined,
      isInline: Boolean(attachment?.isInline),
      contentId: normalizeString(attachment?.contentId),
      content: normalizeBase64Content(attachment?.content),
    }))
    .filter((attachment) => attachment.name);
}

function normalizeBase64Content(value) {
  return normalizeString(value).replace(/^data:[^,]+,/, "");
}

function estimateBase64Size(base64) {
  const raw = normalizeBase64Content(base64);
  if (!raw) return 0;
  const padding = raw.endsWith("==") ? 2 : raw.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((raw.length * 3) / 4) - padding);
}

function normalizeDocumentName(value) {
  return normalizeString(value).replace(/[\\/:*?"<>|]+/g, "_");
}

function normalizeGroupDocumentInput(input = {}) {
  const contentBase64 = normalizeBase64Content(input?.contentBase64 || input?.content);
  return {
    id: normalizeString(input?.id) || `doc_${crypto.randomUUID()}`,
    name: normalizeDocumentName(input?.name) || "documento",
    contentType: normalizeString(input?.contentType) || "application/octet-stream",
    contentBase64,
    size: Number(input?.size || 0) || estimateBase64Size(contentBase64),
    sourceEmailKey: normalizeString(input?.sourceEmailKey),
    sourceItemId: normalizeString(input?.sourceItemId),
    sourceInternetMessageId: normalizeMessageId(input?.sourceInternetMessageId),
    sourceConversationId: normalizeString(input?.sourceConversationId),
    sourceEmailSubject: normalizeString(input?.sourceEmailSubject),
    storageProvider: normalizeString(input?.storageProvider),
    storageBasePath: normalizeString(input?.storageBasePath),
    storagePathHint: normalizeString(input?.storagePathHint),
    createdAt: normalizeString(input?.createdAt) || nowIso(),
    updatedAt: normalizeString(input?.updatedAt) || nowIso(),
  };
}

function normalizeGroupTicketSeriesInput(input = {}, current = {}) {
  const name = normalizeString(input?.name) || normalizeString(current?.name) || "Serie";
  const prefix = normalizeTicketPrefix(input?.prefix ?? current?.prefix ?? name);
  const yearMode = normalizeTicketYearMode(input?.yearMode ?? current?.yearMode);
  const separator = normalizeTicketSeparator(input?.separator ?? current?.separator);
  return {
    id: normalizeString(input?.id) || normalizeString(current?.id) || `ticket_series_${crypto.randomUUID()}`,
    name,
    prefix,
    replyInstructions: normalizeString(input?.replyInstructions ?? current?.replyInstructions),
    yearMode,
    separator,
    nextNumber: normalizePositiveInt(input?.nextNumber ?? current?.nextNumber ?? 1, 1),
    padding: normalizePositiveInt(input?.padding ?? current?.padding ?? 4, 4, 2, 8),
    isActive: typeof input?.isActive === "boolean" ? input.isActive : current?.isActive !== false,
    createdAt: normalizeString(input?.createdAt) || normalizeString(current?.createdAt) || nowIso(),
    updatedAt: normalizeString(input?.updatedAt) || normalizeString(current?.updatedAt) || nowIso(),
  };
}

function normalizeGroupTicketInput(input = {}, current = {}) {
  const sequenceNumber = normalizePositiveInt(input?.sequenceNumber ?? current?.sequenceNumber ?? 1, 1);
  const padding = normalizePositiveInt(input?.padding ?? current?.padding ?? 4, 4, 2, 8);
  const prefix = normalizeTicketPrefix(input?.prefix || current?.prefix);
  const yearMode = normalizeTicketYearMode(input?.yearMode ?? current?.yearMode);
  const separator = normalizeTicketSeparator(input?.separator ?? current?.separator);
  const createdAt = normalizeString(input?.createdAt) || normalizeString(current?.createdAt) || nowIso();
  const yearValue = normalizeString(input?.yearValue || current?.yearValue) || getTicketYearValue(createdAt, yearMode);
  return {
    id: normalizeString(input?.id) || normalizeString(current?.id) || `ticket_${crypto.randomUUID()}`,
    seriesId: normalizeString(input?.seriesId || current?.seriesId),
    seriesName: normalizeString(input?.seriesName || current?.seriesName),
    prefix,
    yearMode,
    separator,
    yearValue,
    code: normalizeString(input?.code) || normalizeString(current?.code) || buildTicketCode(prefix, sequenceNumber, padding, { yearMode, separator, yearValue, dateValue: createdAt }),
    sequenceNumber,
    padding,
    title: normalizeString(input?.title) || normalizeString(current?.title) || "Ticket",
    description: normalizeString(input?.description) || normalizeString(current?.description),
    status: normalizeTicketStatus(input?.status ?? current?.status),
    labels: normalizeGroupLabels(
      Object.prototype.hasOwnProperty.call(input || {}, "labels")
        ? input?.labels
        : current?.labels
    ),
    groupIds: normalizeGroupIds(
      Object.prototype.hasOwnProperty.call(input || {}, "groupIds")
        ? input?.groupIds
        : current?.groupIds
    ),
    createdFromEmailKey: normalizeString(input?.createdFromEmailKey || current?.createdFromEmailKey),
    createdAt,
    updatedAt: normalizeString(input?.updatedAt) || normalizeString(current?.updatedAt) || nowIso(),
  };
}

function parseAttachmentsJson(value) {
  if (Array.isArray(value)) return normalizeAttachments(value);
  const raw = normalizeString(value);
  if (!raw) return [];
  try {
    return normalizeAttachments(JSON.parse(raw));
  } catch {
    return [];
  }
}

function hydrateStore(raw) {
  const store = createEmptyStore();
  const source = raw && typeof raw === "object" ? raw : {};

  store.version = STORE_VERSION;
  store.emails = source.emails && typeof source.emails === "object" ? source.emails : {};
  store.indexes = {
    itemIds: source.indexes?.itemIds && typeof source.indexes.itemIds === "object" ? source.indexes.itemIds : {},
    internetMessageIds: source.indexes?.internetMessageIds && typeof source.indexes.internetMessageIds === "object" ? source.indexes.internetMessageIds : {},
    fingerprints: source.indexes?.fingerprints && typeof source.indexes.fingerprints === "object" ? source.indexes.fingerprints : {},
    conversations: source.indexes?.conversations && typeof source.indexes.conversations === "object" ? source.indexes.conversations : {},
  };
  store.groups = source.groups && typeof source.groups === "object" ? source.groups : {};
  store.groupMembers = source.groupMembers && typeof source.groupMembers === "object" ? source.groupMembers : {};
  store.groupMemberLinks = source.groupMemberLinks && typeof source.groupMemberLinks === "object" ? source.groupMemberLinks : {};
  store.emailGroups = source.emailGroups && typeof source.emailGroups === "object" ? source.emailGroups : {};
  store.conversationGroups = source.conversationGroups && typeof source.conversationGroups === "object" ? source.conversationGroups : {};
  store.entityLinks = source.entityLinks && typeof source.entityLinks === "object" ? source.entityLinks : {};
  store.emailEntityLinks = source.emailEntityLinks && typeof source.emailEntityLinks === "object" ? source.emailEntityLinks : {};
  store.groupDocuments = source.groupDocuments && typeof source.groupDocuments === "object" ? source.groupDocuments : {};
  store.groupAttachmentFlags = source.groupAttachmentFlags && typeof source.groupAttachmentFlags === "object" ? source.groupAttachmentFlags : {};
  store.groupTicketSeries = source.groupTicketSeries && typeof source.groupTicketSeries === "object" ? source.groupTicketSeries : {};
  store.groupTickets = source.groupTickets && typeof source.groupTickets === "object" ? source.groupTickets : {};
  store.groupTicketEmails = source.groupTicketEmails && typeof source.groupTicketEmails === "object" ? source.groupTicketEmails : {};
  store.groupEmailTickets = source.groupEmailTickets && typeof source.groupEmailTickets === "object" ? source.groupEmailTickets : {};

  for (const [groupId, value] of Object.entries(store.groupMemberLinks || {})) {
    const gid = normalizeString(groupId);
    if (!gid || !value || typeof value !== "object") {
      delete store.groupMemberLinks[groupId];
      continue;
    }
    const normalizedEntries = {};
    for (const [emailId, meta] of Object.entries(value)) {
      const eid = normalizeString(emailId);
      if (!eid) continue;
      normalizedEntries[eid] = normalizeGroupMembershipMeta(meta);
      const members = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
      if (!members.includes(eid)) store.groupMembers[gid] = [...members, eid];
      const emailGroups = Array.isArray(store.emailGroups[eid]) ? store.emailGroups[eid] : [];
      if (!emailGroups.includes(gid)) store.emailGroups[eid] = [...emailGroups, gid];
    }
    store.groupMemberLinks[gid] = normalizedEntries;
  }

  for (const [groupId, members] of Object.entries(store.groupMembers || {})) {
    const gid = normalizeString(groupId);
    if (!gid) continue;
    if (!store.groupMemberLinks[gid] || typeof store.groupMemberLinks[gid] !== "object" || Array.isArray(store.groupMemberLinks[gid])) {
      store.groupMemberLinks[gid] = {};
    }
    const dedupedMembers = Array.from(new Set((Array.isArray(members) ? members : []).map((entry) => normalizeString(entry)).filter(Boolean)));
    store.groupMembers[gid] = dedupedMembers;
    for (const emailId of dedupedMembers) {
      if (!store.groupMemberLinks[gid][emailId]) {
        store.groupMemberLinks[gid][emailId] = normalizeGroupMembershipMeta({});
      }
      const emailGroups = Array.isArray(store.emailGroups[emailId]) ? store.emailGroups[emailId] : [];
      if (!emailGroups.includes(gid)) store.emailGroups[emailId] = [...emailGroups, gid];
    }
  }

  for (const [seriesId, value] of Object.entries(store.groupTicketSeries || {})) {
    const sid = normalizeString(seriesId);
    const normalized = normalizeGroupTicketSeriesInput({ id: sid, ...(value || {}) });
    if (!sid || !normalized.prefix) {
      delete store.groupTicketSeries[seriesId];
      continue;
    }
    store.groupTicketSeries[sid] = normalized;
  }

  for (const [ticketId, value] of Object.entries(store.groupTickets || {})) {
    const tid = normalizeString(ticketId);
    const normalized = normalizeGroupTicketInput({ id: tid, ...(value || {}) });
    if (!tid || !normalized.seriesId || !normalized.code) {
      delete store.groupTickets[ticketId];
      continue;
    }
    store.groupTickets[tid] = normalized;
  }

  for (const [ticketId, emailKeys] of Object.entries(store.groupTicketEmails || {})) {
    const tid = normalizeString(ticketId);
    if (!tid || !store.groupTickets[tid]) {
      delete store.groupTicketEmails[ticketId];
      continue;
    }
    store.groupTicketEmails[tid] = Array.from(
      new Set((Array.isArray(emailKeys) ? emailKeys : []).map((entry) => normalizeString(entry)).filter(Boolean))
    );
  }

  for (const [emailKey, ticketIds] of Object.entries(store.groupEmailTickets || {})) {
    const key = normalizeString(emailKey);
    if (!key) {
      delete store.groupEmailTickets[emailKey];
      continue;
    }
    store.groupEmailTickets[key] = Array.from(
      new Set((Array.isArray(ticketIds) ? ticketIds : []).map((entry) => normalizeString(entry)).filter((entry) => Boolean(store.groupTickets[entry])))
    );
  }
  return store;
}

function normalizeAttachmentFlagInput(input = {}) {
  return {
    attachmentKey: normalizeString(input?.attachmentKey),
    emailKey: normalizeString(input?.emailKey),
    attachmentName: normalizeString(input?.attachmentName),
    contentType: normalizeString(input?.contentType),
    size: Number(input?.size || 0) || 0,
    disposition: normalizeString(input?.disposition) || "dismissed",
    createdAt: normalizeString(input?.createdAt) || nowIso(),
    updatedAt: normalizeString(input?.updatedAt) || nowIso(),
  };
}

function upsertConversationIndex(store, conversationId, emailId) {
  const cid = normalizeString(conversationId);
  if (!cid) return;
  const current = Array.isArray(store.indexes.conversations[cid]) ? store.indexes.conversations[cid] : [];
  if (!current.includes(emailId)) store.indexes.conversations[cid] = [...current, emailId];
}

function resolveEmailId(store, input) {
  const normalized = normalizeEmailInput(input);
  if (normalized.itemId && store.indexes.itemIds[normalized.itemId]) {
    return store.indexes.itemIds[normalized.itemId];
  }
  if (normalized.internetMessageId && store.indexes.internetMessageIds[normalized.internetMessageId]) {
    return store.indexes.internetMessageIds[normalized.internetMessageId];
  }
  const fingerprint = makeEmailFingerprint(normalized);
  if (fingerprint && store.indexes.fingerprints[fingerprint]) {
    return store.indexes.fingerprints[fingerprint];
  }
  return "";
}

function upsertEmail(store, input) {
  const normalized = normalizeEmailInput(input);
  const now = nowIso();
  const emailId = resolveEmailId(store, normalized) || `email_${crypto.randomUUID()}`;
  const current = store.emails[emailId] || {
    id: emailId,
    createdAt: now,
  };
  const next = {
    ...current,
    ...Object.fromEntries(Object.entries(normalized).filter(([key, value]) => {
      if (key === "status" || key === "labels" || key === "labelStates") {
        return Object.prototype.hasOwnProperty.call(input || {}, key);
      }
      return Boolean(value);
    })),
    updatedAt: now,
    lastSeenAt: now,
  };

  if (Object.prototype.hasOwnProperty.call(input || {}, "status") && !normalized.status) {
    delete next.status;
  }
  if (Object.prototype.hasOwnProperty.call(input || {}, "labels")) {
    next.labels = Array.isArray(normalized.labels) ? normalized.labels : [];
  }
  if (Object.prototype.hasOwnProperty.call(input || {}, "labelStates")) {
    next.labelStates = normalized.labelStates && typeof normalized.labelStates === "object" ? normalized.labelStates : {};
  }

  if (!next.messageDateIso) {
    next.messageDateIso = next.receivedAtIso || next.sentAtIso || next.linkedAt || now;
  }
  if (!next.receivedAtIso) next.receivedAtIso = next.messageDateIso;

  store.emails[emailId] = next;

  if (next.itemId) store.indexes.itemIds[next.itemId] = emailId;
  if (next.internetMessageId) store.indexes.internetMessageIds[next.internetMessageId] = emailId;
  const fingerprint = makeEmailFingerprint(next);
  if (fingerprint) store.indexes.fingerprints[fingerprint] = emailId;
  if (next.conversationId) upsertConversationIndex(store, next.conversationId, emailId);

  return next;
}

function ensureGroup(store, partial) {
  const now = nowIso();
  const id = normalizeString(partial?.id) || `group_${crypto.randomUUID()}`;
  const current = store.groups[id] || {
    id,
    createdAt: now,
  };
  const archivedRequested = typeof partial?.isArchived === "boolean" ? partial.isArchived : current.isArchived === true;
  const nextArchivedAt = archivedRequested
    ? normalizeString(partial?.archivedAt) || normalizeString(current.archivedAt) || now
    : "";
  const next = {
    ...current,
    kind: normalizeString(partial?.kind) || current.kind || "custom",
    name: normalizeString(partial?.name) || current.name || "Grupo sem nome",
    description: normalizeString(partial?.description) || current.description || "",
    conversationId: normalizeString(partial?.conversationId) || current.conversationId || "",
    status: normalizeGroupStatus(partial?.status || current.status),
    labels: normalizeGroupLabels(
      Object.prototype.hasOwnProperty.call(partial || {}, "labels")
        ? partial?.labels
        : current.labels
    ),
    isArchived: archivedRequested,
    archivedAt: nextArchivedAt,
    documentsEnabled:
      typeof partial?.documentsEnabled === "boolean"
        ? partial.documentsEnabled
        : typeof current.documentsEnabled === "boolean"
          ? current.documentsEnabled
          : true,
    updatedAt: now,
  };
  if (!next.createdAt) next.createdAt = now;
  store.groups[id] = next;
  if (!Array.isArray(store.groupMembers[id])) store.groupMembers[id] = [];
  if (!store.groupMemberLinks[id] || typeof store.groupMemberLinks[id] !== "object") store.groupMemberLinks[id] = {};
  return next;
}

function ensureConversationGroup(store, conversationId, sampleEmail = null) {
  const cid = normalizeString(conversationId);
  if (!cid) return null;
  const existingId = normalizeString(store.conversationGroups[cid]);
  const sampleSubject = normalizeString(sampleEmail?.subject);
  if (existingId && store.groups[existingId]) {
    const current = store.groups[existingId];
    if (sampleSubject && (!current.name || current.name.startsWith("Conversa "))) {
      current.name = `Conversa ${sampleSubject}`;
      current.updatedAt = nowIso();
    }
    return current;
  }

  const labelSource = sampleSubject || cid.slice(0, 10);
  const group = ensureGroup(store, {
    kind: "conversation",
    name: `Conversa ${labelSource}`,
    conversationId: cid,
  });
  store.conversationGroups[cid] = group.id;
  return group;
}

function getEmailMembershipMeta(store, groupId, emailId) {
  const gid = normalizeString(groupId);
  const eid = normalizeString(emailId);
  if (!gid || !eid) return normalizeGroupMembershipMeta({});
  const raw = store?.groupMemberLinks?.[gid]?.[eid];
  return normalizeGroupMembershipMeta(raw || {});
}

function listEmailGroupMemberships(store, emailId) {
  const eid = normalizeString(emailId);
  if (!eid) return [];
  const groupIds = Array.isArray(store?.emailGroups?.[eid]) ? store.emailGroups[eid] : [];
  return groupIds
    .map((groupId) => ({
      groupId: normalizeString(groupId),
      kind: getEmailMembershipMeta(store, groupId, eid).kind,
    }))
    .filter((entry) => entry.groupId);
}

function addEmailMembership(store, groupId, emailId, options = {}) {
  const gid = normalizeString(groupId);
  const eid = normalizeString(emailId);
  if (!gid || !eid) return;
  const now = nowIso();
  const members = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
  if (!members.includes(eid)) store.groupMembers[gid] = [...members, eid];
  const emailGroups = Array.isArray(store.emailGroups[eid]) ? store.emailGroups[eid] : [];
  if (!emailGroups.includes(gid)) store.emailGroups[eid] = [...emailGroups, gid];
  if (!store.groupMemberLinks[gid] || typeof store.groupMemberLinks[gid] !== "object") store.groupMemberLinks[gid] = {};
  const currentMeta = normalizeGroupMembershipMeta(store.groupMemberLinks[gid][eid]);
  store.groupMemberLinks[gid][eid] = {
    kind: normalizeGroupMembershipKind(options?.membershipKind || currentMeta.kind),
    linkedAt: currentMeta.linkedAt || normalizeString(options?.linkedAt) || now,
    updatedAt: now,
  };
  if (store.groups[gid]) store.groups[gid].updatedAt = nowIso();
}

function removeEmailMembership(store, groupId, emailId) {
  const gid = normalizeString(groupId);
  const eid = normalizeString(emailId);
  if (!gid || !eid) return false;
  const members = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
  const nextMembers = members.filter((value) => value !== eid);
  if (nextMembers.length === members.length) return false;
  store.groupMembers[gid] = nextMembers;
  if (store.groupMemberLinks[gid] && typeof store.groupMemberLinks[gid] === "object") {
    delete store.groupMemberLinks[gid][eid];
    if (!Object.keys(store.groupMemberLinks[gid]).length) delete store.groupMemberLinks[gid];
  }

  const emailGroups = Array.isArray(store.emailGroups[eid]) ? store.emailGroups[eid] : [];
  store.emailGroups[eid] = emailGroups.filter((value) => value !== gid);
  if (!store.emailGroups[eid].length) delete store.emailGroups[eid];
  if (store.groups[gid]) store.groups[gid].updatedAt = nowIso();
  return true;
}

function linkEmailToEntity(store, emailId, rawLink) {
  const model = normalizeModel(rawLink?.model);
  const recordId = normalizeRecordId(rawLink?.recordId ?? rawLink?.resId);
  if (!model || !recordId) return null;
  const entityKey = makeEntityKey(model, recordId);
  const email = store.emails[emailId] || {};
  const nextLink = {
    model,
    recordId,
    recordName: normalizeString(rawLink?.recordName || rawLink?.name || rawLink?.title),
    linkedAt: normalizeString(rawLink?.linkedAt) || nowIso(),
    conversationId: normalizeString(rawLink?.conversationId || email.conversationId),
    internetMessageId: normalizeMessageId(rawLink?.internetMessageId || email.internetMessageId),
    itemId: normalizeString(rawLink?.itemId || email.itemId),
    emailWebLink: normalizeString(rawLink?.emailWebLink || rawLink?.url || email.emailWebLink),
    messageDateIso: normalizeString(rawLink?.messageDateIso || rawLink?.receivedAtIso || email.messageDateIso),
    receivedAtIso: normalizeString(rawLink?.receivedAtIso || email.receivedAtIso || rawLink?.messageDateIso),
    subject: normalizeString(rawLink?.subject || email.subject),
    fromEmail: normalizeString(rawLink?.fromEmail || email.fromEmail),
    fromName: normalizeString(rawLink?.fromName || email.fromName),
  };

  const currentByEmail = Array.isArray(store.emailEntityLinks[emailId]) ? store.emailEntityLinks[emailId] : [];
  const nextByEmail = dedupeRecordLinks([{ ...nextLink }, ...currentByEmail]);
  store.emailEntityLinks[emailId] = nextByEmail;

  const currentByEntity = Array.isArray(store.entityLinks[entityKey]) ? store.entityLinks[entityKey] : [];
  if (!currentByEntity.includes(emailId)) store.entityLinks[entityKey] = [emailId, ...currentByEntity];
  return nextLink;
}

function buildLinkEntry(email, entityLink) {
  return {
    conversationId: normalizeString(entityLink?.conversationId || email?.conversationId),
    model: normalizeModel(entityLink?.model),
    recordId: normalizeRecordId(entityLink?.recordId),
    recordName: normalizeString(entityLink?.recordName),
    resId: normalizeRecordId(entityLink?.recordId),
    name: normalizeString(entityLink?.recordName),
    title: normalizeString(entityLink?.recordName),
    linkedAt: normalizeString(entityLink?.linkedAt),
    internetMessageId: normalizeMessageId(entityLink?.internetMessageId || email?.internetMessageId),
    itemId: normalizeString(entityLink?.itemId || email?.itemId),
    emailWebLink: normalizeString(entityLink?.emailWebLink || email?.emailWebLink),
    messageDateIso: normalizeString(entityLink?.messageDateIso || email?.messageDateIso),
    receivedAtIso: normalizeString(entityLink?.receivedAtIso || email?.receivedAtIso),
    subject: normalizeString(entityLink?.subject || email?.subject),
    fromEmail: normalizeString(entityLink?.fromEmail || email?.fromEmail),
    fromName: normalizeString(entityLink?.fromName || email?.fromName),
  };
}

function buildEmailListEntry(email, extra = {}) {
  return {
    emailKey: makePersistentEmailKey(email),
    id: normalizeString(email?.id),
    conversationId: normalizeString(email?.conversationId),
    itemId: normalizeString(email?.itemId),
    internetMessageId: normalizeMessageId(email?.internetMessageId),
    emailWebLink: normalizeString(email?.emailWebLink),
    subject: normalizeString(email?.subject),
    fromEmail: normalizeString(email?.fromEmail),
    fromName: normalizeString(email?.fromName),
    messageDateIso: normalizeString(email?.messageDateIso),
    receivedAtIso: normalizeString(email?.receivedAtIso || email?.messageDateIso),
    sentAtIso: normalizeString(email?.sentAtIso),
    bodyText: normalizeString(email?.bodyText),
    bodyHtml: normalizeString(email?.bodyHtml),
    status: normalizeString(email?.status),
    labels: normalizeGroupLabels(email?.labels),
    labelStates: normalizeEmailLabelStates(email?.labelStates),
    createdAt: normalizeString(email?.createdAt),
    updatedAt: normalizeString(email?.updatedAt),
    attachments: normalizeAttachments(email?.attachments),
    membershipKind: normalizeGroupMembershipKind(extra?.membershipKind || email?.membershipKind),
    ...extra,
  };
}

function buildCurrentEmailContextEntry(store, emailId) {
  const eid = normalizeString(emailId);
  const email = store?.emails?.[eid];
  if (!eid || !email) return null;

  const relatedGroups = listEmailGroupMemberships(store, eid)
    .map((entry) => {
      const group = buildGroupListEntry(store, store?.groups?.[entry.groupId]);
      if (!group) return null;
      return {
        id: group.id,
        name: group.name,
        kind: group.kind === "conversation" ? "conversation" : "group",
        relationKind: normalizeGroupMembershipKind(entry.kind),
      };
    })
    .filter(Boolean);

  const businessGroups = relatedGroups.filter((entry) => entry?.kind !== "conversation");
  const principalGroup = businessGroups.find((entry) => normalizeGroupMembershipKind(entry?.relationKind) === "principal") || businessGroups[0] || null;
  const relatedRecords = Array.isArray(store?.emailEntityLinks?.[eid])
    ? dedupeRecordLinks(store.emailEntityLinks[eid]).map((entry) => ({
      model: entry.model,
      recordId: entry.recordId,
      recordName: entry.recordName,
    }))
    : [];

  return buildEmailListEntry(email, {
    groupId: principalGroup?.id || "",
    groupName: principalGroup?.name || "",
    membershipKind: principalGroup?.relationKind || email?.membershipKind,
    relatedGroups,
    relatedRecords,
    relatedReasons: [],
  });
}

function mergeEmailContextEntries(baseEmail, overlayEmail) {
  if (!baseEmail) return overlayEmail || null;
  if (!overlayEmail) return baseEmail;

  const mergedRelatedGroups = [
    ...(Array.isArray(baseEmail.relatedGroups) ? baseEmail.relatedGroups : []),
    ...(Array.isArray(overlayEmail.relatedGroups) ? overlayEmail.relatedGroups : []),
  ].reduce((acc, entry) => {
    if (!entry?.id || acc.some((current) => current.id === entry.id)) return acc;
    acc.push(entry);
    return acc;
  }, []);

  const mergedRelatedRecords = [
    ...(Array.isArray(baseEmail.relatedRecords) ? baseEmail.relatedRecords : []),
    ...(Array.isArray(overlayEmail.relatedRecords) ? overlayEmail.relatedRecords : []),
  ].reduce((acc, entry) => {
    const key = makeEntityKey(entry?.model, entry?.recordId);
    if (!key || acc.some((current) => makeEntityKey(current?.model, current?.recordId) === key)) return acc;
    acc.push(entry);
    return acc;
  }, []);

  const businessGroups = mergedRelatedGroups.filter((entry) => entry?.kind !== "conversation");
  const principalGroup = businessGroups.find((entry) => normalizeGroupMembershipKind(entry?.relationKind) === "principal") || businessGroups[0] || null;

  return {
    ...baseEmail,
    ...overlayEmail,
    groupId: overlayEmail.groupId || principalGroup?.id || baseEmail.groupId || "",
    groupName: overlayEmail.groupName || principalGroup?.name || baseEmail.groupName || "",
    membershipKind: overlayEmail.membershipKind || principalGroup?.relationKind || baseEmail.membershipKind,
    relatedGroups: mergedRelatedGroups,
    relatedRecords: mergedRelatedRecords,
  };
}

function buildEmailSearchEntry(email, extra = {}) {
  return {
    emailKey: makePersistentEmailKey(email),
    id: normalizeString(email?.id),
    conversationId: normalizeString(email?.conversationId),
    itemId: normalizeString(email?.itemId),
    internetMessageId: normalizeMessageId(email?.internetMessageId),
    emailWebLink: normalizeString(email?.emailWebLink),
    subject: normalizeString(email?.subject),
    fromEmail: normalizeString(email?.fromEmail),
    fromName: normalizeString(email?.fromName),
    messageDateIso: normalizeString(email?.messageDateIso),
    receivedAtIso: normalizeString(email?.receivedAtIso || email?.messageDateIso),
    sentAtIso: normalizeString(email?.sentAtIso),
    createdAt: normalizeString(email?.createdAt),
    updatedAt: normalizeString(email?.updatedAt),
    membershipKind: normalizeGroupMembershipKind(extra?.membershipKind || email?.membershipKind),
    ...extra,
  };
}

function buildRecoveredEmailSnapshot(store, emailId) {
  const existing = store.emails[emailId];
  if (existing) return existing;

  const entityLinks = Array.isArray(store.emailEntityLinks[emailId]) ? dedupeRecordLinks(store.emailEntityLinks[emailId]) : [];
  const source = entityLinks[0];
  if (!source) return null;

  return {
    id: normalizeString(emailId),
    itemId: normalizeString(source?.itemId),
    internetMessageId: normalizeMessageId(source?.internetMessageId),
    conversationId: normalizeString(source?.conversationId),
    subject: normalizeString(source?.subject),
    fromEmail: normalizeString(source?.fromEmail),
    fromName: normalizeString(source?.fromName),
    emailWebLink: normalizeString(source?.emailWebLink),
    messageDateIso: normalizeString(source?.messageDateIso || source?.receivedAtIso || source?.linkedAt),
    receivedAtIso: normalizeString(source?.receivedAtIso || source?.messageDateIso || source?.linkedAt),
    sentAtIso: normalizeString(source?.sentAtIso),
    linkedAt: normalizeString(source?.linkedAt),
    createdAt: normalizeString(source?.linkedAt || nowIso()),
    updatedAt: normalizeString(source?.linkedAt || nowIso()),
  };
}

function makePersistentEmailKey(email) {
  return makeEmailLookupKey(email);
}

function buildGroupListEntry(store, group) {
  if (!group) return null;
  const gid = normalizeString(group.id);
  return {
    ...group,
    status: normalizeGroupStatus(group.status),
    labels: normalizeGroupLabels(group.labels),
    isArchived: group.isArchived === true,
    archivedAt: group.isArchived === true ? normalizeString(group.archivedAt) : "",
    documentsEnabled: group.documentsEnabled !== false,
    memberCount: Array.isArray(store?.groupMembers?.[gid]) ? store.groupMembers[gid].length : Number(group.memberCount || 0) || 0,
    documentCount: Array.isArray(store?.groupDocuments?.[gid]) ? store.groupDocuments[gid].length : Number(group.documentCount || 0) || 0,
  };
}

function buildGroupTicketSeriesEntry(store, series) {
  if (!series) return null;
  const sid = normalizeString(series.id);
  const usageCount = Object.values(store.groupTickets || {}).filter((ticket) => normalizeString(ticket?.seriesId) === sid).length;
  return {
    ...series,
    prefix: normalizeTicketPrefix(series.prefix),
    yearMode: normalizeTicketYearMode(series.yearMode),
    separator: normalizeTicketSeparator(series.separator),
    nextNumber: normalizePositiveInt(series.nextNumber, 1),
    padding: normalizePositiveInt(series.padding, 4, 2, 8),
    isActive: series.isActive !== false,
    usageCount,
  };
}

function buildGroupTicketEntry(store, ticket, extra = {}) {
  if (!ticket) return null;
  const series = store.groupTicketSeries?.[normalizeString(ticket.seriesId)];
  const groupIds = normalizeGroupIds(extra?.groupIds || ticket.groupIds);
  const emailKeys = Array.isArray(store.groupTicketEmails?.[normalizeString(ticket.id)]) ? store.groupTicketEmails[ticket.id] : [];
  return {
    ...ticket,
    seriesName: normalizeString(extra?.seriesName || ticket.seriesName || series?.name),
    prefix: normalizeTicketPrefix(extra?.prefix || ticket.prefix || series?.prefix),
    yearMode: normalizeTicketYearMode(extra?.yearMode ?? ticket.yearMode ?? series?.yearMode),
    separator: normalizeTicketSeparator(extra?.separator ?? ticket.separator ?? series?.separator),
    yearValue: normalizeString(extra?.yearValue || ticket.yearValue) || getTicketYearValue(ticket.createdAt, extra?.yearMode ?? ticket.yearMode ?? series?.yearMode),
    code: normalizeString(ticket.code),
    sequenceNumber: normalizePositiveInt(ticket.sequenceNumber, 1),
    padding: normalizePositiveInt(ticket.padding || series?.padding || 4, 4, 2, 8),
    status: normalizeTicketStatus(ticket.status),
    labels: normalizeGroupLabels(ticket.labels),
    groupIds,
    groups: groupIds
      .map((groupId) => buildGroupListEntry(store, store.groups?.[groupId]))
      .filter(Boolean),
    emailCount: Array.isArray(emailKeys) ? emailKeys.length : 0,
    emailLinked: Boolean(extra?.emailLinked),
  };
}

function resolveEmailKeyFromInput(store, input) {
  const normalized = normalizeEmailInput(input);
  const directKey = normalizeString(input?.emailKey);
  if (directKey) return directKey;
  if (
    !normalized.itemId
    && !normalized.internetMessageId
    && !normalized.conversationId
    && !normalized.subject
    && !normalized.fromEmail
    && !normalized.messageDateIso
    && !normalized.receivedAtIso
    && !normalized.linkedAt
  ) {
    return "";
  }
  const emailId = resolveEmailId(store, normalized);
  if (emailId && store.emails[emailId]) {
    return makePersistentEmailKey(store.emails[emailId]);
  }
  return makePersistentEmailKey(normalized);
}

function ensureTicketEmailLink(store, ticketId, emailKey) {
  const tid = normalizeString(ticketId);
  const key = normalizeString(emailKey);
  if (!tid || !key) return;
  const ticketEmails = Array.isArray(store.groupTicketEmails[tid]) ? store.groupTicketEmails[tid] : [];
  if (!ticketEmails.includes(key)) store.groupTicketEmails[tid] = [...ticketEmails, key];
  const emailTickets = Array.isArray(store.groupEmailTickets[key]) ? store.groupEmailTickets[key] : [];
  if (!emailTickets.includes(tid)) store.groupEmailTickets[key] = [...emailTickets, tid];
}

function removeTicketEmailLink(store, ticketId, emailKey) {
  const tid = normalizeString(ticketId);
  const key = normalizeString(emailKey);
  if (!tid || !key) return;
  if (Array.isArray(store.groupTicketEmails[tid])) {
    store.groupTicketEmails[tid] = store.groupTicketEmails[tid].filter((entry) => normalizeString(entry) !== key);
    if (!store.groupTicketEmails[tid].length) delete store.groupTicketEmails[tid];
  }
  if (Array.isArray(store.groupEmailTickets[key])) {
    store.groupEmailTickets[key] = store.groupEmailTickets[key].filter((entry) => normalizeString(entry) !== tid);
    if (!store.groupEmailTickets[key].length) delete store.groupEmailTickets[key];
  }
}

function listTicketIdsByEmailKey(store, emailKey) {
  const key = normalizeString(emailKey);
  return Array.isArray(store.groupEmailTickets?.[key]) ? store.groupEmailTickets[key].filter(Boolean) : [];
}

function extractTicketCandidates(store, input) {
  const subject = normalizeString(input?.subject);
  const bodyText = normalizeString(input?.bodyText);
  const bodyHtml = stripHtmlForTicketLookup(input?.bodyHtml);
  const haystack = [subject, bodyText, bodyHtml].filter(Boolean).join("\n").toUpperCase();
  if (!haystack) return [];

  const matches = new Map();
  for (const ticket of Object.values(store.groupTickets || {})) {
    const normalizedTicket = normalizeGroupTicketInput(ticket, ticket);
    const code = normalizeString(normalizedTicket.code).toUpperCase();
    if (!code) continue;
    const escapedCode = code.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const regex = new RegExp(`(?:^|[^A-Z0-9])(?:\\[)?(${escapedCode})(?:\\])?(?=$|[^A-Z0-9])`, "i");
    const match = haystack.match(regex);
    if (!match) continue;
    matches.set(normalizedTicket.id, { ticket: normalizedTicket, matchedCode: match[1] || code });
  }

  return Array.from(matches.values());
}

function mapDbGroupRow(row) {
  if (!row) return null;
  return {
    id: normalizeString(row.id),
    kind: CUSTOM_GROUP_KIND,
    name: normalizeString(row.name),
    description: normalizeString(row.description),
    status: normalizeGroupStatus(row.status),
    labels: parseGroupLabelsJson(row.labels_json),
    isArchived: row.is_archived === true,
    archivedAt: row.is_archived === true ? normalizeString(row.archived_at) : "",
    documentsEnabled: row.documents_enabled !== false,
    createdAt: normalizeString(row.created_at),
    updatedAt: normalizeString(row.updated_at),
  };
}

function mapDbGroupMemberRow(row) {
  return buildEmailListEntry({
    id: normalizeString(row.email_key),
    itemId: normalizeString(row.item_id),
    internetMessageId: normalizeMessageId(row.internet_message_id),
    conversationId: normalizeString(row.conversation_id),
    subject: normalizeString(row.subject),
    fromEmail: normalizeString(row.from_email),
    fromName: normalizeString(row.from_name),
    emailWebLink: normalizeString(row.email_web_link),
    messageDateIso: normalizeString(row.message_date_iso),
    receivedAtIso: normalizeString(row.received_at_iso),
    sentAtIso: normalizeString(row.sent_at_iso),
    bodyText: normalizeString(row.body_text),
    bodyHtml: normalizeString(row.body_html),
    createdAt: normalizeString(row.created_at),
    updatedAt: normalizeString(row.updated_at),
    attachments: parseAttachmentsJson(row.attachments_json),
    membershipKind: normalizeGroupMembershipKind(row.relation_kind),
  });
}

function mapDbGroupDocumentRow(row) {
  if (!row) return null;
  return normalizeGroupDocumentInput({
    id: row.id,
    name: row.name,
    contentType: row.content_type,
    contentBase64: row.content_base64,
    size: row.size_bytes,
    sourceEmailKey: row.source_email_key,
    sourceItemId: row.source_item_id,
    sourceInternetMessageId: row.source_internet_message_id,
    sourceConversationId: row.source_conversation_id,
    sourceEmailSubject: row.source_email_subject,
    storageProvider: row.storage_provider,
    storageBasePath: row.storage_base_path,
    storagePathHint: row.storage_path_hint,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  });
}

function mapDbGroupAttachmentFlagRow(row) {
  if (!row) return null;
  return normalizeAttachmentFlagInput({
    attachmentKey: row.attachment_key,
    emailKey: row.email_key,
    attachmentName: row.attachment_name,
    contentType: row.content_type,
    size: row.size_bytes,
    disposition: row.disposition,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  });
}

function mapDbGroupTicketSeriesRow(row) {
  if (!row) return null;
  return normalizeGroupTicketSeriesInput({
    id: row.id,
    name: row.name,
    prefix: row.prefix,
    replyInstructions: row.reply_instructions,
    yearMode: row.year_mode,
    separator: row.separator,
    nextNumber: row.next_number,
    padding: row.padding,
    isActive: row.is_active,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  });
}

function mapDbGroupTicketRow(row) {
  if (!row) return null;
  return normalizeGroupTicketInput({
    id: row.id,
    seriesId: row.series_id,
    seriesName: row.series_name,
    prefix: row.prefix,
    yearMode: row.year_mode,
    separator: row.separator,
    yearValue: row.year_value,
    code: row.code,
    sequenceNumber: row.sequence_number,
    padding: row.padding,
    title: row.title,
    description: row.description,
    status: row.status,
    labels: parseGroupLabelsJson(row.labels_json),
    groupIds: parseGroupLabelsJson(row.group_ids_json),
    createdFromEmailKey: row.created_from_email_key,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  });
}

async function upsertDbGroupTicketSeries(input) {
  if (!db.isEnabled()) return;
  const series = normalizeGroupTicketSeriesInput(input);
  if (!series.id || !series.prefix) return;
  await db.query(
    `INSERT INTO crm_group_ticket_series (id, name, prefix, reply_instructions, year_mode, separator, next_number, padding, is_active, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11)
     ON CONFLICT (id) DO UPDATE SET
       name = EXCLUDED.name,
       prefix = EXCLUDED.prefix,
       reply_instructions = EXCLUDED.reply_instructions,
       year_mode = EXCLUDED.year_mode,
       separator = EXCLUDED.separator,
       next_number = EXCLUDED.next_number,
       padding = EXCLUDED.padding,
       is_active = EXCLUDED.is_active,
       updated_at = EXCLUDED.updated_at`,
    [
      series.id,
      series.name,
      series.prefix,
      series.replyInstructions,
      series.yearMode,
      series.separator,
      series.nextNumber,
      series.padding,
      series.isActive !== false,
      series.createdAt,
      series.updatedAt,
    ]
  );
}

async function upsertDbGroupTicket(input) {
  if (!db.isEnabled()) return;
  const ticket = normalizeGroupTicketInput(input);
  if (!ticket.id || !ticket.seriesId || !ticket.code) return;
  await db.query(
    `INSERT INTO crm_group_tickets (id, series_id, prefix, year_mode, separator, year_value, code, sequence_number, padding, title, description, status, labels_json, group_ids_json, created_from_email_key, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13::jsonb, $14::jsonb, $15, $16, $17)
     ON CONFLICT (id) DO UPDATE SET
       series_id = EXCLUDED.series_id,
       prefix = EXCLUDED.prefix,
       year_mode = EXCLUDED.year_mode,
       separator = EXCLUDED.separator,
       year_value = EXCLUDED.year_value,
       code = EXCLUDED.code,
       sequence_number = EXCLUDED.sequence_number,
       padding = EXCLUDED.padding,
       title = EXCLUDED.title,
       description = EXCLUDED.description,
       status = EXCLUDED.status,
       labels_json = EXCLUDED.labels_json,
       group_ids_json = EXCLUDED.group_ids_json,
       created_from_email_key = EXCLUDED.created_from_email_key,
       updated_at = EXCLUDED.updated_at`,
    [
      ticket.id,
      ticket.seriesId,
      ticket.prefix,
      ticket.yearMode,
      ticket.separator,
      ticket.yearValue,
      ticket.code,
      ticket.sequenceNumber,
      ticket.padding,
      ticket.title,
      ticket.description,
      ticket.status,
      JSON.stringify(normalizeGroupLabels(ticket.labels)),
      JSON.stringify(normalizeGroupIds(ticket.groupIds)),
      ticket.createdFromEmailKey,
      ticket.createdAt,
      ticket.updatedAt,
    ]
  );
}

async function upsertDbGroupTicketEmail(ticketId, emailKey) {
  if (!db.isEnabled()) return;
  const tid = normalizeString(ticketId);
  const key = normalizeString(emailKey);
  if (!tid || !key) return;
  await db.query(
    `INSERT INTO crm_group_ticket_emails (ticket_id, email_key, created_at, updated_at)
     VALUES ($1, $2, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
     ON CONFLICT (ticket_id, email_key) DO UPDATE SET
       updated_at = CURRENT_TIMESTAMP`,
    [tid, key]
  );
}

async function deleteDbGroupTicketEmail(ticketId, emailKey) {
  if (!db.isEnabled()) return;
  const tid = normalizeString(ticketId);
  const key = normalizeString(emailKey);
  if (!tid || !key) return;
  await db.query(`DELETE FROM crm_group_ticket_emails WHERE ticket_id = $1 AND email_key = $2`, [tid, key]);
}

async function deleteDbGroupTicketSeries(seriesId) {
  if (!db.isEnabled()) return;
  const sid = normalizeString(seriesId);
  if (!sid) return;
  await db.query(`DELETE FROM crm_group_ticket_series WHERE id = $1`, [sid]);
}

async function listDbGroupTicketSeries() {
  if (!db.isEnabled()) return [];
  const result = await db.query(`SELECT * FROM crm_group_ticket_series ORDER BY is_active DESC, prefix ASC, name ASC`);
  return (result?.rows || []).map(mapDbGroupTicketSeriesRow).filter(Boolean);
}

async function listDbGroupTickets(query = "", options = {}) {
  if (!db.isEnabled()) return [];
  const q = normalizeString(query).toLowerCase();
  const groupId = normalizeString(options?.groupId);
  const emailKey = normalizeString(options?.emailKey);
  const limit = Math.max(1, Math.min(normalizePositiveInt(options?.limit || 50, 50), 200));
  const where = [];
  const params = [];
  if (q) {
    params.push(`%${q}%`);
    where.push(`(LOWER(t.code) LIKE $${params.length} OR LOWER(t.title) LIKE $${params.length} OR LOWER(COALESCE(t.description, '')) LIKE $${params.length})`);
  }
  if (groupId) {
    params.push(groupId);
    where.push(`t.group_ids_json ? $${params.length}`);
  }
  if (emailKey) {
    params.push(emailKey);
    where.push(`EXISTS (SELECT 1 FROM crm_group_ticket_emails te WHERE te.ticket_id = t.id AND te.email_key = $${params.length})`);
  }
  params.push(limit);
  const querySql = `
    SELECT t.*, s.name AS series_name, s.prefix
    FROM crm_group_tickets t
    LEFT JOIN crm_group_ticket_series s ON s.id = t.series_id
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY t.updated_at DESC, t.created_at DESC
    LIMIT $${params.length}
  `;
  const result = await db.query(querySql, params);
  return (result?.rows || []).map(mapDbGroupTicketRow).filter(Boolean);
}

async function listDbGroupTicketEmailLinks(ticketId = "") {
  if (!db.isEnabled()) return [];
  const tid = normalizeString(ticketId);
  const result = tid
    ? await db.query(`SELECT ticket_id, email_key FROM crm_group_ticket_emails WHERE ticket_id = $1 ORDER BY updated_at DESC, created_at DESC`, [tid])
    : await db.query(`SELECT ticket_id, email_key FROM crm_group_ticket_emails ORDER BY updated_at DESC, created_at DESC`);
  return (result?.rows || [])
    .map((row) => ({
      ticketId: normalizeString(row.ticket_id),
      emailKey: normalizeString(row.email_key),
    }))
    .filter((row) => row.ticketId && row.emailKey);
}

async function syncGroupTicketsFromDb(store) {
  if (!db.isEnabled()) return;
  await ensureCustomGroupDb();
  const [seriesRows, ticketRows, ticketEmailRows] = await Promise.all([
    listDbGroupTicketSeries(),
    listDbGroupTickets("", { limit: 5000 }),
    listDbGroupTicketEmailLinks(),
  ]);

  for (const series of seriesRows) {
    if (!series?.id) continue;
    store.groupTicketSeries[series.id] = normalizeGroupTicketSeriesInput(series, store.groupTicketSeries?.[series.id] || {});
  }

  for (const ticket of ticketRows) {
    if (!ticket?.id) continue;
    store.groupTickets[ticket.id] = normalizeGroupTicketInput(ticket, store.groupTickets?.[ticket.id] || {});
  }

  for (const row of ticketEmailRows) {
    ensureTicketEmailLink(store, row.ticketId, row.emailKey);
  }
}

async function upsertDbCustomGroup(group) {
  if (!db.isEnabled()) return;
  await db.query(
    `INSERT INTO crm_custom_groups (id, name, description, status, is_archived, archived_at, labels_json, documents_enabled, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7::jsonb, $8, $9, $10)
     ON CONFLICT (id) DO UPDATE SET
       name = EXCLUDED.name,
       description = EXCLUDED.description,
       status = EXCLUDED.status,
       is_archived = EXCLUDED.is_archived,
       archived_at = EXCLUDED.archived_at,
       labels_json = EXCLUDED.labels_json,
       documents_enabled = EXCLUDED.documents_enabled,
       updated_at = EXCLUDED.updated_at`,
    [
      normalizeString(group?.id),
      normalizeString(group?.name) || "Grupo sem nome",
      normalizeString(group?.description),
      normalizeGroupStatus(group?.status),
      group?.isArchived === true,
      group?.isArchived === true ? normalizeString(group?.archivedAt) || nowIso() : null,
      JSON.stringify(normalizeGroupLabels(group?.labels)),
      group?.documentsEnabled !== false,
      normalizeString(group?.createdAt) || nowIso(),
      normalizeString(group?.updatedAt) || nowIso(),
    ]
  );
}

async function upsertDbGroupDocument(groupId, input) {
  if (!db.isEnabled()) return;
  const gid = normalizeString(groupId);
  const doc = normalizeGroupDocumentInput(input);
  if (!gid || !doc.id || !doc.contentBase64) return;

  await db.query(
    `INSERT INTO crm_custom_group_documents
      (id, group_id, name, content_type, size_bytes, content_base64, source_email_key, source_item_id, source_internet_message_id, source_conversation_id, source_email_subject, storage_provider, storage_base_path, storage_path_hint, created_at, updated_at)
     VALUES
      ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16)
     ON CONFLICT (id) DO UPDATE SET
      group_id = EXCLUDED.group_id,
      name = EXCLUDED.name,
      content_type = EXCLUDED.content_type,
      size_bytes = EXCLUDED.size_bytes,
      content_base64 = EXCLUDED.content_base64,
      source_email_key = EXCLUDED.source_email_key,
      source_item_id = EXCLUDED.source_item_id,
      source_internet_message_id = EXCLUDED.source_internet_message_id,
      source_conversation_id = EXCLUDED.source_conversation_id,
      source_email_subject = EXCLUDED.source_email_subject,
      storage_provider = EXCLUDED.storage_provider,
      storage_base_path = EXCLUDED.storage_base_path,
      storage_path_hint = EXCLUDED.storage_path_hint,
      updated_at = EXCLUDED.updated_at`,
    [
      doc.id,
      gid,
      doc.name,
      doc.contentType,
      doc.size,
      doc.contentBase64,
      doc.sourceEmailKey,
      doc.sourceItemId,
      doc.sourceInternetMessageId,
      doc.sourceConversationId,
      doc.sourceEmailSubject,
      doc.storageProvider,
      doc.storageBasePath,
      doc.storagePathHint,
      doc.createdAt,
      doc.updatedAt,
    ]
  );
}

async function upsertDbGroupAttachmentFlag(groupId, input) {
  if (!db.isEnabled()) return;
  const gid = normalizeString(groupId);
  const flag = normalizeAttachmentFlagInput(input);
  if (!gid || !flag.attachmentKey) return;

  await db.query(
    `INSERT INTO crm_custom_group_attachment_flags
      (group_id, attachment_key, email_key, attachment_name, content_type, size_bytes, disposition, created_at, updated_at)
     VALUES
      ($1, $2, $3, $4, $5, $6, $7, $8, $9)
     ON CONFLICT (group_id, attachment_key) DO UPDATE SET
      email_key = EXCLUDED.email_key,
      attachment_name = EXCLUDED.attachment_name,
      content_type = EXCLUDED.content_type,
      size_bytes = EXCLUDED.size_bytes,
      disposition = EXCLUDED.disposition,
      updated_at = EXCLUDED.updated_at`,
    [
      gid,
      flag.attachmentKey,
      flag.emailKey,
      flag.attachmentName,
      flag.contentType,
      flag.size,
      flag.disposition,
      flag.createdAt,
      flag.updatedAt,
    ]
  );
}

async function listDbGroupAttachmentFlags(groupId) {
  if (!db.isEnabled()) return [];
  const gid = normalizeString(groupId);
  if (!gid) return [];
  const result = await db.query(
    `SELECT * FROM crm_custom_group_attachment_flags WHERE group_id = $1 ORDER BY updated_at DESC, created_at DESC`,
    [gid]
  );
  return (result?.rows || []).map(mapDbGroupAttachmentFlagRow).filter(Boolean);
}

async function listDbGroupDocuments(groupId) {
  if (!db.isEnabled()) return [];
  const gid = normalizeString(groupId);
  if (!gid) return [];
  const result = await db.query(
    `SELECT * FROM crm_custom_group_documents WHERE group_id = $1 ORDER BY updated_at DESC, created_at DESC`,
    [gid]
  );
  return (result?.rows || []).map(mapDbGroupDocumentRow).filter(Boolean);
}

async function deleteDbGroupDocument(groupId, documentId) {
  if (!db.isEnabled()) return;
  const gid = normalizeString(groupId);
  const did = normalizeString(documentId);
  if (!gid || !did) return;
  await db.query(`DELETE FROM crm_custom_group_documents WHERE group_id = $1 AND id = $2`, [gid, did]);
}

async function upsertDbCustomGroupMember(groupId, email, membershipKind = DEFAULT_GROUP_MEMBERSHIP_KIND) {
  if (!db.isEnabled()) return;
  const emailKey = makePersistentEmailKey(email);
  if (!groupId || !emailKey) return;
  await db.query(
    `INSERT INTO crm_custom_group_members
       (group_id, email_key, relation_kind, item_id, internet_message_id, conversation_id, subject, from_email, from_name, email_web_link, message_date_iso, received_at_iso, sent_at_iso, body_text, body_html, attachments_json, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16::jsonb, $17, $18)
     ON CONFLICT (group_id, email_key) DO UPDATE SET
       relation_kind = EXCLUDED.relation_kind,
       item_id = EXCLUDED.item_id,
       internet_message_id = EXCLUDED.internet_message_id,
       conversation_id = EXCLUDED.conversation_id,
       subject = EXCLUDED.subject,
       from_email = EXCLUDED.from_email,
       from_name = EXCLUDED.from_name,
       email_web_link = EXCLUDED.email_web_link,
       message_date_iso = EXCLUDED.message_date_iso,
       received_at_iso = EXCLUDED.received_at_iso,
       sent_at_iso = EXCLUDED.sent_at_iso,
       body_text = EXCLUDED.body_text,
       body_html = EXCLUDED.body_html,
       attachments_json = EXCLUDED.attachments_json,
       updated_at = EXCLUDED.updated_at`,
    [
      normalizeString(groupId),
      emailKey,
      normalizeGroupMembershipKind(membershipKind),
      normalizeString(email?.itemId),
      normalizeMessageId(email?.internetMessageId),
      normalizeString(email?.conversationId),
      normalizeString(email?.subject),
      normalizeString(email?.fromEmail),
      normalizeString(email?.fromName),
      normalizeString(email?.emailWebLink),
      normalizeString(email?.messageDateIso),
      normalizeString(email?.receivedAtIso),
      normalizeString(email?.sentAtIso),
      normalizeString(email?.bodyText),
      normalizeString(email?.bodyHtml),
      JSON.stringify(normalizeAttachments(email?.attachments)),
      normalizeString(email?.createdAt) || nowIso(),
      normalizeString(email?.updatedAt) || nowIso(),
    ]
  );
}

async function removeDbCustomGroupMember(groupId, identity = {}) {
  if (!db.isEnabled()) return;
  const gid = normalizeString(groupId);
  const key = normalizeString(identity?.emailKey);
  if (!gid) return;
  if (key) {
    await db.query(`DELETE FROM crm_custom_group_members WHERE group_id = $1 AND email_key = $2`, [gid, key]);
    return;
  }

  const normalized = normalizeEmailInput(identity);
  const emailKey = makePersistentEmailKey(normalized);
  if (emailKey) {
    await db.query(`DELETE FROM crm_custom_group_members WHERE group_id = $1 AND email_key = $2`, [gid, emailKey]);
    return;
  }

  if (normalized.itemId) {
    await db.query(`DELETE FROM crm_custom_group_members WHERE group_id = $1 AND item_id = $2`, [gid, normalized.itemId]);
    return;
  }

  if (normalized.internetMessageId) {
    await db.query(
      `DELETE FROM crm_custom_group_members
       WHERE group_id = $1
         AND LOWER(REGEXP_REPLACE(COALESCE(internet_message_id, ''), '[<>[:space:]]', '', 'g')) = $2`,
      [gid, normalized.internetMessageId]
    );
    return;
  }
}

async function getDbCustomGroupById(groupId) {
  if (!db.isEnabled()) return null;
  const gid = normalizeString(groupId);
  if (!gid) return null;
  const result = await db.query(
    `SELECT id, name, description, status, is_archived, archived_at, labels_json, documents_enabled, created_at, updated_at
     FROM crm_custom_groups
     WHERE id = $1`,
    [gid]
  );
  return mapDbGroupRow(result?.rows?.[0]);
}

async function listDbCustomGroups(query = "") {
  if (!db.isEnabled()) return [];
  const q = normalizeString(query).toLowerCase();
  const listAllAlphabetically = q === "/" || q === "*";
  const params = [];
  const where = [];
  if (q && !listAllAlphabetically) {
    params.push(`%${q}%`);
    where.push(`(LOWER(name) LIKE $${params.length} OR LOWER(COALESCE(description, '')) LIKE $${params.length})`);
  }
  const result = await db.query(
    `SELECT g.id, g.name, g.description, g.status, g.is_archived, g.archived_at, g.labels_json, g.documents_enabled, g.created_at, g.updated_at, COUNT(m.email_key)::int AS member_count
     FROM crm_custom_groups g
     LEFT JOIN crm_custom_group_members m ON m.group_id = g.id
     ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
     GROUP BY g.id, g.name, g.description, g.status, g.is_archived, g.archived_at, g.labels_json, g.documents_enabled, g.created_at, g.updated_at
     ORDER BY LOWER(g.name) ASC, g.created_at ASC`,
    params
  );
  return (result?.rows || []).map((row) => ({
    ...mapDbGroupRow(row),
    memberCount: Number(row.member_count || 0),
  }));
}

async function listDbEmailsByGroup(groupId) {
  if (!db.isEnabled()) return [];
  const gid = normalizeString(groupId);
  if (!gid) return [];
  const result = await db.query(
    `SELECT m.*, g.name AS group_name
     FROM crm_custom_group_members m
     JOIN crm_custom_groups g ON g.id = m.group_id
     WHERE m.group_id = $1
     ORDER BY
       COALESCE(NULLIF(m.message_date_iso, ''), NULLIF(m.received_at_iso, ''), NULLIF(m.sent_at_iso, ''), '') DESC,
       m.updated_at DESC,
       m.created_at DESC`,
    [gid]
  );
  return dedupeEmailLinks((result?.rows || []).map((row) =>
    buildEmailListEntry(mapDbGroupMemberRow(row), {
      relatedRecords: [],
      groupId: gid,
      groupName: normalizeString(row.group_name),
      membershipKind: normalizeGroupMembershipKind(row.relation_kind),
    })
  ));
}

async function getDbCustomGroupContext(input) {
  if (!db.isEnabled()) return { groups: [], emails: [], tickets: [] };

  const normalized = normalizeEmailInput(input);
  const emailKey = makePersistentEmailKey(normalized);
  const params = [];
  const where = [];

  if (normalizeString(normalized.itemId)) {
    params.push(normalizeString(normalized.itemId));
    where.push(`m.item_id = $${params.length}`);
  }
  if (normalizeString(normalized.internetMessageId)) {
    params.push(normalizeMessageId(normalized.internetMessageId));
    where.push(`LOWER(REGEXP_REPLACE(COALESCE(m.internet_message_id, ''), '[<>[:space:]]', '', 'g')) = $${params.length}`);
  }
  if (normalizeString(normalized.conversationId)) {
    params.push(normalizeString(normalized.conversationId));
    where.push(`m.conversation_id = $${params.length}`);
  }
  if (emailKey) {
    params.push(emailKey);
    where.push(`m.email_key = $${params.length}`);
  }
  if (!where.length) return { groups: [], emails: [], tickets: [] };

  const currentMemberships = await db.query(
    `SELECT DISTINCT
        g.id,
        g.name,
        g.description,
        g.status,
        g.is_archived,
        g.archived_at,
        g.labels_json,
        g.documents_enabled,
        g.created_at,
        g.updated_at,
        m.email_key,
        m.relation_kind,
        m.item_id,
        m.internet_message_id,
        m.conversation_id,
        m.subject,
        m.from_email,
        m.from_name,
        m.email_web_link,
        m.message_date_iso,
        m.received_at_iso,
        m.sent_at_iso,
        m.body_text,
        m.body_html,
        m.attachments_json
     FROM crm_custom_group_members m
     JOIN crm_custom_groups g ON g.id = m.group_id
     WHERE ${where.join(" OR ")}`,
    params
  );

  const membershipRows = currentMemberships?.rows || [];
  if (!membershipRows.length) {
    const tickets = emailKey
      ? (await listDbGroupTickets("", { emailKey, limit: 100 })).map((ticket) =>
          buildGroupTicketEntry({ groups: {}, groupTicketSeries: {}, groupTicketEmails: {} }, ticket, { emailLinked: true })
        )
      : [];
    return { groups: [], emails: [], tickets };
  }

  const currentEmailKeys = new Set(membershipRows.map((row) => normalizeString(row.email_key)).filter(Boolean));
  const primaryEmailKey = emailKey || Array.from(currentEmailKeys)[0] || "";
  const primaryMembershipRows = membershipRows.filter((row) => normalizeString(row.email_key) === primaryEmailKey);
  const groups = Array.from(
    new Map(membershipRows.map((row) => [normalizeString(row.id), {
      ...mapDbGroupRow(row),
      relationKind: normalizeGroupMembershipKind(row.relation_kind),
    }])).values()
  ).filter(Boolean);
  const currentRelatedGroups = primaryMembershipRows.reduce((acc, row) => {
    const groupId = normalizeString(row.id || row.group_id);
    if (!groupId || acc.some((entry) => entry.id === groupId)) return acc;
    acc.push({
      id: groupId,
      name: normalizeString(row.name || row.group_name),
      kind: "group",
      relationKind: normalizeGroupMembershipKind(row.relation_kind),
    });
    return acc;
  }, []);
  const currentPrincipalGroup = currentRelatedGroups.find((entry) => normalizeGroupMembershipKind(entry?.relationKind) === "principal") || currentRelatedGroups[0] || null;
  const currentEmail = primaryMembershipRows[0]
    ? buildEmailListEntry(mapDbGroupMemberRow(primaryMembershipRows[0]), {
      groupId: currentPrincipalGroup?.id || "",
      groupName: currentPrincipalGroup?.name || "",
      membershipKind: currentPrincipalGroup?.relationKind || normalizeGroupMembershipKind(primaryMembershipRows[0]?.relation_kind),
      relatedGroups: currentRelatedGroups,
      relatedRecords: [],
      relatedReasons: [],
    })
    : null;

  const groupIds = groups.map((group) => group.id).filter(Boolean);
  const tickets = primaryEmailKey
    ? (await listDbGroupTickets("", { emailKey: primaryEmailKey, limit: 100 })).map((ticket) => buildGroupTicketEntry({ groups: Object.fromEntries(groups.map((group) => [group.id, group])), groupTicketSeries: {}, groupTicketEmails: {} }, ticket, { emailLinked: true }))
    : [];
  if (!groupIds.length) return { groups, emails: [], tickets };

  const relatedRows = await db.query(
    `SELECT m.*, g.name AS group_name
     FROM crm_custom_group_members m
     JOIN crm_custom_groups g ON g.id = m.group_id
     WHERE m.group_id = ANY($1::text[])`,
    [groupIds]
  );

  const aggregated = new Map();
  for (const row of relatedRows?.rows || []) {
    const rowEmailKey = normalizeString(row.email_key);
    if (!rowEmailKey || currentEmailKeys.has(rowEmailKey)) continue;

    const email = mapDbGroupMemberRow(row);
    const current = aggregated.get(rowEmailKey) || {
      ...email,
      relatedRecords: [],
      relatedGroups: [],
      relatedReasons: [],
    };

    if (!current.relatedGroups.some((entry) => entry.id === row.group_id)) {
      current.relatedGroups.push({
        id: normalizeString(row.group_id),
        name: normalizeString(row.group_name),
        kind: "group",
        relationKind: normalizeGroupMembershipKind(row.relation_kind),
      });
    }

    const reason = {
      kind: "group",
      groupId: normalizeString(row.group_id),
      groupName: normalizeString(row.group_name),
      conversationId: normalizeString(row.conversation_id),
      relationKind: normalizeGroupMembershipKind(row.relation_kind),
    };
    const reasonKey = JSON.stringify(reason);
    if (!current.relatedReasons.some((entry) => JSON.stringify(entry) === reasonKey)) {
      current.relatedReasons.push(reason);
    }

    aggregated.set(rowEmailKey, current);
  }

  return {
    email: currentEmail,
    groups: groups.map((group) => ({ ...group, memberCount: 0 })),
    tickets,
    emails: Array.from(aggregated.values()).sort((a, b) =>
      String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || ""))
    ),
  };
}

async function ensureCustomGroupDb() {
  if (!db.isEnabled()) return;
  if (customGroupDbInitPromise) return customGroupDbInitPromise;

  customGroupDbInitPromise = (async () => {
    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_custom_groups (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        description TEXT DEFAULT '',
        status TEXT NOT NULL DEFAULT 'em_analise',
        is_archived BOOLEAN NOT NULL DEFAULT FALSE,
        archived_at TIMESTAMP NULL,
        labels_json JSONB NOT NULL DEFAULT '[]'::jsonb,
        documents_enabled BOOLEAN NOT NULL DEFAULT TRUE,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);

    await db.query(`
      ALTER TABLE crm_custom_groups
      ADD COLUMN IF NOT EXISTS documents_enabled BOOLEAN NOT NULL DEFAULT TRUE;
    `);
    await db.query(`
      ALTER TABLE crm_custom_groups
      ADD COLUMN IF NOT EXISTS status TEXT NOT NULL DEFAULT 'em_analise';
    `);
    await db.query(`
      ALTER TABLE crm_custom_groups
      ADD COLUMN IF NOT EXISTS is_archived BOOLEAN NOT NULL DEFAULT FALSE;
    `);
    await db.query(`
      ALTER TABLE crm_custom_groups
      ADD COLUMN IF NOT EXISTS archived_at TIMESTAMP NULL;
    `);
    await db.query(`
      ALTER TABLE crm_custom_groups
      ADD COLUMN IF NOT EXISTS labels_json JSONB NOT NULL DEFAULT '[]'::jsonb;
    `);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_custom_group_members (
        group_id TEXT NOT NULL REFERENCES crm_custom_groups(id) ON DELETE CASCADE,
        email_key TEXT NOT NULL,
        relation_kind TEXT NOT NULL DEFAULT 'principal',
        item_id TEXT,
        internet_message_id TEXT,
        conversation_id TEXT,
        subject TEXT,
        from_email TEXT,
        from_name TEXT,
        email_web_link TEXT,
        message_date_iso TEXT,
        received_at_iso TEXT,
        sent_at_iso TEXT,
        body_text TEXT,
        body_html TEXT,
        attachments_json JSONB DEFAULT '[]'::jsonb,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (group_id, email_key)
      );
    `);

    await db.query(`
      ALTER TABLE crm_custom_group_members
      ADD COLUMN IF NOT EXISTS attachments_json JSONB DEFAULT '[]'::jsonb;
    `);

    await db.query(`
      ALTER TABLE crm_custom_group_members
      ADD COLUMN IF NOT EXISTS body_text TEXT;
    `);

    await db.query(`
      ALTER TABLE crm_custom_group_members
      ADD COLUMN IF NOT EXISTS body_html TEXT;
    `);
    await db.query(`
      ALTER TABLE crm_custom_group_members
      ADD COLUMN IF NOT EXISTS relation_kind TEXT NOT NULL DEFAULT 'principal';
    `);

    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_custom_group_members_item_id ON crm_custom_group_members (item_id);`);
    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_custom_group_members_conversation_id ON crm_custom_group_members (conversation_id);`);
    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_custom_group_members_internet_message_id ON crm_custom_group_members (internet_message_id);`);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_custom_group_documents (
        id TEXT PRIMARY KEY,
        group_id TEXT NOT NULL REFERENCES crm_custom_groups(id) ON DELETE CASCADE,
        name TEXT NOT NULL,
        content_type TEXT,
        size_bytes INTEGER DEFAULT 0,
        content_base64 TEXT NOT NULL,
        source_email_key TEXT,
        source_item_id TEXT,
        source_internet_message_id TEXT,
        source_conversation_id TEXT,
        source_email_subject TEXT,
        storage_provider TEXT DEFAULT '',
        storage_base_path TEXT DEFAULT '',
        storage_path_hint TEXT DEFAULT '',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);

    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_custom_group_documents_group_id ON crm_custom_group_documents (group_id);`);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_custom_group_attachment_flags (
        group_id TEXT NOT NULL REFERENCES crm_custom_groups(id) ON DELETE CASCADE,
        attachment_key TEXT NOT NULL,
        email_key TEXT,
        attachment_name TEXT NOT NULL,
        content_type TEXT,
        size_bytes INTEGER DEFAULT 0,
        disposition TEXT NOT NULL DEFAULT 'dismissed',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (group_id, attachment_key)
      );
    `);

    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_custom_group_attachment_flags_group_id ON crm_custom_group_attachment_flags (group_id);`);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_group_ticket_series (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        prefix TEXT NOT NULL,
        reply_instructions TEXT NOT NULL DEFAULT '',
        year_mode TEXT NOT NULL DEFAULT 'none',
        separator TEXT NOT NULL DEFAULT '-',
        next_number INTEGER NOT NULL DEFAULT 1,
        padding INTEGER NOT NULL DEFAULT 4,
        is_active BOOLEAN NOT NULL DEFAULT TRUE,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_group_tickets (
        id TEXT PRIMARY KEY,
        series_id TEXT NOT NULL REFERENCES crm_group_ticket_series(id) ON DELETE RESTRICT,
        prefix TEXT NOT NULL DEFAULT '',
        year_mode TEXT NOT NULL DEFAULT 'none',
        separator TEXT NOT NULL DEFAULT '-',
        year_value TEXT NOT NULL DEFAULT '',
        code TEXT NOT NULL UNIQUE,
        sequence_number INTEGER NOT NULL,
        padding INTEGER NOT NULL DEFAULT 4,
        title TEXT NOT NULL,
        description TEXT DEFAULT '',
        status TEXT NOT NULL DEFAULT 'open',
        labels_json JSONB NOT NULL DEFAULT '[]'::jsonb,
        group_ids_json JSONB NOT NULL DEFAULT '[]'::jsonb,
        created_from_email_key TEXT DEFAULT '',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);

    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_group_tickets_series_id ON crm_group_tickets (series_id);`);
    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_group_tickets_code ON crm_group_tickets (code);`);
    await db.query(`ALTER TABLE crm_group_ticket_series ADD COLUMN IF NOT EXISTS reply_instructions TEXT NOT NULL DEFAULT '';`);
    await db.query(`ALTER TABLE crm_group_ticket_series ADD COLUMN IF NOT EXISTS year_mode TEXT NOT NULL DEFAULT 'none';`);
    await db.query(`ALTER TABLE crm_group_ticket_series ADD COLUMN IF NOT EXISTS separator TEXT NOT NULL DEFAULT '-';`);
    await db.query(`ALTER TABLE crm_group_tickets ADD COLUMN IF NOT EXISTS prefix TEXT NOT NULL DEFAULT '';`);
    await db.query(`ALTER TABLE crm_group_tickets ADD COLUMN IF NOT EXISTS year_mode TEXT NOT NULL DEFAULT 'none';`);
    await db.query(`ALTER TABLE crm_group_tickets ADD COLUMN IF NOT EXISTS separator TEXT NOT NULL DEFAULT '-';`);
    await db.query(`ALTER TABLE crm_group_tickets ADD COLUMN IF NOT EXISTS year_value TEXT NOT NULL DEFAULT '';`);

    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_group_ticket_emails (
        ticket_id TEXT NOT NULL REFERENCES crm_group_tickets(id) ON DELETE CASCADE,
        email_key TEXT NOT NULL,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (ticket_id, email_key)
      );
    `);

    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_group_ticket_emails_email_key ON crm_group_ticket_emails (email_key);`);

    const store = readState();
    const customGroups = Object.values(store.groups || {}).filter((group) => group?.kind === CUSTOM_GROUP_KIND);
    for (const group of customGroups) {
      await upsertDbCustomGroup(group);
      for (const emailId of store.groupMembers[group.id] || []) {
        const email = buildRecoveredEmailSnapshot(store, emailId);
        if (email) await upsertDbCustomGroupMember(group.id, email, getEmailMembershipMeta(store, group.id, emailId).kind);
      }
      for (const doc of store.groupDocuments[group.id] || []) {
        await upsertDbGroupDocument(group.id, doc);
      }
      for (const flag of store.groupAttachmentFlags[group.id] || []) {
        await upsertDbGroupAttachmentFlag(group.id, flag);
      }
    }
    for (const series of Object.values(store.groupTicketSeries || {})) {
      await upsertDbGroupTicketSeries(series);
    }
    for (const ticket of Object.values(store.groupTickets || {})) {
      await upsertDbGroupTicket(ticket);
      for (const emailKey of store.groupTicketEmails?.[ticket.id] || []) {
        await upsertDbGroupTicketEmail(ticket.id, emailKey);
      }
    }
  })().catch((error) => {
    customGroupDbInitPromise = null;
    throw error;
  });

  return customGroupDbInitPromise;
}

function getMatchedEmailIds(store, lookup) {
  const normalized = normalizeEmailInput(lookup);
  if (normalized.itemId && store.indexes.itemIds[normalized.itemId]) {
    return [store.indexes.itemIds[normalized.itemId]];
  }
  if (normalized.internetMessageId && store.indexes.internetMessageIds[normalized.internetMessageId]) {
    return [store.indexes.internetMessageIds[normalized.internetMessageId]];
  }
  const fingerprint = makeEmailFingerprint(normalized);
  if (fingerprint && store.indexes.fingerprints[fingerprint]) {
    return [store.indexes.fingerprints[fingerprint]];
  }
  if (normalized.conversationId) {
    return Array.isArray(store.indexes.conversations[normalized.conversationId])
      ? store.indexes.conversations[normalized.conversationId]
      : [];
  }
  return [];
}

function readState() {
  const candidates = uniqueFilePaths()
    .map((filePath) => ({ filePath, raw: readRawFile(filePath) }))
    .filter((entry) => entry.raw && typeof entry.raw === "object");

  const firstCanonical = candidates.find((entry) => entry.raw?.version === STORE_VERSION && entry.raw?.emails);
  if (firstCanonical) {
    const canonicalStore = hydrateStore(firstCanonical.raw);
    const hasPrimaryContent = Object.keys(canonicalStore.emails || {}).length
      || Object.keys(canonicalStore.groups || {}).length
      || Object.keys(canonicalStore.entityLinks || {}).length;
    const legacyCandidate = candidates.find((entry) => entry.filePath !== firstCanonical.filePath);
    if (!hasPrimaryContent && legacyCandidate && legacyCandidate.raw?.version !== STORE_VERSION) {
      const migrated = migrateLegacyStore(legacyCandidate.raw);
      writeStore(migrated);
      return migrated;
    }
    if (firstCanonical.filePath !== PRIMARY_FILE_PATH) writeStore(canonicalStore);
    return canonicalStore;
  }

  for (const candidate of candidates) {
    const migrated = migrateLegacyStore(candidate.raw);
    writeStore(migrated);
    return migrated;
  }

  const empty = createEmptyStore();
  writeStore(empty);
  return empty;
}

function migrateLegacyStore(raw) {
  const store = createEmptyStore();
  const input = raw && typeof raw === "object" ? raw : {};
  for (const [key, value] of Object.entries(input)) {
    if (!Array.isArray(value)) continue;
    const parsed = parseLegacyStorageKey(key);
    for (const entry of value) {
      const email = upsertEmail(store, {
        conversationId: entry?.conversationId || parsed.conversationId,
        internetMessageId: entry?.internetMessageId || parsed.internetMessageId,
        itemId: entry?.itemId,
        emailWebLink: entry?.emailWebLink,
        subject: entry?.subject,
        fromEmail: entry?.fromEmail,
        fromName: entry?.fromName,
        receivedAtIso: entry?.receivedAtIso,
        linkedAt: entry?.linkedAt,
      });
      if (email?.conversationId) {
        const group = ensureConversationGroup(store, email.conversationId, email);
        if (group) addEmailMembership(store, group.id, email.id);
      }
      linkEmailToEntity(store, email.id, {
        ...entry,
        conversationId: entry?.conversationId || parsed.conversationId,
        internetMessageId: entry?.internetMessageId || parsed.internetMessageId,
      });
    }
  }
  return store;
}

function listLinksFromFile(conversationId, internetMessageId = "", itemId = "") {
  const store = readState();
  const matchedIds = getMatchedEmailIds(store, { conversationId, internetMessageId, itemId });
  const links = matchedIds.flatMap((emailId) => {
    const email = store.emails[emailId];
    const entityLinks = Array.isArray(store.emailEntityLinks[emailId]) ? store.emailEntityLinks[emailId] : [];
    return entityLinks.map((entityLink) => buildLinkEntry(email, entityLink));
  });
  return dedupeRecordLinks(links);
}

function listLinksByRecordFromFile(model, recordId) {
  const store = readState();
  const entityKey = makeEntityKey(model, recordId);
  if (!entityKey) return [];
  const emailIds = Array.isArray(store.entityLinks[entityKey]) ? store.entityLinks[entityKey] : [];
  const rows = emailIds.map((emailId) => {
    const email = store.emails[emailId];
    const entityLink = (store.emailEntityLinks[emailId] || []).find(
      (entry) => makeEntityKey(entry?.model, entry?.recordId) === entityKey
    );
    return buildLinkEntry(email, entityLink);
  });
  return dedupeEmailLinks(rows);
}

function writeStateWithEmail(store, email) {
  if (email?.conversationId) {
    const group = ensureConversationGroup(store, email.conversationId, email);
    if (group) addEmailMembership(store, group.id, email.id);
  }
  writeCacheStore(store);
}

export async function registerRelevantEmail(input) {
  const store = readState();
  const email = upsertEmail(store, input);
  writeStateWithEmail(store, email);
  return buildEmailListEntry(email, {
    groups: (store.emailGroups[email.id] || []).map((groupId) => store.groups[groupId]).filter(Boolean),
  });
}

export async function createCustomGroup(input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("criar o grupo");
  const group = ensureGroup(store, {
    kind: CUSTOM_GROUP_KIND,
    name: normalizeString(input?.name) || "Grupo sem nome",
    description: normalizeString(input?.description),
    status: normalizeGroupStatus(input?.status),
    labels: normalizeGroupLabels(input?.labels),
    isArchived: typeof input?.isArchived === "boolean" ? input.isArchived : false,
    archivedAt: normalizeString(input?.archivedAt),
    documentsEnabled: typeof input?.documentsEnabled === "boolean" ? input.documentsEnabled : true,
  });
  if (useDurableDb) {
    try {
      await upsertDbCustomGroup(group);
    } catch (error) {
      throw durablePersistenceError("criar o grupo", error);
    }
  }
  writeCacheStore(store);
  return buildGroupListEntry(store, group);
}

export async function updateCustomGroup(groupId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("atualizar o grupo");
  const gid = normalizeString(groupId);
  const current = store.groups[gid];
  if (!gid || !current || current.kind !== CUSTOM_GROUP_KIND) {
    throw new Error("Grupo inválido.");
  }

  const group = ensureGroup(store, {
    ...current,
    id: gid,
    name: normalizeString(input?.name) || current.name,
    description:
      Object.prototype.hasOwnProperty.call(input || {}, "description")
        ? normalizeString(input?.description)
        : current.description,
    status:
      Object.prototype.hasOwnProperty.call(input || {}, "status")
        ? normalizeGroupStatus(input?.status)
        : current.status,
    labels:
      Object.prototype.hasOwnProperty.call(input || {}, "labels")
        ? normalizeGroupLabels(input?.labels)
        : current.labels,
    isArchived:
      typeof input?.isArchived === "boolean"
        ? input.isArchived
        : current.isArchived === true,
    archivedAt:
      Object.prototype.hasOwnProperty.call(input || {}, "archivedAt")
        ? normalizeString(input?.archivedAt)
        : current.archivedAt,
    documentsEnabled:
      typeof input?.documentsEnabled === "boolean"
        ? input.documentsEnabled
        : current.documentsEnabled,
  });
  if (useDurableDb) {
    try {
      await upsertDbCustomGroup(group);
    } catch (error) {
      throw durablePersistenceError("atualizar o grupo", error);
    }
  }
  writeCacheStore(store);
  return buildGroupListEntry(store, group);
}

export async function listCustomGroups(query = "") {
  const store = readState();
  const q = normalizeString(query).toLowerCase();
  const listAllAlphabetically = q === "/" || q === "*";
  const fileGroups = Object.values(store.groups)
    .filter((group) => group?.kind === "custom")
    .filter((group) => {
      if (!q || listAllAlphabetically) return true;
      return String(group?.name || "").toLowerCase().includes(q)
        || String(group?.description || "").toLowerCase().includes(q);
    })
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "pt"))
    .map((group) => buildGroupListEntry(store, group))
    .filter(Boolean);

  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("carregar os grupos");
  }
  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbGroups = await listDbCustomGroups(query);
      const merged = new Map(fileGroups.map((group) => [group.id, group]));
      for (const group of dbGroups) {
        merged.set(group.id, {
          ...group,
          memberCount: Math.max(Number(group.memberCount || 0), Number(merged.get(group.id)?.memberCount || 0)),
          documentCount: Math.max(Number(group.documentCount || 0), Number(merged.get(group.id)?.documentCount || 0)),
        });
      }
      return Array.from(merged.values()).sort((a, b) =>
        String(a?.name || "").localeCompare(String(b?.name || ""), "pt")
      );
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("carregar os grupos", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Custom Group Query Error, falling back to central file store:", error);
    }
  }

  return fileGroups;
}

export async function addEmailToGroup(groupId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("ligar o email ao grupo");
  let existingDbGroup = null;
  if (useDurableDb) {
    try {
      existingDbGroup = await getDbCustomGroupById(groupId);
    } catch (error) {
      throw durablePersistenceError("ligar o email ao grupo", error);
    }
  }

  const group = ensureGroup(store, {
    id: groupId,
    kind: CUSTOM_GROUP_KIND,
    name: existingDbGroup?.name,
    description: existingDbGroup?.description,
    status: existingDbGroup?.status,
    labels: existingDbGroup?.labels,
    isArchived: existingDbGroup?.isArchived,
    archivedAt: existingDbGroup?.archivedAt,
    documentsEnabled: existingDbGroup?.documentsEnabled,
  });
  const email = upsertEmail(store, input);
  const membershipKind = normalizeGroupMembershipKind(input?.membershipKind);
  if (email.conversationId) {
    const conversationGroup = ensureConversationGroup(store, email.conversationId, email);
    if (conversationGroup) addEmailMembership(store, conversationGroup.id, email.id, { membershipKind: DEFAULT_GROUP_MEMBERSHIP_KIND });
  }
  addEmailMembership(store, group.id, email.id, { membershipKind });
  if (useDurableDb) {
    try {
      await upsertDbCustomGroup(group);
      await upsertDbCustomGroupMember(group.id, email, membershipKind);
    } catch (error) {
      throw durablePersistenceError("ligar o email ao grupo", error);
    }
  }
  writeCacheStore(store);

  return {
    group: buildGroupListEntry(store, group),
    email: buildEmailListEntry(email, { membershipKind }),
  };
}

export async function removeEmailFromGroup(groupId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("remover o email do grupo");
  const gid = normalizeString(groupId);
  const emailKey = normalizeString(input?.emailKey);
  let emailId = "";

  if (emailKey) {
    emailId = Object.values(store.emails).find((email) => makePersistentEmailKey(email) === emailKey)?.id || "";
  }

  if (!emailId) {
    emailId = resolveEmailId(store, input);
  }

  let removed = false;
  if (emailId) {
    removed = removeEmailMembership(store, gid, emailId);
  } else if (emailKey && Array.isArray(store.groupMembers[gid])) {
    const match = store.groupMembers[gid].find((candidateEmailId) => {
      const email = store.emails[candidateEmailId];
      return makePersistentEmailKey(email) === emailKey;
    });
    if (match) removed = removeEmailMembership(store, gid, match);
  }

  if (useDurableDb) {
    try {
      await removeDbCustomGroupMember(groupId, input);
    } catch (error) {
      throw durablePersistenceError("remover o email do grupo", error);
    }
  }
  if (removed) {
    writeCacheStore(store);
  }

  return {
    ok: true,
    removed,
    groupId: gid,
    emailKey: emailKey || (emailId ? makePersistentEmailKey(store.emails[emailId]) : ""),
  };
}

export async function deleteCustomGroup(groupId) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("eliminar o grupo");
  const gid = normalizeString(groupId);
  if (!gid) return { ok: true, deleted: false };

  const memberIds = Array.isArray(store.groupMembers[gid]) ? [...store.groupMembers[gid]] : [];
  for (const emailId of memberIds) {
    removeEmailMembership(store, gid, emailId);
  }
  delete store.groupMembers[gid];
  delete store.groupMemberLinks[gid];
  delete store.groupDocuments[gid];
  delete store.groupAttachmentFlags[gid];
  for (const ticket of Object.values(store.groupTickets || {})) {
    if (!normalizeGroupIds(ticket?.groupIds).includes(gid)) continue;
    store.groupTickets[ticket.id] = normalizeGroupTicketInput({
      ...ticket,
      groupIds: normalizeGroupIds(ticket.groupIds).filter((entry) => entry !== gid),
      updatedAt: nowIso(),
    }, ticket);
  }
  delete store.groups[gid];
  if (useDurableDb) {
    try {
      for (const ticket of Object.values(store.groupTickets || {})) {
        if (!ticket?.id) continue;
        await upsertDbGroupTicket(ticket);
      }
      await db.query(`DELETE FROM crm_custom_groups WHERE id = $1`, [gid]);
    } catch (error) {
      throw durablePersistenceError("eliminar o grupo", error);
    }
  }
  writeCacheStore(store);

  return { ok: true, deleted: true, groupId: gid };
}

export async function listEmailsByGroup(groupId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const emailIds = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
  const fileRows = emailIds.map((emailId) => {
    const email = buildRecoveredEmailSnapshot(store, emailId);
    const relatedRecords = Array.isArray(store.emailEntityLinks[emailId])
      ? store.emailEntityLinks[emailId].map((entry) => ({
        model: entry.model,
        recordId: entry.recordId,
        recordName: entry.recordName,
      }))
      : [];
    return buildEmailListEntry(email, {
      relatedRecords,
      groupId: gid,
      groupName: store.groups[gid]?.name || "",
      membershipKind: getEmailMembershipMeta(store, gid, emailId).kind,
    });
  });

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbRows = await listDbEmailsByGroup(groupId);
      return dedupeEmailLinks([...fileRows, ...dbRows]);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Email Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Group Email Query Error, falling back to central file store:", error);
    }
  }

  return dedupeEmailLinks(fileRows);
}

export async function listKnownEmails(query = "", options = {}) {
  const store = readState();
  const q = normalizeString(query).toLowerCase();
  const excludeGroupId = normalizeString(options?.excludeGroupId);
  const limit = Math.max(1, Math.min(Number(options?.limit || 200) || 200, 500));
  const allEmailIds = new Set([
    ...Object.keys(store.emails || {}),
    ...Object.keys(store.emailEntityLinks || {}),
    ...Object.values(store.groupMembers || {}).flatMap((value) => (Array.isArray(value) ? value : [])),
  ]);

  const rows = [];
  for (const emailId of allEmailIds) {
    const source = buildRecoveredEmailSnapshot(store, emailId);
    if (!source) continue;
    const relatedRecords = Array.isArray(store.emailEntityLinks[emailId])
      ? dedupeRecordLinks(store.emailEntityLinks[emailId]).map((entry) => ({
        model: entry.model,
        recordId: entry.recordId,
        recordName: entry.recordName,
      }))
      : [];
    const relatedGroups = listEmailGroupMemberships(store, emailId)
      .map((membership) => {
        const group = store.groups[membership.groupId];
        if (!group) return null;
        return {
          id: normalizeString(group.id),
          name: normalizeString(group.name),
          kind: normalizeString(group.kind),
          relationKind: membership.kind,
        };
      })
      .filter(Boolean);

    if (excludeGroupId && relatedGroups.some((group) => normalizeString(group.id) === excludeGroupId)) {
      continue;
    }

    const row = buildEmailSearchEntry(source, {
      relatedRecords,
      relatedGroups,
    });

    if (q) {
      const haystack = [
        row.subject,
        row.fromEmail,
        row.fromName,
        row.conversationId,
        ...(row.relatedGroups || []).flatMap((group) => [group.name, group.kind]),
        ...(row.relatedRecords || []).flatMap((record) => [record.recordName, record.model]),
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      if (!haystack.includes(q)) continue;
    }

    rows.push(row);
  }

  return dedupeEmailLinks(rows)
    .sort((a, b) =>
      String(b.messageDateIso || b.receivedAtIso || b.updatedAt || "").localeCompare(
        String(a.messageDateIso || a.receivedAtIso || a.updatedAt || "")
      )
    )
    .slice(0, limit);
}

export async function listDocumentsByGroup(groupId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const fileRows = Array.isArray(store.groupDocuments?.[gid])
    ? store.groupDocuments[gid].map((doc) => normalizeGroupDocumentInput(doc))
    : [];

  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("carregar os documentos do grupo");
  }
  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbRows = await listDbGroupDocuments(gid);
      const merged = new Map();
      for (const doc of [...fileRows, ...dbRows]) {
        if (!doc?.id) continue;
        merged.set(doc.id, doc);
      }
      return Array.from(merged.values()).sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("carregar os documentos do grupo", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Document Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Group Document Query Error, falling back to central file store:", error);
    }
  }

  return fileRows.sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
}

export async function listAttachmentFlagsByGroup(groupId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const fileRows = Array.isArray(store.groupAttachmentFlags?.[gid])
    ? store.groupAttachmentFlags[gid].map((entry) => normalizeAttachmentFlagInput(entry)).filter((entry) => entry.attachmentKey)
    : [];

  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("carregar a configuracao de anexos do grupo");
  }
  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbRows = await listDbGroupAttachmentFlags(gid);
      const merged = new Map();
      for (const flag of [...fileRows, ...dbRows]) {
        if (!flag?.attachmentKey) continue;
        merged.set(flag.attachmentKey, flag);
      }
      return Array.from(merged.values()).sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("carregar a configuracao de anexos do grupo", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Attachment Flag Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Attachment Flag Query Error, falling back to central file store:", error);
    }
  }

  return fileRows.sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
}

export async function saveAttachmentFlagsToGroup(groupId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("guardar a configuracao de anexos do grupo");
  const gid = normalizeString(groupId);
  const group = store.groups[gid];
  if (!gid || !group || group.kind !== CUSTOM_GROUP_KIND) throw new Error("Grupo inválido.");

  const entries = Array.isArray(input?.entries)
    ? input.entries.map((entry) => normalizeAttachmentFlagInput(entry)).filter((entry) => entry.attachmentKey)
    : [];
  if (!entries.length) return { ok: true, flags: await listAttachmentFlagsByGroup(gid) };

  const current = Array.isArray(store.groupAttachmentFlags[gid])
    ? store.groupAttachmentFlags[gid].map((entry) => normalizeAttachmentFlagInput(entry))
    : [];
  const byKey = new Map(current.map((entry) => [entry.attachmentKey, entry]));
  for (const entry of entries) {
    byKey.set(entry.attachmentKey, {
      ...byKey.get(entry.attachmentKey),
      ...entry,
      updatedAt: nowIso(),
      createdAt: byKey.get(entry.attachmentKey)?.createdAt || entry.createdAt || nowIso(),
    });
  }
  store.groupAttachmentFlags[gid] = Array.from(byKey.values()).sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
  if (store.groups[gid]) store.groups[gid].updatedAt = nowIso();
  if (useDurableDb) {
    try {
      for (const entry of entries) {
        await upsertDbGroupAttachmentFlag(gid, entry);
      }
    } catch (error) {
      throw durablePersistenceError("guardar a configuracao de anexos do grupo", error);
    }
  }
  writeCacheStore(store);

  return { ok: true, flags: await listAttachmentFlagsByGroup(gid) };
}

export async function saveDocumentsToGroup(groupId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("guardar documentos no grupo");
  const gid = normalizeString(groupId);
  const group = store.groups[gid];
  if (group?.documentsEnabled === false) throw new Error("A gestao documental esta desativada neste grupo.");
  if (!gid || !group || group.kind !== CUSTOM_GROUP_KIND) throw new Error("Grupo inválido.");

  const docs = Array.isArray(input?.documents) ? input.documents.map((doc) => normalizeGroupDocumentInput(doc)).filter((doc) => doc.contentBase64) : [];
  if (!docs.length) throw new Error("Sem documentos válidos para guardar.");

  const current = Array.isArray(store.groupDocuments[gid]) ? store.groupDocuments[gid].map((doc) => normalizeGroupDocumentInput(doc)) : [];
  const byId = new Map(current.map((doc) => [doc.id, doc]));
  for (const doc of docs) {
    byId.set(doc.id, {
      ...doc,
      updatedAt: nowIso(),
      createdAt: byId.get(doc.id)?.createdAt || doc.createdAt || nowIso(),
    });
  }
  store.groupDocuments[gid] = Array.from(byId.values()).sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
  if (store.groups[gid]) store.groups[gid].updatedAt = nowIso();
  if (useDurableDb) {
    try {
      for (const doc of docs) {
        await upsertDbGroupDocument(gid, doc);
      }
    } catch (error) {
      throw durablePersistenceError("guardar documentos no grupo", error);
    }
  }
  writeCacheStore(store);

  return {
    ok: true,
    group: store.groups[gid],
    documents: await listDocumentsByGroup(gid),
  };
}

export async function deleteDocumentFromGroup(groupId, documentId) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("eliminar o documento do grupo");
  const gid = normalizeString(groupId);
  const did = normalizeString(documentId);
  const current = Array.isArray(store.groupDocuments[gid]) ? store.groupDocuments[gid] : [];
  const next = current.filter((doc) => normalizeString(doc?.id) !== did);
  const removed = next.length !== current.length;
  store.groupDocuments[gid] = next;
  if (removed && store.groups[gid]) store.groups[gid].updatedAt = nowIso();
  if (useDurableDb && removed) {
    try {
      await deleteDbGroupDocument(gid, did);
    } catch (error) {
      throw durablePersistenceError("eliminar o documento do grupo", error);
    }
  }
  if (removed) writeCacheStore(store);

  return { ok: true, removed, groupId: gid, documentId: did };
}

function findSeriesConflict(store, prefix, excludeId = "") {
  const normalizedPrefix = normalizeTicketPrefix(prefix);
  return Object.values(store.groupTicketSeries || {}).find((series) =>
    normalizeTicketPrefix(series?.prefix) === normalizedPrefix && normalizeString(series?.id) !== normalizeString(excludeId)
  ) || null;
}

export async function listGroupTicketSeries() {
  const store = readState();
  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("carregar as series de tickets");
  }
  if (db.isEnabled()) {
    try {
      await syncGroupTicketsFromDb(store);
      writeCacheStore(store);
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("carregar as series de tickets", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Ticket Series Sync Error, using central file store:", error.message);
      else console.error("[linkStore] DB Ticket Series Sync Error, using central file store:", error);
    }
  }
  return Object.values(store.groupTicketSeries || {})
    .map((series) => buildGroupTicketSeriesEntry(store, series))
    .filter(Boolean)
    .sort((a, b) =>
      Number(b.isActive !== false) - Number(a.isActive !== false)
      || String(a.prefix || "").localeCompare(String(b.prefix || ""), "pt")
      || String(a.name || "").localeCompare(String(b.name || ""), "pt")
    );
}

export async function createGroupTicketSeries(input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("criar a serie de tickets", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const draft = normalizeGroupTicketSeriesInput(input);
  if (!draft.prefix) throw new Error("Define um prefixo para a serie.");
  const conflict = findSeriesConflict(store, draft.prefix);
  if (conflict) throw new Error(`Ja existe uma serie com o prefixo ${draft.prefix}.`);

  draft.updatedAt = nowIso();
  store.groupTicketSeries[draft.id] = draft;
  if (useDurableDb) {
    try {
      await upsertDbGroupTicketSeries(draft);
    } catch (error) {
      throw durablePersistenceError("criar a serie de tickets", error);
    }
  }
  writeCacheStore(store);

  return buildGroupTicketSeriesEntry(store, draft);
}

export async function updateGroupTicketSeries(seriesId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("atualizar a serie de tickets", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const sid = normalizeString(seriesId);
  const current = store.groupTicketSeries?.[sid];
  if (!sid || !current) throw new Error("Serie de ticket invalida.");

  const next = normalizeGroupTicketSeriesInput({ ...current, ...input, id: sid }, current);
  const conflict = findSeriesConflict(store, next.prefix, sid);
  if (conflict) throw new Error(`Ja existe uma serie com o prefixo ${next.prefix}.`);

  store.groupTicketSeries[sid] = { ...next, updatedAt: nowIso() };

  for (const ticket of Object.values(store.groupTickets || {})) {
    if (normalizeString(ticket?.seriesId) !== sid) continue;
    const yearValue = getTicketYearValue(ticket?.createdAt, next.yearMode);
    const merged = normalizeGroupTicketInput({
      ...ticket,
      prefix: next.prefix,
      yearMode: next.yearMode,
      separator: next.separator,
      yearValue,
      padding: next.padding,
      code: buildTicketCode(next.prefix, ticket.sequenceNumber, next.padding, {
        yearMode: next.yearMode,
        separator: next.separator,
        yearValue,
        dateValue: ticket?.createdAt,
      }),
      seriesName: next.name,
      updatedAt: nowIso(),
    }, ticket);
    store.groupTickets[merged.id] = merged;
  }
  if (useDurableDb) {
    try {
      await upsertDbGroupTicketSeries(store.groupTicketSeries[sid]);
      for (const ticket of Object.values(store.groupTickets || {})) {
        if (normalizeString(ticket?.seriesId) !== sid) continue;
        await upsertDbGroupTicket(ticket);
      }
    } catch (error) {
      throw durablePersistenceError("atualizar a serie de tickets", error);
    }
  }
  writeCacheStore(store);

  return buildGroupTicketSeriesEntry(store, store.groupTicketSeries[sid]);
}

export async function deleteGroupTicketSeries(seriesId) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("eliminar a serie de tickets", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });
  const sid = normalizeString(seriesId);
  if (!sid || !store.groupTicketSeries?.[sid]) throw new Error("Serie de ticket invalida.");
  const usageCount = Object.values(store.groupTickets || {}).filter((ticket) => normalizeString(ticket?.seriesId) === sid).length;
  if (usageCount > 0) throw new Error("Nao podes eliminar uma serie que ja tem tickets.");
  delete store.groupTicketSeries[sid];
  if (useDurableDb) {
    try {
      await deleteDbGroupTicketSeries(sid);
    } catch (error) {
      throw durablePersistenceError("eliminar a serie de tickets", error);
    }
  }
  writeCacheStore(store);

  return { ok: true, deleted: true, seriesId: sid };
}

export async function listGroupTickets(query = "", options = {}) {
  const store = readState();
  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("carregar os tickets");
  }
  if (db.isEnabled()) {
    try {
      await syncGroupTicketsFromDb(store);
      writeCacheStore(store);
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("carregar os tickets", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Ticket Sync Error, using central file store:", error.message);
      else console.error("[linkStore] DB Ticket Sync Error, using central file store:", error);
    }
  }

  const q = normalizeString(query).toLowerCase();
  const groupId = normalizeString(options?.groupId);
  const emailKey = resolveEmailKeyFromInput(store, options?.email || {});
  const limit = Math.max(1, Math.min(normalizePositiveInt(options?.limit || 50, 50), 200));

  return Object.values(store.groupTickets || {})
    .map((ticket) => buildGroupTicketEntry(store, ticket, {
      emailLinked: emailKey ? listTicketIdsByEmailKey(store, emailKey).includes(normalizeString(ticket.id)) : false,
    }))
    .filter((ticket) => {
      if (!ticket) return false;
      if (groupId && !normalizeGroupIds(ticket.groupIds).includes(groupId)) return false;
      if (emailKey && !listTicketIdsByEmailKey(store, emailKey).includes(normalizeString(ticket.id))) return false;
      if (!q) return true;
      const haystack = [
        ticket.code,
        ticket.seriesName,
        ticket.prefix,
        ticket.title,
        ticket.description,
        ticket.status,
        ...(ticket.labels || []),
        ...(ticket.groups || []).map((group) => group?.name),
      ]
        .filter(Boolean)
        .join(" ")
        .toLowerCase();
      return haystack.includes(q);
    })
    .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))
    .slice(0, limit);
}

export async function createGroupTicket(input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("criar o ticket", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const seriesId = normalizeString(input?.seriesId);
  const series = store.groupTicketSeries?.[seriesId];
  if (!seriesId || !series) throw new Error("Seleciona uma serie valida.");
  if (series.isActive === false) throw new Error("A serie selecionada esta desativada.");

  const sequenceNumber = normalizePositiveInt(series.nextNumber, 1);
  const createdAt = nowIso();
  const yearValue = getTicketYearValue(createdAt, series.yearMode);
  const code = buildTicketCode(series.prefix, sequenceNumber, series.padding, {
    yearMode: series.yearMode,
    separator: series.separator,
    yearValue,
    dateValue: createdAt,
  });
  const ticket = normalizeGroupTicketInput({
    seriesId,
    seriesName: series.name,
    prefix: series.prefix,
    yearMode: series.yearMode,
    separator: series.separator,
    yearValue,
    padding: series.padding,
    sequenceNumber,
    code,
    title: normalizeString(input?.title) || code,
    description: normalizeString(input?.description),
    labels: normalizeGroupLabels(input?.labels),
    groupIds: normalizeGroupIds(input?.groupIds).filter((groupId) => Boolean(store.groups?.[groupId])),
    createdFromEmailKey: resolveEmailKeyFromInput(store, input?.email || {}),
    status: DEFAULT_GROUP_TICKET_STATUS,
    createdAt,
  });

  store.groupTickets[ticket.id] = ticket;
  store.groupTicketSeries[seriesId] = {
    ...series,
    nextNumber: sequenceNumber + 1,
    updatedAt: nowIso(),
  };

  if (input?.email) {
    const email = upsertEmail(store, input.email);
    const emailKey = makePersistentEmailKey(email);
    ensureTicketEmailLink(store, ticket.id, emailKey);
    for (const groupId of ticket.groupIds) {
      addEmailMembership(store, groupId, email.id, { membershipKind: normalizeGroupMembershipKind(input?.membershipKind) });
    }
  }
  if (useDurableDb) {
    try {
      await upsertDbGroupTicketSeries(store.groupTicketSeries[seriesId]);
      await upsertDbGroupTicket(ticket);
      for (const emailKey of store.groupTicketEmails?.[ticket.id] || []) {
        await upsertDbGroupTicketEmail(ticket.id, emailKey);
      }
      if (input?.email) {
        const email = normalizeEmailInput(input.email);
        const emailKey = resolveEmailKeyFromInput(store, email);
        for (const groupId of ticket.groupIds) {
          await upsertDbCustomGroupMember(groupId, { ...email, ...store.emails[resolveEmailId(store, email)] }, normalizeGroupMembershipKind(input?.membershipKind));
        }
        if (emailKey) await upsertDbGroupTicketEmail(ticket.id, emailKey);
      }
    } catch (error) {
      throw durablePersistenceError("criar o ticket", error);
    }
  }
  writeCacheStore(store);

  return buildGroupTicketEntry(store, ticket);
}

export async function updateGroupTicket(ticketId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("atualizar o ticket", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const tid = normalizeString(ticketId);
  const current = store.groupTickets?.[tid];
  if (!tid || !current) throw new Error("Ticket invalido.");

  const nextGroupIds = normalizeGroupIds(
    Object.prototype.hasOwnProperty.call(input || {}, "groupIds") ? input?.groupIds : current.groupIds
  ).filter((groupId) => Boolean(store.groups?.[groupId]));

  const next = normalizeGroupTicketInput({
    ...current,
    ...input,
    id: tid,
    groupIds: nextGroupIds,
    updatedAt: nowIso(),
  }, current);
  store.groupTickets[tid] = next;
  if (useDurableDb) {
    try {
      await upsertDbGroupTicket(next);
    } catch (error) {
      throw durablePersistenceError("atualizar o ticket", error);
    }
  }
  writeCacheStore(store);

  return buildGroupTicketEntry(store, next);
}

export async function linkEmailToGroupTicket(ticketId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("ligar o email ao ticket", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const tid = normalizeString(ticketId);
  const current = store.groupTickets?.[tid];
  if (!tid || !current) throw new Error("Ticket invalido.");

  const email = upsertEmail(store, input?.email || {});
  const emailKey = makePersistentEmailKey(email);
  const applyGroups = input?.applyGroups !== false;
  const membershipKind = normalizeGroupMembershipKind(input?.membershipKind || "referencia");
  const ensuredGroupIds = normalizeGroupIds(input?.groupIds).filter((groupId) => Boolean(store.groups?.[groupId]));
  const nextGroupIds = Array.from(new Set([...normalizeGroupIds(current.groupIds), ...ensuredGroupIds]));

  store.groupTickets[tid] = normalizeGroupTicketInput({
    ...current,
    groupIds: nextGroupIds,
    updatedAt: nowIso(),
  }, current);
  ensureTicketEmailLink(store, tid, emailKey);

  const appliedGroups = [];
  if (applyGroups) {
    for (const groupId of nextGroupIds) {
      const group = store.groups?.[groupId];
      if (!group) continue;
      addEmailMembership(store, groupId, email.id, { membershipKind });
      appliedGroups.push(buildGroupListEntry(store, group));
    }
  }
  if (useDurableDb) {
    try {
      await upsertDbGroupTicket(store.groupTickets[tid]);
      await upsertDbGroupTicketEmail(tid, emailKey);
      if (applyGroups) {
        for (const groupId of nextGroupIds) {
          const group = store.groups?.[groupId];
          if (!group) continue;
          await upsertDbCustomGroup(group);
          await upsertDbCustomGroupMember(groupId, email, membershipKind);
        }
      }
    } catch (error) {
      throw durablePersistenceError("ligar o email ao ticket", error);
    }
  }
  writeCacheStore(store);

  return {
    ok: true,
    ticket: buildGroupTicketEntry(store, store.groupTickets[tid], { emailLinked: true }),
    appliedGroups: appliedGroups.filter(Boolean),
    email: buildEmailListEntry(email),
  };
}

export async function unlinkEmailFromGroupTicket(ticketId, input) {
  const store = readState();
  const useDurableDb = await requireDurablePersistence("remover o email do ticket", {
    syncStore: async () => {
      await syncGroupTicketsFromDb(store);
    },
  });

  const tid = normalizeString(ticketId);
  const current = store.groupTickets?.[tid];
  if (!tid || !current) throw new Error("Ticket invalido.");

  const emailKey = resolveEmailKeyFromInput(store, input?.email || input || {});
  if (!emailKey) {
    return {
      ok: true,
      removed: false,
      ticket: buildGroupTicketEntry(store, current, { emailLinked: false }),
      emailKey: "",
    };
  }

  const wasLinked = listTicketIdsByEmailKey(store, emailKey).includes(tid);
  removeTicketEmailLink(store, tid, emailKey);

  if (useDurableDb) {
    try {
      await deleteDbGroupTicketEmail(tid, emailKey);
    } catch (error) {
      throw durablePersistenceError("remover o email do ticket", error);
    }
  }

  writeCacheStore(store);

  return {
    ok: true,
    removed: wasLinked,
    ticket: buildGroupTicketEntry(store, store.groupTickets?.[tid] || current, { emailLinked: false }),
    emailKey,
  };
}

export async function detectGroupTicketsForEmail(input) {
  const store = readState();
  if (hasDurablePersistence() && !db.isEnabled()) {
    throw durablePersistenceError("detetar tickets para o email");
  }
  if (db.isEnabled()) {
    try {
      await syncGroupTicketsFromDb(store);
      writeCacheStore(store);
    } catch (error) {
      if (hasDurablePersistence()) throw durablePersistenceError("detetar tickets para o email", error);
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Ticket Detect Sync Error, using central file store:", error.message);
      else console.error("[linkStore] DB Ticket Detect Sync Error, using central file store:", error);
    }
  }

  const emailKey = resolveEmailKeyFromInput(store, input);
  return extractTicketCandidates(store, input).map(({ ticket, matchedCode }) => ({
    matchedCode,
    ticket: buildGroupTicketEntry(store, ticket, {
      emailLinked: emailKey ? listTicketIdsByEmailKey(store, emailKey).includes(normalizeString(ticket.id)) : false,
    }),
    emailLinked: emailKey ? listTicketIdsByEmailKey(store, emailKey).includes(normalizeString(ticket.id)) : false,
    proposedGroups: normalizeGroupIds(ticket.groupIds)
      .map((groupId) => buildGroupListEntry(store, store.groups?.[groupId]))
      .filter(Boolean),
  }));
}

export async function getRelatedEmails(input) {
  const store = readState();
  const currentEmailIds = getMatchedEmailIds(store, input);
  const currentEmailSet = new Set(currentEmailIds);
  const primaryEmailKey = currentEmailIds[0] && store.emails[currentEmailIds[0]]
    ? makePersistentEmailKey(store.emails[currentEmailIds[0]])
    : resolveEmailKeyFromInput(store, input);
  const aggregated = new Map();

  function appendReason(emailId, reason) {
    if (!emailId || currentEmailSet.has(emailId)) return;
    const email = store.emails[emailId];
    if (!email) return;
    const current = aggregated.get(emailId) || {
      ...buildEmailListEntry(email),
      relatedRecords: [],
      relatedGroups: [],
      relatedReasons: [],
    };

    if (reason.kind === "entity") {
      const entityKey = makeEntityKey(reason.model, reason.recordId);
      if (entityKey && !current.relatedRecords.some((entry) => makeEntityKey(entry.model, entry.recordId) === entityKey)) {
        current.relatedRecords.push({
          model: reason.model,
          recordId: reason.recordId,
          recordName: reason.recordName,
        });
      }
    }

    if ((reason.kind === "group" || reason.kind === "conversation") && reason.groupId) {
      if (!current.relatedGroups.some((entry) => entry.id === reason.groupId)) {
        current.relatedGroups.push({
          id: reason.groupId,
          name: reason.groupName,
          kind: reason.kind,
          relationKind: normalizeGroupMembershipKind(reason.relationKind),
        });
      }
    }

    const reasonKey = JSON.stringify(reason);
    if (!current.relatedReasons.some((entry) => JSON.stringify(entry) === reasonKey)) {
      current.relatedReasons.push(reason);
    }
    aggregated.set(emailId, current);
  }

  for (const emailId of currentEmailIds) {
    const email = store.emails[emailId];
    if (!email) continue;

    const entityLinks = Array.isArray(store.emailEntityLinks[emailId]) ? store.emailEntityLinks[emailId] : [];
    for (const entityLink of entityLinks) {
      const entityKey = makeEntityKey(entityLink.model, entityLink.recordId);
      for (const relatedEmailId of store.entityLinks[entityKey] || []) {
        appendReason(relatedEmailId, {
          kind: "entity",
          model: entityLink.model,
          recordId: entityLink.recordId,
          recordName: entityLink.recordName,
        });
      }
    }

    const emailGroups = listEmailGroupMemberships(store, emailId);
    for (const membership of emailGroups) {
      const groupId = membership.groupId;
      const group = store.groups[groupId];
      if (!group) continue;
      const kind = group.kind === "conversation" ? "conversation" : "group";
      for (const relatedEmailId of store.groupMembers[groupId] || []) {
        appendReason(relatedEmailId, {
          kind,
          groupId,
          groupName: group.name,
          conversationId: group.conversationId,
          relationKind: membership.kind,
        });
      }
    }
  }

  const fileResult = {
    email: currentEmailIds[0] ? buildCurrentEmailContextEntry(store, currentEmailIds[0]) : null,
    groups: Array.from(
      new Map(
        currentEmailIds
          .flatMap((emailId) => listEmailGroupMemberships(store, emailId).map((entry) => ({
            groupId: entry.groupId,
            relationKind: normalizeGroupMembershipKind(entry.kind),
          })))
          .map((entry) => {
            const group = store.groups[entry.groupId];
            if (!group) return null;
            return [entry.groupId, {
              ...group,
              relationKind: entry.relationKind,
            }];
          })
          .filter(Boolean)
      ).values()
    ).map((group) => ({
      ...group,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    })),
    tickets: primaryEmailKey
      ? listTicketIdsByEmailKey(store, primaryEmailKey)
        .map((ticketId) => buildGroupTicketEntry(store, store.groupTickets?.[ticketId], { emailLinked: true }))
        .filter(Boolean)
      : [],
    emails: Array.from(aggregated.values()).sort((a, b) =>
      String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || ""))
    ),
  };

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbResult = await getDbCustomGroupContext(input);
      const mergedGroups = new Map(fileResult.groups.map((group) => [group.id, group]));
      for (const group of dbResult.groups) {
        const current = mergedGroups.get(group.id);
        mergedGroups.set(group.id, {
          ...current,
          ...group,
          memberCount: Math.max(Number(current?.memberCount || 0), Number(group.memberCount || 0)),
        });
      }

      return {
        email: mergeEmailContextEntries(fileResult.email, dbResult.email),
        groups: Array.from(mergedGroups.values()),
        tickets: Array.from(
          new Map([...fileResult.tickets, ...(dbResult.tickets || [])].filter(Boolean).map((ticket) => [normalizeString(ticket.id), ticket]))
            .values()
        ),
        emails: dedupeEmailLinks([...fileResult.emails, ...dbResult.emails]),
      };
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Related Group Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Related Group Query Error, falling back to central file store:", error);
    }
  }

  return fileResult;
}

export async function listLinksByConversation(conversationId, internetMessageId = "", itemId = "") {
  const lookup = splitLookupKey(conversationId, internetMessageId);
  const resolvedConversationId = lookup.conversationId;
  const resolvedInternetMessageId = lookup.internetMessageId;
  const resolvedItemId = normalizeString(itemId);

  if (!resolvedConversationId && !resolvedInternetMessageId && !resolvedItemId) return [];

  const fileLinks = listLinksFromFile(resolvedConversationId, resolvedInternetMessageId, resolvedItemId);

  if (db.isEnabled()) {
    try {
      const params = [];
      const where = [];
      if (resolvedConversationId) {
        params.push(resolvedConversationId);
        where.push(`conversation_id = $${params.length}`);
      }
      if (resolvedInternetMessageId) {
        params.push(resolvedInternetMessageId);
        where.push(`LOWER(REGEXP_REPLACE(COALESCE(internet_message_id, ''), '[<>[:space:]]', '', 'g')) = $${params.length}`);
      }
      if (!where.length) return fileLinks;
      const result = await db.query(
        `SELECT * FROM crm_links WHERE ${where.join(" OR ")} ORDER BY linked_at DESC`,
        params
      );
      const rows = result?.rows || [];
      const dbLinks = rows.map((row) => ({
        conversationId: row.conversation_id,
        model: row.model,
        recordId: row.record_id,
        recordName: row.record_name,
        linkedAt: row.linked_at,
        internetMessageId: normalizeMessageId(row.internet_message_id),
        subject: row.subject,
        fromEmail: row.from_email,
        fromName: row.from_name,
      }));
      return dedupeRecordLinks([...fileLinks, ...dbLinks]);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Query Error, falling back to central file store:", error);
    }
  }

  return fileLinks;
}

export async function listLinksByRecord(model, recordId) {
  const normalizedModel = normalizeModel(model);
  const normalizedRecordId = normalizeRecordId(recordId);
  if (!normalizedModel || !normalizedRecordId) return [];

  const fileLinks = listLinksByRecordFromFile(normalizedModel, normalizedRecordId);

  if (db.isEnabled()) {
    try {
      const result = await db.query(
        `SELECT * FROM crm_links WHERE model = $1 AND record_id = $2 ORDER BY linked_at DESC`,
        [normalizedModel, normalizedRecordId]
      );
      const rows = result?.rows || [];
      const dbLinks = rows.map((row) => ({
        conversationId: row.conversation_id,
        model: row.model,
        recordId: row.record_id,
        recordName: row.record_name,
        linkedAt: row.linked_at,
        internetMessageId: normalizeMessageId(row.internet_message_id),
        subject: row.subject,
        fromEmail: row.from_email,
        fromName: row.from_name,
      }));
      return dedupeEmailLinks([...fileLinks, ...dbLinks]);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Record Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Record Query Error, falling back to central file store:", error);
    }
  }

  return fileLinks;
}

export async function addLink(conversationId, entry) {
  const resolvedConversationId = normalizeString(conversationId || entry?.conversationId);
  if (!resolvedConversationId) throw new Error("Missing conversationId");

  const store = readState();
  const useDurableDb = await requireDurablePersistence("gravar a ligacao CRM");
  const email = upsertEmail(store, { ...entry, conversationId: resolvedConversationId });
  const conversationGroup = ensureConversationGroup(store, resolvedConversationId, email);
  if (conversationGroup) addEmailMembership(store, conversationGroup.id, email.id);
  const nextLink = linkEmailToEntity(store, email.id, { ...entry, conversationId: resolvedConversationId });
  if (useDurableDb) {
    try {
      await db.query(
        `INSERT INTO crm_links
         (conversation_id, model, record_id, record_name, linked_at, internet_message_id, subject, from_email, from_name)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (conversation_id, model, record_id) DO UPDATE SET
           record_name = EXCLUDED.record_name,
           linked_at = EXCLUDED.linked_at,
           internet_message_id = EXCLUDED.internet_message_id,
           subject = EXCLUDED.subject,
           from_email = EXCLUDED.from_email,
           from_name = EXCLUDED.from_name`,
        [
          resolvedConversationId,
          nextLink?.model,
          nextLink?.recordId,
          nextLink?.recordName,
          nextLink?.linkedAt,
          nextLink?.internetMessageId,
          nextLink?.subject,
          nextLink?.fromEmail,
          nextLink?.fromName,
        ]
      );
    } catch (error) {
      throw durablePersistenceError("gravar a ligacao CRM", error);
    }
  }
  writeCacheStore(store);

  return await listLinksByConversation(resolvedConversationId, email.internetMessageId, email.itemId);
}
