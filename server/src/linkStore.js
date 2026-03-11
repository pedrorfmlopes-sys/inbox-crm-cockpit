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

const db = createOptionalPgStore("linkStore");
let customGroupDbInitPromise = null;

function nowIso() {
  return new Date().toISOString();
}

function normalizeString(value) {
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
    emailGroups: {},
    conversationGroups: {},
    entityLinks: {},
    emailEntityLinks: {},
    groupDocuments: {},
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
    ...(attachments.length ? { attachments } : {}),
  };
}

function normalizeAttachments(value) {
  if (!Array.isArray(value)) return [];
  return value
    .map((attachment) => ({
      name: normalizeString(attachment?.name),
      contentType: normalizeString(attachment?.contentType),
      size: Number(attachment?.size || 0) || undefined,
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
  store.emailGroups = source.emailGroups && typeof source.emailGroups === "object" ? source.emailGroups : {};
  store.conversationGroups = source.conversationGroups && typeof source.conversationGroups === "object" ? source.conversationGroups : {};
  store.entityLinks = source.entityLinks && typeof source.entityLinks === "object" ? source.entityLinks : {};
  store.emailEntityLinks = source.emailEntityLinks && typeof source.emailEntityLinks === "object" ? source.emailEntityLinks : {};
  store.groupDocuments = source.groupDocuments && typeof source.groupDocuments === "object" ? source.groupDocuments : {};
  return store;
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
    ...Object.fromEntries(Object.entries(normalized).filter(([, value]) => value)),
    updatedAt: now,
    lastSeenAt: now,
  };

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
  const next = {
    ...current,
    kind: normalizeString(partial?.kind) || current.kind || "custom",
    name: normalizeString(partial?.name) || current.name || "Grupo sem nome",
    description: normalizeString(partial?.description) || current.description || "",
    conversationId: normalizeString(partial?.conversationId) || current.conversationId || "",
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

function addEmailMembership(store, groupId, emailId) {
  const gid = normalizeString(groupId);
  const eid = normalizeString(emailId);
  if (!gid || !eid) return;
  const members = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
  if (!members.includes(eid)) store.groupMembers[gid] = [...members, eid];
  const emailGroups = Array.isArray(store.emailGroups[eid]) ? store.emailGroups[eid] : [];
  if (!emailGroups.includes(gid)) store.emailGroups[eid] = [...emailGroups, gid];
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
    attachments: normalizeAttachments(email?.attachments),
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

function mapDbGroupRow(row) {
  if (!row) return null;
  return {
    id: normalizeString(row.id),
    kind: CUSTOM_GROUP_KIND,
    name: normalizeString(row.name),
    description: normalizeString(row.description),
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
    createdAt: normalizeString(row.created_at),
    updatedAt: normalizeString(row.updated_at),
    attachments: parseAttachmentsJson(row.attachments_json),
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

async function upsertDbCustomGroup(group) {
  if (!db.isEnabled()) return;
  await db.query(
    `INSERT INTO crm_custom_groups (id, name, description, documents_enabled, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6)
     ON CONFLICT (id) DO UPDATE SET
       name = EXCLUDED.name,
       description = EXCLUDED.description,
       documents_enabled = EXCLUDED.documents_enabled,
       updated_at = EXCLUDED.updated_at`,
    [
      normalizeString(group?.id),
      normalizeString(group?.name) || "Grupo sem nome",
      normalizeString(group?.description),
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

async function upsertDbCustomGroupMember(groupId, email) {
  if (!db.isEnabled()) return;
  const emailKey = makePersistentEmailKey(email);
  if (!groupId || !emailKey) return;
  await db.query(
    `INSERT INTO crm_custom_group_members
       (group_id, email_key, item_id, internet_message_id, conversation_id, subject, from_email, from_name, email_web_link, message_date_iso, received_at_iso, sent_at_iso, attachments_json, created_at, updated_at)
     VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13::jsonb, $14, $15)
     ON CONFLICT (group_id, email_key) DO UPDATE SET
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
       attachments_json = EXCLUDED.attachments_json,
       updated_at = EXCLUDED.updated_at`,
    [
      normalizeString(groupId),
      emailKey,
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
    `SELECT id, name, description, documents_enabled, created_at, updated_at
     FROM crm_custom_groups
     WHERE id = $1`,
    [gid]
  );
  return mapDbGroupRow(result?.rows?.[0]);
}

async function listDbCustomGroups(query = "") {
  if (!db.isEnabled()) return [];
  const q = normalizeString(query).toLowerCase();
  const listAllAlphabetically = q === "7" || q === "*";
  const params = [];
  const where = [];
  if (q && !listAllAlphabetically) {
    params.push(`%${q}%`);
    where.push(`(LOWER(name) LIKE $${params.length} OR LOWER(COALESCE(description, '')) LIKE $${params.length})`);
  }
  const result = await db.query(
    `SELECT g.id, g.name, g.description, g.documents_enabled, g.created_at, g.updated_at, COUNT(m.email_key)::int AS member_count
     FROM crm_custom_groups g
     LEFT JOIN crm_custom_group_members m ON m.group_id = g.id
     ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
     GROUP BY g.id, g.name, g.description, g.documents_enabled, g.created_at, g.updated_at
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
    })
  ));
}

async function getDbCustomGroupContext(input) {
  if (!db.isEnabled()) return { groups: [], emails: [] };

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
  if (!where.length) return { groups: [], emails: [] };

  const currentMemberships = await db.query(
    `SELECT DISTINCT g.id, g.name, g.description, g.created_at, g.updated_at, m.email_key
     FROM crm_custom_group_members m
     JOIN crm_custom_groups g ON g.id = m.group_id
     WHERE ${where.join(" OR ")}`,
    params
  );

  const membershipRows = currentMemberships?.rows || [];
  if (!membershipRows.length) return { groups: [], emails: [] };

  const currentEmailKeys = new Set(membershipRows.map((row) => normalizeString(row.email_key)).filter(Boolean));
  const groups = Array.from(
    new Map(membershipRows.map((row) => [normalizeString(row.id), mapDbGroupRow(row)])).values()
  ).filter(Boolean);

  const groupIds = groups.map((group) => group.id).filter(Boolean);
  if (!groupIds.length) return { groups, emails: [] };

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
      });
    }

    const reason = {
      kind: "group",
      groupId: normalizeString(row.group_id),
      groupName: normalizeString(row.group_name),
      conversationId: normalizeString(row.conversation_id),
    };
    const reasonKey = JSON.stringify(reason);
    if (!current.relatedReasons.some((entry) => JSON.stringify(entry) === reasonKey)) {
      current.relatedReasons.push(reason);
    }

    aggregated.set(rowEmailKey, current);
  }

  return {
    groups: groups.map((group) => ({ ...group, memberCount: 0 })),
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
      CREATE TABLE IF NOT EXISTS crm_custom_group_members (
        group_id TEXT NOT NULL REFERENCES crm_custom_groups(id) ON DELETE CASCADE,
        email_key TEXT NOT NULL,
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

    const store = readState();
    const customGroups = Object.values(store.groups || {}).filter((group) => group?.kind === CUSTOM_GROUP_KIND);
    for (const group of customGroups) {
      await upsertDbCustomGroup(group);
      for (const emailId of store.groupMembers[group.id] || []) {
        const email = buildRecoveredEmailSnapshot(store, emailId);
        if (email) await upsertDbCustomGroupMember(group.id, email);
      }
      for (const doc of store.groupDocuments[group.id] || []) {
        await upsertDbGroupDocument(group.id, doc);
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
  writeStore(store);
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
  const group = ensureGroup(store, {
    kind: CUSTOM_GROUP_KIND,
    name: normalizeString(input?.name) || "Grupo sem nome",
    description: normalizeString(input?.description),
    documentsEnabled: typeof input?.documentsEnabled === "boolean" ? input.documentsEnabled : true,
  });
  writeStore(store);
  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await upsertDbCustomGroup(group);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Insert Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Custom Group Insert Error, central file store kept as source of truth:", error);
    }
  }
  return {
    ...group,
    memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
  };
}

export async function updateCustomGroup(groupId, input) {
  const store = readState();
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
    documentsEnabled:
      typeof input?.documentsEnabled === "boolean"
        ? input.documentsEnabled
        : current.documentsEnabled,
  });
  writeStore(store);

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await upsertDbCustomGroup(group);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Update Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Custom Group Update Error, central file store kept as source of truth:", error);
    }
  }

  return {
    ...group,
    memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
  };
}

export async function listCustomGroups(query = "") {
  const store = readState();
  const q = normalizeString(query).toLowerCase();
  const listAllAlphabetically = q === "7" || q === "*";
  const fileGroups = Object.values(store.groups)
    .filter((group) => group?.kind === "custom")
    .filter((group) => {
      if (!q || listAllAlphabetically) return true;
      return String(group?.name || "").toLowerCase().includes(q)
        || String(group?.description || "").toLowerCase().includes(q);
    })
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "pt"))
    .map((group) => ({
      ...group,
      documentsEnabled: group.documentsEnabled !== false,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    }));

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      const dbGroups = await listDbCustomGroups(query);
      const merged = new Map(fileGroups.map((group) => [group.id, group]));
      for (const group of dbGroups) {
        merged.set(group.id, {
          ...group,
          memberCount: Math.max(Number(group.memberCount || 0), Number(merged.get(group.id)?.memberCount || 0)),
        });
      }
      return Array.from(merged.values()).sort((a, b) =>
        String(a?.name || "").localeCompare(String(b?.name || ""), "pt")
      );
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Custom Group Query Error, falling back to central file store:", error);
    }
  }

  return fileGroups;
}

export async function addEmailToGroup(groupId, input) {
  const store = readState();
  let existingDbGroup = null;
  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      existingDbGroup = await getDbCustomGroupById(groupId);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Read Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Custom Group Read Error, central file store kept as source of truth:", error);
    }
  }

  const group = ensureGroup(store, {
    id: groupId,
    kind: CUSTOM_GROUP_KIND,
    name: existingDbGroup?.name,
    description: existingDbGroup?.description,
    documentsEnabled: existingDbGroup?.documentsEnabled,
  });
  const email = upsertEmail(store, input);
  if (email.conversationId) {
    const conversationGroup = ensureConversationGroup(store, email.conversationId, email);
    if (conversationGroup) addEmailMembership(store, conversationGroup.id, email.id);
  }
  addEmailMembership(store, group.id, email.id);
  writeStore(store);

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await upsertDbCustomGroup(group);
      await upsertDbCustomGroupMember(group.id, email);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Membership Insert Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Group Membership Insert Error, central file store kept as source of truth:", error);
    }
  }

  return {
    group: {
      ...group,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    },
    email: buildEmailListEntry(email),
  };
}

export async function removeEmailFromGroup(groupId, input) {
  const store = readState();
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

  if (removed) {
    writeStore(store);
  }

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await removeDbCustomGroupMember(groupId, input);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Membership Delete Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Group Membership Delete Error, central file store kept as source of truth:", error);
    }
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
  const gid = normalizeString(groupId);
  if (!gid) return { ok: true, deleted: false };

  const memberIds = Array.isArray(store.groupMembers[gid]) ? [...store.groupMembers[gid]] : [];
  for (const emailId of memberIds) {
    removeEmailMembership(store, gid, emailId);
  }
  delete store.groupMembers[gid];
  delete store.groupDocuments[gid];
  delete store.groups[gid];
  writeStore(store);

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await db.query(`DELETE FROM crm_custom_groups WHERE id = $1`, [gid]);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Custom Group Delete Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Custom Group Delete Error, central file store kept as source of truth:", error);
    }
  }

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

export async function listDocumentsByGroup(groupId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const fileRows = Array.isArray(store.groupDocuments?.[gid])
    ? store.groupDocuments[gid].map((doc) => normalizeGroupDocumentInput(doc))
    : [];

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
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Document Query Error, falling back to central file store:", error.message);
      else console.error("[linkStore] DB Group Document Query Error, falling back to central file store:", error);
    }
  }

  return fileRows.sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
}

export async function saveDocumentsToGroup(groupId, input) {
  const store = readState();
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
  writeStore(store);

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      for (const doc of docs) {
        await upsertDbGroupDocument(gid, doc);
      }
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Document Save Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Group Document Save Error, central file store kept as source of truth:", error);
    }
  }

  return {
    ok: true,
    group: store.groups[gid],
    documents: await listDocumentsByGroup(gid),
  };
}

export async function deleteDocumentFromGroup(groupId, documentId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const did = normalizeString(documentId);
  const current = Array.isArray(store.groupDocuments[gid]) ? store.groupDocuments[gid] : [];
  const next = current.filter((doc) => normalizeString(doc?.id) !== did);
  const removed = next.length !== current.length;
  store.groupDocuments[gid] = next;
  if (removed && store.groups[gid]) store.groups[gid].updatedAt = nowIso();
  if (removed) writeStore(store);

  if (db.isEnabled()) {
    try {
      await ensureCustomGroupDb();
      await deleteDbGroupDocument(gid, did);
    } catch (error) {
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Group Document Delete Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Group Document Delete Error, central file store kept as source of truth:", error);
    }
  }

  return { ok: true, removed, groupId: gid, documentId: did };
}

export async function getRelatedEmails(input) {
  const store = readState();
  const currentEmailIds = getMatchedEmailIds(store, input);
  const currentEmailSet = new Set(currentEmailIds);
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

    const emailGroups = Array.isArray(store.emailGroups[emailId]) ? store.emailGroups[emailId] : [];
    for (const groupId of emailGroups) {
      const group = store.groups[groupId];
      if (!group) continue;
      const kind = group.kind === "conversation" ? "conversation" : "group";
      for (const relatedEmailId of store.groupMembers[groupId] || []) {
        appendReason(relatedEmailId, {
          kind,
          groupId,
          groupName: group.name,
          conversationId: group.conversationId,
        });
      }
    }
  }

  const fileResult = {
    email: currentEmailIds[0] ? buildEmailListEntry(store.emails[currentEmailIds[0]]) : null,
    groups: Array.from(
      new Map(
        currentEmailIds
          .flatMap((emailId) => store.emailGroups[emailId] || [])
          .map((groupId) => [groupId, store.groups[groupId]])
          .filter(([, group]) => Boolean(group))
      ).values()
    ).map((group) => ({
      ...group,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    })),
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
        email: fileResult.email,
        groups: Array.from(mergedGroups.values()),
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
  const email = upsertEmail(store, { ...entry, conversationId: resolvedConversationId });
  const conversationGroup = ensureConversationGroup(store, resolvedConversationId, email);
  if (conversationGroup) addEmailMembership(store, conversationGroup.id, email.id);
  const nextLink = linkEmailToEntity(store, email.id, { ...entry, conversationId: resolvedConversationId });
  writeStore(store);

  if (db.isEnabled()) {
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
      if (error?.optionalDbFallback) console.warn("[linkStore] DB Insert Error, central file store kept as source of truth:", error.message);
      else console.error("[linkStore] DB Insert Error, central file store kept as source of truth:", error);
    }
  }

  return await listLinksByConversation(resolvedConversationId, email.internetMessageId, email.itemId);
}
