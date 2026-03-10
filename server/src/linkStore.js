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

const db = createOptionalPgStore("linkStore");

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
  };
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
    ...extra,
  };
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
    kind: "custom",
    name: normalizeString(input?.name) || "Grupo sem nome",
    description: normalizeString(input?.description),
  });
  writeStore(store);
  return {
    ...group,
    memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
  };
}

export async function listCustomGroups(query = "") {
  const store = readState();
  const q = normalizeString(query).toLowerCase();
  return Object.values(store.groups)
    .filter((group) => group?.kind === "custom")
    .filter((group) => {
      if (!q) return true;
      return String(group?.name || "").toLowerCase().includes(q)
        || String(group?.description || "").toLowerCase().includes(q);
    })
    .sort((a, b) => String(b.updatedAt || b.createdAt || "").localeCompare(String(a.updatedAt || a.createdAt || "")))
    .map((group) => ({
      ...group,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    }));
}

export async function addEmailToGroup(groupId, input) {
  const store = readState();
  const group = ensureGroup(store, { id: groupId });
  const email = upsertEmail(store, input);
  if (email.conversationId) {
    const conversationGroup = ensureConversationGroup(store, email.conversationId, email);
    if (conversationGroup) addEmailMembership(store, conversationGroup.id, email.id);
  }
  addEmailMembership(store, group.id, email.id);
  writeStore(store);
  return {
    group: {
      ...group,
      memberCount: Array.isArray(store.groupMembers[group.id]) ? store.groupMembers[group.id].length : 0,
    },
    email: buildEmailListEntry(email),
  };
}

export async function listEmailsByGroup(groupId) {
  const store = readState();
  const gid = normalizeString(groupId);
  const emailIds = Array.isArray(store.groupMembers[gid]) ? store.groupMembers[gid] : [];
  const rows = emailIds.map((emailId) => {
    const email = store.emails[emailId];
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
  return dedupeEmailLinks(rows);
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

  return {
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
