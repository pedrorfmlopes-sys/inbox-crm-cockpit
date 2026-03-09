import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import { createOptionalPgStore } from "./optionalPg.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

const DATA_DIR = path.join(process.cwd(), "server", "data");
const FILE_PATH = path.join(DATA_DIR, "links.json");

const db = createOptionalPgStore("linkStore");

/**
 * Simple file store:
 * {
 *   "conversationId:<id>": [
 *     { model, recordId, recordName, linkedAt, internetMessageId, subject, fromEmail, fromName }
 *   ]
 * }
 */
function ensureFile() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(FILE_PATH)) fs.writeFileSync(FILE_PATH, JSON.stringify({}), "utf-8");
}

function readAll() {
  ensureFile();
  const raw = fs.readFileSync(FILE_PATH, "utf-8");
  try {
    return JSON.parse(raw || "{}");
  } catch {
    return {};
  }
}

function writeAll(obj) {
  ensureFile();
  fs.writeFileSync(FILE_PATH, JSON.stringify(obj, null, 2), "utf-8");
}

function normalizeEntry(entry) {
  return {
    conversationId: String(entry.conversationId || "").trim(),
    model: entry.model,
    recordId: Number(entry.recordId),
    recordName: entry.recordName || "",
    linkedAt: entry.linkedAt,
    internetMessageId: normalizeMessageId(entry.internetMessageId),
    itemId: String(entry.itemId || "").trim(),
    emailWebLink: String(entry.emailWebLink || "").trim(),
    receivedAtIso: String(entry.receivedAtIso || "").trim(),
    subject: entry.subject,
    fromEmail: entry.fromEmail,
    fromName: entry.fromName
  };
}

function normalizeMessageId(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[<>\s]/g, "");
}

function splitLookupKey(conversationId, internetMessageId = "") {
  const rawConversationId = String(conversationId || "").trim();
  if (rawConversationId.includes("||")) {
    const [cid, imid] = rawConversationId.split("||");
    return {
      conversationId: String(cid || "").trim(),
      internetMessageId: normalizeMessageId(internetMessageId || imid),
    };
  }
  return {
    conversationId: rawConversationId,
    internetMessageId: normalizeMessageId(internetMessageId),
  };
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
  if (String(incoming?.receivedAtIso || "") > String(current?.receivedAtIso || "")) {
    next.receivedAtIso = incoming.receivedAtIso;
  }
  return next;
}

function dedupeRecordLinks(entries) {
  const seen = new Map();
  for (const raw of entries || []) {
    const entry = normalizeEntry(raw);
    const key = `${entry.model}:${entry.recordId}`;
    const current = seen.get(key);
    seen.set(key, current ? mergeEntryData(current, entry) : entry);
  }
  return Array.from(seen.values()).sort((a, b) => String(b.linkedAt || "").localeCompare(String(a.linkedAt || "")));
}

function makeEmailLookupKey(entry) {
  const normalized = normalizeEntry(entry);
  return normalized.itemId
    || normalized.internetMessageId
    || normalized.conversationId
    || [
      String(normalized.subject || "").trim().toLowerCase(),
      String(normalized.fromEmail || "").trim().toLowerCase(),
      String(normalized.receivedAtIso || normalized.linkedAt || "").trim(),
    ].join("|");
}

function dedupeEmailLinks(entries) {
  const seen = new Map();
  for (const raw of entries || []) {
    const entry = normalizeEntry(raw);
    const key = makeEmailLookupKey(entry);
    if (!key) continue;
    const current = seen.get(key);
    seen.set(key, current ? mergeEntryData(current, entry) : entry);
  }
  return Array.from(seen.values()).sort((a, b) =>
    String(b.receivedAtIso || b.linkedAt || "").localeCompare(String(a.receivedAtIso || a.linkedAt || ""))
  );
}

function parseStorageKey(key) {
  const raw = String(key || "").trim();
  if (raw.startsWith("conversationId:")) {
    return { conversationId: raw.slice("conversationId:".length).trim(), internetMessageId: "" };
  }
  if (raw.startsWith("internetMessageId:")) {
    return { conversationId: "", internetMessageId: normalizeMessageId(raw.slice("internetMessageId:".length)) };
  }
  return { conversationId: "", internetMessageId: "" };
}

function listLinksFromFile(conversationId, internetMessageId = "") {
  if (!conversationId && !internetMessageId) return [];
  const all = readAll();
  const direct = conversationId ? (all[`conversationId:${conversationId}`] || []) : [];
  const byMessageKey = internetMessageId ? (all[`internetMessageId:${internetMessageId}`] || []) : [];
  const byMessageScan = internetMessageId
    ? Object.values(all)
      .flatMap((entries) => Array.isArray(entries) ? entries : [])
      .filter((entry) => normalizeMessageId(entry?.internetMessageId) === internetMessageId)
    : [];
  return dedupeRecordLinks([...direct, ...byMessageKey, ...byMessageScan]);
}

function listLinksByRecordFromFile(model, recordId) {
  const normalizedModel = String(model || "").trim();
  const normalizedRecordId = Number(recordId || 0);
  if (!normalizedModel || !normalizedRecordId) return [];

  const all = readAll();
  const entries = Object.entries(all).flatMap(([key, value]) => {
    const keyMeta = parseStorageKey(key);
    const items = Array.isArray(value) ? value : [];
    return items
      .filter((entry) => String(entry?.model || "").trim() === normalizedModel && Number(entry?.recordId || 0) === normalizedRecordId)
      .map((entry) => normalizeEntry({
        ...entry,
        conversationId: entry?.conversationId || keyMeta.conversationId,
        internetMessageId: entry?.internetMessageId || keyMeta.internetMessageId,
      }));
  });

  return dedupeEmailLinks(entries);
}

function writeLinkToFile(conversationId, entry) {
  const all = readAll();
  const nextEntry = normalizeEntry({ ...entry, conversationId });
  const keys = [`conversationId:${conversationId}`];
  if (nextEntry.internetMessageId) keys.push(`internetMessageId:${nextEntry.internetMessageId}`);

  for (const key of keys) {
    const arr = Array.isArray(all[key]) ? all[key] : [];
    all[key] = dedupeRecordLinks([nextEntry, ...arr]).slice(0, 50);
  }

  writeAll(all);
}

export async function listLinksByConversation(conversationId, internetMessageId = "") {
  const lookup = splitLookupKey(conversationId, internetMessageId);
  const resolvedConversationId = lookup.conversationId;
  const resolvedInternetMessageId = lookup.internetMessageId;

  if (!resolvedConversationId && !resolvedInternetMessageId) return [];

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
      const result = await db.query(
        `SELECT * FROM crm_links WHERE ${where.join(" OR ")} ORDER BY linked_at DESC`,
        params
      );
      const rows = result?.rows || [];
      const dbLinks = rows.map((r) => normalizeEntry({
        conversationId: r.conversation_id,
        model: r.model,
        recordId: r.record_id,
        recordName: r.record_name,
        linkedAt: r.linked_at,
        internetMessageId: r.internet_message_id,
        subject: r.subject,
        fromEmail: r.from_email,
        fromName: r.from_name
      }));
      return dedupeRecordLinks([...dbLinks, ...listLinksFromFile(resolvedConversationId, resolvedInternetMessageId)]);
    } catch (e) {
      if (e?.optionalDbFallback) console.warn("[linkStore] DB Query Error, falling back to file store:", e.message);
      else console.error("[linkStore] DB Query Error, falling back to file store:", e);
    }
  }

  return listLinksFromFile(resolvedConversationId, resolvedInternetMessageId);
}

export async function listLinksByRecord(model, recordId) {
  const normalizedModel = String(model || "").trim();
  const normalizedRecordId = Number(recordId || 0);
  if (!normalizedModel || !normalizedRecordId) return [];

  if (db.isEnabled()) {
    try {
      const result = await db.query(
        `SELECT * FROM crm_links WHERE model = $1 AND record_id = $2 ORDER BY linked_at DESC`,
        [normalizedModel, normalizedRecordId]
      );
      const rows = result?.rows || [];
      const dbLinks = rows.map((r) => normalizeEntry({
        conversationId: r.conversation_id,
        model: r.model,
        recordId: r.record_id,
        recordName: r.record_name,
        linkedAt: r.linked_at,
        internetMessageId: r.internet_message_id,
        subject: r.subject,
        fromEmail: r.from_email,
        fromName: r.from_name,
      }));
      return dedupeEmailLinks([...dbLinks, ...listLinksByRecordFromFile(normalizedModel, normalizedRecordId)]);
    } catch (e) {
      if (e?.optionalDbFallback) console.warn("[linkStore] DB Record Query Error, falling back to file store:", e.message);
      else console.error("[linkStore] DB Record Query Error, falling back to file store:", e);
    }
  }

  return listLinksByRecordFromFile(normalizedModel, normalizedRecordId);
}

export async function addLink(conversationId, entry) {
  if (!conversationId) throw new Error("Missing conversationId");
  const nextEntry = normalizeEntry(entry);

  writeLinkToFile(conversationId, nextEntry);

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
          conversationId,
          nextEntry.model,
          nextEntry.recordId,
          nextEntry.recordName,
          nextEntry.linkedAt,
          nextEntry.internetMessageId,
          nextEntry.subject,
          nextEntry.fromEmail,
          nextEntry.fromName
        ]
      );
      return await listLinksByConversation(conversationId, nextEntry.internetMessageId);
    } catch (e) {
      if (e?.optionalDbFallback) console.warn("[linkStore] DB Insert Error, falling back to file store:", e.message);
      else console.error("[linkStore] DB Insert Error, falling back to file store:", e);
    }
  }

  return await listLinksByConversation(conversationId, nextEntry.internetMessageId);
}

