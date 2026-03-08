import fs from "node:fs";
import path from "node:path";
import pg from "pg";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

const DATA_DIR = path.join(process.cwd(), "server", "data");
const FILE_PATH = path.join(DATA_DIR, "links.json");

// Database configuration
const DATABASE_URL = process.env.DATABASE_URL;
let pool = null;

if (DATABASE_URL) {
  console.log("[linkStore] Using PostgreSQL/Supabase persistence.");
  pool = new pg.Pool({
    connectionString: DATABASE_URL,
    ssl: DATABASE_URL.includes("supabase.co") ? { rejectUnauthorized: false } : false
  });
} else {
  console.log("[linkStore] Using local JSON file persistence.");
}

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
    model: entry.model,
    recordId: Number(entry.recordId),
    recordName: entry.recordName || "",
    linkedAt: entry.linkedAt,
    internetMessageId: normalizeMessageId(entry.internetMessageId),
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

function dedupeLinks(entries) {
  const out = [];
  const seen = new Set();
  for (const raw of entries || []) {
    const entry = normalizeEntry(raw);
    const key = `${entry.model}:${entry.recordId}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(entry);
  }
  return out.sort((a, b) => String(b.linkedAt || "").localeCompare(String(a.linkedAt || "")));
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
  return dedupeLinks([...direct, ...byMessageKey, ...byMessageScan]);
}

function writeLinkToFile(conversationId, entry) {
  const all = readAll();
  const nextEntry = normalizeEntry(entry);
  const keys = [`conversationId:${conversationId}`];
  if (nextEntry.internetMessageId) keys.push(`internetMessageId:${nextEntry.internetMessageId}`);

  for (const key of keys) {
    const arr = Array.isArray(all[key]) ? all[key] : [];
    all[key] = dedupeLinks([nextEntry, ...arr]).slice(0, 50);
  }

  writeAll(all);
}

export async function listLinksByConversation(conversationId, internetMessageId = "") {
  const lookup = splitLookupKey(conversationId, internetMessageId);
  const resolvedConversationId = lookup.conversationId;
  const resolvedInternetMessageId = lookup.internetMessageId;

  if (!resolvedConversationId && !resolvedInternetMessageId) return [];

  if (pool) {
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
      const { rows } = await pool.query(
        `SELECT * FROM crm_links WHERE ${where.join(" OR ")} ORDER BY linked_at DESC`,
        params
      );
      const dbLinks = rows.map((r) => normalizeEntry({
        model: r.model,
        recordId: r.record_id,
        recordName: r.record_name,
        linkedAt: r.linked_at,
        internetMessageId: r.internet_message_id,
        subject: r.subject,
        fromEmail: r.from_email,
        fromName: r.from_name
      }));
      return dedupeLinks([...dbLinks, ...listLinksFromFile(resolvedConversationId, resolvedInternetMessageId)]);
    } catch (e) {
      console.error("[linkStore] DB Query Error, falling back to file store:", e);
    }
  }

  return listLinksFromFile(resolvedConversationId, resolvedInternetMessageId);
}

export async function addLink(conversationId, entry) {
  if (!conversationId) throw new Error("Missing conversationId");
  const nextEntry = normalizeEntry(entry);

  writeLinkToFile(conversationId, nextEntry);

  if (pool) {
    try {
      await pool.query(
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
      console.error("[linkStore] DB Insert Error, falling back to file store:", e);
    }
  }

  return await listLinksByConversation(conversationId, nextEntry.internetMessageId);
}

