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

export async function listLinksByConversation(conversationId) {
  if (!conversationId) return [];

  if (pool) {
    try {
      const { rows } = await pool.query(
        "SELECT * FROM crm_links WHERE conversation_id = $1 ORDER BY linked_at DESC",
        [conversationId]
      );
      // Map DB snake_case to JS camelCase for UI compatibility
      return rows.map(r => ({
        model: r.model,
        recordId: r.record_id,
        recordName: r.record_name,
        linkedAt: r.linked_at,
        internetMessageId: r.internet_message_id,
        subject: r.subject,
        fromEmail: r.from_email,
        fromName: r.from_name
      }));
    } catch (e) {
      console.error("[linkStore] DB Query Error:", e);
      return [];
    }
  }

  const all = readAll();
  return all[`conversationId:${conversationId}`] || [];
}

export async function addLink(conversationId, entry) {
  if (!conversationId) throw new Error("Missing conversationId");

  if (pool) {
    try {
      await pool.query(
        `INSERT INTO crm_links 
         (conversation_id, model, record_id, record_name, linked_at, internet_message_id, subject, from_email, from_name)
         VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9)
         ON CONFLICT (conversation_id, model, record_id) DO UPDATE SET linked_at = $5`,
        [
          conversationId,
          entry.model,
          entry.recordId,
          entry.recordName,
          entry.linkedAt,
          entry.internetMessageId,
          entry.subject,
          entry.fromEmail,
          entry.fromName
        ]
      );
      return await listLinksByConversation(conversationId);
    } catch (e) {
      console.error("[linkStore] DB Insert Error:", e);
      // Fallback logic not implemented for error state, returning empty or previous
      return [];
    }
  }

  const all = readAll();
  const key = `conversationId:${conversationId}`;
  const arr = all[key] || [];

  // Deduplicate by model+recordId
  const exists = arr.some((x) => x.model === entry.model && Number(x.recordId) === Number(entry.recordId));
  if (!exists) arr.unshift(entry);

  // Keep last 50
  all[key] = arr.slice(0, 50);
  writeAll(all);
  return all[key];
}
