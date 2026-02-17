import fs from "node:fs";
import path from "node:path";

const DATA_DIR = path.join(process.cwd(), "server", "data");
const FILE_PATH = path.join(DATA_DIR, "links.json");

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

export function listLinksByConversation(conversationId) {
  if (!conversationId) return [];
  const all = readAll();
  return all[`conversationId:${conversationId}`] || [];
}

export function addLink(conversationId, entry) {
  if (!conversationId) throw new Error("Missing conversationId");
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
