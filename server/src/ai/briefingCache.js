import fs from "node:fs";
import path from "node:path";
import pg from "pg";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../../.env") });

const DATA_DIR = path.join(process.cwd(), "server", "data");
const FILE_PATH = path.join(DATA_DIR, "briefings.json");
const DATABASE_URL = process.env.DATABASE_URL;

let pool = null;
if (DATABASE_URL) {
    pool = new pg.Pool({
        connectionString: DATABASE_URL,
        ssl: DATABASE_URL.includes("supabase.co") ? { rejectUnauthorized: false } : false
    });
}

function ensureFile() {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    if (!fs.existsSync(FILE_PATH)) fs.writeFileSync(FILE_PATH, JSON.stringify({}), "utf-8");
}

function readAll() {
    ensureFile();
    try {
        const raw = fs.readFileSync(FILE_PATH, "utf-8");
        return JSON.parse(raw || "{}");
    } catch {
        return {};
    }
}

function writeAll(obj) {
    ensureFile();
    fs.writeFileSync(FILE_PATH, JSON.stringify(obj, null, 2), "utf-8");
}

/**
 * Gets a cached briefing for a conversation if it hasn't expired.
 * @param {string} conversationId 
 * @returns {Promise<string|null>}
 */
export async function getBriefing(conversationId) {
    if (!conversationId) return null;

    if (pool) {
        try {
            const { rows } = await pool.query(
                "SELECT summary FROM crm_briefings WHERE conversation_id = $1 AND expires_at > CURRENT_TIMESTAMP",
                [conversationId]
            );
            return rows.length > 0 ? rows[0].summary : null;
        } catch (e) {
            console.warn("[briefingCache] DB read error (may need table creation):", e.message);
            // Fallback to local if DB fails
        }
    }

    const all = readAll();
    const entry = all[conversationId];
    if (entry && new Date(entry.expiresAt) > new Date()) {
        return entry.summary;
    }
    return null;
}

/**
 * Saves a briefing to the cache with a 5-day expiration.
 * @param {string} conversationId 
 * @param {string} summary 
 */
export async function saveBriefing(conversationId, summary) {
    if (!conversationId || !summary) return;

    const expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + 5);

    if (pool) {
        try {
            await pool.query(
                `INSERT INTO crm_briefings (conversation_id, summary, expires_at)
         VALUES ($1, $2, $3)
         ON CONFLICT (conversation_id) DO UPDATE SET summary = $2, expires_at = $3, created_at = CURRENT_TIMESTAMP`,
                [conversationId, summary, expiresAt]
            );
            return;
        } catch (e) {
            console.warn("[briefingCache] DB save error:", e.message);
        }
    }

    const all = readAll();
    all[conversationId] = {
        summary,
        expiresAt: expiresAt.toISOString()
    };
    writeAll(all);
}

/**
 * Ensures the briefings table exists in the database.
 */
export async function initBriefingDb() {
    if (!pool) return;
    try {
        await pool.query(`
            CREATE TABLE IF NOT EXISTS crm_briefings (
                conversation_id TEXT PRIMARY KEY,
                summary TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP NOT NULL
            );
        `);
        console.log("[briefingCache] Database table ensured.");
    } catch (e) {
        console.error("[briefingCache] Failed to initialize database table:", e.message);
    }
}
