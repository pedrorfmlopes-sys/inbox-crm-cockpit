import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import { createOptionalPgStore } from "./optionalPg.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

const DATA_DIR = path.join(process.cwd(), "server", "data");
const LOG_FILE_PATH = path.join(DATA_DIR, "learning_logs.json");
const PROFILE_FILE_PATH = path.join(DATA_DIR, "style_profiles.json");

const db = createOptionalPgStore("learningStore");

function ensureFile(filePath) {
    if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
    if (!fs.existsSync(filePath)) fs.writeFileSync(filePath, JSON.stringify(filePath.includes("profiles") ? {} : []), "utf-8");
}

function readAll(filePath) {
    ensureFile(filePath);
    const raw = fs.readFileSync(filePath, "utf-8");
    try {
        return JSON.parse(raw || (filePath.includes("profiles") ? "{}" : "[]"));
    } catch {
        return filePath.includes("profiles") ? {} : [];
    }
}

function writeAll(filePath, obj) {
    ensureFile(filePath);
    fs.writeFileSync(filePath, JSON.stringify(obj, null, 2), "utf-8");
}

export async function initLearningDb() {
    if (!db.isEnabled()) return;
    try {
        await db.query(`
            CREATE TABLE IF NOT EXISTS learning_logs (
                id BIGSERIAL PRIMARY KEY,
                conversation_id TEXT,
                from_email TEXT,
                to_emails JSONB,
                original_subject TEXT,
                original_body TEXT,
                user_response TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        `);

        await db.query(`
            CREATE TABLE IF NOT EXISTS style_profiles (
                user_id TEXT PRIMARY KEY,
                style_data JSONB DEFAULT '{}'::jsonb,
                habits_data JSONB DEFAULT '{}'::jsonb,
                last_updated TIMESTAMP
            );
        `);

        console.log("[learningStore] Database tables ensured.");
    } catch (e) {
        console.error("[learningStore] Failed to initialize database tables:", e.message);
    }
}

/**
 * Logs an interaction
 */
export async function logInteraction(log) {
    const entry = {
        conversationId: log.conversationId,
        fromEmail: log.fromEmail,
        toEmails: Array.isArray(log.toEmails) ? log.toEmails : [],
        originalSubject: log.originalSubject,
        originalBody: log.originalBody,
        userResponse: log.userResponse,
        createdAt: new Date().toISOString()
    };

    if (db.isEnabled()) {
        try {
            await db.query(
                `INSERT INTO learning_logs 
         (conversation_id, from_email, to_emails, original_subject, original_body, user_response, created_at)
         VALUES ($1, $2, $3, $4, $5, $6, $7)`,
                [
                    entry.conversationId,
                    entry.fromEmail,
                    entry.toEmails,
                    entry.originalSubject,
                    entry.originalBody,
                    entry.userResponse,
                    entry.createdAt
                ]
            );
            return { ok: true };
        } catch (e) {
            if (e?.optionalDbFallback) console.warn("[learningStore] DB Insert Error (learning_logs):", e.message);
            else console.error("[learningStore] DB Insert Error (learning_logs):", e);
            // Fallback below
        }
    }

    const all = readAll(LOG_FILE_PATH);
    all.push(entry);
    // Keep last 1000 logs in local file to avoid bloat
    const limited = all.slice(-1000);
    writeAll(LOG_FILE_PATH, limited);
    return { ok: true };
}

/**
 * Gets recent interaction logs
 */
export async function getLogs(limit = 50) {
    if (db.isEnabled()) {
        try {
            const result = await db.query(
                "SELECT * FROM learning_logs ORDER BY created_at DESC LIMIT $1",
                [limit]
            );
            const rows = result?.rows || [];
            return rows.map(r => ({
                conversationId: r.conversation_id,
                fromEmail: r.from_email,
                toEmails: r.to_emails,
                originalSubject: r.original_subject,
                originalBody: r.original_body,
                userResponse: r.user_response,
                createdAt: r.created_at
            }));
        } catch (e) {
            if (e?.optionalDbFallback) console.warn("[learningStore] DB Query Error (learning_logs):", e.message);
            else console.error("[learningStore] DB Query Error (learning_logs):", e);
        }
    }

    const all = readAll(LOG_FILE_PATH);
    return all.slice(-limit).reverse();
}

/**
 * Gets or sets the style profile
 */
export async function getStyleProfile(userId = "global") {
    if (db.isEnabled()) {
        try {
            const result = await db.query(
                "SELECT * FROM style_profiles WHERE user_id = $1",
                [userId]
            );
            const rows = result?.rows || [];
            if (rows.length > 0) {
                return {
                    styleData: rows[0].style_data,
                    habitsData: rows[0].habits_data,
                    lastUpdated: rows[0].last_updated
                };
            }
        } catch (e) {
            if (e?.optionalDbFallback) console.warn("[learningStore] DB Query Error (style_profiles):", e.message);
            else console.error("[learningStore] DB Query Error (style_profiles):", e);
        }
    }

    const all = readAll(PROFILE_FILE_PATH);
    return all[userId] || { styleData: {}, habitsData: {}, lastUpdated: null };
}

export async function updateStyleProfile(userId = "global", profile) {
    const lastUpdated = new Date().toISOString();

    if (db.isEnabled()) {
        try {
            await db.query(
                `INSERT INTO style_profiles (user_id, style_data, habits_data, last_updated)
         VALUES ($1, $2, $3, $4)
         ON CONFLICT (user_id) DO UPDATE SET 
            style_data = $2, 
            habits_data = $3, 
            last_updated = $4`,
                [userId, profile.styleData || {}, profile.habitsData || {}, lastUpdated]
            );
            return { ok: true };
        } catch (e) {
            if (e?.optionalDbFallback) console.warn("[learningStore] DB Update Error (style_profiles):", e.message);
            else console.error("[learningStore] DB Update Error (style_profiles):", e);
        }
    }

    const all = readAll(PROFILE_FILE_PATH);
    all[userId] = {
        ...profile,
        lastUpdated
    };
    writeAll(PROFILE_FILE_PATH, all);
    return { ok: true };
}
