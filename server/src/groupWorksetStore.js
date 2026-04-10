import fs from "node:fs";
import path from "node:path";
import dotenv from "dotenv";
import { fileURLToPath } from "node:url";
import { createOptionalPgStore } from "./optionalPg.js";
import {
  buildGroupWorksetPayloadScore,
  mergeGroupWorksetManifest,
  normalizeGroupWorksetManifest,
} from "./groupWorksetManifest.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

const PRIMARY_DATA_DIR = process.env.ICC_DATA_DIR
  ? path.resolve(process.env.ICC_DATA_DIR)
  : path.resolve(__dirname, "../data");
const WORKSET_FILE_PATH = path.join(PRIMARY_DATA_DIR, "groupWorksets.json");
const WORKSET_FILE_VERSION = 1;

const db = createOptionalPgStore("groupWorksetStore");
let dbInitPromise = null;

function ensurePrimaryDir() {
  if (!fs.existsSync(PRIMARY_DATA_DIR)) fs.mkdirSync(PRIMARY_DATA_DIR, { recursive: true });
}

function createEmptyWorksetStore() {
  return {
    version: WORKSET_FILE_VERSION,
    worksets: {},
  };
}

function readRawFileStore() {
  try {
    if (!fs.existsSync(WORKSET_FILE_PATH)) return null;
    return JSON.parse(fs.readFileSync(WORKSET_FILE_PATH, "utf-8") || "{}");
  } catch {
    return null;
  }
}

function writeFileStore(store) {
  ensurePrimaryDir();
  fs.writeFileSync(WORKSET_FILE_PATH, JSON.stringify(store, null, 2), "utf-8");
}

function readFileStore() {
  const raw = readRawFileStore();
  if (!raw || typeof raw !== "object") {
    const empty = createEmptyWorksetStore();
    writeFileStore(empty);
    return empty;
  }
  const next = createEmptyWorksetStore();
  const input = raw.worksets && typeof raw.worksets === "object" ? raw.worksets : {};
  for (const [worksetKey, manifest] of Object.entries(input)) {
    const normalized = normalizeGroupWorksetManifest({
      ...manifest,
      worksetKey,
    });
    if (!normalized) continue;
    next.worksets[worksetKey] = normalized;
  }
  writeFileStore(next);
  return next;
}

function pickPreferredManifest(left, right) {
  const leftManifest = normalizeGroupWorksetManifest(left);
  const rightManifest = normalizeGroupWorksetManifest(right);
  if (!leftManifest) return rightManifest;
  if (!rightManifest) return leftManifest;
  const leftUpdated = Date.parse(leftManifest.updatedAtIso || leftManifest.createdAtIso || "");
  const rightUpdated = Date.parse(rightManifest.updatedAtIso || rightManifest.createdAtIso || "");
  if (Number.isFinite(leftUpdated) && Number.isFinite(rightUpdated) && leftUpdated !== rightUpdated) {
    return leftUpdated > rightUpdated ? leftManifest : rightManifest;
  }
  return buildGroupWorksetPayloadScore(leftManifest) >= buildGroupWorksetPayloadScore(rightManifest)
    ? leftManifest
    : rightManifest;
}

async function ensureGroupWorksetDb() {
  if (!db.isEnabled()) return;
  if (dbInitPromise) return dbInitPromise;
  dbInitPromise = (async () => {
    await db.query(`
      CREATE TABLE IF NOT EXISTS crm_group_worksets (
        workset_key TEXT PRIMARY KEY,
        anchor_email_key TEXT NOT NULL,
        storage_mode TEXT NOT NULL,
        working_group_id TEXT,
        working_group_name TEXT,
        included_email_keys_json JSONB NOT NULL DEFAULT '[]'::jsonb,
        filters_json JSONB NOT NULL DEFAULT '{}'::jsonb,
        attachments_json JSONB NOT NULL DEFAULT '[]'::jsonb,
        main_location_json JSONB NOT NULL DEFAULT '{}'::jsonb,
        remote_promotion_location_json JSONB NULL,
        promotion_json JSONB NOT NULL DEFAULT '{}'::jsonb,
        manifest_json JSONB NOT NULL DEFAULT '{}'::jsonb,
        payload_score INTEGER NOT NULL DEFAULT 0,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
      );
    `);
    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_group_worksets_anchor_email_key ON crm_group_worksets (anchor_email_key);`);
    await db.query(`CREATE INDEX IF NOT EXISTS idx_crm_group_worksets_updated_at ON crm_group_worksets (updated_at DESC);`);
  })();
  return dbInitPromise;
}

function rowToManifest(row) {
  if (!row) return null;
  const manifestJson = row.manifest_json && typeof row.manifest_json === "object"
    ? row.manifest_json
    : {};
  return normalizeGroupWorksetManifest({
    ...manifestJson,
    worksetKey: row.workset_key || manifestJson.worksetKey,
    anchorEmailKey: row.anchor_email_key || manifestJson.anchorEmailKey,
    storageMode: row.storage_mode || manifestJson.storageMode,
    workingGroupId: row.working_group_id || manifestJson.workingGroupId,
    workingGroupName: row.working_group_name || manifestJson.workingGroupName,
    includedEmailKeys: row.included_email_keys_json || manifestJson.includedEmailKeys,
    filters: row.filters_json || manifestJson.filters,
    attachments: row.attachments_json || manifestJson.attachments,
    mainLocation: row.main_location_json || manifestJson.mainLocation,
    remotePromotionLocation: row.remote_promotion_location_json || manifestJson.remotePromotionLocation,
    promotion: row.promotion_json || manifestJson.promotion,
    createdAtIso: manifestJson.createdAtIso || row.created_at,
    updatedAtIso: manifestJson.updatedAtIso || row.updated_at,
  });
}

async function getDbGroupWorksetManifest(worksetKey) {
  await ensureGroupWorksetDb();
  const result = await db.query(
    `SELECT * FROM crm_group_worksets WHERE workset_key = $1 LIMIT 1`,
    [worksetKey]
  );
  return rowToManifest(result?.rows?.[0] || null);
}

async function upsertDbGroupWorksetManifest(manifest) {
  await ensureGroupWorksetDb();
  const payloadScore = buildGroupWorksetPayloadScore(manifest);
  await db.query(
    `
      INSERT INTO crm_group_worksets (
        workset_key,
        anchor_email_key,
        storage_mode,
        working_group_id,
        working_group_name,
        included_email_keys_json,
        filters_json,
        attachments_json,
        main_location_json,
        remote_promotion_location_json,
        promotion_json,
        manifest_json,
        payload_score,
        created_at,
        updated_at
      ) VALUES (
        $1, $2, $3, $4, $5, $6::jsonb, $7::jsonb, $8::jsonb, $9::jsonb, $10::jsonb, $11::jsonb, $12::jsonb, $13, $14::timestamp, $15::timestamp
      )
      ON CONFLICT (workset_key) DO UPDATE SET
        anchor_email_key = EXCLUDED.anchor_email_key,
        storage_mode = EXCLUDED.storage_mode,
        working_group_id = EXCLUDED.working_group_id,
        working_group_name = EXCLUDED.working_group_name,
        included_email_keys_json = EXCLUDED.included_email_keys_json,
        filters_json = EXCLUDED.filters_json,
        attachments_json = EXCLUDED.attachments_json,
        main_location_json = EXCLUDED.main_location_json,
        remote_promotion_location_json = EXCLUDED.remote_promotion_location_json,
        promotion_json = EXCLUDED.promotion_json,
        manifest_json = EXCLUDED.manifest_json,
        payload_score = GREATEST(crm_group_worksets.payload_score, EXCLUDED.payload_score),
        updated_at = EXCLUDED.updated_at
    `,
    [
      manifest.worksetKey,
      manifest.anchorEmailKey,
      manifest.storageMode,
      manifest.workingGroupId || null,
      manifest.workingGroupName || null,
      JSON.stringify(manifest.includedEmailKeys || []),
      JSON.stringify(manifest.filters || {}),
      JSON.stringify(manifest.attachments || []),
      JSON.stringify(manifest.mainLocation || {}),
      JSON.stringify(manifest.remotePromotionLocation || null),
      JSON.stringify(manifest.promotion || {}),
      JSON.stringify(manifest),
      payloadScore,
      manifest.createdAtIso,
      manifest.updatedAtIso,
    ]
  );
}

function writeWorksetToFileStore(manifest) {
  const store = readFileStore();
  store.worksets[manifest.worksetKey] = manifest;
  writeFileStore(store);
}

export async function getGroupWorksetManifest(worksetKey) {
  const normalizedKey = String(worksetKey || "").trim();
  if (!normalizedKey) return null;

  const fileStore = readFileStore();
  const fileManifest = normalizeGroupWorksetManifest(fileStore.worksets[normalizedKey] || null);
  if (!db.isEnabled()) return fileManifest;

  try {
    const dbManifest = await getDbGroupWorksetManifest(normalizedKey);
    return pickPreferredManifest(fileManifest, dbManifest);
  } catch (error) {
    if (error?.optionalDbFallback) {
      console.warn("[groupWorksetStore] DB read failed, using file fallback:", error.message);
      return fileManifest;
    }
    throw error;
  }
}

export async function saveGroupWorksetManifest(input) {
  const incoming = normalizeGroupWorksetManifest(input);
  if (!incoming) {
    throw new Error("Manifesto de workset invalido.");
  }

  const current = await getGroupWorksetManifest(incoming.worksetKey);
  const merged = mergeGroupWorksetManifest(current, incoming);
  if (!merged) {
    throw new Error("Manifesto de workset invalido.");
  }

  if (db.isEnabled()) {
    try {
      await upsertDbGroupWorksetManifest(merged);
    } catch (error) {
      if (!error?.optionalDbFallback) throw error;
      console.warn("[groupWorksetStore] DB save failed, keeping file fallback:", error.message);
    }
  }

  writeWorksetToFileStore(merged);
  return merged;
}
