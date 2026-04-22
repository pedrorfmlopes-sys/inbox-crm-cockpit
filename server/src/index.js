import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";
import { simpleParser } from "mailparser";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

import express from "express";
import cors from "cors";
import { odooClientFromEnv } from "./odoo.js";
import {
  addEmailToGroup,
  addLink,
  createGroupTicket,
  createGroupTicketSeries,
  createCustomGroup,
  deleteGroupTicketSeries,
  deleteDocumentFromGroup,
  deleteCustomGroup,
  detectGroupTicketsForEmail,
  getDocumentContentFromGroup,
  getEmailAttachmentContent,
  getRelatedEmails,
  listAttachmentFlagsByGroup,
  listDocumentsByGroup,
  listCustomGroups,
  listGroupTicketSeries,
  listGroupTickets,
  listKnownEmails,
  listEmailsByGroup,
  listLinksByConversation,
  listLinksByRecord,
  linkEmailToGroupTicket,
  unlinkEmailFromGroupTicket,
  removeEmailFromGroup,
  registerRelevantEmail,
  saveAttachmentFlagsToGroup,
  saveDocumentsToGroup,
  updateGroupTicket,
  updateGroupTicketSeries,
  updateCustomGroup,
} from "./linkStore.js";
import {
  getGroupWorksetManifest,
  migrateGroupWorksetManifest,
  saveGroupWorksetManifest,
  validateGroupStorageLocation,
} from "./groupWorksetStore.js";
import { extractTextFromPdfBuffer } from "./ai/pdfHelper.js";
import { createAiRouter } from "./routes/aiRoutes.js";
import { createLearningRouter } from "./routes/learningRoutes.js";
import fs from "fs";
import { sessionManager } from "./sessionManager.js";

// --- crash visibility (avoid silent exit) ---
process.on("uncaughtException", (err) => {
  console.error("[server] uncaughtException", err);
});
process.on("unhandledRejection", (reason) => {
  console.error("[server] unhandledRejection", reason);
});

const app = express();
app.use(cors());
app.use(express.json({ limit: "80mb" }));

app.use("/api/links", (_req, res, next) => {
  res.set("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
  res.set("Pragma", "no-cache");
  res.set("Expires", "0");
  next();
});

// AI (email assistant)
app.use("/api/ai", createAiRouter());
app.use("/api/learning", createLearningRouter());

const port = process.env.PORT ? Number(process.env.PORT) : 7071;

function normalizeExternalBaseUrl(url) {
  return String(url || "").trim().replace(/\/+$/, "");
}

function stripInlineBase64Prefix(content) {
  return String(content || "").trim().replace(/^data:[^,]+,/, "");
}

function stripHtmlLight(html) {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;|&#039;/gi, "'")
    .replace(/&quot;/gi, "\"")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function truncateExtractedAttachmentText(text, maxChars = 6000) {
  const value = String(text || "").trim();
  if (!value) return "";
  return value.length > maxChars ? `${value.slice(0, maxChars)}\n\n[texto truncado para deteccao]` : value;
}

async function extractTextFromAttachmentPayload(file) {
  const name = String(file?.name || "").trim();
  const contentType = String(file?.contentType || file?.type || "").trim().toLowerCase();
  const content = stripInlineBase64Prefix(file?.content);
  if (!content) return "";
  try {
    const buffer = Buffer.from(content, "base64");
    const lowerName = name.toLowerCase();
    if (contentType === "application/pdf" || lowerName.endsWith(".pdf")) {
      return truncateExtractedAttachmentText(await extractTextFromPdfBuffer(buffer));
    }
    if (
      contentType.startsWith("text/")
      || /json|xml|csv|html|message\/rfc822/.test(contentType)
      || /\.(txt|csv|json|xml|html?|eml)$/i.test(name)
    ) {
      const decoded = buffer.toString("utf8");
      const normalized = /html/.test(contentType) || /\.html?$/i.test(name) ? stripHtmlLight(decoded) : decoded;
      return truncateExtractedAttachmentText(normalized);
    }
  } catch (error) {
    console.warn("[attachments] Failed to extract attachment text", { name, error: error?.message || error });
  }
  return "";
}

function buildInvoiceStudioBatchId(seed) {
  const normalizedSeed = String(seed || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 48);
  const random = Math.random().toString(36).slice(2, 8);
  return `icc-${normalizedSeed || "batch"}-${Date.now()}-${random}`;
}

async function parseRemoteResponse(response) {
  const contentType = String(response.headers.get("content-type") || "").toLowerCase();
  if (contentType.includes("application/json")) {
    return await response.json();
  }
  return await response.text();
}

async function invoiceStudioLogin({ baseUrl, email, password }) {
  const normalizedBaseUrl = normalizeExternalBaseUrl(baseUrl);
  if (!normalizedBaseUrl || !email || !password) {
    throw new Error("Configuracao InvoiceStudio incompleta.");
  }

  const response = await fetch(`${normalizedBaseUrl}/api/auth/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });

  const body = await parseRemoteResponse(response);
  if (!response.ok || !body?.token) {
    const message = typeof body === "string"
      ? body
      : body?.message || body?.error || JSON.stringify(body);
    throw new Error(`InvoiceStudio login falhou: ${message}`);
  }

  return {
    token: String(body.token || "").trim(),
    baseUrl: normalizedBaseUrl,
  };
}

async function invoiceStudioUploadBatch({ baseUrl, token, project, batchId, files }) {
  const targetUrl = new URL(`${baseUrl}/api/extract`);
  if (project) targetUrl.searchParams.set("project", String(project).trim());
  if (batchId) targetUrl.searchParams.set("batchId", String(batchId).trim());

  const form = new FormData();
  for (const file of files || []) {
    const name = String(file?.name || "").trim();
    const content = String(file?.content || "").trim().replace(/^data:[^,]+,/, "");
    const type = String(file?.type || "application/pdf").trim() || "application/pdf";
    if (!name || !content) continue;
    const blob = new Blob([Buffer.from(content, "base64")], { type });
    form.append("files", blob, name);
  }

  const response = await fetch(targetUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
    },
    body: form,
  });

  const body = await parseRemoteResponse(response);
  if (!response.ok) {
    const message = typeof body === "string"
      ? body
      : body?.message || body?.error || JSON.stringify(body);
    throw new Error(`InvoiceStudio upload falhou: ${message}`);
  }

  return body;
}

async function invoiceStudioGetBatchStatus({ baseUrl, token, project, batchId }) {
  const progressUrl = new URL(`${baseUrl}/api/progress/${encodeURIComponent(batchId)}`);
  if (project) progressUrl.searchParams.set("project", String(project).trim());

  const progressResponse = await fetch(progressUrl, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const progressBody = await parseRemoteResponse(progressResponse);
  if (!progressResponse.ok) {
    const message = typeof progressBody === "string"
      ? progressBody
      : progressBody?.message || progressBody?.error || JSON.stringify(progressBody);
    throw new Error(`InvoiceStudio progress falhou: ${message}`);
  }

  let rows = [];
  const statusLabel = String(progressBody?.status || "").trim().toLowerCase();
  if (statusLabel === "finished" || Number(progressBody?.done || 0) >= Number(progressBody?.total || 0)) {
    const batchUrl = new URL(`${baseUrl}/api/batch/${encodeURIComponent(batchId)}`);
    if (project) batchUrl.searchParams.set("project", String(project).trim());
    const batchResponse = await fetch(batchUrl, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const batchBody = await parseRemoteResponse(batchResponse);
    if (!batchResponse.ok) {
      const message = typeof batchBody === "string"
        ? batchBody
        : batchBody?.message || batchBody?.error || JSON.stringify(batchBody);
      throw new Error(`InvoiceStudio batch falhou: ${message}`);
    }
    rows = Array.isArray(batchBody?.rows) ? batchBody.rows : [];
  }

  return {
    progress: progressBody,
    rows,
  };
}

// --- Odoo Client Cache (Singleton) ---
let cachedOdoo = null;
let odooInitPromise = null;

async function getOdooCached(req) {
  // 1. Try session token from Authorization header
  const authHeader = req.headers.authorization;
  let hadSessionToken = false;
  if (authHeader && authHeader.startsWith("Session ")) {
    hadSessionToken = true;
    const token = authHeader.split(" ")[1];
    const session = sessionManager.getSession(token);
    if (session) return session.client || session;
    throw Object.assign(new Error("Session expired"), { status: 401, code: "SESSION_EXPIRED" });
  }

  // 2. Fallback to global singleton (from .env)
  if (cachedOdoo) return cachedOdoo;
  if (odooInitPromise) return await odooInitPromise;

  odooInitPromise = (async () => {
    try {
      const client = await odooClientFromEnv();
      cachedOdoo = client;
      return client;
    } catch (e) {
      if (hadSessionToken) {
        throw Object.assign(new Error("Session expired"), { status: 401, code: "SESSION_EXPIRED", cause: e });
      }
      odooInitPromise = null; // allow retry
      throw e;
    } finally {
      odooInitPromise = null;
    }
  })();

  return await odooInitPromise;
}

app.get("/health", (_req, res) => res.json({ ok: true }));

app.post("/api/invoice-studio/upload", async (req, res) => {
  try {
    const baseUrl = String(req.body?.baseUrl || "").trim();
    const email = String(req.body?.email || "").trim();
    const password = String(req.body?.password || "").trim();
    const project = String(req.body?.project || "").trim();
    const metadata = req.body?.metadata || {};
    const files = Array.isArray(req.body?.files)
      ? req.body.files
        .map((file) => ({
          name: String(file?.name || "").trim(),
          type: String(file?.type || "").trim(),
          content: String(file?.content || "").trim(),
        }))
        .filter((file) => file.name && file.content)
      : [];

    if (!baseUrl || !email || !password) {
      return res.status(400).json({ ok: false, error: "InvoiceStudio configuracao incompleta." });
    }
    if (!files.length) {
      return res.status(400).json({ ok: false, error: "Sem ficheiros para enviar ao InvoiceStudio." });
    }

    const auth = await invoiceStudioLogin({ baseUrl, email, password });
    const batchId = String(req.body?.batchId || "").trim()
      || buildInvoiceStudioBatchId(
        metadata?.internetMessageId
        || metadata?.itemId
        || metadata?.conversationId
        || metadata?.subject
        || files[0]?.name
      );

    const upload = await invoiceStudioUploadBatch({
      baseUrl: auth.baseUrl,
      token: auth.token,
      project,
      batchId,
      files,
    });

    return res.json({
      ok: true,
      batchId: String(upload?.batchId || batchId),
      count: Number(upload?.count || files.length),
      project: String(upload?.project || project || "").trim() || undefined,
      status: String(upload?.status || "processing"),
      upload,
    });
  } catch (error) {
    console.error("[server] InvoiceStudio upload failed:", error);
    return res.status(500).json({ ok: false, error: error?.message || "Falha ao enviar documentos para o InvoiceStudio." });
  }
});

app.post("/api/invoice-studio/status", async (req, res) => {
  try {
    const baseUrl = String(req.body?.baseUrl || "").trim();
    const email = String(req.body?.email || "").trim();
    const password = String(req.body?.password || "").trim();
    const project = String(req.body?.project || "").trim();
    const batchId = String(req.body?.batchId || "").trim();

    if (!baseUrl || !email || !password || !batchId) {
      return res.status(400).json({ ok: false, error: "Pedido de estado InvoiceStudio incompleto." });
    }

    const auth = await invoiceStudioLogin({ baseUrl, email, password });
    const status = await invoiceStudioGetBatchStatus({
      baseUrl: auth.baseUrl,
      token: auth.token,
      project,
      batchId,
    });

    return res.json({
      ok: true,
      batchId,
      project: project || undefined,
      progress: status.progress || {},
      rows: Array.isArray(status.rows) ? status.rows : [],
    });
  } catch (error) {
    console.error("[server] InvoiceStudio status failed:", error);
    return res.status(500).json({ ok: false, error: error?.message || "Falha ao consultar o estado no InvoiceStudio." });
  }
});

// --- AUTH ---
app.post("/api/auth/login", async (req, res) => {
  try {
    const credentials = req.body;
    const session = await sessionManager.createSession(credentials);
    return res.json({ ok: true, ...session });
  } catch (e) {
    console.warn("[server] Login failed:", e?.message);
    return res.status(401).json({ ok: false, message: e?.message || "Login falhou" });
  }
});

app.get("/api/auth/check", async (req, res) => {
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith("Session ")) {
    const token = authHeader.split(" ")[1];
    const session = sessionManager.getSession(token);
    if (session?.client?.meta) return res.json({ ok: true, authenticated: true, meta: session.client.meta });
  }
  return res.json({ ok: true, authenticated: false, reason: "no_session" });
});

app.get("/api/odoo/meta", async (req, res) => {
  try {
    const odoo = await getOdooCached(req);
    return res.json({ ok: true, meta: odoo.meta });
  } catch (e) {
    console.warn("[server] Odoo Meta lookup failed:", e?.message);
    return res.json({ ok: false, message: "Odoo não configurado ou inacessível." });
  }
});

app.get("/api/odoo/ping", async (req, res) => {
  try {
    const odoo = await getOdooCached(req);
    const ok = await odoo.ping();
    return res.json({ ok });
  } catch (e) {
    return res.json({ ok: false, error: e?.message });
  }
});

/**
 * ✅ Odoo Bridge: Seamless Login
 * Receives a session token, fetches credentials, and renders an auto-submitting form.
 */
app.get("/api/odoo/auto-login", async (req, res) => {
  try {
    const token = String(req.query.token || "").trim();
    const redirect = String(req.query.redirect || "/web").trim();

    let credentials = null;

    // 1. Try Session Manager
    if (token) {
      const session = sessionManager.getSession(token);
      if (session?.credentials) {
        credentials = session.credentials;
      }
    }

    // 2. Fallback to Env if no session (only if Odoo is configured in .env)
    if (!credentials) {
      credentials = {
        url: process.env.ODOO_URL,
        db: process.env.ODOO_DB,
        login: process.env.ODOO_USERNAME || process.env.ODOO_USER,
        password: process.env.ODOO_API_KEY || process.env.ODOO_PASS || process.env.ODOO_PASSWORD
      };
      // Basic check if it's usable
      if (!credentials.url || credentials.url.includes("your-odoo-instance.com")) {
        return res.status(401).send("Sessão expirada ou Odoo não configurado. Por favor, faz login novamente.");
      }
    }

    const { url, db, login, password } = credentials;
    const loginUrl = `${url.replace(/\/+$/, "")}/web/login`;

    // Render Auto-Login Page
    const html = `
<!DOCTYPE html>
<html>
<head>
    <title>A entrar no Odoo...</title>
    <style>
        body { font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; margin: 0; color: #666; }
        .spinner { border: 4px solid #f3f3f3; border-top: 4px solid #714B67; border-radius: 50%; width: 40px; height: 40px; animation: spin 2s linear infinite; margin-bottom: 20px; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    </style>
</head>
<body>
    <div class="spinner"></div>
    <p>A iniciar sessão segura no Odoo...</p>
    
    <form id="loginForm" action="${loginUrl}" method="POST" style="display:none;">
        <input type="hidden" name="db" value="${escapeHtml(db)}">
        <input type="hidden" name="login" value="${escapeHtml(login)}">
        <input type="hidden" name="password" value="${escapeHtml(password)}">
        <input type="hidden" name="redirect" value="${escapeHtml(redirect)}">
    </form>

    <script>
        document.getElementById('loginForm').submit();
    </script>
</body>
</html>
    `;

    res.send(html);

  } catch (e) {
    console.error("[server] Bridge failed:", e);
    res.status(500).send("Erro ao iniciar sessão no Odoo.");
  }
});

// ✅ Alargar (Jira-like): projetos, leads, contactos, tarefas, utilizadores, etapas
const MODEL_WHITELIST = new Set([
  "project.project",
  "crm.lead",
  "res.partner",
  "res.country",
  "res.country.state",
  "project.task",
  "helpdesk.ticket",
  "helpdesk.team",
  "helpdesk.stage",
  "res.users",
  "project.task.type",
  "ir.attachment",
  "crm.stage",
]);

function modelAllowed(model) {
  return MODEL_WHITELIST.has(String(model || "").trim());
}

const PARTNER_PREFERRED_READ_FIELDS = [
  "id", "name", "email", "phone", "mobile", "function", "company_type", "is_company",
  "parent_id", "vat", "street", "street2", "zip", "city", "state_id", "country_id",
  "website", "display_name"
];

async function safeSearchReadCompat(odoo, model, domain, wantedFields, limit = 10, order) {
  if (typeof odoo.safeSearchRead === "function") {
    return await odoo.safeSearchRead(model, domain, wantedFields, limit, order);
  }
  return await odoo.searchRead(model, domain, wantedFields, limit, order);
}

async function safeCreateCompat(odoo, model, data) {
  if (typeof odoo.safeCreate === "function") {
    return await odoo.safeCreate(model, data);
  }
  return await odoo.create(model, data);
}

async function safeWriteCompat(odoo, model, ids, data) {
  if (typeof odoo.safeWrite === "function") {
    return await odoo.safeWrite(model, ids, data);
  }
  return await odoo.call(model, "write", [Array.isArray(ids) ? ids : [ids], data]);
}

async function filterReadFieldsCompat(odoo, model, fields) {
  if (!Array.isArray(fields)) return [];
  if (odoo?.schema?.filterReadFields) {
    return await odoo.schema.filterReadFields(model, fields);
  }
  return fields;
}

async function safeReadCompat(odoo, model, ids, wantedFields) {
  const filtered = await filterReadFieldsCompat(odoo, model, wantedFields);
  const fallback = filtered.length ? filtered : ["id"];
  try {
    return await odoo.read(model, ids, fallback);
  } catch (e) {
    const msg = String(e?.message || "");
    if (!/invalid field/i.test(msg)) throw e;
    if (odoo?.schema?.invalidateModel) {
      odoo.schema.invalidateModel(model);
      const refreshed = await filterReadFieldsCompat(odoo, model, wantedFields);
      const fields2 = refreshed.length ? refreshed : ["id"];
      return await odoo.read(model, ids, fields2);
    }
    throw e;
  }
}

function normalizeEmail(raw) {
  return String(raw || "").trim().toLowerCase();
}

function normalizePartnerPayload(data = {}) {
  const clean = {};
  if (typeof data.name === "string" && data.name.trim()) clean.name = data.name.trim();

  const email = normalizeEmail(data.email);
  if (email) clean.email = email;

  if (data.company_type === "person" || data.company_type === "company") {
    clean.company_type = data.company_type;
  }

  if (Object.prototype.hasOwnProperty.call(data, "parent_id")) {
    const pid = Number(data.parent_id);
    clean.parent_id = pid > 0 ? pid : false;
  }

  if (typeof data.function === "string") clean.function = data.function.trim();
  if (typeof data.phone === "string") clean.phone = data.phone.trim();
  if (typeof data.mobile === "string") clean.mobile = data.mobile.trim();
  if (typeof data.vat === "string") clean.vat = data.vat.trim().toUpperCase().replace(/\s+/g, "");
  if (typeof data.street === "string") clean.street = data.street.trim();
  if (typeof data.street2 === "string") clean.street2 = data.street2.trim();
  if (typeof data.zip === "string") clean.zip = data.zip.trim();
  if (typeof data.city === "string") clean.city = data.city.trim();
  if (typeof data.website === "string") clean.website = data.website.trim();

  if (Object.prototype.hasOwnProperty.call(data, "country_id")) {
    const countryId = Number(data.country_id);
    clean.country_id = countryId > 0 ? countryId : false;
  }

  if (Object.prototype.hasOwnProperty.call(data, "state_id")) {
    const stateId = Number(data.state_id);
    clean.state_id = stateId > 0 ? stateId : false;
  }

  return clean;
}

async function findPartnerExactByEmail(odoo, email) {
  const normalized = normalizeEmail(email);
  if (!normalized) return null;

  const rows = await safeSearchReadCompat(
    odoo,
    "res.partner",
    [["email", "=", normalized]],
    PARTNER_PREFERRED_READ_FIELDS,
    1
  );

  return Array.isArray(rows) && rows.length ? rows[0] : null;
}

function many2oneLabel(value) {
  return Array.isArray(value) ? String(value[1] || "").trim() : "";
}

function formatIsoDate(raw) {
  const value = String(raw || "").trim();
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return date.toLocaleDateString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  });
}

function formatCurrencyAmount(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return "";
  return new Intl.NumberFormat("pt-PT", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  }).format(amount);
}

function isIgnorableOdooLookupError(error) {
  const message = String(error?.message || error || "");
  return /invalid field|keyerror|unknown|does not exist|doesn't exist|not found|accesserror|access denied|forbidden/i.test(message);
}

async function getModelFieldsSafe(odoo, model) {
  try {
    return await odoo.schema.getModelFields(model);
  } catch (error) {
    if (isIgnorableOdooLookupError(error)) return null;
    throw error;
  }
}

async function readPartnerRelationSection(odoo, spec, partnerId) {
  const fields = await getModelFieldsSafe(odoo, spec.model);
  if (!fields) return null;
  if ((spec.requiredFields || []).some((fieldName) => !fields.has(fieldName))) return null;

  try {
    const rows = await safeSearchReadCompat(
      odoo,
      spec.model,
      spec.domain(partnerId),
      spec.fields,
      spec.limit || 6,
      spec.order,
    );

    const items = (Array.isArray(rows) ? rows : []).map((row) => spec.mapRow(row)).filter(Boolean);
    if (!items.length) return null;

    return {
      key: spec.key,
      label: spec.label,
      model: spec.model,
      total: items.length,
      items,
    };
  } catch (error) {
    if (isIgnorableOdooLookupError(error)) return null;
    throw error;
  }
}

async function listPartnerNativeRelations(odoo, partnerId) {
  const specs = [
    {
      key: "leads",
      label: "Leads e oportunidades",
      model: "crm.lead",
      requiredFields: ["partner_id"],
      fields: ["name", "type", "stage_id", "user_id", "probability", "expected_revenue"],
      order: "write_date desc, id desc",
      domain: (id) => [["partner_id", "=", id]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        const kind = row?.type === "opportunity" ? "Oportunidade" : "Lead";
        const stage = many2oneLabel(row?.stage_id);
        const owner = many2oneLabel(row?.user_id);
        const probability = Number(row?.probability);
        const revenue = Number(row?.expected_revenue);
        return {
          model: "crm.lead",
          recordId,
          title: String(row?.name || "").trim() || `Lead #${recordId}`,
          meta: [kind, stage].filter(Boolean).join(" · "),
          secondary: [
            owner,
            Number.isFinite(probability) ? `${Math.round(probability)}%` : "",
            Number.isFinite(revenue) && revenue > 0 ? `${formatCurrencyAmount(revenue)} EUR` : "",
          ].filter(Boolean).join(" · "),
        };
      },
    },
    {
      key: "projects",
      label: "Projetos",
      model: "project.project",
      requiredFields: ["partner_id"],
      fields: ["name", "user_id"],
      order: "write_date desc, id desc",
      domain: (id) => [["partner_id", "=", id]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        return {
          model: "project.project",
          recordId,
          title: String(row?.name || "").trim() || `Projeto #${recordId}`,
          meta: "Projeto",
          secondary: many2oneLabel(row?.user_id),
        };
      },
    },
    {
      key: "tasks",
      label: "Tarefas",
      model: "project.task",
      requiredFields: ["partner_id"],
      fields: ["name", "project_id", "stage_id", "date_deadline"],
      order: "write_date desc, id desc",
      domain: (id) => [["partner_id", "=", id]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        const project = many2oneLabel(row?.project_id);
        const stage = many2oneLabel(row?.stage_id);
        const deadline = formatIsoDate(row?.date_deadline);
        return {
          model: "project.task",
          recordId,
          title: String(row?.name || "").trim() || `Tarefa #${recordId}`,
          meta: [project, stage].filter(Boolean).join(" · "),
          secondary: deadline ? `Prazo ${deadline}` : "",
        };
      },
    },
    {
      key: "tickets",
      label: "Tickets",
      model: "helpdesk.ticket",
      requiredFields: ["partner_id"],
      fields: ["name", "stage_id", "team_id", "user_id", "priority"],
      order: "write_date desc, id desc",
      domain: (id) => [["partner_id", "=", id]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        const stage = many2oneLabel(row?.stage_id);
        const team = many2oneLabel(row?.team_id);
        const owner = many2oneLabel(row?.user_id);
        const priority = String(row?.priority || "").trim();
        return {
          model: "helpdesk.ticket",
          recordId,
          title: String(row?.name || "").trim() || `Ticket #${recordId}`,
          meta: [stage, team].filter(Boolean).join(" · "),
          secondary: [owner, priority ? `Prioridade ${priority}` : ""].filter(Boolean).join(" · "),
        };
      },
    },
    {
      key: "sales",
      label: "Vendas",
      model: "sale.order",
      requiredFields: ["partner_id"],
      fields: ["name", "state", "user_id", "date_order", "amount_total"],
      order: "date_order desc, id desc",
      domain: (id) => [["partner_id", "=", id]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        const state = String(row?.state || "").trim().replace(/_/g, " ");
        const owner = many2oneLabel(row?.user_id);
        const orderedAt = formatIsoDate(row?.date_order);
        const amount = formatCurrencyAmount(row?.amount_total);
        return {
          model: "sale.order",
          recordId,
          title: String(row?.name || "").trim() || `Venda #${recordId}`,
          meta: [state, owner].filter(Boolean).join(" · "),
          secondary: [orderedAt ? `Data ${orderedAt}` : "", amount ? `${amount} EUR` : ""].filter(Boolean).join(" · "),
        };
      },
    },
    {
      key: "meetings",
      label: "Reunioes",
      model: "calendar.event",
      requiredFields: ["partner_ids"],
      fields: ["name", "start", "stop", "user_id"],
      order: "start desc, id desc",
      domain: (id) => [["partner_ids", "in", [id]]],
      mapRow: (row) => {
        const recordId = Number(row?.id || 0);
        if (!recordId) return null;
        const owner = many2oneLabel(row?.user_id);
        const start = formatIsoDate(row?.start);
        const stop = formatIsoDate(row?.stop);
        return {
          model: "calendar.event",
          recordId,
          title: String(row?.name || "").trim() || `Reuniao #${recordId}`,
          meta: [start ? `Inicio ${start}` : "", stop && stop !== start ? `Fim ${stop}` : ""].filter(Boolean).join(" · "),
          secondary: owner,
        };
      },
    },
  ];

  const sections = [];
  for (const spec of specs) {
    const section = await readPartnerRelationSection(odoo, spec, partnerId);
    if (section) sections.push(section);
  }
  return sections;
}

function normalizeStudioLabel(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function decodeXmlEntities(value) {
  return String(value || "")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function extractFormTabTitles(viewArch) {
  const matches = String(viewArch || "").matchAll(/<page\b[^>]*\bstring=(["'])(.*?)\1/gi);
  const titles = [];
  for (const match of matches) {
    const decoded = decodeXmlEntities(match?.[2] || "").trim();
    if (decoded) titles.push(decoded);
  }
  return titles;
}

function fieldExistsInView(viewArch, fieldName) {
  const normalizedFieldName = String(fieldName || "").trim();
  if (!normalizedFieldName) return false;
  const pattern = new RegExp(`<field\\b[^>]*\\bname=(["'])${normalizedFieldName}\\1`, "i");
  return pattern.test(String(viewArch || ""));
}

function createLayoutCheck({
  key,
  label,
  kind,
  configuredName,
  status,
  message,
  details,
  actualType,
  expectedTypes,
  recommendedType,
  presentInFormView,
}) {
  return {
    key,
    label,
    kind,
    configuredName,
    status,
    message,
    details,
    actualType,
    expectedTypes,
    recommendedType,
    presentInFormView,
  };
}

app.post("/api/odoo/layout/validate", async (req, res) => {
  try {
    const layout = req.body?.layout || {};
    const requestedTarget = String(req.body?.target || "project").trim().toLowerCase();
    const target = requestedTarget === "lead"
      ? "lead"
      : requestedTarget === "task"
        ? "task"
        : requestedTarget === "ticket"
          ? "ticket"
          : "project";
    const targetConfig = layout?.[target] || {};
    const mode = String(targetConfig?.mode || layout?.mode || "description_only").trim() === "structured_project"
      ? "structured_project"
      : "description_only";
    const modelFallback = target === "lead"
      ? "crm.lead"
      : target === "task"
        ? "project.task"
        : target === "ticket"
          ? "helpdesk.ticket"
          : "project.project";
    const model = String(targetConfig?.model || modelFallback).trim() || modelFallback;

    if (!modelAllowed(model)) {
      return res.status(400).json({ ok: false, error: "invalid_model", message: `Model not allowed: ${model}` });
    }

    const odoo = await getOdooCached(req);
    const fieldMeta = await odoo.call(model, "fields_get", [], {
      attributes: ["string", "type", "relation", "readonly", "required"],
    });

    let formViewArch = "";
    let formViewAvailable = false;
    let formViewError = "";
    try {
      const formView = await odoo.call(model, "fields_view_get", [], { view_type: "form" });
      formViewArch = String(formView?.arch || "");
      formViewAvailable = Boolean(formViewArch.trim());
    } catch (error) {
      formViewError = String(error?.message || error || "");
    }

    const tabTitles = formViewAvailable ? extractFormTabTitles(formViewArch) : [];
    const normalizedTabTitles = new Set(tabTitles.map(normalizeStudioLabel).filter(Boolean));
    const checks = [];

    function validateField(spec) {
      const configuredName = String(spec.configuredName || "").trim();
      if (!configuredName) {
        checks.push(createLayoutCheck({
          key: spec.key,
          label: spec.label,
          kind: "field",
          configuredName,
          status: "error",
          message: "Campo tecnico por definir.",
          expectedTypes: spec.expectedTypes,
          recommendedType: spec.recommendedType,
        }));
        return;
      }

      const meta = fieldMeta?.[configuredName];
      if (!meta) {
        checks.push(createLayoutCheck({
          key: spec.key,
          label: spec.label,
          kind: "field",
          configuredName,
          status: "error",
          message: "O campo nao existe neste modelo Odoo.",
          details: `Criar no Studio em ${model}.`,
          expectedTypes: spec.expectedTypes,
          recommendedType: spec.recommendedType,
          presentInFormView: false,
        }));
        return;
      }

      const actualType = String(meta?.type || "").trim();
      const expectedTypes = Array.isArray(spec.expectedTypes) ? spec.expectedTypes : [];
      const presentInFormView = formViewAvailable ? fieldExistsInView(formViewArch, configuredName) : undefined;
      let status = expectedTypes.includes(actualType) ? "ok" : "error";
      let message = expectedTypes.includes(actualType)
        ? "Campo encontrado e tipo compativel."
        : `Tipo incompatível: ${actualType || "desconhecido"}.`;

      if (status === "ok" && spec.recommendedType && actualType && actualType !== spec.recommendedType) {
        status = "warning";
        message = `Campo utilizavel, mas o tipo recomendado e ${spec.recommendedType}.`;
      }

      if (status !== "error" && formViewAvailable && presentInFormView === false) {
        status = "warning";
        message = "Campo existe, mas nao esta exposto na vista form atual.";
      }

      checks.push(createLayoutCheck({
        key: spec.key,
        label: spec.label,
        kind: "field",
        configuredName,
        status,
        message,
        actualType,
        expectedTypes,
        recommendedType: spec.recommendedType,
        presentInFormView,
      }));
    }

    function validateTab(spec) {
      const configuredName = String(spec.configuredName || "").trim();
      if (!configuredName) {
        checks.push(createLayoutCheck({
          key: spec.key,
          label: spec.label,
          kind: "tab",
          configuredName,
          status: "warning",
          message: "Etiqueta da aba por definir.",
        }));
        return;
      }

      if (!formViewAvailable) {
        checks.push(createLayoutCheck({
          key: spec.key,
          label: spec.label,
          kind: "tab",
          configuredName,
          status: "warning",
          message: "Nao foi possivel validar a aba na vista form.",
          details: formViewError || "A vista form do modelo nao foi lida.",
        }));
        return;
      }

      const exists = normalizedTabTitles.has(normalizeStudioLabel(configuredName));
      checks.push(createLayoutCheck({
        key: spec.key,
        label: spec.label,
        kind: "tab",
        configuredName,
        status: exists ? "ok" : "warning",
        message: exists
          ? "Aba encontrada na vista form."
          : "Aba nao encontrada na vista form atual.",
      }));
    }

    validateField({
      key: "descriptionField",
      label: "Campo base da descricao",
      configuredName: targetConfig.descriptionField,
      expectedTypes: ["html", "text", "char"],
      recommendedType: "html",
    });

    if (mode === "structured_project") {
      validateField({
        key: "fixedInfoField",
        label: "Campo de informacao fixa",
        configuredName: targetConfig.fixedInfoField,
        expectedTypes: ["html", "text"],
        recommendedType: "html",
      });
      validateField({
        key: "historyField",
        label: "Campo de historico",
        configuredName: targetConfig.historyField,
        expectedTypes: ["html", "text"],
        recommendedType: "html",
      });
      validateField({
        key: "documentsField",
        label: "Campo de documentos",
        configuredName: targetConfig.documentsField,
        expectedTypes: ["html", "text"],
        recommendedType: "html",
      });

      validateTab({
        key: "fixedInfoTabLabel",
        label: "Aba de informacao fixa",
        configuredName: targetConfig.fixedInfoTabLabel,
      });
      validateTab({
        key: "historyTabLabel",
        label: "Aba de historico",
        configuredName: targetConfig.historyTabLabel,
      });
      validateTab({
        key: "documentsTabLabel",
        label: "Aba de documentos",
        configuredName: targetConfig.documentsTabLabel,
      });
    }

    const summary = checks.reduce((acc, check) => {
      if (check.status === "ok") acc.ok += 1;
      else if (check.status === "warning") acc.warning += 1;
      else acc.error += 1;
      return acc;
    }, { ok: 0, warning: 0, error: 0 });

    return res.json({
      ok: true,
      target,
      mode,
      model,
      ready: summary.error === 0,
      summary,
      checks,
      formView: {
        available: formViewAvailable,
        tabTitles,
        error: formViewError || undefined,
      },
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "layout_validation_failed", details: String(e?.message || e) });
  }
});

app.get("/api/odoo/partners/by-email", async (req, res) => {
  try {
    const email = normalizeEmail(req.query.email);
    if (!email) return res.status(400).json({ ok: false, message: "Missing email", partner: null });

    try {
      const odoo = await getOdooCached(req);
      const partner = await findPartnerExactByEmail(odoo, email);
      return res.json({ ok: true, partner: partner || null });
    } catch (e) {
      const message = String(e?.message || e);
      const status = Number(e?.status) || (/session expired/i.test(message) ? 401 : 503);
      console.warn("[server] /partners/by-email failed:", message);
      return res.status(status).json({ ok: false, error: "odoo_unavailable", message, partner: null });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "partner_lookup_failed", details: String(e?.message || e), partner: null });
  }
});

app.get("/api/odoo/partners/:id/relations", async (req, res) => {
  try {
    const partnerId = Number(req.params.id || 0);
    if (!partnerId) {
      return res.status(400).json({ ok: false, message: "Missing partner id", relations: [] });
    }

    const odoo = await getOdooCached(req);
    const partnerRows = await safeReadCompat(odoo, "res.partner", [partnerId], PARTNER_PREFERRED_READ_FIELDS);
    const partner = Array.isArray(partnerRows) && partnerRows.length ? partnerRows[0] : null;
    if (!partner) {
      return res.status(404).json({ ok: false, message: "Partner not found", relations: [] });
    }

    const relations = await listPartnerNativeRelations(odoo, partnerId);
    const total = relations.reduce((sum, section) => sum + Number(section?.total || 0), 0);

    return res.json({
      ok: true,
      partner,
      total,
      relations,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "partner_relations_failed", details: String(e?.message || e), relations: [] });
  }
});

app.get("/api/odoo/companies/search", async (req, res) => {
  try {
    const q = String(req.query.q || "").trim();

    try {
      const odoo = await getOdooCached(req);
      const domain = [["company_type", "=", "company"], ...(q ? [["name", "ilike", q]] : [])];
      const companies = await safeSearchReadCompat(odoo, "res.partner", domain, PARTNER_PREFERRED_READ_FIELDS, 10);
      return res.json({ ok: true, results: companies || [] });
    } catch (e) {
      const message = String(e?.message || e);
      const status = Number(e?.status) || (/session expired/i.test(message) ? 401 : 503);
      console.warn("[server] /companies/search failed:", message);
      return res.status(status).json({ ok: false, error: "odoo_unavailable", message, results: [] });
    }
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "company_search_failed", details: String(e?.message || e), results: [] });
  }
});

app.post("/api/odoo/partners/create-or-update", async (req, res) => {
  try {
    const mode = String(req.body?.mode || "").trim();
    const targetPartnerId = Number(req.body?.targetPartnerId);
    const data = normalizePartnerPayload(req.body?.data || {});

    if (mode !== "create" && mode !== "update") {
      return res.status(400).json({ ok: false, message: "Invalid mode" });
    }

    if (!data.email) {
      return res.status(400).json({ ok: false, message: "Missing email" });
    }

    const odoo = await getOdooCached(req);
    const existing = await findPartnerExactByEmail(odoo, data.email);

    if (mode === "create") {
      if (existing) {
        return res.status(409).json({
          ok: false,
          conflict: true,
          message: "Partner already exists for email; use update",
          existingPartner: existing,
        });
      }

      if (!data.name) data.name = data.email.split("@")[0] || "Contacto";
      if (!data.company_type) data.company_type = "person";

      const id = await safeCreateCompat(odoo, "res.partner", data);
      const partner = await safeReadCompat(odoo, "res.partner", [id], PARTNER_PREFERRED_READ_FIELDS);
      return res.json({ ok: true, mode: "create", id, partner: Array.isArray(partner) ? partner[0] : null });
    }

    const targetId = targetPartnerId || Number(existing?.id);
    if (!targetId) {
      return res.status(404).json({ ok: false, message: "Partner not found to update" });
    }

    if (existing && Number(existing.id) !== Number(targetId)) {
      return res.status(409).json({
        ok: false,
        conflict: true,
        message: "Email already belongs to another partner",
        existingPartner: existing,
      });
    }

    const patch = { ...data };
    if (!Object.keys(patch).length) {
      return res.status(400).json({ ok: false, message: "No fields to update" });
    }

    await safeWriteCompat(odoo, "res.partner", [targetId], patch);
    const partner = await safeReadCompat(odoo, "res.partner", [targetId], PARTNER_PREFERRED_READ_FIELDS);
    return res.json({ ok: true, mode: "update", id: targetId, partner: Array.isArray(partner) ? partner[0] : null });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "partner_create_or_update_failed", details: String(e?.message || e) });
  }
});

app.get("/api/odoo/search", async (req, res) => {
  try {
    const model = String(req.query.model || "").trim();
    const q = String(req.query.q || "").trim();
    const limit = Math.min(Number(req.query.limit || 10), 20);

    if (!modelAllowed(model)) return res.status(400).json({ ok: false, error: "model_not_allowed" });

    const odoo = await getOdooCached(req);

    // Quando a pesquisa está vazia: devolve as primeiras N linhas (útil para dropdown aberto)
    const isEmpty = !q;

    let domain;
    let fields;

    if (model === "res.partner") {
      domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email", "ilike", q]];
      fields = ["name", "email", "phone", "mobile", "display_name", "company_type"];
    } else if (model === "crm.lead") {
      domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email_from", "ilike", q]];
      fields = ["name", "display_name", "email_from", "partner_id"];
    } else if (model === "project.project") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "partner_id", "user_id"];
    } else if (model === "project.task") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "project_id", "parent_id"];
    } else if (model === "helpdesk.ticket") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "partner_id", "stage_id", "team_id", "user_id", "priority"];
    } else if (model === "helpdesk.team") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name"];
    } else if (model === "helpdesk.stage") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "team_ids"];
    } else if (model === "res.users") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "email"];
    } else if (model === "project.task.type" || model === "crm.stage") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name"];
    } else {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name"];
    }

    const items = await safeSearchReadCompat(odoo, model, domain, fields, limit);
    return res.json({ items: items || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

// --- compat endpoints (client expects POST + search-domain/read/write/call) ---
function cleanValuesForModel(model, values) {
  const allowedByModel = {
    "res.partner": new Set([
      "name", "email", "phone", "mobile", "function", "company_type", "parent_id",
      "vat", "is_company", "street", "street2", "zip", "city", "country_id",
      "state_id", "website"
    ]),
    "crm.lead": new Set(["name", "contact_name", "email_from", "phone", "partner_id", "stage_id", "description"]),
    "project.project": new Set(["name", "partner_id", "user_id", "description"]),
    "project.task": new Set(["name", "description", "date_deadline", "project_id", "lead_id", "parent_id", "user_ids", "stage_id"]),
    "helpdesk.ticket": new Set(["name", "description", "partner_id", "team_id", "user_id", "stage_id", "priority"]),
    "ir.attachment": new Set(["name", "datas", "res_model", "res_id", "type", "mimetype", "datas_fname"]),
    "crm.stage": new Set(["name"]),
  }[model];

  if (!allowedByModel) return null;
  if (!values || typeof values !== "object") return null;

  const clean = {};
  for (const [k, v] of Object.entries(values)) {
    if (
      allowedByModel.has(k) ||
      (model === "project.project" && /^x_/i.test(k)) ||
      (model === "project.task" && /^x_/i.test(k)) ||
      (model === "helpdesk.ticket" && /^x_/i.test(k)) ||
      (model === "crm.lead" && (/^x_/i.test(k) || /(^|_)(lead_type|tipo_de_lead|tipo_lead)$/.test(k) || /tipo.*lead/i.test(k)))
    ) {
      clean[k] = v;
    }
  }

  // Normalize M2M for project.task user_ids
  if (model === "project.task" && Array.isArray(clean.user_ids)) {
    const ids = clean.user_ids.map((x) => Number(x)).filter(Boolean);
    if (ids.length) clean.user_ids = [[6, 0, ids]];
    else delete clean.user_ids;
  }

  return clean;
}

function buildSearchSpec(model, q) {
  const isEmpty = !q;
  let domain;
  let fields;

  if (model === "res.partner") {
    domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email", "ilike", q]];
    fields = ["name", "email", "phone", "mobile", "display_name", "company_type"];
  } else if (model === "res.country") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name"];
  } else if (model === "res.country.state") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "country_id"];
  } else if (model === "crm.lead") {
    domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email_from", "ilike", q]];
    fields = ["name", "display_name", "email_from", "partner_id"];
  } else if (model === "project.project") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "partner_id", "user_id"];
  } else if (model === "project.task") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "project_id", "parent_id", "stage_id"];
  } else if (model === "helpdesk.ticket") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "partner_id", "stage_id", "team_id", "user_id", "priority"];
  } else if (model === "helpdesk.team") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name"];
  } else if (model === "helpdesk.stage") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "team_ids"];
  } else if (model === "res.users") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "email"];
  } else if (model === "project.task.type" || model === "crm.stage") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name"];
  } else {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name"];
  }

  return { domain, fields };
}

app.post("/api/odoo/search", async (req, res) => {
  try {
    const body = req.body || {};
    const model = String(body.model || "").trim();

    if (!modelAllowed(model)) return res.status(400).json({ ok: false, error: "model_not_allowed" });

    // Two supported shapes:
    // 1) { model, query, limit }  (free-text)
    // 2) { model, domain, fields, limit, order } (domain search)
    const q = String(body.query ?? body.q ?? "").trim();
    const limit = Math.min(Number(body.limit ?? 20), 80);

    const odoo = await getOdooCached(req);

    if (Array.isArray(body.domain)) {
      const domain = body.domain;
      const fields = Array.isArray(body.fields) ? body.fields : ["id", "name"];
      const order = typeof body.order === "string" ? body.order : undefined;
      const records = await safeSearchReadCompat(odoo, model, domain, fields, limit, order);
      return res.json({ records: records || [] });
    }

    const spec = buildSearchSpec(model, q);
    const records = await safeSearchReadCompat(odoo, model, spec.domain, spec.fields, limit);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

app.post("/api/odoo/search-domain", async (req, res) => {
  try {
    const { model, domain, fields, limit, order } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!Array.isArray(domain)) return res.status(400).json({ ok: false, error: "missing_domain" });

    const lim = Math.min(Number(limit ?? 20), 80);
    const f = Array.isArray(fields) ? fields : ["id", "name"];
    const ord = typeof order === "string" ? order : undefined;

    const odoo = await getOdooCached(req);
    const records = await safeSearchReadCompat(odoo, m, domain, f, lim, ord);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

app.post("/api/odoo/read", async (req, res) => {
  try {
    const { model, ids, fields } = req.body || {};
    const m = String(model || "").trim();
    const idList = (Array.isArray(ids) ? ids : [ids]).map((x) => Number(x)).filter(Boolean).slice(0, 80);

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!idList.length) return res.status(400).json({ ok: false, error: "missing_ids" });

    const f = Array.isArray(fields) ? fields : ["id", "name", "display_name"];

    const odoo = await getOdooCached(req);
    const records = await safeReadCompat(odoo, m, idList, f);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

app.post("/api/odoo/write", async (req, res) => {
  try {
    const { model, id, ids, values } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });

    const idList = (Array.isArray(ids) ? ids : [id]).map((x) => Number(x)).filter(Boolean);
    if (!idList.length) return res.status(400).json({ ok: false, error: "missing_ids" });

    const clean = cleanValuesForModel(m, values);
    if (!clean) return res.status(400).json({ ok: false, error: "missing_values" });

    const odoo = await getOdooCached(req);
    // write accepts a list of ids
    const ok = await safeWriteCompat(odoo, m, idList, clean);
    return res.json({ ok: true, result: ok });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

const ALLOWED_CALL_METHODS = new Set(["search_read", "read", "create", "write", "name_get", "fields_get"]);

app.post("/api/odoo/call", async (req, res) => {
  try {
    const { model, method, args, kwargs } = req.body || {};
    const m = String(model || "").trim();
    const meth = String(method || "").trim();

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!ALLOWED_CALL_METHODS.has(meth)) return res.status(400).json({ ok: false, error: "method_not_allowed" });

    let safeArgs = Array.isArray(args) ? args : [];
    const safeKw = (kwargs && typeof kwargs === "object") ? kwargs : {};

    // sanitize create/write payloads (defense-in-depth)
    if (meth === "create") {
      const clean = cleanValuesForModel(m, safeArgs[0]);
      if (!clean) return res.status(400).json({ ok: false, error: "missing_values" });
      safeArgs = [clean];
    }
    if (meth === "write") {
      const ids0 = Array.isArray(safeArgs[0]) ? safeArgs[0] : [];
      const vals0 = safeArgs[1];
      const clean = cleanValuesForModel(m, vals0);
      if (!ids0.length) return res.status(400).json({ ok: false, error: "missing_ids" });
      if (!clean) return res.status(400).json({ ok: false, error: "missing_values" });
      safeArgs = [ids0, clean];
    }

    const odoo = await getOdooCached(req);

    if (meth === "search_read") {
      const domain = Array.isArray(safeArgs?.[0]) ? safeArgs[0] : [];
      const fields = await filterReadFieldsCompat(odoo, m, Array.isArray(safeKw?.fields) ? safeKw.fields : ["id", "name"]);
      const limit = Math.min(Number(safeKw?.limit ?? 20), 80);
      const order = typeof safeKw?.order === "string" ? safeKw.order : undefined;
      const result = await safeSearchReadCompat(odoo, m, domain, fields, limit, order);
      return res.json({ ok: true, result });
    }

    if (meth === "read") {
      const idsArg = Array.isArray(safeArgs?.[0]) ? safeArgs[0] : [];
      if (!idsArg.length) return res.status(400).json({ ok: false, error: "missing_ids" });
      const fields = Array.isArray(safeKw?.fields) ? safeKw.fields : ["id", "name", "display_name"];
      const result = await safeReadCompat(odoo, m, idsArg, fields);
      return res.json({ ok: true, result });
    }

    if (meth === "fields_get") {
      const result = await odoo.call(m, "fields_get", [], safeKw);
      return res.json({ ok: true, result });
    }

    if (meth === "create") {
      const result = await safeCreateCompat(odoo, m, safeArgs[0]);
      return res.json({ ok: true, result });
    }

    if (meth === "write") {
      const result = await safeWriteCompat(odoo, m, safeArgs[0], safeArgs[1]);
      return res.json({ ok: true, result });
    }

    const result = await odoo.call(m, meth, safeArgs, safeKw);
    return res.json({ ok: true, result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_call_failed", details: String(e?.message || e) });
  }
});


app.post("/api/odoo/create", async (req, res) => {
  try {
    const { model, values } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!values || typeof values !== "object") return res.status(400).json({ ok: false, error: "missing_values" });

    const clean = cleanValuesForModel(m, values);
    if (!clean) return res.status(400).json({ ok: false, error: "missing_values" });

    // Extra validation for attachments
    if (m === "ir.attachment") {
      const rm = String(clean.res_model || "").trim();
      if (!rm || !modelAllowed(rm)) return res.status(400).json({ ok: false, error: "invalid_res_model" });
      const rid = Number(clean.res_id);
      if (!rid) return res.status(400).json({ ok: false, error: "invalid_res_id" });
      if (!clean.datas || typeof clean.datas !== "string") return res.status(400).json({ ok: false, error: "missing_datas" });
      clean.type = clean.type || "binary";
    }

    // Normalização simples de Many2many
    if (m === "project.task" && Array.isArray(clean.user_ids)) {
      if (Array.isArray(clean.user_ids[0])) {
        const cmd = clean.user_ids[0];
        if (!(cmd && cmd[0] === 6)) {
          delete clean.user_ids;
        }
      } else {
        const ids = clean.user_ids.map((x) => Number(x)).filter(Boolean);
        if (ids.length) clean.user_ids = [[6, 0, ids]];
        else delete clean.user_ids;
      }
    }

    if (!clean.name) return res.status(400).json({ ok: false, error: "missing_name" });

    const odoo = await getOdooCached(req);
    const id = await safeCreateCompat(odoo, m, clean);

    return res.json({ ok: true, id });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_create_failed", details: String(e?.message || e) });
  }
});


// ✅ Endpoint "Jira-like": cria ligação oculta email↔entidade no Odoo + guarda link local por conversationId
app.post("/api/odoo/link-email", async (req, res) => {
  try {
    const bodyIn = req.body || {};

    // Aceita variações do cliente (compat)
    const conversationId = bodyIn.conversationId;
    const model = bodyIn.model;
    const recordName = bodyIn.recordName || bodyIn.name || "";

    const recordIdRaw = bodyIn.recordId ?? bodyIn.resId ?? bodyIn.record_id ?? bodyIn.id;
    const rid = Number(recordIdRaw);

    const subject = bodyIn.subject ?? bodyIn.emailSubject;
    const fromEmail = bodyIn.fromEmail ?? bodyIn.emailFrom;
    const fromName = bodyIn.fromName ?? bodyIn.emailFromName;
    const receivedAtIso = bodyIn.receivedAtIso ?? bodyIn.emailReceivedAtIso;
    const emailWebLink = bodyIn.emailWebLink ?? bodyIn.url;
    const internetMessageId = bodyIn.internetMessageId ?? bodyIn.internet_message_id;
    const itemId = bodyIn.itemId ?? bodyIn.item_id;

    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!conversationId) return res.status(400).json({ ok: false, error: "missing_conversation_id" });
    if (!rid) return res.status(400).json({ ok: false, error: "missing_record_id" });

    const odoo = await getOdooCached(req);

    const safeSubject = subject || "(sem assunto)";
    const safeFrom = `${(fromName || "").trim()}${fromEmail ? ` <${fromEmail}>` : ""}`.trim() || "(desconhecido)";
    const shouldPostToChatter = bodyIn.postToChatter !== false;
    const attachmentIds = Array.isArray(bodyIn.attachmentIds)
      ? bodyIn.attachmentIds.map((value) => Number(value)).filter((value) => Number.isFinite(value) && value > 0)
      : [];

    const normalizedEmailBody = normalizeEmailBodyForOdoo(bodyIn.bodyHtml, bodyIn.bodyText);
    const bodyForOdoo = [
      `<div style="font-family: sans-serif; line-height: 1.5;">`,
      `<div style="border-left: 3px solid #714B67; padding-left: 12px; margin-bottom: 16px; color: #666;">`,
      `<p style="margin: 0 0 4px 0;"><b>Assunto:</b> ${escapeHtml(safeSubject)}</p>`,
      `<p style="margin: 0 0 4px 0;"><b>De:</b> ${escapeHtml(safeFrom)}</p>`,
      receivedAtIso ? `<p style="margin: 0 0 4px 0;"><b>Data:</b> ${escapeHtml(receivedAtIso)}</p>` : "",
      internetMessageId ? `<p style="margin: 0 0 4px 0;"><b>ID:</b> <code>${escapeHtml(internetMessageId)}</code></p>` : "",
      emailWebLink ? `<p style="margin: 0 0 4px 0;"><a href="${escapeHtml(emailWebLink)}" target="_blank" rel="noreferrer" style="color: #0078d4; text-decoration: none;">Ver no Outlook</a></p>` : "",
      `</div>`,
      normalizedEmailBody.html
        ? (normalizedEmailBody.isRichHtml
          ? `<div style="margin: 0; color: #333; overflow-x: auto;">${normalizedEmailBody.html}</div>`
          : `<blockquote style="margin: 0; padding: 0 0 0 12px; border-left: 1px solid #ccc; color: #333;">${normalizedEmailBody.html}</blockquote>`)
        : "",
      `</div>`
    ].filter(Boolean).join("\n");

    if (shouldPostToChatter) {
      // message_post no chatter do registo
      await odoo.messagePost(m, rid, bodyForOdoo, safeSubject, {
        message_id: internetMessageId || false,
        ...(attachmentIds.length ? { attachment_ids: attachmentIds } : {}),
      });
    }

    const entry = {
      model: m,
      recordId: rid,
      recordName: recordName || "",
      linkedAt: new Date().toISOString(),
      internetMessageId: internetMessageId || "",
      itemId: itemId || "",
      emailWebLink: emailWebLink || "",
      receivedAtIso: receivedAtIso || "",
      subject: safeSubject,
      fromEmail: fromEmail || "",
      fromName: fromName || "",
    };

    const list = await addLink(conversationId, entry);

    return res.json({ ok: true, links: list });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

// ✅ Alias compatível com o UI (evita "Cannot POST /api/links/link")
app.post("/api/links/link", (req, res) => {
  // reusa o handler principal
  req.url = "/api/odoo/link-email";
  app._router.handle(req, res);
});

app.get("/api/links", async (req, res) => {
  try {
    const conversationId = String(req.query.conversationId || "").trim();
    const internetMessageId = String(req.query.internetMessageId || "").trim();
    const itemId = String(req.query.itemId || "").trim();
    if (!conversationId && !internetMessageId && !itemId) return res.json({ links: [] });
    const links = await listLinksByConversation(conversationId, internetMessageId, itemId);
    return res.json({ links });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/by-record", async (req, res) => {
  try {
    const model = String(req.query.model || "").trim();
    const recordId = Number(req.query.recordId || 0);
    if (!modelAllowed(model)) return res.status(400).json({ ok: false, error: "model_not_allowed" });
    if (!recordId) return res.status(400).json({ ok: false, error: "missing_record_id" });
    const links = await listLinksByRecord(model, recordId);
    return res.json({ links });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "odoo_endpoint_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/email", async (req, res) => {
  try {
    const email = await registerRelevantEmail(req.body || {});
    return res.json({ ok: true, email });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "email_registration_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/related", async (req, res) => {
  try {
    const related = await getRelatedEmails({
      conversationId: req.query.conversationId,
      internetMessageId: req.query.internetMessageId,
      itemId: req.query.itemId,
      subject: req.query.subject,
      fromEmail: req.query.fromEmail,
      receivedAtIso: req.query.receivedAtIso,
    });
    return res.json({ ok: true, ...related });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "related_lookup_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups", async (req, res) => {
  try {
    const groups = await listCustomGroups(String(req.query.q || ""));
    return res.json({ ok: true, groups });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_lookup_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/emails", async (req, res) => {
  try {
    const emails = await listKnownEmails(String(req.query.q || ""), {
      excludeGroupId: req.query.excludeGroupId,
      limit: req.query.limit,
    });
    return res.json({ ok: true, emails });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "known_emails_lookup_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups", async (req, res) => {
  try {
    const group = await createCustomGroup(req.body || {});
    return res.json({ ok: true, group });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_create_failed", details: String(e?.message || e) });
  }
});

app.patch("/api/links/groups/:groupId", async (req, res) => {
  try {
    const group = await updateCustomGroup(req.params.groupId, req.body || {});
    return res.json({ ok: true, group });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_update_failed", details: String(e?.message || e) });
  }
});

app.delete("/api/links/groups/:groupId", async (req, res) => {
  try {
    const result = await deleteCustomGroup(req.params.groupId);
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_delete_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups/:groupId/emails", async (req, res) => {
  try {
    const emails = await listEmailsByGroup(req.params.groupId);
    return res.json({ ok: true, emails });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_emails_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/:groupId/emails", async (req, res) => {
  try {
    const result = await addEmailToGroup(req.params.groupId, req.body || {});
    return res.json({ ok: true, ...result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_link_failed", details: String(e?.message || e) });
  }
});

app.delete("/api/links/groups/:groupId/emails", async (req, res) => {
  try {
    const result = await removeEmailFromGroup(req.params.groupId, req.body || {});
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_unlink_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups/:groupId/documents", async (req, res) => {
  try {
    const documents = await listDocumentsByGroup(req.params.groupId);
    return res.json({ ok: true, documents });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_documents_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups/:groupId/documents/:documentId/content", async (req, res) => {
  try {
    const groupId = String(req.params.groupId || "").trim();
    const documentId = String(req.params.documentId || "").trim();
    const download = String(req.query.download || "").trim() === "1";
    const payload = await getDocumentContentFromGroup(groupId, documentId);
    const document = payload?.document || null;
    const buffer = payload?.buffer || null;

    if (!document || !buffer?.length) {
      return res.status(404).json({ ok: false, error: "group_document_not_found" });
    }

    const contentType = normalizeStoredDocumentMimeType(document.contentType, document.name);
    const fileName = String(document.name || "documento").replace(/["\r\n]/g, "_");

    res.setHeader("Content-Type", contentType);
    res.setHeader("Content-Length", String(buffer.length));
    res.setHeader("X-File-Name", fileName);
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader(
      "Content-Disposition",
      `${download ? "attachment" : "inline"}; filename="${fileName}"`
    );
    return res.send(buffer);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_document_content_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/emails/:emailId/attachments/:attachmentKey/content", async (req, res) => {
  try {
    const emailId = String(req.params.emailId || "").trim();
    const attachmentKey = String(req.params.attachmentKey || "").trim();
    const download = String(req.query.download || "").trim() === "1";
    const payload = await getEmailAttachmentContent(emailId, attachmentKey);
    const attachment = payload?.attachment || null;
    const buffer = payload?.buffer || null;

    if (!attachment || !buffer?.length) {
      return res.status(404).json({ ok: false, error: "email_attachment_not_found" });
    }

    const contentType = normalizeStoredDocumentMimeType(attachment.contentType, attachment.name);
    const fileName = String(attachment.name || "anexo").replace(/["\r\n]/g, "_");

    res.setHeader("Content-Type", contentType);
    res.setHeader("Content-Length", String(buffer.length));
    res.setHeader("X-File-Name", fileName);
    res.setHeader("X-Content-Type-Options", "nosniff");
    res.setHeader(
      "Content-Disposition",
      `${download ? "attachment" : "inline"}; filename="${fileName}"`
    );
    return res.send(buffer);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "email_attachment_content_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups/:groupId/attachment-flags", async (req, res) => {
  try {
    const flags = await listAttachmentFlagsByGroup(req.params.groupId);
    return res.json({ ok: true, flags });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_attachment_flags_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/:groupId/attachment-flags", async (req, res) => {
  try {
    const result = await saveAttachmentFlagsToGroup(req.params.groupId, req.body || {});
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_attachment_flags_save_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/:groupId/documents", async (req, res) => {
  try {
    const result = await saveDocumentsToGroup(req.params.groupId, req.body || {});
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_documents_save_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/groups/worksets/:worksetKey", async (req, res) => {
  try {
    const manifest = await getGroupWorksetManifest(req.params.worksetKey, {
      location: {
        mode: req.query?.mode,
        baseFolderPath: req.query?.basePath,
        chosenFolderKind: req.query?.chosenFolderKind,
        primaryTarget: req.query?.primaryTarget,
      },
    });
    if (!manifest) {
      return res.json({ ok: true, exists: false, manifest: null });
    }
    return res.json({ ok: true, exists: true, manifest });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_workset_load_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/worksets", async (req, res) => {
  try {
    const manifest = await saveGroupWorksetManifest(req.body?.manifest || req.body || {});
    return res.json({ ok: true, manifest });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_workset_save_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/worksets/migrate", async (req, res) => {
  try {
    const result = await migrateGroupWorksetManifest(req.body || {});
    return res.json({ ok: true, ...result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_workset_migrate_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/groups/storage/validate", async (req, res) => {
  try {
    const result = validateGroupStorageLocation(req.body || {});
    return res.json({ ok: true, result });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_storage_validate_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/attachments/extract-text", async (req, res) => {
  try {
    const files = Array.isArray(req.body?.files) ? req.body.files.slice(0, 8) : [];
    const results = [];
    for (const file of files) {
      const key = String(file?.key || file?.id || file?.name || "").trim();
      if (!key) continue;
      const text = await extractTextFromAttachmentPayload(file);
      results.push({
        key,
        name: String(file?.name || "").trim() || key,
        contentType: String(file?.contentType || file?.type || "").trim() || undefined,
        text,
      });
    }
    return res.json({ ok: true, results });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "attachment_text_extract_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/eml/extract", async (req, res) => {
  try {
    const emlBase64 = stripInlineBase64Prefix(req.body?.emlBase64 || req.body?.content || "");
    if (!emlBase64) {
      return res.status(400).json({ ok: false, error: "eml_missing" });
    }

    const parsed = await simpleParser(Buffer.from(emlBase64, "base64"));
    const attachments = Array.isArray(parsed?.attachments)
      ? parsed.attachments
        .map((attachment) => ({
          id: String(attachment?.cid || attachment?.checksum || attachment?.filename || "").trim() || undefined,
          name: String(attachment?.filename || "").trim(),
          contentType: String(attachment?.contentType || "application/octet-stream").trim(),
          size: Number(attachment?.size || attachment?.content?.length || 0) || undefined,
          isInline: Boolean(attachment?.contentDisposition === "inline"),
          contentId: String(attachment?.cid || "").trim() || undefined,
          content: Buffer.isBuffer(attachment?.content) ? attachment.content.toString("base64") : "",
        }))
        .filter((attachment) => attachment.name && attachment.content)
      : [];

    return res.json({
      ok: true,
      subject: String(parsed?.subject || "").trim(),
      bodyText: String(parsed?.text || "").trim(),
      bodyHtml: String(parsed?.html || "").trim(),
      attachments,
    });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "eml_extract_failed", details: String(e?.message || e) });
  }
});

app.delete("/api/links/groups/:groupId/documents/:documentId", async (req, res) => {
  try {
    const result = await deleteDocumentFromGroup(req.params.groupId, req.params.documentId);
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_document_delete_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/group-ticket-series", async (_req, res) => {
  try {
    const series = await listGroupTicketSeries();
    return res.json({ ok: true, series });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_series_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/group-ticket-series", async (req, res) => {
  try {
    const series = await createGroupTicketSeries(req.body || {});
    return res.json({ ok: true, series });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_series_create_failed", details: String(e?.message || e) });
  }
});

app.patch("/api/links/group-ticket-series/:seriesId", async (req, res) => {
  try {
    const series = await updateGroupTicketSeries(req.params.seriesId, req.body || {});
    return res.json({ ok: true, series });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_series_update_failed", details: String(e?.message || e) });
  }
});

app.delete("/api/links/group-ticket-series/:seriesId", async (req, res) => {
  try {
    const result = await deleteGroupTicketSeries(req.params.seriesId);
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_series_delete_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/group-tickets/search", async (req, res) => {
  try {
    const tickets = await listGroupTickets(String(req.body?.q || ""), {
      groupId: req.body?.groupId,
      email: req.body?.email,
      limit: req.body?.limit,
    });
    return res.json({ ok: true, tickets });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_tickets_search_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/group-tickets", async (req, res) => {
  try {
    const ticket = await createGroupTicket(req.body || {});
    return res.json({ ok: true, ticket });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_create_failed", details: String(e?.message || e) });
  }
});

app.patch("/api/links/group-tickets/:ticketId", async (req, res) => {
  try {
    const ticket = await updateGroupTicket(req.params.ticketId, req.body || {});
    return res.json({ ok: true, ticket });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_update_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/group-tickets/:ticketId/email", async (req, res) => {
  try {
    const result = await linkEmailToGroupTicket(req.params.ticketId, req.body || {});
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_email_link_failed", details: String(e?.message || e) });
  }
});

app.delete("/api/links/group-tickets/:ticketId/email", async (req, res) => {
  try {
    const result = await unlinkEmailFromGroupTicket(req.params.ticketId, req.body || {});
    return res.json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_email_unlink_failed", details: String(e?.message || e) });
  }
});

app.post("/api/links/group-tickets/detect", async (req, res) => {
  try {
    const matches = await detectGroupTicketsForEmail(req.body?.email || {});
    return res.json({ ok: true, matches });
  } catch (e) {
    console.error(e);
    return res.status(500).json({ ok: false, error: "group_ticket_detect_failed", details: String(e?.message || e) });
  }
});

app.get("/api/links/:conversationId", async (req, res) => {
  const conversationId = req.params.conversationId;
  const links = await listLinksByConversation(conversationId);
  return res.json({ links });
});

function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function simpleDecodeHtml(s) {
  if (!s || !s.includes("&")) return s;
  return s
    .replaceAll("&nbsp;", " ")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'")
    .replaceAll("&#039;", "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code) || 0));
}

function htmlToReadableText(html) {
  if (!html) return "";
  let s = simpleDecodeHtml(String(html || ""));
  s = s.replace(/<!--[\s\S]*?-->/g, " ");
  s = s.replace(/<(script|style|head|meta|link|title|xml|o:p|svg|img)[\s\S]*?<\/\1>/gi, " ");
  s = s.replace(/<(br|hr)\s*\/?>/gi, "\n");
  s = s.replace(/<\/\s*(p|div|section|article|header|footer|blockquote|tr|table|h[1-6])\s*>/gi, "\n\n");
  s = s.replace(/<\/\s*li\s*>/gi, "\n");
  s = s.replace(/<li[^>]*>/gi, "* ");
  s = s.replace(/<[^>]+>/g, " ");
  s = simpleDecodeHtml(s);
  s = s.replace(/\r/g, "");
  s = s.replace(/\t/g, " ");
  s = s.replace(/\u00a0/g, " ");
  s = s.replace(/[ ]{2,}/g, " ");
  s = s.replace(/\n{3,}/g, "\n\n");
  return s.trim();
}

function plainTextToHtml(text) {
  const normalized = String(text || "")
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[\t ]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (!normalized) return "";

  return normalized
    .split(/\n{2,}/)
    .map((block) => `<p style="margin: 0 0 10px 0;">${escapeHtml(block).replace(/\n/g, "<br/>")}</p>`)
    .join("\n");
}

function normalizeStoredDocumentMimeType(value, name) {
  const raw = String(value || "").trim().toLowerCase();
  const fileName = String(name || "").trim().toLowerCase();
  if (raw === "application/x-pdf" || (!raw && /\.pdf$/.test(fileName))) return "application/pdf";
  if (raw === "image/jpg") return "image/jpeg";
  if (raw) return raw;
  if (/\.docx$/.test(fileName)) return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  if (/\.doc$/.test(fileName)) return "application/msword";
  if (/\.xlsx$/.test(fileName)) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (/\.xls$/.test(fileName)) return "application/vnd.ms-excel";
  if (/\.pptx$/.test(fileName)) return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
  if (/\.ppt$/.test(fileName)) return "application/vnd.ms-powerpoint";
  if (/\.txt$/.test(fileName)) return "text/plain; charset=utf-8";
  return "application/octet-stream";
}

function sanitizeEmailHtmlForOdoo(html) {
  let cleaned = String(html || "").trim();
  if (!cleaned) return "";

  const bodyMatch = cleaned.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (bodyMatch?.[1]) cleaned = bodyMatch[1];

  cleaned = cleaned.replace(/<!--[\s\S]*?-->/g, "");
  cleaned = cleaned.replace(/<(script|style|meta|link|title|head|xml|svg|canvas|noscript|iframe)[^>]*>[\s\S]*?<\/\1>/gi, "");
  cleaned = cleaned.replace(/<img\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/<picture\b[^>]*>[\s\S]*?<\/picture>/gi, "");
  cleaned = cleaned.replace(/<source\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/<\/?(o:p|v:[^>\s]+|w:[^>\s]+)\b[^>]*>/gi, "");

  // Remove common signature wrappers while keeping the rest of the email HTML intact.
  cleaned = cleaned.replace(/<(div|table|section)[^>]*(?:class|id)=["'][^"']*(?:gmail_signature|signature|x_signature|moz-signature|apple-mail-signature)[^"']*["'][^>]*>[\s\S]*?<\/\1>/gi, "");

  // Strip risky handlers and javascript: links, but keep inline table/text formatting.
  cleaned = cleaned.replace(/\s+on[a-z]+\s*=\s*(['"]).*?\1/gi, "");
  cleaned = cleaned.replace(/\s+on[a-z]+\s*=\s*[^\s>]+/gi, "");
  cleaned = cleaned.replace(/\s(href|src)\s*=\s*(['"])\s*javascript:[\s\S]*?\2/gi, "");

  cleaned = cleaned.replace(/<(\/?)(html|body)\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/(?:<div>\s*<\/div>|<p>\s*<\/p>|<span>\s*<\/span>)+/gi, "");
  cleaned = cleaned.trim();

  return cleaned;
}

function normalizeEmailBodyForOdoo(bodyHtml, bodyText) {
  const richHtml = sanitizeEmailHtmlForOdoo(bodyHtml);
  if (richHtml) {
    return { html: richHtml, isRichHtml: true };
  }

  const fromHtml = htmlToReadableText(bodyHtml);
  const fromText = String(bodyText || "").trim();
  return {
    html: plainTextToHtml(fromHtml || fromText),
    isRichHtml: false,
  };
}

const host = process.env.HOST || "0.0.0.0"; // force IPv4 bind

// --- static files (UI) ---
const distPath = path.join(__dirname, "../../client/dist");
console.log(`[server] initial distPath: ${distPath}`);

if (fs.existsSync(distPath)) {
  console.log(`[server] OK distPath exists: ${distPath}`);
  const indexPath = path.join(distPath, "index.html");
  if (fs.existsSync(indexPath)) {
    console.log(`[server] OK index.html found at: ${indexPath}`);
  } else {
    console.error(`[server] ERROR index.html NOT found at: ${indexPath}`);
  }
} else {
  console.error(`[server] ERROR distPath does NOT exist: ${distPath}`);
}

app.use(express.static(distPath));

// Fallback: serve index.html for any other route (SPA) - EXCEPT /api
app.get("*", (req, res, next) => {
  if (req.url.startsWith("/api") || req.url === "/health") {
    return next();
  }
  const indexPath = path.join(distPath, "index.html");
  res.sendFile(indexPath, (err) => {
    if (err) {
      console.error(`[server] ERROR res.sendFile error: ${err.message}`);
      if (!res.headersSent) {
        res.status(404).send("SPA index.html not found");
      }
    }
  });
});

app.listen(port, host, () => {
  console.log(`[server] listening on http://${host}:${port}`);
});
