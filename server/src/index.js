import dotenv from "dotenv";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config({ path: path.resolve(__dirname, "../.env") });

import express from "express";
import cors from "cors";
import { odooClientFromEnv } from "./odoo.js";
import { addLink, listLinksByConversation } from "./linkStore.js";
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
app.use(express.json({ limit: "20mb" }));

// AI (email assistant)
app.use("/api/ai", createAiRouter());
app.use("/api/learning", createLearningRouter());

const port = process.env.PORT ? Number(process.env.PORT) : 7071;

// --- Odoo Client Cache (Singleton) ---
let cachedOdoo = null;
let odooInitPromise = null;

async function getOdooCached(req) {
  // 1. Try session token from Authorization header
  const authHeader = req.headers.authorization;
  if (authHeader && authHeader.startsWith("Session ")) {
    const token = authHeader.split(" ")[1];
    const session = sessionManager.getSession(token);
    if (session) return session.client || session;
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
      odooInitPromise = null; // allow retry
      throw e;
    } finally {
      odooInitPromise = null;
    }
  })();

  return await odooInitPromise;
}

app.get("/health", (_req, res) => res.json({ ok: true }));

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
    const sessionClient = sessionManager.getSession(token);
    if (sessionClient) return res.json({ ok: true, meta: sessionClient.meta });
  }
  return res.json({ ok: false });
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
  "project.task",
  "res.users",
  "project.task.type",
  "ir.attachment",
]);

function modelAllowed(model) {
  return MODEL_WHITELIST.has(String(model || "").trim());
}

app.get("/api/odoo/search", async (req, res) => {
  try {
    const model = String(req.query.model || "").trim();
    const q = String(req.query.q || "").trim();
    const limit = Math.min(Number(req.query.limit || 10), 20);

    if (!modelAllowed(model)) return res.status(400).send("Model not allowed");

    const odoo = await getOdooCached(req);

    // Quando a pesquisa está vazia: devolve as primeiras N linhas (útil para dropdown aberto)
    const isEmpty = !q;

    let domain;
    let fields;

    if (model === "res.partner") {
      domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email", "ilike", q]];
      fields = ["name", "email", "phone", "mobile", "display_name"];
    } else if (model === "crm.lead") {
      domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email_from", "ilike", q]];
      fields = ["name", "display_name", "email_from", "partner_id"];
    } else if (model === "project.project") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "partner_id", "user_id"];
    } else if (model === "project.task") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "project_id", "parent_id"];
    } else if (model === "res.users") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name", "email"];
    } else if (model === "project.task.type") {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name"];
    } else {
      domain = isEmpty ? [] : [["name", "ilike", q]];
      fields = ["name", "display_name"];
    }

    const items = await odoo.searchRead(model, domain, fields, limit);
    return res.json({ items: items || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

// --- compat endpoints (client expects POST + search-domain/read/write/call) ---
function cleanValuesForModel(model, values) {
  const allowedByModel = {
    "res.partner": new Set(["name", "email", "phone", "mobile"]),
    "crm.lead": new Set(["name", "email_from", "partner_id"]),
    "project.project": new Set(["name", "partner_id", "user_id"]),
    "project.task": new Set(["name", "description", "date_deadline", "project_id", "lead_id", "parent_id", "user_ids", "stage_id"]),
    "ir.attachment": new Set(["name", "datas", "res_model", "res_id", "type", "mimetype", "datas_fname"]),
  }[model];

  if (!allowedByModel) return null;
  if (!values || typeof values !== "object") return null;

  const clean = {};
  for (const [k, v] of Object.entries(values)) {
    if (allowedByModel.has(k)) clean[k] = v;
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
    fields = ["name", "email", "phone", "mobile", "display_name"];
  } else if (model === "crm.lead") {
    domain = isEmpty ? [] : ["|", ["name", "ilike", q], ["email_from", "ilike", q]];
    fields = ["name", "display_name", "email_from", "partner_id"];
  } else if (model === "project.project") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "partner_id", "user_id"];
  } else if (model === "project.task") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "project_id", "parent_id", "stage_id"];
  } else if (model === "res.users") {
    domain = isEmpty ? [] : [["name", "ilike", q]];
    fields = ["name", "display_name", "email"];
  } else if (model === "project.task.type") {
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

    if (!modelAllowed(model)) return res.status(400).send("Model not allowed");

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
      const records = await odoo.searchRead(model, domain, fields, limit, order);
      return res.json({ records: records || [] });
    }

    const spec = buildSearchSpec(model, q);
    const records = await odoo.searchRead(model, spec.domain, spec.fields, limit);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.post("/api/odoo/search-domain", async (req, res) => {
  try {
    const { model, domain, fields, limit, order } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");
    if (!Array.isArray(domain)) return res.status(400).send("Missing domain");

    const lim = Math.min(Number(limit ?? 20), 80);
    const f = Array.isArray(fields) ? fields : ["id", "name"];
    const ord = typeof order === "string" ? order : undefined;

    const odoo = await getOdooCached(req);
    const records = await odoo.searchRead(m, domain, f, lim, ord);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.post("/api/odoo/read", async (req, res) => {
  try {
    const { model, ids, fields } = req.body || {};
    const m = String(model || "").trim();
    const idList = (Array.isArray(ids) ? ids : [ids]).map((x) => Number(x)).filter(Boolean).slice(0, 80);

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");
    if (!idList.length) return res.status(400).send("Missing ids");

    const f = Array.isArray(fields) ? fields : ["id", "name", "display_name"];

    const odoo = await getOdooCached(req);
    const records = await odoo.read(m, idList, f);
    return res.json({ records: records || [] });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.post("/api/odoo/write", async (req, res) => {
  try {
    const { model, id, ids, values } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");

    const idList = (Array.isArray(ids) ? ids : [id]).map((x) => Number(x)).filter(Boolean);
    if (!idList.length) return res.status(400).send("Missing id(s)");

    const clean = cleanValuesForModel(m, values);
    if (!clean) return res.status(400).send("Missing values");

    const odoo = await getOdooCached(req);
    // write accepts a list of ids
    const ok = await odoo.call(m, "write", [idList, clean]);
    return res.json({ ok: true, result: ok });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

const ALLOWED_CALL_METHODS = new Set(["search_read", "read", "create", "write", "name_get"]);

app.post("/api/odoo/call", async (req, res) => {
  try {
    const { model, method, args, kwargs } = req.body || {};
    const m = String(model || "").trim();
    const meth = String(method || "").trim();

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");
    if (!ALLOWED_CALL_METHODS.has(meth)) return res.status(400).send("Method not allowed");

    let safeArgs = Array.isArray(args) ? args : [];
    const safeKw = (kwargs && typeof kwargs === "object") ? kwargs : {};

    // sanitize create/write payloads (defense-in-depth)
    if (meth === "create") {
      const clean = cleanValuesForModel(m, safeArgs[0]);
      if (!clean) return res.status(400).send("Missing values");
      safeArgs = [clean];
    }
    if (meth === "write") {
      const ids0 = Array.isArray(safeArgs[0]) ? safeArgs[0] : [];
      const vals0 = safeArgs[1];
      const clean = cleanValuesForModel(m, vals0);
      if (!ids0.length) return res.status(400).send("Missing ids");
      if (!clean) return res.status(400).send("Missing values");
      safeArgs = [ids0, clean];
    }

    const odoo = await getOdooCached(req);
    const result = await odoo.call(m, meth, safeArgs, safeKw);
    return res.json({ ok: true, result });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.post("/api/odoo/create", async (req, res) => {
  try {
    const { model, values } = req.body || {};
    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");
    if (!values || typeof values !== "object") return res.status(400).send("Missing values");

    const allowedByModel = {
      "res.partner": new Set(["name", "email", "phone", "mobile"]),
      "crm.lead": new Set(["name", "email_from", "partner_id"]),
      "project.project": new Set(["name", "partner_id", "user_id"]),
      // project.task:
      // - project_id (opcional)
      // - lead_id (opcional; só funciona se o módulo criar o campo)
      // - parent_id (subtarefa)
      // - user_ids (m2m) é convertido abaixo
      "project.task": new Set(["name", "description", "date_deadline", "project_id", "lead_id", "parent_id", "user_ids", "stage_id"]),
    "ir.attachment": new Set(["name", "datas", "res_model", "res_id", "type", "mimetype", "datas_fname"]),
    }[m];

    if (!allowedByModel) return res.status(400).send("Model not allowed");

    const clean = {};
    for (const [k, v] of Object.entries(values)) {
      if (allowedByModel.has(k)) clean[k] = v;
    }

        // Extra validation for attachments
    if (m === "ir.attachment") {
      const rm = String(clean.res_model || "").trim();
      if (!rm || !modelAllowed(rm)) return res.status(400).send("Invalid res_model");
      const rid = Number(clean.res_id);
      if (!rid) return res.status(400).send("Invalid res_id");
      if (!clean.datas || typeof clean.datas !== "string") return res.status(400).send("Missing datas");
      // default
      clean.type = clean.type || "binary";
    }

// Normalização simples de Many2many (aceita [ids] ou command [[6,0,[ids]]])
    if (m === "project.task" && Array.isArray(clean.user_ids)) {
      // already in command form?
      if (Array.isArray(clean.user_ids[0])) {
        const cmd = clean.user_ids[0];
        if (cmd && cmd[0] === 6) {
          // keep as is
        } else {
          // unknown command -> drop for safety
          delete clean.user_ids;
        }
      } else {
        const ids = clean.user_ids.map((x) => Number(x)).filter(Boolean);
        if (ids.length) clean.user_ids = [[6, 0, ids]];
        else delete clean.user_ids;
      }
    }

    if (!clean.name) return res.status(400).send("Missing name");

    const odoo = await getOdooCached(req);
    const id = await odoo.create(m, clean);

    return res.json({ ok: true, id });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
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

    const m = String(model || "").trim();

    if (!modelAllowed(m)) return res.status(400).send("Model not allowed");
    if (!conversationId) return res.status(400).send("Missing conversationId");
    if (!rid) return res.status(400).send("Missing recordId");

    const odoo = await getOdooCached(req);

    const safeSubject = subject || "(sem assunto)";
    const safeFrom = `${(fromName || "").trim()}${fromEmail ? ` <${fromEmail}>` : ""}`.trim() || "(desconhecido)";

    // HTML limpo e legível dentro do chatter do Odoo
    const body = [
      `<p><b>Ligação criada a partir do Outlook</b></p>`,
      `<p><b>Assunto:</b> ${escapeHtml(safeSubject)}</p>`,
      `<p><b>De:</b> ${escapeHtml(safeFrom)}</p>`,
      receivedAtIso ? `<p><b>Data:</b> ${escapeHtml(receivedAtIso)}</p>` : "",
      internetMessageId ? `<p><b>InternetMessageId:</b> <code>${escapeHtml(internetMessageId)}</code></p>` : "",
      `<p style="color:#666;"><small><b>Thread/ConversationId:</b> ${escapeHtml(conversationId)}</small></p>`,
      emailWebLink ? `<p><b>Outlook link:</b> <a href="${escapeHtml(emailWebLink)}" target="_blank" rel="noreferrer">Abrir email</a></p>` : "",
      `<p style="color:#888;"><small>(Anexos: MVP ainda não envia. Próxima fase.)</small></p>`,
    ].filter(Boolean).join("\n");

    // message_post no chatter do registo
    await odoo.messagePost(m, rid, body, safeSubject);

    const entry = {
      model: m,
      recordId: rid,
      recordName: recordName || "",
      linkedAt: new Date().toISOString(),
      internetMessageId: internetMessageId || "",
      subject: safeSubject,
      fromEmail: fromEmail || "",
      fromName: fromName || "",
    };

    const list = await addLink(conversationId, entry);

    return res.json({ ok: true, links: list });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
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
    if (!conversationId) return res.json({ links: [] });
    const links = await listLinksByConversation(conversationId);
    return res.json({ links });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
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

const host = process.env.HOST || "0.0.0.0"; // force IPv4 bind

// --- static files (UI) ---
const distPath = path.join(__dirname, "../../client/dist");
console.log(`[server] initial distPath: ${distPath}`);

if (fs.existsSync(distPath)) {
  console.log(`[server] ✅ distPath exists: ${distPath}`);
  const indexPath = path.join(distPath, "index.html");
  if (fs.existsSync(indexPath)) {
    console.log(`[server] ✅ index.html found at: ${indexPath}`);
  } else {
    console.error(`[server] ❌ index.html NOT found at: ${indexPath}`);
  }
} else {
  console.error(`[server] ❌ distPath does NOT exist: ${distPath}`);
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
      console.error(`[server] ❌ res.sendFile error: ${err.message}`);
      if (!res.headersSent) {
        res.status(404).send("SPA index.html not found");
      }
    }
  });
});

app.listen(port, host, () => {
  console.log(`[server] listening on http://${host}:${port}`);
});
