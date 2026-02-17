import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import { odooClientFromEnv } from "./odoo.js";
import { addLink, listLinksByConversation } from "./linkStore.js";
import { createAiRouter } from "./routes/aiRoutes.js";
import { fileURLToPath } from "url";
import path from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

dotenv.config();

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

const port = process.env.PORT ? Number(process.env.PORT) : 7071;

app.get("/health", (_req, res) => res.json({ ok: true }));

app.get("/api/odoo/meta", async (_req, res) => {
  try {
    const odoo = await odooClientFromEnv();
    return res.json({ ok: true, meta: odoo.meta });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.get("/api/odoo/ping", async (_req, res) => {
  try {
    const odoo = await odooClientFromEnv();
    const ok = await odoo.ping();
    return res.json({ ok });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
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

    const odoo = await odooClientFromEnv();

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

    const odoo = await odooClientFromEnv();

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

    const odoo = await odooClientFromEnv();
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

    const odoo = await odooClientFromEnv();
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

    const odoo = await odooClientFromEnv();
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

    const odoo = await odooClientFromEnv();
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
    }[m];

    if (!allowedByModel) return res.status(400).send("Model not allowed");

    const clean = {};
    for (const [k, v] of Object.entries(values)) {
      if (allowedByModel.has(k)) clean[k] = v;
    }

    // Normalização simples de Many2many (Dialog envia [id], Odoo quer command)
    if (m === "project.task" && Array.isArray(clean.user_ids)) {
      const ids = clean.user_ids.map((x) => Number(x)).filter(Boolean);
      if (ids.length) clean.user_ids = [[6, 0, ids]];
      else delete clean.user_ids;
    }

    if (!clean.name) return res.status(400).send("Missing name");

    const odoo = await odooClientFromEnv();
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

    const odoo = await odooClientFromEnv();

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

    const list = addLink(conversationId, entry);

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

app.get("/api/links", (req, res) => {
  try {
    const conversationId = String(req.query.conversationId || "").trim();
    if (!conversationId) return res.json({ links: [] });
    const links = listLinksByConversation(conversationId);
    return res.json({ links });
  } catch (e) {
    console.error(e);
    return res.status(500).send(String(e?.message || e));
  }
});

app.get("/api/links/:conversationId", (req, res) => {
  const conversationId = req.params.conversationId;
  const links = listLinksByConversation(conversationId);
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
app.use(express.static(distPath));

// Fallback: serve index.html for any other route (SPA) - EXCEPT /api
app.get("*", (req, res, next) => {
  if (req.url.startsWith("/api") || req.url === "/health") {
    return next();
  }
  res.sendFile(path.join(distPath, "index.html"));
});

app.listen(port, host, () => {
  console.log(`[server] listening on http://${host}:${port}`);
});
