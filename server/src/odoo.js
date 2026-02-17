import axios from "axios";
import { CookieJar } from "tough-cookie";
import { wrapper } from "axios-cookiejar-support";
import https from "node:https";

/**
 * Odoo JSON-RPC client with session cookies.
 * Auth: /web/session/authenticate
 * Calls: /web/dataset/call_kw
 *
 * Optional troubleshooting:
 *   ODOO_INSECURE_TLS=true
 */

function requireEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}

function sanitizeBaseUrl(url) {
  return String(url || "").replace(/\/+$/, "");
}

function safeJson(data) {
  try {
    return JSON.stringify(data);
  } catch {
    return String(data);
  }
}

function buildHttpsAgentIfNeeded(baseUrl) {
  if (!/^https:\/\//i.test(baseUrl)) return undefined;
  const insecure = String(process.env.ODOO_INSECURE_TLS || "").toLowerCase() === "true";
  if (!insecure) return undefined;
  return new https.Agent({ rejectUnauthorized: false });
}

export async function odooClientFromEnv() {
  const baseUrl = sanitizeBaseUrl(requireEnv("ODOO_URL")); // IMPORTANT: no /web
  const db = requireEnv("ODOO_DB");
  const login = requireEnv("ODOO_USERNAME");
  const password = requireEnv("ODOO_API_KEY"); // currently used as password

  const jar = new CookieJar();
  const httpsAgent = buildHttpsAgentIfNeeded(baseUrl);

  const http = wrapper(
    axios.create({
      baseURL: baseUrl,
      jar,
      withCredentials: true,
      timeout: 20000,
      httpsAgent,
    })
  );

  async function postJson(path, payload) {
    let resp;
    try {
      resp = await http.post(path, payload, {
        headers: { "Content-Type": "application/json" },
        maxRedirects: 0,
        validateStatus: () => true,
      });
    } catch (e) {
      const msg = e?.response ? `HTTP ${e.response.status} ${safeJson(e.response.data)}` : e?.message || String(e);
      throw new Error(`Erro de rede ao ligar ao Odoo (${path}): ${msg}`);
    }
    return resp;
  }

  const authPayload = {
    jsonrpc: "2.0",
    method: "call",
    params: { db, login, password },
    id: Date.now(),
  };

  const authResp = await postJson("/web/session/authenticate", authPayload);

  if (authResp.status !== 200) {
    throw new Error(`Odoo respondeu HTTP ${authResp.status} em authenticate. Body: ${safeJson(authResp.data)}`);
  }

  const uid = authResp?.data?.result?.uid;
  if (!uid) {
    throw new Error(`Auth falhou (uid=false). Resposta: ${safeJson(authResp.data)}`);
  }

  const webBaseUrl = authResp?.data?.result?.["web.base.url"] || baseUrl;

  async function callKw({ model, method, args = [], kwargs = {} }) {
    const payload = {
      jsonrpc: "2.0",
      method: "call",
      params: { model, method, args, kwargs },
      id: Date.now(),
    };

    const r = await postJson("/web/dataset/call_kw", payload);

    if (r.status !== 200) {
      throw new Error(`Odoo respondeu HTTP ${r.status} em call_kw. Body: ${safeJson(r.data)}`);
    }
    if (r?.data?.error) {
      throw new Error(`Odoo JSON-RPC error: ${safeJson(r.data.error)}`);
    }
    return r?.data?.result;
  }

  return {
    meta: {
      webBaseUrl,
      db,
      uid,
      login,
      serverVersion: authResp?.data?.result?.server_version,
    },

    async ping() {
      const result = await callKw({
        model: "res.partner",
        method: "search_read",
        args: [[["id", "=", 1]]],
        kwargs: { fields: ["name"], limit: 1 },
      });
      return Array.isArray(result);
    },

    async searchRead(model, domain, fields, limit = 10, order) {
      const kwargs = { fields, limit };
      if (order) kwargs.order = order;
      return await callKw({
        model,
        method: "search_read",
        args: [domain],
        kwargs,
      });
    },

    async create(model, vals) {
      return await callKw({
        model,
        method: "create",
        args: [vals],
      });
    },

    async write(model, id, vals) {
      return await callKw({
        model,
        method: "write",
        args: [[Number(id)], vals],
      });
    },

    async read(model, ids, fields) {
      const idList = (Array.isArray(ids) ? ids : [ids]).map((x) => Number(x)).filter(Boolean);
      return await callKw({
        model,
        method: "read",
        args: [idList],
        kwargs: { fields },
      });
    },

    async call(model, method, args = [], kwargs = {}) {
      return await callKw({ model, method, args, kwargs });
    },

    async messagePost(model, id, body, subject) {
      return await callKw({
        model,
        method: "message_post",
        args: [[Number(id)]],
        kwargs: {
          body,
          subject: subject || "",
          message_type: "comment",
          subtype_xmlid: "mail.mt_comment",
        },
      });
    },

    async findPartnerByEmail(email) {
      const result = await this.searchRead(
        "res.partner",
        [["email", "=", email]],
        ["name", "email", "phone", "mobile"],
        1
      );
      const p = Array.isArray(result) ? result[0] : null;
      if (!p) return null;
      return {
        id: p.id,
        name: p.name,
        email: p.email,
        phone: p.phone,
        mobile: p.mobile,
      };
    },

    async createLead({ name, email_from, partner_id }) {
      const vals = { name };
      if (email_from) vals.email_from = email_from;
      if (partner_id) vals.partner_id = partner_id;
      return await this.create("crm.lead", vals);
    },
  };
}
