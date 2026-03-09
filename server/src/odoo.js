import axios from "axios";
import { CookieJar } from "tough-cookie";
import { wrapper } from "axios-cookiejar-support";
import https from "node:https";
import { createOdooSchemaCache } from "./odoo_schema_cache.js";

/**
 * Odoo JSON-RPC client with session cookies.
 * Auth: /web/session/authenticate
 * Calls: /web/dataset/call_kw
 *
 * Optional troubleshooting:
 *   ODOO_INSECURE_TLS=true
 */

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

function getEnvValue(...values) {
  for (const value of values) {
    const normalized = String(value || "").trim();
    if (normalized) return normalized;
  }
  return "";
}

export function getOdooRuntimeConfig(config = null) {
  const baseUrl = sanitizeBaseUrl(getEnvValue(config?.url, process.env.ODOO_URL));
  const db = getEnvValue(config?.db, process.env.ODOO_DB);
  const login = getEnvValue(config?.login, process.env.ODOO_USERNAME, process.env.ODOO_USER);
  const password = getEnvValue(config?.password, process.env.ODOO_API_KEY, process.env.ODOO_PASSWORD, process.env.ODOO_PASS);
  return { baseUrl, db, login, password };
}

export function getMissingOdooConfigKeys(config = null) {
  const runtime = getOdooRuntimeConfig(config);
  const missing = [];
  if (!runtime.baseUrl || runtime.baseUrl.includes("your-odoo-instance.com")) missing.push("ODOO_URL");
  if (!runtime.db || runtime.db === "your_db_name") missing.push("ODOO_DB");
  if (!runtime.login || runtime.login === "your_username") missing.push("ODOO_USERNAME");
  if (!runtime.password || runtime.password === "your_password") missing.push("ODOO_API_KEY");
  return missing;
}

export function hasOdooRuntimeConfig(config = null) {
  return getMissingOdooConfigKeys(config).length === 0;
}

export async function odooClientFromEnv(config = null) {
  const missing = getMissingOdooConfigKeys(config);
  if (missing.length) {
    throw Object.assign(
      new Error(`Odoo configuration incomplete: missing ${missing.join(", ")}`),
      { status: 503, code: "ODOO_CONFIG_MISSING", missing }
    );
  }

  const { baseUrl, db, login, password } = getOdooRuntimeConfig(config);
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

  const schemaCache = createOdooSchemaCache({
    fetchFields: async (model) => {
      return await callKw({
        model,
        method: "fields_get",
        args: [],
        kwargs: { attributes: ["type", "readonly", "required"] },
      });
    },
  });

  async function rawSearchRead(model, domain, fields, limit = 10, order) {
    const kwargs = { fields, limit };
    if (order) kwargs.order = order;
    return await callKw({ model, method: "search_read", args: [domain], kwargs });
  }

  async function rawCreate(model, vals) {
    return await callKw({ model, method: "create", args: [vals] });
  }

  async function rawWrite(model, ids, vals) {
    const idList = (Array.isArray(ids) ? ids : [ids]).map((x) => Number(x)).filter(Boolean);
    return await callKw({ model, method: "write", args: [idList, vals] });
  }

  return {
    meta: {
      baseUrl,
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
      return await rawSearchRead(model, domain, fields, limit, order);
    },

    async safeSearchRead(model, domain, wantedFields, limit = 10, order) {
      return await schemaCache.safeSearchRead(
        model,
        domain,
        wantedFields,
        limit,
        { order },
        async (m, d, filteredFields, lim, opts) => rawSearchRead(m, d, filteredFields, lim, opts?.order)
      );
    },

    async create(model, vals) {
      return await rawCreate(model, vals);
    },

    async safeCreate(model, vals) {
      return await schemaCache.safeCreate(model, vals, async (m, clean) => rawCreate(m, clean));
    },

    async write(model, id, vals) {
      return await rawWrite(model, [Number(id)], vals);
    },

    async safeWrite(model, ids, vals) {
      return await schemaCache.safeWrite(model, ids, vals, async (m, idList, clean) => rawWrite(m, idList, clean));
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

    async messagePost(model, id, body, subject, extraKwargs = {}) {
      return await callKw({
        model,
        method: "message_post",
        args: [[Number(id)]],
        kwargs: {
          body,
          body_is_html: true,
          subject: subject || "",
          message_type: "comment",
          subtype_xmlid: "mail.mt_comment",
          ...extraKwargs,
        },
      });
    },

    async findPartnerByEmail(email) {
      const result = await this.safeSearchRead(
        "res.partner",
        [["email", "=", email]],
        ["id", "name", "email", "phone", "mobile", "function", "company_type", "is_company", "parent_id", "vat", "street", "zip", "city", "country_id"],
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

    schema: {
      getModelFields: (model, opts) => schemaCache.getModelFields(model, opts),
      invalidateModel: (model) => schemaCache.invalidateModel(model),
      invalidateAll: () => schemaCache.invalidateAll(),
      filterReadFields: (model, wantedFields) => schemaCache.filterReadFields(model, wantedFields),
      sanitizeWriteData: (model, data) => schemaCache.sanitizeWriteData(model, data),
    },

    async createLead({ name, email_from, partner_id }) {
      const vals = { name };
      if (email_from) vals.email_from = email_from;
      if (partner_id) vals.partner_id = partner_id;
      return await this.create("crm.lead", vals);
    },
  };
}
