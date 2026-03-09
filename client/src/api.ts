// Goal: keep UI stable even if server endpoints evolve.

let _sessionToken: string | null = null;

export function setApiSessionToken(token: string | null) {
  _sessionToken = token;
}

export type OdooMeta = {
  ok: boolean;
  baseUrl?: string;     // preferred
  webBaseUrl?: string;  // compat (algum código ainda usa isto)
  url?: string;         // compat (outro código ainda usa isto)
  db?: string;
  user?: { id: number; name: string; login?: string } | null;
  models?: string[];
};

export type OdooMetaResponse = { ok: boolean; meta: OdooMeta };

const isMock = () => typeof window !== "undefined" && !!(window as any).__ICCC_MOCK__;
const getMockData = (type: string, model: string) => (window as any).__ICCC_MOCK_DATA?.[type]?.[model] || null;

export type LinkEntry = {
  id?: string;
  conversationId: string;
  model: string;

  // preferred
  recordId?: number;
  recordName?: string;

  // compat aliases
  resId?: number;
  name?: string;

  // display helpers
  title?: string;
  url?: string;

  createdAt?: string;
  updatedAt?: string;
};

export type LinkPayload = {
  conversationId: string;
  model: string;

  // preferred
  recordId: number;
  recordName?: string;

  emailSubject?: string;
  emailFrom?: string;
  emailWebLink?: string;
  internetMessageId?: string;
  receivedAtIso?: string;
  bodyHtml?: string;
  bodyText?: string;

  // compat aliases
  resId?: number;
  name?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  url?: string;
};

export type AiGenerateResponse =
  | {
    ok: true;
    html?: string;
    text: string;
    data?: any;
    suggestedRecipients?: { to: string[]; cc: string[] };
    suggestedSubject?: string;
  }
  | { ok: false; error: string };

export type AuthResponse = { ok: true; token: string; meta: OdooMeta } | { ok: false; message: string };

type Json = any;

async function requestJSON<T = Json>(path: string, init?: RequestInit): Promise<T> {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), 10000); // 10s timeout

  try {
    const res = await fetch(path, {
      ...init,
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        ...(_sessionToken ? { "Authorization": `Session ${_sessionToken}` } : {}),
        ...(init?.headers || {}),
      },
    });

    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const body = ct.includes("application/json") ? await res.json() : await res.text();

    if (!res.ok) {
      const msg =
        typeof body === "string"
          ? body
          : body?.error || body?.message || JSON.stringify(body);
      throw new Error(`HTTP ${res.status}: ${msg}`);
    }
    return body as T;
  } finally {
    clearTimeout(id);
  }
}

// -------- Auth --------
export async function login(credentials: any): Promise<AuthResponse> {
  return await requestJSON(`/api/auth/login`, {
    method: "POST",
    body: JSON.stringify(credentials),
  });
}

export async function checkAuth(): Promise<{ ok: boolean; meta?: OdooMeta }> {
  return await requestJSON(`/api/auth/check`);
}

// -------- Odoo meta / ping --------
export async function getOdooMeta(): Promise<OdooMeta> {
  const r: any = await requestJSON(`/api/odoo/meta`);
  const m = (r?.meta ?? r) as OdooMeta;
  if (m) {
    // Standardize URL property
    m.baseUrl = m.baseUrl || m.webBaseUrl || m.url;
  }
  return m;
}

export function getOdooAutoLoginUrl(_token: string | null, redirect: string = "/web", odooBaseUrl?: string): string {
  let targetPath = redirect;
  if (odooBaseUrl && redirect.startsWith(odooBaseUrl)) {
    targetPath = redirect.substring(odooBaseUrl.length);
  }
  if (!targetPath.startsWith("/")) targetPath = "/" + targetPath;

  if (!odooBaseUrl) {
    return targetPath;
  }

  const base = String(odooBaseUrl).replace(/\/+$/, "");
  const redirectUrl = new URL(targetPath, base + "/");
  const db = redirectUrl.searchParams.get("db");
  const loginUrl = new URL(base + "/web/login");

  if (db) {
    loginUrl.searchParams.set("db", db);
  }
  loginUrl.searchParams.set("redirect", redirectUrl.pathname + redirectUrl.search + redirectUrl.hash);

  return loginUrl.toString();
}

export async function odooPing(): Promise<{ ok: boolean }> {
  return await requestJSON(`/api/odoo/ping`);
}

// -------- Links --------
export async function getLinks(conversationId?: string, internetMessageId?: string): Promise<LinkEntry[]> {
  const params = new URLSearchParams();
  const normalizedConversationId = String(conversationId || "").trim();
  const normalizedInternetMessageId = String(internetMessageId || "")
    .trim()
    .toLowerCase()
    .replace(/[<>\s]/g, "");
  const lookupKey = normalizedConversationId || normalizedInternetMessageId
    ? `${normalizedConversationId}||${normalizedInternetMessageId}`
    : "";
  if (lookupKey) params.set("conversationId", lookupKey);
  if (normalizedInternetMessageId) params.set("internetMessageId", normalizedInternetMessageId);
  const r: any = await requestJSON(`/api/links?${params.toString()}`);
  const links: LinkEntry[] = r?.links ?? r ?? [];
  return (Array.isArray(links) ? links : []).map((l: any) => ({
    ...l,
    resId: l.resId ?? l.recordId,
    recordId: l.recordId ?? l.resId,
    name: l.name ?? l.recordName ?? l.title,
    title: l.title ?? l.recordName ?? l.name ?? l.model,
    url: l.url ?? l.emailWebLink,
  }));
}

export async function linkEmailToRecord(payload: LinkPayload): Promise<{ ok: boolean; link?: LinkEntry }> {
  try {
    return await requestJSON(`/api/links/link`, { method: "POST", body: JSON.stringify(payload) });
  } catch {
    // fallback for older servers
    return await requestJSON(`/api/odoo/link-email`, { method: "POST", body: JSON.stringify(payload) });
  }
}


export type PartnerLite = {
  id: number;
  name?: string;
  email?: string;
  company_type?: "person" | "company" | string;
  parent_id?: [number, string] | number | null;
  function?: string;
  phone?: string;
  mobile?: string;
};

export async function getPartnerByEmail(email: string): Promise<PartnerLite | null> {
  const q = encodeURIComponent(String(email || "").trim());
  const r: any = await requestJSON(`/api/odoo/partners/by-email?email=${q}`);
  return r?.partner ?? null;
}

export async function createOrUpdatePartner(payload: {
  mode: "create" | "update";
  targetPartnerId?: number;
  data: {
    name?: string;
    email?: string;
    company_type?: "person" | "company";
    parent_id?: number | null;
    function?: string;
    phone?: string;
    mobile?: string;
  };
}): Promise<any> {
  return await requestJSON(`/api/odoo/partners/create-or-update`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export async function searchCompanies(q: string): Promise<any[]> {
  const query = encodeURIComponent(String(q || "").trim());
  const r: any = await requestJSON(`/api/odoo/companies/search?q=${query}`);
  // Merge-conflict resolution: keep compatibility with both payload shapes
  // and always return a safe array.
  const results = r?.results ?? r?.companies ?? [];
  return Array.isArray(results) ? results : [];
}

// -------- Odoo generic helpers --------
export async function readOdoo(model: string, ids: number[] | number, fields: string[]): Promise<any[]> {
  const idList = Array.isArray(ids) ? ids : [ids];
  if (isMock()) {
    console.log(`[Mock] Reading ${model}`, idList);
    const mock = getMockData("read", model);
    return idList.map(id => mock?.[id] || { id, name: `Mock ${model} ${id}` });
  }
  try {
    const r: any = await requestJSON(`/api/odoo/read`, {
      method: "POST",
      body: JSON.stringify({ model, ids: idList, fields }),
    });
    return r?.records ?? r?.result ?? r ?? [];
  } catch {
    // fallback to domain search (if read endpoint absent)
    const r2: any = await requestJSON(`/api/odoo/search-domain`, {
      method: "POST",
      body: JSON.stringify({ model, domain: [["id", "in", idList]], fields, limit: idList.length }),
    });
    return r2?.records ?? r2?.result ?? r2 ?? [];
  }
}

// searchOdoo: supports both old (model, query, limit) and new (args object)
export async function searchOdoo(
  modelOrArgs:
    | string
    | { model: string; domain: any[]; fields?: string[]; limit?: number; order?: string },
  query?: string,
  limit?: number
): Promise<any[]> {
  if (typeof modelOrArgs === "string") {
    const model = modelOrArgs;
    const q = (query ?? "").trim();
    const lim = limit ?? 20;

    // if server implements free-text search, use it; else fallback to name ilike
    try {
      const r: any = await requestJSON(`/api/odoo/search`, {
        method: "POST",
        body: JSON.stringify({ model, query: q, limit: lim }),
      });
      return r?.records ?? r?.result ?? r ?? [];
    } catch {
      const domain = q ? [["name", "ilike", q]] : [];
      const r2: any = await requestJSON(`/api/odoo/search-domain`, {
        method: "POST",
        body: JSON.stringify({ model, domain, fields: ["id", "name"], limit: lim }),
      });
      return r2?.records ?? r2?.result ?? r2 ?? [];
    }
  }

  // args object
  try {
    const r: any = await requestJSON(`/api/odoo/search`, { method: "POST", body: JSON.stringify(modelOrArgs) });
    return r?.records ?? r?.result ?? r ?? [];
  } catch {
    const r2: any = await requestJSON(`/api/odoo/search-domain`, { method: "POST", body: JSON.stringify(modelOrArgs) });
    return r2?.records ?? r2?.result ?? r2 ?? [];
  }
}

// searchOdooDomain: supports both signatures
export async function searchOdooDomain(args: { model: string; domain: any[]; fields?: string[]; limit?: number; order?: string }): Promise<any[]>;
export async function searchOdooDomain(model: string, domain: any[], fields?: string[], limit?: number): Promise<any[]>;
export async function searchOdooDomain(
  a: any,
  b?: any,
  c?: any,
  d?: any
): Promise<any[]> {
  const payload =
    typeof a === "string"
      ? { model: a, domain: b ?? [], fields: c, limit: d }
      : a;

  const r: any = isMock() ? { records: getMockData("search_domain", payload.model) || [] } : await requestJSON(`/api/odoo/search-domain`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return r?.records ?? r?.result ?? r ?? [];
}

export async function callOdoo(payload: { model: string; method: string; args: any[]; kwargs?: Record<string, any> }): Promise<any> {
  const response: any = await requestJSON(`/api/odoo/call`, { method: "POST", body: JSON.stringify(payload) });
  return response?.result ?? response;
}

export type OdooFieldMeta = {
  name: string;
  string?: string;
  type?: string;
  relation?: string;
  selection?: Array<[string, string]>;
};

function normalizeFieldLabel(value: string): string {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

export async function findOdooFieldByLabel(model: string, label: string): Promise<OdooFieldMeta | null> {
  const target = normalizeFieldLabel(label);
  if (!target) return null;

  const result: Record<string, any> = await callOdoo({
    model,
    method: "fields_get",
    args: [],
    kwargs: { attributes: ["string", "type", "relation", "selection"] },
  });

  for (const [name, meta] of Object.entries(result || {})) {
    if (normalizeFieldLabel(String(meta?.string || "")) !== target) continue;
    return {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };
  }

  return null;
}

export async function findOdooField(
  model: string,
  options: {
    labels?: string[];
    nameCandidates?: string[];
    namePatterns?: RegExp[];
    preferredTypes?: string[];
  }
): Promise<OdooFieldMeta | null> {
  const result: Record<string, any> = await callOdoo({
    model,
    method: "fields_get",
    args: [],
    kwargs: { attributes: ["string", "type", "relation", "selection"] },
  });

  const labelTargets = (options.labels || []).map(normalizeFieldLabel).filter(Boolean);
  const candidateNames = new Set((options.nameCandidates || []).map((name) => String(name || "").trim()).filter(Boolean));
  const normalizedCandidateNames = new Set(
    (options.nameCandidates || []).map((name) => normalizeFieldLabel(String(name || ""))).filter(Boolean)
  );
  const preferredTypes = new Set(options.preferredTypes || []);

  for (const [name, meta] of Object.entries(result || {})) {
    const normalizedName = normalizeFieldLabel(name);
    if (!candidateNames.has(name) && !normalizedCandidateNames.has(normalizedName)) continue;
    return {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };
  }

  let best: { score: number; field: OdooFieldMeta } | null = null;

  for (const [name, meta] of Object.entries(result || {})) {
    const normalizedLabel = normalizeFieldLabel(String(meta?.string || ""));
    const normalizedName = normalizeFieldLabel(name);
    let score = -1;

    if (candidateNames.has(name) || normalizedCandidateNames.has(normalizedName)) score = 110;
    else if (labelTargets.includes(normalizedLabel)) score = 100;
    else if (labelTargets.some((target) => normalizedLabel.includes(target) || target.includes(normalizedLabel))) score = 80;
    else if (labelTargets.some((target) => target.split(/\s+/).every((token) => normalizedLabel.includes(token)))) score = 70;
    else if ((options.namePatterns || []).some((pattern) => pattern.test(name))) score = 60;

    if (score < 0) continue;
    if (preferredTypes.has(String(meta?.type || ""))) score += 5;
    if (labelTargets.some((target) => target.split(/\s+/).every((token) => normalizedName.includes(token)))) score += 2;
    if (String(meta?.type || "") === "selection" && Array.isArray(meta?.selection) && meta.selection.length) score += 20;
    if (String(meta?.type || "") === "many2one" && meta?.relation) score += 10;

    const field: OdooFieldMeta = {
      name,
      string: meta?.string,
      type: meta?.type,
      relation: meta?.relation,
      selection: Array.isArray(meta?.selection) ? meta.selection : [],
    };

    if (!best || score > best.score) {
      best = { score, field };
    }
  }

  return best?.field || null;
}

// createOdoo: return number (DialogApp expects number)
export async function createOdoo(model: string, values: Record<string, any>): Promise<number> {
  // prefer dedicated endpoint
  if (isMock()) {
    const id = Math.floor(Math.random() * 10000);
    console.log(`[Mock] Created ${model} with ID ${id}`, values);
    return id;
  }
  try {
    const r: any = await requestJSON(`/api/odoo/create`, { method: "POST", body: JSON.stringify({ model, values }) });
    const id = r?.id ?? r?.result ?? r;
    return Number(id);
  } catch {
    const r2: any = await callOdoo({ model, method: "create", args: [values] });
    const id = r2?.id ?? r2?.result ?? r2;
    return Number(id);
  }
}

export async function writeOdoo(model: string, ids: number[] | number, values: Record<string, any>): Promise<boolean> {
  const idList = Array.isArray(ids) ? ids : [ids];
  // try direct endpoint if it exists
  try {
    const r: any = await requestJSON(`/api/odoo/write`, { method: "POST", body: JSON.stringify({ model, id: idList[0], ids: idList, values }) });
    return Boolean(r?.ok ?? r?.result ?? r ?? true);
  } catch {
    const r2: any = await callOdoo({ model, method: "write", args: [idList, values] });
    return Boolean(r2?.ok ?? r2?.result ?? r2 ?? true);
  }
}

// -------- AI --------
export async function aiSelftest(customModels?: any): Promise<{ ok: boolean; openai: { ok: boolean; error?: string }; gemini: { ok: boolean; error?: string }; error?: string }> {
  return await requestJSON(`/api/ai/selftest`, { method: "POST", body: JSON.stringify({ customModels }) });
}

export async function aiGenerate(payload: any, customModels?: any): Promise<AiGenerateResponse> {
  const mergedPayload = { ...payload, customModels: customModels ?? payload.customModels };
  return await requestJSON(`/api/ai/generate`, { method: "POST", body: JSON.stringify(mergedPayload) });
}

export async function aiExtractAnchors(emailBody: string, customModels?: any, emailContext?: any): Promise<{ ok: boolean; anchors: any }> {
  return await requestJSON(`/api/ai/extract-anchors`, { method: "POST", body: JSON.stringify({ emailBody, emailContext, customModels }) });
}

export async function aiGenerateBriefing(context: string, history: any[] = [], customModels?: any, conversationId?: string): Promise<{ ok: boolean; summary: string }> {
  return await requestJSON(`/api/ai/briefing`, { method: "POST", body: JSON.stringify({ context, history, customModels, conversationId }) });
}

export async function aiVoiceCommand(commandText: string, context: any): Promise<{ ok: boolean; actions: string[] }> {
  return await requestJSON(`/api/ai/voice-command`, { method: "POST", body: JSON.stringify({ commandText, context }) });
}

export async function aiListModels(): Promise<{ ok: boolean; openai: string[]; gemini: string[] }> {
  return await requestJSON(`/api/ai/list-models`);
}

// -------- Learning --------
export async function logLearningInteraction(log: any): Promise<{ ok: boolean }> {
  return await requestJSON(`/api/learning/log`, {
    method: "POST",
    body: JSON.stringify(log),
  });
}

export async function getLearningProfile(userId: string = "global"): Promise<{ ok: boolean; profile: any }> {
  return await requestJSON(`/api/learning/profile?userId=${encodeURIComponent(userId)}`);
}
