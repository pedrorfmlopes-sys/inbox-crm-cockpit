// client/src/api.ts
// Inbox CRM Cockpit - client-side API helper (superset + backwards compatible)
// Goal: keep UI stable even if server endpoints evolve.

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

  // compat aliases
  resId?: number;
  name?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  url?: string;
};

export type AiGenerateResponse =
  | { ok: true; html?: string; text?: string; data?: any }
  | { ok: false; error: string };

type Json = any;

async function requestJSON<T = Json>(path: string, init?: RequestInit): Promise<T> {
  const res = await fetch(path, {
    ...init,
    headers: {
      "Content-Type": "application/json",
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
}

// -------- Odoo meta / ping --------
export async function getOdooMeta(): Promise<OdooMeta> {
  const r: any = await requestJSON(`/api/odoo/meta`);
  return (r?.meta ?? r) as OdooMeta;
}

export async function odooPing(): Promise<{ ok: boolean }> {
  return await requestJSON(`/api/odoo/ping`);
}

// -------- Links --------
export async function getLinks(conversationId: string): Promise<LinkEntry[]> {
  const q = encodeURIComponent(conversationId);
  const r: any = await requestJSON(`/api/links?conversationId=${q}`);
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
  } catch (e1: any) {
    // fallback for older servers
    return await requestJSON(`/api/odoo/link-email`, { method: "POST", body: JSON.stringify(payload) });
  }
}

// -------- Odoo generic helpers --------
export async function readOdoo(model: string, ids: number[] | number, fields: string[]): Promise<any[]> {
  const idList = Array.isArray(ids) ? ids : [ids];
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

  const r: any = await requestJSON(`/api/odoo/search-domain`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return r?.records ?? r?.result ?? r ?? [];
}

export async function callOdoo(payload: { model: string; method: string; args: any[]; kwargs?: Record<string, any> }): Promise<any> {
  return await requestJSON(`/api/odoo/call`, { method: "POST", body: JSON.stringify(payload) });
}

// createOdoo: return number (DialogApp expects number)
export async function createOdoo(model: string, values: Record<string, any>): Promise<number> {
  // prefer dedicated endpoint
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
export async function aiSelftest(): Promise<{ ok: boolean; text?: string; error?: string }> {
  return await requestJSON(`/api/ai/selftest`, { method: "POST", body: "{}" });
}

export async function aiGenerate(payload: any): Promise<AiGenerateResponse> {
  return await requestJSON(`/api/ai/generate`, { method: "POST", body: JSON.stringify(payload) });
}
