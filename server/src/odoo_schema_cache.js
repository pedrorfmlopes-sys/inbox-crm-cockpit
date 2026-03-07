const DEFAULT_TTL_MS = Number(process.env.ODOO_SCHEMA_TTL_MS || 12 * 60 * 60 * 1000);

function nowMs() {
  return Date.now();
}

function isInvalidFieldError(error) {
  const msg = String(error?.message || error || "");
  return /invalid field/i.test(msg);
}

export function createOdooSchemaCache({ fetchFields, ttlMs = DEFAULT_TTL_MS }) {
  if (typeof fetchFields !== "function") {
    throw new Error("createOdooSchemaCache requires fetchFields(model) function");
  }

  const cache = new Map();

  async function getModelFields(model, { forceRefresh = false } = {}) {
    const m = String(model || "").trim();
    if (!m) return new Set();

    const current = cache.get(m);
    const expired = !current || nowMs() - current.fetchedAt > ttlMs;
    if (!forceRefresh && current && !expired) return current.fields;

    const raw = await fetchFields(m);
    const fields = new Set(Object.keys(raw || {}));
    cache.set(m, { fields, fetchedAt: nowMs() });
    return fields;
  }

  function invalidateModel(model) {
    cache.delete(String(model || "").trim());
  }

  function invalidateAll() {
    cache.clear();
  }

  async function filterReadFields(model, wantedFields = []) {
    const wanted = Array.isArray(wantedFields) ? wantedFields.filter(Boolean) : [];
    if (!wanted.length) return [];
    const fields = await getModelFields(model);
    return wanted.filter((f) => fields.has(f));
  }

  async function sanitizeWriteData(model, data = {}) {
    if (!data || typeof data !== "object") return {};
    const fields = await getModelFields(model);
    const out = {};
    for (const [k, v] of Object.entries(data)) {
      if (fields.has(k)) out[k] = v;
    }
    return out;
  }

  async function safeSearchRead(model, domain, wantedFields, limit = 10, opts = {}, onSearchRead) {
    if (typeof onSearchRead !== "function") {
      throw new Error("safeSearchRead requires onSearchRead callback");
    }

    const filteredFields = await filterReadFields(model, wantedFields);

    try {
      return await onSearchRead(model, domain, filteredFields, limit, opts);
    } catch (e) {
      if (!isInvalidFieldError(e)) throw e;
      invalidateModel(model);
      const refreshedFields = await filterReadFields(model, wantedFields);
      return await onSearchRead(model, domain, refreshedFields, limit, opts);
    }
  }

  async function safeCreate(model, data, onCreate) {
    if (typeof onCreate !== "function") {
      throw new Error("safeCreate requires onCreate callback");
    }

    let clean = await sanitizeWriteData(model, data);

    try {
      return await onCreate(model, clean);
    } catch (e) {
      if (!isInvalidFieldError(e)) throw e;
      invalidateModel(model);
      clean = await sanitizeWriteData(model, data);
      return await onCreate(model, clean);
    }
  }

  async function safeWrite(model, ids, data, onWrite) {
    if (typeof onWrite !== "function") {
      throw new Error("safeWrite requires onWrite callback");
    }

    let clean = await sanitizeWriteData(model, data);

    try {
      return await onWrite(model, ids, clean);
    } catch (e) {
      if (!isInvalidFieldError(e)) throw e;
      invalidateModel(model);
      clean = await sanitizeWriteData(model, data);
      return await onWrite(model, ids, clean);
    }
  }

  return {
    getModelFields,
    invalidateModel,
    invalidateAll,
    filterReadFields,
    sanitizeWriteData,
    safeSearchRead,
    safeCreate,
    safeWrite,
    isInvalidFieldError,
  };
}
