const WORKSET_KIND = "groups_v1_workset_manifest";
const WORKSET_VERSION = 1;
const STORAGE_MODES = new Set(["supabase", "local_device", "chosen_folder", "hybrid"]);
const LOCATION_KINDS = new Set(["session", "supabase", "local_device", "filesystem", "document_library", "hybrid"]);
const LOCATION_PROVIDERS = new Set(["cloud", "local", "onedrive", "supabase", "hybrid"]);
const PROMOTION_STATES = new Set(["not_requested", "pending", "promoted", "partial", "skipped", "blocked"]);
const PROMOTION_SCOPES = new Set(["manifest", "email_metadata", "attachment_metadata", "attachment_binary"]);

function normalizeString(value) {
  return String(value || "").trim();
}

function normalizeIsoString(value, fallback = "") {
  const raw = normalizeString(value);
  if (!raw) return fallback;
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? fallback : parsed.toISOString();
}

function normalizeUniqueStringList(values) {
  return Array.from(new Set((Array.isArray(values) ? values : []).map((value) => normalizeString(value)).filter(Boolean)));
}

function normalizePromotionScopeList(values) {
  return normalizeUniqueStringList(values).filter((value) => PROMOTION_SCOPES.has(value));
}

function normalizeAttachmentSelection(value) {
  const normalized = normalizeString(value).toLowerCase();
  if (normalized === "pending") return "pending";
  if (normalized === "rejected") return "rejected";
  return "selected";
}

function normalizeStorageDisposition(value) {
  const normalized = normalizeString(value).toLowerCase();
  if (normalized === "reference" || normalized === "skip") return normalized;
  if (normalized === "binary") return "binary";
  return undefined;
}

function normalizeFilters(input) {
  return {
    query: normalizeString(input?.query) || undefined,
    fromEmail: normalizeString(input?.fromEmail) || undefined,
    labels: normalizeUniqueStringList(input?.labels),
    dateFromIso: normalizeIsoString(input?.dateFromIso) || undefined,
    dateToIso: normalizeIsoString(input?.dateToIso) || undefined,
    attachmentMode: input?.attachmentMode === "with" || input?.attachmentMode === "without" ? input.attachmentMode : "all",
    groupMode: input?.groupMode === "with_group" || input?.groupMode === "without_group" ? input.groupMode : "all",
  };
}

function normalizeAttachments(values) {
  return (Array.isArray(values) ? values : [])
    .map((entry) => ({
      key: normalizeString(entry?.key),
      emailKey: normalizeString(entry?.emailKey),
      name: normalizeString(entry?.name),
      contentType: normalizeString(entry?.contentType) || undefined,
      size: Number(entry?.size || 0) || undefined,
      isInline: entry?.isInline === true,
      hasContent: entry?.hasContent === true,
      selection: normalizeAttachmentSelection(entry?.selection),
      storageDisposition: normalizeStorageDisposition(entry?.storageDisposition),
      requiresDecision: entry?.requiresDecision === true,
    }))
    .filter((entry) => entry.key && entry.emailKey && entry.name);
}

function normalizeLocation(input, fallbackLabel = "Sessao local") {
  const kind = normalizeString(input?.kind);
  const provider = normalizeString(input?.provider);
  return {
    kind: LOCATION_KINDS.has(kind) ? kind : "session",
    provider: LOCATION_PROVIDERS.has(provider) ? provider : "cloud",
    label: normalizeString(input?.label) || fallbackLabel,
    basePath: normalizeString(input?.basePath) || undefined,
    relativePath: normalizeString(input?.relativePath) || undefined,
    folderHint: normalizeString(input?.folderHint) || undefined,
    isRemote: input?.isRemote === true,
    isConfigured: input?.isConfigured === true,
  };
}

function normalizePromotion(input) {
  const state = normalizeString(input?.state);
  return {
    state: PROMOTION_STATES.has(state) ? state : "not_requested",
    lastAttemptAtIso: normalizeIsoString(input?.lastAttemptAtIso) || undefined,
    promotedScopes: normalizePromotionScopeList(input?.promotedScopes),
    blockedScopes: normalizePromotionScopeList(input?.blockedScopes),
    note: normalizeString(input?.note) || undefined,
  };
}

export function normalizeGroupWorksetManifest(input) {
  const worksetKey = normalizeString(input?.worksetKey);
  const anchorEmailKey = normalizeString(input?.anchorEmailKey);
  const storageMode = normalizeString(input?.storageMode);
  if (!worksetKey || !anchorEmailKey || !STORAGE_MODES.has(storageMode)) return null;

  const createdAtIso = normalizeIsoString(input?.createdAtIso, new Date().toISOString());
  const updatedAtIso = normalizeIsoString(input?.updatedAtIso, createdAtIso || new Date().toISOString());

  return {
    kind: WORKSET_KIND,
    version: WORKSET_VERSION,
    worksetKey,
    createdAtIso,
    updatedAtIso,
    storageMode,
    anchorEmailKey,
    includedEmailKeys: normalizeUniqueStringList(input?.includedEmailKeys),
    workingGroupId: normalizeString(input?.workingGroupId) || undefined,
    workingGroupName: normalizeString(input?.workingGroupName) || undefined,
    filters: normalizeFilters(input?.filters),
    attachments: normalizeAttachments(input?.attachments),
    mainLocation: normalizeLocation(input?.mainLocation, "Persistencia principal"),
    remotePromotionLocation: input?.remotePromotionLocation
      ? normalizeLocation(input.remotePromotionLocation, "Promocao remota")
      : null,
    promotion: normalizePromotion(input?.promotion),
  };
}

export function hasMeaningfulGroupWorksetPayload(input) {
  const manifest = normalizeGroupWorksetManifest(input);
  if (!manifest) return false;
  return Boolean(
    manifest.includedEmailKeys.length
    || manifest.attachments.length
    || manifest.workingGroupId
    || manifest.workingGroupName
    || manifest.filters.query
    || manifest.filters.fromEmail
    || manifest.filters.labels.length
    || manifest.filters.dateFromIso
    || manifest.filters.dateToIso
    || manifest.filters.attachmentMode !== "all"
    || manifest.filters.groupMode !== "all"
  );
}

export function buildGroupWorksetPayloadScore(input) {
  const manifest = normalizeGroupWorksetManifest(input);
  if (!manifest) return 0;
  return (
    manifest.includedEmailKeys.length * 4
    + manifest.attachments.length * 3
    + (manifest.workingGroupId ? 3 : 0)
    + (manifest.workingGroupName ? 2 : 0)
    + (manifest.filters.query ? 2 : 0)
    + (manifest.filters.fromEmail ? 1 : 0)
    + manifest.filters.labels.length
    + (manifest.filters.attachmentMode !== "all" ? 1 : 0)
    + (manifest.filters.groupMode !== "all" ? 1 : 0)
  );
}

export function mergeGroupWorksetManifest(currentInput, incomingInput) {
  const incoming = normalizeGroupWorksetManifest(incomingInput);
  if (!incoming) return normalizeGroupWorksetManifest(currentInput);

  const current = normalizeGroupWorksetManifest(currentInput);
  if (!current) return incoming;
  if (!hasMeaningfulGroupWorksetPayload(incoming) && hasMeaningfulGroupWorksetPayload(current)) {
    return current;
  }

  return normalizeGroupWorksetManifest({
    ...current,
    ...incoming,
    createdAtIso: current.createdAtIso,
    updatedAtIso: incoming.updatedAtIso || new Date().toISOString(),
    mainLocation: incoming.mainLocation || current.mainLocation,
    remotePromotionLocation: incoming.remotePromotionLocation === undefined
      ? current.remotePromotionLocation
      : incoming.remotePromotionLocation,
    promotion: {
      ...current.promotion,
      ...incoming.promotion,
      promotedScopes: Array.isArray(incoming.promotion?.promotedScopes)
        ? incoming.promotion.promotedScopes
        : current.promotion.promotedScopes,
      blockedScopes: Array.isArray(incoming.promotion?.blockedScopes)
        ? incoming.promotion.blockedScopes
        : current.promotion.blockedScopes,
    },
  });
}
