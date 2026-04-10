export type GroupsPrepareSubview = "list" | "attachments" | "summary";
export type GroupsPrepareAttachmentMode = "all" | "with" | "without";
export type GroupsPrepareGroupMode = "all" | "with_group" | "without_group";
export type GroupsPrepareSessionSaveReason =
  | "manual"
  | "debounced"
  | "before_exit"
  | "before_context_change"
  | "before_subview_change"
  | "before_classify";

export interface GroupsPrepareSessionState {
  subview: GroupsPrepareSubview;
  showGroupPanel: boolean;
  showFiltersPanel: boolean;
  workingGroupId: string;
  workingGroupQuery: string;
  filterQuery: string;
  attachmentMode: GroupsPrepareAttachmentMode;
  groupMode: GroupsPrepareGroupMode;
  selectedEmailKeys: string[];
  expandedEmailKeys: string[];
  selectedAttachmentKeys: string[];
  updatedAtIso: string;
}

export interface GroupsPrepareSessionRecord {
  kind: "groups_prepare_session";
  version: 1;
  anchorEmailKey: string;
  storage: "sessionStorage";
  isCanonical: false;
  savedAtIso: string;
  lastReason: GroupsPrepareSessionSaveReason;
  state: GroupsPrepareSessionState;
}

export interface GroupPreparationSeed {
  kind: "groups_prepare_classify_seed";
  version: 1;
  anchorEmailKey: string;
  selectedEmailKeys: string[];
  selectedAttachmentKeys: string[];
  workingGroupId: string;
  filterQuery: string;
  attachmentMode: GroupsPrepareAttachmentMode;
  groupMode: GroupsPrepareGroupMode;
  savedAtIso: string;
  expiresAtIso: string;
}

export const GROUPS_PREPARE_SESSION_STORAGE_PREFIX = "iccc_groups_prepare_session_v1:";
export const GROUPS_PREPARE_CLASSIFY_SEED_STORAGE_PREFIX = "iccc_groups_prepare_classify_seed_v1:";
export const GROUPS_PREPARE_CLASSIFY_PARAM = "prepareSeedKey";
export const GROUPS_PREPARE_SESSION_SAVE_DEBOUNCE_MS = 700;
export const GROUPS_PREPARE_CLASSIFY_SEED_TTL_MS = 12 * 60 * 60 * 1000;
export const GROUPS_PREPARE_SESSION_POLICY = {
  storage: "sessionStorage",
  isCanonical: false,
  remotePromotion: "future-phase",
  includes: [
    "subview",
    "showGroupPanel",
    "showFiltersPanel",
    "workingGroupId",
    "workingGroupQuery",
    "filterQuery",
    "attachmentMode",
    "groupMode",
    "selectedEmailKeys",
    "expandedEmailKeys",
    "selectedAttachmentKeys",
  ],
  excludes: [
    "persisted email bodies/html",
    "attachment binary content",
    "known email search results",
    "group catalog payloads",
    "backend persistence state",
    "final classification payload",
  ],
} as const;

export const DEFAULT_GROUPS_PREPARE_SESSION_STATE: GroupsPrepareSessionState = {
  subview: "list",
  showGroupPanel: false,
  showFiltersPanel: false,
  workingGroupId: "",
  workingGroupQuery: "",
  filterQuery: "",
  attachmentMode: "all",
  groupMode: "all",
  selectedEmailKeys: [],
  expandedEmailKeys: [],
  selectedAttachmentKeys: [],
  updatedAtIso: "",
};

function normalizeToken(value: string | null | undefined): string {
  return String(value || "").trim();
}

function normalizeUniqueList(values: unknown): string[] {
  if (!Array.isArray(values)) return [];
  return Array.from(new Set(values.map((value) => normalizeToken(String(value || ""))).filter(Boolean)));
}

function normalizeSubview(value: unknown): GroupsPrepareSubview {
  return value === "attachments" || value === "summary" ? value : "list";
}

function normalizeAttachmentMode(value: unknown): GroupsPrepareAttachmentMode {
  return value === "with" || value === "without" ? value : "all";
}

function normalizeGroupMode(value: unknown): GroupsPrepareGroupMode {
  return value === "with_group" || value === "without_group" ? value : "all";
}

function normalizeIsoDate(value: unknown): string {
  const raw = normalizeToken(typeof value === "string" ? value : "");
  if (!raw) return "";
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? "" : parsed.toISOString();
}

function isGroupsPrepareSessionRecord(value: unknown): value is GroupsPrepareSessionRecord {
  return Boolean(
    value
    && typeof value === "object"
    && (value as GroupsPrepareSessionRecord).kind === "groups_prepare_session"
    && (value as GroupsPrepareSessionRecord).version === 1
    && (value as GroupsPrepareSessionRecord).state
  );
}

function isGroupPreparationSeed(value: unknown): value is GroupPreparationSeed {
  return Boolean(
    value
    && typeof value === "object"
    && (value as GroupPreparationSeed).kind === "groups_prepare_classify_seed"
    && (value as GroupPreparationSeed).version === 1
    && normalizeToken((value as GroupPreparationSeed).anchorEmailKey)
  );
}

export function buildGroupsPrepareSessionKey(anchorEmailKey: string | null | undefined): string {
  const key = normalizeToken(anchorEmailKey);
  return key ? `${GROUPS_PREPARE_SESSION_STORAGE_PREFIX}${key}` : "";
}

export function buildGroupsPrepareSessionSnapshot(
  input: Partial<GroupsPrepareSessionState> | null | undefined
): GroupsPrepareSessionState {
  return {
    subview: normalizeSubview(input?.subview),
    showGroupPanel: input?.showGroupPanel === true,
    showFiltersPanel: input?.showFiltersPanel === true,
    workingGroupId: normalizeToken(input?.workingGroupId),
    workingGroupQuery: normalizeToken(input?.workingGroupQuery),
    filterQuery: normalizeToken(input?.filterQuery),
    attachmentMode: normalizeAttachmentMode(input?.attachmentMode),
    groupMode: normalizeGroupMode(input?.groupMode),
    selectedEmailKeys: normalizeUniqueList(input?.selectedEmailKeys),
    expandedEmailKeys: normalizeUniqueList(input?.expandedEmailKeys),
    selectedAttachmentKeys: normalizeUniqueList(input?.selectedAttachmentKeys),
    updatedAtIso: normalizeIsoDate(input?.updatedAtIso),
  };
}

export function sanitizeGroupsPrepareSessionState(
  input: Partial<GroupsPrepareSessionState> | null | undefined
): GroupsPrepareSessionState {
  return buildGroupsPrepareSessionSnapshot(input);
}

export function getGroupsPrepareSessionSignature(
  input: Partial<GroupsPrepareSessionState> | null | undefined
): string {
  const snapshot = buildGroupsPrepareSessionSnapshot(input);
  return JSON.stringify({
    ...snapshot,
    updatedAtIso: "",
  });
}

export function readGroupsPrepareSession(anchorEmailKey: string | null | undefined): GroupsPrepareSessionState {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  if (!storageKey || typeof sessionStorage === "undefined") {
    return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
  }
  try {
    const raw = sessionStorage.getItem(storageKey);
    if (!raw) return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
    const parsed = JSON.parse(raw);
    if (isGroupsPrepareSessionRecord(parsed)) {
      return buildGroupsPrepareSessionSnapshot({
        ...parsed.state,
        updatedAtIso: parsed.state.updatedAtIso || parsed.savedAtIso,
      });
    }
    return buildGroupsPrepareSessionSnapshot(parsed);
  } catch {
    return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
  }
}

export function hasGroupsPrepareSession(anchorEmailKey: string | null | undefined): boolean {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  if (!storageKey || typeof sessionStorage === "undefined") return false;
  try {
    return Boolean(sessionStorage.getItem(storageKey));
  } catch {
    return false;
  }
}

export function writeGroupsPrepareSession(
  anchorEmailKey: string | null | undefined,
  state: Partial<GroupsPrepareSessionState>,
  options?: { reason?: GroupsPrepareSessionSaveReason }
): boolean {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  const normalizedAnchorEmailKey = normalizeToken(anchorEmailKey);
  if (!storageKey || !normalizedAnchorEmailKey || typeof sessionStorage === "undefined") return false;
  try {
    const savedAtIso = new Date().toISOString();
    const snapshot = buildGroupsPrepareSessionSnapshot({
      ...state,
      updatedAtIso: savedAtIso,
    });
    const record: GroupsPrepareSessionRecord = {
      kind: "groups_prepare_session",
      version: 1,
      anchorEmailKey: normalizedAnchorEmailKey,
      storage: "sessionStorage",
      isCanonical: false,
      savedAtIso,
      lastReason: options?.reason || "manual",
      state: snapshot,
    };
    sessionStorage.setItem(storageKey, JSON.stringify(record));
    return true;
  } catch {
    return false;
  }
}

export function clearGroupsPrepareSession(anchorEmailKey: string | null | undefined): void {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  if (!storageKey || typeof sessionStorage === "undefined") return;
  try {
    sessionStorage.removeItem(storageKey);
  } catch {
    // ignore
  }
}

function getGroupPreparationSeedExpiry(savedAtIso: string): string {
  return new Date(new Date(savedAtIso).getTime() + GROUPS_PREPARE_CLASSIFY_SEED_TTL_MS).toISOString();
}

function cleanupStaleGroupPreparationSeeds(now = Date.now()): void {
  if (typeof localStorage === "undefined") return;
  try {
    const staleKeys: string[] = [];
    for (let index = 0; index < localStorage.length; index += 1) {
      const key = localStorage.key(index);
      if (!key || !key.startsWith(GROUPS_PREPARE_CLASSIFY_SEED_STORAGE_PREFIX)) continue;
      const raw = localStorage.getItem(key);
      if (!raw) {
        staleKeys.push(key);
        continue;
      }
      try {
        const parsed = JSON.parse(raw);
        if (!isGroupPreparationSeed(parsed)) {
          staleKeys.push(key);
          continue;
        }
        const expiresAt = normalizeIsoDate(parsed.expiresAtIso);
        if (!expiresAt || new Date(expiresAt).getTime() <= now) {
          staleKeys.push(key);
        }
      } catch {
        staleKeys.push(key);
      }
    }
    staleKeys.forEach((key) => localStorage.removeItem(key));
  } catch {
    // ignore
  }
}

export function buildGroupPreparationSeed(input: {
  anchorEmailKey: string | null | undefined;
  selectedEmailKeys?: string[];
  selectedAttachmentKeys?: string[];
  workingGroupId?: string | null;
  filterQuery?: string;
  attachmentMode?: GroupsPrepareAttachmentMode;
  groupMode?: GroupsPrepareGroupMode;
}): GroupPreparationSeed | null {
  const anchorEmailKey = normalizeToken(input.anchorEmailKey);
  if (!anchorEmailKey) return null;
  const savedAtIso = new Date().toISOString();
  return {
    kind: "groups_prepare_classify_seed",
    version: 1,
    anchorEmailKey,
    selectedEmailKeys: normalizeUniqueList(input.selectedEmailKeys),
    selectedAttachmentKeys: normalizeUniqueList(input.selectedAttachmentKeys),
    workingGroupId: normalizeToken(input.workingGroupId),
    filterQuery: normalizeToken(input.filterQuery),
    attachmentMode: normalizeAttachmentMode(input.attachmentMode),
    groupMode: normalizeGroupMode(input.groupMode),
    savedAtIso,
    expiresAtIso: getGroupPreparationSeedExpiry(savedAtIso),
  };
}

export function readGroupPreparationSeed(seedKey: string | null | undefined): GroupPreparationSeed | null {
  const key = normalizeToken(seedKey);
  if (!key || typeof localStorage === "undefined") return null;
  try {
    cleanupStaleGroupPreparationSeeds();
    const raw = localStorage.getItem(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!isGroupPreparationSeed(parsed)) return null;
    const expiresAt = normalizeIsoDate(parsed.expiresAtIso);
    if (!expiresAt || new Date(expiresAt).getTime() <= Date.now()) {
      localStorage.removeItem(key);
      return null;
    }
    return {
      ...parsed,
      savedAtIso: normalizeIsoDate(parsed.savedAtIso) || new Date().toISOString(),
      expiresAtIso: expiresAt,
    };
  } catch {
    return null;
  }
}

export function consumeGroupPreparationSeed(seedKey: string | null | undefined): GroupPreparationSeed | null {
  const key = normalizeToken(seedKey);
  const seed = readGroupPreparationSeed(key);
  if (seed && typeof localStorage !== "undefined") {
    try {
      localStorage.removeItem(key);
    } catch {
      // ignore
    }
  }
  return seed;
}

export function clearGroupPreparationSeed(seedKey: string | null | undefined): void {
  const key = normalizeToken(seedKey);
  if (!key || typeof localStorage === "undefined") return;
  try {
    localStorage.removeItem(key);
  } catch {
    // ignore
  }
}

export function writeGroupPreparationSeed(seed: GroupPreparationSeed | null): string {
  if (!seed || typeof localStorage === "undefined") return "";
  const key = `${GROUPS_PREPARE_CLASSIFY_SEED_STORAGE_PREFIX}${Date.now()}:${seed.anchorEmailKey}`;
  try {
    cleanupStaleGroupPreparationSeeds();
    localStorage.setItem(key, JSON.stringify(seed));
    return key;
  } catch {
    return "";
  }
}
