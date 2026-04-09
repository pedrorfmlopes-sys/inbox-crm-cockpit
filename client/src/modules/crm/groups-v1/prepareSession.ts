export type GroupsPrepareSubview = "list" | "attachments" | "summary";
export type GroupsPrepareAttachmentMode = "all" | "with" | "without";
export type GroupsPrepareGroupMode = "all" | "with_group" | "without_group";

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

export interface GroupPreparationSeed {
  anchorEmailKey: string;
  selectedEmailKeys: string[];
  selectedAttachmentKeys: string[];
  workingGroupId: string;
  filterQuery: string;
  attachmentMode: GroupsPrepareAttachmentMode;
  groupMode: GroupsPrepareGroupMode;
  openedAtIso: string;
}

export const GROUPS_PREPARE_SESSION_STORAGE_PREFIX = "iccc_groups_prepare_session_v1:";
export const GROUPS_PREPARE_CLASSIFY_SEED_STORAGE_PREFIX = "iccc_groups_prepare_classify_seed_v1:";
export const GROUPS_PREPARE_CLASSIFY_PARAM = "prepareSeedKey";

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

export function buildGroupsPrepareSessionKey(anchorEmailKey: string | null | undefined): string {
  const key = normalizeToken(anchorEmailKey);
  return key ? `${GROUPS_PREPARE_SESSION_STORAGE_PREFIX}${key}` : "";
}

export function sanitizeGroupsPrepareSessionState(
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
    updatedAtIso: normalizeToken(input?.updatedAtIso),
  };
}

export function readGroupsPrepareSession(anchorEmailKey: string | null | undefined): GroupsPrepareSessionState {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  if (!storageKey || typeof sessionStorage === "undefined") {
    return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
  }
  try {
    const raw = sessionStorage.getItem(storageKey);
    if (!raw) return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
    return sanitizeGroupsPrepareSessionState(JSON.parse(raw));
  } catch {
    return { ...DEFAULT_GROUPS_PREPARE_SESSION_STATE };
  }
}

export function writeGroupsPrepareSession(
  anchorEmailKey: string | null | undefined,
  state: Partial<GroupsPrepareSessionState>
): boolean {
  const storageKey = buildGroupsPrepareSessionKey(anchorEmailKey);
  if (!storageKey || typeof sessionStorage === "undefined") return false;
  try {
    sessionStorage.setItem(storageKey, JSON.stringify(sanitizeGroupsPrepareSessionState({
      ...state,
      updatedAtIso: new Date().toISOString(),
    })));
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
  return {
    anchorEmailKey,
    selectedEmailKeys: normalizeUniqueList(input.selectedEmailKeys),
    selectedAttachmentKeys: normalizeUniqueList(input.selectedAttachmentKeys),
    workingGroupId: normalizeToken(input.workingGroupId),
    filterQuery: normalizeToken(input.filterQuery),
    attachmentMode: normalizeAttachmentMode(input.attachmentMode),
    groupMode: normalizeGroupMode(input.groupMode),
    openedAtIso: new Date().toISOString(),
  };
}

export function writeGroupPreparationSeed(seed: GroupPreparationSeed | null): string {
  if (!seed || typeof localStorage === "undefined") return "";
  const key = `${GROUPS_PREPARE_CLASSIFY_SEED_STORAGE_PREFIX}${Date.now()}:${seed.anchorEmailKey}`;
  try {
    localStorage.setItem(key, JSON.stringify(seed));
    return key;
  } catch {
    return "";
  }
}
