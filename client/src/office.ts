import { clientLog } from "./logger";
import { getLinks, getRelatedEmailContext } from "./api";
import { getSettings } from "./settings";
import {
  buildOutlookCategorySourceFromRelatedContext,
  GROUP_CATEGORY_PREFIX,
  REFERENCE_CATEGORY_PREFIX,
  TICKET_CATEGORY_PREFIX,
  LEGACY_STATUS_CATEGORY_PREFIX,
  GROUP_STATUS_CATEGORY_PREFIX,
  TICKET_STATUS_CATEGORY_PREFIX,
  LABEL_STATUS_CATEGORY_PREFIX,
  LEGACY_TICKET_CATEGORY_PREFIX,
  LEGACY_LABEL_CATEGORY_PREFIX,
  ODOO_LINKED_CATEGORY,
  buildOutlookCategoryPlan,
  buildOutlookCategorySourceFromLegacyInput,
  getOutlookCategoryPlanSignature,
  getOutlookCategorySourceSignature,
  isManagedCategoryFamilyName,
  isReservedOutlookCategoryName,
  normalizeOutlookCategorySource,
  normalizeUniqueCategoryValues,
  type LegacyManagedOutlookCategoryInput,
  type OutlookCategoryPlan,
  type OutlookCategorySource,
} from "./outlookCategories";



export type Recipient = { name: string; email: string };

export type OutlookMessageContext = {
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  internetMessageId?: string;
  conversationId?: string;
  itemId?: string;
  receivedDateTimeIso?: string;

  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  isCompose?: boolean;
};

export type OutlookAttachment = {
  id?: string;
  name: string;
  contentType: string;
  content: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
};

const GRAPH_NAA_CLIENT_ID = "67f42759-4576-461a-b87b-c78332f7a1e7";
const GRAPH_NAA_TENANT_ID = "7eff32de-43f1-447b-af3e-af8d2939e93d";
const GRAPH_NAA_AUTHORITY = `https://login.microsoftonline.com/${GRAPH_NAA_TENANT_ID}`;
const GRAPH_ATTACHMENT_SCOPES = ["Mail.Read", "User.Read"];
const GRAPH_PEOPLE_SCOPES = ["People.Read", "User.Read"];
export const GRAPH_DRIVE_SELF_TEST_SCOPES = ["openid", "profile", "offline_access", "User.Read", "Files.ReadWrite"] as const;
const GRAPH_NAA_REDIRECT_URI = `${window.location.origin}/`;
const GRAPH_RUNTIME_ENABLED = false;

let nestableMsalPromise: Promise<any> | null = null;
let browserMsalPromise: Promise<any> | null = null;
export const OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT = "iccc-outlook-category-context-invalidated";
export const OUTLOOK_CATEGORY_SYNC_REQUEST_EVENT = "iccc-outlook-category-sync-request";
export const OUTLOOK_CATEGORY_SYNC_REQUEST_STORAGE_KEY = "iccc-outlook-category-sync-request-v1";
export const OUTLOOK_CATEGORY_SYNC_DEBUG_STORAGE_KEY = "iccc-outlook-category-sync-debug-v1";
export const OUTLOOK_CATEGORY_OPERATION_DEBUG_STORAGE_KEY = "iccc-outlook-category-op-debug-v1";
export const OUTLOOK_CATEGORY_SYNC_RESULT_EVENT = "iccc-outlook-category-sync-result";
export const OUTLOOK_CATEGORY_SYNC_RESULT_STORAGE_KEY = "iccc-outlook-category-sync-result-v1";
const HOST_ACTION_WINDOW_MESSAGE_TYPE = "iccc-host-action-window";
const HOST_ACTION_WINDOW_RESULT_TYPE = "iccc-host-action-window-result";
const OUTLOOK_CATEGORY_OPERATION_STORAGE_PREFIX = "iccc-outlook-category-op-v1:";
const OUTLOOK_CATEGORY_OPERATION_ACTIVE_PREFIX = "iccc-outlook-category-op-active-v1:";
const OUTLOOK_CATEGORY_OPERATION_DEFAULT_LEASE_MS = 45_000;

export type OutlookCategorySyncTarget = {
  itemId?: string;
  internetMessageId?: string;
  conversationId?: string;
};

export type GraphDriveSelfTestConclusion =
  | "tenant_allows_user_write"
  | "tenant_blocks_user_write"
  | "implementation_cannot_complete_test";

export type GraphDriveSelfTestStep = {
  attempted: boolean;
  ok: boolean;
  status?: number;
  errorCode?: string;
  errorMessage?: string;
  detail?: string;
  response?: unknown;
};

export type GraphDriveSelfTestResult = {
  scopes: string[];
  authMode: "nested_app_auth" | "browser_msal" | "unavailable";
  consent: GraphDriveSelfTestStep & {
    result: "accepted" | "need_admin_approval" | "auth_error" | "timeout" | "not_available";
    account?: string;
  };
  meDrive: GraphDriveSelfTestStep;
  createFolder: GraphDriveSelfTestStep & {
    folderId?: string;
    folderName?: string;
  };
  cleanup: GraphDriveSelfTestStep;
  conclusion: GraphDriveSelfTestConclusion;
  conclusionMessage: string;
};

export type OutlookCategorySyncRequest = {
  requestId: string;
  createdAtIso: string;
  reason?: string;
  operationId?: string;
  mode: "source" | "current-item-context";
  target?: OutlookCategorySyncTarget;
  source?: Partial<OutlookCategorySource> | null;
};

type OutlookCategorySyncMode = "source" | "current-item-context";
export type OutlookCategoryWriterResult =
  | "success"
  | "failed"
  | "item-mismatch"
  | "stale"
  | "duplicate"
  | "cancelled"
  | "timeout";

export type OutlookCategoryOperationPhase =
  | "opening"
  | "saving"
  | "refreshing"
  | "rehydrating"
  | "planning"
  | "writingOutlook"
  | "verifying"
  | "completed"
  | "failed"
  | "cancelled";

export type OutlookCategoryOperationStatus =
  | "active"
  | "completed"
  | "failed"
  | "cancelled"
  | "timeout";

export type OutlookCategoryOperationRecord = {
  operationId: string;
  itemIdentity: string;
  target?: OutlookCategorySyncTarget;
  owner: string;
  startedAtIso: string;
  startedAtMs: number;
  lastUpdatedAtIso: string;
  phase: OutlookCategoryOperationPhase;
  status: OutlookCategoryOperationStatus;
  result?: OutlookCategoryWriterResult;
  requestId?: string;
  expectedItemToken?: string;
  leaseExpiresAtMs: number;
};

export type OutlookCategorySyncResult = {
  requestId: string;
  operationId?: string;
  reason: string;
  mode: OutlookCategorySyncMode;
  itemIdentity: string;
  target?: OutlookCategorySyncTarget;
  sourceSignature?: string;
  planSignature?: string;
  result: OutlookCategoryWriterResult;
  detail?: string;
  finishedAtIso: string;
};

type OutlookCategoryWriterShortCircuit = {
  result: OutlookCategoryWriterResult;
  itemIdentity: string;
  detail?: string;
};

type OutlookCategorySyncWriterRequest = {
  requestId: string;
  operationId?: string;
  requestedAtMs: number;
  reason: string;
  mode: OutlookCategorySyncMode;
  target?: OutlookCategorySyncTarget;
  source?: Partial<OutlookCategorySource> | null;
  expectedItemToken?: string;
  manageClassificationFamilies?: boolean;
};

type PreparedOutlookCategorySyncWriterRequest = OutlookCategorySyncWriterRequest & {
  itemIdentity: string;
  expectedItemToken: string;
  source: OutlookCategorySource;
  sourceSignature: string;
  plan: OutlookCategoryPlan;
  planSignature: string;
};

type OutlookCategoryMutationResult = {
  ok: boolean;
  rawStatus?: string;
  error?: string;
};

type OutlookCategoryReadbackResult = {
  categories: string[];
  source: "getAsync" | "array-fallback" | "unavailable";
  rawStatus?: string;
  error?: string;
};

type OutlookCategoryPlanDiff = {
  currentCategories: string[];
  currentManagedCategories: string[];
  desiredCategories: string[];
  toAdd: string[];
  toRemove: string[];
  missingManagedCategories: string[];
  unexpectedManagedCategories: string[];
};

type ApplyOutlookCategoryPlanResult = {
  result: "success" | "noop" | "failed" | "stale" | "item-mismatch";
  detail?: string;
  diff: OutlookCategoryPlanDiff;
};

type OutlookCategoryWriterFreshness = {
  requestedAtMs: number;
  order: number;
  requestId: string;
};

type OutlookCategoryWriterState = {
  tail: Promise<OutlookCategorySyncResult>;
  latestFreshness: OutlookCategoryWriterFreshness | null;
  lastAppliedPlanSignature: string;
  lastAppliedSourceSignature: string;
  recentRequestIds: string[];
};

const OUTLOOK_CATEGORY_SYNC_PREFIX = "[outlook-category-sync]";
const OUTLOOK_CATEGORY_WRITER_RECENT_REQUEST_LIMIT = 12;
const outlookCategoryWriterStateByItem = new Map<string, OutlookCategoryWriterState>();
let outlookCategoryWriterOrder = 0;

function sleep(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
}

function normalizeOutlookIdentityString(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function hasPreciseOutlookItemIdentity(context: Pick<OutlookMessageContext, "itemId" | "internetMessageId"> | null | undefined): boolean {
  return Boolean(String(context?.itemId || "").trim() || normalizeOutlookIdentityString(context?.internetMessageId));
}

function doesResolvedEmailMatchCurrentOutlookItem(
  email: { itemId?: string; internetMessageId?: string } | null | undefined,
  context: Pick<OutlookMessageContext, "itemId" | "internetMessageId">
): boolean {
  const currentItemId = String(context?.itemId || "").trim();
  const currentInternetMessageId = normalizeOutlookIdentityString(context?.internetMessageId);
  const emailItemId = String(email?.itemId || "").trim();
  const emailInternetMessageId = normalizeOutlookIdentityString(email?.internetMessageId);

  if (currentItemId) return Boolean(emailItemId) && emailItemId === currentItemId;
  if (currentInternetMessageId) return Boolean(emailInternetMessageId) && emailInternetMessageId === currentInternetMessageId;
  return false;
}

function isOutlookCategorySyncDebugEnabled(): boolean {
  try {
    return Boolean((import.meta as any)?.env?.DEV)
      || window.localStorage?.getItem(OUTLOOK_CATEGORY_SYNC_DEBUG_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

function isOutlookCategoryOperationDebugEnabled(): boolean {
  try {
    return isOutlookCategorySyncDebugEnabled()
      || window.localStorage?.getItem(OUTLOOK_CATEGORY_OPERATION_DEBUG_STORAGE_KEY) === "1";
  } catch {
    return isOutlookCategorySyncDebugEnabled();
  }
}

function logOutlookCategorySync(level: "debug" | "info" | "warn" | "error", event: string, data?: any) {
  if (!isOutlookCategorySyncDebugEnabled()) return;
  const message = `${OUTLOOK_CATEGORY_SYNC_PREFIX} ${event}`;
  if (level === "warn") {
    clientLog.warn(message, data);
    return;
  }
  if (level === "error") {
    clientLog.error(message, data);
    return;
  }
  if (level === "info") {
    clientLog.log(message, data);
    return;
  }
  clientLog.debug(message, data);
}

function logOutlookCategoryOperation(level: "debug" | "info" | "warn" | "error", event: string, data?: any) {
  if (!isOutlookCategoryOperationDebugEnabled()) return;
  const message = `[outlook-category-op] ${event}`;
  if (level === "warn") {
    clientLog.warn(message, data);
    return;
  }
  if (level === "error") {
    clientLog.error(message, data);
    return;
  }
  if (level === "info") {
    clientLog.log(message, data);
    return;
  }
  clientLog.debug(message, data);
}

function createOutlookCategorySyncRequestId(reason: string): string {
  const normalizedReason = String(reason || "sync").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-") || "sync";
  return `outlook-category-sync:${normalizedReason}:${Date.now()}:${Math.random().toString(36).slice(2)}`;
}

function normalizeOutlookCategorySyncTarget(target?: OutlookCategorySyncTarget | null): OutlookCategorySyncTarget | undefined {
  if (!target) return undefined;
  const itemId = String(target.itemId || "").trim() || undefined;
  const internetMessageId = String(target.internetMessageId || "").trim() || undefined;
  const conversationId = String(target.conversationId || "").trim() || undefined;
  if (!itemId && !internetMessageId && !conversationId) return undefined;
  return { itemId, internetMessageId, conversationId };
}

function buildOutlookCategoryWriterItemIdentity(target?: OutlookCategorySyncTarget | null): string {
  const normalized = normalizeOutlookCategorySyncTarget(target);
  const itemId = String(normalized?.itemId || "").trim();
  const internetMessageId = normalizeOutlookIdentityString(normalized?.internetMessageId);
  const conversationId = String(normalized?.conversationId || "").trim();
  if (itemId) return `item:${itemId}`;
  if (internetMessageId) return `internet:${internetMessageId}`;
  if (conversationId) return `conversation:${conversationId}`;
  return "";
}

function buildOutlookCategoryOperationItemIdentity(input?: {
  target?: OutlookCategorySyncTarget | null;
  expectedItemToken?: string;
}): string {
  const identity = buildOutlookCategoryWriterItemIdentity(input?.target);
  if (identity) return identity;
  const expectedItemToken = String(input?.expectedItemToken || "").trim();
  return expectedItemToken ? `token:${expectedItemToken}` : "";
}

function getOutlookCategoryOperationStorageKey(operationId: string): string {
  return `${OUTLOOK_CATEGORY_OPERATION_STORAGE_PREFIX}${String(operationId || "").trim()}`;
}

function getOutlookCategoryOperationActiveStorageKey(itemIdentity: string): string {
  return `${OUTLOOK_CATEGORY_OPERATION_ACTIVE_PREFIX}${String(itemIdentity || "").trim()}`;
}

function cloneOutlookCategoryOperationRecord(record: OutlookCategoryOperationRecord | null | undefined): OutlookCategoryOperationRecord | null {
  if (!record) return null;
  return {
    ...record,
    target: record.target ? { ...record.target } : undefined,
  };
}

function readOutlookCategoryOperationRecord(operationId: string): OutlookCategoryOperationRecord | null {
  if (typeof window === "undefined" || !window.localStorage) return null;
  const normalizedOperationId = String(operationId || "").trim();
  if (!normalizedOperationId) return null;
  try {
    const raw = window.localStorage.getItem(getOutlookCategoryOperationStorageKey(normalizedOperationId));
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    const itemIdentity = String(parsed.itemIdentity || "").trim();
    if (!itemIdentity) return null;
    return {
      operationId: normalizedOperationId,
      itemIdentity,
      target: normalizeOutlookCategorySyncTarget(parsed.target),
      owner: String(parsed.owner || "").trim() || "unknown",
      startedAtIso: String(parsed.startedAtIso || "").trim() || new Date().toISOString(),
      startedAtMs: Number(parsed.startedAtMs || 0) || Date.now(),
      lastUpdatedAtIso: String(parsed.lastUpdatedAtIso || "").trim() || new Date().toISOString(),
      phase: parsed.phase as OutlookCategoryOperationPhase || "opening",
      status: parsed.status as OutlookCategoryOperationStatus || "active",
      result: parsed.result as OutlookCategoryWriterResult | undefined,
      requestId: String(parsed.requestId || "").trim() || undefined,
      expectedItemToken: String(parsed.expectedItemToken || "").trim() || undefined,
      leaseExpiresAtMs: Number(parsed.leaseExpiresAtMs || 0) || 0,
    };
  } catch {
    return null;
  }
}

function persistOutlookCategoryOperationRecord(record: OutlookCategoryOperationRecord) {
  if (typeof window === "undefined" || !window.localStorage) return;
  try {
    window.localStorage.setItem(
      getOutlookCategoryOperationStorageKey(record.operationId),
      JSON.stringify(record)
    );
  } catch {
    // best effort
  }
}

function clearOutlookCategoryOperationActiveLock(itemIdentity: string, operationId?: string) {
  if (typeof window === "undefined" || !window.localStorage) return;
  const normalizedItemIdentity = String(itemIdentity || "").trim();
  if (!normalizedItemIdentity) return;
  try {
    const activeKey = getOutlookCategoryOperationActiveStorageKey(normalizedItemIdentity);
    const currentOperationId = String(window.localStorage.getItem(activeKey) || "").trim();
    if (!operationId || !currentOperationId || currentOperationId === String(operationId || "").trim()) {
      window.localStorage.removeItem(activeKey);
    }
  } catch {
    // best effort
  }
}

function isOutlookCategoryOperationExpired(record: OutlookCategoryOperationRecord | null | undefined): boolean {
  if (!record) return true;
  return record.status !== "active" || Number(record.leaseExpiresAtMs || 0) <= Date.now();
}

function readActiveOutlookCategoryOperationByItemIdentity(itemIdentity: string): OutlookCategoryOperationRecord | null {
  if (typeof window === "undefined" || !window.localStorage) return null;
  const normalizedItemIdentity = String(itemIdentity || "").trim();
  if (!normalizedItemIdentity) return null;
  try {
    const activeOperationId = String(window.localStorage.getItem(getOutlookCategoryOperationActiveStorageKey(normalizedItemIdentity)) || "").trim();
    if (!activeOperationId) return null;
    const record = readOutlookCategoryOperationRecord(activeOperationId);
    if (!record || record.itemIdentity !== normalizedItemIdentity || isOutlookCategoryOperationExpired(record)) {
      clearOutlookCategoryOperationActiveLock(normalizedItemIdentity, activeOperationId);
      return null;
    }
    return record;
  } catch {
    return null;
  }
}

function touchOutlookCategoryOperationRecord(
  record: OutlookCategoryOperationRecord,
  patch?: Partial<OutlookCategoryOperationRecord> & { leaseMs?: number }
): OutlookCategoryOperationRecord {
  const leaseMs = Math.max(5_000, Number(patch?.leaseMs || 0) || OUTLOOK_CATEGORY_OPERATION_DEFAULT_LEASE_MS);
  const next: OutlookCategoryOperationRecord = {
    ...record,
    ...patch,
    target: patch?.target ? normalizeOutlookCategorySyncTarget(patch.target) : record.target,
    lastUpdatedAtIso: new Date().toISOString(),
    leaseExpiresAtMs: Date.now() + leaseMs,
  };
  persistOutlookCategoryOperationRecord(next);
  return next;
}

export function beginOutlookCategoryOperation(input: {
  owner: string;
  target?: OutlookCategorySyncTarget | null;
  expectedItemToken?: string;
  operationId?: string;
  leaseMs?: number;
}): { ok: true; operation: OutlookCategoryOperationRecord } | { ok: false; reason: "locked" | "no-identity"; activeOperation?: OutlookCategoryOperationRecord | null } {
  const itemIdentity = buildOutlookCategoryOperationItemIdentity({
    target: input.target,
    expectedItemToken: input.expectedItemToken,
  });
  if (!itemIdentity) {
    logOutlookCategoryOperation("warn", "begin-failed-no-identity", {
      owner: input.owner,
      target: input.target,
    });
    return { ok: false, reason: "no-identity" };
  }

  const activeOperation = readActiveOutlookCategoryOperationByItemIdentity(itemIdentity);
  if (activeOperation) {
    logOutlookCategoryOperation("warn", "begin-failed-locked", {
      owner: input.owner,
      itemIdentity,
      activeOperationId: activeOperation.operationId,
      activePhase: activeOperation.phase,
    });
    return {
      ok: false,
      reason: "locked",
      activeOperation,
    };
  }

  const operationId = String(input.operationId || "").trim()
    || `outlook-category-op:${String(input.owner || "operation").trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-")}:${Date.now()}:${Math.random().toString(36).slice(2)}`;
  const leaseMs = Math.max(5_000, Number(input.leaseMs || 0) || OUTLOOK_CATEGORY_OPERATION_DEFAULT_LEASE_MS);
  const nowIso = new Date().toISOString();
  const record: OutlookCategoryOperationRecord = {
    operationId,
    itemIdentity,
    target: normalizeOutlookCategorySyncTarget(input.target),
    owner: String(input.owner || "").trim() || "unknown",
    startedAtIso: nowIso,
    startedAtMs: Date.now(),
    lastUpdatedAtIso: nowIso,
    phase: "opening",
    status: "active",
    expectedItemToken: String(input.expectedItemToken || "").trim() || undefined,
    leaseExpiresAtMs: Date.now() + leaseMs,
  };
  persistOutlookCategoryOperationRecord(record);
  try {
    window.localStorage?.setItem(getOutlookCategoryOperationActiveStorageKey(itemIdentity), operationId);
  } catch {
    // best effort
  }
  logOutlookCategoryOperation("info", "lock-acquired", {
    operationId,
    owner: record.owner,
    itemIdentity,
    phase: record.phase,
  });
  return { ok: true, operation: cloneOutlookCategoryOperationRecord(record)! };
}

export function getActiveOutlookCategoryOperation(target?: OutlookCategorySyncTarget | null, expectedItemToken?: string): OutlookCategoryOperationRecord | null {
  const itemIdentity = buildOutlookCategoryOperationItemIdentity({ target, expectedItemToken });
  return cloneOutlookCategoryOperationRecord(readActiveOutlookCategoryOperationByItemIdentity(itemIdentity));
}

export function setOutlookCategoryOperationPhase(
  operationId: string,
  phase: OutlookCategoryOperationPhase,
  patch?: Partial<OutlookCategoryOperationRecord> & { leaseMs?: number }
): OutlookCategoryOperationRecord | null {
  const record = readOutlookCategoryOperationRecord(operationId);
  if (!record || record.status !== "active") return null;
  const next = touchOutlookCategoryOperationRecord(record, {
    ...patch,
    phase,
  });
  logOutlookCategoryOperation("debug", "phase", {
    operationId,
    itemIdentity: next.itemIdentity,
    phase,
    requestId: next.requestId,
  });
  return cloneOutlookCategoryOperationRecord(next);
}

export function completeOutlookCategoryOperation(
  operationId: string,
  input: {
    result: OutlookCategoryWriterResult;
    requestId?: string;
    phase?: Extract<OutlookCategoryOperationPhase, "completed" | "failed" | "cancelled">;
    detail?: string;
  }
): OutlookCategoryOperationRecord | null {
  const record = readOutlookCategoryOperationRecord(operationId);
  if (!record) return null;
  const result = input.result;
  const status: OutlookCategoryOperationStatus =
    result === "success" || result === "duplicate"
      ? "completed"
      : result === "cancelled"
        ? "cancelled"
        : result === "timeout"
          ? "timeout"
          : "failed";
  const phase = input.phase
    || (status === "completed" ? "completed" : status === "cancelled" ? "cancelled" : "failed");
  const next = touchOutlookCategoryOperationRecord(record, {
    phase,
    status,
    result,
    requestId: String(input.requestId || record.requestId || "").trim() || undefined,
    leaseMs: OUTLOOK_CATEGORY_OPERATION_DEFAULT_LEASE_MS,
  });
  clearOutlookCategoryOperationActiveLock(next.itemIdentity, next.operationId);
  logOutlookCategoryOperation(status === "completed" ? "info" : "warn", "lock-released", {
    operationId,
    itemIdentity: next.itemIdentity,
    phase,
    status,
    result,
    detail: input.detail,
  });
  return cloneOutlookCategoryOperationRecord(next);
}

function publishOutlookCategorySyncResult(result: OutlookCategorySyncResult) {
  if (typeof window === "undefined" || !window.localStorage) return;
  try {
    window.localStorage.setItem(OUTLOOK_CATEGORY_SYNC_RESULT_STORAGE_KEY, JSON.stringify(result));
    window.dispatchEvent(new CustomEvent(OUTLOOK_CATEGORY_SYNC_RESULT_EVENT, { detail: result }));
  } catch {
    // best effort
  }
}

function readOutlookCategorySyncResult(requestId: string): OutlookCategorySyncResult | null {
  if (typeof window === "undefined" || !window.localStorage) return null;
  try {
    const raw = window.localStorage.getItem(OUTLOOK_CATEGORY_SYNC_RESULT_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    if (String(parsed.requestId || "").trim() !== String(requestId || "").trim()) return null;
    return parsed as OutlookCategorySyncResult;
  } catch {
    return null;
  }
}

export async function waitForOutlookCategorySyncResult(
  requestId: string,
  options?: { timeoutMs?: number }
): Promise<OutlookCategorySyncResult | null> {
  const normalizedRequestId = String(requestId || "").trim();
  if (!normalizedRequestId) return null;
  const immediate = readOutlookCategorySyncResult(normalizedRequestId);
  if (immediate) return immediate;
  const timeoutMs = Math.max(1_000, Number(options?.timeoutMs || 0) || 15_000);
  return await new Promise<OutlookCategorySyncResult | null>((resolve) => {
    let finished = false;
    const finish = (value: OutlookCategorySyncResult | null) => {
      if (finished) return;
      finished = true;
      try {
        window.removeEventListener("storage", handleStorage as EventListener);
        window.removeEventListener(OUTLOOK_CATEGORY_SYNC_RESULT_EVENT, handleEvent as EventListener);
      } catch {
        // ignore
      }
      window.clearTimeout(timerId);
      resolve(value);
    };
    const maybeResolve = (candidate: OutlookCategorySyncResult | null | undefined) => {
      if (!candidate) return;
      if (String(candidate.requestId || "").trim() !== normalizedRequestId) return;
      finish(candidate);
    };
    const handleStorage = (event: StorageEvent) => {
      if (event.key && event.key !== OUTLOOK_CATEGORY_SYNC_RESULT_STORAGE_KEY) return;
      maybeResolve(readOutlookCategorySyncResult(normalizedRequestId));
    };
    const handleEvent = (event: Event) => {
      const detail = (event as CustomEvent<OutlookCategorySyncResult>).detail;
      maybeResolve(detail || null);
    };
    const timerId = window.setTimeout(() => finish(null), timeoutMs);
    window.addEventListener("storage", handleStorage as EventListener);
    window.addEventListener(OUTLOOK_CATEGORY_SYNC_RESULT_EVENT, handleEvent as EventListener);
  });
}

function doesOutlookCategorySyncTargetMatchContext(
  target: OutlookCategorySyncTarget | null | undefined,
  context: Pick<OutlookMessageContext, "itemId" | "internetMessageId" | "conversationId">
): boolean {
  const normalizedTarget = normalizeOutlookCategorySyncTarget(target);
  if (!normalizedTarget) return true;

  const targetItemId = String(normalizedTarget.itemId || "").trim();
  const targetInternetMessageId = normalizeOutlookIdentityString(normalizedTarget.internetMessageId);
  const targetConversationId = String(normalizedTarget.conversationId || "").trim();

  const currentItemId = String(context.itemId || "").trim();
  const currentInternetMessageId = normalizeOutlookIdentityString(context.internetMessageId);
  const currentConversationId = String(context.conversationId || "").trim();

  if (targetItemId) return Boolean(currentItemId) && targetItemId === currentItemId;
  if (targetInternetMessageId) return Boolean(currentInternetMessageId) && targetInternetMessageId === currentInternetMessageId;
  if (targetConversationId) return Boolean(currentConversationId) && targetConversationId === currentConversationId;
  return true;
}

function getOutlookCategoryValueSignature(values: readonly string[]): string {
  return JSON.stringify(
    normalizeUniqueCategoryValues(values).sort((left, right) =>
      String(left || "").trim().toLowerCase().localeCompare(String(right || "").trim().toLowerCase(), "pt")
    )
  );
}

function formatOutlookAsyncError(error: any): string {
  const code = String(error?.code || "").trim();
  const message = String(error?.message || "").trim();
  if (code && message) return `${code}: ${message}`;
  return code || message || "unknown";
}

function buildOutlookCategoryPlanDiff(currentCategories: string[], plan: OutlookCategoryPlan): OutlookCategoryPlanDiff {
  const normalizedCurrentCategories = normalizeUniqueCategoryValues(currentCategories);
  const desiredCategories = normalizeUniqueCategoryValues(plan.desiredCategories);
  const currentManagedCategories = getCurrentManagedCategoryNames(normalizedCurrentCategories, plan);
  const desiredCategorySet = new Set(desiredCategories);
  const currentManagedSet = new Set(currentManagedCategories);
  return {
    currentCategories: normalizedCurrentCategories,
    currentManagedCategories,
    desiredCategories,
    toAdd: desiredCategories.filter((name) => !normalizedCurrentCategories.includes(name)),
    toRemove: Array.from(new Set(currentManagedCategories.filter((name) => !desiredCategorySet.has(name)))),
    missingManagedCategories: desiredCategories.filter((name) => !currentManagedSet.has(name)),
    unexpectedManagedCategories: currentManagedCategories.filter((name) => !desiredCategorySet.has(name)),
  };
}

function isOutlookCategoryPlanDiffSatisfied(diff: OutlookCategoryPlanDiff): boolean {
  return getOutlookCategoryValueSignature(diff.currentManagedCategories)
    === getOutlookCategoryValueSignature(diff.desiredCategories);
}

function isOutlookCategoryPlanHostConfirmed(
  diff: OutlookCategoryPlanDiff,
  readback: OutlookCategoryReadbackResult
): boolean {
  return readback.source === "getAsync" && isOutlookCategoryPlanDiffSatisfied(diff);
}

function formatOutlookCategoryPlanDiffDetail(
  diff: OutlookCategoryPlanDiff,
  errors: string[] = []
): string {
  const parts: string[] = [];
  if (diff.missingManagedCategories.length) {
    parts.push(`missing=${diff.missingManagedCategories.join(", ")}`);
  }
  if (diff.unexpectedManagedCategories.length) {
    parts.push(`unexpected=${diff.unexpectedManagedCategories.join(", ")}`);
  }
  if (errors.length) {
    parts.push(`errors=${errors.join(" | ")}`);
  }
  return parts.join("; ") || "managed categories did not converge";
}

function getOutlookCategoryApplyLogItemId(syncMeta?: {
  itemIdentity?: string;
  target?: OutlookCategorySyncTarget;
}): string {
  return String(syncMeta?.target?.itemId || syncMeta?.target?.internetMessageId || syncMeta?.itemIdentity || "").trim();
}

function logOutlookCategoryApplyDiagnostic(
  level: "log" | "warn",
  event: string,
  data?: Record<string, unknown>
) {
  const message = `[TEMP][outlook-category-apply] ${event}`;
  if (level === "warn") {
    clientLog.warn(message, data);
    return;
  }
  clientLog.log(message, data);
}

function compareOutlookCategoryWriterFreshness(
  left: OutlookCategoryWriterFreshness | null | undefined,
  right: OutlookCategoryWriterFreshness | null | undefined
): number {
  if (!left && !right) return 0;
  if (!left) return -1;
  if (!right) return 1;
  if (left.requestedAtMs !== right.requestedAtMs) return left.requestedAtMs - right.requestedAtMs;
  return left.order - right.order;
}

function getOutlookCategoryWriterState(itemIdentity: string): OutlookCategoryWriterState {
  const existing = outlookCategoryWriterStateByItem.get(itemIdentity);
  if (existing) return existing;
  const created: OutlookCategoryWriterState = {
    tail: Promise.resolve({
      requestId: "__bootstrap__",
      reason: "bootstrap",
      mode: "source" as OutlookCategorySyncMode,
      itemIdentity,
      result: "success" as OutlookCategoryWriterResult,
      finishedAtIso: new Date().toISOString(),
    }),
    latestFreshness: null,
    lastAppliedPlanSignature: "",
    lastAppliedSourceSignature: "",
    recentRequestIds: [],
  };
  outlookCategoryWriterStateByItem.set(itemIdentity, created);
  return created;
}

function rememberOutlookCategoryWriterRequestId(state: OutlookCategoryWriterState, requestId: string) {
  const normalized = String(requestId || "").trim();
  if (!normalized) return;
  state.recentRequestIds = [normalized, ...state.recentRequestIds.filter((entry) => entry !== normalized)]
    .slice(0, OUTLOOK_CATEGORY_WRITER_RECENT_REQUEST_LIMIT);
}

function hasSeenOutlookCategoryWriterRequestId(state: OutlookCategoryWriterState, requestId: string): boolean {
  const normalized = String(requestId || "").trim();
  return Boolean(normalized) && state.recentRequestIds.includes(normalized);
}

function dispatchOutlookCategoryContextInvalidated() {
  if (typeof window === "undefined" || typeof window.dispatchEvent !== "function") return;
  try {
    window.dispatchEvent(new CustomEvent(OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT));
  } catch {
    // best effort only
  }
}

export function readPendingOutlookCategorySyncRequest(): OutlookCategorySyncRequest | null {
  if (typeof window === "undefined" || !window.localStorage) return null;
  try {
    const raw = window.localStorage.getItem(OUTLOOK_CATEGORY_SYNC_REQUEST_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== "object") return null;
    const requestId = String(parsed.requestId || "").trim();
    const mode = parsed.mode === "source" ? "source" : parsed.mode === "current-item-context" ? "current-item-context" : "";
    if (!requestId || !mode) return null;
    return {
      requestId,
      createdAtIso: String(parsed.createdAtIso || "").trim() || new Date().toISOString(),
      reason: String(parsed.reason || "").trim() || undefined,
      operationId: String(parsed.operationId || "").trim() || undefined,
      mode,
      target: parsed.target && typeof parsed.target === "object"
        ? {
            itemId: String(parsed.target.itemId || "").trim() || undefined,
            internetMessageId: String(parsed.target.internetMessageId || "").trim() || undefined,
            conversationId: String(parsed.target.conversationId || "").trim() || undefined,
          }
        : undefined,
      source: parsed.source && typeof parsed.source === "object"
        ? parsed.source as Partial<OutlookCategorySource>
        : undefined,
    };
  } catch {
    return null;
  }
}

export function enqueueOutlookCategorySyncRequest(input: {
  requestId?: string;
  createdAtIso?: string;
  reason?: string;
  operationId?: string;
  mode: "source" | "current-item-context";
  target?: { itemId?: string; internetMessageId?: string; conversationId?: string };
  source?: Partial<OutlookCategorySource> | null;
}): OutlookCategorySyncRequest | null {
  if (typeof window === "undefined" || !window.localStorage) return null;
  try {
    const requestId = String(input.requestId || "").trim() || createOutlookCategorySyncRequestId(input.reason || input.mode);
    const createdAtIso = String(input.createdAtIso || "").trim() || new Date().toISOString();
    const request: OutlookCategorySyncRequest = {
      requestId,
      createdAtIso,
      reason: String(input.reason || "").trim() || undefined,
      operationId: String(input.operationId || "").trim() || undefined,
      mode: input.mode,
      target: input.target ? {
        itemId: String(input.target.itemId || "").trim() || undefined,
        internetMessageId: String(input.target.internetMessageId || "").trim() || undefined,
        conversationId: String(input.target.conversationId || "").trim() || undefined,
      } : undefined,
      source: input.mode === "source" ? (input.source || {}) : undefined,
    };
    window.localStorage.setItem(OUTLOOK_CATEGORY_SYNC_REQUEST_STORAGE_KEY, JSON.stringify(request));
    try {
      window.dispatchEvent(new CustomEvent(OUTLOOK_CATEGORY_SYNC_REQUEST_EVENT, { detail: request }));
    } catch {
      // best effort only
    }
    return request;
  } catch {
    return null;
  }
}

function collectKnownOutlookCategoryLabelNames(input: {
  settings: Awaited<ReturnType<typeof getSettings>> | null;
  email: any;
  groups: any[];
  tickets: any[];
}): string[] {
  return Array.from(new Set([
    ...(Array.isArray(input.settings?.groups?.labels?.catalog)
      ? input.settings.groups.labels.catalog.map((entry) => String(entry?.label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.labels)
      ? input.email.labels.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.removedInheritedLabels)
      ? input.email.removedInheritedLabels.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.classificationMeta?.categorizedLabelNames)
      ? input.email.classificationMeta.categorizedLabelNames.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.groups)
      ? input.groups.flatMap((group: any) => Array.isArray(group?.labels) ? group.labels : []).map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.tickets)
      ? input.tickets.flatMap((ticket: any) => Array.isArray(ticket?.labels) ? ticket.labels : []).map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
  ]));
}

async function waitForOffice(maxWaitMs = 5000): Promise<any> {
  const start = Date.now();
  while (true) {
    const OfficeAny = (window as any).Office;
    if (OfficeAny) return OfficeAny;
    if (Date.now() - start > maxWaitMs) return null;
    await sleep(50);
  }
}

async function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: any;
  const timeoutPromise = new Promise<T>((resolve) => {
    timer = setTimeout(() => resolve(fallback), ms);
  });
  return Promise.race([
    promise.then(v => { clearTimeout(timer); return v; }),
    timeoutPromise
  ]);
}

async function ensureOfficeReady(): Promise<any> {
  const OfficeAny = await waitForOffice(5000);
  if (!OfficeAny) throw new Error("Office.js não está disponível (Script não carregou).");

  // Bypass: If legacy initialize already fired and we have context, don't wait for onReady (it might hang)
  if (OfficeAny.context) {
    return OfficeAny;
  }

  clientLog.log("[office] waiting for onReady...");
  // Modern Office.onReady returns a promise. We wait up to 10s for the handshake.
  const ready = await withTimeout(OfficeAny.onReady(), 10000, null);

  if (!ready) {
    // Final check before throwing: maybe context appeared meanwhile?
    if (OfficeAny.context) return OfficeAny;
    clientLog.error("[office] onReady timeout (10s)");
    throw new Error("O Office.js demorou demasiado tempo a carregar (Handshake timeout). Reabre o Cockpit.");
  }

  clientLog.log("[office] onReady finished");
  return OfficeAny;
}

async function getNestableMsalInstance(): Promise<any | null> {
  const OfficeAny = await ensureOfficeReady();
  const supportsNaa = Boolean(OfficeAny?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1"));
  if (!supportsNaa) return null;
  if (!nestableMsalPromise) {
    nestableMsalPromise = (async () => {
      const { createNestablePublicClientApplication } = await import("@azure/msal-browser");
      return await createNestablePublicClientApplication({
        auth: {
          clientId: GRAPH_NAA_CLIENT_ID,
          authority: GRAPH_NAA_AUTHORITY,
          redirectUri: GRAPH_NAA_REDIRECT_URI,
        },
        cache: {
          cacheLocation: "localStorage",
        },
      });
    })().catch((error) => {
      nestableMsalPromise = null;
      throw error;
    });
  }
  return await nestableMsalPromise;
}

async function acquireGraphTokenWithNaa(scopes: string[], options?: { allowPrompts?: boolean }): Promise<string> {
  try {
    const msal = await getNestableMsalInstance();
    if (!msal) return "";
    const { InteractionRequiredAuthError } = await import("@azure/msal-browser");
    const tokenRequest = { scopes };
    try {
      const authResult = await msal.acquireTokenSilent(tokenRequest);
      return String(authResult?.accessToken || "").trim();
    } catch (error: any) {
      const interactionRequired =
        error instanceof InteractionRequiredAuthError
        || String(error?.name || "").trim() === "InteractionRequiredAuthError"
        || /interaction_required|consent_required|login_required/i.test(String(error?.errorCode || error?.message || ""));
      if (!interactionRequired) throw error;
      if (options?.allowPrompts === false) return "";
      const authResult = await msal.acquireTokenPopup(tokenRequest);
      return String(authResult?.accessToken || "").trim();
    }
  } catch (error) {
    clientLog.warn("[office] NAA Graph token acquisition failed", error);
    return "";
  }
}

async function getGraphAccessToken(scopes: string[], options?: { allowPrompts?: boolean }): Promise<string> {
  if (!GRAPH_RUNTIME_ENABLED) return "";
  return await acquireGraphTokenWithNaa(scopes, options);
}

function summarizeGraphError(error: any): { errorCode?: string; errorMessage: string } {
  const errorCode = String(error?.errorCode || error?.code || error?.name || "").trim() || undefined;
  const errorMessage = String(
    error?.message
    || error?.errorMessage
    || error?.error_description
    || error?.toString?.()
    || "Erro desconhecido"
  ).trim();
  return { errorCode, errorMessage };
}

function isAdminApprovalGraphError(error: any): boolean {
  const raw = [
    error?.errorCode,
    error?.subError,
    error?.message,
    error?.errorMessage,
    error?.error_description,
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(" ");
  return /admin approval|administrator approval|aadsts65001|aadsts65004|consent_required|permission(s)? requested require admin/i.test(raw);
}

async function getBrowserMsalInstance(): Promise<any> {
  if (!browserMsalPromise) {
    browserMsalPromise = (async () => {
      const { PublicClientApplication } = await import("@azure/msal-browser");
      const app = new PublicClientApplication({
        auth: {
          clientId: GRAPH_NAA_CLIENT_ID,
          authority: GRAPH_NAA_AUTHORITY,
          redirectUri: GRAPH_NAA_REDIRECT_URI,
        },
        cache: {
          cacheLocation: "localStorage",
        },
      });
      await app.initialize();
      return app;
    })().catch((error) => {
      browserMsalPromise = null;
      throw error;
    });
  }
  return await browserMsalPromise;
}

type GraphSelfTestAuthOutcome = {
  accessToken: string;
  authMode: GraphDriveSelfTestResult["authMode"];
  consent: GraphDriveSelfTestResult["consent"];
};

async function acquireGraphTokenForSelfTest(scopes: string[]): Promise<GraphSelfTestAuthOutcome> {
  try {
    const msal = await getNestableMsalInstance();
    if (msal) {
      const { InteractionRequiredAuthError } = await import("@azure/msal-browser");
      const tokenRequest = { scopes: [...scopes] };
      try {
        const silentResult = await msal.acquireTokenSilent(tokenRequest);
        const accessToken = String(silentResult?.accessToken || "").trim();
        return {
          accessToken,
          authMode: "nested_app_auth",
          consent: {
            attempted: true,
            ok: Boolean(accessToken),
            result: accessToken ? "accepted" : "auth_error",
            detail: "Token obtido por Nested App Auth (silent).",
            account: String(silentResult?.account?.username || "").trim() || undefined,
          },
        };
      } catch (error: any) {
        const interactionRequired =
          error instanceof InteractionRequiredAuthError
          || String(error?.name || "").trim() === "InteractionRequiredAuthError"
          || /interaction_required|consent_required|login_required/i.test(String(error?.errorCode || error?.message || ""));

        if (!interactionRequired) {
          const summary = summarizeGraphError(error);
          return {
            accessToken: "",
            authMode: "nested_app_auth",
            consent: {
              attempted: true,
              ok: false,
              result: isAdminApprovalGraphError(error) ? "need_admin_approval" : "auth_error",
              errorCode: summary.errorCode,
              errorMessage: summary.errorMessage,
            },
          };
        }

        const popupResult = await withTimeout(
          msal.acquireTokenPopup(tokenRequest),
          45_000,
          null as any,
        );
        if (!popupResult) {
          return {
            accessToken: "",
            authMode: "nested_app_auth",
            consent: {
              attempted: true,
              ok: false,
              result: "timeout",
              errorMessage: "Timeout a aguardar conclusao do popup de consentimento/login (Nested App Auth).",
            },
          };
        }
        const accessToken = String(popupResult?.accessToken || "").trim();
        return {
          accessToken,
          authMode: "nested_app_auth",
          consent: {
            attempted: true,
            ok: Boolean(accessToken),
            result: accessToken ? "accepted" : "auth_error",
            detail: "Token obtido por Nested App Auth (popup).",
            account: String(popupResult?.account?.username || "").trim() || undefined,
          },
        };
      }
    }
  } catch (error) {
    const summary = summarizeGraphError(error);
    if (String(summary.errorMessage || "").trim()) {
      return {
        accessToken: "",
        authMode: "nested_app_auth",
        consent: {
          attempted: true,
          ok: false,
          result: isAdminApprovalGraphError(error) ? "need_admin_approval" : "auth_error",
          errorCode: summary.errorCode,
          errorMessage: summary.errorMessage,
        },
      };
    }
  }

  try {
    const msal = await getBrowserMsalInstance();
    const existingAccounts = Array.isArray(msal.getAllAccounts?.()) ? msal.getAllAccounts() : [];
    const knownAccount = existingAccounts[0] || null;
    const tokenRequest = {
      scopes: [...scopes],
      account: knownAccount || undefined,
    };

    if (knownAccount) {
      try {
        const silentResult = await msal.acquireTokenSilent(tokenRequest);
        const accessToken = String(silentResult?.accessToken || "").trim();
        if (silentResult?.account) {
          try {
            msal.setActiveAccount?.(silentResult.account);
          } catch {
            // best effort
          }
        }
        return {
          accessToken,
          authMode: "browser_msal",
          consent: {
            attempted: true,
            ok: Boolean(accessToken),
            result: accessToken ? "accepted" : "auth_error",
            detail: "Token obtido por MSAL browser (silent).",
            account: String(silentResult?.account?.username || "").trim() || undefined,
          },
        };
      } catch {
        // continue to popup flow
      }
    }

    const popupResult = await withTimeout(
      knownAccount
        ? msal.acquireTokenPopup({ scopes: [...scopes], account: knownAccount })
        : msal.loginPopup({ scopes: [...scopes] }),
      45_000,
      null as any,
    );

    if (!popupResult) {
      return {
        accessToken: "",
        authMode: "browser_msal",
        consent: {
          attempted: true,
          ok: false,
          result: "timeout",
          errorMessage: "Timeout a aguardar conclusao do popup de consentimento/login (MSAL browser).",
        },
      };
    }

    try {
      if (popupResult?.account) {
        msal.setActiveAccount?.(popupResult.account);
      }
    } catch {
      // best effort
    }

    let accessToken = String(popupResult?.accessToken || "").trim();
    if (!accessToken) {
      const popupAccount = popupResult?.account || msal.getActiveAccount?.() || msal.getAllAccounts?.()[0];
      if (popupAccount) {
        const silentResult = await msal.acquireTokenSilent({
          scopes: [...scopes],
          account: popupAccount,
        });
        accessToken = String(silentResult?.accessToken || "").trim();
      }
    }

    return {
      accessToken,
      authMode: "browser_msal",
      consent: {
        attempted: true,
        ok: Boolean(accessToken),
        result: accessToken ? "accepted" : "auth_error",
        detail: "Token obtido por MSAL browser (popup).",
        account: String(popupResult?.account?.username || "").trim() || undefined,
      },
    };
  } catch (error) {
    const summary = summarizeGraphError(error);
    return {
      accessToken: "",
      authMode: "browser_msal",
      consent: {
        attempted: true,
        ok: false,
        result: isAdminApprovalGraphError(error) ? "need_admin_approval" : "auth_error",
        errorCode: summary.errorCode,
        errorMessage: summary.errorMessage,
      },
    };
  }
}

async function readGraphJsonResponse(response: Response): Promise<GraphDriveSelfTestStep> {
  const rawText = await response.text();
  let parsed: unknown = rawText;
  try {
    parsed = rawText ? JSON.parse(rawText) : null;
  } catch {
    parsed = rawText;
  }
  const graphError = (parsed as any)?.error;
  return {
    attempted: true,
    ok: response.ok,
    status: response.status,
    response: parsed,
    errorCode: response.ok ? undefined : String(graphError?.code || "").trim() || undefined,
    errorMessage: response.ok
      ? undefined
      : String(graphError?.message || response.statusText || "Falha Graph").trim() || undefined,
  };
}

export async function runGraphDriveWriteSelfTest(folderName = "InboxCockpit-Graph-Write-Test"): Promise<GraphDriveSelfTestResult> {
  const scopes = [...GRAPH_DRIVE_SELF_TEST_SCOPES];
  const auth = await acquireGraphTokenForSelfTest(scopes);

  if (!auth.accessToken) {
    return {
      scopes,
      authMode: auth.authMode,
      consent: auth.consent,
      meDrive: { attempted: false, ok: false, detail: "Saltado por falta de token." },
      createFolder: { attempted: false, ok: false, detail: "Saltado por falta de token." },
      cleanup: { attempted: false, ok: false, detail: "Saltado por falta de pasta criada." },
      conclusion: auth.consent.result === "need_admin_approval"
        ? "tenant_blocks_user_write"
        : "implementation_cannot_complete_test",
      conclusionMessage: auth.consent.result === "need_admin_approval"
        ? "O tenant bloqueou o consentimento delegado necessario para o write test."
        : String(auth.consent.errorMessage || "A implementacao atual nao conseguiu fechar o fluxo de token Graph."),
    };
  }

  const headers = {
    Authorization: `Bearer ${auth.accessToken}`,
    "Content-Type": "application/json",
  };

  const meDriveResponse = await fetch("https://graph.microsoft.com/v1.0/me/drive", {
    headers: { Authorization: `Bearer ${auth.accessToken}` },
  });
  const meDrive = await readGraphJsonResponse(meDriveResponse);
  if (!meDrive.ok) {
    const blocked = Number(meDrive.status || 0) === 403;
    return {
      scopes,
      authMode: auth.authMode,
      consent: auth.consent,
      meDrive,
      createFolder: { attempted: false, ok: false, detail: "Saltado porque /me/drive falhou." },
      cleanup: { attempted: false, ok: false, detail: "Saltado por falta de pasta criada." },
      conclusion: blocked ? "tenant_blocks_user_write" : "implementation_cannot_complete_test",
      conclusionMessage: blocked
        ? String(meDrive.errorMessage || "O tenant recusou o acesso ao drive do utilizador.")
        : String(meDrive.errorMessage || "A chamada /me/drive falhou por motivo tecnico."),
    };
  }

  const createFolderResponse = await fetch("https://graph.microsoft.com/v1.0/me/drive/root/children", {
    method: "POST",
    headers,
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    }),
  });
  const createFolderStepBase = await readGraphJsonResponse(createFolderResponse);
  const createFolder: GraphDriveSelfTestResult["createFolder"] = {
    ...createFolderStepBase,
    folderId: String((createFolderStepBase.response as any)?.id || "").trim() || undefined,
    folderName: String((createFolderStepBase.response as any)?.name || "").trim() || undefined,
  };

  let cleanup: GraphDriveSelfTestResult["cleanup"] = {
    attempted: false,
    ok: false,
    detail: "Saltado por falta de pasta criada.",
  };

  if (createFolder.ok && createFolder.folderId) {
    const cleanupResponse = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(createFolder.folderId)}`,
      {
        method: "DELETE",
        headers: { Authorization: `Bearer ${auth.accessToken}` },
      }
    );
    cleanup = {
      attempted: true,
      ok: cleanupResponse.ok || cleanupResponse.status === 204,
      status: cleanupResponse.status,
      detail: cleanupResponse.ok || cleanupResponse.status === 204
        ? "Pasta de teste removida com sucesso."
        : "A limpeza da pasta de teste falhou.",
    };
    if (!cleanup.ok) {
      const cleanupBody = await cleanupResponse.text();
      cleanup.errorMessage = String(cleanupBody || cleanupResponse.statusText || "Falha a limpar pasta de teste.").trim();
    }
  }

  if (!createFolder.ok) {
    const blocked = Number(createFolder.status || 0) === 403;
    return {
      scopes,
      authMode: auth.authMode,
      consent: auth.consent,
      meDrive,
      createFolder,
      cleanup,
      conclusion: blocked ? "tenant_blocks_user_write" : "implementation_cannot_complete_test",
      conclusionMessage: blocked
        ? String(createFolder.errorMessage || "O tenant nao permitiu criar a pasta de teste no OneDrive do utilizador.")
        : String(createFolder.errorMessage || "A criacao da pasta de teste falhou por motivo tecnico."),
    };
  }

  return {
    scopes,
    authMode: auth.authMode,
    consent: auth.consent,
    meDrive,
    createFolder,
    cleanup,
    conclusion: "tenant_allows_user_write",
    conclusionMessage: "Consentimento, /me/drive e criacao da pasta de teste funcionaram com o utilizador autenticado.",
  };
}

async function getCurrentMessageRestId(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const mailbox = OfficeAny?.context?.mailbox;
    const item = mailbox?.item;
    const itemId = String(item?.itemId || "").trim();
    if (!itemId) return "";
    try {
      if (typeof mailbox?.convertToRestId === "function" && OfficeAny?.MailboxEnums?.RestVersion?.v2_0) {
        return String(mailbox.convertToRestId(itemId, OfficeAny.MailboxEnums.RestVersion.v2_0) || "").trim() || itemId;
      }
    } catch (error) {
      clientLog.warn("[office] convertToRestId failed", error);
    }
    return itemId;
  } catch (error) {
    clientLog.warn("[office] getCurrentMessageRestId failed", error);
    return "";
  }
}

function mergeOutlookAttachments(primary: OutlookAttachment[], fallback: OutlookAttachment[]): OutlookAttachment[] {
  const byKey = new Map<string, OutlookAttachment>();
  const makeKey = (attachment: Partial<OutlookAttachment>) =>
    String(attachment?.id || attachment?.contentId || attachment?.name || "").trim().toLowerCase();
  for (const attachment of fallback) {
    const key = makeKey(attachment);
    if (!key) continue;
    byKey.set(key, { ...attachment });
  }
  for (const attachment of primary) {
    const key = makeKey(attachment);
    if (!key) continue;
    const existing = byKey.get(key);
    byKey.set(key, {
      ...(existing || {}),
      ...attachment,
      content: String(attachment.content || existing?.content || "").trim(),
    });
  }
  return Array.from(byKey.values()).filter((attachment) => String(attachment.name || "").trim());
}

async function getAttachmentsViaGraphForCurrentItem(): Promise<OutlookAttachment[]> {
  const messageId = await getCurrentMessageRestId();
  if (!messageId) return [];
  const token = await getGraphAccessToken(GRAPH_ATTACHMENT_SCOPES, { allowPrompts: true });
  if (!token) return [];
  const headers = { Authorization: `Bearer ${token}` };
  try {
    const listRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline,contentId`,
      { headers }
    );
    if (!listRes.ok) {
      clientLog.warn("[office] Graph attachment list failed", { status: listRes.status });
      return [];
    }
    const listBody: any = await listRes.json();
    const rawItems = Array.isArray(listBody?.value) ? listBody.value : [];
    const results: OutlookAttachment[] = [];
    for (const summary of rawItems) {
      const attachmentId = String(summary?.id || "").trim();
      const attachmentName = String(summary?.name || "").trim();
      if (!attachmentId || !attachmentName) continue;
      try {
        const detailRes = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
          { headers }
        );
        if (!detailRes.ok) {
          clientLog.warn("[office] Graph attachment detail failed", { attachmentId, status: detailRes.status });
          continue;
        }
        const detail: any = await detailRes.json();
        results.push({
          id: attachmentId || undefined,
          name: attachmentName,
          contentType: String(detail?.contentType || summary?.contentType || "application/octet-stream").trim(),
          size: Number(detail?.size || summary?.size || 0) || undefined,
          isInline: Boolean(detail?.isInline ?? summary?.isInline),
          contentId: String(detail?.contentId || summary?.contentId || "").trim() || undefined,
          content: String(detail?.contentBytes || "").trim(),
        });
      } catch (error) {
        clientLog.warn("[office] Graph attachment detail exception", { attachmentId, error });
      }
    }
    return results.filter((attachment) => attachment.name);
  } catch (error) {
    clientLog.warn("[office] Graph attachment fallback failed", error);
    return [];
  }
}

function normalizeRecipients(arr: any): Recipient[] {
  if (!Array.isArray(arr)) return [];
  return arr
    .map((r) => ({
      name: String(r?.displayName || "").trim(),
      email: String(r?.emailAddress || "").trim(),
    }))
    .filter((r) => r.email);
}

export async function getOutlookContext(): Promise<OutlookMessageContext> {
  try {
    const OfficeAny = await ensureOfficeReady();

    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) {
      clientLog.warn("[office] mailbox.item is empty");
      return {};
    }

    const getAsyncValue = async (obj: any, coercer: (v: any) => string): Promise<string> => {
      if (!obj?.getAsync) return "";
      const p = new Promise<string>((resolve) => {
        try {
          obj.getAsync((r: any) => {
            try {
              if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(coercer(r.value));
              else resolve("");
            } catch { resolve(""); }
          });
        } catch { resolve(""); }
      });
      return await withTimeout(p, 2000, "");
    };

    const getSubject = async (): Promise<string> => {
      const s = item.subject;
      if (typeof s === "string") return s;
      // Compose: subject is an object with getAsync/setAsync
      return await getAsyncValue(s, (v) => String(v ?? ""));
    };

    const getRecipients = async (recips: any): Promise<Recipient[]> => {
      if (Array.isArray(recips)) return normalizeRecipients(recips);
      // Compose: recipients are an object with getAsync/addAsync
      if (recips?.getAsync) {
        const raw = await new Promise<any[]>((resolve) => {
          try {
            recips.getAsync((r: any) => {
              try {
                if (r?.status === OfficeAny.AsyncResultStatus.Succeeded && Array.isArray(r.value)) resolve(r.value);
                else resolve([]);
              } catch {
                resolve([]);
              }
            });
          } catch {
            resolve([]);
          }
        });
        return normalizeRecipients(raw);
      }
      return [];
    };

    const subject = await getSubject();

    // From is only reliable in Read. In Compose it may be missing/unsupported.
    const from = item.from;
    const fromEmail = from?.emailAddress ? String(from.emailAddress) : "";
    const fromName = from?.displayName ? String(from.displayName) : "";

    const conversationId = typeof item.conversationId === "string" ? item.conversationId : "";
    const internetMessageId = typeof item.internetMessageId === "string" ? item.internetMessageId : "";

    const itemId = typeof item.itemId === "string" ? item.itemId : "";

    const receivedDateTimeIso = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : "";

    const toRecipients = await getRecipients(item.to);
    const ccRecipients = await getRecipients(item.cc);

    let isCompose = false;
    try {
      isCompose = await isComposeMode();
    } catch {
      isCompose = false;
    }

    return {
      subject,
      fromEmail,
      fromName,
      conversationId,
      internetMessageId,
      itemId,
      receivedDateTimeIso,
      toRecipients,
      ccRecipients,
      isCompose,
    };
  } catch (error) {
    clientLog.error("[office] getOutlookContext error", error);
    return {};
  }
}


// Backwards-compat with older UI code
export const getSelectedMessageContext = getOutlookContext;

const ODOO_LINKED_NOTICE = "iccc-odoo-linked";

function firstCategoryColor(colors: any, candidates: string[]): any {
  for (const candidate of candidates) {
    if (colors?.[candidate]) return colors[candidate];
  }
  return colors?.Preset0;
}

function hashCategorySeed(value: string): number {
  let hash = 0;
  for (const char of String(value || "")) {
    hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
  }
  return Math.abs(hash);
}

function extractTicketSeriesKey(ticketCode: string): string {
  const raw = String(ticketCode || "").trim().replace(/^Ticket:\s*/i, "").replace(/^TK:\s*/i, "");
  return String(raw.split(/[-/_\s]/)[0] || "").trim().toUpperCase();
}

function isReservedManagedCategoryName(name: string): boolean {
  return isReservedOutlookCategoryName(name);
}

function resolveStatusCategoryColor(label: string, colors: any): any {
  const normalized = String(label || "").trim().toLowerCase();
  if (normalized === "em analise") return firstCategoryColor(colors, ["Preset3", "Preset1", "Preset0"]);
  if (normalized === "em progresso") return firstCategoryColor(colors, ["Preset1", "Preset5", "Preset0"]);
  if (normalized === "concluido") return firstCategoryColor(colors, ["Preset4", "Preset14", "Preset0"]);
  if (normalized === "aberto") return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  if (normalized === "fechado") return firstCategoryColor(colors, ["Preset4", "Preset14", "Preset0"]);
  return firstCategoryColor(colors, ["Preset12", "Preset0"]);
}

function resolveManagedCategoryColor(displayName: string, colors: any, preferredStatus?: string): any {
  const label = String(displayName || "").trim();
  if (!label || !colors) return colors?.Preset0;

  if (preferredStatus) {
    return resolveStatusCategoryColor(preferredStatus, colors);
  }

  if (label === ODOO_LINKED_CATEGORY) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(GROUP_CATEGORY_PREFIX)) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(REFERENCE_CATEGORY_PREFIX)) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(LEGACY_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(LEGACY_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(GROUP_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(GROUP_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(TICKET_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(TICKET_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(LABEL_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(LABEL_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(TICKET_CATEGORY_PREFIX)) {
    const seriesKey = extractTicketSeriesKey(label.slice(TICKET_CATEGORY_PREFIX.length));
    if (/^(RCL|REC|CLA|RET|RMA)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset6", "Preset9", "Preset0"]);
    }
    if (/^(EDD|ENC|ORD|PED|PO)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
    }
    if (/^(SUP|TCK|INC|SRV|TEC)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset5", "Preset8", "Preset0"]);
    }
    const palette = ["Preset22", "Preset5", "Preset8", "Preset4", "Preset1", "Preset6", "Preset9", "Preset14"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seriesKey) % palette.length], "Preset0"]);
  }

  if (label.startsWith(LEGACY_TICKET_CATEGORY_PREFIX)) {
    const seriesKey = extractTicketSeriesKey(label.slice(LEGACY_TICKET_CATEGORY_PREFIX.length));
    if (/^(RCL|REC|CLA|RET|RMA)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset6", "Preset9", "Preset0"]);
    }
    if (/^(EDD|ENC|ORD|PED|PO)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
    }
    if (/^(SUP|TCK|INC|SRV|TEC)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset5", "Preset8", "Preset0"]);
    }
    const palette = ["Preset22", "Preset5", "Preset8", "Preset4", "Preset1", "Preset6", "Preset9", "Preset14"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seriesKey) % palette.length], "Preset0"]);
  }

  if (label.startsWith(LEGACY_LABEL_CATEGORY_PREFIX)) {
    const seed = label.slice(LEGACY_LABEL_CATEGORY_PREFIX.length).trim();
    const palette = ["Preset12", "Preset10", "Preset11", "Preset13", "Preset15", "Preset16", "Preset17", "Preset18"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seed) % palette.length], "Preset0"]);
  }

  if (!isReservedManagedCategoryName(label)) {
    const palette = ["Preset12", "Preset10", "Preset11", "Preset13", "Preset15", "Preset16", "Preset17", "Preset18"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(label) % palette.length], "Preset0"]);
  }

  return colors?.Preset0;
}


export type OutlookContactSuggestion = {
  name?: string;
  company?: string;
  jobTitle?: string;
  phones?: string[];
  email?: string;
};

export async function getOutlookContactSuggestionByEmail(emailRaw: string): Promise<OutlookContactSuggestion | null> {
  const email = String(emailRaw || "").trim().toLowerCase();
  if (!email) return null;
  if (!GRAPH_RUNTIME_ENABLED) return null;

  try {
    const token = await getGraphAccessToken(GRAPH_PEOPLE_SCOPES, { allowPrompts: false });
    if (!token) return null;

    const q = encodeURIComponent(`"${email}"`);
    const url = `https://graph.microsoft.com/v1.0/me/people?$search=${q}&$top=10`;
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual",
      },
    });
    if (!res.ok) return null;

    const body: any = await res.json();
    const arr = Array.isArray(body?.value) ? body.value : [];

    const exact = arr.find((p: any) => {
      const emails = Array.isArray(p?.scoredEmailAddresses) ? p.scoredEmailAddresses.map((x: any) => String(x?.address || "").trim().toLowerCase()) : [];
      return emails.includes(email);
    });

    if (!exact) return null;

    const phones = [
      ...(Array.isArray(exact?.businessPhones) ? exact.businessPhones : []),
      exact?.mobilePhone,
    ].map((x: any) => String(x || "").trim()).filter(Boolean);

    return {
      name: String(exact?.displayName || "").trim() || undefined,
      company: String(exact?.companyName || "").trim() || undefined,
      jobTitle: String(exact?.jobTitle || "").trim() || undefined,
      phones,
      email,
    };
  } catch {
    return null;
  }
}

// Ler corpo do email (texto simples) — usado pela IA
export async function getEmailBodyText(): Promise<string> {
  clientLog.log("[office] getEmailBodyText start");
  try {
    const OfficeAny = await ensureOfficeReady();
    const item: any = OfficeAny?.context?.mailbox?.item;
    if (!item?.body?.getAsync) return "";

    const p = new Promise<string>((resolve) => {
      item.body.getAsync("text", (r: any) => {
        try {
          if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(String(r.value ?? ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    });

    const result = await withTimeout(p, 3000, "");
    clientLog.log("[office] getEmailBodyText end");
    return result;
  } catch (e) {
    clientLog.error("[office] getEmailBodyText error", e);
    return "";
  }
}

export async function getEmailBodyHtml(): Promise<string> {
  clientLog.log("[office] getEmailBodyHtml start");
  try {
    const OfficeAny = await ensureOfficeReady();
    const item: any = OfficeAny?.context?.mailbox?.item;
    if (!item?.body?.getAsync) return "";

    const p = new Promise<string>((resolve) => {
      item.body.getAsync("html", (r: any) => {
        try {
          if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(String(r.value ?? ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    });

    const result = await withTimeout(p, 4000, "");
    clientLog.log("[office] getEmailBodyHtml end");
    return result;
  } catch (e) {
    clientLog.error("[office] getEmailBodyHtml error", e);
    return "";
  }
}


// Token barato para detetar mudanca de email (para polling fallback ao ItemChanged)
export async function getCurrentItemToken(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) return "";

    // "Context Poke": Read a basic property to force some hosts (Outlook Desktop) 
    // to refresh the internal state of the proxy object.
    void item.itemId;

    const cid = typeof item.conversationId === "string" ? item.conversationId : "";
    const imid = typeof item.internetMessageId === "string" ? item.internetMessageId : "";
    const itemId = typeof item.itemId === "string" ? item.itemId : "";
    const created = item.dateTimeCreated ? String(item.dateTimeCreated) : "";
    const subj = typeof item.subject === "string" ? item.subject : "";

    // Using a more structured token for better comparison
    return [cid, imid, itemId, created, subj].filter(Boolean).join("|");
  } catch {
    return "";
  }
}

async function ensureMasterCategory(displayName: string, preferredStatus?: string): Promise<OutlookCategoryMutationResult> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.masterCategories) {
    return { ok: true };
  }
  const categoryColor = resolveManagedCategoryColor(displayName, OfficeAny.MailboxEnums?.CategoryColor, preferredStatus);

  return await new Promise<OutlookCategoryMutationResult>((resolve) => {
    try {
      OfficeAny.context.mailbox.masterCategories.getAsync((res: any) => {
        if (res.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          const error = formatOutlookAsyncError(res?.error);
          clientLog.warn("[office] masterCategories.getAsync failed", { displayName, error });
          return resolve({ ok: false, error: `masterCategories.getAsync failed for ${displayName}: ${error}` });
        }
        const list = Array.isArray(res.value) ? res.value : [];
        const existing = list.find((c: any) => (c.displayName || c.name) === displayName);
        const addCategory = () =>
          OfficeAny.context.mailbox.masterCategories.addAsync([{ displayName, color: categoryColor }], (addResult: any) => {
            if (addResult?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              const error = formatOutlookAsyncError(addResult?.error);
              clientLog.warn("[office] masterCategories.addAsync failed", {
                displayName,
                error,
              });
              return resolve({ ok: false, error: `masterCategories.addAsync failed for ${displayName}: ${error}` });
            }
            resolve({ ok: true });
          });

        if (existing) {
          if (!categoryColor || String(existing?.color || "") === String(categoryColor)) {
            return resolve({ ok: true });
          }
          if (typeof OfficeAny.context.mailbox.masterCategories.removeAsync !== "function") {
            return resolve({ ok: true });
          }
          return OfficeAny.context.mailbox.masterCategories.removeAsync([displayName], (removeResult: any) => {
            if (removeResult?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              const error = formatOutlookAsyncError(removeResult?.error);
              clientLog.warn("[office] masterCategories.removeAsync failed", {
                displayName,
                error,
              });
              return resolve({ ok: false, error: `masterCategories.removeAsync failed for ${displayName}: ${error}` });
            }
            addCategory();
          });
        }

        addCategory();
      });
    } catch {
      resolve({ ok: false, error: `master category preparation threw for ${displayName}` });
    }
  });
}

async function addCategoryToCurrentItem(displayName: string): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.addAsync([displayName], () => resolve());
    } catch {
      resolve();
    }
  });
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...slice);
  }
  return globalThis.btoa(binary);
}

function uint8ArraysToBase64(chunks: Uint8Array[]): string {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return arrayBufferToBase64(merged.buffer);
}

async function getCurrentMessageAsEmlBase64(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item?.getAsFileAsync) return "";

    const file: any = await new Promise((resolve, reject) => {
      try {
        item.getAsFileAsync((result: any) => {
          if (result?.status === OfficeAny.AsyncResultStatus.Succeeded && result?.value) resolve(result.value);
          else reject(new Error(result?.error?.message || "getAsFileAsync failed"));
        });
      } catch (error) {
        reject(error);
      }
    });

    const sliceCount = Number(file?.sliceCount || 0);
    if (!sliceCount || !file?.getSliceAsync) return "";

    const chunks: Uint8Array[] = [];
    try {
      for (let index = 0; index < sliceCount; index += 1) {
        const sliceResult: any = await new Promise((resolve, reject) => {
          try {
            file.getSliceAsync(index, (result: any) => {
              if (result?.status === OfficeAny.AsyncResultStatus.Succeeded && result?.value) resolve(result.value);
              else reject(new Error(result?.error?.message || `getSliceAsync failed @${index}`));
            });
          } catch (error) {
            reject(error);
          }
        });

        const data = sliceResult?.data;
        if (Array.isArray(data)) {
          chunks.push(Uint8Array.from(data));
        } else if (data instanceof ArrayBuffer) {
          chunks.push(new Uint8Array(data));
        } else if (ArrayBuffer.isView(data)) {
          chunks.push(new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
        } else if (typeof data === "string" && data) {
          chunks.push(Uint8Array.from(Array.from(data).map((char) => char.charCodeAt(0) & 0xff)));
        }
      }
    } finally {
      try {
        if (typeof file?.closeAsync === "function") {
          file.closeAsync(() => {});
        }
      } catch {
        // noop
      }
    }

    if (!chunks.length) return "";
    return uint8ArraysToBase64(chunks);
  } catch (error) {
    clientLog.warn("[office] getCurrentMessageAsEmlBase64 failed", error);
    return "";
  }
}

export async function waitForStableSelectedMessageContext(options?: {
  maxAttempts?: number;
  delayMs?: number;
  requirePreciseIdentity?: boolean;
}): Promise<{ context: OutlookMessageContext; itemToken: string }> {
  const maxAttempts = Math.max(1, Number(options?.maxAttempts) || 4);
  const delayMs = Math.max(40, Number(options?.delayMs) || 120);
  const requirePreciseIdentity = options?.requirePreciseIdentity !== false;
  let lastContext: OutlookMessageContext = {};
  let lastItemToken = "";
  let lastIdentity = "";
  let stableHits = 0;

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    const context = await getSelectedMessageContext().catch(() => ({} as OutlookMessageContext));
    const itemToken = await getCurrentItemToken().catch(() => "");
    const identity = [
      String(context.itemId || "").trim(),
      normalizeOutlookIdentityString(context.internetMessageId),
      String(context.conversationId || "").trim(),
    ].filter(Boolean).join("|");
    lastContext = context;
    lastItemToken = itemToken;

    if (identity && identity === lastIdentity) stableHits += 1;
    else stableHits = identity ? 1 : 0;
    if (identity) lastIdentity = identity;

    const hasPreciseIdentity = hasPreciseOutlookItemIdentity(context);
    if ((!requirePreciseIdentity || hasPreciseIdentity) && stableHits >= 2) {
      return {
        context,
        itemToken: itemToken || identity,
      };
    }
    if (attempt < maxAttempts - 1) {
      await sleep(delayMs);
    }
  }

  return {
    context: lastContext,
    itemToken: lastItemToken || lastIdentity,
  };
}

async function getAttachmentsViaEmlForCurrentItem(): Promise<OutlookAttachment[]> {
  try {
    const emlBase64 = await getCurrentMessageAsEmlBase64();
    if (!emlBase64) return [];

    const response = await fetch("/api/links/eml/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ emlBase64 }),
    });
    if (!response.ok) {
      clientLog.warn("[office] EML attachment extract failed", { status: response.status });
      return [];
    }
    const payload: any = await response.json();
    const rawAttachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
    return rawAttachments
      .map((attachment: any) => ({
        id: String(attachment?.id || "").trim() || undefined,
        name: String(attachment?.name || "").trim(),
        contentType: String(attachment?.contentType || "application/octet-stream").trim(),
        size: Number(attachment?.size || 0) || undefined,
        isInline: Boolean(attachment?.isInline),
        contentId: String(attachment?.contentId || "").trim() || undefined,
        content: String(attachment?.content || "").trim(),
      }))
      .filter((attachment: OutlookAttachment) => attachment.name && attachment.content);
  } catch (error) {
    clientLog.warn("[office] EML attachment fallback failed", error);
    return [];
  }
}

async function addCategoriesToCurrentItem(displayNames: string[]): Promise<OutlookCategoryMutationResult> {
  const uniqueNames = Array.from(new Set((displayNames || []).map((name) => String(name || "").trim()).filter(Boolean)));
  if (!uniqueNames.length) return { ok: true };
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) {
    return { ok: false, rawStatus: "unavailable", error: "item.categories.addAsync unavailable" };
  }

  return await new Promise<OutlookCategoryMutationResult>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.addAsync(uniqueNames, (result: any) => {
        const rawStatus = String(result?.status || "").trim() || "unknown";
        if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          const error = formatOutlookAsyncError(result?.error);
          clientLog.warn("[office] item.categories.addAsync failed", {
            categories: uniqueNames,
            error,
          });
          return resolve({
            ok: false,
            rawStatus,
            error: `item.categories.addAsync failed for ${uniqueNames.join(", ")}: ${error}`,
          });
        }
        resolve({ ok: true, rawStatus });
      });
    } catch {
      resolve({
        ok: false,
        rawStatus: "throw",
        error: `item.categories.addAsync threw for ${uniqueNames.join(", ")}`,
      });
    }
  });
}

async function removeCategoryFromCurrentItem(displayName: string): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.removeAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.removeAsync([displayName], () => resolve());
    } catch {
      resolve();
    }
  });
}

async function removeCategoriesFromCurrentItem(displayNames: string[]): Promise<OutlookCategoryMutationResult> {
  const uniqueNames = Array.from(new Set((displayNames || []).map((name) => String(name || "").trim()).filter(Boolean)));
  if (!uniqueNames.length) return { ok: true };
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.removeAsync) {
    return { ok: false, rawStatus: "unavailable", error: "item.categories.removeAsync unavailable" };
  }

  return await new Promise<OutlookCategoryMutationResult>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.removeAsync(uniqueNames, (result: any) => {
        const rawStatus = String(result?.status || "").trim() || "unknown";
        if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          const error = formatOutlookAsyncError(result?.error);
          clientLog.warn("[office] item.categories.removeAsync failed", {
            categories: uniqueNames,
            error,
          });
          return resolve({
            ok: false,
            rawStatus,
            error: `item.categories.removeAsync failed for ${uniqueNames.join(", ")}: ${error}`,
          });
        }
        resolve({ ok: true, rawStatus });
      });
    } catch {
      resolve({
        ok: false,
        rawStatus: "throw",
        error: `item.categories.removeAsync threw for ${uniqueNames.join(", ")}`,
      });
    }
  });
}

async function readCurrentItemCategoryNamesFromHost(): Promise<OutlookCategoryReadbackResult> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  const categoriesApi = OfficeAny?.context?.mailbox?.item?.categories;
  if (!categoriesApi) return { categories: [], source: "unavailable", rawStatus: "unavailable", error: "item.categories unavailable" };

  if (typeof categoriesApi.getAsync === "function") {
    const asyncRead = await new Promise<OutlookCategoryReadbackResult | null>((resolve) => {
      try {
        categoriesApi.getAsync((result: any) => {
          const rawStatus = String(result?.status || "").trim() || "unknown";
          if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
            const error = formatOutlookAsyncError(result?.error);
            return resolve({
              categories: [],
              source: "unavailable",
              rawStatus,
              error: `item.categories.getAsync failed: ${error}`,
            });
          }
          const value = Array.isArray(result.value) ? result.value : [];
          resolve({
            categories: normalizeUniqueCategoryValues(
              value
              .map((entry: any) => String(entry?.displayName || entry?.name || entry || "").trim())
              .filter(Boolean)
            ),
            source: "getAsync",
            rawStatus,
          });
        });
      } catch {
        resolve({
          categories: [],
          source: "unavailable",
          rawStatus: "throw",
          error: "item.categories.getAsync threw",
        });
      }
    });
    if (asyncRead?.source === "getAsync") return asyncRead;
    if (Array.isArray(categoriesApi)) {
      return {
        categories: normalizeUniqueCategoryValues(
          categoriesApi
            .map((entry: any) => String(entry?.displayName || entry?.name || entry || "").trim())
            .filter(Boolean)
        ),
        source: "array-fallback",
        rawStatus: asyncRead?.rawStatus || "fallback",
        error: asyncRead?.error,
      };
    }
    return asyncRead || { categories: [], source: "unavailable", rawStatus: "unknown", error: "item.categories.getAsync returned no result" };
  }

  if (Array.isArray(categoriesApi)) {
    return {
      categories: normalizeUniqueCategoryValues(
        categoriesApi
          .map((entry: any) => String(entry?.displayName || entry?.name || entry || "").trim())
          .filter(Boolean)
      ),
      source: "array-fallback",
      rawStatus: "array",
      error: "item.categories.getAsync unavailable",
    };
  }

  return { categories: [], source: "unavailable", rawStatus: "unavailable", error: "item.categories.getAsync unavailable" };
}

async function getCurrentItemCategoryNames(): Promise<string[]> {
  const readback = await readCurrentItemCategoryNamesFromHost();
  return readback.categories;
}

async function hasExpectedCurrentItemToken(expectedItemToken?: string): Promise<boolean> {
  if (!expectedItemToken) return true;
  const currentToken = await getCurrentItemToken().catch(() => "");
  if (!currentToken) {
    clientLog.warn("[office] current item token unavailable during category sync; continuing without strict guard", {
      expectedItemToken,
    });
    return true;
  }
  return currentToken === expectedItemToken;
}

function getCurrentManagedCategoryNames(currentCategories: string[], plan: OutlookCategoryPlan): string[] {
  const managedLabelSet = new Set(plan.managedLabelNames.map((label) => label.toLowerCase()));
  const managedSpecialSet = new Set(plan.managedSpecialCategories.map((label) => label.toLowerCase()));
  return currentCategories.filter((name) => {
    const normalized = String(name || "").trim().toLowerCase();
    return Boolean(
      (plan.manageClassificationFamilies && isManagedCategoryFamilyName(name))
      || managedLabelSet.has(normalized)
      || managedSpecialSet.has(normalized)
    );
  });
}

async function doesCurrentItemMatchOutlookCategoryPlan(
  plan: OutlookCategoryPlan,
  options?: { expectedItemToken?: string }
): Promise<boolean> {
  if (!(await hasExpectedCurrentItemToken(String(options?.expectedItemToken || "").trim()))) return false;
  const readback = await readCurrentItemCategoryNamesFromHost();
  return isOutlookCategoryPlanHostConfirmed(buildOutlookCategoryPlanDiff(readback.categories, plan), readback);
}

export async function applyOutlookCategoryPlan(
  plan: OutlookCategoryPlan,
  options?: {
    expectedItemToken?: string;
    isExecutionCurrent?: () => boolean;
    syncMeta?: {
      requestId: string;
      reason: string;
      mode: OutlookCategorySyncMode;
      itemIdentity: string;
      sourceSignature: string;
      planSignature: string;
      requestedAtMs: number;
      target?: OutlookCategorySyncTarget;
    };
  }
): Promise<ApplyOutlookCategoryPlanResult> {
  const expectedItemToken = String(options?.expectedItemToken || "").trim();
  const syncMeta = options?.syncMeta;
  const isExecutionCurrent = options?.isExecutionCurrent;
  const emptyDiff = buildOutlookCategoryPlanDiff([], plan);
  const diagnosticItemId = getOutlookCategoryApplyLogItemId(syncMeta);

  const readbackCurrentCategories = async (reason: string): Promise<OutlookCategoryReadbackResult> => {
    const readback = await readCurrentItemCategoryNamesFromHost();
    if (readback.source !== "getAsync") {
      logOutlookCategoryApplyDiagnostic("warn", "readback-fallback", {
        itemId: diagnosticItemId,
        reason,
        categoriesRequested: plan.desiredCategories,
        categoriesReadback: readback.categories,
        readbackSource: readback.source,
        readbackRawStatus: readback.rawStatus,
        fallbackReason: readback.error,
      });
    }
    return readback;
  };

  if (isExecutionCurrent && !isExecutionCurrent()) {
    logOutlookCategorySync("debug", "writer-skip-stale-before-read", syncMeta);
    return { result: "stale", detail: "stale-before-read", diff: emptyDiff };
  }
  if (!(await hasExpectedCurrentItemToken(expectedItemToken))) {
    return { result: "item-mismatch", detail: "item-token-before-read", diff: emptyDiff };
  }

  let readback = await readbackCurrentCategories("pre-apply");
  let diff = buildOutlookCategoryPlanDiff(readback.categories, plan);

  logOutlookCategorySync("debug", "writer-plan-diff", {
    ...syncMeta,
    currentCategories: diff.currentCategories,
    currentManagedCategories: diff.currentManagedCategories,
    desiredCategories: diff.desiredCategories,
    toAdd: diff.toAdd,
    toRemove: diff.toRemove,
  });
  logOutlookCategoryApplyDiagnostic("log", "resolved-plan", {
    itemId: diagnosticItemId,
    categoriesRequested: diff.desiredCategories,
    categoriesReadback: readback.categories,
    readbackSource: readback.source,
    readbackRawStatus: readback.rawStatus,
  });

  if (isOutlookCategoryPlanHostConfirmed(diff, readback)) {
    logOutlookCategorySync("info", "writer-noop", {
      ...syncMeta,
      currentCategories: diff.currentCategories,
      desiredCategories: diff.desiredCategories,
    });
    logOutlookCategoryApplyDiagnostic("log", "confirmed-noop-readback", {
      itemId: diagnosticItemId,
      categoriesRequested: diff.desiredCategories,
      categoriesReadback: readback.categories,
      readbackSource: readback.source,
    });
    return { result: "noop", diff };
  }

  const preparationErrors: string[] = [];
  for (const categoryName of diff.desiredCategories) {
    const masterCategoryResult = await ensureMasterCategory(categoryName, plan.desiredCategoryColors?.[categoryName]);
    if (!masterCategoryResult.ok && masterCategoryResult.error) {
      preparationErrors.push(masterCategoryResult.error);
    }
  }

  if (isExecutionCurrent && !isExecutionCurrent()) {
    logOutlookCategorySync("debug", "writer-skip-stale-after-prepare", syncMeta);
    return { result: "stale", detail: "stale-after-prepare", diff };
  }
  if (!(await hasExpectedCurrentItemToken(expectedItemToken))) {
    clientLog.warn("[office] applyOutlookCategoryPlan aborted after category preparation because the item changed", {
      expectedItemToken,
    });
    logOutlookCategorySync("warn", "writer-skip-item-mismatch-after-prepare", syncMeta);
    return { result: "item-mismatch", detail: "item-token-after-prepare", diff };
  }

  const writeErrors = [...preparationErrors];
  const maxAttempts = 3;

  for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
    if (attempt > 0) {
      await sleep(250);
      readback = await readbackCurrentCategories(`retry-pre-attempt-${attempt}`);
      diff = buildOutlookCategoryPlanDiff(readback.categories, plan);
    }

    if (isOutlookCategoryPlanHostConfirmed(diff, readback)) {
      logOutlookCategorySync("info", "writer-applied", {
        ...syncMeta,
        desiredCategories: diff.desiredCategories,
        currentCategories: diff.currentCategories,
        currentManagedCategories: diff.currentManagedCategories,
        toAdd: diff.toAdd,
        toRemove: diff.toRemove,
        attempt,
      });
      logOutlookCategoryApplyDiagnostic("log", "confirmed-readback", {
        itemId: diagnosticItemId,
        attempt,
        categoriesRequested: diff.desiredCategories,
        categoriesReadback: readback.categories,
        readbackSource: readback.source,
      });
      return { result: "success", diff };
    }

    if (diff.toAdd.length) {
      const addResult = await addCategoriesToCurrentItem(diff.toAdd);
      logOutlookCategoryApplyDiagnostic(addResult.ok ? "log" : "warn", "apply-add-result", {
        itemId: diagnosticItemId,
        attempt,
        categoriesRequested: diff.desiredCategories,
        categoriesApplied: diff.toAdd,
        applyRawStatus: addResult.rawStatus,
        applyError: addResult.error,
      });
      if (!addResult.ok && addResult.error) {
        writeErrors.push(addResult.error);
      }
    }

    if (isExecutionCurrent && !isExecutionCurrent()) {
      logOutlookCategorySync("debug", "writer-skip-stale-after-add", syncMeta);
      return { result: "stale", detail: "stale-after-add", diff };
    }
    if (!(await hasExpectedCurrentItemToken(expectedItemToken))) {
      clientLog.warn("[office] applyOutlookCategoryPlan skipped removals because the item changed", {
        expectedItemToken,
      });
      logOutlookCategorySync("warn", "writer-skip-item-mismatch-before-remove", syncMeta);
      return { result: "item-mismatch", detail: "item-token-before-remove", diff };
    }

    readback = await readbackCurrentCategories(`post-add-attempt-${attempt}`);
    diff = buildOutlookCategoryPlanDiff(readback.categories, plan);
    if (diff.toRemove.length) {
      const removeResult = await removeCategoriesFromCurrentItem(diff.toRemove);
      logOutlookCategoryApplyDiagnostic(removeResult.ok ? "log" : "warn", "apply-remove-result", {
        itemId: diagnosticItemId,
        attempt,
        categoriesRequested: diff.desiredCategories,
        categoriesRemoved: diff.toRemove,
        applyRawStatus: removeResult.rawStatus,
        applyError: removeResult.error,
      });
      if (!removeResult.ok && removeResult.error) {
        writeErrors.push(removeResult.error);
      }
    }

    await sleep(200);
    readback = await readbackCurrentCategories(`post-remove-attempt-${attempt}`);
    diff = buildOutlookCategoryPlanDiff(readback.categories, plan);

    logOutlookCategorySync("debug", "writer-attempt-finished", {
      ...syncMeta,
      attempt,
      currentCategories: diff.currentCategories,
      currentManagedCategories: diff.currentManagedCategories,
      desiredCategories: diff.desiredCategories,
      toAdd: diff.toAdd,
      toRemove: diff.toRemove,
      missingManagedCategories: diff.missingManagedCategories,
      unexpectedManagedCategories: diff.unexpectedManagedCategories,
      writeErrors,
    });
    logOutlookCategoryApplyDiagnostic(
      isOutlookCategoryPlanHostConfirmed(diff, readback) ? "log" : "warn",
      "attempt-readback",
      {
        itemId: diagnosticItemId,
        attempt,
        categoriesRequested: diff.desiredCategories,
        categoriesReadback: readback.categories,
        readbackSource: readback.source,
        readbackRawStatus: readback.rawStatus,
        fallbackReason: readback.error,
        missingManagedCategories: diff.missingManagedCategories,
        unexpectedManagedCategories: diff.unexpectedManagedCategories,
      }
    );

    if (isOutlookCategoryPlanHostConfirmed(diff, readback)) {
      logOutlookCategorySync("info", "writer-applied", {
        ...syncMeta,
        desiredCategories: diff.desiredCategories,
        currentCategories: diff.currentCategories,
        currentManagedCategories: diff.currentManagedCategories,
        attempt,
      });
      logOutlookCategoryApplyDiagnostic("log", "confirmed-readback", {
        itemId: diagnosticItemId,
        attempt,
        categoriesRequested: diff.desiredCategories,
        categoriesReadback: readback.categories,
        readbackSource: readback.source,
      });
      return { result: "success", diff };
    }

    for (const categoryName of diff.missingManagedCategories) {
      const masterCategoryResult = await ensureMasterCategory(categoryName, plan.desiredCategoryColors?.[categoryName]);
      if (!masterCategoryResult.ok && masterCategoryResult.error) {
        writeErrors.push(masterCategoryResult.error);
      }
    }
  }

  const detail = formatOutlookCategoryPlanDiffDetail(diff, writeErrors);
  logOutlookCategorySync("warn", "writer-verify-failed", {
    ...syncMeta,
    currentCategories: diff.currentCategories,
    currentManagedCategories: diff.currentManagedCategories,
    desiredCategories: diff.desiredCategories,
    missingManagedCategories: diff.missingManagedCategories,
    unexpectedManagedCategories: diff.unexpectedManagedCategories,
    detail,
  });
  logOutlookCategoryApplyDiagnostic("warn", "final-confirmation-failed", {
    itemId: diagnosticItemId,
    categoriesRequested: diff.desiredCategories,
    categoriesReadback: readback.categories,
    readbackSource: readback.source,
    readbackRawStatus: readback.rawStatus,
    fallbackReason: readback.error || detail,
    detail,
  });
  return { result: "failed", detail, diff };
}

function createOutlookCategorySyncResult(
  request: OutlookCategorySyncWriterRequest,
  input: {
    itemIdentity: string;
    result: OutlookCategoryWriterResult;
    detail?: string;
    target?: OutlookCategorySyncTarget;
    sourceSignature?: string;
    planSignature?: string;
  }
): OutlookCategorySyncResult {
  return {
    requestId: request.requestId,
    operationId: String(request.operationId || "").trim() || undefined,
    reason: request.reason,
    mode: request.mode,
    itemIdentity: input.itemIdentity,
    target: input.target,
    sourceSignature: input.sourceSignature,
    planSignature: input.planSignature,
    result: input.result,
    detail: input.detail,
    finishedAtIso: new Date().toISOString(),
  };
}

async function prepareCurrentItemOutlookCategorySyncWriterRequest(
  request: OutlookCategorySyncWriterRequest
): Promise<PreparedOutlookCategorySyncWriterRequest | OutlookCategoryWriterShortCircuit> {
  const stableSelection = await waitForStableSelectedMessageContext({
    maxAttempts: 4,
    delayMs: 120,
    requirePreciseIdentity: true,
  }).catch(() => ({
    context: {} as OutlookMessageContext,
    itemToken: "",
  }));
  const currentContext = stableSelection.context;
  if (!hasPreciseOutlookItemIdentity(currentContext)) {
    return {
      result: "failed",
      itemIdentity: buildOutlookCategoryWriterItemIdentity(request.target),
      detail: "no-identity",
    };
  }

  if (!doesOutlookCategorySyncTargetMatchContext(request.target, currentContext)) {
    return {
      result: "item-mismatch",
      itemIdentity: buildOutlookCategoryWriterItemIdentity(request.target) || buildOutlookCategoryWriterItemIdentity(currentContext),
    };
  }

  const expectedItemToken = String(request.expectedItemToken || "").trim() || stableSelection.itemToken || await getCurrentItemToken().catch(() => "");
  const payload = {
    itemId: String(currentContext.itemId || "").trim() || undefined,
    internetMessageId: String(currentContext.internetMessageId || "").trim() || undefined,
    conversationId: String(currentContext.conversationId || "").trim() || undefined,
    subject: String(currentContext.subject || "").trim() || undefined,
    fromEmail: String(currentContext.fromEmail || "").trim() || undefined,
    fromName: String(currentContext.fromName || "").trim() || undefined,
    receivedAtIso: String(currentContext.receivedDateTimeIso || "").trim() || undefined,
    messageDateIso: String(currentContext.receivedDateTimeIso || "").trim() || undefined,
  };
  const [settings, related, links] = await Promise.all([
    getSettings().catch(() => null),
    getRelatedEmailContext(payload).catch(() => null),
    getLinks(payload.conversationId, payload.internetMessageId, payload.itemId).catch(() => []),
  ]);
  const resolvedEmailCandidates = [
    related?.email || null,
    ...(Array.isArray(related?.emails) ? related.emails : []),
  ].filter(Boolean);
  const resolvedCurrentEmail = resolvedEmailCandidates.find((email) =>
    doesResolvedEmailMatchCurrentOutlookItem(email, currentContext)
  ) || null;
  if (!resolvedCurrentEmail) {
    clientLog.warn("[office] syncCurrentItemOutlookCategoriesFromContext skipped because related context resolved to a different email", {
      currentItemId: String(currentContext.itemId || "").trim(),
      currentInternetMessageId: normalizeOutlookIdentityString(currentContext.internetMessageId),
      resolvedItemId: String(related?.email?.itemId || "").trim(),
      resolvedInternetMessageId: normalizeOutlookIdentityString(related?.email?.internetMessageId),
    });
    return {
      result: "item-mismatch",
      itemIdentity: buildOutlookCategoryWriterItemIdentity(currentContext),
    };
  }
  const knownLabelNames = collectKnownOutlookCategoryLabelNames({
    settings,
    email: resolvedCurrentEmail,
    groups: Array.isArray(related?.groups) ? related.groups : [],
    tickets: Array.isArray(related?.tickets) ? related.tickets : [],
  });
  const snapshot = await getManagedOutlookCategorySnapshot(knownLabelNames).catch(() => null);
  const source = buildOutlookCategorySourceFromRelatedContext({
    email: resolvedCurrentEmail,
    groups: Array.isArray(related?.groups) ? related.groups : [],
    tickets: Array.isArray(related?.tickets) ? related.tickets : [],
    settings,
    currentOutlookLabelNames: snapshot?.labelNames || [],
    specialCategories: Array.isArray(links) && links.length ? [ODOO_LINKED_CATEGORY] : [],
    managedSpecialCategories: [ODOO_LINKED_CATEGORY],
  });
  const normalizedSource = normalizeOutlookCategorySource(source);
  const plan = buildOutlookCategoryPlan(normalizedSource, {
    manageClassificationFamilies: request.manageClassificationFamilies,
  });
  return {
    ...request,
    target: normalizeOutlookCategorySyncTarget(payload),
    itemIdentity: buildOutlookCategoryWriterItemIdentity(payload),
    expectedItemToken,
    source: normalizedSource,
    sourceSignature: getOutlookCategorySourceSignature(normalizedSource),
    plan,
    planSignature: getOutlookCategoryPlanSignature(plan),
  };
}

async function prepareSourceOutlookCategorySyncWriterRequest(
  request: OutlookCategorySyncWriterRequest
): Promise<PreparedOutlookCategorySyncWriterRequest | OutlookCategoryWriterShortCircuit> {
  const stableSelection = await waitForStableSelectedMessageContext({
    maxAttempts: 4,
    delayMs: 120,
    requirePreciseIdentity: false,
  }).catch(() => ({
    context: {} as OutlookMessageContext,
    itemToken: "",
  }));
  const normalizedTarget = normalizeOutlookCategorySyncTarget(request.target) || normalizeOutlookCategorySyncTarget(stableSelection.context);
  const fallbackTokenIdentity = String(stableSelection.itemToken || request.expectedItemToken || "").trim();
  const itemIdentity = buildOutlookCategoryWriterItemIdentity(normalizedTarget)
    || (fallbackTokenIdentity ? `token:${fallbackTokenIdentity}` : "");
  if (!itemIdentity) {
    return {
      result: "failed",
      itemIdentity: "",
      detail: "no-identity",
    };
  }

  if (normalizedTarget && !doesOutlookCategorySyncTargetMatchContext(normalizedTarget, stableSelection.context)) {
    return {
      result: "item-mismatch",
      itemIdentity,
    };
  }

  const source = normalizeOutlookCategorySource(request.source);
  const plan = buildOutlookCategoryPlan(source, {
    manageClassificationFamilies: request.manageClassificationFamilies,
  });
  return {
    ...request,
    target: normalizedTarget,
    itemIdentity,
    expectedItemToken: String(request.expectedItemToken || "").trim() || stableSelection.itemToken || await getCurrentItemToken().catch(() => ""),
    source,
    sourceSignature: getOutlookCategorySourceSignature(source),
    plan,
    planSignature: getOutlookCategoryPlanSignature(plan),
  };
}

async function prepareOutlookCategorySyncWriterRequest(
  request: OutlookCategorySyncWriterRequest
): Promise<PreparedOutlookCategorySyncWriterRequest | OutlookCategoryWriterShortCircuit> {
  if (request.mode === "current-item-context") {
    return await prepareCurrentItemOutlookCategorySyncWriterRequest(request);
  }
  return await prepareSourceOutlookCategorySyncWriterRequest(request);
}

function isSuccessfulOutlookCategoryWriterResult(result: OutlookCategoryWriterResult): boolean {
  return result === "success" || result === "duplicate";
}

async function runOutlookCategoryWriterRequest(
  request: OutlookCategorySyncWriterRequest
): Promise<OutlookCategorySyncResult> {
  const prepared = await prepareOutlookCategorySyncWriterRequest(request);
  if ("result" in prepared) {
    const shortCircuitResult = createOutlookCategorySyncResult(request, {
      itemIdentity: prepared.itemIdentity,
      result: prepared.result,
      detail: prepared.detail,
      target: request.target,
    });
    publishOutlookCategorySyncResult(shortCircuitResult);
    logOutlookCategorySync(prepared.result === "failed" ? "warn" : "debug", "request-short-circuited", {
      requestId: request.requestId,
      operationId: request.operationId,
      reason: request.reason,
      mode: request.mode,
      target: request.target,
      itemIdentity: prepared.itemIdentity,
      result: prepared.result,
      detail: prepared.detail,
    });
    return shortCircuitResult;
  }

  const state = getOutlookCategoryWriterState(prepared.itemIdentity);
  const activeOperation = readActiveOutlookCategoryOperationByItemIdentity(prepared.itemIdentity);
  if (activeOperation) {
    if (!prepared.operationId || activeOperation.operationId !== prepared.operationId) {
      const blockedResult = createOutlookCategorySyncResult(request, {
        itemIdentity: prepared.itemIdentity,
        result: "cancelled",
        detail: "blocked-by-active-operation",
        target: prepared.target,
        sourceSignature: prepared.sourceSignature,
        planSignature: prepared.planSignature,
      });
      publishOutlookCategorySyncResult(blockedResult);
      logOutlookCategorySync("info", "request-blocked-by-operation", {
        requestId: prepared.requestId,
        operationId: prepared.operationId,
        activeOperationId: activeOperation.operationId,
        reason: prepared.reason,
        mode: prepared.mode,
        itemIdentity: prepared.itemIdentity,
      });
      return blockedResult;
    }
  } else if (prepared.operationId) {
    const missingOperationResult = createOutlookCategorySyncResult(request, {
      itemIdentity: prepared.itemIdentity,
      result: "cancelled",
      detail: "operation-not-active",
      target: prepared.target,
      sourceSignature: prepared.sourceSignature,
      planSignature: prepared.planSignature,
    });
    publishOutlookCategorySyncResult(missingOperationResult);
    logOutlookCategorySync("warn", "request-cancelled-missing-operation", {
      requestId: prepared.requestId,
      operationId: prepared.operationId,
      reason: prepared.reason,
      mode: prepared.mode,
      itemIdentity: prepared.itemIdentity,
    });
    return missingOperationResult;
  }

  const freshness: OutlookCategoryWriterFreshness = {
    requestedAtMs: prepared.requestedAtMs,
    order: ++outlookCategoryWriterOrder,
    requestId: prepared.requestId,
  };
  if (compareOutlookCategoryWriterFreshness(freshness, state.latestFreshness) >= 0) {
    state.latestFreshness = freshness;
  }

  logOutlookCategorySync("debug", "request-enqueued", {
    requestId: prepared.requestId,
    reason: prepared.reason,
    mode: prepared.mode,
    itemIdentity: prepared.itemIdentity,
    target: prepared.target,
    sourceSignature: prepared.sourceSignature,
    planSignature: prepared.planSignature,
    requestedAtMs: prepared.requestedAtMs,
    freshness,
    latestFreshness: state.latestFreshness,
  });

  const run = async (): Promise<OutlookCategorySyncResult> => {
    const syncMeta = {
      requestId: prepared.requestId,
      operationId: prepared.operationId,
      reason: prepared.reason,
      mode: prepared.mode,
      itemIdentity: prepared.itemIdentity,
      sourceSignature: prepared.sourceSignature,
      planSignature: prepared.planSignature,
      requestedAtMs: prepared.requestedAtMs,
      target: prepared.target,
    };
    const isExecutionCurrent = () => compareOutlookCategoryWriterFreshness(freshness, state.latestFreshness) >= 0;

    logOutlookCategorySync("debug", "request-dequeued", {
      ...syncMeta,
      latestFreshness: state.latestFreshness,
    });

    if (hasSeenOutlookCategoryWriterRequestId(state, prepared.requestId)) {
      const duplicateResult = createOutlookCategorySyncResult(request, {
        itemIdentity: prepared.itemIdentity,
        result: "duplicate",
        detail: "request-id",
        target: prepared.target,
        sourceSignature: prepared.sourceSignature,
        planSignature: prepared.planSignature,
      });
      publishOutlookCategorySyncResult(duplicateResult);
      logOutlookCategorySync("info", "request-ignored-duplicate", syncMeta);
      return duplicateResult;
    }

    if (!isExecutionCurrent()) {
      const staleResult = createOutlookCategorySyncResult(request, {
        itemIdentity: prepared.itemIdentity,
        result: "stale",
        target: prepared.target,
        sourceSignature: prepared.sourceSignature,
        planSignature: prepared.planSignature,
      });
      publishOutlookCategorySyncResult(staleResult);
      logOutlookCategorySync("info", "request-ignored-stale", syncMeta);
      rememberOutlookCategoryWriterRequestId(state, prepared.requestId);
      return staleResult;
    }

    if (
      state.lastAppliedPlanSignature === prepared.planSignature
      && state.lastAppliedSourceSignature === prepared.sourceSignature
      && await doesCurrentItemMatchOutlookCategoryPlan(prepared.plan, { expectedItemToken: prepared.expectedItemToken })
    ) {
      const equivalentResult = createOutlookCategorySyncResult(request, {
        itemIdentity: prepared.itemIdentity,
        result: "duplicate",
        detail: "equivalent",
        target: prepared.target,
        sourceSignature: prepared.sourceSignature,
        planSignature: prepared.planSignature,
      });
      publishOutlookCategorySyncResult(equivalentResult);
      logOutlookCategorySync("info", "request-ignored-equivalent", syncMeta);
      rememberOutlookCategoryWriterRequestId(state, prepared.requestId);
      return equivalentResult;
    }

    const applied = await applyOutlookCategoryPlan(prepared.plan, {
      expectedItemToken: prepared.expectedItemToken,
      isExecutionCurrent,
      syncMeta,
    });
    rememberOutlookCategoryWriterRequestId(state, prepared.requestId);
    if (applied.result !== "success" && applied.result !== "noop") {
      const failedResult = createOutlookCategorySyncResult(request, {
        itemIdentity: prepared.itemIdentity,
        result: applied.result,
        detail: applied.detail,
        target: prepared.target,
        sourceSignature: prepared.sourceSignature,
        planSignature: prepared.planSignature,
      });
      publishOutlookCategorySyncResult(failedResult);
      logOutlookCategorySync(failedResult.result === "failed" ? "warn" : "info", "request-finished-without-apply", {
        ...syncMeta,
        result: failedResult.result,
        detail: failedResult.detail,
      });
      return failedResult;
    }

    state.lastAppliedPlanSignature = prepared.planSignature;
    state.lastAppliedSourceSignature = prepared.sourceSignature;
    dispatchOutlookCategoryContextInvalidated();
    const successResult = createOutlookCategorySyncResult(request, {
      itemIdentity: prepared.itemIdentity,
      result: "success",
      target: prepared.target,
      sourceSignature: prepared.sourceSignature,
      planSignature: prepared.planSignature,
    });
    publishOutlookCategorySyncResult(successResult);
    logOutlookCategorySync("info", "request-executed", syncMeta);
    return successResult;
  };

  const resultPromise = state.tail.catch(() => createOutlookCategorySyncResult(request, {
    itemIdentity: prepared.itemIdentity,
    result: "failed",
    target: prepared.target,
    sourceSignature: prepared.sourceSignature,
    planSignature: prepared.planSignature,
  })).then(run);
  state.tail = resultPromise;
  return await resultPromise;
}

export async function executeOutlookCategorySourceSync(
  source: Partial<OutlookCategorySource> | null | undefined,
  options?: {
    expectedItemToken?: string;
    manageClassificationFamilies?: boolean;
    requestId?: string;
    operationId?: string;
    requestedAtIso?: string;
    reason?: string;
    target?: OutlookCategorySyncTarget;
  }
): Promise<OutlookCategorySyncResult> {
  const requestId = String(options?.requestId || "").trim() || createOutlookCategorySyncRequestId(options?.reason || "source");
  const requestedAtMs = Date.parse(String(options?.requestedAtIso || "").trim()) || Date.now();
  return await runOutlookCategoryWriterRequest({
    requestId,
    operationId: String(options?.operationId || "").trim() || undefined,
    requestedAtMs,
    reason: String(options?.reason || "source-sync"),
    mode: "source",
    target: options?.target,
    source,
    expectedItemToken: options?.expectedItemToken,
    manageClassificationFamilies: options?.manageClassificationFamilies,
  });
}

export async function executeCurrentItemOutlookCategorySync(
  options?: {
    expectedItemToken?: string;
    requestId?: string;
    operationId?: string;
    requestedAtIso?: string;
    reason?: string;
    target?: OutlookCategorySyncTarget;
  }
): Promise<OutlookCategorySyncResult> {
  const requestId = String(options?.requestId || "").trim() || createOutlookCategorySyncRequestId(options?.reason || "current-item-context");
  const requestedAtMs = Date.parse(String(options?.requestedAtIso || "").trim()) || Date.now();
  return await runOutlookCategoryWriterRequest({
    requestId,
    operationId: String(options?.operationId || "").trim() || undefined,
    requestedAtMs,
    reason: String(options?.reason || "current-item-context"),
    mode: "current-item-context",
    target: options?.target,
    expectedItemToken: options?.expectedItemToken,
  });
}

export async function syncOutlookCategorySource(
  source: Partial<OutlookCategorySource> | null | undefined,
  options?: {
    expectedItemToken?: string;
    manageClassificationFamilies?: boolean;
    requestId?: string;
    operationId?: string;
    requestedAtIso?: string;
    reason?: string;
    target?: OutlookCategorySyncTarget;
  }
): Promise<boolean> {
  const result = await executeOutlookCategorySourceSync(source, options);
  return isSuccessfulOutlookCategoryWriterResult(result.result);
}

export async function syncCurrentItemOutlookCategoriesFromContext(
  options?: {
    expectedItemToken?: string;
    requestId?: string;
    operationId?: string;
    requestedAtIso?: string;
    reason?: string;
    target?: OutlookCategorySyncTarget;
  }
): Promise<boolean> {
  const result = await executeCurrentItemOutlookCategorySync(options);
  return isSuccessfulOutlookCategoryWriterResult(result.result);
}

export async function syncOdooLinkedCategory(hasLinks: boolean): Promise<void> {
  await syncOutlookCategorySource(
    {
      specialCategories: hasLinks ? [ODOO_LINKED_CATEGORY] : [],
      managedSpecialCategories: [ODOO_LINKED_CATEGORY],
    },
    {
      manageClassificationFamilies: false,
      reason: "odoo-linked",
    }
  );
}

export async function syncManagedOutlookCategories(input: LegacyManagedOutlookCategoryInput): Promise<void> {
  await syncOutlookCategorySource(buildOutlookCategorySourceFromLegacyInput(input), {
    reason: "legacy-managed-categories",
  });
}

export async function getManagedOutlookCategorySnapshot(): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}>;
export async function getManagedOutlookCategorySnapshot(knownLabelNames: string[]): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}>;
export async function getManagedOutlookCategorySnapshot(knownLabelNames?: string[]): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}> {
  const currentCategories = await getCurrentItemCategoryNames();
  const knownLabelSet = new Set(
    normalizeUniqueCategoryValues(knownLabelNames).map((label) => String(label || "").trim().toLowerCase())
  );
  const principalGroupNames = currentCategories
    .filter((name) => name.startsWith(GROUP_CATEGORY_PREFIX))
    .map((name) => name.slice(GROUP_CATEGORY_PREFIX.length).trim())
    .filter(Boolean);
  const referenceGroupNames = currentCategories
    .filter((name) => name.startsWith(REFERENCE_CATEGORY_PREFIX))
    .map((name) => name.slice(REFERENCE_CATEGORY_PREFIX.length).trim())
    .filter(Boolean);
  return {
    principalGroupNames,
    referenceGroupNames,
    groupNames: [...principalGroupNames, ...referenceGroupNames],
    ticketCodes: currentCategories
      .filter((name) => name.startsWith(TICKET_CATEGORY_PREFIX) || name.startsWith(LEGACY_TICKET_CATEGORY_PREFIX))
      .map((name) => name.startsWith(TICKET_CATEGORY_PREFIX) ? name.slice(TICKET_CATEGORY_PREFIX.length).trim() : name.slice(LEGACY_TICKET_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    statuses: currentCategories
      .filter((name) => name.startsWith(LEGACY_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(LEGACY_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    groupStatuses: currentCategories
      .filter((name) => name.startsWith(GROUP_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(GROUP_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    ticketStatuses: currentCategories
      .filter((name) => name.startsWith(TICKET_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(TICKET_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    labelStatuses: currentCategories
      .filter((name) => name.startsWith(LABEL_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(LABEL_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    labelNames: currentCategories
      .filter((name) =>
        name.startsWith(LEGACY_LABEL_CATEGORY_PREFIX)
        || (!isReservedManagedCategoryName(name) && knownLabelSet.has(String(name || "").trim().toLowerCase()))
      )
      .map((name) => name.startsWith(LEGACY_LABEL_CATEGORY_PREFIX) ? name.slice(LEGACY_LABEL_CATEGORY_PREFIX.length).trim() : String(name || "").trim())
      .filter(Boolean),
  };
}

export async function syncLinkCategoriesToComposeDraft(
  input: (Partial<OutlookCategorySource> | (LegacyManagedOutlookCategoryInput & { hasOdooLinks?: boolean })) | null | undefined,
  options?: { attempts?: number; delayMs?: number }
): Promise<void> {
  const attempts = Math.max(1, Number(options?.attempts || 12));
  const delayMs = Math.max(150, Number(options?.delayMs || 450));
  const source = "specialCategories" in (input || {}) || "managedSpecialCategories" in (input || {})
    ? buildOutlookCategoryPlan(input as Partial<OutlookCategorySource>).source
    : buildOutlookCategorySourceFromLegacyInput({
        ...(input as LegacyManagedOutlookCategoryInput & { hasOdooLinks?: boolean }),
        specialCategories: (input as { hasOdooLinks?: boolean } | null | undefined)?.hasOdooLinks ? [ODOO_LINKED_CATEGORY] : [],
        managedSpecialCategories: typeof (input as { hasOdooLinks?: boolean } | null | undefined)?.hasOdooLinks === "boolean" ? [ODOO_LINKED_CATEGORY] : [],
      });
  const hasManagedCategories = Boolean(
    source.principalGroupNames.length
    || source.referenceGroupNames.length
    || source.ticketCodes.length
    || source.groupStatuses.length
    || source.ticketStatuses.length
    || source.labelStatuses.length
    || source.labelNames.length
    || source.managedLabelNames.length
    || source.specialCategories.length
    || source.managedSpecialCategories.length
  );
  const hasOdooLinks = (input as { hasOdooLinks?: boolean } | null | undefined)?.hasOdooLinks === true;

  if (!hasManagedCategories && !hasOdooLinks) return;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (attempt > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }

    const composeReady = await isComposeMode().catch(() => false);
    if (!composeReady) continue;

    await syncOutlookCategorySource(source).catch(() => {
      // best-effort
    });
    return;
  }

  clientLog.warn("[office] syncLinkCategoriesToComposeDraft: compose draft not ready in time");
}

export async function setSubjectInComposeDraft(subject: string, options?: { attempts?: number; delayMs?: number }): Promise<void> {
  const desiredSubject = String(subject || "").trim();
  if (!desiredSubject) return;

  const attempts = Math.max(1, Number(options?.attempts || 12));
  const delayMs = Math.max(150, Number(options?.delayMs || 450));

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (attempt > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }

    const composeReady = await isComposeMode().catch(() => false);
    if (!composeReady) continue;

    const OfficeAny = await ensureOfficeReady().catch(() => null);
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item?.subject?.setAsync) continue;

    await new Promise<void>((resolve, reject) => {
      item.subject.setAsync(desiredSubject, (result: any) => {
        if (result?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result?.error?.message || "Erro ao definir assunto"));
      });
    });
    return;
  }

  clientLog.warn("[office] setSubjectInComposeDraft: compose draft subject not ready in time");
}

export async function syncOdooLinkedNotification(hasLinks: boolean, count = 0): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  const notifications = OfficeAny?.context?.mailbox?.item?.notificationMessages;
  if (!notifications?.replaceAsync || !notifications?.removeAsync) return;

  await new Promise<void>((resolve) => {
    try {
      if (!hasLinks) {
        notifications.removeAsync(ODOO_LINKED_NOTICE, () => resolve());
        return;
      }

      notifications.replaceAsync(
        ODOO_LINKED_NOTICE,
        {
          type: OfficeAny.MailboxEnums?.ItemNotificationMessageType?.InformationalMessage,
          message: count > 0
            ? `Este email tem ${count} ligacao(oes) ativa(s) ao Odoo.`
            : "Este email tem ligacoes ativas ao Odoo.",
          icon: "Icon.80x80",
          persistent: false,
        },
        () => resolve()
      );
    } catch {
      resolve();
    }
  });
}

let activeDialog: any = null;

export type AiReplyTargetSelection = {
  emailKey: string;
  itemId?: string;
  emailWebLink?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  messageDateIso?: string;
  receivedAtIso?: string;
  bodyText?: string;
  bodyHtml?: string;
};

type CockpitHostAction =
  | { type: "close" }
  | { type: "open-email"; itemId?: string; emailWebLink?: string }
  | { type: "reply-current" }
  | { type: "forward-current" }
  | {
      type: "sync-current-item-categories";
      requestId?: string;
      operationId?: string;
      requestedAtIso?: string;
      reason?: string;
      target?: OutlookCategorySyncTarget;
    }
  | {
      type: "sync-managed-categories";
      payload: (LegacyManagedOutlookCategoryInput & Partial<OutlookCategorySource>);
      requestId?: string;
      operationId?: string;
      requestedAtIso?: string;
      reason?: string;
      target?: OutlookCategorySyncTarget;
    };

async function executeCockpitHostAction(action: CockpitHostAction): Promise<boolean> {
  if (action.type === "close") {
    try {
      if (activeDialog) activeDialog.close();
    } catch { }
    activeDialog = null;
    return true;
  }

  if (action.type === "open-email") {
    await openLinkedOutlookEmail({ itemId: action.itemId, emailWebLink: action.emailWebLink });
    return true;
  }

  if (action.type === "reply-current") {
    await displayReplyForm("", true);
    return true;
  }

  if (action.type === "forward-current") {
    await displayForwardForm("", true);
    return true;
  }

  if (action.type === "sync-current-item-categories") {
    return await syncCurrentItemOutlookCategoriesFromContext({
      requestId: action.requestId,
      operationId: action.operationId,
      requestedAtIso: action.requestedAtIso,
      reason: action.reason || "host-action-current-item-context",
      target: action.target,
    });
  }

  if (action.type === "sync-managed-categories") {
    if ("groupNames" in (action.payload || {}) || "statuses" in (action.payload || {})) {
      await syncManagedOutlookCategories(action.payload || {});
      return true;
    }
    const synced = await syncOutlookCategorySource(action.payload || {}, {
      requestId: action.requestId,
      operationId: action.operationId,
      requestedAtIso: action.requestedAtIso,
      reason: action.reason || "host-action-source",
      target: action.target,
    });
    return synced;
  }

  return false;
}

function tryParseCockpitHostMessage(rawMessage: any): CockpitHostAction | null {
  const text = String(rawMessage || "").trim();
  if (!text) return null;
  if (text === "close") return { type: "close" };
  try {
    const parsed = JSON.parse(text);
    if (parsed?.type !== "host-action" || !parsed?.action?.type) return null;
    return parsed.action as CockpitHostAction;
  } catch {
    return null;
  }
}

function buildCockpitViewUrl(view: string, params: Record<string, string>) {
  const url = new URL(window.location.origin);
  url.searchParams.set("view", view);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));
  return url;
}

function isCockpitHostSyncAction(action: CockpitHostAction): boolean {
  return action.type === "sync-current-item-categories" || action.type === "sync-managed-categories";
}

function enqueueCockpitHostSyncFallback(action: CockpitHostAction): boolean {
  if (action.type === "sync-current-item-categories") {
    enqueueOutlookCategorySyncRequest({
      requestId: String(action.requestId || "").trim() || createOutlookCategorySyncRequestId(action.reason || "host-sync-current-item-fallback"),
      operationId: String(action.operationId || "").trim() || undefined,
      createdAtIso: String(action.requestedAtIso || "").trim() || new Date().toISOString(),
      reason: action.reason || "host-sync-current-item-fallback",
      mode: "current-item-context",
      target: action.target,
    });
    return true;
  }

  if (action.type === "sync-managed-categories") {
    enqueueOutlookCategorySyncRequest({
      requestId: String(action.requestId || "").trim() || createOutlookCategorySyncRequestId(action.reason || "host-sync-source-fallback"),
      operationId: String(action.operationId || "").trim() || undefined,
      createdAtIso: String(action.requestedAtIso || "").trim() || new Date().toISOString(),
      reason: action.reason || "host-sync-source-fallback",
      mode: "source",
      target: action.target,
      source: action.payload || {},
    });
    return true;
  }

  return false;
}

function tryOpenStandaloneWindow(url: URL, name: string, features: string): boolean {
  try {
    const popup = window.open(url.toString(), name, features);
    if (popup) {
      try {
        popup.focus();
      } catch {
        // best effort
      }
      return true;
    }
  } catch (error) {
    clientLog.warn("[office] standalone window fallback failed", error);
  }
  return false;
}

/**
 * Opens a separate window using Office Dialog API.
 * Guard: only one dialog at a time (evita "já existe uma dialog ativa").
 */
async function openCockpitView<T = void>(view: string, params: Record<string, string>, options?: { height?: number; width?: number; displayInIframe?: boolean; timeoutMs?: number }) {
  const OfficeAny = await ensureOfficeReady();
  const url = buildCockpitViewUrl(view, params);

  clientLog.log(`[office] openDialog ${url.toString()}`);

  // close previous if any
  try {
    if (activeDialog) activeDialog.close();
  } catch { }
  activeDialog = null;

  return await new Promise<T | null>((resolve, reject) => {
    let settled = false;
    const resolveOnce = (value: T | null = null) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    const rejectOnce = (error: Error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };
    const timer = setTimeout(() => {
      clientLog.warn(`[office] displayDialogAsync timeout for ${view}`);
      rejectOnce(new Error("A abertura da janela demorou demasiado tempo."));
    }, Math.max(2000, Number(options?.timeoutMs || 4000)));

    OfficeAny.context.ui.displayDialogAsync(
      url.toString(),
      { height: options?.height || 65, width: options?.width || 40, displayInIframe: Boolean(options?.displayInIframe) },
      (result: any) => {
        clearTimeout(timer);
        if (result.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.error(`[office] displayDialogAsync failed: ${result.error?.message || "unknown"}`);
          rejectOnce(new Error(result.error?.message || "Falha ao abrir janela (Dialog)."));
          return;
        }
        const dialog = result.value;
        activeDialog = dialog;

        dialog.addEventHandler(OfficeAny.EventType.DialogMessageReceived, (arg: any) => {
          const action = tryParseCockpitHostMessage(arg?.message);
          if (action) {
            void executeCockpitHostAction(action)
              .catch((error) => clientLog.error("[office] host action failed", error))
              .finally(() => {
                if (action.type === "close") {
                  resolveOnce();
                }
              });
            return;
          }

          const resultPayload = tryParseCockpitDialogResult<T>(arg?.message);
          if (typeof resultPayload !== "undefined") {
            try {
              if (activeDialog) activeDialog.close();
            } catch { }
            activeDialog = null;
            resolveOnce(resultPayload);
          }
        });

        dialog.addEventHandler(OfficeAny.EventType.DialogEventReceived, () => {
          activeDialog = null;
          resolveOnce();
        });
      }
    );
  });
}

export async function openCockpitDialog(params: Record<string, string>) {
  return await openCockpitView("dialog", params, { height: 65, width: 40, displayInIframe: false });
}

export async function openGroupExplorer(params: Record<string, string>) {
  try {
    return await openCockpitView("group-explorer", params, { height: 78, width: 52, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-explorer", params);
    clientLog.warn("[office] group explorer fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

function tryParseCockpitDialogResult<T>(rawMessage: any): T | undefined {
  const text = String(rawMessage || "").trim();
  if (!text) return undefined;
  try {
    const parsed = JSON.parse(text);
    if (parsed?.type !== "dialog-result") return undefined;
    return parsed.result as T;
  } catch {
    return undefined;
  }
}

export async function openGroupManager(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("group-manager", params, { height: 82, width: 58, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-manager", params);
    clientLog.warn("[office] group manager fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openAiSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("ai-settings", params, { height: 84, width: 60, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("ai-settings", params);
    clientLog.warn("[office] ai settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openGroupSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("group-settings", params, { height: 84, width: 60, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-settings", params);
    clientLog.warn("[office] group settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openGroupsTabSettings(params: Record<string, string> = {}) {
  return await openGroupSettings({
    surface: "groups-tab",
    ...params,
  });
}

export async function openGroupClassificationStudio(params: Record<string, string> = {}) {
  const url = buildCockpitViewUrl("group-classification-studio", params);
  try {
    return await openCockpitView("group-classification-studio", params, {
      height: 84,
      width: 74,
      displayInIframe: false,
      timeoutMs: 16000,
    });
  } catch (error) {
    clientLog.warn("[office] group classification studio dialog failed; retrying standalone dialog", error);
  }

  await sleep(450);

  try {
    return await openCockpitView("group-classification-studio", params, {
      height: 84,
      width: 74,
      displayInIframe: false,
      timeoutMs: 16000,
    });
  } catch (error) {
    clientLog.warn("[office] group classification studio retry failed; trying popup fallback", error);
  }

  const popupOpened = tryOpenStandaloneWindow(
    url,
    "iccc-group-classification-studio",
    "popup=yes,width=1520,height=980,resizable=yes,scrollbars=yes"
  );

  if (popupOpened) return null;

  throw new Error("Nao foi possivel abrir a janela externa do Classificar.");
}

export async function openAppSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("app-settings", params, { height: 86, width: 64, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("app-settings", params);
    clientLog.warn("[office] app settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openAiReplyTargetPicker(params: Record<string, string> = {}) {
  return await openCockpitView<AiReplyTargetSelection>("ai-reply-target-picker", params, { height: 84, width: 74, displayInIframe: true });
}

async function postCockpitHostActionToOpener(action: CockpitHostAction): Promise<boolean> {
  if (typeof window === "undefined") return false;
  const openerWindow = window.opener;
  if (!openerWindow || openerWindow === window || typeof openerWindow.postMessage !== "function") return false;
  if (action.type === "close") return false;

  const requestId = `host-action:${Date.now()}:${Math.random().toString(36).slice(2)}`;
  return await new Promise<boolean>((resolve) => {
    let settled = false;
    const finish = (ok: boolean) => {
      if (settled) return;
      settled = true;
      try {
        window.removeEventListener("message", handleMessage as EventListener);
      } catch {
        // ignore cleanup failures
      }
      window.clearTimeout(timeoutId);
      resolve(ok);
    };
    const handleMessage = (event: MessageEvent) => {
      if (event.origin !== window.location.origin) return;
      const payload: any = event.data;
      if (payload?.type !== HOST_ACTION_WINDOW_RESULT_TYPE) return;
      if (String(payload?.requestId || "") !== requestId) return;
      finish(payload?.ok === true);
    };
    const timeoutId = window.setTimeout(() => finish(false), 8000);
    window.addEventListener("message", handleMessage as EventListener);
    try {
      openerWindow.postMessage({
        type: HOST_ACTION_WINDOW_MESSAGE_TYPE,
        requestId,
        action,
      }, window.location.origin);
    } catch {
      finish(false);
    }
  });
}

export async function requestCockpitHostAction(action: CockpitHostAction): Promise<boolean> {
  const isSyncAction = isCockpitHostSyncAction(action);
  const hasOpenerBridge = typeof window !== "undefined"
    && Boolean(window.opener && window.opener !== window && typeof window.opener.postMessage === "function");
  const openerResult = await postCockpitHostActionToOpener(action).catch(() => false);
  if (openerResult) return true;

  let hasMessageParentBridge = false;
  try {
    const OfficeAny = await ensureOfficeReady();
    hasMessageParentBridge = typeof OfficeAny?.context?.ui?.messageParent === "function";
    if (hasMessageParentBridge) {
      OfficeAny.context.ui.messageParent(JSON.stringify({ type: "host-action", action }));
      return true;
    }
  } catch {
    // fall through
  }

  if (isSyncAction && (hasOpenerBridge || hasMessageParentBridge)) {
    return enqueueCockpitHostSyncFallback(action);
  }

  try {
    return await executeCockpitHostAction(action);
  } catch {
    return false;
  }
}

type WindowHostActionMessage = {
  type: typeof HOST_ACTION_WINDOW_MESSAGE_TYPE;
  requestId: string;
  action: CockpitHostAction;
};

function installCockpitWindowHostActionBridge() {
  if (typeof window === "undefined" || typeof window.addEventListener !== "function") return;
  const host = window as typeof window & { __icccWindowHostActionBridgeInstalled?: boolean };
  if (host.__icccWindowHostActionBridgeInstalled) return;
  host.__icccWindowHostActionBridgeInstalled = true;
  window.addEventListener("message", (event: MessageEvent) => {
    if (event.origin !== window.location.origin) return;
    const payload = event.data as WindowHostActionMessage | undefined;
    if (payload?.type !== HOST_ACTION_WINDOW_MESSAGE_TYPE) return;
    const sourceWindow = event.source as WindowProxy | null;
    const requestId = String(payload?.requestId || "").trim();
    const action = payload?.action;
    if (!requestId || !action?.type) return;

    void (async () => {
      let ok = false;
      try {
        ok = action.type === "close"
          ? false
          : await requestCockpitHostAction(action);
      } catch {
        ok = false;
      }
      try {
        sourceWindow?.postMessage({
          type: HOST_ACTION_WINDOW_RESULT_TYPE,
          requestId,
          ok,
        }, event.origin);
      } catch {
        // best effort only
      }
    })();
  });
}

installCockpitWindowHostActionBridge();

/**
 * Subscribe to selection change (when user clicks a different email).
 * IMPORTANT: This must NEVER open dialogs. Only refresh the taskpane state.
 */
export async function subscribeToItemChanges(onChanged: () => void): Promise<() => void> {
  const OfficeAny = await ensureOfficeReady();

  const handler = () => {
    try {
      onChanged();
    } catch (e) {
      clientLog.error("[office] ItemChanged handler error", e);
    }
  };

  try {
    if (OfficeAny?.context?.mailbox?.addHandlerAsync) {
      OfficeAny.context.mailbox.addHandlerAsync(OfficeAny.EventType.ItemChanged, handler);
      clientLog.log("[office] subscribed ItemChanged");
      return () => {
        try {
          OfficeAny.context.mailbox.removeHandlerAsync(OfficeAny.EventType.ItemChanged, { handler });
          clientLog.log("[office] unsubscribed ItemChanged");
        } catch { }
      };
    }
  } catch (e) {
    clientLog.warn("[office] ItemChanged not supported here", e);
  }

  return () => { };
}

/**
 * Checks if the current item is in compose mode (editable).
 */
export async function isComposeMode(): Promise<boolean> {
  const OfficeAny = await ensureOfficeReady();
  return Boolean(OfficeAny?.context?.mailbox?.item?.body?.setSelectedDataAsync);
}

/**
 * Inserts HTML or Text into the message body at the current cursor position.
 * Only works in Compose mode.
 */
export async function insertTextToBody(content: string, isHtml = true): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  clientLog.log(`[office] insertTextToBody called. Item exists: ${!!item}, setSelectedDataAsync exists: ${!!item?.body?.setSelectedDataAsync}`);

  if (!item?.body?.setSelectedDataAsync) {
    clientLog.warn("[office] insertTextToBody: Not in compose mode or not supported (setSelectedDataAsync missing).");
    throw new Error("Não é possível inserir texto: o item não está em modo de edição ou a funcionalidade não é suportada.");
  }

  // Formatting fix: Convert newlines to <br> if HTML
  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;

  return await new Promise<void>((resolve, reject) => {
    item.body.setSelectedDataAsync(
      finalContent,
      { coercionType: isHtml ? OfficeAny.CoercionType.Html : OfficeAny.CoercionType.Text },
      (result: any) => {
        if (result.status === OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.log("[office] insertTextToBody: Success");
          resolve();
        } else {
          clientLog.error(`[office] setSelectedDataAsync failed: ${result.error?.message || "unknown"}`);
          reject(new Error(result.error?.message || "Falha ao inserir no email."));
        }
      }
    );
  });
}

/**
 * Opens a new Reply form with pre-filled content.
 */
export async function displayReplyForm(content: string, isHtml = true, options?: { replyAll?: boolean }): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  if (!item?.displayReplyAllForm && !item?.displayReplyForm) {
    throw new Error("Funcionalidade de resposta não disponível neste item.");
  }

  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;
  const replyAll = options?.replyAll !== false;

  if (replyAll && typeof item.displayReplyAllForm === "function") {
    if (isHtml) item.displayReplyAllForm({ htmlBody: finalContent });
    else item.displayReplyAllForm(finalContent);
    return;
  }

  if (typeof item.displayReplyForm === "function") {
    if (isHtml) item.displayReplyForm({ htmlBody: finalContent });
    else item.displayReplyForm(finalContent);
    return;
  }

  if (isHtml) item.displayReplyAllForm({ htmlBody: finalContent });
  else item.displayReplyAllForm(finalContent);
}

/**
 * Opens a new Forward form with pre-filled content.
 */
export async function displayForwardForm(content: string, isHtml = true): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  if (!item?.displayForwardForm) {
    throw new Error("Funcionalidade de reenvio não disponível neste item.");
  }

  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;

  item.displayForwardForm({ htmlBody: isHtml ? finalContent : undefined, textBody: !isHtml ? finalContent : undefined });
}

/**
 * Opens a brand new email form with pre-filled recipients, subject and body.
 */
export async function displayNewMessageForm(params: {
  toRecipients?: string[];
  ccRecipients?: string[];
  bccRecipients?: string[];
  subject?: string;
  body?: string;
  isHtml?: boolean;
}): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const mailbox = OfficeAny?.context?.mailbox;
  if (!mailbox?.displayNewMessageForm) {
    throw new Error("Funcionalidade de nova mensagem nÃ£o disponÃ­vel neste ambiente.");
  }

  const isHtml = params.isHtml !== false;
  mailbox.displayNewMessageForm({
    toRecipients: Array.isArray(params.toRecipients) ? params.toRecipients : [],
    ccRecipients: Array.isArray(params.ccRecipients) ? params.ccRecipients : [],
    bccRecipients: Array.isArray(params.bccRecipients) ? params.bccRecipients : [],
    subject: params.subject,
    htmlBody: isHtml ? String(params.body || "") : undefined,
    body: !isHtml ? String(params.body || "") : undefined,
  });
}

/**
 * Opens the New Appointment form with pre-filled details.
 */
export async function displayNewMeetingForm(params: {
  subject?: string;
  body?: string;
  location?: string;
  start?: Date;
  end?: Date;
  requiredAttendees?: string[];
}) {
  const OfficeAny = await ensureOfficeReady();
  const mailbox = OfficeAny?.context?.mailbox;

  if (!mailbox?.displayNewAppointmentForm) {
    throw new Error("Calendário não suportado neste ambiente.");
  }

  mailbox.displayNewAppointmentForm({
    subject: params.subject,
    body: params.body,
    location: params.location,
    start: params.start,
    end: params.end,
    requiredAttendees: params.requiredAttendees,
  });
}

/**
 * Sets recipients in a compose item.
 * @param type 'to' | 'cc' | 'bcc'
 * @param recipients Array of email strings or Recipient objects
 */
export async function setRecipients(type: 'to' | 'cc' | 'bcc', recipients: (string | Recipient)[]): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  const target = item?.[type];

  if (!target?.setAsync) {
    clientLog.warn(`[office] setRecipients: target ${type} does not support setAsync`);
    return;
  }

  const formatted = recipients.map(r => typeof r === 'string' ? r : r.email);

  return await new Promise<void>((resolve, reject) => {
    target.setAsync(formatted, (result: any) => {
      if (result.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || `Erro ao definir destinatários ${type}`));
    });
  });
}

/**
 * Sets the subject of a compose item.
 */
export async function setSubject(subject: string): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  if (!item?.subject?.setAsync) {
    clientLog.warn("[office] setSubject: item.subject does not support setAsync");
    return;
  }

  return await new Promise<void>((resolve, reject) => {
    item.subject.setAsync(subject, (result: any) => {
      if (result.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Erro ao definir assunto"));
    });
  });
}

export async function addBase64AttachmentToCompose(name: string, contentBase64: string): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  if (!item?.addFileAttachmentFromBase64Async) {
    throw new Error("Abre uma mensagem em modo de edição para anexar documentos.");
  }

  const safeName = String(name || "").trim() || "documento";
  const base64 = String(contentBase64 || "").trim().replace(/^data:[^,]+,/, "");
  if (!base64) {
    throw new Error("O documento não tem conteúdo disponível para anexar.");
  }

  await new Promise<void>((resolve, reject) => {
    try {
      item.addFileAttachmentFromBase64Async(base64, safeName, { isInline: false }, (result: any) => {
        if (result?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result?.error?.message || "Falha ao anexar o documento."));
      });
    } catch (error: any) {
      reject(error);
    }
  });
}

/**
 * Fetch attachments from the current item.
 * Returns array of attachment metadata plus content when available.
 */
export async function getAttachments(): Promise<OutlookAttachment[]> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;

    if (!item?.attachments) return [];

    const attachments = item.attachments;
    const fileAttachments = Array.from(attachments).filter((att: any) => att?.attachmentType === "file");
    const results: OutlookAttachment[] = [];

    for (const att of fileAttachments) {
      // Only process file attachments
      try {
        const content = await new Promise<string>((resolve, reject) => {
          item.getAttachmentContentAsync((att as any).id, async (result: any) => {
            if (result.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              reject(new Error(result.error?.message));
              return;
            }

            const format = String(result?.value?.format || "").trim().toLowerCase();
            const rawContent = String(result?.value?.content || "").trim();
            if (!rawContent) {
              resolve("");
              return;
            }

            if (!format || format === "base64") {
              resolve(rawContent);
              return;
            }

            if (format === "url") {
              try {
                const response = await fetch(rawContent);
                if (!response.ok) {
                  reject(new Error(`Falha ao descarregar conteudo do anexo (${response.status})`));
                  return;
                }
                const buffer = await response.arrayBuffer();
                resolve(arrayBufferToBase64(buffer));
                return;
              } catch (error: any) {
                reject(error);
                return;
              }
            }

            resolve(rawContent);
          });
        });

        results.push({
          id: String((att as any).id || "").trim() || undefined,
          name: String((att as any).name || "").trim(),
          contentType: String((att as any).contentType || "").trim(),
          size: Number((att as any).size || 0) || undefined,
          isInline: Boolean((att as any).isInline),
          contentId: String((att as any).contentId || "").trim() || undefined,
          content: content,
        });
      } catch (e) {
        clientLog.error(`[office] Failed to download attachment ${(att as any).name}`, e);
      }
    }

    let mergedResults = results;

    const needsEmlFallback =
      fileAttachments.length > 0 &&
      (mergedResults.length < fileAttachments.length || !mergedResults.some((attachment) => String(attachment.content || "").trim()));
    if (needsEmlFallback) {
      const emlResults = await getAttachmentsViaEmlForCurrentItem();
      if (emlResults.length) {
        mergedResults = mergeOutlookAttachments(mergedResults, emlResults);
      }
    }

    const needsGraphFallback =
      fileAttachments.length > 0 &&
      (mergedResults.length < fileAttachments.length || !mergedResults.some((attachment) => String(attachment.content || "").trim()));
    if (GRAPH_RUNTIME_ENABLED && needsGraphFallback) {
      const graphResults = await getAttachmentsViaGraphForCurrentItem();
      if (graphResults.length) {
        return mergeOutlookAttachments(mergedResults, graphResults);
      }
    }

    return mergedResults;
  } catch (error) {
    clientLog.error("[office] getAttachments error", error);
    return [];
  }
}

export async function openLinkedOutlookEmail(target: { itemId?: string; emailWebLink?: string }): Promise<boolean> {
  const itemId = String(target?.itemId || "").trim();
  if (itemId) {
    const OfficeAny: any = await ensureOfficeReady().catch(() => null);
    const mailbox = OfficeAny?.context?.mailbox;
    if (typeof mailbox?.displayMessageForm === "function") {
      mailbox.displayMessageForm(itemId);
      return true;
    }
  }

  const emailWebLink = String(target?.emailWebLink || "").trim();
  if (emailWebLink) {
    window.open(emailWebLink, "_blank", "noopener,noreferrer");
    return true;
  }

  return false;
}
