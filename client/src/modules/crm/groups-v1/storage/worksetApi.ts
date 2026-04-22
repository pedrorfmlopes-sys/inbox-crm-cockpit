import type { GroupWorksetManifest } from "./types";

export type GroupStorageValidationResult = {
  mode: string;
  provider: string;
  fileBacked: boolean;
  supported: boolean;
  basePath: string;
  normalizedBasePath: string;
  isWebUrl: boolean;
  requiresServerAccessiblePath: boolean;
  canStoreManifest: boolean;
  canStoreBinary: boolean;
  pickerAvailable: boolean;
  pickerBlockedReason?: string;
  blockingReason?: string;
  notes: string[];
};

const WORKSET_API_TIMEOUT_MS = 30000;

async function requestWorksetJson<T>(path: string, init?: RequestInit): Promise<T> {
  const controller = new AbortController();
  const timeoutMessage = `Pedido de workset excedeu ${Math.round(WORKSET_API_TIMEOUT_MS / 1000)}s: ${path}`;
  const id = setTimeout(() => controller.abort(timeoutMessage), WORKSET_API_TIMEOUT_MS);

  try {
    const res = await fetch(path, {
      ...init,
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        ...(init?.headers || {}),
      },
    });
    const ct = (res.headers.get("content-type") || "").toLowerCase();
    const body = ct.includes("application/json") ? await res.json() : await res.text();
    if (!res.ok || (body && typeof body === "object" && body.ok === false)) {
      const message = typeof body === "string" ? body : body?.details || body?.message || body?.error || JSON.stringify(body);
      throw new Error(`HTTP ${res.status}: ${message}`);
    }
    return body as T;
  } catch (error: unknown) {
    const aborted = controller.signal.aborted || (error instanceof Error && error.name === "AbortError");
    if (!aborted) throw error;
    const reason =
      typeof controller.signal.reason === "string" && controller.signal.reason.trim()
        ? controller.signal.reason.trim()
        : timeoutMessage;
    throw new Error(reason);
  } finally {
    clearTimeout(id);
  }
}

export async function getGroupWorksetManifest(
  worksetKey: string,
  options?: {
    mode?: string;
    basePath?: string;
    chosenFolderKind?: string;
    primaryTarget?: string;
  }
): Promise<GroupWorksetManifest | null> {
  const normalizedWorksetKey = String(worksetKey || "").trim();
  if (!normalizedWorksetKey) return null;
  const params = new URLSearchParams();
  params.set("_ts", String(Date.now()));
  if (options?.mode) params.set("mode", String(options.mode));
  if (options?.basePath) params.set("basePath", String(options.basePath));
  if (options?.chosenFolderKind) params.set("chosenFolderKind", String(options.chosenFolderKind));
  if (options?.primaryTarget) params.set("primaryTarget", String(options.primaryTarget));
  try {
    const response = await requestWorksetJson<{ manifest?: GroupWorksetManifest | null }>(
      `/api/links/groups/worksets/${encodeURIComponent(normalizedWorksetKey)}?${params.toString()}`
    );
    return response?.manifest || null;
  } catch (error: unknown) {
    const message = error instanceof Error ? error.message : String(error || "");
    if (/HTTP 404/i.test(message)) return null;
    throw error;
  }
}

export async function saveGroupWorksetManifest(payload: {
  manifest: GroupWorksetManifest;
  keepalive?: boolean;
}): Promise<GroupWorksetManifest | null> {
  const response = await requestWorksetJson<{ manifest?: GroupWorksetManifest | null }>(`/api/links/groups/worksets`, {
    method: "POST",
    keepalive: payload.keepalive === true,
    body: JSON.stringify({
      manifest: payload.manifest,
    }),
  });
  return response?.manifest || null;
}

export async function migrateGroupWorksetManifest(payload: {
  worksetKey: string;
  sourceLocation?: Record<string, unknown> | null;
  targetLocation: Record<string, unknown>;
  removeSource?: boolean;
}): Promise<{
  manifest?: GroupWorksetManifest | null;
  removedSourceMirror?: boolean;
  targetValidation?: GroupStorageValidationResult;
}> {
  return await requestWorksetJson(`/api/links/groups/worksets/migrate`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
}

export async function validateGroupStorageTarget(payload: Record<string, unknown>): Promise<GroupStorageValidationResult> {
  const response = await requestWorksetJson<{ result?: GroupStorageValidationResult }>(`/api/links/groups/storage/validate`, {
    method: "POST",
    body: JSON.stringify(payload),
  });
  return response?.result || {
    mode: String(payload.mode || "supabase"),
    provider: "cloud",
    fileBacked: false,
    supported: false,
    basePath: "",
    normalizedBasePath: "",
    isWebUrl: false,
    requiresServerAccessiblePath: false,
    canStoreManifest: false,
    canStoreBinary: false,
    pickerAvailable: false,
    blockingReason: "Nao foi possivel validar o destino configurado.",
    notes: [],
  };
}
